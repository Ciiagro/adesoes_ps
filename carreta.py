"""
Lógica do módulo "Carreta do Agro": pedidos de disponibilização de carretas
pra eventos, com calendário de disponibilidade (pra evitar pedido em cima de
data já reservada). As rotas Flask (em app.py) só chamam essas funções.
"""
import calendar
from datetime import datetime, date, timedelta
from zoneinfo import ZoneInfo

from flask import url_for
from sqlalchemy import or_

from email_utils import enviar_email
from models import SolicitacaoCarreta, ArquivoCarreta, BloqueioCarreta, SindicatoRural, MunicipioSindicato
import distancia_municipios
import supabase_storage
import pdf_compressao

TIPOS_RECURSO = {
    "carreta_agro": "Carreta do Agro",
    "carreta_saude": "Carreta da Saúde",
}

STATUS_LABEL = {
    "pendente": "Pendente",
    "aprovada": "Aprovada",
    "realizada": "Realizada",
    "recusada": "Cancelada",
}

# Status que "ocupam" o motorista/veículo — usado tanto na checagem de
# conflito quanto na linha do tempo do mapa. "realizada" entra aqui porque
# é a mesma reserva de antes, só que já passou da data.
STATUS_CONFIRMADOS = ("pendente", "aprovada", "realizada")

ARQUIVO_TIPOS_PERMITIDOS = {"application/pdf"}
ARQUIVO_TAMANHO_MAXIMO_BYTES = 10 * 1024 * 1024  # 10 MB por arquivo

# Categorias de arquivo anexado a uma solicitação — hoje são 2 documentos
# obrigatórios no formulário público: o ofício e o termo de compromisso
# (baixado do site, assinado no gov.br e reanexado aqui já assinado).
CATEGORIA_OFICIO = "oficio"
CATEGORIA_TERMO_COMPROMISSO = "termo_compromisso"

CATEGORIA_ARQUIVO_LABEL = {
    CATEGORIA_OFICIO: "Ofício",
    CATEGORIA_TERMO_COMPROMISSO: "Termo de Ciência assinado",
}


def _processar_anexos(db, solicitacao_id, arquivos, categoria):
    """Valida (precisa ser PDF e caber no limite de tamanho) e sobe pro
    Storage cada arquivo da lista, registrando um ArquivoCarreta pra cada
    um com a `categoria` informada (ofício ou termo de compromisso
    assinado). Compartilhada entre `criar_solicitacao` e
    `atualizar_solicitacao` pra não duplicar a mesma validação duas vezes."""
    for arquivo in arquivos or []:
        if not arquivo or not arquivo.filename:
            continue
        tipo = (arquivo.mimetype or "").lower()
        if tipo not in ARQUIVO_TIPOS_PERMITIDOS:
            raise ValueError(f'O arquivo "{arquivo.filename}" precisa ser um PDF.')
        conteudo = pdf_compressao.comprimir_pdf_se_necessario(arquivo.read())
        if len(conteudo) > ARQUIVO_TAMANHO_MAXIMO_BYTES:
            raise ValueError(f'O arquivo "{arquivo.filename}" passou de 10 MB (mesmo após compressão automática).')

        try:
            caminho = supabase_storage.upload_oficio_carreta(solicitacao_id, conteudo, arquivo.filename)
        except supabase_storage.SupabaseStorageError as erro:
            raise ValueError(f"Não consegui enviar o arquivo pro Storage: {erro}")

        db.add(ArquivoCarreta(
            solicitacao_id=solicitacao_id,
            nome_arquivo=arquivo.filename,
            storage_path=caminho,
            tipo_mime=tipo,
            categoria=categoria,
        ))


def _limpar(valor):
    if valor is None:
        return None
    valor = str(valor).strip()
    return valor or None


def obter_sindicatos_e_municipios(db):
    """Pra alimentar o formulário público: a lista de sindicatos ativos
    (em ordem alfabética) e, pra cada um, a lista de municípios pelos quais
    ele responde (também em ordem alfabética) — usado pro campo de
    Município se limitar às opções do sindicato escolhido.

    Faz só DUAS consultas ao banco (uma pros sindicatos, uma pra TODOS os
    municípios de uma vez) e junta tudo em Python — evita repetir uma
    consulta por sindicato (isso é o que tava deixando a tela lenta)."""
    sindicatos = (
        db.query(SindicatoRural)
        .filter(SindicatoRural.ativo.is_(True))
        .order_by(SindicatoRural.nome)
        .all()
    )

    municipios_do_sindicato = {}
    linhas = (
        db.query(MunicipioSindicato.sindicato_id, MunicipioSindicato.nome)
        .filter(MunicipioSindicato.sindicato_id.isnot(None))
        .order_by(MunicipioSindicato.nome)
        .all()
    )
    for sindicato_id, nome_municipio in linhas:
        municipios_do_sindicato.setdefault(sindicato_id, []).append(nome_municipio)

    municipios_por_sindicato = {
        sindicato.nome: municipios_do_sindicato.get(sindicato.id, [])
        for sindicato in sindicatos
    }
    return [s.nome for s in sindicatos], municipios_por_sindicato


def _titulo(valor):
    """Padroniza texto digitado (evento, responsável, sindicato...): cada
    palavra com inicial maiúscula e o resto minúsculo. Ex: 'JOÃO da SILVA'
    -> 'João Da Silva'. Preserva múltiplos espaços como um só (efeito
    colateral do split), o que é o comportamento esperado aqui."""
    valor = _limpar(valor)
    if not valor:
        return valor
    return " ".join(p[:1].upper() + p[1:].lower() for p in valor.split(" ") if p)


def _texto_para_data(valor):
    valor = _limpar(valor)
    if not valor:
        return None
    try:
        return datetime.strptime(valor, "%Y-%m-%d").date()
    except ValueError:
        return None


def atualizar_realizadas_automaticamente(db, hoje=None):
    """Muda sozinho pra 'realizada' toda solicitação que estava 'aprovada'
    mas cuja data final já passou — não precisa de nenhuma tarefa agendada
    (cron): isso roda de leve toda vez que a listagem ou o mapa da carreta
    são abertos, então o status fica sempre em dia sem trabalho manual.
    Quem quiser marcar como realizada ANTES da data acabar (evento que
    terminou mais cedo, por exemplo) ainda pode fazer isso na mão, pelo
    botão "Marcar como realizada" na tela de detalhe."""
    hoje = hoje or date.today()
    pendentes_de_atualizar = (
        db.query(SolicitacaoCarreta)
        .filter(SolicitacaoCarreta.status == "aprovada")
        .filter(SolicitacaoCarreta.data_fim < hoje)
        .all()
    )
    for solicitacao in pendentes_de_atualizar:
        solicitacao.status = "realizada"
    if pendentes_de_atualizar:
        db.commit()
    return len(pendentes_de_atualizar)


def listar_solicitacoes(db, status=None, tipo_recurso=None, ano=None, mes=None, busca=None):
    """Lista de solicitações, mais recentes primeiro. Filtra por status,
    tipo de recurso, mês (ano+mes juntos) e/ou busca livre (qualquer trecho
    de município, evento, sindicato ou responsável) quando informado. O
    filtro de mês considera qualquer solicitação cujo período TOQUE aquele
    mês (não só as que começam nele) — então um evento que atravessa a
    virada do mês aparece nos dois."""
    atualizar_realizadas_automaticamente(db)
    query = db.query(SolicitacaoCarreta).order_by(SolicitacaoCarreta.data_inicio.desc())
    if status:
        query = query.filter(SolicitacaoCarreta.status == status)
    if tipo_recurso:
        query = query.filter(SolicitacaoCarreta.tipo_recurso == tipo_recurso)
    if ano and mes:
        primeiro_dia = date(ano, mes, 1)
        ultimo_dia = date(ano, mes, calendar.monthrange(ano, mes)[1])
        query = query.filter(SolicitacaoCarreta.data_inicio <= ultimo_dia)
        query = query.filter(SolicitacaoCarreta.data_fim >= primeiro_dia)
    busca = _limpar(busca)
    if busca:
        termo = f"%{busca}%"
        query = query.filter(
            or_(
                SolicitacaoCarreta.municipio_nome.ilike(termo),
                SolicitacaoCarreta.evento.ilike(termo),
                SolicitacaoCarreta.sindicato_nome.ilike(termo),
                SolicitacaoCarreta.responsavel_nome.ilike(termo),
                SolicitacaoCarreta.numero_oficio.ilike(termo),
            )
        )
    return query.all()


def contar_pendentes(db):
    """Quantas solicitações (de qualquer programa) estão com status
    'pendente' agora — usado pro contador no menu do admin, pra quem
    aprova ver de cara que chegou pedido novo, sem precisar abrir a tela."""
    return db.query(SolicitacaoCarreta).filter(SolicitacaoCarreta.status == "pendente").count()


def situacao_atual(db, hoje=None, limite_anteriores=15, tipo_recurso=None):
    """Onde a carreta está AGORA, pra onde vai depois, e onde já esteve —
    pro mapa. Considera solicitações APROVADAS ou já marcadas REALIZADAS —
    pendente é só pedido, ainda não é compromisso confirmado.

    `tipo_recurso`, se informado ("carreta_agro" ou "carreta_saude"), filtra
    só as solicitações daquele programa — usado pra montar um mapa por
    carreta (são DUAS carretas físicas, mesmo dividindo cavalo/motorista).
    Sem esse filtro, os dois programas contam juntos, como antes.

    Devolve um dicionário:
      - "anteriores": últimas solicitações que já terminaram, mais recente
        primeiro (limitado a `limite_anteriores` pra não virar uma lista
        infinita com o tempo).
      - "atual": a solicitação cujo período cobre hoje (ou None se não tem
        nada acontecendo agora — a carreta está "parada"/disponível).
      - "proximas": lista das próximas solicitações que ainda vão começar,
        em ordem cronológica (pra desenhar o trajeto no mapa).
    Se por acaso tiver mais de uma solicitação cobrindo o mesmo dia (não
    deveria acontecer, já que `verificar_conflito` avisa disso), pega a que
    começou primeiro.
    """
    hoje = hoje or date.today()
    atualizar_realizadas_automaticamente(db, hoje=hoje)

    query_base = db.query(SolicitacaoCarreta).filter(
        SolicitacaoCarreta.status.in_(["aprovada", "realizada"])
    )
    if tipo_recurso:
        query_base = query_base.filter(SolicitacaoCarreta.tipo_recurso == tipo_recurso)

    aprovadas = (
        query_base
        .filter(SolicitacaoCarreta.data_fim >= hoje)
        .order_by(SolicitacaoCarreta.data_inicio)
        .all()
    )

    atual = None
    proximas = []
    for solicitacao in aprovadas:
        if solicitacao.data_inicio <= hoje <= solicitacao.data_fim:
            if atual is None:
                atual = solicitacao
        elif solicitacao.data_inicio > hoje:
            proximas.append(solicitacao)

    query_anteriores = db.query(SolicitacaoCarreta).filter(
        SolicitacaoCarreta.status.in_(["aprovada", "realizada"])
    )
    if tipo_recurso:
        query_anteriores = query_anteriores.filter(SolicitacaoCarreta.tipo_recurso == tipo_recurso)

    anteriores = (
        query_anteriores
        .filter(SolicitacaoCarreta.data_fim < hoje)
        .order_by(SolicitacaoCarreta.data_fim.desc())
        .limit(limite_anteriores)
        .all()
    )

    return {"anteriores": anteriores, "atual": atual, "proximas": proximas}


def trajeto_do_mes(db, ano, mes, hoje=None, tipo_recurso=None):
    """Igual a `situacao_atual`, mas olhando só pra um mês específico — pra
    navegação mês a mês no mapa (← mês anterior / próximo mês →), em vez de
    sempre 'últimos 15 + todo o futuro'. Considera qualquer solicitação
    aprovada/realizada cujo período toque esse mês (mesmo que comece ou
    termine em outro mês). Classifica cada uma como já aconteceu / hoje /
    ainda vai acontecer, comparando com a data de hoje de verdade (não com
    o mês em exibição) — então se você olhar o mês atual, o "agora" aparece
    normalmente; olhando um mês passado ou futuro, tudo cai em "anteriores"
    ou "proximas". `tipo_recurso`, se informado, filtra só um dos dois
    programas — mesma lógica de `situacao_atual`."""
    hoje = hoje or date.today()
    atualizar_realizadas_automaticamente(db, hoje=hoje)

    primeiro_dia = date(ano, mes, 1)
    ultimo_dia = date(ano, mes, calendar.monthrange(ano, mes)[1])

    query = db.query(SolicitacaoCarreta).filter(
        SolicitacaoCarreta.status.in_(["aprovada", "realizada"])
    )
    if tipo_recurso:
        query = query.filter(SolicitacaoCarreta.tipo_recurso == tipo_recurso)

    do_mes = (
        query
        .filter(SolicitacaoCarreta.data_inicio <= ultimo_dia)
        .filter(SolicitacaoCarreta.data_fim >= primeiro_dia)
        .order_by(SolicitacaoCarreta.data_inicio)
        .all()
    )

    atual = None
    anteriores = []
    proximas = []
    for solicitacao in do_mes:
        if solicitacao.data_inicio <= hoje <= solicitacao.data_fim:
            if atual is None:
                atual = solicitacao
        elif solicitacao.data_fim < hoje:
            anteriores.append(solicitacao)
        else:
            proximas.append(solicitacao)

    anteriores.sort(key=lambda s: s.data_fim, reverse=True)
    return {"anteriores": anteriores, "atual": atual, "proximas": proximas}


def obter_solicitacao(db, solicitacao_id):
    solicitacao = db.get(SolicitacaoCarreta, solicitacao_id)
    if solicitacao and solicitacao.status == "aprovada" and solicitacao.data_fim < date.today():
        solicitacao.status = "realizada"
        db.commit()
    return solicitacao


def atualizar_solicitacao(db, solicitacao, form, arquivos_upload=None, arquivos_termo=None):
    """Edita os dados de uma solicitação já cadastrada — corrige qualquer
    campo (programa, município, evento, datas, ofício, responsável...),
    com as mesmas validações da criação. `arquivos_upload`/`arquivos_termo`,
    se informados, ADICIONAM novos ofícios/termo em PDF (não mexe nos que
    já estavam anexados). Levanta ValueError se faltar algo obrigatório ou
    a data estiver inválida."""
    tipo_recurso = form.get("tipo_recurso") or solicitacao.tipo_recurso
    municipio_nome = _limpar(form.get("municipio_nome"))
    evento = _titulo(form.get("evento"))
    data_inicio = _texto_para_data(form.get("data_inicio"))
    data_fim = _texto_para_data(form.get("data_fim"))
    responsavel_email = _limpar(form.get("responsavel_email"))
    responsavel_nome = _titulo(form.get("responsavel_nome"))
    responsavel_telefone = _limpar(form.get("responsavel_telefone"))

    if tipo_recurso not in TIPOS_RECURSO:
        raise ValueError("Programa inválido.")
    if not municipio_nome:
        raise ValueError("Informe o município.")
    if not evento:
        raise ValueError("Informe o nome do evento.")
    if not data_inicio or not data_fim:
        raise ValueError("Informe as datas de início e fim do evento.")
    if data_fim < data_inicio:
        raise ValueError("A data final não pode ser antes da data inicial.")
    bloqueios = verificar_bloqueio(db, data_inicio, data_fim)
    if bloqueios:
        motivos = "; ".join(
            f"{b.motivo} ({b.data_inicio.strftime('%d/%m')} a {b.data_fim.strftime('%d/%m')})" for b in bloqueios
        )
        raise ValueError(f"Essas datas estão bloqueadas pelo administrador: {motivos}. Escolha outro período.")
    if not responsavel_nome:
        raise ValueError("Informe o nome do responsável.")
    if not responsavel_telefone:
        raise ValueError("Informe o telefone do responsável.")
    if not responsavel_email:
        raise ValueError("Informe o e-mail do responsável.")

    solicitacao.tipo_recurso = tipo_recurso
    solicitacao.municipio_nome = municipio_nome
    solicitacao.sindicato_nome = _titulo(form.get("sindicato_nome"))
    solicitacao.evento = evento
    solicitacao.data_inicio = data_inicio
    solicitacao.data_fim = data_fim
    solicitacao.numero_oficio = _limpar(form.get("numero_oficio"))
    solicitacao.data_oficio = _texto_para_data(form.get("data_oficio"))
    solicitacao.responsavel_nome = responsavel_nome
    solicitacao.responsavel_telefone = responsavel_telefone
    solicitacao.responsavel_email = responsavel_email
    solicitacao.observacoes = _limpar(form.get("observacoes"))

    _processar_anexos(db, solicitacao.id, arquivos_upload, CATEGORIA_OFICIO)
    _processar_anexos(db, solicitacao.id, arquivos_termo, CATEGORIA_TERMO_COMPROMISSO)

    db.commit()



# Distância (linha reta, km) a partir da qual consideramos que o motorista
# precisa de pelo menos 1 dia livre entre um evento e outro só pra viajar de
# uma cidade pra outra. É uma estimativa (ver distancia_municipios.py) —
# ajuste esse número se ele estiver gerando aviso demais ou de menos na
# prática.
DISTANCIA_MINIMA_KM_PARA_DESLOCAMENTO = 150

# Quantos dias de intervalo esse deslocamento exige, quando a distância
# passa do limite acima.
DIAS_NECESSARIOS_PARA_DESLOCAMENTO = 1


def verificar_conflito(db, data_inicio, data_fim, excluir_id=None):
    """Solicitações (pendentes ou aprovadas) — de QUALQUER programa (Carreta
    do Agro ou Carreta da Saúde) — que já ocupam algum dia dentro do
    período pedido. Os dois programas usam o MESMO motorista/veículo, então
    uma data comprometida num programa bloqueia o outro também: não dá pra
    tratar os calendários como independentes."""
    query = (
        db.query(SolicitacaoCarreta)
        .filter(SolicitacaoCarreta.status.in_(STATUS_CONFIRMADOS))
        .filter(SolicitacaoCarreta.data_inicio <= data_fim)
        .filter(SolicitacaoCarreta.data_fim >= data_inicio)
    )
    if excluir_id:
        query = query.filter(SolicitacaoCarreta.id != excluir_id)
    return query.order_by(SolicitacaoCarreta.data_inicio).all()


def verificar_bloqueio(db, data_inicio, data_fim):
    """Bloqueios de data (feriado, manutenção, motorista de férias...)
    CADASTRADOS PELO ADMIN — nunca criados automaticamente pelo sistema —
    que cobrem algum dia dentro do período pedido. Usado tanto pra travar
    pedido novo (público ou pelo admin) quanto pra mostrar no calendário."""
    return (
        db.query(BloqueioCarreta)
        .filter(BloqueioCarreta.data_inicio <= data_fim)
        .filter(BloqueioCarreta.data_fim >= data_inicio)
        .order_by(BloqueioCarreta.data_inicio)
        .all()
    )


def criar_bloqueio(db, form, criado_por=None):
    """Cadastra um bloqueio de data — SEMPRE uma decisão manual do admin,
    preenchendo esse formulário; o sistema nunca bloqueia data nenhuma
    sozinho. Levanta ValueError se faltar algo ou a data estiver
    inválida."""
    data_inicio = _texto_para_data(form.get("data_inicio"))
    data_fim = _texto_para_data(form.get("data_fim"))
    motivo = _limpar(form.get("motivo"))

    if not data_inicio or not data_fim:
        raise ValueError("Informe as datas de início e fim do bloqueio.")
    if data_fim < data_inicio:
        raise ValueError("A data final não pode ser antes da data inicial.")
    if not motivo:
        raise ValueError("Informe o motivo do bloqueio.")

    bloqueio = BloqueioCarreta(
        data_inicio=data_inicio, data_fim=data_fim, motivo=motivo,
        criado_por=_limpar(criado_por) or "não informado",
    )
    db.add(bloqueio)
    db.commit()
    return bloqueio


def listar_bloqueios(db):
    """Todos os bloqueios cadastrados, mais recente (data de início) primeiro."""
    return db.query(BloqueioCarreta).order_by(BloqueioCarreta.data_inicio.desc()).all()


def obter_bloqueio(db, bloqueio_id):
    return db.get(BloqueioCarreta, bloqueio_id)


def excluir_bloqueio(db, bloqueio):
    """Remove um bloqueio — a data volta a ficar disponível pra pedido
    novo. Também é uma decisão manual do admin."""
    db.delete(bloqueio)
    db.commit()


def verificar_alerta_deslocamento(db, municipio_nome, data_inicio, data_fim, excluir_id=None):
    """Verifica se dá tempo do motorista chegar de/ir pra outro compromisso
    logo antes ou logo depois desse período, considerando a distância entre
    os municípios. Não bloqueia o pedido (a equipe pode decidir manualmente),
    só avisa. Devolve uma lista de dicionários
    {solicitacao, distancia_km, dias_de_intervalo} — vazia se não houver
    nada preocupante (ou se algum município não for reconhecido)."""
    query = db.query(SolicitacaoCarreta).filter(
        SolicitacaoCarreta.status.in_(STATUS_CONFIRMADOS)
    )
    if excluir_id:
        query = query.filter(SolicitacaoCarreta.id != excluir_id)

    avisos = []
    for outra in query.all():
        if outra.data_fim < data_inicio:
            dias_de_intervalo = (data_inicio - outra.data_fim).days - 1
        elif outra.data_inicio > data_fim:
            dias_de_intervalo = (outra.data_inicio - data_fim).days - 1
        else:
            continue  # tem sobreposição de datas — isso já é conflito, não deslocamento

        if dias_de_intervalo >= DIAS_NECESSARIOS_PARA_DESLOCAMENTO:
            continue  # já tem intervalo suficiente, não importa a distância

        distancia = distancia_municipios.distancia_km(municipio_nome, outra.municipio_nome)
        if distancia is None or distancia < DISTANCIA_MINIMA_KM_PARA_DESLOCAMENTO:
            continue

        avisos.append({
            "solicitacao": outra,
            "distancia_km": round(distancia),
            "dias_de_intervalo": dias_de_intervalo,
        })

    avisos.sort(key=lambda a: a["solicitacao"].data_inicio)
    return avisos


def criar_solicitacao(db, form, arquivos_upload, arquivos_termo=None):
    """Cria uma nova solicitação de carreta a partir do formulário público.
    `arquivos_upload` (ofício) e `arquivos_termo` (termo de compromisso
    assinado) são listas de FileStorage do Flask (request.files).
    Levanta ValueError se faltar algo obrigatório ou a data estiver
    inválida (fim antes do início). Retorna a solicitação criada."""
    municipio_nome = _limpar(form.get("municipio_nome"))
    evento = _titulo(form.get("evento"))
    data_inicio = _texto_para_data(form.get("data_inicio"))
    data_fim = _texto_para_data(form.get("data_fim"))
    responsavel_email = _limpar(form.get("responsavel_email"))
    responsavel_nome = _titulo(form.get("responsavel_nome"))
    responsavel_telefone = _limpar(form.get("responsavel_telefone"))
    tipo_recurso = form.get("tipo_recurso") or "carreta_agro"

    if tipo_recurso not in TIPOS_RECURSO:
        raise ValueError("Escolha qual carreta você está solicitando (Agro ou Saúde).")
    if not municipio_nome:
        raise ValueError("Informe o município.")
    if not evento:
        raise ValueError("Informe o nome do evento.")
    if not data_inicio or not data_fim:
        raise ValueError("Informe as datas de início e fim do evento.")
    if data_fim < data_inicio:
        raise ValueError("A data final não pode ser antes da data inicial.")
    bloqueios = verificar_bloqueio(db, data_inicio, data_fim)
    if bloqueios:
        motivos = "; ".join(
            f"{b.motivo} ({b.data_inicio.strftime('%d/%m')} a {b.data_fim.strftime('%d/%m')})" for b in bloqueios
        )
        raise ValueError(f"Essas datas estão bloqueadas pelo administrador: {motivos}. Escolha outro período.")
    if not responsavel_nome:
        raise ValueError("Informe o nome do responsável.")
    if not responsavel_telefone:
        raise ValueError("Informe o telefone do responsável.")
    if not responsavel_email:
        raise ValueError("Informe o e-mail do responsável — é pra onde mandamos a confirmação do pedido.")

    solicitacao = SolicitacaoCarreta(
        tipo_recurso=tipo_recurso,
        municipio_nome=municipio_nome,
        sindicato_nome=_titulo(form.get("sindicato_nome")),
        evento=evento,
        data_inicio=data_inicio,
        data_fim=data_fim,
        numero_oficio=_limpar(form.get("numero_oficio")),
        data_oficio=_texto_para_data(form.get("data_oficio")),
        responsavel_nome=responsavel_nome,
        responsavel_telefone=responsavel_telefone,
        responsavel_email=responsavel_email,
        observacoes=_limpar(form.get("observacoes")),
        status="pendente",
    )
    db.add(solicitacao)
    db.flush()

    _processar_anexos(db, solicitacao.id, arquivos_upload, CATEGORIA_OFICIO)
    _processar_anexos(db, solicitacao.id, arquivos_termo, CATEGORIA_TERMO_COMPROMISSO)

    db.commit()
    return solicitacao


def atualizar_status(db, solicitacao, status, motivo_recusa=None, alterado_por=None):
    if status not in STATUS_LABEL:
        raise ValueError("Situação inválida.")
    solicitacao.status = status
    solicitacao.motivo_recusa = _titulo(motivo_recusa) if status == "recusada" else None
    solicitacao.atualizado_em = datetime.utcnow()
    solicitacao.atualizado_por = _limpar(alterado_por) or "não informado"
    db.commit()


def vizinhos_da_solicitacao(db, solicitacao):
    """Pra ajudar a decidir uma solicitação (aprovar ou não): acha o
    compromisso CONFIRMADO (aprovada/realizada) mais próximo ANTES e o
    mais próximo DEPOIS dela no calendário — independente da distância —
    pra dar uma ideia visual de onde a carreta estaria logo antes e logo
    depois, se essa solicitação for aprovada. Exclui ela mesma da busca
    (relevante se ela já estiver aprovada e você estiver só conferindo)."""
    confirmados = (
        db.query(SolicitacaoCarreta)
        .filter(SolicitacaoCarreta.status.in_(["aprovada", "realizada"]))
        .filter(SolicitacaoCarreta.id != solicitacao.id)
        .order_by(SolicitacaoCarreta.data_inicio)
        .all()
    )
    anterior = None
    proximo = None
    for outra in confirmados:
        if outra.data_fim < solicitacao.data_inicio:
            if anterior is None or outra.data_fim > anterior.data_fim:
                anterior = outra
        elif outra.data_inicio > solicitacao.data_fim:
            if proximo is None or outra.data_inicio < proximo.data_inicio:
                proximo = outra
    return {"anterior": anterior, "proximo": proximo}


def excluir_solicitacao(db, solicitacao):
    for arquivo in solicitacao.arquivos:
        if arquivo.storage_path:
            supabase_storage.excluir_arquivo_privado(arquivo.storage_path)
    db.delete(solicitacao)
    db.commit()


MESES_NOMES = [
    "", "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
    "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro",
]


def matriz_semanas(ano, mes):
    """Lista de semanas (cada uma com 7 dias, começando na segunda-feira) do
    mês — dias fora do mês vêm como None, pra desenhar a grade do
    calendário sem precisar de lógica de datas dentro do template."""
    cal = calendar.Calendar(firstweekday=0)  # 0 = segunda-feira
    semanas = []
    for semana in cal.monthdatescalendar(ano, mes):
        semanas.append([dia if dia.month == mes else None for dia in semana])
    return semanas


def dados_calendario(db, ano, mes, tipo_recurso=None):
    """Monta os dias do mês com as solicitações (pendentes/aprovadas) que
    ocupam cada dia — pra desenhar um calendário mensal simples. Mostra
    solicitações dos DOIS programas (Agro e Saúde): como usam o mesmo
    motorista/veículo, um dia ocupado por qualquer um dos dois bloqueia o
    outro também. O parâmetro `tipo_recurso` é mantido só por compatibilidade
    com as telas que ainda o enviam (ex: pra saber qual aba está selecionada)
    — não filtra mais a disponibilidade em si."""
    primeiro_dia = date(ano, mes, 1)
    ultimo_dia = date(ano, mes, calendar.monthrange(ano, mes)[1])

    atualizar_realizadas_automaticamente(db)
    solicitacoes = verificar_conflito(db, primeiro_dia, ultimo_dia)

    dias = {}
    dia_atual = primeiro_dia
    while dia_atual <= ultimo_dia:
        dias[dia_atual] = []
        dia_atual = date.fromordinal(dia_atual.toordinal() + 1)

    for solicitacao in solicitacoes:
        inicio = max(solicitacao.data_inicio, primeiro_dia)
        fim = min(solicitacao.data_fim, ultimo_dia)
        dia_atual = inicio
        while dia_atual <= fim:
            dias[dia_atual].append(solicitacao)
            dia_atual = date.fromordinal(dia_atual.toordinal() + 1)

    return dias


def dias_bloqueados_do_mes(db, ano, mes):
    """Monta um dict {dia: [BloqueioCarreta, ...]} com os bloqueios que
    cobrem esse mês — separado de `dados_calendario` (que é sobre
    solicitações) porque um BloqueioCarreta não tem os mesmos campos
    (evento, município...) que uma solicitação; misturar os dois na mesma
    lista quebraria os templates que esperam um ou outro."""
    primeiro_dia = date(ano, mes, 1)
    ultimo_dia = date(ano, mes, calendar.monthrange(ano, mes)[1])
    bloqueios = verificar_bloqueio(db, primeiro_dia, ultimo_dia)

    dias = {}
    for bloqueio in bloqueios:
        inicio = max(bloqueio.data_inicio, primeiro_dia)
        fim = min(bloqueio.data_fim, ultimo_dia)
        dia_atual = inicio
        while dia_atual <= fim:
            dias.setdefault(dia_atual, []).append(bloqueio)
            dia_atual = date.fromordinal(dia_atual.toordinal() + 1)
    return dias


def _credenciais_email_carreta():
    """Credenciais PRÓPRIAS do módulo (CARRETA_MAIL_USERNAME/
    CARRETA_MAIL_PASSWORD). Retorna (None, None) se não estiverem
    configuradas — nesse caso, quem chamar não deve nem tentar enviar,
    porque `enviar_email` cairia pra conta de e-mail compartilhada do
    resto do sistema, e a Carreta do Agro precisa de um endereço próprio,
    separado."""
    import os
    usuario = os.environ.get("CARRETA_MAIL_USERNAME")
    senha = os.environ.get("CARRETA_MAIL_PASSWORD")
    if not usuario or not senha:
        return None, None
    return usuario, senha


FUSO_BRASIL = ZoneInfo("America/Fortaleza")  # UTC-3, sem horário de verão


def _data_hora_brasil_str(valor):
    """Formata um datetime (do banco, geralmente em UTC) já convertido pro
    horário do Brasil, pra mostrar nos e-mails (ex: 'Solicitado em ...')."""
    if not valor:
        return "—"
    if valor.tzinfo is None:
        valor = valor.replace(tzinfo=ZoneInfo("UTC"))
    return valor.astimezone(FUSO_BRASIL).strftime("%d/%m/%Y às %H:%M")


def _data_brasil_str(valor):
    """Igual _data_hora_brasil_str, mas só a data (sem hora) — pra colunas
    de listagem onde a hora exata não é necessária."""
    if not valor:
        return "—"
    if valor.tzinfo is None:
        valor = valor.replace(tzinfo=ZoneInfo("UTC"))
    return valor.astimezone(FUSO_BRASIL).strftime("%d/%m/%Y")


def _envelope_email(cor, titulo_faixa, corpo_interno):
    """Deixa todos os e-mails da Carreta com a mesma cara: uma faixa
    colorida no topo (a cor muda conforme o tipo de aviso) e o conteúdo
    dentro de um cartão branco arredondado."""
    return f"""
    <div style="font-family:Georgia,'Times New Roman',serif;max-width:560px;margin:0 auto;">
      <div style="background:{cor};color:#fff;padding:16px 22px;border-radius:12px 12px 0 0;">
        <h2 style="margin:0;font-size:1.15rem;font-family:Georgia,serif;">{titulo_faixa}</h2>
      </div>
      <div style="border:1px solid #E3DCC8;border-top:none;border-radius:0 0 12px 12px;
                  padding:22px;background:#fff;font-family:Arial,Helvetica,sans-serif;
                  font-size:.95rem;color:#3A3226;line-height:1.5;">
        {corpo_interno}
      </div>
      <p style="color:#948A76;font-size:.72rem;margin-top:14px;text-align:center;
                font-family:Arial,Helvetica,sans-serif;">
        Sistema de Solicitação de Carreta — SENAR-CE
      </p>
    </div>
    """


def _tabela_detalhes_email(solicitacao, linhas_extras=None):
    """Tabelinha padrão com os dados do pedido, usada em todos os e-mails."""
    linhas = [
        ("Programa", TIPOS_RECURSO.get(solicitacao.tipo_recurso, solicitacao.tipo_recurso)),
        ("Evento", solicitacao.evento),
        ("Município", solicitacao.municipio_nome),
        ("Período", f"{solicitacao.data_inicio.strftime('%d/%m/%Y')} a {solicitacao.data_fim.strftime('%d/%m/%Y')}"),
        ("Solicitado em", _data_hora_brasil_str(solicitacao.criado_em)),
    ]
    if linhas_extras:
        linhas += linhas_extras
    linhas_html = "".join(
        f'<tr><td style="padding:7px 12px 7px 0;color:#8A8168;white-space:nowrap;vertical-align:top;">{rotulo}</td>'
        f'<td style="padding:7px 0;font-weight:700;">{valor}</td></tr>'
        for rotulo, valor in linhas
    )
    return f'<table style="width:100%;border-collapse:collapse;margin:14px 0;">{linhas_html}</table>'


def enviar_confirmacao_solicitante(solicitacao):
    """Manda pro PRÓPRIO solicitante uma confirmação de que o pedido foi
    recebido — usa a mesma conta de e-mail própria do módulo
    (CARRETA_MAIL_USERNAME/CARRETA_MAIL_PASSWORD)."""
    usuario, senha = _credenciais_email_carreta()
    if not usuario or not solicitacao.responsavel_email:
        return False

    corpo_interno = f"""
    <p>Olá, {solicitacao.responsavel_nome or ''}.</p>
    <p>Recebemos sua solicitação. Confira os dados abaixo:</p>
    {_tabela_detalhes_email(solicitacao)}
    <p>Situação atual: <b>Pendente</b>. Assim que o pedido for analisado, você recebe um novo
    e-mail avisando se foi aprovado ou não.</p>
    """
    corpo = _envelope_email("#B5851A", "📨 Solicitação recebida", corpo_interno)
    return enviar_email(
        solicitacao.responsavel_email, f"Recebemos sua solicitação — {solicitacao.evento}", corpo,
        usuario=usuario, senha=senha,
    )


def enviar_notificacao_status(solicitacao):
    """Avisa o solicitante quando o pedido for aprovado ou recusado."""
    usuario, senha = _credenciais_email_carreta()
    if not usuario or not solicitacao.responsavel_email or solicitacao.status not in ("aprovada", "recusada"):
        return False

    if solicitacao.status == "aprovada":
        assunto = f"Solicitação aprovada — {solicitacao.evento}"
        corpo_interno = f"""
        <p>Olá, {solicitacao.responsavel_nome or ''}.</p>
        <p>Boa notícia: sua solicitação foi <b>aprovada</b>! Confira os dados abaixo:</p>
        {_tabela_detalhes_email(solicitacao)}
        """
        corpo = _envelope_email("#3E8E52", "✅ Solicitação aprovada", corpo_interno)
    else:
        linhas_extras = [("Motivo do cancelamento", solicitacao.motivo_recusa)] if solicitacao.motivo_recusa else None
        assunto = f"Solicitação cancelada — {solicitacao.evento}"
        corpo_interno = f"""
        <p>Olá, {solicitacao.responsavel_nome or ''}.</p>
        <p>Sua solicitação foi <b>cancelada</b>. Confira os dados abaixo:</p>
        {_tabela_detalhes_email(solicitacao, linhas_extras)}
        """
        corpo = _envelope_email("#A6412B", "✖ Solicitação cancelada", corpo_interno)

    return enviar_email(
        solicitacao.responsavel_email, assunto, corpo,
        usuario=usuario, senha=senha,
    )


def enviar_notificacao_nova_solicitacao(solicitacao, destinatario):
    """Manda um e-mail avisando que chegou uma solicitação nova — usa uma
    conta de e-mail PRÓPRIA do módulo (CARRETA_MAIL_USERNAME/
    CARRETA_MAIL_PASSWORD), separada da conta usada pro resto do sistema.
    Se essas variáveis não estiverem configuradas ainda, simplesmente não
    envia (retorna False) sem quebrar o fluxo de quem preencheu o formulário.

    O e-mail vem com um link direto pra tela de detalhe (não precisa ficar
    procurando na listagem) e com os ofícios em PDF anexados de verdade —
    assim o admin já visualiza a demanda direto no e-mail, mesmo sem abrir
    o sistema."""
    usuario, senha = _credenciais_email_carreta()
    if not usuario:
        return False

    try:
        link = url_for("admin_carreta_detalhe", solicitacao_id=solicitacao.id, _external=True)
    except RuntimeError:
        # Fora de um request do Flask (ex: chamado de um script/console) não
        # dá pra montar link externo — o e-mail ainda sai, só sem o link.
        link = None

    anexos = []
    for arquivo in solicitacao.arquivos:
        conteudo = None
        if arquivo.storage_path:
            try:
                conteudo = supabase_storage.baixar_bytes_privado(arquivo.storage_path)
            except Exception as erro:  # noqa: BLE001 — não vale travar o e-mail por causa de 1 anexo
                print(f"[carreta] aviso: não consegui anexar '{arquivo.nome_arquivo}' no e-mail: {erro}")
        elif arquivo.conteudo:
            conteudo = arquivo.conteudo
        if conteudo:
            anexos.append({"nome": arquivo.nome_arquivo, "conteudo": conteudo, "tipo_mime": "application/pdf"})

    corpo_interno = f"""
    <p>Chegou uma nova solicitação, aguardando análise:</p>
    {_tabela_detalhes_email(solicitacao, [("Responsável", solicitacao.responsavel_nome or "não informado")])}
    {f'<p style="margin-top:18px;"><a href="{link}" style="display:inline-block;background:#B5851A;color:#fff;padding:10px 18px;border-radius:8px;text-decoration:none;font-weight:700;">Ver e aprovar essa solicitação →</a></p>' if link else ''}
    <p style="margin-top:16px;color:#8A8168;font-size:.85rem;">
      {'📎 Os documentos anexados (ofício e termo de compromisso) também estão neste e-mail.' if anexos else 'Nenhum documento foi anexado a esse pedido.'}
    </p>
    """
    corpo = _envelope_email("#B5851A", "🔔 Nova solicitação de carreta", corpo_interno)
    return enviar_email(
        destinatario, f"Nova solicitação de {TIPOS_RECURSO.get(solicitacao.tipo_recurso, solicitacao.tipo_recurso)} — {solicitacao.municipio_nome}", corpo,
        usuario=usuario, senha=senha, anexos=anexos,
    )
