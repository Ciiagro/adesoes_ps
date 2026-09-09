import os
import secrets
from datetime import datetime
from io import BytesIO

from flask import (
    Flask, render_template, request, redirect, url_for, session, flash, abort,
    Response, send_file
)
from sqlalchemy import func
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

from database import SessionLocal, init_db
from models import (
    Programa, Municipio, Coordenador, Adesao, Escola, AdesaoEscola, Serie,
    Matricula, Assinatura, MunicipioSindicato, SindicatoRural
)
from municipios_ce import MUNICIPIOS_CE
from ceps_municipios import CEP_POR_MUNICIPIO
from email_utils import enviar_email
from pdf_generator import gerar_pdf_adesao_bytes
import sindicatos_admin as sind
import carreta as carr
import supabase_storage
import distancia_municipios

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "dev")
app.jinja_env.filters["data_brasil"] = carr._data_hora_brasil_str


@app.teardown_appcontext
def _fechar_sessao_db(exception=None):
    """Fecha (e devolve) a conexão com o banco ao final de CADA requisição.
    Sem isso, como o engine usa NullPool, toda chamada a SessionLocal() abre
    uma conexão nova que nunca era liberada — em pouco tempo estoura o limite
    de conexões simultâneas do pooler do Supabase (erro "max clients reached
    in session mode"). Rodar isso aqui garante que a conexão é sempre
    devolvida, mesmo quando a rota termina com erro."""
    SessionLocal.remove()


PROGRAMAS_SLUG = {
    "agrinho": "AGRINHO",
    "valores": "VALORES",
}

NOME_PROGRAMA = {
    "AGRINHO": "Programa Agrinho",
    "VALORES": "Projeto Valores",
}

# Papel -> rótulo amigável, usado nos e-mails e na tela de assinatura
PAPEIS = {
    "prefeito": "Prefeito(a) Municipal",
    "secretario": "Secretário(a) de Educação",
    "coordenador": "Coordenador(a) do Projeto",
}

# Cores do tema usadas na ficha em PDF, por programa
CORES_PDF = {
    "AGRINHO": {"tema": "#3F6F4C", "escura": "#2C5137"},
    "VALORES": {"tema": "#C98A2B", "escura": "#8C5F1D"},
}


# ---------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------

def get_programa(db, slug):
    nome = PROGRAMAS_SLUG.get(slug)
    if not nome:
        abort(404)
    programa = db.query(Programa).filter_by(nome=nome).first()
    if not programa:
        abort(404, description="Rode o schema.sql no Supabase antes de usar o sistema.")
    return programa


def _gerar_pdf_adesao(db, adesao):
    """Monta a ficha de adesão (com termo e assinaturas) e devolve os bytes do PDF.
    Usa reportlab (Python puro) — não depende de bibliotecas de sistema, então
    funciona em ambientes serverless como o Vercel."""
    series = (
        db.query(Serie)
        .filter_by(programa_id=adesao.programa_id)
        .order_by(Serie.ordem)
        .all()
    )

    totais_series = {s.id: 0 for s in series}
    for ae in adesao.escolas:
        for m in ae.matriculas:
            totais_series[m.serie_id] = totais_series.get(m.serie_id, 0) + (m.quantidade or 0)

    cores = CORES_PDF.get(adesao.programa.nome, CORES_PDF["AGRINHO"])
    assinaturas_por_papel = {a.papel: a for a in adesao.assinaturas}
    nome_programa = NOME_PROGRAMA.get(adesao.programa.nome, adesao.programa.nome)

    return gerar_pdf_adesao_bytes(
        adesao=adesao,
        series=series,
        totais_series=totais_series,
        gerado_em=datetime.now(),
        cor_tema=cores["tema"],
        cor_tema_escura=cores["escura"],
        assinaturas_por_papel=assinaturas_por_papel,
        nome_programa=nome_programa,
    )


def _serializar_adesao_para_form(adesao):
    """Converte uma Adesao já salva em um dicionário no mesmo formato dos campos
    do formulário, para pré-preencher a tela quando alguém retoma um rascunho."""
    m = adesao.municipio
    c = adesao.coordenador
    escolas = []
    for ae in adesao.escolas:
        e = ae.escola
        escolas.append({
            "nome": e.nome,
            "endereco": e.endereco,
            "telefone": e.contato_telefone,
            "email": e.contato_email,
            "professores": ae.quantidade_professores,
            "gestor1": e.gestor1_nome,
            "gestor2": e.gestor2_nome,
            "latitude": e.latitude,
            "longitude": e.longitude,
            "series": {str(mat.serie_id): mat.quantidade for mat in ae.matriculas},
        })
    return {
        "ano": adesao.ano,
        "municipio_nome": m.nome,
        "prefeito_nome": m.prefeito_nome,
        "prefeito_rg": m.prefeito_rg,
        "prefeito_cpf": m.prefeito_cpf,
        "prefeitura_endereco": m.prefeitura_endereco,
        "prefeitura_cep": m.prefeitura_cep,
        "prefeitura_telefone": m.prefeitura_telefone,
        "prefeitura_email": m.prefeitura_email,
        "secretario_nome": m.secretario_nome,
        "secretaria_endereco": m.secretaria_endereco,
        "secretaria_cep": m.secretaria_cep,
        "secretaria_telefone": m.secretaria_telefone,
        "secretaria_email": m.secretaria_email,
        "coordenador_nome": c.nome if c else None,
        "coordenador_rg": c.rg if c else None,
        "coordenador_cpf": c.cpf if c else None,
        "coordenador_telefone1": c.telefone1 if c else None,
        "coordenador_telefone2": c.telefone2 if c else None,
        "coordenador_email": c.email if c else None,
        "total_escolas": adesao.total_escolas_informado,
        "total_professores": adesao.total_professores_informado,
        "responsavel_preenchimento": adesao.responsavel_preenchimento,
        "escolas": escolas,
    }


def _texto_para_float(valor):
    """Converte o texto digitado (latitude/longitude) para float, aceitando
    vírgula ou ponto como separador decimal. Ex.: '-32,779' ou '-32.779'."""
    if valor is None:
        return None
    valor = valor.strip()
    if not valor:
        return None
    try:
        return float(valor.replace(",", "."))
    except ValueError:
        return None


def _montar_dados_do_formulario(form, series):
    """Reconstrói o dicionário 'dados' (mesmo formato usado para pré-preencher
    o formulário) diretamente do que a pessoa acabou de digitar — usado pra
    devolver a tela sem perder nada quando alguma validação falha antes de
    gravar no banco."""
    escolas = []
    idx = 0
    while f"escola_nome_{idx}" in form:
        nome_escola = form.get(f"escola_nome_{idx}", "").strip()
        if nome_escola:
            escolas.append({
                "nome": nome_escola,
                "endereco": form.get(f"escola_endereco_{idx}"),
                "telefone": form.get(f"escola_telefone_{idx}"),
                "email": form.get(f"escola_email_{idx}"),
                "professores": form.get(f"escola_professores_{idx}", type=int),
                "gestor1": form.get(f"escola_gestor1_{idx}"),
                "gestor2": form.get(f"escola_gestor2_{idx}"),
                "latitude": form.get(f"escola_latitude_{idx}"),
                "longitude": form.get(f"escola_longitude_{idx}"),
                "series": {
                    str(serie.id): form.get(f"serie_{serie.id}_{idx}", type=int) or 0
                    for serie in series
                },
            })
        idx += 1

    return {
        "ano": form.get("ano", type=int),
        "municipio_nome": form.get("municipio_nome", "").strip().upper(),
        "prefeito_nome": form.get("prefeito_nome"),
        "prefeito_rg": form.get("prefeito_rg"),
        "prefeito_cpf": form.get("prefeito_cpf"),
        "prefeitura_endereco": form.get("prefeitura_endereco"),
        "prefeitura_cep": form.get("prefeitura_cep"),
        "prefeitura_telefone": form.get("prefeitura_telefone"),
        "prefeitura_email": form.get("prefeitura_email"),
        "secretario_nome": form.get("secretario_nome"),
        "secretaria_endereco": form.get("secretaria_endereco"),
        "secretaria_cep": form.get("secretaria_cep"),
        "secretaria_telefone": form.get("secretaria_telefone"),
        "secretaria_email": form.get("secretaria_email"),
        "coordenador_nome": form.get("coordenador_nome"),
        "coordenador_rg": form.get("coordenador_rg"),
        "coordenador_cpf": form.get("coordenador_cpf"),
        "coordenador_telefone1": form.get("coordenador_telefone1"),
        "coordenador_telefone2": form.get("coordenador_telefone2"),
        "coordenador_email": form.get("coordenador_email"),
        "total_escolas": form.get("total_escolas", type=int),
        "total_professores": form.get("total_professores", type=int),
        "responsavel_preenchimento": form.get("responsavel_preenchimento"),
        "escolas": escolas,
    }


def _papeis_assinatura(adesao):
    """Lista (papel, nome, email) dos 3 responsáveis que precisam assinar,
    a partir dos dados já salvos do município/coordenador."""
    m = adesao.municipio
    c = adesao.coordenador
    return [
        ("prefeito", m.prefeito_nome, m.prefeitura_email),
        ("secretario", m.secretario_nome, m.secretaria_email),
        ("coordenador", c.nome if c else None, c.email if c else None),
    ]


def _enviar_emails_assinatura(db, adesao):
    """(Re)cria os registros de assinatura dos 3 responsáveis e envia o link de
    confirmação por e-mail para cada um. Retorna lista de avisos (e-mails que
    não puderam ser enviados, para mostrar na tela pra quem preencheu)."""
    m = adesao.municipio
    dados_papeis = _papeis_assinatura(adesao)

    avisos = []
    nome_programa = NOME_PROGRAMA.get(adesao.programa.nome, adesao.programa.nome)

    for papel, nome, email in dados_papeis:
        assinatura = next((a for a in adesao.assinaturas if a.papel == papel), None)
        if not assinatura:
            assinatura = Assinatura(adesao_id=adesao.id, papel=papel, token=secrets.token_urlsafe(24))
        assinatura.nome = nome
        assinatura.email = email
        assinatura.assinado_em = None  # qualquer reenvio invalida a assinatura anterior
        assinatura.ip = None
        assinatura.user_agent = None
        db.add(assinatura)

        link = url_for("assinar_adesao", token=assinatura.token, _external=True)
        if email:
            corpo = f"""
            <p>Olá, {nome or ''}.</p>
            <p>O município de <strong>{m.nome}</strong> enviou a adesão ao
            <strong>{nome_programa} {adesao.ano}</strong> e sua confirmação, como
            <strong>{PAPEIS.get(papel, papel)}</strong>, é uma das etapas finais.</p>
            <p>Clique no link abaixo para revisar os dados da sua função e confirmar
            sua assinatura eletrônica — não é necessário criar conta nem acessar
            nenhum sistema:</p>
            <p><a href="{link}">{link}</a></p>
            <p>Se você não solicitou isso, apenas ignore este e-mail.</p>
            """
            enviado = enviar_email(
                email,
                f"Confirme sua assinatura — {nome_programa} {adesao.ano} ({m.nome})",
                corpo,
            )
            if not enviado:
                avisos.append((papel, nome, email, link))
        else:
            avisos.append((papel, nome, email, link))

    return avisos


# ---------------------------------------------------------------
# Páginas públicas
# ---------------------------------------------------------------

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/adesao/<programa_slug>", methods=["GET", "POST"])
def adesao_form(programa_slug):
    db = SessionLocal()
    programa = get_programa(db, programa_slug)
    series = db.query(Serie).filter_by(programa_id=programa.id).order_by(Serie.ordem).all()

    if request.method == "POST":
        acao = request.form.get("acao", "enviar")  # 'rascunho' ou 'enviar'
        ano = request.form.get("ano", type=int) or datetime.now().year

        # --- Validação: soma de professores por escola deve bater com o
        # total informado (só bloqueia no envio final; rascunho pode ficar
        # incompleto, já que é só um "salvar e continuar depois") ---
        soma_professores = 0
        escolas_preenchidas = 0
        idx_check = 0
        while f"escola_nome_{idx_check}" in request.form:
            if request.form.get(f"escola_nome_{idx_check}", "").strip():
                escolas_preenchidas += 1
                soma_professores += request.form.get(f"escola_professores_{idx_check}", type=int) or 0
            idx_check += 1
        total_professores_informado = request.form.get("total_professores", type=int)
        total_escolas_informado = request.form.get("total_escolas", type=int)

        erros_validacao = []
        if (
            total_professores_informado is not None
            and soma_professores != total_professores_informado
        ):
            erros_validacao.append(
                f"O total de professores informado ({total_professores_informado}) não bate com a "
                f"soma dos professores por escola ({soma_professores}). Confira os números "
                f"(ou o total geral, ou a quantidade em alguma escola)."
            )
        if (
            total_escolas_informado is not None
            and escolas_preenchidas != total_escolas_informado
        ):
            erros_validacao.append(
                f"Você declarou {total_escolas_informado} escola(s), mas preencheu os dados de "
                f"{escolas_preenchidas}. Preencha (ou remova) escolas até os números baterem."
            )

        if acao == "enviar" and erros_validacao:
            for erro in erros_validacao:
                flash(erro)
            dados_reenvio = _montar_dados_do_formulario(request.form, series)
            series_data = [{"id": s.id, "nome": s.nome} for s in series]
            ano_atual = datetime.now().year
            anos_disponiveis = list(range(ano_atual - 1, ano_atual + 3))
            ano_travado = dados_reenvio.get("ano") or ano_atual
            return render_template(
                "adesao_form.html",
                programa=programa,
                series=series_data,
                ano_atual=ano_atual,
                anos_disponiveis=anos_disponiveis,
                ano_travado=ano_travado,
                municipios=MUNICIPIOS_CE,
                ceps_por_municipio=CEP_POR_MUNICIPIO,
                dados=dados_reenvio,
                continuar_token=request.form.get("continuar_token", ""),
                municipio_selecionado=dados_reenvio.get("municipio_nome"),
            )

        # --- Município ---
        municipio_nome = request.form["municipio_nome"].strip().upper()
        municipio = db.query(Municipio).filter_by(nome=municipio_nome, uf="CE").first()
        if not municipio:
            municipio = Municipio(nome=municipio_nome, uf="CE")

        municipio.prefeito_nome = request.form.get("prefeito_nome")
        municipio.prefeito_rg = request.form.get("prefeito_rg")
        municipio.prefeito_cpf = request.form.get("prefeito_cpf")
        municipio.prefeitura_endereco = request.form.get("prefeitura_endereco")
        municipio.prefeitura_cep = request.form.get("prefeitura_cep")
        municipio.prefeitura_telefone = request.form.get("prefeitura_telefone")
        municipio.prefeitura_email = request.form.get("prefeitura_email")
        municipio.secretario_nome = request.form.get("secretario_nome")
        municipio.secretaria_endereco = request.form.get("secretaria_endereco")
        municipio.secretaria_cep = request.form.get("secretaria_cep")
        municipio.secretaria_telefone = request.form.get("secretaria_telefone")
        municipio.secretaria_email = request.form.get("secretaria_email")
        db.add(municipio)
        db.flush()  # garante municipio.id disponível

        # --- Coordenador (um novo registro a cada envio, histórico fica preservado) ---
        coordenador = Coordenador(
            municipio_id=municipio.id,
            nome=request.form.get("coordenador_nome"),
            rg=request.form.get("coordenador_rg"),
            cpf=request.form.get("coordenador_cpf"),
            telefone1=request.form.get("coordenador_telefone1"),
            telefone2=request.form.get("coordenador_telefone2"),
            email=request.form.get("coordenador_email"),
        )
        db.add(coordenador)
        db.flush()

        # --- Adesão (programa + município + ano) ---
        adesao = (
            db.query(Adesao)
            .filter_by(programa_id=programa.id, municipio_id=municipio.id, ano=ano)
            .first()
        )
        if not adesao:
            adesao = Adesao(programa_id=programa.id, municipio_id=municipio.id, ano=ano)
        adesao.coordenador_id = coordenador.id
        adesao.responsavel_preenchimento = request.form.get("responsavel_preenchimento")
        adesao.total_escolas_informado = request.form.get("total_escolas", type=int)
        adesao.total_professores_informado = request.form.get("total_professores", type=int)
        db.add(adesao)
        db.flush()

        # --- Escolas dinâmicas (escola_nome_0, escola_nome_1, ...) ---
        idx = 0
        while f"escola_nome_{idx}" in request.form:
            nome_escola = request.form.get(f"escola_nome_{idx}", "").strip()
            if nome_escola:
                escola = (
                    db.query(Escola)
                    .filter_by(municipio_id=municipio.id, nome=nome_escola)
                    .first()
                )
                if not escola:
                    escola = Escola(municipio_id=municipio.id, nome=nome_escola)
                escola.endereco = request.form.get(f"escola_endereco_{idx}")
                escola.contato_telefone = request.form.get(f"escola_telefone_{idx}")
                escola.contato_email = request.form.get(f"escola_email_{idx}")
                escola.gestor1_nome = request.form.get(f"escola_gestor1_{idx}")
                escola.gestor2_nome = request.form.get(f"escola_gestor2_{idx}")
                escola.latitude = _texto_para_float(request.form.get(f"escola_latitude_{idx}"))
                escola.longitude = _texto_para_float(request.form.get(f"escola_longitude_{idx}"))
                db.add(escola)
                db.flush()

                adesao_escola = (
                    db.query(AdesaoEscola)
                    .filter_by(adesao_id=adesao.id, escola_id=escola.id)
                    .first()
                )
                if not adesao_escola:
                    adesao_escola = AdesaoEscola(adesao_id=adesao.id, escola_id=escola.id)
                adesao_escola.quantidade_professores = request.form.get(
                    f"escola_professores_{idx}", type=int
                )
                db.add(adesao_escola)
                db.flush()

                for serie in series:
                    qtd = request.form.get(f"serie_{serie.id}_{idx}", type=int) or 0
                    matricula = (
                        db.query(Matricula)
                        .filter_by(adesao_escola_id=adesao_escola.id, serie_id=serie.id)
                        .first()
                    )
                    if not matricula:
                        matricula = Matricula(
                            adesao_escola_id=adesao_escola.id, serie_id=serie.id
                        )
                    matricula.quantidade = qtd
                    db.add(matricula)
            idx += 1

        # --- Rascunho x Envio para assinaturas ---
        if acao == "rascunho":
            adesao.status = "rascunho"
            if not adesao.rascunho_token:
                adesao.rascunho_token = secrets.token_urlsafe(24)
            db.add(adesao)
            db.commit()

            link_rascunho = url_for(
                "adesao_form", programa_slug=programa_slug,
                continuar=adesao.rascunho_token, _external=True
            )
            destinatario = coordenador.email or municipio.prefeitura_email
            email_enviado = False
            if destinatario:
                nome_programa = NOME_PROGRAMA.get(programa.nome, programa.nome)
                corpo = f"""
                <p>Olá!</p>
                <p>O preenchimento da adesão ao <strong>{nome_programa} {ano}</strong>
                do município de <strong>{municipio.nome}</strong> foi salvo como
                rascunho. Use o link abaixo para continuar de onde parou, quando
                quiser — não é necessário fazer login:</p>
                <p><a href="{link_rascunho}">{link_rascunho}</a></p>
                """
                email_enviado = enviar_email(
                    destinatario,
                    f"Rascunho salvo — Adesão {nome_programa} {ano} ({municipio.nome})",
                    corpo,
                )
            return render_template(
                "rascunho_salvo.html",
                link=link_rascunho,
                destinatario=destinatario,
                email_enviado=email_enviado,
            )

        # acao == "enviar": os dados são salvos, mas os e-mails de assinatura
        # só são disparados depois que o responsável confirmar, na tela de
        # prévia, que a ficha ficou do jeito certo.
        adesao.status = "aguardando_confirmacao"
        if not adesao.rascunho_token:
            adesao.rascunho_token = secrets.token_urlsafe(24)
        db.add(adesao)
        db.commit()

        destinatarios = [
            {"papel": papel, "papel_label": PAPEIS.get(papel, papel), "nome": nome, "email": email}
            for papel, nome, email in _papeis_assinatura(adesao)
        ]
        link_editar = url_for(
            "adesao_form", programa_slug=programa_slug,
            continuar=adesao.rascunho_token,
        )
        return render_template(
            "confirmar_envio.html",
            adesao=adesao,
            nome_programa=NOME_PROGRAMA.get(programa.nome, programa.nome),
            destinatarios=destinatarios,
            link_editar=link_editar,
        )

    # --- GET: formulário novo ou retomando um rascunho ---
    dados = None
    token = request.args.get("continuar")
    if token:
        adesao_existente = (
            db.query(Adesao)
            .filter_by(rascunho_token=token, programa_id=programa.id)
            .first()
        )
        if adesao_existente:
            dados = _serializar_adesao_para_form(adesao_existente)

    series_data = [{"id": s.id, "nome": s.nome} for s in series]
    ano_atual = datetime.now().year
    # permite lançar adesão do ano anterior (atraso) até 2 anos à frente
    anos_disponiveis = list(range(ano_atual - 1, ano_atual + 3))
    # o formulário público não deixa mais escolher o ano — trava sempre no
    # ano atual, a não ser que seja um rascunho antigo sendo retomado
    # (nesse caso mantém o ano em que ele foi criado).
    ano_travado = (dados.get("ano") if dados else None) or ano_atual
    return render_template(
        "adesao_form.html",
        programa=programa,
        series=series_data,
        ano_atual=ano_atual,
        anos_disponiveis=anos_disponiveis,
        ano_travado=ano_travado,
        municipios=MUNICIPIOS_CE,
        ceps_por_municipio=CEP_POR_MUNICIPIO,
        dados=dados,
        continuar_token=token or "",
        municipio_selecionado=(dados.get("municipio_nome") if dados else None),
    )


@app.route("/adesao/confirmar/<int:adesao_id>", methods=["POST"])
def adesao_confirmar_envio(adesao_id):
    """Chamada depois que o responsável já viu a prévia da ficha (rota
    confirmar_envio.html) e confirmou que está tudo certo. Só a partir daqui
    os e-mails de assinatura são realmente disparados."""
    db = SessionLocal()
    adesao = db.get(Adesao, adesao_id)
    if not adesao:
        abort(404)

    adesao.status = "aguardando_assinaturas"
    adesao.enviada_em = datetime.now()
    db.add(adesao)
    db.commit()

    _enviar_emails_assinatura(db, adesao)
    db.commit()

    return redirect(url_for("sucesso", adesao_id=adesao.id))


@app.route("/sucesso")
def sucesso():
    adesao_id = request.args.get("adesao_id", type=int)
    return render_template("sucesso.html", adesao_id=adesao_id)


@app.route("/assinar/<token>", methods=["GET", "POST"])
def assinar_adesao(token):
    """Página pública (sem login) onde cada responsável confirma/assina
    eletronicamente apenas a sua parte da adesão."""
    db = SessionLocal()
    assinatura = db.query(Assinatura).filter_by(token=token).first()
    if not assinatura:
        abort(404)

    adesao = assinatura.adesao
    m = adesao.municipio
    c = adesao.coordenador

    dados_papel = {
        "prefeito": {
            "Nome": m.prefeito_nome, "RG": m.prefeito_rg, "CPF": m.prefeito_cpf,
            "Endereço da Prefeitura": m.prefeitura_endereco,
            "Telefone": m.prefeitura_telefone, "E-mail": m.prefeitura_email,
        },
        "secretario": {
            "Nome": m.secretario_nome, "CPF": m.secretaria_cep,
            "Endereço": m.secretaria_endereco,
            "Telefone": m.secretaria_telefone, "E-mail": m.secretaria_email,
        },
        "coordenador": {
            "Nome": c.nome if c else None, "RG": c.rg if c else None,
            "CPF": c.cpf if c else None,
            "Telefone(s)": ", ".join(filter(None, [c.telefone1 if c else None, c.telefone2 if c else None])),
            "E-mail": c.email if c else None,
        },
    }.get(assinatura.papel, {})

    if request.method == "POST":
        if not assinatura.assinado_em:
            nome_digitado = request.form.get("confirmacao_nome", "").strip()
            assinatura.assinado_em = datetime.now()
            assinatura.ip = request.headers.get("X-Forwarded-For", request.remote_addr)
            assinatura.user_agent = request.headers.get("User-Agent", "")[:255]
            if nome_digitado:
                assinatura.nome = nome_digitado
            db.add(assinatura)
            db.commit()

            # se todos os 3 papéis já assinaram, conclui a adesão
            todas = db.query(Assinatura).filter_by(adesao_id=adesao.id).all()
            if todas and all(a.assinado_em for a in todas):
                adesao.status = "concluida"
                adesao.concluida_em = datetime.now()
                db.add(adesao)
                db.commit()

        return render_template(
            "assinatura_confirmada.html",
            adesao=adesao,
            papel_label=PAPEIS.get(assinatura.papel, assinatura.papel),
            nome_programa=NOME_PROGRAMA.get(adesao.programa.nome, adesao.programa.nome),
        )

    return render_template(
        "assinar.html",
        assinatura=assinatura,
        adesao=adesao,
        papel_label=PAPEIS.get(assinatura.papel, assinatura.papel),
        dados_papel=dados_papel,
        nome_programa=NOME_PROGRAMA.get(adesao.programa.nome, adesao.programa.nome),
    )


@app.route("/adesao/pdf/<int:adesao_id>")
def adesao_pdf(adesao_id):
    """Gera (sob demanda) a ficha de adesão em PDF, pronta para impressão e
    assinatura do prefeito, secretário(a), coordenador(a) e gestores das escolas.
    Responsáveis que já assinaram eletronicamente aparecem como tal no PDF."""
    db = SessionLocal()
    adesao = db.get(Adesao, adesao_id)
    if not adesao:
        abort(404)

    pdf_bytes = _gerar_pdf_adesao(db, adesao)

    nome_arquivo = f"Adesao_{adesao.municipio.nome}_{adesao.programa.nome}_{adesao.ano}.pdf"
    nome_arquivo = nome_arquivo.replace(" ", "_")

    return Response(
        pdf_bytes,
        mimetype="application/pdf",
        headers={"Content-Disposition": f"inline; filename=\"{nome_arquivo}\""},
    )


# ---------------------------------------------------------------
# Área Administrativa
# ---------------------------------------------------------------

def admin_logado():
    return session.get("admin_ok", False)


def gestao_sindicato_logado():
    return session.get("gestao_sindicato_ok", False)


def pode_ver_sindicatos():
    """Admin/Comissão (visão só de leitura, colunas básicas) OU a pessoa
    responsável pela gestão dos sindicatos (visão completa, com edição)."""
    return admin_logado() or gestao_sindicato_logado()


def carreta_admin_logado():
    return session.get("carreta_admin_ok", False)


def pode_ver_carreta():
    """Admin (visão geral) OU a pessoa responsável pela Carreta do Agro
    (login próprio, senha separada)."""
    return admin_logado() or carreta_admin_logado()


def acesso_adesoes_obrigatorio():
    """Usado nas páginas de Adesões/Escolas/Estatísticas: só quem logou como
    Admin/Comissão entra. Quem logou só como gestão de sindicatos é mandado
    de volta para o módulo dele, com uma explicação."""
    if admin_logado():
        return None
    if gestao_sindicato_logado():
        flash("Esse login tem acesso apenas ao módulo de Sindicatos x Municípios.")
        return redirect(url_for("admin_sindicatos"))
    return redirect(url_for("admin_login"))


def acesso_gestao_sindicatos_obrigatorio():
    """Usado em editar/histórico/exportar de sindicatos: só quem logou com a
    senha de gestão de sindicatos pode alterar ou baixar a planilha. O
    Admin/Comissão vê a lista (básica), mas não edita nem exporta daqui."""
    if gestao_sindicato_logado():
        return None
    if admin_logado():
        flash("Esse login só visualiza os dados de sindicatos. Peça para quem faz a gestão dos sindicatos.")
        return redirect(url_for("admin_sindicatos"))
    return redirect(url_for("admin_login"))


def _filtros_programa_ano_sessao():
    """Lê os filtros de Programa e Ano: se vieram agora (formulário
    enviado nesta requisição), usa e grava na sessão; se a pessoa só
    navegou por um link do menu, usa o que ficou guardado da última vez.
    Assim, uma vez escolhido, o filtro continua valendo em todas as telas
    (Adesões, Planilha, Exportar, Mapa de Escolas) até ser trocado de novo.
    """
    if "programa" in request.args:
        programa_slug = request.args.get("programa") or ""
        session["admin_programa"] = programa_slug
    else:
        programa_slug = session.get("admin_programa", "")

    if "ano" in request.args:
        ano = request.args.get("ano", type=int)
        session["admin_ano"] = ano
    else:
        ano = session.get("admin_ano")

    return programa_slug, ano


@app.route("/admin/login", methods=["GET", "POST"])
def admin_login():
    if request.method == "POST":
        senha = request.form.get("senha")
        # Cada tentativa de login começa "limpa": nunca soma com uma sessão
        # anterior (senão dava pra ficar com admin_ok E gestao_sindicato_ok
        # ao mesmo tempo, se a pessoa logasse duas vezes sem sair antes).
        session.pop("admin_ok", None)
        session.pop("gestao_sindicato_ok", None)
        session.pop("carreta_admin_ok", None)

        if senha and senha == os.environ.get("ADMIN_PASSWORD"):
            session["admin_ok"] = True
            return redirect(url_for("admin_dashboard"))
        if senha and senha == os.environ.get("COMISSAO_PASSWORD"):
            session["admin_ok"] = True
            return redirect(url_for("admin_dashboard"))
        if senha and senha == os.environ.get("GESTAO_SINDICATO_PASSWORD"):
            session["gestao_sindicato_ok"] = True
            return redirect(url_for("admin_sindicatos"))
        if senha and senha == os.environ.get("CARRETA_PASSWORD"):
            session["carreta_admin_ok"] = True
            return redirect(url_for("admin_carreta"))
        flash("Senha incorreta.")
    return render_template("admin_login.html")


@app.route("/admin/logout")
def admin_logout():
    session.pop("admin_ok", None)
    session.pop("gestao_sindicato_ok", None)
    session.pop("carreta_admin_ok", None)
    session.pop("admin_programa", None)
    session.pop("admin_ano", None)
    return redirect(url_for("index"))


@app.route("/admin")
def admin_dashboard():
    """Visão geral: cards de estatísticas + adesões mais recentes."""
    resp = acesso_adesoes_obrigatorio()
    if resp:
        return resp

    db = SessionLocal()

    contagem_status = dict(
        db.query(Adesao.status, func.count(Adesao.id)).group_by(Adesao.status).all()
    )
    total_adesoes = sum(contagem_status.values())
    total_escolas = db.query(func.count(Escola.id)).scalar() or 0
    total_municipios = (
        db.query(func.count(func.distinct(Adesao.municipio_id))).scalar() or 0
    )

    recentes = (
        db.query(Adesao)
        .order_by(Adesao.criado_em.desc())
        .limit(6)
        .all()
    )

    return render_template(
        "admin_dashboard.html",
        pagina_ativa="dashboard",
        total_adesoes=total_adesoes,
        total_rascunho=contagem_status.get("rascunho", 0),
        total_aguardando=contagem_status.get("aguardando_assinaturas", 0),
        total_concluida=contagem_status.get("concluida", 0),
        total_escolas=total_escolas,
        total_municipios=total_municipios,
        recentes=recentes,
    )


@app.route("/admin/adesoes")
def admin_adesoes():
    """Lista completa de adesões, com filtro por programa, ano e status.

    O filtro de Programa fica "fixo": ao escolher Agrinho ou Projeto Valores
    aqui (ou na tela de Estatísticas), a escolha fica guardada na sessão e
    continua valendo ao navegar pelo menu (Rascunhos, Aguardando, Concluídas
    etc.), até a pessoa trocar de programa de novo ou escolher "Todos".
    """
    resp = acesso_adesoes_obrigatorio()
    if resp:
        return resp

    db = SessionLocal()
    query = db.query(Adesao)

    programa_slug, ano = _filtros_programa_ano_sessao()
    status = request.args.get("status")

    if programa_slug:
        programa = get_programa(db, programa_slug)
        query = query.filter(Adesao.programa_id == programa.id)
    if ano:
        query = query.filter(Adesao.ano == ano)
    if status:
        query = query.filter(Adesao.status == status)

    adesoes = query.order_by(Adesao.ano.desc(), Adesao.criado_em.desc()).all()

    anos_existentes = [
        row[0]
        for row in db.query(Adesao.ano).distinct().order_by(Adesao.ano.desc()).all()
    ]

    pagina_ativa = {
        "rascunho": "rascunho",
        "aguardando_assinaturas": "aguardando",
        "concluida": "concluida",
    }.get(status, "adesoes")

    return render_template(
        "admin_adesoes.html",
        pagina_ativa=pagina_ativa,
        adesoes=adesoes,
        anos=anos_existentes,
        filtro_programa=programa_slug or "",
        filtro_ano=ano,
        filtro_status=status or "",
    )


STATUS_LABEL = {
    "rascunho": "Rascunho",
    "aguardando_confirmacao": "Aguardando confirmação do responsável",
    "aguardando_assinaturas": "Aguardando assinaturas",
    "concluida": "Concluída",
}


def _montar_dados_planilha(db, programa, ano, status):
    """Monta as colunas e as linhas da planilha de UM programa — usado tanto
    para gerar o .xlsx quanto para mostrar a mesma tabela na tela. As
    colunas de série mudam de acordo com o programa (Agrinho: 2º ao 9º ano
    | Valores: Infantil 3, 4, 5).

    Ordem das colunas: Ano, Programa, Município, Total de Escolas, Total de
    Professores, [uma coluna por série], Total de Alunos, Status — o Total
    de Alunos vem logo depois das séries, de propósito, porque ele é a soma
    delas (por isso ganha um destaque de cor na hora de montar a planilha).
    """
    series = (
        db.query(Serie).filter_by(programa_id=programa.id).order_by(Serie.ordem).all()
    )

    colunas = ["Ano", "Programa", "Município", "Total de Escolas", "Total de Professores"]
    colunas += [s.nome for s in series]
    colunas += ["Total de Alunos", "Status"]

    query = db.query(Adesao).filter(Adesao.programa_id == programa.id)
    if ano:
        query = query.filter(Adesao.ano == ano)
    if status:
        query = query.filter(Adesao.status == status)
    adesoes = query.order_by(Adesao.ano.desc()).all()
    # ordena por município em Python (evita depender de join extra acima)
    adesoes.sort(key=lambda a: (-(a.ano), a.municipio.nome))

    nome_programa = NOME_PROGRAMA.get(programa.nome, programa.nome)
    linhas = []
    escolas_por_linha = []
    for adesao in adesoes:
        totais_serie = {s.id: 0 for s in series}
        total_alunos = 0
        total_escolas = len(adesao.escolas)
        total_professores = 0
        nomes_escolas = []
        for ae in adesao.escolas:
            professores_escola = ae.quantidade_professores or 0
            total_professores += professores_escola
            alunos_escola = 0
            for mat in ae.matriculas:
                totais_serie[mat.serie_id] = totais_serie.get(mat.serie_id, 0) + mat.quantidade
                total_alunos += mat.quantidade
                alunos_escola += mat.quantidade
            nomes_escolas.append({
                "nome": ae.escola.nome,
                "total_professores": professores_escola,
                "total_alunos": alunos_escola,
            })

        linha = [adesao.ano, nome_programa, adesao.municipio.nome, total_escolas, total_professores]
        linha += [totais_serie.get(s.id, 0) for s in series]
        linha += [total_alunos, STATUS_LABEL.get(adesao.status, adesao.status)]
        linhas.append(linha)
        escolas_por_linha.append(sorted(nomes_escolas, key=lambda e: e["nome"]))

    # índice (0-based) da coluna "Total de Alunos" — usado tanto no Excel
    # quanto na tela pra destacar essa coluna com uma cor bem clarinha.
    indice_total_alunos = len(colunas) - 2
    # índice (0-based) da coluna "Total de Escolas" — usada na tela pra
    # deixar o número clicável e mostrar a lista de escolas daquela linha.
    indice_total_escolas = 3

    return colunas, linhas, indice_total_alunos, indice_total_escolas, escolas_por_linha


def _clarear_cor(cor_hex, fator=0.85):
    """Clareia uma cor hex misturando ela com branco. fator=0.85 dá uma cor
    bem clarinha (85% na direção do branco)."""
    cor_hex = cor_hex.lstrip("#")
    r, g, b = int(cor_hex[0:2], 16), int(cor_hex[2:4], 16), int(cor_hex[4:6], 16)
    r = round(r + (255 - r) * fator)
    g = round(g + (255 - g) * fator)
    b = round(b + (255 - b) * fator)
    return f"{r:02X}{g:02X}{b:02X}"


def _escrever_aba_planilha(ws, colunas, linhas, indice_total_alunos, cores):
    """Escreve os dados já montados (_montar_dados_planilha) numa aba do
    Excel, com a formatação (cabeçalho colorido, coluna de Total de Alunos
    destacada em cor clarinha, filtro, larguras)."""
    cor_tema = cores["tema"].lstrip("#")
    cor_clara = _clarear_cor(cores["tema"], fator=0.85)
    fonte_cabecalho = Font(bold=True, color="FFFFFF")
    fundo_cabecalho = PatternFill("solid", fgColor=cor_tema)
    fundo_destaque = PatternFill("solid", fgColor=cor_clara)
    fonte_destaque = Font(bold=True)
    alinhamento_centro = Alignment(horizontal="center", vertical="center")

    col_total_alunos = indice_total_alunos + 1  # 1-based

    for col, titulo in enumerate(colunas, start=1):
        cel = ws.cell(row=1, column=col, value=titulo)
        cel.alignment = alinhamento_centro
        if col == col_total_alunos:
            cel.font = fonte_destaque
            cel.fill = fundo_destaque
        else:
            cel.font = fonte_cabecalho
            cel.fill = fundo_cabecalho
    ws.freeze_panes = "A2"

    for i, linha_dados in enumerate(linhas, start=2):
        for col, valor in enumerate(linha_dados, start=1):
            cel = ws.cell(row=i, column=col, value=valor)
            if col == col_total_alunos:
                cel.fill = fundo_destaque
                cel.font = Font(bold=True)

    n_series = len(colunas) - 7  # Ano, Programa, Município, Escolas, Professores, Total de Alunos, Status
    larguras = [8, 18, 22, 15, 17] + [10] * max(n_series, 0) + [15, 22]
    for i, largura in enumerate(larguras, start=1):
        ws.column_dimensions[get_column_letter(i)].width = largura

    if linhas:
        ws.auto_filter.ref = f"A1:{get_column_letter(len(colunas))}{len(linhas) + 1}"


@app.route("/admin/adesoes/exportar")
def admin_adesoes_exportar():
    """Gera a planilha (.xlsx) de adesões: uma aba por programa (as colunas
    de série mudam entre Agrinho e Projeto Valores), respeitando os filtros
    de ano/status vindos da tela de Adesões. Se um programa específico já
    estiver selecionado (filtro fixo da sessão ou na URL), a planilha sai
    só com a aba daquele programa."""
    resp = acesso_adesoes_obrigatorio()
    if resp:
        return resp

    db = SessionLocal()

    programa_slug, ano = _filtros_programa_ano_sessao()
    status = request.args.get("status")

    wb = Workbook()
    wb.remove(wb.active)

    if programa_slug:
        programas = [get_programa(db, programa_slug)]
    else:
        programas = db.query(Programa).order_by(Programa.nome).all()

    for programa in programas:
        cores = CORES_PDF.get(programa.nome, CORES_PDF["AGRINHO"])
        nome_aba = NOME_PROGRAMA.get(programa.nome, programa.nome)[:31]
        ws = wb.create_sheet(title=nome_aba)
        colunas, linhas, indice_total_alunos, _, _ = _montar_dados_planilha(db, programa, ano, status)
        _escrever_aba_planilha(ws, colunas, linhas, indice_total_alunos, cores)

    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)

    partes_nome = ["adesoes"]
    if programa_slug:
        partes_nome.append(programa_slug)
    if ano:
        partes_nome.append(str(ano))
    if status:
        partes_nome.append(status)
    nome_arquivo = "_".join(partes_nome) + ".xlsx"

    return send_file(
        buffer,
        as_attachment=True,
        download_name=nome_arquivo,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.route("/admin/adesoes/planilha")
def admin_adesoes_planilha():
    """Mesma planilha da exportação, só que mostrada na tela (uma aba por
    programa, com abas/tabs se houver mais de um)."""
    resp = acesso_adesoes_obrigatorio()
    if resp:
        return resp

    db = SessionLocal()

    programa_slug, ano = _filtros_programa_ano_sessao()
    status = request.args.get("status")

    if programa_slug:
        programas = [get_programa(db, programa_slug)]
    else:
        programas = db.query(Programa).order_by(Programa.nome).all()

    abas = []
    for programa in programas:
        cores = CORES_PDF.get(programa.nome, CORES_PDF["AGRINHO"])
        colunas, linhas, indice_total_alunos, indice_total_escolas, escolas_por_linha = (
            _montar_dados_planilha(db, programa, ano, status)
        )
        abas.append({
            "nome": NOME_PROGRAMA.get(programa.nome, programa.nome),
            "slug": programa.nome.lower(),
            "cor": cores["tema"],
            "cor_clara": "#" + _clarear_cor(cores["tema"], fator=0.85),
            "colunas": colunas,
            "linhas": linhas,
            "indice_total_alunos": indice_total_alunos,
            "indice_total_escolas": indice_total_escolas,
            "escolas_por_linha": escolas_por_linha,
        })

    anos_existentes = [
        row[0]
        for row in db.query(Adesao.ano).distinct().order_by(Adesao.ano.desc()).all()
    ]

    return render_template(
        "admin_planilha.html",
        pagina_ativa="adesoes",
        abas=abas,
        anos=anos_existentes,
        filtro_programa=programa_slug or "",
        filtro_ano=ano,
        filtro_status=status or "",
    )


@app.route("/admin/escolas")
def admin_escolas():
    """Lista de todas as escolas já cadastradas (cadastro reaproveitado ano a ano)."""
    resp = acesso_adesoes_obrigatorio()
    if resp:
        return resp

    db = SessionLocal()
    filtro_municipio = request.args.get("municipio", "").strip()

    query = db.query(Escola).join(Municipio)
    if filtro_municipio:
        query = query.filter(Municipio.nome == filtro_municipio)
    escolas = query.order_by(Municipio.nome, Escola.nome).all()

    return render_template(
        "admin_escolas.html",
        pagina_ativa="escolas",
        escolas=escolas,
        municipios=MUNICIPIOS_CE,
        filtro_municipio=filtro_municipio,
    )


@app.route("/admin/mapa-escolas")
def admin_mapa_escolas():
    """Mapa com as escolas georreferenciadas que foram contempladas em cada
    adesão (ano + programa) — não é o cadastro geral de escolas, é o
    recorte de quem participou naquele ano específico."""
    resp = acesso_adesoes_obrigatorio()
    if resp:
        return resp

    db = SessionLocal()

    programa_slug, ano = _filtros_programa_ano_sessao()

    query = (
        db.query(Escola, Adesao.ano, Adesao.programa_id)
        .join(AdesaoEscola, AdesaoEscola.escola_id == Escola.id)
        .join(Adesao, AdesaoEscola.adesao_id == Adesao.id)
        .join(Municipio, Municipio.id == Escola.municipio_id)
        .filter(Escola.latitude.isnot(None), Escola.longitude.isnot(None))
    )
    if programa_slug:
        programa = get_programa(db, programa_slug)
        query = query.filter(Adesao.programa_id == programa.id)
    if ano:
        query = query.filter(Adesao.ano == ano)

    linhas = query.order_by(Municipio.nome, Escola.nome).all()

    # Uma escola pode aparecer em mais de uma adesão (anos diferentes) — o
    # marcador no mapa é um só por escola, mas mostra todos os anos em que
    # ela participou dentro do filtro atual.
    escolas_mapa = {}
    for escola, ano_adesao, programa_id in linhas:
        item = escolas_mapa.setdefault(escola.id, {
            "nome": escola.nome,
            "municipio": escola.municipio.nome,
            "latitude": float(escola.latitude),
            "longitude": float(escola.longitude),
            "anos": [],
        })
        if ano_adesao not in item["anos"]:
            item["anos"].append(ano_adesao)

    anos_existentes = [
        row[0]
        for row in db.query(Adesao.ano).distinct().order_by(Adesao.ano.desc()).all()
    ]

    return render_template(
        "admin_mapa_escolas.html",
        pagina_ativa="mapa_escolas",
        escolas_mapa=list(escolas_mapa.values()),
        anos=anos_existentes,
        filtro_programa=programa_slug or "",
        filtro_ano=ano,
    )


@app.route("/admin/estatisticas")
def admin_estatisticas():
    """Painel com estatísticas: status das adesões, matrículas por série e
    ranking de municípios com mais escolas cadastradas."""
    resp = acesso_adesoes_obrigatorio()
    if resp:
        return resp

    db = SessionLocal()

    contagem_status = dict(
        db.query(Adesao.status, func.count(Adesao.id)).group_by(Adesao.status).all()
    )
    total_adesoes = sum(contagem_status.values())

    programa_slug = request.args.get("programa")
    if programa_slug:
        session["admin_programa"] = programa_slug
    else:
        # Sem programa explícito na URL: usa o que ficou fixado na sessão
        # (ex.: escolhido na tela de Adesões). Estatísticas sempre precisa
        # de um programa específico, então "Todos" (vazio) cai no Agrinho.
        programa_slug = session.get("admin_programa") or "agrinho"
    programa = get_programa(db, programa_slug)
    series = (
        db.query(Serie).filter_by(programa_id=programa.id).order_by(Serie.ordem).all()
    )

    totais_series = {s.id: 0 for s in series}
    linhas = (
        db.query(Matricula.serie_id, func.sum(Matricula.quantidade))
        .join(AdesaoEscola, Matricula.adesao_escola_id == AdesaoEscola.id)
        .join(Adesao, AdesaoEscola.adesao_id == Adesao.id)
        .filter(Adesao.programa_id == programa.id)
        .group_by(Matricula.serie_id)
        .all()
    )
    for serie_id, soma in linhas:
        totais_series[serie_id] = soma or 0
    maior_serie = max(totais_series.values()) if totais_series and max(totais_series.values()) else 1

    top_municipios = (
        db.query(Municipio.nome, func.count(Escola.id).label("qtd"))
        .join(Escola, Escola.municipio_id == Municipio.id)
        .group_by(Municipio.nome)
        .order_by(func.count(Escola.id).desc())
        .limit(10)
        .all()
    )
    maior_municipio = top_municipios[0][1] if top_municipios else 1

    return render_template(
        "admin_estatisticas.html",
        pagina_ativa="estatisticas",
        contagem_status=contagem_status,
        total_adesoes=total_adesoes,
        programa_slug=programa_slug,
        series=series,
        totais_series=totais_series,
        maior_serie=maior_serie or 1,
        top_municipios=top_municipios,
        maior_municipio=maior_municipio or 1,
    )


# ---------------------------------------------------------------
# Sindicatos Rurais x Municípios
# ---------------------------------------------------------------

@app.route("/admin/sindicatos")
def admin_sindicatos():
    """Lista de municípios com o sindicato/presidente atual, com filtro por
    nome, região FAEC ou sindicato."""
    if not pode_ver_sindicatos():
        return redirect(url_for("admin_login"))

    db = SessionLocal()
    busca = request.args.get("busca", "").strip()
    regiao_faec = request.args.get("regiao_faec", "")
    sindicato_id = request.args.get("sindicato_id", type=int)

    municipios = sind.listar_municipios(db, busca, regiao_faec, sindicato_id)

    avisos_gestao = sind.sindicatos_com_gestao_vencendo(db) if gestao_sindicato_logado() else []
    avisos_vencidos = [(s, dias) for s, dias in avisos_gestao if dias < 0]
    avisos_a_vencer = [(s, dias) for s, dias in avisos_gestao if dias >= 0]

    return render_template(
        "admin_sindicatos.html",
        pagina_ativa="sindicatos",
        municipios=municipios,
        sindicatos=sind.listar_sindicatos(db, regiao_faec),
        regioes_faec=sind.listar_regioes_faec(db),
        filtro_busca=busca,
        filtro_regiao_faec=regiao_faec,
        filtro_sindicato_id=sindicato_id,
        pode_editar_sindicatos=gestao_sindicato_logado(),
        apenas_sindicatos=(gestao_sindicato_logado() and not admin_logado()),
        avisos_gestao=avisos_gestao,
        avisos_vencidos=avisos_vencidos,
        avisos_a_vencer=avisos_a_vencer,
    )


@app.route("/admin/sindicatos/mapa")
def admin_sindicatos_mapa():
    """Dois mapas do Ceará lado a lado: um com os sindicatos
    georreferenciados (um ponto por sindicato, na sede cadastrada) e outro
    com os 184 municípios coloridos por região FAEC — mesma cor nos dois,
    pra dar pra comparar."""
    if not pode_ver_sindicatos():
        return redirect(url_for("admin_login"))

    db = SessionLocal()
    dados = sind.dados_mapa_sindicatos(db)

    return render_template(
        "admin_sindicatos_mapa.html",
        pagina_ativa="mapa_sindicatos",
        apenas_sindicatos=(gestao_sindicato_logado() and not admin_logado()),
        dados_mapa=dados,
    )


@app.route("/admin/sindicatos/exportar")
def admin_sindicatos_exportar():
    """Gera a planilha .xlsx com o estado ATUAL do banco (mesmas colunas da
    planilha original), respeitando o mesmo filtro aplicado na tela."""
    resp = acesso_gestao_sindicatos_obrigatorio()
    if resp:
        return resp

    db = SessionLocal()
    busca = request.args.get("busca", "").strip()
    regiao_faec = request.args.get("regiao_faec", "")
    sindicato_id = request.args.get("sindicato_id", type=int)

    conteudo = sind.gerar_planilha_xlsx(db, busca, regiao_faec, sindicato_id)
    nome_arquivo = f"Municipios_Ceara_Sindicatos_{datetime.now():%Y-%m-%d}.xlsx"

    return Response(
        conteudo,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{nome_arquivo}"'},
    )


@app.route("/admin/sindicatos/ficha")
def admin_sindicatos_ficha():
    """Gera a "Proposta/Ficha de Distribuição dos Municípios nos Sindicatos"
    em PDF, a partir do estado ATUAL do banco."""
    resp = acesso_gestao_sindicatos_obrigatorio()
    if resp:
        return resp

    db = SessionLocal()
    conteudo = sind.gerar_ficha_distribuicao_pdf(db)
    nome_arquivo = f"Proposta_Distribuicao_Municipios_Sindicatos_{datetime.now():%Y-%m-%d}.pdf"

    return Response(
        conteudo,
        mimetype="application/pdf",
        headers={"Content-Disposition": f'inline; filename="{nome_arquivo}"'},
    )


@app.route("/admin/sindicatos/novo", methods=["GET", "POST"])
def admin_sindicato_novo():
    """Cadastra um sindicato novo, já com presidente/contatos e os
    municípios que devem ficar vinculados a ele."""
    resp = acesso_gestao_sindicatos_obrigatorio()
    if resp:
        return resp

    db = SessionLocal()

    if request.method == "POST":
        alterado_por = request.form.get("alterado_por", "").strip() or "não informado"
        cod_ibges_selecionados = [int(v) for v in request.form.getlist("municipios")]

        try:
            sindicato, total = sind.criar_sindicato_com_municipios(
                db, request.form.get("nome", ""), request.form,
                cod_ibges_selecionados,
                alterado_por=alterado_por,
                observacao=request.form.get("observacao", ""),
            )
        except ValueError as erro:
            flash(str(erro))
        else:
            if total:
                flash(f'Sindicato "{sindicato.nome}" cadastrado, com {total} município(s) vinculado(s).')
            else:
                flash(f'Sindicato "{sindicato.nome}" cadastrado. Nenhum município foi vinculado ainda.')
            return redirect(url_for("admin_sindicato_grupo_editar", sindicato_id=sindicato.id))

    return render_template(
        "admin_sindicato_novo.html", pagina_ativa="sindicatos", apenas_sindicatos=True,
        municipios=sind.listar_municipios_para_selecao(db),
    )


@app.route("/admin/sindicatos/grupo/<int:sindicato_id>/ficha")
def admin_sindicato_ficha_individual(sindicato_id):
    """Ficha em PDF de UM sindicato: dados atuais + histórico de alterações."""
    resp = acesso_gestao_sindicatos_obrigatorio()
    if resp:
        return resp

    db = SessionLocal()
    sindicato = sind.obter_sindicato(db, sindicato_id)
    if not sindicato:
        abort(404)

    conteudo = sind.gerar_ficha_sindicato_pdf(db, sindicato)
    nome_arquivo = f"Ficha_{sindicato.nome}_{datetime.now():%Y-%m-%d}.pdf".replace(" ", "_")

    return Response(
        conteudo,
        mimetype="application/pdf",
        headers={"Content-Disposition": f'inline; filename="{nome_arquivo}"'},
    )


@app.route("/admin/sindicatos/processos")
def admin_todos_processos_eleitorais():
    """Visão geral: todos os sindicatos com o processo eleitoral mais
    recente de cada um, pra ver rapidamente a situação de todos sem entrar
    sindicato por sindicato."""
    if not pode_ver_sindicatos():
        return redirect(url_for("admin_login"))

    db = SessionLocal()
    sindicatos_processos = sind.listar_sindicatos_com_processo_recente(db)
    com_processo = [(s, p) for s, p in sindicatos_processos if p is not None]
    sem_processo_todos = [s for s, p in sindicatos_processos if p is None]
    sem_processo = sind.filtrar_e_ordenar_por_mandato_vencendo(sem_processo_todos, anos=(2025, 2026))

    return render_template(
        "admin_todos_processos_eleitorais.html",
        pagina_ativa="processos_eleitorais",
        apenas_sindicatos=(gestao_sindicato_logado() and not admin_logado()),
        com_processo=com_processo,
        sem_processo=sem_processo,
        total_sem_processo=len(sem_processo_todos),
        situacao_label=sind.SITUACAO_PROCESSO_LABEL,
    )


@app.route("/admin/sindicatos/processos/exportar")
def admin_todos_processos_eleitorais_exportar():
    """Baixa em .xlsx a situação do processo eleitoral mais recente de cada
    sindicato."""
    if not pode_ver_sindicatos():
        return redirect(url_for("admin_login"))

    db = SessionLocal()
    conteudo = sind.gerar_planilha_processos_eleitorais_xlsx(db)
    nome_arquivo = f"Processos_Eleitorais_{datetime.now():%Y-%m-%d}.xlsx"

    return Response(
        conteudo,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{nome_arquivo}"'},
    )


@app.route("/admin/sindicatos/processos/cronograma")
def admin_cronograma_eleitoral_exportar():
    """Baixa em .xlsx o Cronograma Consolidado de Eleições Sindicais: uma
    linha por sindicato, colorida por status (verde/amarelo/cinza)."""
    if not pode_ver_sindicatos():
        return redirect(url_for("admin_login"))

    db = SessionLocal()
    conteudo = sind.gerar_cronograma_eleitoral_xlsx(db)
    nome_arquivo = f"Cronograma_Eleicoes_Sindicais_{datetime.now():%Y-%m-%d}.xlsx"

    return Response(
        conteudo,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{nome_arquivo}"'},
    )


@app.route("/admin/sindicatos/grupo/<int:sindicato_id>/processos")
def admin_processos_eleitorais(sindicato_id):
    """Histórico de todos os processos eleitorais (edital, chapa, eleição,
    posse) já registrados para este sindicato."""
    resp = acesso_gestao_sindicatos_obrigatorio()
    if resp:
        return resp

    db = SessionLocal()
    sindicato = sind.obter_sindicato(db, sindicato_id)
    if not sindicato:
        abort(404)

    return render_template(
        "admin_processos_eleitorais.html",
        pagina_ativa="processos_eleitorais",
        apenas_sindicatos=(gestao_sindicato_logado() and not admin_logado()),
        sindicato=sindicato,
        processos=sind.listar_processos_eleitorais(db, sindicato_id),
        situacao_label=sind.SITUACAO_PROCESSO_LABEL,
    )


@app.route("/admin/sindicatos/grupo/<int:sindicato_id>/processos/novo", methods=["GET", "POST"])
def admin_processo_eleitoral_novo(sindicato_id):
    """Cadastra um novo processo eleitoral (fica no histórico do sindicato,
    junto com os outros já cadastrados antes)."""
    resp = acesso_gestao_sindicatos_obrigatorio()
    if resp:
        return resp

    db = SessionLocal()
    sindicato = sind.obter_sindicato(db, sindicato_id)
    if not sindicato:
        abort(404)

    if request.method == "POST":
        alterado_por = request.form.get("alterado_por", "").strip() or "não informado"
        processo = sind.criar_processo_eleitoral(db, sindicato_id, request.form, alterado_por)
        if processo.situacao == "posse_realizada" and processo.posse_data and processo.mandato_fim:
            flash("Processo eleitoral cadastrado — e já virou o mandato atual do sindicato.")
        else:
            flash("Processo eleitoral cadastrado.")
        return redirect(url_for("admin_processos_eleitorais", sindicato_id=sindicato_id))

    return render_template(
        "admin_processo_eleitoral_editar.html",
        pagina_ativa="processos_eleitorais",
        apenas_sindicatos=(gestao_sindicato_logado() and not admin_logado()),
        sindicato=sindicato,
        processo=None,
        cargos_diretoria=sind.CARGOS_DIRETORIA,
        situacao_label=sind.SITUACAO_PROCESSO_LABEL,
    )


@app.route("/admin/sindicatos/grupo/<int:sindicato_id>/processos/<int:processo_id>/editar", methods=["GET", "POST"])
def admin_processo_eleitoral_editar(sindicato_id, processo_id):
    """Edita um processo eleitoral já cadastrado (datas, situação, membros
    da diretoria eleita)."""
    resp = acesso_gestao_sindicatos_obrigatorio()
    if resp:
        return resp

    db = SessionLocal()
    sindicato = sind.obter_sindicato(db, sindicato_id)
    processo = sind.obter_processo_eleitoral(db, processo_id)
    if not sindicato or not processo or processo.sindicato_id != sindicato_id:
        abort(404)

    if request.method == "POST":
        alterado_por = request.form.get("alterado_por", "").strip() or "não informado"
        sind.atualizar_processo_eleitoral(db, processo, request.form, alterado_por)
        if processo.situacao == "posse_realizada" and processo.posse_data and processo.mandato_fim:
            flash("Processo eleitoral atualizado — e já virou o mandato atual do sindicato.")
        else:
            flash("Processo eleitoral atualizado.")
        return redirect(url_for("admin_processos_eleitorais", sindicato_id=sindicato_id))

    return render_template(
        "admin_processo_eleitoral_editar.html",
        pagina_ativa="processos_eleitorais",
        apenas_sindicatos=(gestao_sindicato_logado() and not admin_logado()),
        sindicato=sindicato,
        processo=processo,
        cargos_diretoria=sind.CARGOS_DIRETORIA,
        situacao_label=sind.SITUACAO_PROCESSO_LABEL,
    )


@app.route("/admin/sindicatos/grupo/<int:sindicato_id>/processos/<int:processo_id>/aplicar", methods=["POST"])
def admin_processo_eleitoral_aplicar(sindicato_id, processo_id):
    """Usa a posse/fim de mandato deste processo como o mandato ATUAL do
    sindicato (o que aparece no aviso de \"mandato vencendo\" e na ficha)."""
    resp = acesso_gestao_sindicatos_obrigatorio()
    if resp:
        return resp

    db = SessionLocal()
    processo = sind.obter_processo_eleitoral(db, processo_id)
    if not processo or processo.sindicato_id != sindicato_id:
        abort(404)

    sind.aplicar_processo_como_mandato_atual(db, processo)
    flash("Este processo agora é o mandato atual do sindicato.")
    return redirect(url_for("admin_processos_eleitorais", sindicato_id=sindicato_id))


@app.route("/admin/sindicatos/grupo/<int:sindicato_id>/processos/<int:processo_id>/excluir", methods=["POST"])
def admin_processo_eleitoral_excluir(sindicato_id, processo_id):
    """Remove um processo eleitoral cadastrado por engano."""
    resp = acesso_gestao_sindicatos_obrigatorio()
    if resp:
        return resp

    db = SessionLocal()
    processo = sind.obter_processo_eleitoral(db, processo_id)
    if not processo or processo.sindicato_id != sindicato_id:
        abort(404)

    sind.excluir_processo_eleitoral(db, processo)
    flash("Processo eleitoral removido.")
    return redirect(url_for("admin_processos_eleitorais", sindicato_id=sindicato_id))


@app.route("/admin/sindicatos/grupo/<int:sindicato_id>/foto")
def admin_sindicato_foto(sindicato_id):
    """Serve a foto do presidente de um sindicato, guardada no banco."""
    resp = acesso_gestao_sindicatos_obrigatorio()
    if resp:
        return resp

    db = SessionLocal()
    sindicato = sind.obter_sindicato(db, sindicato_id)
    if not sindicato:
        abort(404)

    if sindicato.foto_presidente_url:
        # Foto já migrada pro Storage: manda o navegador direto pra lá.
        # O Storage já responde com Cache-Control de longa duração, então
        # isso some do egress do banco a partir da segunda visita.
        return redirect(sindicato.foto_presidente_url, code=302)

    if not sindicato.foto_presidente:
        abort(404)

    # Fallback: linha antiga que ainda não passou pela migração
    # (migrar_fotos_para_storage.py). Ao menos aqui adicionamos cache pro
    # navegador não rebaixar a mesma foto a cada página.
    resp = Response(sindicato.foto_presidente, mimetype=sindicato.foto_presidente_tipo or "image/jpeg")
    resp.headers["Cache-Control"] = "private, max-age=86400"
    return resp


@app.route("/admin/sindicatos/grupo/<int:sindicato_id>/editar", methods=["GET", "POST"])
def admin_sindicato_grupo_editar(sindicato_id):
    """Edita os dados de contato (presidente, vice-presidente, telefones,
    e-mails) de UM SINDICATO, aplicando a todos os municípios vinculados a
    ele de uma vez — pra não precisar repetir a edição município por
    município quando só o presidente do sindicato trocou."""
    resp = acesso_gestao_sindicatos_obrigatorio()
    if resp:
        return resp

    db = SessionLocal()
    sindicato = sind.obter_sindicato(db, sindicato_id)
    if not sindicato:
        abort(404)

    municipios = sind.municipios_do_sindicato(db, sindicato_id)

    if request.method == "POST":
        alterado_por = request.form.get("alterado_por", "").strip() or "não informado"
        try:
            sind.renomear_sindicato(db, sindicato, request.form.get("nome", ""))
        except ValueError as erro:
            flash(str(erro))
            return redirect(url_for("admin_sindicato_grupo_editar", sindicato_id=sindicato_id))
        ativo_novo = request.form.get("ativo") == "on"
        sind.atualizar_status_sindicato(db, sindicato, ativo_novo)

        try:
            sind.atualizar_endereco_geo_foto_sindicato(
                db, sindicato,
                request.form.get("endereco", ""),
                request.form.get("latitude", ""),
                request.form.get("longitude", ""),
                request.files.get("foto_presidente"),
                gestao_inicio=request.form.get("gestao_inicio", ""),
                gestao_fim=request.form.get("gestao_fim", ""),
            )
        except ValueError as erro:
            flash(str(erro))
            return redirect(url_for("admin_sindicato_grupo_editar", sindicato_id=sindicato_id))

        total, qtd = sind.salvar_edicao_em_lote_sindicato(
            db, sindicato, request.form,
            alterado_por=alterado_por,
            observacao=request.form.get("observacao", ""),
        )
        if total:
            flash(f"Dados atualizados em {total} de {qtd} município(s) vinculado(s) a {sindicato.nome}.")
        else:
            flash("Nenhum campo foi alterado.")
        return redirect(url_for("admin_sindicato_grupo_editar", sindicato_id=sindicato_id))

    referencia = sind.melhor_referencia(municipios)
    municipios_disponiveis = [
        m for m in sind.listar_municipios_para_selecao(db)
        if not m.sindicato or m.sindicato.id != sindicato_id
    ]
    return render_template(
        "admin_sindicato_grupo_editar.html",
        pagina_ativa="sindicatos",
        sindicato=sindicato,
        municipios=municipios,
        municipios_disponiveis=municipios_disponiveis,
        referencia=referencia,
        apenas_sindicatos=True,
    )


@app.route("/admin/sindicatos/grupo/<int:sindicato_id>/municipios", methods=["POST"])
def admin_sindicato_municipios_vinculo(sindicato_id):
    """Remove ou adiciona um município vinculado a este sindicato, direto da
    tela de edição do sindicato — sem precisar abrir a página de cada
    município separadamente."""
    resp = acesso_gestao_sindicatos_obrigatorio()
    if resp:
        return resp

    db = SessionLocal()
    sindicato = sind.obter_sindicato(db, sindicato_id)
    if not sindicato:
        abort(404)

    alterado_por = request.form.get("alterado_por", "").strip() or "não informado"

    cod_ibge_remover = request.form.get("remover_cod_ibge", type=int)
    cod_ibge_adicionar = request.form.get("adicionar_cod_ibge", type=int)

    if cod_ibge_remover:
        municipio = sind.obter_municipio(db, cod_ibge_remover)
        if municipio and sind.desvincular_municipio(db, municipio, alterado_por):
            flash(f"{municipio.nome} foi removido de {sindicato.nome}.")
        else:
            flash("Não encontrei esse município (ou ele já não estava vinculado a este sindicato).")
    elif cod_ibge_adicionar:
        municipio = sind.obter_municipio(db, cod_ibge_adicionar)
        if not municipio:
            flash("Não encontrei esse município.")
        elif sind.vincular_municipio_a_sindicato(db, municipio, sindicato, alterado_por):
            flash(f"{municipio.nome} agora está vinculado a {sindicato.nome}.")
        else:
            flash(f"{municipio.nome} já estava vinculado a {sindicato.nome}.")
    else:
        flash("Selecione um município pra remover ou adicionar.")

    return redirect(url_for("admin_sindicato_grupo_editar", sindicato_id=sindicato_id))


@app.route("/admin/sindicatos/<int:cod_ibge>/editar", methods=["GET", "POST"])
def admin_sindicato_editar(cod_ibge):
    resp = acesso_gestao_sindicatos_obrigatorio()
    if resp:
        return resp

    db = SessionLocal()
    municipio = sind.obter_municipio(db, cod_ibge)
    if not municipio:
        abort(404)

    if request.method == "POST":
        alterado_por = request.form.get("alterado_por", "").strip() or "não informado"
        mudou = sind.salvar_edicao(
            db, municipio, request.form,
            alterado_por=alterado_por,
            observacao=request.form.get("observacao", ""),
        )
        flash("Alterações salvas." if mudou else "Nenhum campo foi alterado.")
        return redirect(url_for("admin_sindicato_editar", cod_ibge=cod_ibge))

    return render_template(
        "admin_sindicato_editar.html",
        pagina_ativa="sindicatos",
        municipio=municipio,
        sindicatos=sind.listar_sindicatos(db),
        historico=sind.obter_historico(db, cod_ibge)[:8],
        apenas_sindicatos=True,
    )


@app.route("/admin/sindicatos/<int:cod_ibge>/historico")
def admin_sindicato_historico(cod_ibge):
    resp = acesso_gestao_sindicatos_obrigatorio()
    if resp:
        return resp

    db = SessionLocal()
    municipio = sind.obter_municipio(db, cod_ibge)
    if not municipio:
        abort(404)

    return render_template(
        "admin_sindicato_historico.html",
        pagina_ativa="sindicatos",
        municipio=municipio,
        historico=sind.obter_historico(db, cod_ibge),
        apenas_sindicatos=True,
    )


@app.route("/admin/adesao/<int:adesao_id>")
def admin_adesao_detalhe(adesao_id):
    resp = acesso_adesoes_obrigatorio()
    if resp:
        return resp

    db = SessionLocal()
    adesao = db.get(Adesao, adesao_id)
    if not adesao:
        abort(404)
    return render_template(
        "admin_adesao_detalhe.html", adesao=adesao, papeis=PAPEIS, pagina_ativa=""
    )


@app.route("/admin/adesao/<int:adesao_id>/reenviar/<papel>")
def admin_reenviar_assinatura(adesao_id, papel):
    """Permite ao admin reenviar manualmente o e-mail de assinatura de um papel
    específico (ex.: se o responsável perdeu o e-mail original)."""
    resp = acesso_adesoes_obrigatorio()
    if resp:
        return resp

    db = SessionLocal()
    adesao = db.get(Adesao, adesao_id)
    if not adesao:
        abort(404)

    assinatura = next((a for a in adesao.assinaturas if a.papel == papel), None)
    if assinatura and assinatura.email:
        link = url_for("assinar_adesao", token=assinatura.token, _external=True)
        nome_programa = NOME_PROGRAMA.get(adesao.programa.nome, adesao.programa.nome)
        corpo = f"""
        <p>Olá, {assinatura.nome or ''}.</p>
        <p>Lembrete: sua confirmação como <strong>{PAPEIS.get(papel, papel)}</strong>
        ainda está pendente para a adesão ao <strong>{nome_programa} {adesao.ano}</strong>
        do município de <strong>{adesao.municipio.nome}</strong>.</p>
        <p><a href="{link}">{link}</a></p>
        """
        if enviar_email(assinatura.email, f"Lembrete: confirme sua assinatura — {nome_programa}", corpo):
            flash(f"E-mail reenviado para {assinatura.email}.")
        else:
            flash("Não foi possível enviar o e-mail agora. Copie o link diretamente na tela de detalhes.")
    else:
        flash("Não há e-mail cadastrado para esse responsável.")

    return redirect(url_for("admin_adesao_detalhe", adesao_id=adesao_id))


# =========================================================
# Carreta do Agro: formulário público de solicitação + painel admin
# =========================================================

@app.context_processor
def _injetar_meses_nomes():
    return {"MESES_NOMES": carr.MESES_NOMES}


@app.context_processor
def _injetar_contexto_menu():
    """Calcula sozinho, pra QUALQUER tela do admin, se quem está logado só
    deveria ver a seção de Sindicatos (login de Gestão de Sindicatos) ou só
    a seção de Carreta do Agro (login da Carreta) — sem precisar que cada
    rota se lembre de passar isso na mão. Se a rota passar esses valores
    explicitamente, o valor dela vale (isso aqui é só o padrão).
    Também calcula quantas solicitações de carreta estão pendentes, pra
    mostrar um contador no menu (só quem tem acesso à Carreta enxerga)."""
    pendentes = 0
    if pode_ver_carreta():
        db = SessionLocal()
        pendentes = carr.contar_pendentes(db)
    return {
        "apenas_sindicatos": gestao_sindicato_logado() and not admin_logado(),
        "apenas_carreta": carreta_admin_logado() and not admin_logado(),
        "carreta_pendentes_count": pendentes,
    }


@app.route("/carreta", methods=["GET", "POST"])
def carreta_form():
    """Formulário público pra sindicatos/municípios pedirem a Carreta do
    Agro pra um evento — mostra o calendário de disponibilidade pra evitar
    pedir em cima de uma data já comprometida."""
    db = SessionLocal()

    def _renderizar_formulario():
        # Usado tanto na visita normal (GET) quanto quando o POST dá erro
        # de validação — nesse segundo caso, os campos já digitados
        # continuam preenchidos (o template lê de request.form sozinho),
        # em vez de mandar a pessoa pra uma página em branco de novo.
        hoje = datetime.now()
        ano = request.args.get("ano", type=int) or hoje.year
        mes = request.args.get("mes", type=int) or hoje.month
        mes_anterior, ano_anterior = (12, ano - 1) if mes == 1 else (mes - 1, ano)
        mes_proximo, ano_proximo = (1, ano + 1) if mes == 12 else (mes + 1, ano)
        return render_template(
            "carreta_form.html",
            municipios=MUNICIPIOS_CE,
            tipos_recurso=carr.TIPOS_RECURSO,
            calendario=carr.dados_calendario(db, ano, mes),
            bloqueados=carr.dias_bloqueados_do_mes(db, ano, mes),
            semanas=carr.matriz_semanas(ano, mes),
            ano=ano, mes=mes,
            mes_anterior=mes_anterior, ano_anterior=ano_anterior,
            mes_proximo=mes_proximo, ano_proximo=ano_proximo,
        )

    if request.method == "POST":
        try:
            solicitacao = carr.criar_solicitacao(db, request.form, request.files.getlist("oficios"))
        except ValueError as erro:
            flash(str(erro))
            return _renderizar_formulario()

        conflitos = carr.verificar_conflito(
            db, solicitacao.data_inicio, solicitacao.data_fim, excluir_id=solicitacao.id
        )
        avisos_deslocamento = carr.verificar_alerta_deslocamento(
            db, solicitacao.municipio_nome, solicitacao.data_inicio, solicitacao.data_fim,
            excluir_id=solicitacao.id,
        )
        destinatario_interno = os.environ.get("CARRETA_MAIL_DESTINATARIO") or os.environ.get("CARRETA_MAIL_USERNAME")
        if destinatario_interno:
            carr.enviar_notificacao_nova_solicitacao(solicitacao, destinatario_interno)
        else:
            print(
                "[carreta] aviso: não enviei notificação de solicitação nova — "
                "nem CARRETA_MAIL_DESTINATARIO nem CARRETA_MAIL_USERNAME estão configurados."
            )
        carr.enviar_confirmacao_solicitante(solicitacao)

        return render_template(
            "carreta_sucesso.html", solicitacao=solicitacao,
            tipos_recurso=carr.TIPOS_RECURSO,
            tem_conflito=bool(conflitos),
            avisos_deslocamento=avisos_deslocamento,
        )

    return _renderizar_formulario()


@app.route("/admin/carreta")
def admin_carreta():
    """Lista de solicitações de Carreta do Agro, com filtro por situação
    e por mês. Por padrão já abre filtrado no mês/ano atual — pra ver
    tudo, tem a opção "Todos os meses" no filtro (ou o link "Limpar mês")."""
    if not pode_ver_carreta():
        return redirect(url_for("admin_login"))

    db = SessionLocal()
    hoje = datetime.now().date()
    status = request.args.get("status", "")
    if "mes" in request.args:
        # Chave "mes" apareceu na URL de propósito (dropdown ou "Limpar
        # mês") — respeita o que a pessoa escolheu, mesmo que seja "vazio"
        # (= todos os meses), sem cair no padrão do mês atual.
        ano = request.args.get("ano", type=int)
        mes = request.args.get("mes", type=int)
    else:
        # Primeira visita à página (sem nada na URL): abre já no mês atual.
        ano, mes = hoje.year, hoje.month

    # Pendente é sempre "precisa de ação" — não faz sentido esconder atrás
    # de um filtro de mês (o evento pode ser daqui a 2 meses e mesmo assim
    # precisar de decisão AGORA). Por isso, ao filtrar por Pendente, ignora
    # o mês e mostra todas, de qualquer período.
    ano_da_busca, mes_da_busca = (None, None) if status == "pendente" else (ano, mes)

    return render_template(
        "admin_carreta.html",
        pagina_ativa="carreta",
        solicitacoes=carr.listar_solicitacoes(db, status=status or None, ano=ano_da_busca, mes=mes_da_busca),
        status_label=carr.STATUS_LABEL,
        tipos_recurso=carr.TIPOS_RECURSO,
        filtro_status=status,
        filtro_ano=ano,
        filtro_mes=mes,
        ignorando_mes_por_pendente=(status == "pendente"),
    )


@app.route("/admin/carreta/nova", methods=["GET", "POST"])
def admin_carreta_nova():
    """Cadastro de uma solicitação direto pelo admin (não precisa vir do
    formulário público — por exemplo, um pedido que chegou por telefone ou
    já aconteceu e está sendo registrado depois)."""
    if not pode_ver_carreta():
        return redirect(url_for("admin_login"))

    db = SessionLocal()

    if request.method == "POST":
        try:
            solicitacao = carr.criar_solicitacao(db, request.form, request.files.getlist("oficios"))
        except ValueError as erro:
            flash(str(erro))
            return redirect(url_for("admin_carreta_nova"))

        if request.form.get("status") in carr.STATUS_LABEL:
            carr.atualizar_status(
                db, solicitacao, request.form.get("status"),
                alterado_por=request.form.get("alterado_por", ""),
            )

        flash("Solicitação cadastrada.")
        return redirect(url_for("admin_carreta_detalhe", solicitacao_id=solicitacao.id))

    hoje = datetime.now()
    ano = request.args.get("ano", type=int) or hoje.year
    mes = request.args.get("mes", type=int) or hoje.month
    tipo = request.args.get("tipo") if request.args.get("tipo") in carr.TIPOS_RECURSO else "carreta_agro"
    return render_template(
        "admin_carreta_nova.html",
        pagina_ativa="carreta",
        municipios=MUNICIPIOS_CE,
        tipos_recurso=carr.TIPOS_RECURSO,
        status_label=carr.STATUS_LABEL,
        calendario=carr.dados_calendario(db, ano, mes, tipo_recurso=tipo),
        bloqueados=carr.dias_bloqueados_do_mes(db, ano, mes),
        semanas=carr.matriz_semanas(ano, mes),
        ano=ano, mes=mes, tipo=tipo,
    )


@app.route("/admin/carreta/calendario")
def admin_carreta_calendario():
    """Calendário de disponibilidade — visão admin. Os dois programas
    (Carreta do Agro / Carreta da Saúde) usam o MESMO motorista/veículo,
    então o calendário mostra os dois juntos: uma data ocupada por qualquer
    um dos dois bloqueia o outro também. `tipo` aqui só decide qual aba fica
    selecionada na tela, não filtra a disponibilidade."""
    if not pode_ver_carreta():
        return redirect(url_for("admin_login"))

    db = SessionLocal()
    ano = request.args.get("ano", type=int) or datetime.now().year
    mes = request.args.get("mes", type=int) or datetime.now().month
    tipo = request.args.get("tipo") if request.args.get("tipo") in carr.TIPOS_RECURSO else "carreta_agro"
    mes_anterior, ano_anterior = (12, ano - 1) if mes == 1 else (mes - 1, ano)
    mes_proximo, ano_proximo = (1, ano + 1) if mes == 12 else (mes + 1, ano)
    return render_template(
        "admin_carreta_calendario.html",
        pagina_ativa="carreta_calendario",
        ano=ano, mes=mes, tipo=tipo,
        mes_anterior=mes_anterior, ano_anterior=ano_anterior,
        mes_proximo=mes_proximo, ano_proximo=ano_proximo,
        calendario=carr.dados_calendario(db, ano, mes, tipo_recurso=tipo),
        bloqueados=carr.dias_bloqueados_do_mes(db, ano, mes),
        semanas=carr.matriz_semanas(ano, mes),
        status_label=carr.STATUS_LABEL,
        tipos_recurso=carr.TIPOS_RECURSO,
    )


@app.route("/admin/carreta/<int:solicitacao_id>")
def admin_carreta_detalhe(solicitacao_id):
    """Detalhe de uma solicitação: dados, ofícios anexados, e ações de
    aprovar/recusar."""
    if not pode_ver_carreta():
        return redirect(url_for("admin_login"))

    db = SessionLocal()
    solicitacao = carr.obter_solicitacao(db, solicitacao_id)
    if not solicitacao:
        abort(404)

    conflitos = carr.verificar_conflito(
        db, solicitacao.data_inicio, solicitacao.data_fim, excluir_id=solicitacao.id
    )
    avisos_deslocamento = carr.verificar_alerta_deslocamento(
        db, solicitacao.municipio_nome, solicitacao.data_inicio, solicitacao.data_fim,
        excluir_id=solicitacao.id,
    )

    # Mini-mapa de impacto: onde a carreta estaria logo ANTES e logo DEPOIS
    # dessa solicitação, pra ajudar a decidir se vale aprovar (mesmo antes
    # de qualquer aprovação, isso já é útil — é o "e se eu aprovar isso?").
    vizinhos = carr.vizinhos_da_solicitacao(db, solicitacao)
    ponto_solicitacao = _ponto_carreta(solicitacao)
    ponto_anterior = _ponto_carreta(vizinhos["anterior"]) if vizinhos["anterior"] else None
    ponto_proximo = _ponto_carreta(vizinhos["proximo"]) if vizinhos["proximo"] else None
    distancia_antes_km = None
    distancia_depois_km = None
    if vizinhos["anterior"]:
        distancia_antes_km = distancia_municipios.distancia_km(
            vizinhos["anterior"].municipio_nome, solicitacao.municipio_nome
        )
    if vizinhos["proximo"]:
        distancia_depois_km = distancia_municipios.distancia_km(
            solicitacao.municipio_nome, vizinhos["proximo"].municipio_nome
        )

    # Além do vizinho no calendário (que pode estar longe no tempo, se essa
    # solicitação for pra uma data distante), mostra também a comparação com
    # a última localização REAL conhecida da carreta HOJE — mas só quando
    # isso for realmente relevante: se já existe um compromisso confirmado
    # AGENDADO entre hoje e essa solicitação (`vizinhos["anterior"]`), é ELE
    # que importa pra decidir — "onde a carreta estava em setembro" vira
    # ruído sem sentido pra uma solicitação de novembro, por exemplo, já
    # que até lá ela vai ter passado por outros lugares mesmo.
    ponto_ultima_localizacao = None
    distancia_ultima_localizacao_km = None
    ultima_localizacao_eh_atual = False
    if vizinhos["anterior"] is None:
        situacao_global = carr.situacao_atual(db, limite_anteriores=1)
        ultima_localizacao_solicitacao = situacao_global["atual"] or (
            situacao_global["anteriores"][0] if situacao_global["anteriores"] else None
        )
        ultima_localizacao_eh_atual = bool(situacao_global["atual"])
        if ultima_localizacao_solicitacao and ultima_localizacao_solicitacao.id != solicitacao.id:
            ponto_ultima_localizacao = _ponto_carreta(ultima_localizacao_solicitacao)
            distancia_ultima_localizacao_km = distancia_municipios.distancia_km(
                ultima_localizacao_solicitacao.municipio_nome, solicitacao.municipio_nome
            )

    # Monta a sequência do mini-mapa JÁ ORDENADA POR DATA (não pela ordem em
    # que os pontos foram calculados), e numerada — assim fica claro o
    # fluxo real: 1º, 2º, 3º... em vez de vários pinos parecidos sem
    # nenhuma pista de ordem. Evita duplicar o mesmo ponto (ex: "última
    # localização" pode ser o mesmo registro que "antes").
    candidatos = []
    if ponto_ultima_localizacao:
        candidatos.append({
            "ponto": ponto_ultima_localizacao,
            "papel": "atual" if ultima_localizacao_eh_atual else "ultima_localizacao",
            "data_ordenacao": ultima_localizacao_solicitacao.data_fim,
        })
    if ponto_anterior and (not ponto_ultima_localizacao or ponto_anterior["id"] != ponto_ultima_localizacao["id"]):
        candidatos.append({"ponto": ponto_anterior, "papel": "antes", "data_ordenacao": vizinhos["anterior"].data_fim})
    if ponto_solicitacao:
        candidatos.append({"ponto": ponto_solicitacao, "papel": "solicitacao", "data_ordenacao": solicitacao.data_inicio})
    if ponto_proximo:
        candidatos.append({"ponto": ponto_proximo, "papel": "depois", "data_ordenacao": vizinhos["proximo"].data_inicio})

    candidatos.sort(key=lambda c: c["data_ordenacao"])
    sequencia_impacto = [
        {**c["ponto"], "papel": c["papel"], "numero": i + 1} for i, c in enumerate(candidatos)
    ]

    return render_template(
        "admin_carreta_detalhe.html",
        pagina_ativa="carreta",
        solicitacao=solicitacao,
        conflitos=conflitos,
        avisos_deslocamento=avisos_deslocamento,
        status_label=carr.STATUS_LABEL,
        tipos_recurso=carr.TIPOS_RECURSO,
        ponto_solicitacao=ponto_solicitacao,
        ponto_anterior=ponto_anterior,
        ponto_proximo=ponto_proximo,
        sequencia_impacto=sequencia_impacto,
        distancia_antes_km=round(distancia_antes_km) if distancia_antes_km is not None else None,
        distancia_depois_km=round(distancia_depois_km) if distancia_depois_km is not None else None,
        ponto_ultima_localizacao=ponto_ultima_localizacao,
        ultima_localizacao_eh_atual=ultima_localizacao_eh_atual,
        distancia_ultima_localizacao_km=(
            round(distancia_ultima_localizacao_km) if distancia_ultima_localizacao_km is not None else None
        ),
    )


def _ponto_carreta(solicitacao):
    """Monta os dados de um pino do mapa (coordenadas + info) a partir de
    uma SolicitacaoCarreta. Devolve None se o município não for reconhecido
    no GeoJSON (aí quem chamar decide como avisar disso)."""
    centro = distancia_municipios.centro_municipio(solicitacao.municipio_nome)
    if not centro:
        return None
    return {
        "id": solicitacao.id,
        "evento": solicitacao.evento,
        "municipio_nome": solicitacao.municipio_nome,
        "tipo_recurso_chave": solicitacao.tipo_recurso,
        "tipo_recurso": carr.TIPOS_RECURSO.get(solicitacao.tipo_recurso, solicitacao.tipo_recurso),
        "data_inicio": solicitacao.data_inicio.strftime("%d/%m/%Y"),
        "data_fim": solicitacao.data_fim.strftime("%d/%m/%Y"),
        "lat": centro[0], "lng": centro[1],
    }


@app.route("/admin/carreta/mapa")
def admin_carreta_mapa():
    """Mapa do Ceará mostrando o trajeto da carreta num mês específico —
    o que já aconteceu, o que está rolando (se o mês em exibição incluir
    hoje) e o que ainda vai acontecer, com navegação mês a mês."""
    if not pode_ver_carreta():
        return redirect(url_for("admin_login"))

    db = SessionLocal()
    hoje = datetime.now().date()
    ano = request.args.get("ano", type=int) or hoje.year
    mes = request.args.get("mes", type=int) or hoje.month
    mes_anterior, ano_anterior = (12, ano - 1) if mes == 1 else (mes - 1, ano)
    mes_proximo, ano_proximo = (1, ano + 1) if mes == 12 else (mes + 1, ano)
    eh_mes_atual = (ano, mes) == (hoje.year, hoje.month)

    situacao = carr.trajeto_do_mes(db, ano, mes, hoje=hoje)

    # Independente do mês em exibição: onde a carreta está agora, ou (se não
    # tiver nada rolando hoje) qual foi o último compromisso concluído — pra
    # sempre dar pra responder "onde ela está/esteve por último", mesmo
    # navegando pra um mês sem nenhum compromisso.
    situacao_global = carr.situacao_atual(db, hoje=hoje, limite_anteriores=1)
    ultima_solicitacao = situacao_global["atual"] or (
        situacao_global["anteriores"][0] if situacao_global["anteriores"] else None
    )
    ultima_eh_atual = bool(situacao_global["atual"])

    atual = _ponto_carreta(situacao["atual"]) if situacao["atual"] else None
    proximas = [p for p in (_ponto_carreta(s) for s in situacao["proximas"]) if p]
    anteriores = [p for p in (_ponto_carreta(s) for s in situacao["anteriores"]) if p]
    ultima_posicao = _ponto_carreta(ultima_solicitacao) if ultima_solicitacao else None

    return render_template(
        "admin_carreta_mapa.html",
        pagina_ativa="carreta_mapa",
        atual=atual,
        proximas=proximas,
        anteriores=anteriores,
        tem_atual_sem_coordenada=bool(situacao["atual"] and not atual),
        ultima_posicao=ultima_posicao,
        ultima_eh_atual=ultima_eh_atual,
        tem_ultima_sem_coordenada=bool(ultima_solicitacao and not ultima_posicao),
        ano=ano, mes=mes,
        mes_anterior=mes_anterior, ano_anterior=ano_anterior,
        mes_proximo=mes_proximo, ano_proximo=ano_proximo,
        eh_mes_atual=eh_mes_atual,
    )


@app.route("/admin/carreta/<int:solicitacao_id>/status", methods=["POST"])
def admin_carreta_status(solicitacao_id):
    """Aprova ou recusa uma solicitação."""
    if not pode_ver_carreta():
        return redirect(url_for("admin_login"))

    db = SessionLocal()
    solicitacao = carr.obter_solicitacao(db, solicitacao_id)
    if not solicitacao:
        abort(404)

    status = request.form.get("status", "")
    alterado_por = request.form.get("alterado_por", "").strip() or "não informado"
    try:
        carr.atualizar_status(db, solicitacao, status, request.form.get("motivo_recusa", ""), alterado_por)
        flash(f"Solicitação marcada como \"{carr.STATUS_LABEL.get(status, status)}\".")
        if not carr.enviar_notificacao_status(solicitacao):
            flash(
                "⚠️ O status foi salvo, mas o e-mail de aviso pro solicitante NÃO foi enviado "
                "(confira se CARRETA_MAIL_USERNAME/CARRETA_MAIL_PASSWORD estão configurados). "
                "Avise a pessoa por outro meio, se for urgente."
            )
    except ValueError as erro:
        flash(str(erro))

    return redirect(url_for("admin_carreta_detalhe", solicitacao_id=solicitacao_id))


@app.route("/admin/carreta/<int:solicitacao_id>/editar", methods=["GET", "POST"])
def admin_carreta_editar(solicitacao_id):
    """Edita os dados de uma solicitação já cadastrada (programa,
    município, evento, datas, ofício, responsável...) — pra corrigir um
    cadastro feito errado, sem precisar excluir e criar de novo."""
    if not pode_ver_carreta():
        return redirect(url_for("admin_login"))

    db = SessionLocal()
    solicitacao = carr.obter_solicitacao(db, solicitacao_id)
    if not solicitacao:
        abort(404)

    if request.method == "POST":
        try:
            carr.atualizar_solicitacao(db, solicitacao, request.form, request.files.getlist("oficios"))
            flash("Solicitação atualizada.")
            return redirect(url_for("admin_carreta_detalhe", solicitacao_id=solicitacao.id))
        except ValueError as erro:
            flash(str(erro))

    return render_template(
        "admin_carreta_editar.html",
        pagina_ativa="carreta",
        solicitacao=solicitacao,
        municipios=MUNICIPIOS_CE,
        tipos_recurso=carr.TIPOS_RECURSO,
    )


@app.route("/admin/carreta/<int:solicitacao_id>/excluir", methods=["POST"])
def admin_carreta_excluir(solicitacao_id):
    """Remove uma solicitação cadastrada por engano."""
    if not pode_ver_carreta():
        return redirect(url_for("admin_login"))

    db = SessionLocal()
    solicitacao = carr.obter_solicitacao(db, solicitacao_id)
    if not solicitacao:
        abort(404)

    carr.excluir_solicitacao(db, solicitacao)
    flash("Solicitação removida.")
    return redirect(url_for("admin_carreta"))


@app.route("/admin/carreta/bloqueios", methods=["GET", "POST"])
def admin_carreta_bloqueios():
    """Bloqueio de datas — SEMPRE cadastrado manualmente pelo admin (o
    sistema nunca bloqueia data nenhuma sozinho). Serve pra feriado,
    manutenção do veículo, motorista de férias/licença, etc: enquanto o
    bloqueio existir, ninguém consegue pedir a carreta (nem pelo
    formulário público, nem pelo admin) nessas datas."""
    if not pode_ver_carreta():
        return redirect(url_for("admin_login"))

    db = SessionLocal()

    if request.method == "POST":
        try:
            carr.criar_bloqueio(db, request.form, criado_por=request.form.get("criado_por", ""))
            flash("Bloqueio cadastrado.")
        except ValueError as erro:
            flash(str(erro))
        return redirect(url_for("admin_carreta_bloqueios"))

    return render_template(
        "admin_carreta_bloqueios.html",
        pagina_ativa="carreta_bloqueios",
        bloqueios=carr.listar_bloqueios(db),
    )


@app.route("/admin/carreta/bloqueios/<int:bloqueio_id>/excluir", methods=["POST"])
def admin_carreta_bloqueio_excluir(bloqueio_id):
    """Remove um bloqueio — a data volta a ficar disponível pra pedido novo."""
    if not pode_ver_carreta():
        return redirect(url_for("admin_login"))

    db = SessionLocal()
    bloqueio = carr.obter_bloqueio(db, bloqueio_id)
    if not bloqueio:
        abort(404)

    carr.excluir_bloqueio(db, bloqueio)
    flash("Bloqueio removido.")
    return redirect(url_for("admin_carreta_bloqueios"))


@app.route("/admin/carreta/arquivo/<int:arquivo_id>")
def admin_carreta_arquivo(arquivo_id):
    """Baixa/visualiza um ofício em PDF anexado a uma solicitação."""
    if not pode_ver_carreta():
        return redirect(url_for("admin_login"))

    db = SessionLocal()
    from models import ArquivoCarreta
    arquivo = db.get(ArquivoCarreta, arquivo_id)
    if not arquivo:
        abort(404)

    if arquivo.storage_path:
        try:
            conteudo = supabase_storage.baixar_bytes_privado(arquivo.storage_path)
        except Exception:
            abort(404)
        return Response(
            conteudo,
            mimetype=arquivo.tipo_mime or "application/pdf",
            headers={"Content-Disposition": f'inline; filename="{arquivo.nome_arquivo}"'},
        )

    if not arquivo.conteudo:
        abort(404)

    return Response(
        arquivo.conteudo,
        mimetype=arquivo.tipo_mime or "application/pdf",
        headers={"Content-Disposition": f'inline; filename="{arquivo.nome_arquivo}"'},
    )


if __name__ == "__main__":
    init_db()
    app.run(debug=os.environ.get("FLASK_DEBUG") == "1")
