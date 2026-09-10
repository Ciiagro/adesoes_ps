from sqlalchemy import (
    Column, Integer, String, ForeignKey, UniqueConstraint, DateTime, func,
    BigInteger, Text, Boolean, Numeric, LargeBinary, Date
)
from sqlalchemy.orm import relationship
from database import Base


class Programa(Base):
    __tablename__ = "programas"
    __table_args__ = {"schema": "educacao"}

    id = Column(Integer, primary_key=True)
    nome = Column(String(50), unique=True, nullable=False)


class Municipio(Base):
    __tablename__ = "municipios"
    __table_args__ = (UniqueConstraint("nome", "uf"), {"schema": "educacao"})

    id = Column(Integer, primary_key=True)
    nome = Column(String(100), nullable=False)
    uf = Column(String(2), nullable=False, default="CE")
    prefeito_nome = Column(String(150))
    prefeito_rg = Column(String(30))
    prefeito_cpf = Column(String(20))
    prefeitura_endereco = Column(String(200))
    prefeitura_cep = Column(String(10))
    prefeitura_telefone = Column(String(20))
    prefeitura_email = Column(String(120))
    secretario_nome = Column(String(150))
    secretaria_endereco = Column(String(200))
    secretaria_cep = Column(String(20))  # guarda o CPF do(a) secretário(a), nome da coluna mantido por compatibilidade
    secretaria_telefone = Column(String(20))
    secretaria_email = Column(String(120))


class Coordenador(Base):
    __tablename__ = "coordenadores"
    __table_args__ = {"schema": "educacao"}

    id = Column(Integer, primary_key=True)
    municipio_id = Column(Integer, ForeignKey("educacao.municipios.id"), nullable=False)
    nome = Column(String(150), nullable=False)
    rg = Column(String(30))
    cpf = Column(String(20))
    telefone1 = Column(String(20))
    telefone2 = Column(String(20))
    email = Column(String(120))

    municipio = relationship("Municipio")


class Adesao(Base):
    __tablename__ = "adesoes"
    __table_args__ = (
        UniqueConstraint("programa_id", "municipio_id", "ano"),
        {"schema": "educacao"},
    )

    id = Column(Integer, primary_key=True)
    programa_id = Column(Integer, ForeignKey("educacao.programas.id"), nullable=False)
    municipio_id = Column(Integer, ForeignKey("educacao.municipios.id"), nullable=False)
    ano = Column(Integer, nullable=False)
    coordenador_id = Column(Integer, ForeignKey("educacao.coordenadores.id"))
    responsavel_preenchimento = Column(String(150))
    total_escolas_informado = Column(Integer)
    total_professores_informado = Column(Integer)
    criado_em = Column(DateTime, server_default=func.now())

    # --- Rascunho e fluxo de assinaturas ---
    status = Column(String(30), nullable=False, default="rascunho")
    # 'rascunho' | 'aguardando_assinaturas' | 'concluida'
    rascunho_token = Column(String(64), unique=True)
    enviada_em = Column(DateTime)
    concluida_em = Column(DateTime)

    programa = relationship("Programa")
    municipio = relationship("Municipio")
    coordenador = relationship("Coordenador")
    escolas = relationship(
        "AdesaoEscola", backref="adesao", cascade="all, delete-orphan"
    )
    assinaturas = relationship(
        "Assinatura", backref="adesao", cascade="all, delete-orphan"
    )


class Escola(Base):
    __tablename__ = "escolas"
    __table_args__ = (UniqueConstraint("municipio_id", "nome"), {"schema": "educacao"})

    id = Column(Integer, primary_key=True)
    municipio_id = Column(Integer, ForeignKey("educacao.municipios.id"), nullable=False)
    nome = Column(String(200), nullable=False)
    endereco = Column(String(250))
    contato_telefone = Column(String(20))
    contato_email = Column(String(120))
    gestor1_nome = Column(String(150))
    gestor2_nome = Column(String(150))
    latitude = Column(Numeric(10, 7))
    longitude = Column(Numeric(10, 7))

    municipio = relationship("Municipio")


class AdesaoEscola(Base):
    __tablename__ = "adesao_escolas"
    __table_args__ = (
        UniqueConstraint("adesao_id", "escola_id"),
        {"schema": "educacao"},
    )

    id = Column(Integer, primary_key=True)
    adesao_id = Column(Integer, ForeignKey("educacao.adesoes.id"), nullable=False)
    escola_id = Column(Integer, ForeignKey("educacao.escolas.id"), nullable=False)
    quantidade_professores = Column(Integer)

    escola = relationship("Escola")
    matriculas = relationship(
        "Matricula", backref="adesao_escola", cascade="all, delete-orphan"
    )


class Serie(Base):
    __tablename__ = "series"
    __table_args__ = (
        UniqueConstraint("programa_id", "nome"),
        {"schema": "educacao"},
    )

    id = Column(Integer, primary_key=True)
    programa_id = Column(Integer, ForeignKey("educacao.programas.id"), nullable=False)
    nome = Column(String(30), nullable=False)
    ordem = Column(Integer, nullable=False)


class Matricula(Base):
    __tablename__ = "matriculas"
    __table_args__ = (
        UniqueConstraint("adesao_escola_id", "serie_id"),
        {"schema": "educacao"},
    )

    id = Column(Integer, primary_key=True)
    adesao_escola_id = Column(
        Integer, ForeignKey("educacao.adesao_escolas.id"), nullable=False
    )
    serie_id = Column(Integer, ForeignKey("educacao.series.id"), nullable=False)
    quantidade = Column(Integer, nullable=False, default=0)

    serie = relationship("Serie")


class Assinatura(Base):
    """Registro de assinatura eletrônica de um responsável (prefeito, secretário
    ou coordenador) sobre uma adesão específica. Cada um recebe um link único
    por e-mail e confirma sem precisar acessar o sistema."""

    __tablename__ = "assinaturas"
    __table_args__ = (
        UniqueConstraint("adesao_id", "papel"),
        {"schema": "educacao"},
    )

    id = Column(Integer, primary_key=True)
    adesao_id = Column(Integer, ForeignKey("educacao.adesoes.id"), nullable=False)
    papel = Column(String(20), nullable=False)  # 'prefeito' | 'secretario' | 'coordenador'
    nome = Column(String(150))
    email = Column(String(120))
    token = Column(String(64), unique=True, nullable=False)
    assinado_em = Column(DateTime)
    ip = Column(String(45))
    user_agent = Column(String(255))
    criado_em = Column(DateTime, server_default=func.now())


# =========================================================
# Módulo Sindicatos Rurais x Municípios (schema "sindicatos")
# Independente do restante do sistema (schema "educacao").
# =========================================================

class SindicatoRural(Base):
    """Um sindicato rural (pode ser responsável por vários municípios)."""

    __tablename__ = "sindicatos"
    __table_args__ = {"schema": "sindicatos"}

    id = Column(Integer, primary_key=True)
    nome = Column(String(150), unique=True, nullable=False)
    ativo = Column(Boolean, nullable=False, default=True, server_default="true")
    endereco = Column(String(300))
    # regiao_faec: região do PRÓPRIO sindicato, independente dos municípios
    # vinculados a ele. Normalmente a região "correta" já vem dos municípios
    # (cada um guarda a sua), mas um sindicato pode ficar temporariamente SEM
    # nenhum município (ex: todos foram reatribuídos a outro sindicato) e
    # mesmo assim continuar pertencendo a uma região — esse campo é o que
    # garante que ele não "suma" dos relatórios agrupados por região nesse
    # caso (ver `montar_estrutura_relatorio_contatos`).
    regiao_faec = Column(String(100))
    latitude = Column(Numeric(10, 7))
    longitude = Column(Numeric(10, 7))
    # foto_presidente / foto_presidente_tipo: campos ANTIGOS, guardavam a foto
    # em binário direto na linha do banco. Mantidos só para não quebrar dados
    # já existentes (ver migrar_fotos_para_storage.py) — não escrever mais
    # neles. A foto nova mora no Supabase Storage; aqui fica só a URL.
    foto_presidente = Column(LargeBinary)
    foto_presidente_tipo = Column(String(50))
    foto_presidente_url = Column(String(500))
    gestao_inicio = Column(Date)
    gestao_fim = Column(Date)


class MunicipioSindicato(Base):
    """Um município do Ceará e o sindicato rural responsável por ele hoje,
    junto com os contatos (presidente, vice/diretor regional, telefones,
    e-mails). É o "estado atual" — a mudança de sindicato/presidente de um
    município simplesmente atualiza a linha dele aqui, e o valor anterior
    fica registrado em `historico`."""

    __tablename__ = "municipios"
    __table_args__ = {"schema": "sindicatos"}

    cod_ibge = Column(BigInteger, primary_key=True)
    nome = Column(String(100), nullable=False)
    regiao_faec = Column(String(100))
    regiao_sebrae = Column(String(100))
    regiao_planejamento = Column(String(100))
    sindicato_id = Column(Integer, ForeignKey("sindicatos.sindicatos.id"))
    presidente_nome = Column(String(150))
    secretario_nome = Column(String(150))
    # --- Contato do Sindicato ---
    telefone1 = Column(String(60))
    telefone2 = Column(String(60))
    email1 = Column(String(120))
    email2 = Column(String(120))
    email3 = Column(String(120))
    email4 = Column(String(120))
    # --- Contato do Vice-Presidente Regional (pessoa distinta, que às vezes
    # é também presidente de algum sindicato específico — mas aqui é sempre
    # o contato dele(a) NO PAPEL de vice-presidente/diretor(a) regional) ---
    vice_presidente_texto = Column(String(400))
    telefone_vice_presidente = Column(String(60))
    email_vice_presidente = Column(String(120))
    atualizado_em = Column(DateTime, server_default=func.now())
    atualizado_por = Column(String(150))

    sindicato = relationship("SindicatoRural")


class HistoricoSindicato(Base):
    """Auditoria: uma linha por CAMPO alterado (município, campo, valor
    antigo, valor novo, quem alterou e quando). Nunca é apagada."""

    __tablename__ = "historico"
    __table_args__ = {"schema": "sindicatos"}

    id = Column(Integer, primary_key=True)
    cod_ibge = Column(BigInteger, ForeignKey("sindicatos.municipios.cod_ibge"), nullable=False)
    campo = Column(String(60), nullable=False)
    valor_antigo = Column(Text)
    valor_novo = Column(Text)
    alterado_por = Column(String(150), nullable=False)
    alterado_em = Column(DateTime, server_default=func.now())
    observacao = Column(Text)

    municipio = relationship("MunicipioSindicato")


class ProcessoEleitoral(Base):
    """Um ciclo eleitoral de um sindicato (edital, registro de chapa,
    eleição, posse). Fica tudo guardado — o histórico completo de todos os
    processos já feitos por aquele sindicato, não só o atual."""

    __tablename__ = "processos_eleitorais"
    __table_args__ = {"schema": "sindicatos"}

    id = Column(Integer, primary_key=True)
    sindicato_id = Column(Integer, ForeignKey("sindicatos.sindicatos.id"), nullable=False)

    edital_data = Column(Date)
    registro_chapa_prazo = Column(Date)
    eleicao_data = Column(Date)
    posse_data = Column(Date)
    mandato_fim = Column(Date)

    situacao = Column(String(30), nullable=False, default="planejado")
    presidente_eleito = Column(String(150))
    secretario_eleito = Column(String(150))
    observacoes = Column(Text)

    criado_em = Column(DateTime, server_default=func.now())
    atualizado_em = Column(DateTime, server_default=func.now())
    atualizado_por = Column(String(150))

    sindicato = relationship("SindicatoRural")
    membros = relationship(
        "DiretoriaMembro", order_by="DiretoriaMembro.ordem",
        cascade="all, delete-orphan", backref="processo",
    )


class DiretoriaMembro(Base):
    """Um cargo da diretoria eleito em um processo eleitoral específico
    (Vice-Presidente, Tesoureiro, Conselho Fiscal etc. — presidente e
    secretário(a) já ficam registrados à parte, no próprio sindicato)."""

    __tablename__ = "diretoria_membros"
    __table_args__ = {"schema": "sindicatos"}

    id = Column(Integer, primary_key=True)
    processo_eleitoral_id = Column(Integer, ForeignKey("sindicatos.processos_eleitorais.id"), nullable=False)
    cargo = Column(String(60), nullable=False)
    nome = Column(String(150), nullable=False)
    ordem = Column(Integer, nullable=False, default=0)


class SolicitacaoCarreta(Base):
    """Um pedido de disponibilização de uma carreta (Carreta do Agro, e no
    futuro outras, por causa do campo `tipo_recurso`) pra um evento, num
    período de dias. É o que alimenta o calendário de disponibilidade."""

    __tablename__ = "solicitacoes"
    __table_args__ = {"schema": "carreta"}

    id = Column(Integer, primary_key=True)
    tipo_recurso = Column(String(30), nullable=False, default="carreta_agro")

    municipio_nome = Column(String(120), nullable=False)
    sindicato_nome = Column(String(150))
    evento = Column(String(200), nullable=False)
    data_inicio = Column(Date, nullable=False)
    data_fim = Column(Date, nullable=False)

    numero_oficio = Column(String(60))
    data_oficio = Column(Date)

    responsavel_nome = Column(String(150))
    responsavel_telefone = Column(String(60))
    responsavel_email = Column(String(120))
    observacoes = Column(Text)

    status = Column(String(20), nullable=False, default="pendente")
    motivo_recusa = Column(Text)

    criado_em = Column(DateTime, server_default=func.now())
    atualizado_em = Column(DateTime, server_default=func.now())
    atualizado_por = Column(String(150))

    arquivos = relationship(
        "ArquivoCarreta", order_by="ArquivoCarreta.id",
        cascade="all, delete-orphan", backref="solicitacao",
    )


class ArquivoCarreta(Base):
    """Um arquivo (PDF) anexado a uma solicitação de carreta — pode ser o
    ofício ou o termo de compromisso assinado (ver `categoria`). Por
    enquanto o conteúdo fica guardado aqui mesmo no banco (igual à foto do
    presidente do sindicato); quando a integração com o Google Drive
    estiver pronta, `drive_url` passa a ser preenchido e `conteudo` pode
    ficar vazio."""

    __tablename__ = "arquivos"
    __table_args__ = {"schema": "carreta"}

    id = Column(Integer, primary_key=True)
    solicitacao_id = Column(Integer, ForeignKey("carreta.solicitacoes.id"), nullable=False)
    nome_arquivo = Column(String(200), nullable=False)
    # categoria: "oficio" ou "termo_compromisso" — separa o ofício do termo
    # de compromisso assinado (baixado do site, assinado no gov.br e
    # reanexado). Registros antigos (antes desse campo existir) vêm como
    # "oficio" pelo valor padrão do banco.
    categoria = Column(String(30), nullable=False, server_default="oficio")
    # conteudo: campo ANTIGO (bytea direto no banco) — mantido só pra não
    # quebrar arquivos já existentes (ver migrar_oficios_para_storage.py).
    # Não escrever mais nele; o arquivo novo vai pro Storage e aqui fica só
    # o caminho (storage_path é de um bucket PRIVADO, então isso não é uma
    # URL utilizável direto — precisa passar pela rota do Flask, que baixa
    # com a service_role key só depois de checar login).
    conteudo = Column(LargeBinary)
    tipo_mime = Column(String(80))
    storage_path = Column(String(400))
    drive_url = Column(String(400))
    enviado_em = Column(DateTime, server_default=func.now())


class BloqueioCarreta(Base):
    """Um período em que a carreta NÃO pode ser usada — feriado, manutenção
    do veículo, motorista de férias/licença, etc — cadastrado pelo admin,
    não por quem preenche o formulário público. Bloqueia a data pra pedido
    novo (público ou pelo admin), igual uma solicitação aprovada
    bloquearia, mas sem estar amarrado a nenhum evento/município
    específico.

    `tipo_recurso`: "carreta_agro", "carreta_saude", ou None/vazio pra
    valer pros dois programas juntos (ex: motorista de férias afeta as
    duas, já que é o mesmo motorista; mas a carreta da Saúde quebrada na
    oficina só afeta ela, não a do Agro)."""

    __tablename__ = "bloqueios"
    __table_args__ = {"schema": "carreta"}

    id = Column(Integer, primary_key=True)
    data_inicio = Column(Date, nullable=False)
    data_fim = Column(Date, nullable=False)
    motivo = Column(String(200), nullable=False)
    tipo_recurso = Column(String(30))
    criado_por = Column(String(150))
    criado_em = Column(DateTime, server_default=func.now())
