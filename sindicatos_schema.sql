-- =========================================================
-- Módulo Sindicatos Rurais x Municípios (Ceará)
-- Execute este arquivo UMA VEZ no SQL Editor do Supabase
-- (o mesmo projeto/base já usado pelo sistema de adesões).
-- Não mexe em nada do schema "educacao" já existente.
-- =========================================================

CREATE SCHEMA IF NOT EXISTS sindicatos;

-- Sindicatos rurais (cada um responsável por 1+ municípios)
CREATE TABLE IF NOT EXISTS sindicatos.sindicatos (
    id SERIAL PRIMARY KEY,
    nome VARCHAR(150) UNIQUE NOT NULL,
    ativo BOOLEAN NOT NULL DEFAULT TRUE,
    endereco VARCHAR(300),
    latitude NUMERIC(10, 7),
    longitude NUMERIC(10, 7),
    foto_presidente BYTEA,
    foto_presidente_tipo VARCHAR(50),
    gestao_inicio DATE,
    gestao_fim DATE
);

-- Um registro por município do Ceará, com o sindicato responsável atual
-- e os contatos (presidente, vice/diretor regional, telefones, e-mails).
CREATE TABLE IF NOT EXISTS sindicatos.municipios (
    cod_ibge BIGINT PRIMARY KEY,
    nome VARCHAR(100) NOT NULL,
    regiao_faec VARCHAR(100),
    regiao_sebrae VARCHAR(100),
    regiao_planejamento VARCHAR(100),
    sindicato_id INTEGER REFERENCES sindicatos.sindicatos(id),
    presidente_nome VARCHAR(150),
    secretario_nome VARCHAR(150),
    telefone1 VARCHAR(60),
    telefone2 VARCHAR(60),
    email1 VARCHAR(120),
    email2 VARCHAR(120),
    email3 VARCHAR(120),
    email4 VARCHAR(120),
    vice_presidente_texto VARCHAR(400),
    telefone_vice_presidente VARCHAR(60),
    email_vice_presidente VARCHAR(120),
    atualizado_em TIMESTAMP DEFAULT NOW(),
    atualizado_por VARCHAR(150)
);

-- Histórico de alterações: uma linha por CAMPO alterado, sempre que a
-- pessoa responsável salva uma edição. Nunca é apagado (auditoria).
CREATE TABLE IF NOT EXISTS sindicatos.historico (
    id SERIAL PRIMARY KEY,
    cod_ibge BIGINT NOT NULL REFERENCES sindicatos.municipios(cod_ibge),
    campo VARCHAR(60) NOT NULL,
    valor_antigo TEXT,
    valor_novo TEXT,
    alterado_por VARCHAR(150) NOT NULL,
    alterado_em TIMESTAMP DEFAULT NOW(),
    observacao TEXT
);

CREATE INDEX IF NOT EXISTS idx_hist_sindicatos_municipio
    ON sindicatos.historico (cod_ibge, alterado_em DESC);

CREATE INDEX IF NOT EXISTS idx_municipios_sindicatos_sindicato
    ON sindicatos.municipios (sindicato_id);

-- Ajuste de tamanho de coluna (execute isso se você já rodou este arquivo
-- antes e ainda não tinha essa correção — é seguro rodar de novo, o
-- IF NOT EXISTS acima não recria nada que já existe):
ALTER TABLE sindicatos.municipios ALTER COLUMN telefone1 TYPE VARCHAR(60);
ALTER TABLE sindicatos.municipios ALTER COLUMN telefone2 TYPE VARCHAR(60);

-- Marca se um sindicato ainda está em atividade (execute isso se você já
-- rodou este arquivo antes de ter essa coluna — seguro rodar de novo):
ALTER TABLE sindicatos.sindicatos ADD COLUMN IF NOT EXISTS ativo BOOLEAN NOT NULL DEFAULT TRUE;

-- Endereço, coordenadas e foto do presidente do sindicato (execute isso se
-- você já rodou este arquivo antes de ter essas colunas — seguro rodar de
-- novo):
ALTER TABLE sindicatos.sindicatos ADD COLUMN IF NOT EXISTS endereco VARCHAR(300);
ALTER TABLE sindicatos.sindicatos ADD COLUMN IF NOT EXISTS latitude NUMERIC(10, 7);
ALTER TABLE sindicatos.sindicatos ADD COLUMN IF NOT EXISTS longitude NUMERIC(10, 7);
ALTER TABLE sindicatos.sindicatos ADD COLUMN IF NOT EXISTS foto_presidente BYTEA;
ALTER TABLE sindicatos.sindicatos ADD COLUMN IF NOT EXISTS foto_presidente_tipo VARCHAR(50);
-- Foto passou a morar no Supabase Storage (ver supabase_storage.py); aqui
-- fica só a URL pública. Os dois campos acima ficam só pra linhas antigas
-- que ainda não passaram por migrar_fotos_para_storage.py.
ALTER TABLE sindicatos.sindicatos ADD COLUMN IF NOT EXISTS foto_presidente_url VARCHAR(500);

-- Início e fim da gestão do presidente atual (execute isso se você já rodou
-- este arquivo antes de ter essas colunas — seguro rodar de novo):
ALTER TABLE sindicatos.sindicatos ADD COLUMN IF NOT EXISTS gestao_inicio DATE;
ALTER TABLE sindicatos.sindicatos ADD COLUMN IF NOT EXISTS gestao_fim DATE;

-- Nome do(a) Secretário(a) do Sindicato + contato próprio (telefone/e-mail)
-- do(a) Vice-Presidente Regional, separado do contato geral do sindicato
-- (execute isso se você já rodou este arquivo antes de ter essas colunas —
-- seguro rodar de novo):
ALTER TABLE sindicatos.municipios ADD COLUMN IF NOT EXISTS secretario_nome VARCHAR(150);
ALTER TABLE sindicatos.municipios ADD COLUMN IF NOT EXISTS telefone_vice_presidente VARCHAR(60);
ALTER TABLE sindicatos.municipios ADD COLUMN IF NOT EXISTS email_vice_presidente VARCHAR(120);

-- Histórico de processos eleitorais de cada sindicato (edital, registro de
-- chapa, eleição, posse) e os membros da diretoria eleitos em cada um
-- (seguro rodar de novo, mesmo se essas tabelas já existirem):
CREATE TABLE IF NOT EXISTS sindicatos.processos_eleitorais (
    id SERIAL PRIMARY KEY,
    sindicato_id INTEGER NOT NULL REFERENCES sindicatos.sindicatos(id),
    edital_data DATE,
    registro_chapa_prazo DATE,
    eleicao_data DATE,
    posse_data DATE,
    mandato_fim DATE,
    situacao VARCHAR(30) NOT NULL DEFAULT 'planejado',
    presidente_eleito VARCHAR(150),
    secretario_eleito VARCHAR(150),
    observacoes TEXT,
    criado_em TIMESTAMP DEFAULT now(),
    atualizado_em TIMESTAMP DEFAULT now(),
    atualizado_por VARCHAR(150)
);

CREATE TABLE IF NOT EXISTS sindicatos.diretoria_membros (
    id SERIAL PRIMARY KEY,
    processo_eleitoral_id INTEGER NOT NULL REFERENCES sindicatos.processos_eleitorais(id) ON DELETE CASCADE,
    cargo VARCHAR(60) NOT NULL,
    nome VARCHAR(150) NOT NULL,
    ordem INTEGER NOT NULL DEFAULT 0
);

CREATE INDEX IF NOT EXISTS ix_processos_eleitorais_sindicato ON sindicatos.processos_eleitorais(sindicato_id);
CREATE INDEX IF NOT EXISTS ix_diretoria_membros_processo ON sindicatos.diretoria_membros(processo_eleitoral_id);
