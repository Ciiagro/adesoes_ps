-- =========================================================
-- Sistema de Adesões - Agrinho / Projeto Valores
-- Execute este arquivo UMA VEZ no SQL Editor do Supabase
-- =========================================================

CREATE SCHEMA IF NOT EXISTS educacao;

-- Programas disponíveis (Agrinho, Valores)
CREATE TABLE IF NOT EXISTS educacao.programas (
    id SERIAL PRIMARY KEY,
    nome VARCHAR(50) UNIQUE NOT NULL
);

INSERT INTO educacao.programas (nome) VALUES ('AGRINHO'), ('VALORES')
ON CONFLICT (nome) DO NOTHING;

-- Municípios (dados do prefeito / secretaria) - cadastro único, reaproveitado todo ano
CREATE TABLE IF NOT EXISTS educacao.municipios (
    id SERIAL PRIMARY KEY,
    nome VARCHAR(100) NOT NULL,
    uf CHAR(2) NOT NULL DEFAULT 'CE',
    prefeito_nome VARCHAR(150),
    prefeito_rg VARCHAR(30),
    prefeito_cpf VARCHAR(20),
    prefeitura_endereco VARCHAR(200),
    prefeitura_cep VARCHAR(10),
    prefeitura_telefone VARCHAR(20),
    prefeitura_email VARCHAR(120),
    secretario_nome VARCHAR(150),
    secretaria_endereco VARCHAR(200),
    secretaria_cep VARCHAR(20), -- guarda o CPF do(a) secretário(a), nome da coluna mantido por compatibilidade
    secretaria_telefone VARCHAR(20),
    secretaria_email VARCHAR(120),
    UNIQUE(nome, uf)
);

-- Coordenador(a) do projeto - pode mudar a cada ano, por isso fica ligado à adesão
CREATE TABLE IF NOT EXISTS educacao.coordenadores (
    id SERIAL PRIMARY KEY,
    municipio_id INTEGER NOT NULL REFERENCES educacao.municipios(id),
    nome VARCHAR(150) NOT NULL,
    rg VARCHAR(30),
    cpf VARCHAR(20),
    telefone1 VARCHAR(20),
    telefone2 VARCHAR(20),
    email VARCHAR(120)
);

-- Uma linha por (programa, município, ano) = "a adesão daquele ano"
CREATE TABLE IF NOT EXISTS educacao.adesoes (
    id SERIAL PRIMARY KEY,
    programa_id INTEGER NOT NULL REFERENCES educacao.programas(id),
    municipio_id INTEGER NOT NULL REFERENCES educacao.municipios(id),
    ano INTEGER NOT NULL,
    coordenador_id INTEGER REFERENCES educacao.coordenadores(id),
    responsavel_preenchimento VARCHAR(150),
    total_escolas_informado INTEGER,
    total_professores_informado INTEGER,
    criado_em TIMESTAMP NOT NULL DEFAULT now(),
    status VARCHAR(30) NOT NULL DEFAULT 'rascunho',
    rascunho_token VARCHAR(64) UNIQUE,
    enviada_em TIMESTAMP,
    concluida_em TIMESTAMP,
    UNIQUE(programa_id, municipio_id, ano)
);

-- Migração: bancos criados antes do fluxo de rascunho/assinaturas precisam destas colunas.
ALTER TABLE educacao.adesoes ADD COLUMN IF NOT EXISTS status VARCHAR(30) NOT NULL DEFAULT 'rascunho';
ALTER TABLE educacao.adesoes ADD COLUMN IF NOT EXISTS rascunho_token VARCHAR(64) UNIQUE;
ALTER TABLE educacao.adesoes ADD COLUMN IF NOT EXISTS enviada_em TIMESTAMP;
ALTER TABLE educacao.adesoes ADD COLUMN IF NOT EXISTS concluida_em TIMESTAMP;

-- Escolas - cadastro reaproveitável (uma escola pode aparecer em vários anos)
CREATE TABLE IF NOT EXISTS educacao.escolas (
    id SERIAL PRIMARY KEY,
    municipio_id INTEGER NOT NULL REFERENCES educacao.municipios(id),
    nome VARCHAR(200) NOT NULL,
    endereco VARCHAR(250),
    contato_telefone VARCHAR(20),
    contato_email VARCHAR(120),
    gestor1_nome VARCHAR(150),
    gestor2_nome VARCHAR(150),
    latitude NUMERIC(10, 7),
    longitude NUMERIC(10, 7),
    UNIQUE(municipio_id, nome)
);

-- Migração: bancos criados antes da geolocalização das escolas precisam destas colunas.
ALTER TABLE educacao.escolas ADD COLUMN IF NOT EXISTS latitude NUMERIC(10, 7);
ALTER TABLE educacao.escolas ADD COLUMN IF NOT EXISTS longitude NUMERIC(10, 7);

-- Liga uma escola a uma adesão (ano) específica -> aqui entram as "escolas novas do ano"
CREATE TABLE IF NOT EXISTS educacao.adesao_escolas (
    id SERIAL PRIMARY KEY,
    adesao_id INTEGER NOT NULL REFERENCES educacao.adesoes(id) ON DELETE CASCADE,
    escola_id INTEGER NOT NULL REFERENCES educacao.escolas(id),
    quantidade_professores INTEGER,
    UNIQUE(adesao_id, escola_id)
);

-- Séries de cada programa (Agrinho: 2º ao 9º ano | Valores: Infantil 3,4,5)
CREATE TABLE IF NOT EXISTS educacao.series (
    id SERIAL PRIMARY KEY,
    programa_id INTEGER NOT NULL REFERENCES educacao.programas(id),
    nome VARCHAR(30) NOT NULL,
    ordem INTEGER NOT NULL,
    UNIQUE(programa_id, nome)
);

INSERT INTO educacao.series (programa_id, nome, ordem)
SELECT p.id, s.nome, s.ordem
FROM educacao.programas p
CROSS JOIN (VALUES
    ('2º Ano', 1), ('3º Ano', 2), ('4º Ano', 3), ('5º Ano', 4),
    ('6º Ano', 5), ('7º Ano', 6), ('8º Ano', 7), ('9º Ano', 8)
) AS s(nome, ordem)
WHERE p.nome = 'AGRINHO'
ON CONFLICT (programa_id, nome) DO NOTHING;

INSERT INTO educacao.series (programa_id, nome, ordem)
SELECT p.id, s.nome, s.ordem
FROM educacao.programas p
CROSS JOIN (VALUES
    ('Infantil 3', 1), ('Infantil 4', 2), ('Infantil 5', 3)
) AS s(nome, ordem)
WHERE p.nome = 'VALORES'
ON CONFLICT (programa_id, nome) DO NOTHING;

-- Quantidade de matrículas por série, por escola, dentro de uma adesão (ano)
CREATE TABLE IF NOT EXISTS educacao.matriculas (
    id SERIAL PRIMARY KEY,
    adesao_escola_id INTEGER NOT NULL REFERENCES educacao.adesao_escolas(id) ON DELETE CASCADE,
    serie_id INTEGER NOT NULL REFERENCES educacao.series(id),
    quantidade INTEGER NOT NULL DEFAULT 0,
    UNIQUE(adesao_escola_id, serie_id)
);

-- Assinatura eletrônica de cada responsável (prefeito, secretário, coordenador)
-- sobre uma adesão específica. Cada um recebe um link único por e-mail.
CREATE TABLE IF NOT EXISTS educacao.assinaturas (
    id SERIAL PRIMARY KEY,
    adesao_id INTEGER NOT NULL REFERENCES educacao.adesoes(id) ON DELETE CASCADE,
    papel VARCHAR(20) NOT NULL, -- 'prefeito' | 'secretario' | 'coordenador'
    nome VARCHAR(150),
    email VARCHAR(120),
    token VARCHAR(64) UNIQUE NOT NULL,
    assinado_em TIMESTAMP,
    ip VARCHAR(45),
    user_agent VARCHAR(255),
    criado_em TIMESTAMP NOT NULL DEFAULT now(),
    UNIQUE(adesao_id, papel)
);

-- Migração: se o banco já existia com secretaria_cep VARCHAR(10) (era pra ser CPF), alarga a coluna.
ALTER TABLE educacao.municipios ALTER COLUMN secretaria_cep TYPE VARCHAR(20);
