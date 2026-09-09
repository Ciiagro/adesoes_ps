-- Schema do módulo "Carreta do Agro": pedidos de disponibilização de
-- carretas (SENAR) pra eventos, com calendário de disponibilidade.
-- Execute isso no SQL Editor do Supabase. Seguro rodar de novo (usa
-- IF NOT EXISTS em tudo).

CREATE SCHEMA IF NOT EXISTS carreta;

CREATE TABLE IF NOT EXISTS carreta.solicitacoes (
    id SERIAL PRIMARY KEY,
    tipo_recurso VARCHAR(30) NOT NULL DEFAULT 'carreta_agro',

    municipio_nome VARCHAR(120) NOT NULL,
    sindicato_nome VARCHAR(150),
    evento VARCHAR(200) NOT NULL,
    data_inicio DATE NOT NULL,
    data_fim DATE NOT NULL,

    numero_oficio VARCHAR(60),
    data_oficio DATE,

    responsavel_nome VARCHAR(150),
    responsavel_telefone VARCHAR(60),
    responsavel_email VARCHAR(120),
    observacoes TEXT,

    status VARCHAR(20) NOT NULL DEFAULT 'pendente',
    motivo_recusa TEXT,

    criado_em TIMESTAMP DEFAULT now(),
    atualizado_em TIMESTAMP DEFAULT now(),
    atualizado_por VARCHAR(150)
);

CREATE TABLE IF NOT EXISTS carreta.arquivos (
    id SERIAL PRIMARY KEY,
    solicitacao_id INTEGER NOT NULL REFERENCES carreta.solicitacoes(id) ON DELETE CASCADE,
    nome_arquivo VARCHAR(200) NOT NULL,
    categoria VARCHAR(30) NOT NULL DEFAULT 'oficio',
    conteudo BYTEA,
    tipo_mime VARCHAR(80),
    storage_path VARCHAR(400),
    drive_url VARCHAR(400),
    enviado_em TIMESTAMP DEFAULT now()
);

-- Se a tabela já existia antes desse campo (rodando esse script de novo
-- num banco antigo), garante que a coluna nova seja criada:
ALTER TABLE carreta.arquivos ADD COLUMN IF NOT EXISTS storage_path VARCHAR(400);

-- categoria: "oficio" ou "termo_compromisso" — separa o ofício do termo de
-- compromisso assinado (novo, ver Guia Técnico da Carreta). Registros já
-- existentes (de antes desse campo) recebem "oficio" pelo DEFAULT.
ALTER TABLE carreta.arquivos ADD COLUMN IF NOT EXISTS categoria VARCHAR(30) NOT NULL DEFAULT 'oficio';

CREATE TABLE IF NOT EXISTS carreta.bloqueios (
    id SERIAL PRIMARY KEY,
    data_inicio DATE NOT NULL,
    data_fim DATE NOT NULL,
    motivo VARCHAR(200) NOT NULL,
    criado_por VARCHAR(150),
    criado_em TIMESTAMP DEFAULT now()
);

CREATE INDEX IF NOT EXISTS ix_carreta_solicitacoes_datas ON carreta.solicitacoes(data_inicio, data_fim);
CREATE INDEX IF NOT EXISTS ix_carreta_solicitacoes_status ON carreta.solicitacoes(status);
CREATE INDEX IF NOT EXISTS ix_carreta_arquivos_solicitacao ON carreta.arquivos(solicitacao_id);
CREATE INDEX IF NOT EXISTS ix_carreta_bloqueios_datas ON carreta.bloqueios(data_inicio, data_fim);
