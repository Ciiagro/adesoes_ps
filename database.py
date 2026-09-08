import os
from dotenv import load_dotenv
from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker, scoped_session, declarative_base
from sqlalchemy.pool import NullPool

load_dotenv()

DATABASE_URL = os.environ["DATABASE_URL"]

# Usamos o driver "pg8000" (Python puro) em vez de psycopg2-binary.
# Isso evita erros de compilação/encoding no Windows, que acontecem
# quando o pip precisa compilar o psycopg2 do zero.
# O .env pode continuar com "postgresql://..." (formato padrão do Supabase);
# aqui a gente troca automaticamente para "postgresql+pg8000://...".
if DATABASE_URL.startswith("postgresql://"):
    DATABASE_URL = DATABASE_URL.replace("postgresql://", "postgresql+pg8000://", 1)

# NullPool: cada requisição abre e fecha sua própria conexão, sem manter um
# "pool" de conexões vivas em memória entre chamadas. Isso é importante em
# ambientes serverless (Vercel, AWS Lambda etc.), onde cada invocação pode
# rodar num processo novo — um pool tradicional acabaria criando conexões
# demais no banco (cada processo com seu próprio pool) e estourando o limite
# de conexões do Supabase. Rodando localmente (python app.py) o único efeito
# é abrir uma conexão nova por requisição, que é imperceptível na prática.
#
# Dica para produção no Vercel: use a connection string do "Transaction
# pooler" do Supabase (porta 6543, não 5432) — ela foi feita exatamente para
# esse tipo de ambiente serverless com muitas conexões curtas.
engine = create_engine(DATABASE_URL, poolclass=NullPool, pool_pre_ping=True)

SessionLocal = scoped_session(
    sessionmaker(bind=engine, autoflush=False, autocommit=False)
)

Base = declarative_base()


def init_db():
    """Cria as tabelas que ainda não existirem.
    O schema 'educacao' e os dados iniciais (programas/series) devem
    ser criados rodando schema.sql uma vez no SQL Editor do Supabase.
    """
    import models  # noqa: garante que os modelos foram registrados no Base
    Base.metadata.create_all(bind=engine)
