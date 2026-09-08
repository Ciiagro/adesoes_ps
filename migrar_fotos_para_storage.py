"""Migração ÚNICA: move as fotos de presidente que ainda estão guardadas em
binário (bytea) direto na tabela `sindicatos` pro Supabase Storage, e limpa
os campos antigos depois de confirmar que o upload deu certo.

Rode uma vez, depois de configurar SUPABASE_URL / SUPABASE_SERVICE_KEY no
.env e criar o bucket (ver instruções no topo de supabase_storage.py):

    python migrar_fotos_para_storage.py

É seguro rodar mais de uma vez — sindicatos que já têm foto_presidente_url
preenchida são pulados.
"""

from database import SessionLocal
from models import SindicatoRural
import supabase_storage


def main():
    db = SessionLocal()
    pendentes = (
        db.query(SindicatoRural)
        .filter(SindicatoRural.foto_presidente.isnot(None))
        .filter(SindicatoRural.foto_presidente_url.is_(None))
        .all()
    )

    if not pendentes:
        print("Nada pra migrar — todas as fotos já estão no Storage (ou não há fotos).")
        return

    print(f"Encontrei {len(pendentes)} sindicato(s) com foto ainda em binário no banco.")

    ok, falhas = 0, []
    for sindicato in pendentes:
        try:
            url = supabase_storage.upload_foto_presidente(
                sindicato.id,
                sindicato.foto_presidente,
                sindicato.foto_presidente_tipo or "image/jpeg",
            )
            sindicato.foto_presidente_url = url
            sindicato.foto_presidente = None
            sindicato.foto_presidente_tipo = None
            db.commit()
            ok += 1
            print(f"  ✓ {sindicato.nome} -> {url}")
        except Exception as erro:  # noqa: BLE001
            db.rollback()
            falhas.append((sindicato.nome, str(erro)))
            print(f"  ✗ {sindicato.nome}: {erro}")

    print(f"\nMigrados: {ok}/{len(pendentes)}")
    if falhas:
        print("Falharam (ficaram como estavam, pode rodar o script de novo depois):")
        for nome, erro in falhas:
            print(f"  - {nome}: {erro}")


if __name__ == "__main__":
    main()
