"""Ponto de entrada usado pelo Vercel.

O Vercel Python runtime (@vercel/python) procura, dentro de /api, uma variável
chamada `app` que seja um app WSGI — e é exatamente o que o Flask já é. Este
arquivo só existe para "apontar" pra esse app; toda a lógica de verdade
continua em app.py, na raiz do projeto (não duplicamos nada aqui).
"""
import os
import sys

# garante que a raiz do projeto (onde ficam app.py, models.py, templates/...)
# está no caminho de import, independente de onde o Vercel executa este arquivo
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from app import app  # noqa: E402  (import depois do sys.path de propósito)

# o Vercel detecta esta variável `app` (WSGI) automaticamente
