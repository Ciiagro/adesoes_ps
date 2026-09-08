"""Distância aproximada entre municípios do Ceará, pra estimar se dá tempo
do motorista da carreta ir de uma cidade a outra entre dois eventos.

Usa o mesmo GeoJSON que já alimenta os mapas do sistema
(static/ceara_municipios.geojson) — sem precisar de nenhuma API externa
nem depender de internet. A distância é calculada a partir do "centro"
aproximado (média dos pontos) do polígono de cada município e a fórmula de
Haversine (linha reta, não é distância de estrada real — serve como
estimativa pra decidir se precisa de um dia de deslocamento entre dois
eventos, não pra roteirização precisa)."""

import json
import math
import os

_CAMINHO_GEOJSON = os.path.join(os.path.dirname(__file__), "static", "ceara_municipios.geojson")

_centros_cache = None


def _pontos_da_geometria(geometria):
    """Devolve uma lista achatada de (lon, lat) de qualquer Polygon ou
    MultiPolygon do GeoJSON."""
    tipo = geometria.get("type")
    coords = geometria.get("coordinates")
    pontos = []

    def _extrair(anel):
        for item in anel:
            if isinstance(item[0], (int, float)):
                pontos.append((item[0], item[1]))
            else:
                _extrair(item)

    if tipo in ("Polygon", "MultiPolygon"):
        _extrair(coords)
    return pontos


def _carregar_centros():
    global _centros_cache
    if _centros_cache is not None:
        return _centros_cache

    centros = {}
    try:
        with open(_CAMINHO_GEOJSON, encoding="utf-8") as f:
            dados = json.load(f)
        for feature in dados.get("features", []):
            nome = feature.get("properties", {}).get("name")
            if not nome:
                continue
            pontos = _pontos_da_geometria(feature.get("geometry") or {})
            if not pontos:
                continue
            media_lon = sum(p[0] for p in pontos) / len(pontos)
            media_lat = sum(p[1] for p in pontos) / len(pontos)
            centros[nome] = (media_lat, media_lon)
    except (OSError, json.JSONDecodeError):
        pass

    _centros_cache = centros
    return centros


def centro_municipio(municipio_nome):
    """Devolve (lat, lon) do centro aproximado de um município do Ceará, ou
    None se o nome não for reconhecido no GeoJSON."""
    return _carregar_centros().get(municipio_nome)


def distancia_km(municipio_a, municipio_b):
    """Distância em linha reta (km) entre os centros aproximados de dois
    municípios do Ceará. Devolve None se algum dos dois nomes não for
    encontrado no GeoJSON (ex: erro de digitação, ou município de fora do
    Ceará) — quem chamar deve tratar esse caso sem travar o fluxo."""
    if not municipio_a or not municipio_b:
        return None
    centros = _carregar_centros()
    a = centros.get(municipio_a)
    b = centros.get(municipio_b)
    if not a or not b:
        return None
    if municipio_a == municipio_b:
        return 0.0

    lat1, lon1 = a
    lat2, lon2 = b
    raio_terra_km = 6371.0
    d_lat = math.radians(lat2 - lat1)
    d_lon = math.radians(lon2 - lon1)
    hav = (
        math.sin(d_lat / 2) ** 2
        + math.cos(math.radians(lat1)) * math.cos(math.radians(lat2)) * math.sin(d_lon / 2) ** 2
    )
    return raio_terra_km * 2 * math.asin(math.sqrt(hav))
