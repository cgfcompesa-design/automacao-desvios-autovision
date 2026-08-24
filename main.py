import os
import re
import time
import shutil
import unicodedata
import math
import json
import requests
from pathlib import Path
from datetime import datetime, timedelta, date

import pandas as pd

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.options import Options
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC

from selenium.common.exceptions import (
    TimeoutException,
    StaleElementReferenceException,
    NoSuchElementException,
    WebDriverException,
    ElementClickInterceptedException,
)

from selenium.webdriver.edge import service as edge_service

from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter


# ============================================================
# CONFIGURAÇÕES
# ============================================================

URL_LOGIN = "https://www.autovision.com.br/v3/"

URL_RELATORIO_POSICAO = (
    "https://www.autovision.com.br/v3/"
    "modulos/relatorios/"
    "relatorio_posicao.php"
)


# ===========================================================
# GOOGLE APPS SCRIPT - ENVIO DE RESULTADOS
# ============================================================

GOOGLE_SCRIPT_URL = os.environ.get(
    "GOOGLE_SCRIPT_URL",
    "https://script.google.com/macros/s/AKfycbxREO251djkCbe1HKo8wIxDhXM9CVeaBsMF3lzphYDTjM0272WTzne3PnFoMl9sUNWRhw/exec"
)



# ============================================================
# FONTE OFICIAL DE ABASTECIMENTO
# ============================================================

URL_ABASTECIMENTO = (
    "https://docs.google.com/spreadsheets/d/e/"
    "2PACX-1vTNyx3mdkh9hF027_l61y7O7dwYr_gF5ofFwi0mzRY0eNQuKCu3KR3peiCn7Q_832YRjaxR3rqxQGaB/"
    "pub?gid=1282350705&single=true&output=csv"
)

URL_ABASTECIMENTO_EDIT = (
    "https://docs.google.com/spreadsheets/d/"
    "1NdRe71fCKND18sIOdcivRhhwTJYS4JnMjJfGUaAbZuM/"
    "edit?gid=1282350705#gid=1282350705"
)


# ============================================================
# TELEMETRIA OCIOSIDADE
#
# AS PLACAS SERÃO LIDAS AUTOMATICAMENTE DE:
#
# Aba: Geral
# Coluna: PLACA
# ============================================================

URL_TELEMETRIA_OCIOSIDADE = (
    "https://docs.google.com/spreadsheets/d/e/"
    "2PACX-1vQU22jWG-2LuGLDuLYvVQoBbv5JdArq9WUwGcJ3znHPZWHrFABji3IFNFwYCtVX7u8uo-Rd7YJFb9fZ/"
    "pub?gid=756259345&single=true&output=csv"
)

URL_TELEMETRIA_OCIOSIDADE_EDIT = (
    "https://docs.google.com/spreadsheets/d/"
    "1q-SxCJ4C97uEzuPUyvykM8maMw0VxF540nwfK7bv0SA/"
    "edit?gid=756259345#gid=756259345"
)

ABA_TELEMETRIA_OCIOSIDADE = "Geral"

COLUNA_PLACA_TELEMETRIA = "PLACA"


# ============================================================
# PLACAS
#
# NÃO PREENCHER MANUALMENTE.
#
# Será preenchida automaticamente pela função:
#
# carregar_placas_telemetria()
# ============================================================

PLACAS = []


# ============================================================
# PASTAS
# ============================================================

PASTA_BASE = Path.cwd()

PASTA_HTML = (
    PASTA_BASE /
    "relatorios_posicao"
)

PASTA_EXCEL_NEXUS = (
    PASTA_BASE /
    "analises_nexus"
)

ARQUIVO_UNIFICADO = (
    PASTA_BASE /
    "unificado_desvios.xlsx"
)

PASTA_HTML.mkdir(
    parents=True,
    exist_ok=True
)

PASTA_EXCEL_NEXUS.mkdir(
    parents=True,
    exist_ok=True
)


# ============================================================
# MICROSOFT EDGE DRIVER
#
# O msedgedriver será baixado automaticamente pelo Selenium.
# ============================================================

PASTA_DRIVER = (
    PASTA_BASE /
    "drivers"
)

PASTA_DRIVER.mkdir(
    parents=True,
    exist_ok=True
)


# ============================================================
# AUDITORIA LOCAL DOS RELATÓRIOS AUTOVISION
# O Nexus não é mais utilizado para identificar os desvios.
# ============================================================

PADRAO_ARQUIVO = re.compile(
    r"^(?P<placa>[A-Z0-9-]+)_(?P<data>\d{8})_(?P<hora>\d{6})$",
    re.IGNORECASE,
)

PASTA_SAIDA_AUDITORIA = (
    PASTA_BASE /
    "analises_telemetria"
)

ARQUIVO_AUDITORIA = (
    PASTA_SAIDA_AUDITORIA /
    "auditoria_telemetria.xlsx"
)

PASTA_SAIDA_AUDITORIA.mkdir(
    parents=True,
    exist_ok=True
)

TEMPO_MINIMO_PARADA = 15

GAP_MAX_MINUTOS = 20

RAIO_PARADA_METROS = 100

RAIO_BUSCA_POI_METROS = 300

RAIO_DIVERGENCIA_ENDERECO_METROS = 200

TOLERANCIA_ABASTECIMENTO_MIN = 3

TOLERANCIA_IGNICAO_ABASTECIMENTO_MIN = 3




