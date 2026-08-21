import os
import re
import time
import shutil
import unicodedata
import math
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


# ============================================================
# GOOGLE PLACES API - OPCIONAL / ÚLTIMO RECURSO
# ============================================================

GOOGLE_PLACES_API_KEY = os.environ.get(
    "GOOGLE_PLACES_API_KEY",
    ""
)

OVERPASS_URL = os.environ.get(
    "OVERPASS_URL",
    "https://overpass.kumi.systems/api/interpreter"
)

OVERPASS_URLS = [
    OVERPASS_URL,
    "https://overpass-api.de/api/interpreter",
    "https://overpass.private.coffee/api/interpreter",
]

NOMINATIM_USER_AGENT = os.environ.get(
    "NOMINATIM_USER_AGENT",
    "COMPESA-AuditoriaFrota/1.0"
)

OSM_MIN_INTERVAL_S = 1.1

_ULTIMA_CONSULTA_OSM = 0.0

_CACHE_LOCAIS = {}


CRITICIDADE_MAP = {
    "lodging": "A",
    "hotel": "A",
    "motel": "A",
    "beach": "A",
    "bar": "A",
    "night_club": "A",
    "shopping_mall": "A",
    "campground": "A",
    "rv_park": "A",

    "restaurant": "B",
    "meal_takeaway": "B",
    "supermarket": "B",
    "gym": "B",
    "school": "B",
    "university": "B",
    "movie_theater": "B",
    "performing_arts_theater": "B",
    "hospital": "B",
    "doctor": "B",
    "place_of_worship": "B",
    "airport": "B",
    "bus_station": "B",

    "park": "C",
    "bank": "C",
    "atm": "C",
    "bakery": "C",
    "grocery_or_supermarket": "C",
    "cafe": "C",
    "ice_cream_shop": "C",
    "gas_station": "C",
    "parking": "C",
    "car_repair": "C",
    "car_wash": "C",
}


OSM_CRITICIDADE_MAP = {
    ("tourism", "hotel"): "A",
    ("tourism", "motel"): "A",
    ("tourism", "guest_house"): "A",
    ("tourism", "hostel"): "A",
    ("tourism", "apartment"): "A",
    ("tourism", "chalet"): "A",
    ("tourism", "camp_site"): "A",
    ("tourism", "caravan_site"): "A",
    ("tourism", "resort"): "A",

    ("natural", "beach"): "A",

    ("amenity", "bar"): "A",
    ("amenity", "nightclub"): "A",
    ("amenity", "pub"): "A",
    ("amenity", "casino"): "A",
    ("amenity", "love_hotel"): "A",

    ("shop", "mall"): "A",
    ("shop", "department_store"): "A",

    ("amenity", "restaurant"): "B",
    ("amenity", "fast_food"): "B",

    ("shop", "supermarket"): "B",

    ("leisure", "fitness_centre"): "B",

    ("amenity", "school"): "B",
    ("amenity", "university"): "B",
    ("amenity", "college"): "B",

    ("amenity", "cinema"): "B",
    ("amenity", "theatre"): "B",

    ("amenity", "hospital"): "B",
    ("amenity", "clinic"): "B",
    ("amenity", "doctors"): "B",

    ("amenity", "place_of_worship"): "B",

    ("aeroway", "aerodrome"): "B",

    ("amenity", "bus_station"): "B",

    ("leisure", "park"): "C",
    ("amenity", "bank"): "C",
    ("amenity", "atm"): "C",

    ("shop", "bakery"): "C",
    ("shop", "convenience"): "C",
    ("shop", "grocery"): "C",

    ("amenity", "cafe"): "C",
    ("amenity", "ice_cream"): "C",

    ("shop", "ice_cream"): "C",

    ("amenity", "fuel"): "C",
    ("amenity", "parking"): "C",

    ("shop", "car_repair"): "C",
    ("amenity", "car_wash"): "C",
}


OSM_TAG_KEYS = [
    "amenity",
    "shop",
    "tourism",
    "leisure",
    "aeroway"
]

CRITICIDADE_PESO = {
    "A": 3,
    "B": 2,
    "C": 1
}

CRITICIDADE_COR = {
    "A": "🔴",
    "B": "🟡",
    "C": "🔵"
}


# ============================================================
# TEMPOS
# ============================================================

TIMEOUT = 30

TIMEOUT_RELATORIO = 120

TIMEOUT_DOWNLOAD = 90

INTERVALO_ENTRE_PLACAS = 1.0

INTERVALO_ENTRE_ANALISES = 2.0


# ============================================================
# AUTOVISION
# ============================================================

USUARIO = os.environ.get(
    "AUTOVISION_USUARIO",
    "nayarasilva"
)

SENHA = os.environ.get(
    "AUTOVISION_SENHA",
    "20245"
)


# ============================================================
# NEXUS
# ============================================================

NEXUS_EMAIL = os.environ.get(
    "NEXUS_EMAIL",
    "nayarasilva@compesa.com.br"
)

NEXUS_SENHA = os.environ.get(
    "NEXUS_SENHA",
    "acessologin"
)


# ============================================================
# NORMALIZAR TEXTO
# ============================================================

def normalizar_texto(valor):

    if valor is None:
        return ""

    valor = str(valor)

    valor = unicodedata.normalize(
        "NFKD",
        valor
    )

    valor = "".join(
        c
        for c in valor
        if not unicodedata.combining(c)
    )

    valor = re.sub(
        r"\s+",
        " ",
        valor
    )

    return valor.strip().upper()


# ============================================================
# NORMALIZAR PLACA
# ============================================================

def normalizar_placa(valor):

    if valor is None:
        return ""

    valor = normalizar_texto(
        valor
    )

    valor = re.sub(
        r"[^A-Z0-9]",
        "",
        valor
    )

    return valor


# ============================================================
# CARREGAR PLACAS DA TELEMETRIA OCIOSIDADE
#
# Fonte:
# Google Sheets -> aba Geral -> coluna PLACA
#
# Não utiliza mais:
#
# PLACAS = [...]
# ============================================================

def carregar_placas_telemetria():

    global PLACAS

    print()
    print("=" * 70)
    print("CARREGANDO PLACAS DA TELEMETRIA OCIOSIDADE")
    print("=" * 70)

    print(
        "Fonte:",
        URL_TELEMETRIA_OCIOSIDADE
    )

    try:

        # ----------------------------------------------------
        # A URL publicada já aponta para a planilha/aba Geral.
        # ----------------------------------------------------

        df = pd.read_csv(
            URL_TELEMETRIA_OCIOSIDADE,
            dtype=str
        )

    except Exception as erro:

        raise RuntimeError(
            "Não foi possível carregar a planilha "
            "Telemetria Ociosidade.\n"
            f"URL: {URL_TELEMETRIA_OCIOSIDADE}\n"
            f"Erro: {erro}"
        )

    if df.empty:

        raise RuntimeError(
            "A planilha Telemetria Ociosidade "
            "foi carregada, porém está vazia."
        )

    # --------------------------------------------------------
    # NORMALIZAR NOMES DAS COLUNAS
    # --------------------------------------------------------

    mapa_colunas = {}

    for coluna in df.columns:

        coluna_normalizada = normalizar_texto(
            coluna
        )

        coluna_normalizada = re.sub(
            r"[^A-Z0-9]",
            "",
            coluna_normalizada
        )

        mapa_colunas[
            coluna_normalizada
        ] = coluna

    chave_placa = re.sub(
        r"[^A-Z0-9]",
        "",
        normalizar_texto(
            COLUNA_PLACA_TELEMETRIA
        )
    )

    coluna_placa = mapa_colunas.get(
        chave_placa
    )

    if coluna_placa is None:

        raise RuntimeError(
            "A coluna 'PLACA' não foi encontrada "
            "na aba Geral da Telemetria Ociosidade.\n"
            f"Colunas encontradas: {list(df.columns)}"
        )

    # --------------------------------------------------------
    # EXTRAIR PLACAS
    # --------------------------------------------------------

    placas = []

    for valor in df[coluna_placa]:

        placa = normalizar_placa(
            valor
        )

        if placa:

            placas.append(
                placa
            )

    # --------------------------------------------------------
    # REMOVER DUPLICADAS PRESERVANDO ORDEM
    # --------------------------------------------------------

    PLACAS = list(
        dict.fromkeys(
            placas
        )
    )

    if not PLACAS:

        raise RuntimeError(
            "Nenhuma placa válida foi encontrada "
            "na coluna PLACA da aba Geral."
        )

    # --------------------------------------------------------
    # EXIBIR RESULTADO
    # --------------------------------------------------------

    print()
    print("✓ Planilha carregada.")
    print(
        f"✓ Coluna utilizada: {coluna_placa}"
    )

    print(
        f"✓ Total de placas encontradas: "
        f"{len(PLACAS)}"
    )

    print()
    print("PLACAS CARREGADAS:")

    for indice, placa in enumerate(
        PLACAS,
        start=1
    ):

        print(
            f"{indice}. {placa}"
        )

    print("=" * 70)

    return PLACAS


# ============================================================
# VALIDAR PLACAS
#
# Agora a validação é feita depois da leitura da planilha.
# ============================================================

def validar_placas():

    global PLACAS

    placas_normalizadas = []

    for placa in PLACAS:

        placa_normalizada = (
            normalizar_placa(
                placa
            )
        )

        if placa_normalizada:

            placas_normalizadas.append(
                placa_normalizada
            )

    PLACAS = list(
        dict.fromkeys(
            placas_normalizadas
        )
    )

    if not PLACAS:

        raise RuntimeError(
            "Nenhuma placa foi encontrada "
            "na aba Geral, coluna PLACA, "
            "da Telemetria Ociosidade."
        )

    print()
    print("=" * 70)
    print("PLACAS CONFIGURADAS")
    print("=" * 70)

    for indice, placa in enumerate(
        PLACAS,
        start=1
    ):

        print(
            f"{indice}. {placa}"
        )

    print(
        f"Total: {len(PLACAS)}"
    )


# ============================================================
# CRIAR DRIVER EDGE
#
# O Selenium Manager baixa automaticamente o
# msedgedriver.exe compatível com o Edge.
#
# Não é necessário informar manualmente o caminho.
# ============================================================

def criar_driver():

    print()
    print("=" * 70)
    print("INICIANDO MICROSOFT EDGE")
    print("=" * 70)

    options = Options()

    options.add_argument(
        "--start-maximized"
    )

    options.add_argument(
        "--disable-blink-features=AutomationControlled"
    )

    options.add_argument(
        "--disable-notifications"
    )

    options.add_argument(
        "--disable-popup-blocking"
    )

    options.add_experimental_option(
        "excludeSwitches",
        ["enable-automation"]
    )

    preferencias = {
        "download.default_directory":
            str(PASTA_HTML.resolve()),

        "download.prompt_for_download":
            False,

        "download.directory_upgrade":
            True,

        "safebrowsing.enabled":
            True,

        "profile.default_content_setting_values.automatic_downloads":
            1,
    }

    options.add_experimental_option(
        "prefs",
        preferencias
    )

    # --------------------------------------------------------
    # NÃO INFORMAR executable_path.
    #
    # O Selenium Manager identifica a versão do Edge e baixa
    # automaticamente o msedgedriver.exe necessário.
    # --------------------------------------------------------

    try:

        driver = webdriver.Edge(
            options=options
        )

    except Exception as erro:

        raise RuntimeError(
            "Não foi possível iniciar o Microsoft Edge.\n\n"
            "O Selenium tentou localizar/baixar "
            "automaticamente o msedgedriver.exe.\n\n"
            f"Erro original: {erro}"
        )

    driver.set_page_load_timeout(
        120
    )

    print(
        "✓ Microsoft Edge iniciado."
    )

    print(
        "✓ msedgedriver gerenciado automaticamente "
        "pelo Selenium."
    )

    return driver


# ============================================================
# OBTER PERÍODO
# ============================================================

def obter_periodo():

    hoje = datetime.now()

    print()
    print("=" * 70)
    print("CALCULANDO PERÍODO")
    print("=" * 70)

    print(
        "Data atual:",
        hoje.strftime(
            "%d/%m/%Y %H:%M:%S"
        )
    )

    if hoje.weekday() == 0:

        inicio = (
            hoje -
            timedelta(days=3)
        ).replace(
            hour=0,
            minute=0,
            second=0,
            microsecond=0
        )

        fim = (
            hoje -
            timedelta(days=1)
        ).replace(
            hour=23,
            minute=59,
            second=59,
            microsecond=0
        )

        print(
            "Hoje é segunda-feira."
        )

        print(
            "Período: sexta-feira até domingo."
        )

    else:

        ontem = (
            hoje -
            timedelta(days=1)
        )

        inicio = ontem.replace(
            hour=0,
            minute=0,
            second=0,
            microsecond=0
        )

        fim = ontem.replace(
            hour=23,
            minute=59,
            second=59,
            microsecond=0
        )

        print(
            "Período: dia anterior."
        )

    print(
        "INÍCIO:",
        inicio.strftime(
            "%d/%m/%Y %H:%M:%S"
        )
    )

    print(
        "FIM:",
        fim.strftime(
            "%d/%m/%Y %H:%M:%S"
        )
    )

    return inicio, fim


# ============================================================
# LOGIN AUTOVISION
# ============================================================

def fazer_login(driver):

    print()
    print("=" * 70)
    print("LOGIN AUTOVISION")
    print("=" * 70)

    driver.get(
        URL_LOGIN
    )

    wait = WebDriverWait(
        driver,
        TIMEOUT
    )

    campo_usuario = wait.until(
        EC.visibility_of_element_located(
            (
                By.ID,
                "usuario"
            )
        )
    )

    campo_usuario.clear()

    campo_usuario.send_keys(
        USUARIO
    )

    campo_senha = wait.until(
        EC.visibility_of_element_located(
            (
                By.ID,
                "senha"
            )
        )
    )

    campo_senha.clear()

    campo_senha.send_keys(
        SENHA
    )

    botao_login = wait.until(
        EC.element_to_be_clickable(
            (
                By.CSS_SELECTOR,
                "button[type='submit']"
            )
        )
    )

    driver.execute_script(
        "arguments[0].click();",
        botao_login
    )

    print(
        "Login enviado..."
    )

    try:

        wait.until(
            lambda d:
            len(
                d.find_elements(
                    By.ID,
                    "usuario"
                )
            ) == 0
        )

    except TimeoutException:

        raise RuntimeError(
            "Login não concluído."
        )

    print(
        "✓ Login concluído."
    )


# ============================================================
# PREENCHER DATETIME
# ============================================================

def preencher_datetime(
    driver,
    elemento_id,
    valor
):

    wait = WebDriverWait(
        driver,
        TIMEOUT
    )

    campo = wait.until(
        EC.presence_of_element_located(
            (
                By.ID,
                elemento_id
            )
        )
    )

    driver.execute_script(
        """
        const campo = arguments[0];
        const valor = arguments[1];

        campo.value = valor;

        campo.dispatchEvent(
            new Event('input', {
                bubbles: true
            })
        );

        campo.dispatchEvent(
            new Event('change', {
                bubbles: true
            })
        );

        campo.dispatchEvent(
            new Event('blur', {
                bubbles: true
            })
        );
        """,
        campo,
        valor
    )

    WebDriverWait(
        driver,
        TIMEOUT
    ).until(
        lambda d:
        d.find_element(
            By.ID,
            elemento_id
        ).get_attribute(
            "value"
        ) == valor
    )


# ============================================================
# CHECKBOX
# ============================================================

def selecionar_checkbox_por_texto(
    driver,
    texto_procurado
):

    texto_procurado = normalizar_texto(
        texto_procurado
    )

    labels = driver.find_elements(
        By.CSS_SELECTOR,
        "label.custom-control-label"
    )

    for label in labels:

        try:

            texto_label = normalizar_texto(
                label.text
            )

            if texto_procurado not in texto_label:
                continue

            input_element = None

            try:

                input_element = label.find_element(
                    By.CSS_SELECTOR,
                    "input"
                )

            except NoSuchElementException:
                pass

            if input_element is None:

                try:

                    input_element = label.find_element(
                        By.XPATH,
                        "./preceding-sibling::input[1]"
                    )

                except NoSuchElementException:
                    pass

            if input_element is None:

                try:

                    pai = label.find_element(
                        By.XPATH,
                        ".."
                    )

                    input_element = pai.find_element(
                        By.CSS_SELECTOR,
                        "input"
                    )

                except NoSuchElementException:
                    pass

            if input_element is not None:

                if not input_element.is_selected():

                    driver.execute_script(
                        "arguments[0].click();",
                        input_element
                    )

            else:

                driver.execute_script(
                    "arguments[0].click();",
                    label
                )

            print(
                f"✓ Campo selecionado: "
                f"{texto_procurado}"
            )

            return True

        except StaleElementReferenceException:

            continue

    print(
        f"⚠ Campo não encontrado: "
        f"{texto_procurado}"
    )

    return False


# ============================================================
# CONFIGURAR RELATÓRIO
# ============================================================

def configurar_relatorio_posicao(
    driver,
    inicio=None,
    fim=None
):

    wait = WebDriverWait(
        driver,
        TIMEOUT
    )

    if inicio is not None and fim is not None:

        preencher_datetime(
            driver,
            "data_inicial",
            inicio.strftime(
                "%Y-%m-%dT%H:%M"
            )
        )

        preencher_datetime(
            driver,
            "data_final",
            fim.strftime(
                "%Y-%m-%dT%H:%M"
            )
        )

    select_ponto = wait.until(
        EC.presence_of_element_located(
            (
                By.ID,
                "ponto"
            )
        )
    )

    Select(
        select_ponto
    ).select_by_value(
        "off"
    )

    print(
        "✓ Referência: Endereço"
    )

    selecionar_checkbox_por_texto(
        driver,
        "Velocidade"
    )

    selecionar_checkbox_por_texto(
        driver,
        "Motorista"
    )

    selecionar_checkbox_por_texto(
        driver,
        "Município"
    )


# ============================================================
# SELECIONAR PLACA
# ============================================================

def selecionar_placa_autovision(
    driver,
    placa
):

    placa = normalizar_placa(
        placa
    )

    wait = WebDriverWait(
        driver,
        TIMEOUT
    )

    print(
        f"Selecionando placa: {placa}"
    )

    campo_filtro = wait.until(
        EC.visibility_of_element_located(
            (
                By.CSS_SELECTOR,
                "input[name='q']"
            )
        )
    )

    campo_filtro.clear()

    campo_filtro.send_keys(
        placa
    )

    time.sleep(1)

    select_element = wait.until(
        EC.presence_of_element_located(
            (
                By.ID,
                "multiselect"
            )
        )
    )

    def encontrar_option(driver):

        try:

            elemento = driver.find_element(
                By.ID,
                "multiselect"
            )

            select = Select(
                elemento
            )

            for option in select.options:

                texto = normalizar_placa(
                    option.text
                )

                valor = normalizar_placa(
                    option.get_attribute(
                        "value"
                    )
                )

                if (
                    texto == placa
                    or
                    valor == placa
                ):

                    return True

        except Exception:
            pass

        return False

    wait.until(
        encontrar_option
    )

    encontrada = driver.execute_script(
        """
        const select = arguments[0];
        const placa = arguments[1];

        let encontrada = false;

        for (
            let i = 0;
            i < select.options.length;
            i++
        ) {

            const option = select.options[i];

            const texto =
                (option.text || "")
                .toUpperCase()
                .replace(/[^A-Z0-9]/g, "");

            const valor =
                (option.value || "")
                .toUpperCase()
                .replace(/[^A-Z0-9]/g, "");

            if (
                texto === placa ||
                valor === placa
            ) {

                option.selected = true;

                encontrada = true;

            } else {

                option.selected = false;
            }
        }

        select.dispatchEvent(
            new Event(
                "change",
                {
                    bubbles: true
                }
            )
        );

        return encontrada;
        """,
        select_element,
        placa
    )

    if not encontrada:

        raise RuntimeError(
            f"Placa {placa} não encontrada."
        )

    print(
        f"✓ {placa} encontrada."
    )

    botao_right_all = wait.until(
        EC.presence_of_element_located(
            (
                By.ID,
                "multiselect_rightAll"
            )
        )
    )

    driver.execute_script(
        """
        arguments[0].scrollIntoView({
            block: 'center'
        });

        arguments[0].click();
        """,
        botao_right_all
    )

    print(
        "✓ Placa movida para selecionados."
    )

    time.sleep(1)


# ============================================================
# LOCALIZAR BOTÃO RELATÓRIO
# ============================================================

def localizar_botao_relatorio_analitico(
    driver
):

    wait = WebDriverWait(
        driver,
        TIMEOUT
    )

    seletores = [

        (
            By.XPATH,
            "/html/body/section/form/div[2]/"
            "fieldset/div[1]/button[1]"
        ),

        (
            By.CSS_SELECTOR,
            "#formulario > div:nth-child(2) > "
            "fieldset > div.col-xs-7 > "
            "button.btn.submit-button"
        ),

        (
            By.CSS_SELECTOR,
            "button[onclick='relatorioPosicao(0,0);']"
        ),

        (
            By.XPATH,
            "//button[contains("
            "normalize-space(.),"
            "'Relatório analítico'"
            ")]"
        ),

        (
            By.XPATH,
            "//button[contains("
            "normalize-space(.),"
            "'Relatorio analitico'"
            ")]"
        ),
    ]

    for by, seletor in seletores:

        try:

            return wait.until(
                EC.presence_of_element_located(
                    (
                        by,
                        seletor
                    )
                )
            )

        except TimeoutException:
            continue

    try:

        botoes = driver.find_elements(
            By.CSS_SELECTOR,
            "button"
        )

        for botao in botoes:

            try:

                texto = normalizar_texto(
                    botao.text
                )

                onclick = (
                    botao.get_attribute(
                        "onclick"
                    )
                    or ""
                )

                if (
                    "RELATORIOPOSICAO"
                    in normalizar_texto(
                        onclick
                    )
                ):

                    return botao

                if (
                    "RELATORIO ANALITICO"
                    in texto
                ):

                    return botao

            except Exception:
                continue

    except Exception:
        pass

    return None


# ============================================================
# CLICAR RELATÓRIO
# ============================================================

def clicar_relatorio_analitico(
    driver
):

    abas_antes = set(
        driver.window_handles
    )

    botao = (
        localizar_botao_relatorio_analitico(
            driver
        )
    )

    if botao is None:

        raise RuntimeError(
            "Botão Relatório analítico "
            "não encontrado."
        )

    try:

        driver.execute_script(
            """
            arguments[0].scrollIntoView({
                block: 'center'
            });
            """,
            botao
        )

    except Exception:
        pass

    time.sleep(1)

    try:

        botao.click()

    except Exception:

        driver.execute_script(
            "arguments[0].click();",
            botao
        )

    try:

        WebDriverWait(
            driver,
            10
        ).until(
            lambda d:
            len(
                set(
                    d.window_handles
                )
                -
                abas_antes
            ) > 0
        )

    except TimeoutException:

        botao = (
            localizar_botao_relatorio_analitico(
                driver
            )
        )

        if botao is not None:

            try:

                driver.execute_script(
                    "arguments[0].click();",
                    botao
                )

            except Exception:
                pass

        try:

            novas = (
                set(
                    driver.window_handles
                )
                -
                abas_antes
            )

            if not novas:

                resultado = driver.execute_script(
                    """
                    if (
                        typeof relatorioPosicao ===
                        'function'
                    ) {

                        relatorioPosicao(0, 0);

                        return true;
                    }

                    return false;
                    """
                )

                if not resultado:

                    raise RuntimeError(
                        "relatorioPosicao indisponível."
                    )

        except Exception:
            pass

        WebDriverWait(
            driver,
            20
        ).until(
            lambda d:
            len(
                set(
                    d.window_handles
                )
                -
                abas_antes
            ) > 0
        )

    novas_abas = (
        set(
            driver.window_handles
        )
        -
        abas_antes
    )

    if not novas_abas:

        raise RuntimeError(
            "Nova aba não encontrada."
        )

    nova_aba = next(
        iter(novas_abas)
    )

    driver.switch_to.window(
        nova_aba
    )

    print(
        "✓ Nova aba aberta."
    )

    return nova_aba


# ============================================================
# AGUARDAR RELATÓRIO
# ============================================================

def aguardar_relatorio(
    driver
):

    wait = WebDriverWait(
        driver,
        TIMEOUT_RELATORIO
    )

    try:

        wait.until(
            lambda d:
            d.execute_script(
                "return document.readyState"
            ) == "complete"
        )

    except TimeoutException:
        pass

    wait.until(
        EC.presence_of_element_located(
            (
                By.TAG_NAME,
                "body"
            )
        )
    )

    def existe_tabela(driver):

        try:

            tabelas = driver.find_elements(
                By.TAG_NAME,
                "table"
            )

            for tabela in tabelas:

                linhas = tabela.find_elements(
                    By.CSS_SELECTOR,
                    "tbody tr"
                )

                if len(linhas) > 0:

                    return True

        except Exception:
            pass

        return False

    try:

        wait.until(
            existe_tabela
        )

        print(
            "✓ Tabela encontrada."
        )

        return True

    except TimeoutException:

        try:

            texto = normalizar_texto(
                driver.find_element(
                    By.TAG_NAME,
                    "body"
                ).text
            )

            if any(
                termo in texto
                for termo in [
                    "NENHUM REGISTRO",
                    "NENHUM DADO",
                    "SEM DADOS",
                ]
            ):

                return False

        except Exception:
            pass

        raise RuntimeError(
            "Tabela não encontrada."
        )


# ============================================================
# LOCALIZAR ÍCONE HTML
# ============================================================

def localizar_link_html(
    driver
):

    try:

        imagem = driver.find_element(
            By.ID,
            "btnHTML"
        )

        link = imagem.find_element(
            By.XPATH,
            "./ancestor::a[1]"
        )

        return link

    except Exception:
        pass

    try:

        links = driver.find_elements(
            By.CSS_SELECTOR,
            "a[download]"
        )

        for link in links:

            download = (
                link.get_attribute(
                    "download"
                )
                or ""
            )

            if (
                "RELATORIO_POSICAO"
                in normalizar_texto(
                    download
                )
            ):

                return link

    except Exception:
        pass

    return None


# ============================================================
# LISTAR HTMLS
# ============================================================

def listar_htmls():

    if not PASTA_HTML.exists():
        return set()

    return {
        arquivo.name
        for arquivo in PASTA_HTML.iterdir()
        if (
            arquivo.is_file()
            and
            arquivo.suffix.lower()
            in {
                ".html",
                ".htm"
            }
        )
    }


# ============================================================
# CLICAR HTML
# ============================================================

def clicar_icone_html(
    driver
):

    link = (
        localizar_link_html(
            driver
        )
    )

    if link is None:

        raise RuntimeError(
            "Ícone HTML não encontrado."
        )

    print(
        "✓ Ícone HTML encontrado."
    )

    driver.execute_script(
        """
        arguments[0].scrollIntoView({
            block: 'center'
        });
        """,
        link
    )

    time.sleep(1)

    try:

        link.click()

    except Exception:

        driver.execute_script(
            "arguments[0].click();",
            link
        )

    print(
        "✓ Clique no HTML realizado."
    )


# ============================================================
# ESPERAR DOWNLOAD HTML
# ============================================================

def esperar_download_html(
    arquivos_antes
):

    limite = (
        time.time() +
        TIMEOUT_DOWNLOAD
    )

    while time.time() < limite:

        arquivos = [
            arquivo
            for arquivo in PASTA_HTML.iterdir()
            if arquivo.is_file()
        ]

        novos = [
            arquivo
            for arquivo in arquivos
            if (
                arquivo.name
                not in arquivos_antes
            )
            and
            arquivo.suffix.lower()
            in {
                ".html",
                ".htm"
            }
            and
            not arquivo.name.endswith(
                ".crdownload"
            )
        ]

        if novos:

            arquivo = max(
                novos,
                key=lambda x:
                x.stat().st_mtime
            )

            tamanho_1 = (
                arquivo.stat().st_size
            )

            time.sleep(1)

            if arquivo.exists():

                tamanho_2 = (
                    arquivo.stat().st_size
                )

                if (
                    tamanho_1 ==
                    tamanho_2
                ):

                    return arquivo

        time.sleep(0.5)

    return None


# ============================================================
# RENOMEAR HTML
# ============================================================

def renomear_html(
    arquivo,
    placa
):

    nome = (
        f"{nome_seguro(placa)}_"
        f"{datetime.now().strftime('%Y%m%d_%H%M%S')}.html"
    )

    destino = (
        PASTA_HTML /
        nome
    )

    if destino.exists():

        destino.unlink()

    shutil.move(
        str(arquivo),
        str(destino)
    )

    return destino


# ============================================================
# PROCESSAR PLACA
# ============================================================

def processar_placa(
    driver,
    aba_principal,
    placa,
    inicio,
    fim
):

    aba_relatorio = None

    try:

        print()
        print("#" * 70)
        print(
            f"PROCESSANDO: {placa}"
        )
        print("#" * 70)

        driver.switch_to.window(
            aba_principal
        )

        driver.get(
            URL_RELATORIO_POSICAO
        )

        WebDriverWait(
            driver,
            TIMEOUT
        ).until(
            lambda d:
            d.execute_script(
                "return document.readyState"
            ) == "complete"
        )

        preencher_datetime(
            driver,
            "data_inicial",
            inicio.strftime(
                "%Y-%m-%dT%H:%M"
            )
        )

        preencher_datetime(
            driver,
            "data_final",
            fim.strftime(
                "%Y-%m-%dT%H:%M"
            )
        )

        print(
            "✓ Datas preenchidas."
        )

        configurar_relatorio_posicao(
            driver
        )

        selecionar_placa_autovision(
            driver,
            placa
        )

        aba_relatorio = (
            clicar_relatorio_analitico(
                driver
            )
        )

        possui_dados = (
            aguardar_relatorio(
                driver
            )
        )

        if not possui_dados:

            print(
                f"⚠ {placa}: sem registros."
            )

            return None

        arquivos_antes = listar_htmls()

        clicar_icone_html(
            driver
        )

        print(
            "⬇ Aguardando download HTML..."
        )

        arquivo = (
            esperar_download_html(
                arquivos_antes
            )
        )

        if arquivo is None:

            raise TimeoutException(
                "HTML não foi baixado."
            )

        arquivo_final = (
            renomear_html(
                arquivo,
                placa
            )
        )

        print(
            f"✓ HTML salvo: "
            f"{arquivo_final.name}"
        )

        return arquivo_final

    finally:

        try:

            if (
                aba_relatorio
                and
                aba_relatorio
                in driver.window_handles
            ):

                driver.close()

        except Exception:
            pass

        try:

            if (
                aba_principal
                in driver.window_handles
            ):

                driver.switch_to.window(
                    aba_principal
                )

        except Exception:
            pass


# ============================================================
# PROCESSAR TODAS AS PLACAS
# ============================================================

def processar_placas(
    driver,
    inicio,
    fim
):

    arquivos_html = []

    aba_principal = (
        driver.current_window_handle
    )

    total = len(
        PLACAS
    )

    for indice, placa in enumerate(
        PLACAS,
        start=1
    ):

        print()
        print(
            f"[{indice}/{total}]"
        )

        try:

            arquivo = (
                processar_placa(
                    driver,
                    aba_principal,
                    placa,
                    inicio,
                    fim
                )
            )

            if arquivo:

                arquivos_html.append(
                    arquivo
                )

        except Exception as erro:

            print()
            print("!" * 70)

            print(
                f"ERRO NA PLACA {placa}"
            )

            print(
                erro
            )

            print("!" * 70)

        time.sleep(
            INTERVALO_ENTRE_PLACAS
        )

    return arquivos_html
