
import os
import time
import json
import sys

from datetime import datetime, timedelta
from zoneinfo import ZoneInfo

import pandas as pd
import requests

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException


# ============================================================
# CORRIGIR CODIFICAÇÃO DO CONSOLE
# ============================================================

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(
        encoding="utf-8",
        errors="replace"
    )

if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(
        encoding="utf-8",
        errors="replace"
    )


# ============================================================
# CONFIGURAÇÕES
# ============================================================

URL_LOGIN = "https://www.autovision.com.br/v3/"

TIMEOUT = 60

FUSO_BRASILIA = ZoneInfo(
    "America/Sao_Paulo"
)


# ============================================================
# AUTOVISION
# ============================================================

USUARIO = os.environ.get(
    "AUTOVISION_USUARIO",
    "nayarasilva"
)

SENHA = os.environ.get(
    "AUTOVISION_SENHA",
    "20005"
)


# ============================================================
# GOOGLE APPS SCRIPT
# ============================================================

GOOGLE_SCRIPT_URL = (
    "https://script.google.com/macros/s/"
    "AKfycbwVgLOtn5n92eCZchaKo9naF_ix0lmjzcwNdm-HQafSOH48ZcMrCf9_MyfDHsAwp14"
    "/exec"
)


# ============================================================
# CRIAR NAVEGADOR
# ============================================================

def criar_driver():

    print("\n" + "=" * 70)
    print("INICIANDO NAVEGADOR")
    print("=" * 70)

    options = Options()

    options.add_argument("--headless=new")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--window-size=1920,1080")

    driver = webdriver.Edge(
        options=options
    )

    driver.set_page_load_timeout(
        TIMEOUT
    )

    print("✓ Navegador iniciado.")

    return driver


# ============================================================
# LOGIN AUTOVISION
# ============================================================

def fazer_login(driver):

    print("\n" + "=" * 70)
    print("LOGIN AUTOVISION")
    print("=" * 70)

    if not USUARIO:
        raise RuntimeError(
            "AUTOVISION_USUARIO não configurado."
        )

    if not SENHA:
        raise RuntimeError(
            "AUTOVISION_SENHA não configurada."
        )

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

    print("Login enviado...")

    try:

        wait.until(
            lambda d: len(
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

    print("✓ Login concluído.")


# ============================================================
# ACESSAR RELATÓRIO DE OCIOSIDADE
# ============================================================

def acessar_ociosidade(driver):

    print("\n" + "=" * 70)
    print("ACESSANDO RELATÓRIO DE OCIOSIDADE")
    print("=" * 70)

    wait = WebDriverWait(
        driver,
        TIMEOUT
    )

    link = wait.until(
        EC.presence_of_element_located(
            (
                By.CSS_SELECTOR,
                'a[href="modulos/relatorios/relatorio_ociosidade.php"]'
            )
        )
    )

    driver.execute_script(
        "arguments[0].click();",
        link
    )

    time.sleep(3)

    # --------------------------------------------------------
    # TENTA ENCONTRAR O RELATÓRIO DENTRO DE IFRAME
    # --------------------------------------------------------

    frames = driver.find_elements(
        By.TAG_NAME,
        "iframe"
    )

    for frame in frames:

        try:

            driver.switch_to.default_content()

            driver.switch_to.frame(
                frame
            )

            if driver.find_elements(
                By.ID,
                "data_inicial"
            ):

                print(
                    "✓ Página de Ociosidade encontrada."
                )

                return

        except Exception:

            pass

    # --------------------------------------------------------
    # CASO NÃO ESTEJA EM IFRAME
    # --------------------------------------------------------

    driver.switch_to.default_content()

    wait.until(
        EC.presence_of_element_located(
            (
                By.ID,
                "data_inicial"
            )
        )
    )

    print(
        "✓ Página de Ociosidade encontrada."
    )


# ============================================================
# PREENCHER DATETIME
# ============================================================

def preencher_datetime(
    driver,
    elemento_id,
    valor
):

    campo = WebDriverWait(
        driver,
        TIMEOUT
    ).until(
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
            new Event(
                'input',
                { bubbles: true }
            )
        );

        campo.dispatchEvent(
            new Event(
                'change',
                { bubbles: true }
            )
        );

        campo.dispatchEvent(
            new Event(
                'blur',
                { bubbles: true }
            )
        );
        """,
        campo,
        valor
    )

    print(
        f"✓ {elemento_id}: {valor}"
    )


# ============================================================
# PREENCHER VALOR = 0
# ============================================================

def preencher_valor(driver):

    campo = WebDriverWait(
        driver,
        TIMEOUT
    ).until(
        EC.presence_of_element_located(
            (
                By.ID,
                "valor"
            )
        )
    )

    driver.execute_script(
        """
        arguments[0].value = '0';

        arguments[0].dispatchEvent(
            new Event(
                'input',
                { bubbles: true }
            )
        );

        arguments[0].dispatchEvent(
            new Event(
                'change',
                { bubbles: true }
            )
        );
        """,
        campo
    )

    print(
        "✓ Valor preenchido: 0"
    )


# ============================================================
# SELECIONAR TODAS AS PLACAS
# ============================================================

def selecionar_todas_placas(driver):

    botao = WebDriverWait(
        driver,
        TIMEOUT
    ).until(
        EC.element_to_be_clickable(
            (
                By.ID,
                "multiselect_rightAll"
            )
        )
    )

    driver.execute_script(
        "arguments[0].click();",
        botao
    )

    time.sleep(2)

    print(
        "✓ Todas as placas selecionadas."
    )


# ============================================================
# GERAR RELATÓRIO
# ============================================================

def gerar_relatorio(driver):

    botao = WebDriverWait(
        driver,
        TIMEOUT
    ).until(
        EC.element_to_be_clickable(
            (
                By.XPATH,
                "//button[contains(@onclick, 'relatorio(0)')]"
            )
        )
    )

    driver.execute_script(
        "arguments[0].click();",
        botao
    )

    print(
        "✓ Gerando relatório..."
    )


# ============================================================
# EXTRAIR TABELA HTML
# ============================================================

def extrair_tabela(driver):

    print("\n" + "=" * 70)
    print("EXTRAINDO DADOS DO RELATÓRIO")
    print("=" * 70)

    print(
        "Aguardando tabela do relatório..."
    )

    tabela = None

    fim = time.time() + 120

    while time.time() < fim:

        tabelas = driver.find_elements(
            By.TAG_NAME,
            "table"
        )

        for item in tabelas:

            try:

                if not item.is_displayed():
                    continue

                linhas = item.find_elements(
                    By.CSS_SELECTOR,
                    "tbody tr"
                )

                if not linhas:

                    linhas = item.find_elements(
                        By.CSS_SELECTOR,
                        "tr"
                    )

                if len(linhas) > 1:

                    tabela = item

                    break

            except Exception:

                continue

        if tabela is not None:
            break

        time.sleep(2)

    if tabela is None:

        print(
            "Nenhuma tabela com dados foi encontrada."
        )

        print(
            f"URL atual: {driver.current_url}"
        )

        print(
            f"Título: {driver.title}"
        )

        try:

            driver.save_screenshot(
                "erro_relatorio_ociosidade.png"
            )

            print(
                "Screenshot salvo: "
                "erro_relatorio_ociosidade.png"
            )

        except Exception as erro:

            print(
                f"Erro ao salvar screenshot: {erro}"
            )

        try:

            with open(
                "erro_relatorio_ociosidade.html",
                "w",
                encoding="utf-8"
            ) as arquivo:

                arquivo.write(
                    driver.page_source
                )

            print(
                "HTML salvo: "
                "erro_relatorio_ociosidade.html"
            )

        except Exception as erro:

            print(
                f"Erro ao salvar HTML: {erro}"
            )

        raise TimeoutException(
            "A tabela do relatório não apareceu "
            "com dados dentro de 120 segundos."
        )

    # --------------------------------------------------------
    # CABEÇALHOS
    # --------------------------------------------------------

    cabecalhos = [

        elemento.text.strip()

        for elemento in tabela.find_elements(
            By.CSS_SELECTOR,
            "thead th"
        )

        if elemento.text.strip()
    ]

    # --------------------------------------------------------
    # LINHAS
    # --------------------------------------------------------

    linhas = tabela.find_elements(
        By.CSS_SELECTOR,
        "tbody tr"
    )

    if not linhas:

        todas_linhas = tabela.find_elements(
            By.CSS_SELECTOR,
            "tr"
        )

        if not cabecalhos and todas_linhas:

            cabecalhos = [

                elemento.text.strip()

                for elemento in todas_linhas[0].find_elements(
                    By.CSS_SELECTOR,
                    "th, td"
                )
            ]

            linhas = todas_linhas[1:]

        else:

            linhas = todas_linhas

    print(
        "Cabeçalhos encontrados:"
    )

    print(
        cabecalhos
    )

    # --------------------------------------------------------
    # EXTRAIR DADOS
    # --------------------------------------------------------

    dados = []

    for linha in linhas:

        try:

            colunas = linha.find_elements(
                By.TAG_NAME,
                "td"
            )

            valores = [

                coluna.text.strip()

                for coluna in colunas
            ]

            if not valores:
                continue

            if cabecalhos:

                while len(cabecalhos) < len(valores):

                    cabecalhos.append(
                        f"COLUNA_{len(cabecalhos) + 1}"
                    )

                registro = dict(
                    zip(
                        cabecalhos,
                        valores
                    )
                )

            else:

                registro = {

                    f"COLUNA_{i + 1}": valor

                    for i, valor
                    in enumerate(valores)
                }

            dados.append(
                registro
            )

        except Exception as erro:

            print(
                f"Erro ao ler linha: {erro}"
            )

    df = pd.DataFrame(
        dados
    )

    print(
        f"Registros encontrados: {len(df)}"
    )

    return df


# ============================================================
# NORMALIZAR NOME DE COLUNA
# ============================================================

def normalizar_coluna(valor):

    valor = str(valor)

    valor = (
        valor
        .strip()
        .lower()
    )

    valor = valor.replace(
        "á",
        "a"
    )

    valor = valor.replace(
        "à",
        "a"
    )

    valor = valor.replace(
        "ã",
        "a"
    )

    valor = valor.replace(
        "â",
        "a"
    )

    valor = valor.replace(
        "é",
        "e"
    )

    valor = valor.replace(
        "ê",
        "e"
    )

    valor = valor.replace(
        "í",
        "i"
    )

    valor = valor.replace(
        "ó",
        "o"
    )

    valor = valor.replace(
        "ô",
        "o"
    )

    valor = valor.replace(
        "õ",
        "o"
    )

    valor = valor.replace(
        "ú",
        "u"
    )

    valor = valor.replace(
        "ç",
        "c"
    )

    return valor


# ============================================================
# PREPARAR DADOS FINAIS
# ============================================================

def preparar_dados_finais(
    df,
    data_referencia
):

    print("\n" + "=" * 70)
    print("PREPARANDO DADOS FINAIS")
    print("=" * 70)

    if df.empty:

        print(
            "⚠ DataFrame vazio."
        )

        return pd.DataFrame(
            columns=[
                "Cliente",
                "Frota",
                "PLACA",
                "Distância (Km)",
                "Data"
            ]
        )

    df = df.copy()

    # --------------------------------------------------------
    # REMOVER ESPAÇOS DOS NOMES DAS COLUNAS
    # --------------------------------------------------------

    df.columns = [

        str(coluna).strip()

        for coluna in df.columns
    ]

    print(
        "\nColunas encontradas no HTML:"
    )

    for coluna in df.columns:

        print(
            f" - {coluna}"
        )

    # --------------------------------------------------------
    # IDENTIFICAR COLUNAS
    # --------------------------------------------------------

    mapa_colunas = {}

    for coluna in df.columns:

        coluna_normalizada = normalizar_coluna(
            coluna
        )

        # CLIENTE
        if (
            coluna_normalizada == "cliente"
            or "cliente" in coluna_normalizada
        ):

            mapa_colunas[
                "Cliente"
            ] = coluna

        # FROTA
        elif (
            coluna_normalizada == "frota"
            or "frota" in coluna_normalizada
        ):

            mapa_colunas[
                "Frota"
            ] = coluna

        # PLACA
        elif (
            coluna_normalizada == "placa"
            or coluna_normalizada == "placas"
            or "placa" in coluna_normalizada
        ):

            mapa_colunas[
                "PLACA"
            ] = coluna

        # DISTÂNCIA
        elif (
            "distancia" in coluna_normalizada
            or "km" in coluna_normalizada
        ):

            if (
                "Distância (Km)"
                not in mapa_colunas
            ):

                mapa_colunas[
                    "Distância (Km)"
                ] = coluna

    print(
        "\nMapeamento identificado:"
    )

    print(
        mapa_colunas
    )

    # --------------------------------------------------------
    # VALIDAR COLUNAS
    # --------------------------------------------------------

    colunas_obrigatorias = [

        "Cliente",
        "Frota",
        "PLACA",
        "Distância (Km)"
    ]

    faltando = [

        coluna

        for coluna
        in colunas_obrigatorias

        if coluna not in mapa_colunas
    ]

    if faltando:

        raise RuntimeError(
            "\nNão foi possível localizar as seguintes "
            "colunas no relatório HTML:\n\n"
            + ", ".join(faltando)
            + "\n\nColunas encontradas:\n"
            + str(
                df.columns.tolist()
            )
        )

    # --------------------------------------------------------
    # CRIAR DATAFRAME FINAL
    # --------------------------------------------------------

    df_final = pd.DataFrame()

    for coluna_final in colunas_obrigatorias:

        coluna_origem = mapa_colunas[
            coluna_final
        ]

        df_final[
            coluna_final
        ] = df[
            coluna_origem
        ]

    # --------------------------------------------------------
    # ADICIONAR DATA DO DIA CONSULTADO
    # --------------------------------------------------------

    df_final[
        "Data"
    ] = data_referencia.strftime(
        "%d/%m/%Y"
    )

    # --------------------------------------------------------
    # LIMPAR VALORES NULOS
    # --------------------------------------------------------

    df_final = (
        df_final
        .astype(object)
        .where(
            pd.notnull(df_final),
            ""
        )
    )

    # --------------------------------------------------------
    # REMOVER LINHAS COMPLETAMENTE VAZIAS
    # --------------------------------------------------------

    df_final = df_final[
        df_final[
            [
                "Cliente",
                "Frota",
                "PLACA",
                "Distância (Km)"
            ]
        ]
        .astype(str)
        .apply(
            lambda linha:
            any(
                valor.strip()
                for valor in linha
            ),
            axis=1
        )
    ]

    df_final = df_final.reset_index(
        drop=True
    )

    print(
        "\n✓ Dados preparados com sucesso."
    )

    print(
        f"Total de registros finais: {len(df_final)}"
    )

    print(
        "\nColunas finais:"
    )

    print(
        df_final.columns.tolist()
    )

    return df_final


# ============================================================
# PREPARAR DADOS PARA JSON
# ============================================================

def preparar_registros(df):

    df_envio = df.copy()

    # --------------------------------------------------------
    # CONVERTER DATAS
    # --------------------------------------------------------

    for coluna in df_envio.columns:

        if pd.api.types.is_datetime64_any_dtype(
            df_envio[coluna]
        ):

            df_envio[
                coluna
            ] = (

                df_envio[coluna]
                .dt.strftime(
                    "%d/%m/%Y %H:%M:%S"
                )
            )

    # --------------------------------------------------------
    # CONVERTER NaN E NaT
    # --------------------------------------------------------

    df_envio = (

        df_envio
        .astype(object)
        .where(
            pd.notnull(df_envio),
            ""
        )
    )

    # --------------------------------------------------------
    # CONVERTER PARA REGISTROS
    # --------------------------------------------------------

    registros = df_envio.to_dict(
        orient="records"
    )

    # Validação de serialização
    json.dumps(
        registros,
        default=str,
        ensure_ascii=False
    )

    return registros


# ============================================================
# ENVIAR PARA GOOGLE SHEETS
# ============================================================

def enviar_para_google(df):

    print("\n" + "=" * 70)
    print("ENVIANDO PARA GOOGLE SHEETS")
    print("=" * 70)

    if not GOOGLE_SCRIPT_URL:

        raise RuntimeError(
            "GOOGLE_SCRIPT_URL não configurada."
        )

    if df.empty:

        print(
            "⚠ Nenhum registro para enviar."
        )

        return False

    registros = preparar_registros(
        df
    )

    payload = {

        "acao": "ociosidade",

        "resultados": registros
    }

    # --------------------------------------------------------
    # SERIALIZAÇÃO SEGURA
    # --------------------------------------------------------

    corpo = json.dumps(
        payload,
        default=str,
        ensure_ascii=False
    )

    print(
        f"✓ Registros preparados: {len(registros)}"
    )

    # --------------------------------------------------------
    # ENVIAR PARA APPS SCRIPT
    # --------------------------------------------------------

    resposta = requests.post(

        GOOGLE_SCRIPT_URL,

        data=corpo.encode(
            "utf-8"
        ),

        headers={
            "Content-Type":
            "application/json; charset=utf-8"
        },

        timeout=120,

        allow_redirects=True
    )

    print(
        f"✓ HTTP: {resposta.status_code}"
    )

    print(
        "\nResposta do Apps Script:"
    )

    print(
        resposta.text
    )

    resposta.raise_for_status()

    # --------------------------------------------------------
    # INTERPRETAR RESPOSTA
    # --------------------------------------------------------

    try:

        retorno = resposta.json()

    except ValueError:

        raise RuntimeError(
            "O Apps Script não retornou JSON válido.\n\n"
            "Resposta recebida:\n"
            + resposta.text
        )

    if not retorno.get(
        "success",
        False
    ):

        raise RuntimeError(
            "Apps Script retornou erro:\n"
            + json.dumps(
                retorno,
                ensure_ascii=False,
                indent=2
            )
        )

    quantidade = retorno.get(
        "quantidade",
        len(registros)
    )

    print(
        f"\n✓ Registros gravados com sucesso: "
        f"{quantidade}"
    )

    return True


# ============================================================
# PROCESSO PRINCIPAL
# ============================================================

def executar():

    driver = None

    try:

        print("\n" + "=" * 70)
        print("AUTOMAÇÃO DE OCIOSIDADE")
        print("=" * 70)

        # ----------------------------------------------------
        # DATA ATUAL NO HORÁRIO DE BRASÍLIA
        # ----------------------------------------------------

        agora = datetime.now(
            FUSO_BRASILIA
        )

        print(
            "Data da execução:",
            agora.strftime(
                "%d/%m/%Y %H:%M:%S"
            )
        )

        # ----------------------------------------------------
        # DEFINIR DIA ANTERIOR
        # ----------------------------------------------------

        dia_anterior = (
            agora - timedelta(days=1)
        ).date()

        # ----------------------------------------------------
        # DATA INICIAL
        # DIA ANTERIOR ÀS 00:00:00
        # ----------------------------------------------------

        data_inicial = datetime(

            dia_anterior.year,
            dia_anterior.month,
            dia_anterior.day,

            0,
            0,
            0
        )

        # ----------------------------------------------------
        # DATA FINAL
        # DIA ANTERIOR ÀS 23:59:59
        # ----------------------------------------------------

        data_final = datetime(

            dia_anterior.year,
            dia_anterior.month,
            dia_anterior.day,

            23,
            59,
            59
        )

        print("\n" + "=" * 70)
        print("PERÍODO QUE SERÁ CONSULTADO")
        print("=" * 70)

        print(
            "Data inicial:",
            data_inicial.strftime(
                "%d/%m/%Y %H:%M:%S"
            )
        )

        print(
            "Data final:",
            data_final.strftime(
                "%d/%m/%Y %H:%M:%S"
            )
        )

        print(
            "Data que será gravada no Google Sheets:",
            dia_anterior.strftime(
                "%d/%m/%Y"
            )
        )

        # ----------------------------------------------------
        # INICIAR NAVEGADOR
        # ----------------------------------------------------

        driver = criar_driver()

        # ----------------------------------------------------
        # LOGIN
        # ----------------------------------------------------

        fazer_login(
            driver
        )

        # ----------------------------------------------------
        # ACESSAR OCIOSIDADE
        # ----------------------------------------------------

        acessar_ociosidade(
            driver
        )

        # ----------------------------------------------------
        # PREENCHER DATA INICIAL
        # ----------------------------------------------------

        preencher_datetime(

            driver,

            "data_inicial",

            data_inicial.strftime(
                "%Y-%m-%dT%H:%M"
            )
        )

        # ----------------------------------------------------
        # PREENCHER DATA FINAL
        # ----------------------------------------------------

        preencher_datetime(

            driver,

            "data_final",

            data_final.strftime(
                "%Y-%m-%dT%H:%M"
            )
        )

        # ----------------------------------------------------
        # PREENCHER VALOR = 0
        # ----------------------------------------------------

        preencher_valor(
            driver
        )

        # ----------------------------------------------------
        # SELECIONAR TODAS AS PLACAS
        # ----------------------------------------------------

        selecionar_todas_placas(
            driver
        )

        # ----------------------------------------------------
        # GERAR RELATÓRIO
        # ----------------------------------------------------

        gerar_relatorio(
            driver
        )

        # Aguarda processamento inicial
        time.sleep(5)

        # ----------------------------------------------------
        # EXTRAIR TABELA HTML
        # ----------------------------------------------------

        df = extrair_tabela(
            driver
        )

        # ----------------------------------------------------
        # PREPARAR APENAS AS COLUNAS NECESSÁRIAS
        # ----------------------------------------------------

        df = preparar_dados_finais(

            df,

            dia_anterior
        )

        print("\n" + "=" * 70)
        print("PRIMEIROS REGISTROS")
        print("=" * 70)

        print(
            df.head(10)
        )

        # ----------------------------------------------------
        # ENVIAR PARA GOOGLE SHEETS
        # ----------------------------------------------------

        if df.empty:

            print(
                "\n⚠ Nenhum dado encontrado."
            )

        else:

            enviar_para_google(
                df
            )

        print("\n" + "=" * 70)
        print("PROCESSO FINALIZADO COM SUCESSO")
        print("=" * 70)

    except Exception as erro:

        print("\n" + "=" * 70)
        print("ERRO NA AUTOMAÇÃO")
        print("=" * 70)

        print(
            str(erro)
        )

        raise

    finally:

        if driver:

            try:

                driver.quit()

                print(
                    "\n✓ Navegador fechado."
                )

            except Exception as erro:

                print(
                    f"\nErro ao fechar navegador: {erro}"
                )


# ============================================================
# EXECUÇÃO
# ============================================================

if __name__ == "__main__":

    executar()

