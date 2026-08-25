
import os
import time
import json
from datetime import datetime

import pandas as pd
import requests

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException


# ============================================================
# CONFIGURAÇÕES
# ============================================================

URL_LOGIN = "https://www.autovision.com.br/v3/"

TIMEOUT = 60


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

    driver = webdriver.Edge(options=options)

    driver.set_page_load_timeout(TIMEOUT)

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

    driver.get(URL_LOGIN)

    wait = WebDriverWait(
        driver,
        TIMEOUT
    )

    campo_usuario = wait.until(
        EC.visibility_of_element_located(
            (By.ID, "usuario")
        )
    )

    campo_usuario.clear()

    campo_usuario.send_keys(
        USUARIO
    )

    campo_senha = wait.until(
        EC.visibility_of_element_located(
            (By.ID, "senha")
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
# ACESSAR RELATÓRIO OCIOSIDADE
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

    # Tenta encontrar o conteúdo dentro de iframe
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

    # Caso não esteja dentro de iframe
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

    tabela = WebDriverWait(
        driver,
        TIMEOUT
    ).until(
        EC.presence_of_element_located(
            (
                By.XPATH,
                "//table[contains(@id, 'tabela_excel')]"
            )
        )
    )

    # Aguarda até existir pelo menos uma linha
    WebDriverWait(
        driver,
        TIMEOUT
    ).until(
        lambda d: len(
            tabela.find_elements(
                By.CSS_SELECTOR,
                "tbody tr"
            )
        ) > 0
    )

    cabecalhos = [
        elemento.text.strip()
        for elemento in tabela.find_elements(
            By.CSS_SELECTOR,
            "thead th"
        )
    ]

    print("Cabeçalhos encontrados:")

    print(cabecalhos)

    dados = []

    linhas = tabela.find_elements(
        By.CSS_SELECTOR,
        "tbody tr"
    )

    for linha in linhas:

        colunas = linha.find_elements(
            By.TAG_NAME,
            "td"
        )

        valores = [
            coluna.text.strip()
            for coluna in colunas
        ]

        if len(valores) == len(cabecalhos):

            registro = dict(
                zip(
                    cabecalhos,
                    valores
                )
            )

            dados.append(
                registro
            )

    df = pd.DataFrame(
        dados
    )

    print(
        f"✓ Registros encontrados: {len(df)}"
    )

    return df


# ============================================================
# ADICIONAR DATA
# ============================================================

def adicionar_data(df):

    df = df.copy()

    df["Data"] = datetime.now().strftime(
        "%d/%m/%Y"
    )

    return df


# ============================================================
# PREPARAR DADOS PARA JSON
# ============================================================

def preparar_registros(df):

    df_envio = df.copy()

    # Converter datas
    for coluna in df_envio.columns:

        if pd.api.types.is_datetime64_any_dtype(
            df_envio[coluna]
        ):

            df_envio[coluna] = (
                df_envio[coluna]
                .dt.strftime(
                    "%d/%m/%Y %H:%M:%S"
                )
            )

    # Converter NaN/NaT
    df_envio = (
        df_envio
        .astype(object)
        .where(
            pd.notnull(df_envio),
            ""
        )
    )

    registros = df_envio.to_dict(
        orient="records"
    )

    # Garante que qualquer tipo não serializável
    # seja convertido para texto
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

    registros = preparar_registros(
        df
    )

    payload = {

        "acao": "ociosidade",

        "resultados": registros

    }

    # Serialização segura
    corpo = json.dumps(
        payload,
        default=str,
        ensure_ascii=False
    )

    print(
        f"✓ Registros preparados: {len(registros)}"
    )

    resposta = requests.post(
        GOOGLE_SCRIPT_URL,
        data=corpo.encode("utf-8"),
        headers={
            "Content-Type":
                "application/json; charset=utf-8"
        },
        timeout=120
    )

    print(
        f"✓ HTTP: {resposta.status_code}"
    )

    print(
        "Resposta:"
    )

    print(
        resposta.text
    )

    resposta.raise_for_status()

    retorno = resposta.json()

    if not retorno.get(
        "success",
        False
    ):

        raise RuntimeError(
            "Apps Script retornou erro: "
            + str(retorno)
        )

    print(
        f"✓ Registros gravados: "
        f"{retorno.get('quantidade', 0)}"
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

        print(
            "Data da execução:",
            datetime.now().strftime(
                "%d/%m/%Y %H:%M:%S"
            )
        )

        # ----------------------------------------------------
        # NAVEGADOR
        # ----------------------------------------------------

        driver = criar_driver()

        # ----------------------------------------------------
        # LOGIN
        # ----------------------------------------------------

        fazer_login(
            driver
        )

        # ----------------------------------------------------
        # OCIOSIDADE
        # ----------------------------------------------------

        acessar_ociosidade(
            driver
        )

        # ----------------------------------------------------
        # DEFINIR PERÍODO
        # ----------------------------------------------------

        agora = datetime.now()

        data_inicial = agora.replace(
            hour=0,
            minute=0,
            second=0,
            microsecond=0
        )

        data_final = agora.replace(
            hour=23,
            minute=59,
            second=0,
            microsecond=0
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
        # VALOR = 0
        # ----------------------------------------------------

        preencher_valor(
            driver
        )

        # ----------------------------------------------------
        # TODAS AS PLACAS
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

        # Aguarda o AutoVision processar
        time.sleep(5)

        # ----------------------------------------------------
        # EXTRAIR DADOS
        # ----------------------------------------------------

        df = extrair_tabela(
            driver
        )

        # ----------------------------------------------------
        # ADICIONAR DATA
        # ----------------------------------------------------

        df = adicionar_data(
            df
        )

        print("\nPRIMEIROS REGISTROS:")

        print(
            df.head()
        )

        # ----------------------------------------------------
        # ENVIAR PARA GOOGLE
        # ----------------------------------------------------

        if df.empty:

            print(
                "⚠ Nenhum dado encontrado."
            )

        else:

            enviar_para_google(
                df
            )

        print("\n" + "=" * 70)
        print("✅ PROCESSO FINALIZADO COM SUCESSO")
        print("=" * 70)

    except Exception as erro:

        print("\n" + "=" * 70)
        print("❌ ERRO NA AUTOMAÇÃO")
        print("=" * 70)

        print(
            str(erro)
        )

        raise

    finally:

        if driver:

            driver.quit()

            print(
                "\n✓ Navegador fechado."
            )


# ============================================================
# EXECUÇÃO
# ============================================================

if __name__ == "__main__":

    executar()
