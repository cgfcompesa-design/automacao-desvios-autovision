
import os
import time
import json
import sys
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
# CORRIGIR CODIFICAÇÃO DO CONSOLE
# ============================================================

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")


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

    print("
" + "=" * 70)
    print("LOGIN AUTOVISION")
    print("=" * 70)

    if not USUARIO:
        raise RuntimeError("AUTOVISION_USUARIO não configurado.")

    if not SENHA:
        raise RuntimeError("AUTOVISION_SENHA não configurada.")

    driver.switch_to.default_content()
    driver.get(URL_LOGIN)

    wait = WebDriverWait(driver, TIMEOUT)

    campo_usuario = wait.until(
        EC.visibility_of_element_located((By.ID, "usuario"))
    )
    campo_usuario.clear()
    campo_usuario.send_keys(USUARIO)

    campo_senha = wait.until(
        EC.visibility_of_element_located((By.ID, "senha"))
    )
    campo_senha.clear()
    campo_senha.send_keys(SENHA)

    botao_login = wait.until(
        EC.element_to_be_clickable(
            (By.CSS_SELECTOR, "button[type='submit']")
        )
    )

    driver.execute_script("arguments[0].click();", botao_login)

    print("Login enviado...")

    try:
        wait.until(
            lambda d: len(d.find_elements(By.ID, "usuario")) == 0
        )

        print("Aguardando carregamento do sistema...")
        time.sleep(5)

        WebDriverWait(driver, TIMEOUT).until(
            lambda d: d.execute_script(
                "return document.readyState"
            ) == "complete"
        )

    except TimeoutException:
        raise RuntimeError(
            "Login não concluído ou página não terminou de carregar."
        )

    print("✓ Login concluído.")
    print(f"URL após login: {driver.current_url}")


def acessar_ociosidade(driver):

    print("
" + "=" * 70)
    print("ACESSANDO RELATÓRIO DE OCIOSIDADE")
    print("=" * 70)

    driver.switch_to.default_content()

    print("Aguardando AutoVision carregar...")
    time.sleep(5)

    acessou = False

    seletores = [
        'a[href*="relatorio_ociosidade.php"]',
        'a[href*="relatorio_ociosidade"]',
        'a[href*="ociosidade"]',
    ]

    print("Tentativa 1: procurando link de Ociosidade...")

    try:
        for seletor in seletores:
            elementos = driver.find_elements(
                By.CSS_SELECTOR,
                seletor
            )

            if elementos:
                print(f"✓ Link encontrado: {seletor}")

                driver.execute_script(
                    "arguments[0].click();",
                    elementos[0]
                )

                acessou = True
                time.sleep(5)
                break

    except Exception as erro:
        print(f"⚠ Erro ao procurar link: {erro}")

    if not acessou:
        print("Tentativa 2: acessando relatório diretamente...")

        driver.switch_to.default_content()

        driver.get(
            "https://www.autovision.com.br/v3/"
            "modulos/relatorios/"
            "relatorio_ociosidade.php"
        )

        time.sleep(5)

    driver.switch_to.default_content()

    try:
        print("Verificando página principal...")

        WebDriverWait(driver, 20).until(
            lambda d: len(
                d.find_elements(By.ID, "data_inicial")
            ) > 0
        )

        print("✓ Página de Ociosidade encontrada.")
        return

    except TimeoutException:
        print("Campo não encontrado na página principal.")

    driver.switch_to.default_content()

    frames = driver.find_elements(By.TAG_NAME, "iframe")

    print(f"Frames encontrados: {len(frames)}")

    for indice, frame in enumerate(frames):
        try:
            driver.switch_to.default_content()
            driver.switch_to.frame(frame)

            print(f"Verificando iframe {indice + 1}...")

            WebDriverWait(driver, 10).until(
                lambda d: len(
                    d.find_elements(By.ID, "data_inicial")
                ) > 0
            )

            print(
                f"✓ Página de Ociosidade encontrada "
                f"no iframe {indice + 1}."
            )

            return

        except TimeoutException:
            continue

        except Exception as erro:
            print(
                f"Erro no iframe {indice + 1}: {erro}"
            )

    driver.switch_to.default_content()

    print("
" + "=" * 70)
    print("ERRO: PÁGINA DE OCIOSIDADE NÃO ENCONTRADA")
    print("=" * 70)

    print(f"URL atual: {driver.current_url}")
    print(f"Título da página: {driver.title}")
    print(f"Quantidade de iframes: {len(frames)}")

    raise RuntimeError(
        "Não foi possível localizar a página "
        "ou os campos do relatório de Ociosidade."
    )


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

    print("
" + "=" * 70)
    print("EXTRAINDO DADOS DO RELATÓRIO")
    print("=" * 70)

    # O relatório pode carregar lentamente e a tabela pode ser criada
    # depois do clique em "Gerar relatório".
    print("Aguardando tabela do relatório...")

    seletores_tabela = [
        "//table[contains(@id, 'tabela_excel')]",
        "//table[contains(@id, 'excel')]",
        "//table",
    ]

    tabela = None

    # Aguarda até 120 segundos, verificando todos os seletores.
    fim = time.time() + 120

    while time.time() < fim:

        for seletor in seletores_tabela:

            tabelas = driver.find_elements(
                By.XPATH,
                seletor
            )

            for item in tabelas:

                try:

                    if not item.is_displayed():
                        continue

                    linhas = item.find_elements(
                        By.CSS_SELECTOR,
                        "tbody tr"
                    )

                    if linhas:

                        tabela = item
                        break

                except Exception:

                    continue

            if tabela is not None:
                break

        if tabela is not None:
            break

        time.sleep(2)

    # ------------------------------------------------------------
    # SE NÃO ENCONTRAR, SALVA DIAGNÓSTICOS
    # ------------------------------------------------------------

    if tabela is None:

        print("⚠ Nenhuma tabela com linhas foi encontrada.")
        print(f"URL atual: {driver.current_url}")
        print(f"Título: {driver.title}")

        try:
            driver.save_screenshot(
                "erro_relatorio_ociosidade.png"
            )

            print(
                "✓ Screenshot salvo: "
                "erro_relatorio_ociosidade.png"
            )

        except Exception as erro:
            print(
                f"Não foi possível salvar screenshot: {erro}"
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
                "✓ HTML salvo: "
                "erro_relatorio_ociosidade.html"
            )

        except Exception as erro:
            print(
                f"Não foi possível salvar HTML: {erro}"
            )

        # Mostra todas as tabelas existentes para diagnóstico
        todas_tabelas = driver.find_elements(
            By.TAG_NAME,
            "table"
        )

        print(
            f"Total de tabelas encontradas: "
            f"{len(todas_tabelas)}"
        )

        for indice, item in enumerate(
            todas_tabelas[:10],
            start=1
        ):

            try:

                print(
                    f"Tabela {indice}: "
                    f"id='{item.get_attribute('id')}' "
                    f"class='{item.get_attribute('class')}'"
                )

            except Exception:

                pass

        raise TimeoutException(
            "A tabela do relatório não apareceu "
            "com registros dentro de 120 segundos."
        )

    # ------------------------------------------------------------
    # EXTRAIR CABEÇALHOS
    # ------------------------------------------------------------

    cabecalhos = [

        elemento.text.strip()

        for elemento in tabela.find_elements(
            By.CSS_SELECTOR,
            "thead th"
        )

        if elemento.text.strip()
    ]

    # Caso não exista thead, tenta primeira linha
    if not cabecalhos:

        primeira_linha = tabela.find_elements(
            By.CSS_SELECTOR,
            "tr"
        )

        if primeira_linha:

            cabecalhos = [

                elemento.text.strip()

                for elemento in primeira_linha[0].find_elements(
                    By.CSS_SELECTOR,
                    "th, td"
                )

            ]

    print("Cabeçalhos encontrados:")
    print(cabecalhos)

    dados = []

    linhas = tabela.find_elements(
        By.CSS_SELECTOR,
        "tbody tr"
    )

    # Caso a tabela não use tbody
    if not linhas:

        todas_linhas = tabela.find_elements(
            By.CSS_SELECTOR,
            "tr"
        )

        if cabecalhos and todas_linhas:
            linhas = todas_linhas[1:]
        else:
            linhas = todas_linhas

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

                # Se houver menos cabeçalhos que valores, cria nomes extras
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
                    for i, valor in enumerate(valores)
                }

            dados.append(registro)

        except Exception as erro:

            print(
                f"⚠ Erro ao ler uma linha: {erro}"
            )

    df = pd.DataFrame(dados)

    print(
        f"✓ Registros encontrados: {len(df)}"
    )

    return df


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

        # A extração aguarda a tabela por até 120 segundos
        time.sleep(2)

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
