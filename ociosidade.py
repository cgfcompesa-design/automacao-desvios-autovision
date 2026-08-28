```python
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
from selenium.common.exceptions import (
    TimeoutException,
    StaleElementReferenceException
)


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

TEMPO_MAXIMO_RELATORIO = 180

FUSO_BRASILIA = ZoneInfo(
    "America/Sao_Paulo"
)

ID_TABELA_RELATORIO = (
    "tabela_excel_grafico_mais utilizados"
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
# FUNÇÃO AUXILIAR - MOSTRAR ABAS
# ============================================================

def mostrar_abas(driver):

    print("\nABAS ATUALMENTE ABERTAS:")

    abas = driver.window_handles

    for indice, aba in enumerate(abas, start=1):

        try:

            driver.switch_to.window(
                aba
            )

            print(
                f"{indice}. "
                f"Handle: {aba} | "
                f"URL: {driver.current_url} | "
                f"Título: {driver.title}"
            )

        except Exception as erro:

            print(
                f"{indice}. "
                f"Erro ao verificar aba: {erro}"
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

    print(
        "✓ Navegador iniciado."
    )

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

    print(
        "Login enviado..."
    )

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

    print(
        "✓ Login concluído."
    )


# ============================================================
# ACESSAR RELATÓRIO DE OCIOSIDADE
# ============================================================

def acessar_ociosidade(driver):

    print("\n" + "=" * 70)
    print("ACESSANDO RELATÓRIO DE OCIOSIDADE")
    print("=" * 70)

    driver.switch_to.default_content()

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

    print(
        "✓ Link do relatório clicado."
    )

    time.sleep(3)

    # --------------------------------------------------------
    # TENTA ENCONTRAR O FORMULÁRIO DENTRO DE IFRAME
    # --------------------------------------------------------

    frames = driver.find_elements(
        By.TAG_NAME,
        "iframe"
    )

    print(
        f"Quantidade de iframes encontrados: {len(frames)}"
    )

    for indice, frame in enumerate(frames):

        try:

            driver.switch_to.default_content()

            driver.switch_to.frame(
                frame
            )

            elementos = driver.find_elements(
                By.ID,
                "data_inicial"
            )

            if elementos:

                print(
                    f"✓ Relatório encontrado "
                    f"no iframe {indice + 1}."
                )

                return

        except Exception:

            continue

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

        campo.focus();
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
        const campo = arguments[0];

        campo.focus();
        campo.value = '0';

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
        campo
    )

    print(
        "✓ Valor preenchido: 0"
    )


# ============================================================
# SELECIONAR TODAS AS PLACAS
# ============================================================

def selecionar_todas_placas(driver):

    print("\n" + "=" * 70)
    print("SELECIONANDO TODAS AS PLACAS")
    print("=" * 70)

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
        "arguments[0].scrollIntoView({block: 'center'});",
        botao
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

    print("\n" + "=" * 70)
    print("GERANDO RELATÓRIO")
    print("=" * 70)

    # --------------------------------------------------------
    # GUARDA A ABA ATUAL E TODAS AS ABAS EXISTENTES
    # --------------------------------------------------------

    aba_origem = driver.current_window_handle

    abas_antes = driver.window_handles.copy()

    print(
        f"Quantidade de abas antes: {len(abas_antes)}"
    )

    print(
        f"Aba de origem: {aba_origem}"
    )

    # --------------------------------------------------------
    # ENCONTRAR BOTÃO
    # --------------------------------------------------------

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

    # --------------------------------------------------------
    # CLICAR
    # --------------------------------------------------------

    driver.execute_script(
        "arguments[0].scrollIntoView({block: 'center'});",
        botao
    )

    driver.execute_script(
        "arguments[0].click();",
        botao
    )

    print(
        "✓ Botão Gerar Relatório clicado."
    )

    # --------------------------------------------------------
    # ESPERAR:
    # 1. NOVA ABA
    # OU
    # 2. A TABELA CARREGAR NA MESMA ABA
    # --------------------------------------------------------

    fim = time.time() + TIMEOUT

    while time.time() < fim:

        abas_atuais = driver.window_handles

        # ----------------------------------------------------
        # CASO 1 - NOVA ABA ABERTA
        # ----------------------------------------------------

        if len(abas_atuais) > len(abas_antes):

            novas_abas = [
                aba
                for aba in abas_atuais
                if aba not in abas_antes
            ]

            if novas_abas:

                nova_aba = novas_abas[-1]

                driver.switch_to.default_content()

                driver.switch_to.window(
                    nova_aba
                )

                print(
                    "✓ Nova aba do relatório encontrada."
                )

                print(
                    f"URL: {driver.current_url}"
                )

                return nova_aba

        # ----------------------------------------------------
        # CASO 2 - RELATÓRIO ABRIU NA MESMA ABA
        # ----------------------------------------------------

        try:

            driver.switch_to.default_content()

            tabela = driver.find_elements(
                By.ID,
                ID_TABELA_RELATORIO
            )

            if tabela:

                print(
                    "✓ Relatório carregado "
                    "na mesma aba."
                )

                return driver.current_window_handle

        except Exception:

            pass

        # Volta para a aba original para continuar verificando
        try:

            driver.switch_to.window(
                aba_origem
            )

        except Exception:

            pass

        time.sleep(1)

    # --------------------------------------------------------
    # DIAGNÓSTICO
    # --------------------------------------------------------

    print(
        "\n⚠ Não foi detectada nova aba "
        "dentro do tempo esperado."
    )

    mostrar_abas(
        driver
    )

    raise TimeoutException(
        "O relatório não abriu em nova aba "
        "e a tabela também não apareceu "
        "na aba atual."
    )


# ============================================================
# EXTRAIR TABELA HTML
# ============================================================

def extrair_tabela(driver):

    print("\n" + "=" * 70)
    print("EXTRAINDO DADOS DO RELATÓRIO")
    print("=" * 70)

    print(
        f"URL atual: {driver.current_url}"
    )

    print(
        f"Título atual: {driver.title}"
    )

    print(
        f"Procurando tabela: {ID_TABELA_RELATORIO}"
    )

    # --------------------------------------------------------
    # AGUARDAR A TABELA
    # --------------------------------------------------------

    try:

        WebDriverWait(
            driver,
            TEMPO_MAXIMO_RELATORIO
        ).until(
            EC.presence_of_element_located(
                (
                    By.ID,
                    ID_TABELA_RELATORIO
                )
            )
        )

        print(
            "✓ Tabela encontrada."
        )

    except TimeoutException:

        print(
            "\nNenhuma tabela foi encontrada."
        )

        print(
            f"URL atual: {driver.current_url}"
        )

        print(
            f"Título: {driver.title}"
        )

        print(
            f"Quantidade de abas: "
            f"{len(driver.window_handles)}"
        )

        mostrar_abas(
            driver
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
            "A tabela do relatório não apareceu."
        )

    # --------------------------------------------------------
    # AGUARDAR DADOS REAIS NA TABELA
    # --------------------------------------------------------

    print(
        "Aguardando carregamento dos dados..."
    )

    fim = (
        time.time()
        + TEMPO_MAXIMO_RELATORIO
    )

    while time.time() < fim:

        try:

            tabela = driver.find_element(
                By.ID,
                ID_TABELA_RELATORIO
            )

            linhas = tabela.find_elements(
                By.CSS_SELECTOR,
                "tbody tr"
            )

            quantidade_valida = 0

            for linha in linhas:

                colunas = linha.find_elements(
                    By.TAG_NAME,
                    "td"
                )

                valores = [
                    coluna.text.strip()
                    for coluna in colunas
                ]

                # O relatório possui 4 colunas
                if len(valores) >= 4:

                    # Garante que exista algum valor
                    if any(valores):

                        quantidade_valida += 1

            if quantidade_valida > 0:

                print(
                    f"✓ Dados carregados: "
                    f"{quantidade_valida} registros."
                )

                break

        except StaleElementReferenceException:

            pass

        except Exception:

            pass

        time.sleep(2)

    else:

        raise TimeoutException(
            "A tabela apareceu, mas nenhum "
            "registro válido foi carregado "
            f"em {TEMPO_MAXIMO_RELATORIO} segundos."
        )

    # --------------------------------------------------------
    # EXTRAIR CABEÇALHOS
    # --------------------------------------------------------

    tabela = driver.find_element(
        By.ID,
        ID_TABELA_RELATORIO
    )

    cabecalhos = [
        elemento.text.strip()
        for elemento in tabela.find_elements(
            By.CSS_SELECTOR,
            "thead th"
        )
        if elemento.text.strip()
    ]

    print(
        "Cabeçalhos encontrados:"
    )

    print(
        cabecalhos
    )

    # --------------------------------------------------------
    # VALIDAR CABEÇALHOS
    # --------------------------------------------------------

    if len(cabecalhos) < 4:

        raise RuntimeError(
            "A tabela encontrada não possui "
            "os quatro cabeçalhos esperados. "
            f"Cabeçalhos encontrados: {cabecalhos}"
        )

    # --------------------------------------------------------
    # EXTRAIR LINHAS
    # --------------------------------------------------------

    dados = []

    linhas = tabela.find_elements(
        By.CSS_SELECTOR,
        "tbody tr"
    )

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

            # Ignora linhas vazias
            if not valores:
                continue

            if not any(valores):
                continue

            # Precisa possuir pelo menos 4 colunas
            if len(valores) < 4:
                continue

            registro = {
                "Cliente": valores[0],
                "Frota": valores[1],
                "PLACA": valores[2],
                "Distância (Km)": valores[3]
            }

            dados.append(
                registro
            )

        except StaleElementReferenceException:

            continue

        except Exception as erro:

            print(
                f"Erro ao ler linha: {erro}"
            )

    df = pd.DataFrame(
        dados,
        columns=[
            "Cliente",
            "Frota",
            "PLACA",
            "Distância (Km)"
        ]
    )

    print(
        f"✓ Total de registros extraídos: {len(df)}"
    )

    return df


# ============================================================
# ADICIONAR DATA DE REFERÊNCIA
# ============================================================

def adicionar_data(
    df,
    data_referencia
):

    df = df.copy()

    df["Data"] = data_referencia.strftime(
        "%d/%m/%Y"
    )

    return df


# ============================================================
# LIMPAR DADOS FINAIS
# ============================================================

def preparar_dados_finais(
    df,
    data_referencia
):

    print("\n" + "=" * 70)
    print("PREPARANDO DADOS FINAIS")
    print("=" * 70)

    colunas_finais = [
        "Cliente",
        "Frota",
        "PLACA",
        "Distância (Km)",
        "Data"
    ]

    if df.empty:

        print(
            "⚠ Nenhum dado foi extraído."
        )

        return pd.DataFrame(
            columns=colunas_finais
        )

    df = df.copy()

    # --------------------------------------------------------
    # GARANTIR AS COLUNAS ESPERADAS
    # --------------------------------------------------------

    for coluna in [
        "Cliente",
        "Frota",
        "PLACA",
        "Distância (Km)"
    ]:

        if coluna not in df.columns:

            df[coluna] = ""

    # --------------------------------------------------------
    # MANTER APENAS AS COLUNAS NECESSÁRIAS
    # --------------------------------------------------------

    df = df[
        [
            "Cliente",
            "Frota",
            "PLACA",
            "Distância (Km)"
        ]
    ].copy()

    # --------------------------------------------------------
    # LIMPAR ESPAÇOS
    # --------------------------------------------------------

    for coluna in df.columns:

        df[coluna] = (
            df[coluna]
            .fillna("")
            .astype(str)
            .str.strip()
        )

    # --------------------------------------------------------
    # REMOVER LINHAS SEM PLACA
    # --------------------------------------------------------

    df = df[
        df["PLACA"] != ""
    ].copy()

    # --------------------------------------------------------
    # ADICIONAR DATA DO DIA CONSULTADO
    # --------------------------------------------------------

    df = adicionar_data(
        df,
        data_referencia
    )

    # --------------------------------------------------------
    # ORDEM FINAL
    # --------------------------------------------------------

    df = df[
        colunas_finais
    ].reset_index(
        drop=True
    )

    print(
        f"✓ Registros finais: {len(df)}"
    )

    print(
        "✓ Colunas finais:"
    )

    print(
        df.columns.tolist()
    )

    return df


# ============================================================
# PREPARAR DADOS PARA JSON
# ============================================================

def preparar_registros(df):

    df_envio = df.copy()

    # --------------------------------------------------------
    # CONVERTER DATAS PANDAS
    # --------------------------------------------------------

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

    registros = df_envio.to_dict(
        orient="records"
    )

    # Validação
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

    corpo = json.dumps(
        payload,
        default=str,
        ensure_ascii=False
    )

    print(
        f"✓ Registros preparados: {len(registros)}"
    )

    try:

        resposta = requests.post(
            GOOGLE_SCRIPT_URL,
            data=corpo.encode("utf-8"),
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
            "Resposta do Apps Script:"
        )

        print(
            resposta.text
        )

        resposta.raise_for_status()

    except requests.RequestException as erro:

        raise RuntimeError(
            "Erro ao enviar dados para "
            f"Google Apps Script: {erro}"
        )

    try:

        retorno = resposta.json()

    except ValueError:

        raise RuntimeError(
            "O Apps Script não retornou "
            "um JSON válido.\n\n"
            f"Resposta:\n{resposta.text}"
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
        f"✓ Registros gravados: {quantidade}"
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
        # DATA E HORA ATUAL - BRASÍLIA
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
        # CALCULAR DIA ANTERIOR
        # ----------------------------------------------------

        dia_anterior = (
            agora - timedelta(days=1)
        ).date()

        # ----------------------------------------------------
        # PERÍODO DO DIA ANTERIOR
        # ----------------------------------------------------

        data_inicial = datetime(
            dia_anterior.year,
            dia_anterior.month,
            dia_anterior.day,
            0,
            0,
            0
        )

        data_final = datetime(
            dia_anterior.year,
            dia_anterior.month,
            dia_anterior.day,
            23,
            59,
            0
        )

        print("\n" + "=" * 70)
        print("PERÍODO CONSULTADO")
        print("=" * 70)

        print(
            "Data inicial:",
            data_inicial.strftime(
                "%d/%m/%Y %H:%M"
            )
        )

        print(
            "Data final:",
            data_final.strftime(
                "%d/%m/%Y %H:%M"
            )
        )

        print(
            "Data que será gravada:",
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
        # ACESSAR RELATÓRIO OCIOSIDADE
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
        # VALOR = 0
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
        #
        # A FUNÇÃO DETECTA SE O RELATÓRIO ABRIU:
        # - EM UMA NOVA ABA
        # - OU NA MESMA ABA
        # ----------------------------------------------------

        gerar_relatorio(
            driver
        )

        # Pequena pausa para estabilidade
        time.sleep(3)

        # ----------------------------------------------------
        # EXTRAIR DADOS
        # ----------------------------------------------------

        df = extrair_tabela(
            driver
        )

        # ----------------------------------------------------
        # PREPARAR DADOS FINAIS
        # ----------------------------------------------------

        df = preparar_dados_finais(
            df,
            dia_anterior
        )

        print("\n" + "=" * 70)
        print("PRIMEIROS REGISTROS")
        print("=" * 70)

        print(
            df.head(10).to_string(
                index=False
            )
        )

        # ----------------------------------------------------
        # VALIDAR RESULTADOS
        # ----------------------------------------------------

        if df.empty:

            print(
                "\n⚠ Nenhum dado encontrado."
            )

            print(
                "O processo não enviará "
                "uma planilha vazia."
            )

        else:

            # ------------------------------------------------
            # ENVIAR PARA GOOGLE SHEETS
            # ------------------------------------------------

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
                    f"Erro ao fechar navegador: {erro}"
                )


# ============================================================
# EXECUÇÃO
# ============================================================

if __name__ == "__main__":

    executar()
```
