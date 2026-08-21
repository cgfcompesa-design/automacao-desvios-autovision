import os
import sys
import traceback
from datetime import datetime

import pandas as pd
import requests

from autovision import Autovision
from desvios import processar_desvios


# ============================================================
# CONFIGURAÇÕES
# ============================================================

CSV_URL = (
    "https://docs.google.com/spreadsheets/d/e/"
    "2PACX-1vQU22jWG-2LuGLDuLYvVQoBbv5JdArq9WUwGcJ3znHPZWHrFABji3IFNFwYCtVX7u8uo-Rd7YJFb9fZ/"
    "pub?gid=756259345&single=true&output=csv"
)

APPS_SCRIPT_URL = os.getenv(
    "APPS_SCRIPT_URL",
    "https://script.google.com/macros/s/"
    "AKfycbwiKp58t3hyHHXXnjqLBbLcrY1o8pupLMOSLPSxOqpL/dev"
)

TIMEOUT_HTTP = 60


# ============================================================
# LOG
# ============================================================

def log(mensagem):
    agora = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    print(f"[{agora}] {mensagem}", flush=True)


# ============================================================
# NORMALIZAÇÃO DE PLACA
# ============================================================

def normalizar_placa(valor):

    if pd.isna(valor):
        return ""

    placa = str(valor).strip().upper()

    placa = placa.replace("-", "")
    placa = placa.replace(" ", "")

    return placa


# ============================================================
# BAIXAR GOOGLE SHEETS
# ============================================================

def baixar_google_sheets():

    log("Baixando CSV da planilha TELEMETRIA_OCIOSIDADE...")

    resposta = requests.get(
        CSV_URL,
        timeout=TIMEOUT_HTTP
    )

    resposta.raise_for_status()

    texto = resposta.content.decode("utf-8-sig")

    if not texto.strip():
        raise RuntimeError(
            "O CSV retornado pelo Google Sheets está vazio."
        )

    from io import StringIO

    df = pd.read_csv(
        StringIO(texto),
        dtype=str
    )

    df.columns = [
        str(col).strip()
        for col in df.columns
    ]

    log(
        f"CSV carregado com sucesso: "
        f"{len(df)} linhas."
    )

    return df


# ============================================================
# VALIDAR PLANILHA
# ============================================================

def validar_planilha(df):

    colunas_obrigatorias = [
        "PLACA"
    ]

    faltantes = [
        coluna
        for coluna in colunas_obrigatorias
        if coluna not in df.columns
    ]

    if faltantes:

        raise RuntimeError(
            "Colunas obrigatórias não encontradas: "
            + ", ".join(faltantes)
            + "\n"
            + "Colunas encontradas: "
            + ", ".join(df.columns)
        )


# ============================================================
# LER PLACAS DA ABA GERAL
# ============================================================

def obter_placas(df):

    validar_planilha(df)

    placas = (
        df["PLACA"]
        .apply(normalizar_placa)
    )

    placas = [
        placa
        for placa in placas.tolist()
        if placa
    ]

    # Remove duplicadas mantendo a ordem
    placas_unicas = list(
        dict.fromkeys(placas)
    )

    log(
        f"Placas encontradas na aba Geral: "
        f"{len(placas_unicas)}"
    )

    for placa in placas_unicas:
        log(f"  - {placa}")

    return placas_unicas


# ============================================================
# ENVIAR RESULTADOS PARA APPS SCRIPT
# ============================================================

def enviar_resultados(resultados):

    log("Enviando resultados para o Apps Script...")

    payload = {
        "acao": "desvios",
        "resultados": resultados
    }

    resposta = requests.post(
        APPS_SCRIPT_URL,
        json=payload,
        timeout=TIMEOUT_HTTP
    )

    resposta.raise_for_status()

    try:
        retorno = resposta.json()
    except Exception:
        retorno = {
            "texto": resposta.text
        }

    log(
        "Resposta do Apps Script:"
    )

    log(
        str(retorno)
    )

    if isinstance(retorno, dict):

        if retorno.get("success") is False:

            raise RuntimeError(
                "Apps Script retornou erro: "
                + str(retorno)
            )

    return retorno


# ============================================================
# PROCESSAR UMA PLACA
# ============================================================

def processar_placa(
    autovision,
    placa,
    linha_planilha
):

    log("=" * 70)

    log(
        f"Processando placa: {placa}"
    )

    try:

        dados_autovision = (
            autovision.consultar_placa(
                placa=placa,
                data=linha_planilha.get("Data"),
            )
        )

        if dados_autovision is None:

            log(
                f"Nenhum dado encontrado para {placa}"
            )

            return {
                "PLACA": placa,
                "status": "SEM_DADOS",
                "desvios": []
            }

        desvios = processar_desvios(
            dados_autovision
        )

        return {
            "PLACA": placa,
            "status": "OK",
            "desvios": desvios
        }

    except Exception as exc:

        log(
            f"ERRO na placa {placa}: "
            f"{exc}"
        )

        traceback.print_exc()

        return {
            "PLACA": placa,
            "status": "ERRO",
            "erro": str(exc),
            "desvios": []
        }


# ============================================================
# MAIN
# ============================================================

def main():

    inicio = datetime.now()

    log("=" * 70)
    log("INICIANDO AUTOMAÇÃO DE DESVIOS")
    log("=" * 70)

    try:

        # ----------------------------------------------------
        # 1. BAIXAR GOOGLE SHEETS
        # ----------------------------------------------------

        df = baixar_google_sheets()

        # ----------------------------------------------------
        # 2. MOSTRAR COLUNAS
        # ----------------------------------------------------

        log(
            "Colunas encontradas:"
        )

        for coluna in df.columns:
            log(
                f"  - {coluna}"
            )

        # ----------------------------------------------------
        # 3. OBTER PLACAS
        # ----------------------------------------------------

        placas = obter_placas(df)

        if not placas:

            raise RuntimeError(
                "Nenhuma placa foi encontrada "
                "na coluna PLACA da aba Geral."
            )

        # ----------------------------------------------------
        # 4. INICIAR AUTOVISION
        # ----------------------------------------------------

        log(
            "Inicializando Autovision..."
        )

        autovision = Autovision()

        autovision.iniciar()

        # ----------------------------------------------------
        # 5. PROCESSAR PLACAS
        # ----------------------------------------------------

        resultados = []

        for placa in placas:

            registros = df[
                df["PLACA"]
                .apply(normalizar_placa)
                == placa
            ]

            if registros.empty:

                linha = {}

            else:

                linha = (
                    registros.iloc[0]
                    .fillna("")
                    .to_dict()
                )

            resultado = processar_placa(
                autovision,
                placa,
                linha
            )

            resultados.append(
                resultado
            )

        # ----------------------------------------------------
        # 6. ENCERRAR AUTOVISION
        # ----------------------------------------------------

        try:

            autovision.fechar()

        except Exception as exc:

            log(
                "Aviso ao fechar Autovision: "
                + str(exc)
            )

        # ----------------------------------------------------
        # 7. ENVIAR RESULTADOS
        # ----------------------------------------------------

        enviar_resultados(
            resultados
        )

        # ----------------------------------------------------
        # FINAL
        # ----------------------------------------------------

        duracao = (
            datetime.now() - inicio
        )

        log("=" * 70)

        log(
            "AUTOMAÇÃO FINALIZADA COM SUCESSO"
        )

        log(
            f"Placas processadas: "
            f"{len(resultados)}"
        )

        log(
            f"Tempo total: "
            f"{duracao}"
        )

        log("=" * 70)

    except Exception as exc:

        log("=" * 70)
        log("AUTOMAÇÃO FINALIZADA COM ERRO")
        log("=" * 70)

        log(
            str(exc)
        )

        traceback.print_exc()

        sys.exit(1)


if __name__ == "__main__":
    main()
