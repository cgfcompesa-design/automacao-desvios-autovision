
import math
import re
import unicodedata

import pandas as pd


# ============================================================
# UTILITÁRIOS
# ============================================================

def texto(valor):

    if valor is None:
        return ""

    try:

        if pd.isna(valor):
            return ""

    except Exception:

        pass

    return str(valor).strip()


# ============================================================
# NORMALIZAR TEXTO
# ============================================================

def normalizar_texto(valor):

    valor = texto(valor)

    if not valor:
        return ""

    valor = unicodedata.normalize(
        "NFKD",
        valor
    )

    valor = "".join(
        caractere
        for caractere in valor
        if not unicodedata.combining(
            caractere
        )
    )

    valor = valor.upper()

    valor = re.sub(
        r"\s+",
        " ",
        valor
    )

    return valor.strip()


# ============================================================
# MOTORISTA VÁLIDO
# ============================================================

def motorista_valido(valor):

    valor = normalizar_texto(
        valor
    )

    if not valor:
        return False

    invalidos = {
        "NAN",
        "SEM MOTORISTA",
        "SEM MOTORISTA INFORMADO",
        "NAO INFORMADO",
        "NÃO INFORMADO",
        "NONE",
        "NULL",
        "-"
    }

    return valor not in invalidos


# ============================================================
# ENDEREÇO VÁLIDO
# ============================================================

def endereco_valido(valor):

    valor = normalizar_texto(
        valor
    )

    if not valor:
        return False

    invalidos = {
        "NAN",
        "SEM ENDERECO",
        "SEM ENDEREÇO",
        "NAO INFORMADO",
        "NÃO INFORMADO",
        "NONE",
        "NULL",
        "-"
    }

    return valor not in invalidos


# ============================================================
# MOTORISTAS DIFERENTES
# ============================================================

def identificar_motoristas_diferentes(
    registros
):

    motoristas = []

    for registro in registros:

        if not isinstance(
            registro,
            dict
        ):
            continue

        motorista = registro.get(
            "Motorista",
            registro.get(
                "motorista",
                ""
            )
        )

        if not motorista_valido(
            motorista
        ):
            continue

        motorista_normalizado = (
            normalizar_texto(
                motorista
            )
        )

        if motorista_normalizado not in motoristas:

            motoristas.append(
                motorista_normalizado
            )

    if len(motoristas) <= 1:

        return None

    return {
        "tipo": "Motorista diferente",
        "descricao": (
            "Foram identificados "
            "motoristas diferentes "
            "para a mesma placa."
        ),
        "motoristas": motoristas,
        "quantidade": len(motoristas)
    }


# ============================================================
# ENDEREÇOS DIFERENTES
# ============================================================

def identificar_enderecos_diferentes(
    registros
):

    enderecos = []

    enderecos_originais = {}

    for registro in registros:

        if not isinstance(
            registro,
            dict
        ):
            continue

        endereco = registro.get(
            "Endereço",
            registro.get(
                "endereco",
                ""
            )
        )

        if not endereco_valido(
            endereco
        ):
            continue

        endereco_normalizado = (
            normalizar_texto(
                endereco
            )
        )

        if (
            endereco_normalizado
            not in enderecos
        ):

            enderecos.append(
                endereco_normalizado
            )

            enderecos_originais[
                endereco_normalizado
            ] = texto(
                endereco
            )

    if len(enderecos) <= 1:

        return None

    return {
        "tipo": "Endereço divergente",
        "descricao": (
            "Foram identificados "
            "endereços diferentes "
            "para a mesma placa."
        ),
        "enderecos": [
            enderecos_originais[
                endereco
            ]
            for endereco in enderecos
        ],
        "quantidade": len(enderecos)
    }


# ============================================================
# NORMALIZAR REGISTROS
# ============================================================

def normalizar_registros(
    dados
):

    if dados is None:

        return []

    if isinstance(
        dados,
        dict
    ):

        registros = dados.get(
            "registros",
            []
        )

        if isinstance(
            registros,
            list
        ):

            return registros

        return []

    if isinstance(
        dados,
        list
    ):

        return dados

    if isinstance(
        dados,
        pd.DataFrame
    ):

        return dados.to_dict(
            orient="records"
        )

    return []


# ============================================================
# PROCESSAR DESVIOS
# ============================================================

def processar_desvios(
    dados
):

    registros = normalizar_registros(
        dados
    )

    desvios = []

    # --------------------------------------------------------
    # MOTORISTA
    # --------------------------------------------------------

    motorista = (
        identificar_motoristas_diferentes(
            registros
        )
    )

    if motorista:

        desvios.append(
            motorista
        )

    # --------------------------------------------------------
    # ENDEREÇO
    # --------------------------------------------------------

    endereco = (
        identificar_enderecos_diferentes(
            registros
        )
    )

    if endereco:

        desvios.append(
            endereco
        )

    return desvios


# ============================================================
# PROCESSAR DATAFRAME
# ============================================================

def processar_dataframe(
    df
):

    if df is None or df.empty:

        return []

    registros = df.to_dict(
        orient="records"
    )

    return processar_desvios(
        registros
    )


# ============================================================
# RESUMO
# ============================================================

def gerar_resumo(
    placa,
    desvios
):

    quantidade = len(
        desvios
    )

    if quantidade == 0:

        return {
            "PLACA": placa,
            "status": "SEM DESVIO",
            "quantidade_desvios": 0
        }

    tipos = [
        desvio.get(
            "tipo",
            ""
        )
        for desvio in desvios
    ]

    return {
        "PLACA": placa,
        "status": "COM DESVIO",
        "quantidade_desvios": quantidade,
        "tipos": tipos,
        "desvios": desvios
    }
