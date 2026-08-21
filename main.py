
import requests
import sys


# ============================================================
# CONFIGURAÇÃO
# ============================================================

APPS_SCRIPT_URL = (
    "https://script.google.com/macros/s/"
    "AKfycbwiKp58t3hyHHXXnjqLBbLcrY1o8pupLMOSLPSxOqpL"
    "/exec"
)


# ============================================================
# BUSCAR PLACAS
# ============================================================

def buscar_placas():

    url = APPS_SCRIPT_URL

    params = {
        "acao": "placas"
    }

    print("Consultando Apps Script...")
    print(url)

    resposta = requests.get(
        url,
        params=params,
        timeout=60
    )

    print("HTTP:", resposta.status_code)

    resposta.raise_for_status()

    dados = resposta.json()

    if not dados.get("success"):

        raise RuntimeError(
            "Apps Script retornou erro: "
            + str(dados)
        )

    placas = dados.get("placas", [])

    print()
    print("=" * 60)
    print("PLACAS RECEBIDAS")
    print("=" * 60)

    for item in placas:

        print(
            f"Linha: {item.get('linha')} | "
            f"Placa: {item.get('placa')}"
        )

    print()
    print("Total:", len(placas))

    return placas


# ============================================================
# MAIN
# ============================================================

def main():

    try:

        placas = buscar_placas()

        if not placas:

            print("Nenhuma placa encontrada.")

        else:

            print(
                f"{len(placas)} placa(s) serão "
                "disponibilizadas para a análise."
            )

    except Exception as erro:

        print()
        print("ERRO:")
        print(erro)

        sys.exit(1)


if __name__ == "__main__":
    main()
