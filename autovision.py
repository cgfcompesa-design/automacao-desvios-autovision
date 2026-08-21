
import os
import time

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from selenium.common.exceptions import (
    TimeoutException,
    NoSuchElementException,
    WebDriverException,
)


class Autovision:

    def __init__(self):

        self.driver = None

        self.url = os.getenv(
            "AUTOVISION_URL",
            ""
        )

        self.usuario = os.getenv(
            "AUTOVISION_USUARIO",
            ""
        )

        self.senha = os.getenv(
            "AUTOVISION_SENHA",
            ""
        )

        self.timeout = 30

    # ========================================================
    # INICIAR NAVEGADOR
    # ========================================================

    def iniciar(self):

        print(
            "[AUTOVISION] Iniciando navegador..."
        )

        options = Options()

        options.add_argument(
            "--headless=new"
        )

        options.add_argument(
            "--disable-gpu"
        )

        options.add_argument(
            "--no-sandbox"
        )

        options.add_argument(
            "--disable-dev-shm-usage"
        )

        options.add_argument(
            "--window-size=1920,1080"
        )

        options.add_argument(
            "--disable-blink-features=AutomationControlled"
        )

        try:

            self.driver = webdriver.Edge(
                options=options
            )

        except WebDriverException as exc:

            raise RuntimeError(
                "Não foi possível iniciar o Microsoft Edge/WebDriver: "
                + str(exc)
            )

        self.driver.set_page_load_timeout(
            self.timeout
        )

        if self.url:

            self.driver.get(
                self.url
            )

            print(
                "[AUTOVISION] Página inicial aberta."
            )

            self._login()

        else:

            print(
                "[AUTOVISION] AUTOVISION_URL não configurada."
            )

    # ========================================================
    # LOGIN
    # ========================================================

    def _login(self):

        """
        Coloque aqui o login do seu código original
        do Autovision.

        Esta função é chamada automaticamente por iniciar().
        """

        if not self.usuario or not self.senha:

            print(
                "[AUTOVISION] Usuário/senha não configurados."
            )

            return

        print(
            "[AUTOVISION] Executando login..."
        )

        # ====================================================
        # IMPORTANTE
        # ====================================================
        #
        # Os seletores abaixo são exemplos.
        #
        # Substitua pelos seletores do seu código original.
        #
        # ====================================================

        try:

            # ------------------------------------------------
            # EXEMPLO:
            # ------------------------------------------------
            #
            # campo_usuario = WebDriverWait(
            #     self.driver,
            #     self.timeout
            # ).until(
            #     EC.presence_of_element_located(
            #         (By.NAME, "usuario")
            #     )
            # )
            #
            # campo_usuario.send_keys(
            #     self.usuario
            # )
            #
            # campo_senha = self.driver.find_element(
            #     By.NAME,
            #     "senha"
            # )
            #
            # campo_senha.send_keys(
            #     self.senha
            # )
            #
            # botao = self.driver.find_element(
            #     By.XPATH,
            #     "//button[contains(., 'Entrar')]"
            # )
            #
            # botao.click()
            #
            # ------------------------------------------------

            print(
                "[AUTOVISION] "
                "Login: inserir código original aqui."
            )

        except TimeoutException:

            raise RuntimeError(
                "Timeout durante login do Autovision."
            )

    # ========================================================
    # CONSULTAR PLACA
    # ========================================================

    def consultar_placa(
        self,
        placa,
        data=None
    ):

        if self.driver is None:

            raise RuntimeError(
                "O Autovision não foi iniciado."
            )

        print(
            f"[AUTOVISION] Consultando placa {placa}"
        )

        # ====================================================
        # AQUI ENTRA O CÓDIGO ORIGINAL DO AUTOVISION
        # ====================================================

        #
        # O fluxo deverá:
        #
        # 1. acessar o relatório necessário
        #
        # 2. preencher a placa
        #
        # 3. preencher a data
        #
        # 4. clicar em pesquisar
        #
        # 5. aguardar o relatório
        #
        # 6. extrair os dados
        #
        # 7. retornar um dicionário
        #
        # ====================================================

        resultado = {
            "PLACA": placa,
            "DATA": data,
            "registros": []
        }

        return resultado

    # ========================================================
    # EXTRAIR TABELA
    # ========================================================

    def extrair_tabela(self):

        if self.driver is None:

            return []

        registros = []

        try:

            tabelas = self.driver.find_elements(
                By.TAG_NAME,
                "table"
            )

            for tabela in tabelas:

                linhas = tabela.find_elements(
                    By.TAG_NAME,
                    "tr"
                )

                for linha in linhas:

                    celulas = linha.find_elements(
                        By.TAG_NAME,
                        "td"
                    )

                    if not celulas:
                        continue

                    valores = [
                        celula.text.strip()
                        for celula in celulas
                    ]

                    registros.append(
                        valores
                    )

        except Exception as exc:

            print(
                "[AUTOVISION] Erro ao extrair tabela:",
                exc
            )

        return registros

    # ========================================================
    # FECHAR
    # ========================================================

    def fechar(self):

        if self.driver is not None:

            try:

                self.driver.quit()

            except Exception:

                pass

            finally:

                self.driver = None

        print(
            "[AUTOVISION] Navegador encerrado."
        )
