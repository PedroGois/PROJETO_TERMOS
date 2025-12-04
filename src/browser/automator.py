from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import time
import subprocess
import os


class Automator:

    def __init__(self, bat_path="abrir_chrome_debug.bat"):
        # Executa o .bat para iniciar Chrome em modo debug
        bat_completo = os.path.join(os.path.dirname(__file__), bat_path)
        subprocess.Popen(bat_completo)
        time.sleep(3)

        chrome_options = Options()
        # Conecta à porta de debug aberta pelo .bat
        chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")

        self.driver = webdriver.Chrome(options=chrome_options)
        print("Conectado com sucesso!")
        print("Titulo da pagina:", self.driver.title)
        time.sleep(2)

    def abrir_link(self, url):
        # Navega até a URL
        self.driver.get(url)
        time.sleep(1)

    def enviar_termo(self):
        # Localiza e clica no botão de reenvio
        try:
            print("Procurando botão...")
            wait = WebDriverWait(self.driver, 15)

            # Tenta localizar pelo ID primeiro, depois pelo XPath
            botao = None
            try:
                botao = wait.until(EC.element_to_be_clickable((By.ID, "resend-asset-control-email-confirmation-button")))
            except (TimeoutException, NoSuchElementException):
                botao = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Reenviar E-mail de Movimentação')]")))

            # Rola a página até o botão
            print("Rolando a página até o botão...")
            self.driver.execute_script("arguments[0].scrollIntoView(true);", botao)
            time.sleep(0.5)

            # Clica via JavaScript para garantir sucesso
            print("Clicando no botão...")
            self.driver.execute_script("arguments[0].click();", botao)
            # botao.click() #também funciona
            print("Botão clicado com sucesso!")
            time.sleep(1)

        except Exception as e:
            print(f"Erro ao clicar no botão: {e}")

    def fechar(self):
        # Fecha a conexão com o browser
        self.driver.quit()


# Exemplo de uso
if __name__ == "__main__":
    bot = Automator()

