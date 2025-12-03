from selenium import webdriver
from selenium.webdriver.common.by import By
import time

class Automator:

    def __init__(self):
        options = webdriver.ChromeOptions()
        options.add_argument("--start-maximized")
        
        # >>> opção para usar SEU Chrome já logado <<<
        options.add_argument(r'--user-data-dir=C:\Users\Estagiario\AppData\Local\Google\Chrome\User Data')
        options.add_argument(r'--profile-directory=Profile 2')  # normalmente "Default"
        self.driver = webdriver.Chrome(options=options)

    def abrir_link(self, url):
        self.driver.get(url)
        time.sleep(1)

    def enviar_termo(self):
        # exemplo de clique num botão
        botao = self.driver.find_element(By.XPATH, "//button[contains(text(), 'Enviar termo')]")
        botao.click()
        time.sleep(1)

    def fechar(self):
        self.driver.quit()