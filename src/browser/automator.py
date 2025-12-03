from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException

class Automator:
    def __init__(self):
        self.driver = webdriver.Chrome()

    def reenviar_termo(self, link):
        try:
            self.driver.get(link)

            try:
                botao = self.driver.find_element(By.XPATH, '//button[@title="Reenviar E-mail de Movimentação"]')
                botao.click()
                return True
            except NoSuchElementException:
                return False

        except Exception as e:
            print(f"Erro no Selenium: {e}")
            return False

    def fechar(self):
        self.driver.quit()