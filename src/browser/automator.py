from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
import time

class Automator:

    def __init__(self):
        chrome_options = Options()
        # Conecta na porta que abrimos pelo .bat
        chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")

        # CORREÇÃO: Usar 'self.driver' para que fique acessível na classe toda
        self.driver = webdriver.Chrome(options=chrome_options)

        print("Conectado com sucesso!")
        print("Titulo da pagina:", self.driver.title)
        
    def abrir_link(self, url):
        self.driver.get(url)
        time.sleep(1)

    def enviar_termo(self):
        try:
            # Exemplo de clique num botão
            print("Procurando botão...")
            botao = self.driver.find_element(By.XPATH, "//button[contains(text(), 'Enviar termo')]")
            botao.click()
            print("Botão clicado!")
            time.sleep(1)
        except Exception as e:
            print(f"Erro ao tentar clicar: {e}")

    def fechar(self):
        # Atenção: Isso fecha a conexão. Se quiser fechar o navegador todo, use quit()
        # Se quiser apenas desconectar o script e deixar o Chrome aberto, não chame nada ou use close() na aba
        self.driver.quit()

# --- COMO USAR ---
if __name__ == "__main__":
    # 1. Certifique-se que rodou o arquivo .bat antes!
    
    # Instancia a automação
    bot = Automator()
    
    # Exemplo de uso
    # bot.abrir_link("https://seu-site.com")
    # bot.enviar_termo()