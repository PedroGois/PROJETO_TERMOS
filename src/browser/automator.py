from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, ElementClickInterceptedException
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
import time
import subprocess
import os
from datetime import datetime


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
        
        # Arquivos de log
        self.log_confirmados = "confirmados.log"  # Pessoas que já confirmaram o termo
        self.log_pendentes = "pendentes.log"  # Pessoas que receberam mensagem (precisa confirmar)
        self.log_erros = "erros.log"  # Pessoas com erro
        
        print("Conectado com sucesso!")
        print("Titulo da pagina:", self.driver.title)
        time.sleep(2)

    def registrar_confirmado(self, nome):
        # Registra pessoa que já confirmou (botão não encontrado)
        with open(self.log_confirmados, "a", encoding="utf-8") as f:
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            f.write(f"[{timestamp}] {nome}\n")

    def registrar_pendente(self, nome):
        # Registra pessoa que precisa confirmar (recebeu mensagem)
        with open(self.log_pendentes, "a", encoding="utf-8") as f:
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            f.write(f"[{timestamp}] {nome}\n")

    def registrar_erro(self, nome, erro):
        # Registra pessoa com erro
        with open(self.log_erros, "a", encoding="utf-8") as f:
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            f.write(f"[{timestamp}] {nome} - {erro}\n")

    def abrir_link(self, url):
        # Navega até a URL
        self.driver.get(url)
        time.sleep(1)

    def enviar_termo(self):
        # Fase 1: Localiza o botão de reenvio
        try:
            print("[TERMO] Fase 1: Procurando botão de reenvio...")
            wait = WebDriverWait(self.driver, 15)

            botao = None
            try:
                botao = wait.until(EC.element_to_be_clickable((By.ID, "resend-asset-control-email-confirmation-button")))
                print("[TERMO] Botão localizado por ID.")
            except (TimeoutException, NoSuchElementException):
                try:
                    botao = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Reenviar E-mail de Movimentação')]")))
                    print("[TERMO] Botão localizado por XPath.")
                except (TimeoutException, NoSuchElementException):
                    print("[TERMO] Botão não encontrado — pessoa já assinou o termo.")
                    return False

            # Fase 2: Rola e clica no botão
            print("[TERMO] Fase 2: Rolando e clicando no botão...")
            self.driver.execute_script("arguments[0].scrollIntoView(true);", botao)
            time.sleep(0.5)

            try:
                self.driver.execute_script("arguments[0].click();", botao)
            except ElementClickInterceptedException:
                try:
                    ActionChains(self.driver).move_to_element(botao).click().perform()
                except Exception as e:
                    print(f"[TERMO] Erro ao clicar: {e}")
                    return False

            print("[TERMO] Botão clicado com sucesso!")
            time.sleep(1)
            return True

        except Exception as e:
            print(f"[TERMO] Erro inesperado: {e}")
            return False

    def fechar(self):
        # Fecha a conexão com o browser
        self.driver.quit()

    def enviar_mensagem(self, nome):
        # Mensagem a enviar no Teams
        linhas_mensagem = [
            "Olá! Estou passando aqui para te fazer um lembrete importante sobre o seu equipamento. ⚠️",
            "",
            "Ainda falta confirmar o seu **Termo de Responsabilidade**. Pedimos que finalize este processo para **evitar o bloqueio do seu e-mail**.",
            "",
            "Siga estes 3 passos simples:",
            "1 - Procure o e-mail com o assunto: '[VC-X Sonar] Entrega em andamento de ativo'",
            "2 - Clique em 'Ver Proposta de Movimentação de Ativo'",
            "3 - Vá em 'Confirmar movimentação' e pronto! ✅",
            "",
            "Qualquer dúvida, é só nos chamar! Agradecemos a atenção."
        ]

        # Fase 1: Abre o Teams
        try:
            print(f"[TEAMS] Fase 1: Abrindo Teams para {nome}...")
            self.abrir_link("https://teams.microsoft.com/v2/")
        except Exception as e:
            print(f"[TEAMS] Erro ao abrir Teams: {e}")
            return False

        wait = WebDriverWait(self.driver, 5)

        # Fase 2: Localiza a barra de pesquisa
        print(f"[TEAMS] Fase 2: Procurando usuário {nome}...")
        try:
            search_input = wait.until(EC.element_to_be_clickable((By.ID, "ms-searchux-input")))
            print("[TEAMS] Barra de pesquisa localizada por ID.")
        except TimeoutException:
            try:
                search_input = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@aria-label='Search box']")))
                print("[TEAMS] Barra de pesquisa localizada por XPath.")
            except TimeoutException as e:
                print(f"[TEAMS] Erro: barra de pesquisa não encontrada.")
                return False

        # Fase 3: Pesquisa o usuário
        try:
            print(f"[TEAMS] Fase 3: Digitando nome {nome}...")
            search_input.clear()
            search_input.send_keys(nome)
            search_input.send_keys(Keys.ENTER)
            time.sleep(1)

            # Fase 4: Seleciona o resultado
            print(f"[TEAMS] Fase 4: Selecionando resultado para {nome}...")
            try:
                card_xpath = "//div[@data-tid='search-people-card' and .//span[contains(normalize-space(.), '{}')]]".format(nome)
                button_xpath = card_xpath + "//button[@data-tid='carousel-card-button']"
                card_button = wait.until(EC.element_to_be_clickable((By.XPATH, button_xpath)))
                print("[TEAMS] Resultado localizado por XPath (nome exato).")
                try:
                    self.driver.execute_script("arguments[0].scrollIntoView({block:'center'});", card_button)
                    time.sleep(0.2)
                    self.driver.execute_script("arguments[0].click();", card_button)
                    print("[TEAMS] Resultado clicado via JavaScript.")
                except ElementClickInterceptedException:
                    ActionChains(self.driver).move_to_element(card_button).click().perform()
                    print("[TEAMS] Resultado clicado via ActionChains.")

            except TimeoutException:
                print("[TEAMS] Resultado por nome não encontrado, tentando primeiro...")
                try:
                    first_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[@data-tid='search-people-card']//button[@data-tid='carousel-card-button']")))
                    print("[TEAMS] Primeiro resultado localizado por XPath.")
                    try:
                        self.driver.execute_script("arguments[0].scrollIntoView({block:'center'});", first_button)
                        time.sleep(0.2)
                        self.driver.execute_script("arguments[0].click();", first_button)
                        print("[TEAMS] Primeiro resultado clicado via JavaScript.")
                    except Exception:
                        first_button.click()
                        print("[TEAMS] Primeiro resultado clicado via click() nativo.")
                except Exception as e:
                    print(f"[TEAMS] Erro: resultado não encontrado.")
                    return False

            # Fase 5: Envia a mensagem
            print(f"[TEAMS] Fase 5: Enviando mensagem para {nome}...")
            message_box = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[@contenteditable='true']")))
            message_box.click()
            time.sleep(0.3)
            
            # Limpa qualquer rascunho que possa estar na caixa de texto
            message_box.send_keys(Keys.CONTROL + "a")  # Seleciona tudo
            message_box.send_keys(Keys.DELETE)  # Deleta tudo
            time.sleep(0.2)

            for i, linha in enumerate(linhas_mensagem):
                message_box.send_keys(linha)
                if i < len(linhas_mensagem) - 1:
                    message_box.send_keys(Keys.SHIFT + Keys.ENTER)
            
            message_box.send_keys(Keys.ENTER)
            print(f"[TEAMS] Mensagem enviada com sucesso para {nome}!")
            return True

        except TimeoutException as e:
            print(f"[TEAMS] Erro: timeout ao processar {nome}.")
            return False
        except Exception as e:
            print(f"[TEAMS] Erro inesperado ao processar {nome}: {e}")
            return False

# Exemplo de uso
if __name__ == "__main__":
    bot = Automator()

