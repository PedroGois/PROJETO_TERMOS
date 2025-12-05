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

    def enviar_mensagem(self, nome):
        # Lista com linhas de mensagem (vai usar Shift+Enter para quebras)
        linhas_mensagem = [
            "Olá! Estamos entrando em contato para te lembrar de confirmar o seu termo de responsabilidade referente a seu equipamento.",
            "",  # Linha vazia = parágrafo
            "Por favor, para evitar o bloqueio de seu email, verifique seu e-mail e siga as instruções para completar o processo.",
            "",  # Linha vazia = parágrafo
            "Orientações:",
            "1 - Acessar o e-mail que chegou para você com o assunto '[VC-X Sonar] Entrega em andamento de ativo' >",
            "2 - Clicar em 'Ver Proposta de Movimentação de Ativo' >",
            "3- Ir em 'Confirmar movimentação' e pronto.",
            "",  # Linha vazia = parágrafo
            "Agradecemos sua atenção!"]

        # Abre o Teams e captura erro de navegação
        try:
            self.abrir_link("https://teams.microsoft.com/v2/")
        except Exception as e:
            print(f"[TEAMS] Erro ao abrir Teams: {e}")
            return False

        wait = WebDriverWait(self.driver, 5)
        print(f"[TEAMS] Iniciando busca por {nome}...")

        # Localiza a barra de pesquisa (tenta ID, depois XPath)
        try:
            search_input = wait.until(EC.element_to_be_clickable((By.ID, "ms-searchux-input")))
            print("[TEAMS] Barra de pesquisa encontrada por ID.")
        except TimeoutException:
            try:
                print("[TEAMS] Tentando localizar a barra de pesquisa por XPath...")
                search_input = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@aria-label='Search box']")))
                print("[TEAMS] Barra de pesquisa encontrada por XPath.")
            except TimeoutException as e:
                print(f"[TEAMS] ERRO: barra de pesquisa não encontrada: {e}")
                return False

        # Pesquisa o usuário e seleciona o resultado
        try:
            search_input.clear()
            search_input.send_keys(nome)
            search_input.send_keys(Keys.ENTER)
            time.sleep(1)  # deixa o Teams atualizar resultados

            # Tenta localizar o cartão de pessoa que contenha o nome pesquisado
            try:
                card_xpath = "//div[@data-tid='search-people-card' and .//span[contains(normalize-space(.), '{}')]]".format(nome)
                button_xpath = card_xpath + "//button[@data-tid='carousel-card-button']"
                card_button = wait.until(EC.element_to_be_clickable((By.XPATH, button_xpath)))
                print("testando encontro do card da pessoa...")
                # Rola o cartão para o centro e tenta clicar via JS (evita interceptação)
                try:
                    self.driver.execute_script("arguments[0].scrollIntoView({block:'center'});", card_button)
                    time.sleep(0.2)
                    self.driver.execute_script("arguments[0].click();", card_button)
                    print(f"[TEAMS] Resultado selecionado por cartão para {nome}.")
                except ElementClickInterceptedException:
                    # Se elemento interceptado, tenta ActionChains
                    try:
                        ActionChains(self.driver).move_to_element(card_button).click().perform()
                        print(f"[TEAMS] Resultado selecionado por cartão (ActionChains) para {nome}.")
                    except Exception as e:
                        print(f"[TEAMS] Falha ao clicar no cartão: {e}")
                        raise

            except TimeoutException:
                # Fallback: tenta clicar no primeiro cartão disponível
                try:
                    first_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[@data-tid='search-people-card']//button[@data-tid='carousel-card-button']")))
                    try:
                        self.driver.execute_script("arguments[0].scrollIntoView({block:'center'});", first_button)
                        time.sleep(0.2)
                        self.driver.execute_script("arguments[0].click();", first_button)
                        print(f"[TEAMS] Selecionou primeiro resultado disponível para {nome}.")
                    except Exception:
                        first_button.click()
                        print(f"[TEAMS] Selecionou primeiro resultado disponível para {nome} (click()).")
                except Exception as e:
                    print(f"[TEAMS] Não foi possível selecionar resultado de pesquisa: {e}")
                    return False

            # Aguarda o campo de mensagem e envia
            message_box = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[@contenteditable='true']")))
            message_box.click()

            # Envia cada linha com Shift+Enter para quebra (exceto a última, que usa Enter)
            for i, linha in enumerate(linhas_mensagem):
                message_box.send_keys(linha)
                if i < len(linhas_mensagem) - 1:
                    # Shift+Enter para quebra de linha
                    message_box.send_keys(Keys.SHIFT + Keys.ENTER)
            
            # Enter final para enviar
            message_box.send_keys(Keys.ENTER)

            print(f"[TEAMS] Mensagem enviada com sucesso para {nome}.")
            return True

        except TimeoutException as e:
            print(f"[TEAMS] ERRO DE TIMEOUT ao enviar mensagem: {e}")
            return False
        except Exception as e:
            print(f"[TEAMS] ERRO INESPERADO ao enviar mensagem: {e}")
            return False

# Exemplo de uso
if __name__ == "__main__":
    bot = Automator()

