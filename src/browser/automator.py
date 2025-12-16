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
        # aguarda um curto período para que o chrome debug inicie (se necessário)
        time.sleep(1)

        chrome_options = Options()
        # Conecta à porta de debug aberta pelo .bat
        chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")

        self.driver = webdriver.Chrome(options=chrome_options)
        
        self.vcx_handle = self.driver.current_window_handle
        self.teams_handle = None
        print("Conectado com sucesso!")

    def alternar_aba(self, aba):
        """Gerencia a troca entre VCX e Teams sem recarregar páginas."""
        if aba == "TEAMS":
            if self.teams_handle is None:
                print("[SISTEMA] Abrindo Teams pela primeira vez (novo tab via Selenium)...")
                # tenta abrir nova aba via Selenium 4 API
                try:
                    self.driver.switch_to.new_window('tab')
                    self.driver.get('https://teams.microsoft.com/')
                    try:
                        WebDriverWait(self.driver, 20).until(
                            EC.presence_of_element_located((By.ID, 'ms-searchux-input'))
                        )
                    except Exception:
                        pass
                    self.teams_handle = self.driver.current_window_handle
                    print(f"[SISTEMA] Nova aba criada: {self.teams_handle}")
                except Exception as e:
                    print(f"[SISTEMA] Falha ao criar nova aba com Selenium: {e} - tentando fallback JS...")
                    try:
                        self.driver.execute_script("window.open('https://teams.microsoft.com/v2/', '_blank');")
                        try:
                            WebDriverWait(self.driver, 20).until(
                                lambda d: len(d.window_handles) > 1
                            )
                        except Exception:
                            pass
                        self.teams_handle = self.driver.window_handles[-1]
                        print(f"[SISTEMA] Fallback criou aba: {self.teams_handle}")
                    except Exception as ex:
                        print(f"[SISTEMA] Fallback também falhou: {ex}")
                        raise

            print(f"[SISTEMA] Alternando para handle do Teams: {self.teams_handle}")
            self.driver.switch_to.window(self.teams_handle)

    def abrir_link(self, url):
        """Abre o `url` na aba VCX (garante alternância para a aba correta)."""
        try:
            # garante que temos um handle do VCX
            if not hasattr(self, 'vcx_handle') or self.vcx_handle is None:
                self.vcx_handle = self.driver.current_window_handle

            self.driver.switch_to.window(self.vcx_handle)
            self.driver.get(url)
            try:
                WebDriverWait(self.driver, 10).until(
                    lambda d: d.execute_script("return document.readyState") == "complete"
                )
            except Exception:
                pass
        except Exception as e:
            print(f"[SISTEMA] Erro ao abrir link {url}: {e}")

    def enviar_termo(self):
        # Fase 0: Validação de login (verifica se existe botão sign-in-button)
        try:
            print("[TERMO] Fase 0: Verificando se está na tela de login...")
            wait = WebDriverWait(self.driver, 5)
            login_button = wait.until(EC.element_to_be_clickable((By.ID, "sign-in-button")))
            print("[TERMO] Botão de login encontrado! Clicando...")
            login_button.click()
            try:
                WebDriverWait(self.driver, 10).until(EC.staleness_of(login_button))
            except Exception:
                pass
        except TimeoutException:
            print("[TERMO] Nenhum botão de login encontrado - você já está logado!")
        except Exception as e:
            print(f"[TERMO] Erro na validação de login: {e}")

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

            try:
                self.driver.execute_script("arguments[0].click();", botao)
            except ElementClickInterceptedException:
                try:
                    ActionChains(self.driver).move_to_element(botao).click().perform()
                except Exception as e:
                    print(f"[TERMO] Erro ao clicar: {e}")
                    return False

            try:
                WebDriverWait(self.driver, 5).until(EC.staleness_of(botao))
            except Exception:
                pass

            print("[TERMO] Botão clicado com sucesso!")
            return True

        except Exception as e:
            print(f"[TERMO] Erro inesperado: {e}")
            return False

    # O método fechar foi removido por solicitação do usuário.


    def enviar_mensagem(self, nome):
        # Mensagem a enviar no Teams
        linhas_mensagem = [
            "Olá! Estou passando aqui para te fazer um lembrete importante sobre o seu equipamento. ⚠️",
            "",
            "Ainda falta confirmar o seu *Termo de Responsabilidade*. Pedimos que finalize este processo para *evitar o bloqueio do seu e-mail*.",
            "",
            "Siga estes 3 passos simples:",
            "1 - Procure o e-mail com o assunto: '[VC-X Sonar] Entrega em andamento de ativo'",
            "2 - Clique em 'Ver Proposta de Movimentação de Ativo'",
            "3 - Vá em 'Confirmar movimentação' e pronto! ✅",
            "",
            "Qualquer dúvida, é só nos chamar! Agradecemos a atenção."
        ]
        
        try:
            print(f"[TEAMS] Fase 1: Alternando para aba do Teams para {nome}...")
            self.alternar_aba("TEAMS")
            wait = WebDriverWait(self.driver, 15)
        except Exception as e:
            print(f"[TEAMS] Erro ao alternar para aba do Teams: {e}")
            return False
            
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
            try:
                WebDriverWait(self.driver, 8).until(
                    EC.presence_of_element_located((By.XPATH, "//div[@data-tid='search-people-card']"))
                )
            except Exception:
                pass

            # Fase 4: Seleciona o resultado
            print(f"[TEAMS] Fase 4: Selecionando resultado para {nome}...")
            try:
                card_xpath = "//div[@data-tid='search-people-card' and .//span[contains(normalize-space(.), '{}')]]".format(nome)
                button_xpath = card_xpath + "//button[@data-tid='carousel-card-button']"
                card_button = wait.until(EC.element_to_be_clickable((By.XPATH, button_xpath)))
                print("[TEAMS] Resultado localizado por XPath (nome exato).")
                try:
                    self.driver.execute_script("arguments[0].scrollIntoView({block:'center'});", card_button)
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

            # Limpa qualquer rascunho que possa estar na caixa de texto
            message_box.send_keys(Keys.CONTROL + "a")  # Seleciona tudo
            message_box.send_keys(Keys.DELETE)  # Deleta tudo

            # Pequeno delay controlado enquanto digita cada linha (evita problemas de envio acelerado)
            for i, linha in enumerate(linhas_mensagem):
                message_box.send_keys(linha)
                if i < len(linhas_mensagem) - 1:
                    message_box.send_keys(Keys.SHIFT + Keys.ENTER)
                time.sleep(0.05)

            message_box.send_keys(Keys.ENTER)
            print(f"[TEAMS] Mensagem enviada com sucesso para {nome}!")
            self.alternar_aba("VCX")
            return True

        except TimeoutException as e:
            print(f"[TEAMS] Erro: timeout ao processar {nome}.")
            return False
        except Exception as e:
            print(f"[TEAMS] Erro inesperado ao processar {nome}: {e}")
            return False

# Módulo pronto para importação e uso por outros scripts. Exemplo de execução
# foi removido; o próximo passo será implementar gravação em planilha com openpyxl.

