from excel.reader import carregar_planilha
from browser.automator import Automator
import time

def main():
    caminho_excel = r"C:\Users\Estagiario\JFI Silvicultura Ltda\Suporte JFI - Ti\Pedro\CONTROLE\TERMO\COLETA DE TERMO.xlsx"

    registros = carregar_planilha(caminho_excel)

    auto = Automator()

    for i, reg in enumerate(registros, 1):
        print(f"\nRegistro {i} de {len(registros)}")
        print(f"\nProcessando: {reg['pessoa']}")
        print(f"Link: {reg['link']}\n")

        nome = reg['pessoa']
        
        try:
            auto.abrir_link(reg['link'])
            
            # Verifica se o termo já foi confirmado (botão não encontrado)
            if not auto.enviar_termo(nome):
                print(f"✓ {nome} já confirmou o termo")
                auto.registrar_confirmado(nome)
            else:
                # Se tem botão, significa que precisa confirmar - envia mensagem no Teams
                print(f"✗ {nome} precisa confirmar - enviando mensagem Teams")
                try:
                    auto.enviar_mensagem(nome)
                    auto.registrar_pendente(nome)  # Registra como pendente (enviou mensagem)
                except Exception as e:
                    print(f"[ERRO] Falha ao enviar mensagem para {nome}: {e}")
                    auto.registrar_erro(nome, f"Erro ao enviar mensagem Teams: {e}")
        
        except Exception as e:
            print(f"[ERRO] Falha ao processar {nome}: {e}")
            auto.registrar_erro(nome, f"Erro geral ao processar: {e}")

    auto.fechar()

if __name__ == "__main__":
    main()
