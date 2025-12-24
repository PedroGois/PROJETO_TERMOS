from excel.reader import carregar_planilha
import excel.writer as writer
from browser.automator import Automator
import os
import subprocess
import sys
import time

# Fecha Excel antes de rodar
subprocess.call([sys.executable, r"C:\Users\Estagiario\JFI Silvicultura Ltda\Suporte JFI - Ti\Pedro\PROJETOS\PROJETO_TERMOS\src\excel\fechar_excel.py"])

def main():
    caminho_dir = r"C:\Users\Estagiario\JFI Silvicultura Ltda\Suporte JFI - Ti\Pedro\CONTROLE\TERMO"
    caminho_coleta = os.path.join(caminho_dir, "COLETA DE TERMO.xlsx")
    caminho_status = os.path.join(caminho_dir, "STATUS TERMOS.xlsx")

    registros = carregar_planilha(caminho_coleta)
    auto = Automator()

    for i, reg in enumerate(registros, 1):
        nome = reg['pessoa']
        link = reg['link']
        
        print(f"\n[{i}/{len(registros)}] Processando: {nome}")

        status_email = "ERRO"
        status_msg = "ERRO"

        try:
            auto.abrir_link(link)

            # 1. Verifica Termo no VCX
            if not auto.enviar_termo():
                print(f"✓ {nome} já confirmou.")
                status_email = "CONFIRMADO"
                status_msg = "CONFIRMADO"
            else:
                print(f"✗ {nome} pendente - Acionando Teams...")
                status_email = "ENVIADO"
                # 2. Envia Mensagem no Teams
                enviado = auto.enviar_mensagem(nome)
                status_msg = "ENVIADO" if enviado else "ERRO"
                
                if enviado:
                    print(f"Mensagem enviada no Teams para {nome}.")
                else:
                    print(f"Falha ao enviar mensagem no Teams para {nome}.")

        except Exception as e:
            print(f"[ERRO] Falha geral em {nome}: {e}")
            status_email = f"Erro: {str(e)[:20]}"
        
        # 3. Alimenta a planilha de status via módulo Writer
        writer.registrar_status(caminho_status, nome, link, status_email, status_msg)

if __name__ == "__main__":
    main()