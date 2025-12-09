from excel.reader import carregar_planilha
from browser.automator import Automator
import time

def main():
    caminho_excel = r"C:\Users\Estagiario\JFI Silvicultura Ltda\Suporte JFI - Ti\Pedro\CONTROLE\TERMO\COLETA DE TERMO.xlsx"

    registros = carregar_planilha(caminho_excel)

    auto = Automator()

    for reg in registros:
        print(f"\nProcessando: {reg['pessoa']}")
        print(f"Link: {reg['link']}\n")

        nome = reg['pessoa']
        auto.abrir_link(reg['link'])
        
        # Verifica se o termo já foi confirmado (botão não encontrado)
        if not auto.enviar_termo():
            print(f"✓ {nome} já confirmou o termo")
            auto.registrar_confirmado(nome)
        else:
            # Se tem botão, significa que precisa confirmar - envia mensagem no Teams
            print(f"✗ {nome} precisa confirmar - enviando mensagem Teams")
            auto.enviar_mensagem(nome)
            auto.registrar_pendente(nome)  # Registra como pendente (enviou mensagem)

    auto.fechar()

if __name__ == "__main__":
    main()
