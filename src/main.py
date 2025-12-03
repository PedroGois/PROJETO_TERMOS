from excel.reader import carregar_planilha
from browser.automator import Automator
import time

def main():
    caminho_excel = r"C:\Users\Estagiario\JFI Silvicultura Ltda\Suporte JFI - Ti\Pedro\CONTROLE\TERMO\COLETA TERMO - 03-12-25.xlsx"

    registros = carregar_planilha(caminho_excel)

    auto = Automator()

    for reg in registros:
        print("Processando:", reg["pessoa"], reg["link"])

        auto.abrir_link(reg["link"])
        time.sleep(2)
        auto.enviar_termo()

    auto.fechar()

if __name__ == "__main__":
    main()
