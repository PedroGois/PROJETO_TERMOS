from excel.reader import carregar_planilha
from browser.automator import Automator


def main():
    # Caminho do arquivo Excel
    caminho_excel = r"C:\Users\Estagiario\JFI Silvicultura Ltda\Suporte JFI - Ti\Pedro\CONTROLE\TERMO\COLETA TERMO - 03-12-25.xlsx"

    # Carrega registros da planilha
    registros = carregar_planilha(caminho_excel)

    # Inicia a automação
    auto = Automator()

    # Processa cada registro
    for reg in registros[11:12]:
        print("Processando:", reg["pessoa"], reg["link"])
        auto.abrir_link(reg["link"])
        auto.enviar_termo()
        
        nome = reg["pessoa"]
        auto.enviar_mensagem(nome)


if __name__ == "__main__":
    main()
