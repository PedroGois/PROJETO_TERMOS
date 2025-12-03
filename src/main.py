from excel.reader import carregar_planilha
from browser.automator import Automator
from teams.sender import enviar_mensagem_teams

def main():
    caminho = "data/termos.xlsx"
    registros = carregar_planilha(caminho)

    automator = Automator()

    faltando_botao = []
    faltando_teams = []

    for r in registros:
        nome = r["Nome"]
        link = r["Link"]

        print(f"Processando {nome}...")

        # 1. Selenium
        if not automator.reenviar_termo(link):
            faltando_botao.append(nome)

        # 2. Teams
        if not enviar_mensagem_teams(nome):
            faltando_teams.append(nome)

    automator.fechar()

    print("\n=== PROCESSO FINALIZADO ===")
    print("Sem botão de reenviar:")
    print(faltando_botao)

    print("\nSem Teams:")
    print(faltando_teams)

if __name__ == "__main__":
    main()
