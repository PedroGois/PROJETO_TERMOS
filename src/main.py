from excel.reader import carregar_planilha
from browser.automator import Automator


def main():
    # Inicia o processamento
    print("=== INICIANDO PROCESSAMENTO ===\n")
    caminho_excel = r"C:\Users\Estagiario\JFI Silvicultura Ltda\Suporte JFI - Ti\Pedro\CONTROLE\TERMO\COLETA TERMO.xlsx"

    # Carrega registros da planilha
    registros = carregar_planilha(caminho_excel)

    if not registros:
        print("Nenhum registro para processar. Encerrando.")
        return

    # Inicia a automação
    auto = Automator()

    # Processa cada registro
    total = len(registros)
    for idx, reg in enumerate(registros, start=1):
        print(f"\n[PROCESSO] {idx}/{total} - Processando {reg['pessoa']}...")
        print(f"[PROCESSO] Link: {reg['link']}\n")

        auto.abrir_link(reg["link"])

        # Se enviar_termo retornar False, pula para próximo registro
        if not auto.enviar_termo():
            print(f"[PROCESSO] Pulando {reg['pessoa']} — botão não encontrado.\n")
            continue

        nome = reg["pessoa"]
        auto.enviar_mensagem(nome)
        print(f"[PROCESSO] Concluído para {nome}.\n")

    print("=== FIM DO PROCESSAMENTO ===")
    auto.fechar()


if __name__ == "__main__":
    main()
