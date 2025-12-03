from excel.reader import carregar_planilha

CAMINHO = r"C:\Users\Estagiario\JFI Silvicultura Ltda\Suporte JFI - Ti\Pedro\CONTROLE\TERMO\COLETA TERMO - 03-12-25.xlsx"

def main():
    registros = carregar_planilha(CAMINHO)

    print("\n=== TOTAL DE REGISTROS LIDOS:", len(registros), "===\n")

    # Mostra os 5 primeiros registros para conferência
    for r in registros[:5]:
        print(f"Pessoa: {r['pessoa']}")
        print(f"Ativo: {r['ativo']}")
        print(f"Movimentação: {r['tipo_mov']}")
        print(f"Link: {r['link']}")
        print("----")

    # Conferência extra
    sem_link = [r for r in registros if r["link"] is None]
    print("\n QTD Registros sem link:", len(sem_link))

if __name__ == "__main__":
    main()