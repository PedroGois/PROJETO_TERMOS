from excel.reader import carregar_planilha
from browser.automator import Automator


def main():
    caminho_excel = r"C:\Users\Estagiario\JFI Silvicultura Ltda\Suporte JFI - Ti\Pedro\CONTROLE\TERMO\COLETA DE TERMO.xlsx"
    registros = carregar_planilha(caminho_excel)
    auto = Automator()

    for i, reg in enumerate(registros, 1):
        nome = reg['pessoa']
        link = reg['link']
        
        print(f"\n[{i}/{len(registros)}] Processando: {nome}")

        try:
            # 1. Abre o link no VCX (Aba 1)
            auto.abrir_link(link)

            # 2. Tenta o reenvio no VCX (passa o nome para contexto de log)
            if not auto.enviar_termo():
                print(f"✓ {nome} já confirmou.")
            else:
                # 3. Se precisou reenviar, vai para o Teams (Aba 2) e envia a mensagem
                print(f"✗ {nome} pendente - Acionando Teams...")
                enviado = auto.enviar_mensagem(nome)
                if enviado:
                    print(f"Mensagem enviada no Teams para {nome}.")
                else:
                    print(f"Falha ao enviar mensagem no Teams para {nome}.")
        except Exception as e:
            print(f"[ERRO] Falha geral em {nome}: {e}")

    # encerramento do browser foi removido do Automator; feche manualmente se necessário
if __name__ == "__main__":
    main()