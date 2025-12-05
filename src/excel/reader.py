import openpyxl

def carregar_planilha(caminho):
    # Abre a planilha com valores apenas (sem fórmulas)
    print("[READER] Abrindo planilha...")
    wb = openpyxl.load_workbook(caminho, data_only=True)
    ws = wb.active

    registros = []

    # Itera linhas a partir da linha 2 (pulando cabeçalho)
    for linha in ws.iter_rows(min_row=2):
        pessoa = linha[0].value
        ativo = linha[1].value
        tipo_mov = linha[3].value

        # Extrai o link do hiperlink, se houver
        link = linha[1].hyperlink.target if linha[1].hyperlink else None

        # Filtra registros por tipo de movimentação e pessoa
        if tipo_mov == "Entrega em andamento" and pessoa != "DIGITAL BRENDA":
            registros.append({
                "pessoa": pessoa,
                "ativo": ativo,
                "tipo_mov": tipo_mov,
                "link": link
            })

    print(f"[READER] Total de registros encontrados: {len(registros)}")
    return registros