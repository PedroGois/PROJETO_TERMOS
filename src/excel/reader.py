import openpyxl

def carregar_planilha(caminho):
    wb = openpyxl.load_workbook(caminho, data_only=True)
    ws = wb.active

    registros = []

    for linha in ws.iter_rows(min_row=2):
        pessoa = linha[0].value
        ativo = linha[1].value
        tipo_mov = linha[3].value

        # pego o link do hiperlink
        link = linha[1].hyperlink.target if linha[1].hyperlink else None

        if tipo_mov == "Entrega em andamento" and pessoa != "DIGITAL BRENDA":
            registros.append({
            "pessoa": pessoa,
            "ativo": ativo,
            "tipo_mov": tipo_mov,
            "link": link
        })

    return registros