import openpyxl
import os

def registrar_status(caminho, nome, link, status_email, status_mensagem):
    """
    Abre a planilha de status e adiciona uma nova linha com os resultados.
    """
    # Verifica se o arquivo existe, se não, cria um novo com cabeçalho
    if not os.path.exists(caminho):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["NOME", "LINK", "STATUS EMAIL", "STATUS MENSAGEM"])
    else:
        wb = openpyxl.load_workbook(caminho)
        ws = wb.active

    # Adiciona os dados na próxima linha disponível
    ws.append([nome, link, status_email, status_mensagem])

    # Salva as alterações
    wb.save(caminho)