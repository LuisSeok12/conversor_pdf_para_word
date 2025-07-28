import os
from pdf2docx import Converter
import tkinter as tk
from tkinter import filedialog, messagebox


def converter_pdf_para_docx():
    """
    Converte um arquivo PDF em DOCX mantendo a formatação o mais fiel possível.
    O arquivo convertido será salvo no mesmo diretório.
    """

    PDF_FILE = filedialog.askopenfilename(title="Selecione o PDF para ser convertido: ",
                                          filetypes=[("Arquivos Excel", "*.xlsx *.xls")])
    DOCX_FILE = "saida.docx"    

    # Verificação simples se o PDF existe no diretório
    if not os.path.isfile(PDF_FILE):
        print(f"[ERRO] O arquivo '{PDF_FILE}' não foi encontrado no diretório atual.")
        return

    # Verifica se a extensão de entrada é .pdf
    if not PDF_FILE.lower().endswith(".pdf"):
        print(f"[ERRO] O arquivo '{PDF_FILE}' não é um PDF válido.")
        return

    try:
        # Cria o conversor
        converter = Converter(PDF_FILE)
        # Converte todas as páginas (start=0, end=None)
        converter.convert(DOCX_FILE, start=0, end=None)
        converter.close()

        print(f"[SUCESSO] Conversão concluída! O arquivo gerado foi: '{DOCX_FILE}'")

    except Exception as e:
        print(f"[ERRO] Falha na conversão do arquivo '{PDF_FILE}': {str(e)}")


if __name__ == "__main__":
    root = tk.TK()
    root.withdraw()
    converter_pdf_para_docx()

