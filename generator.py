from pathlib import Path
from datetime import datetime
from docxtpl import DocxTemplate
from docx2pdf import convert
import os
import json

path_config = Path(__file__).parent / "config"  

def write_json(dic):
    # Escreve dic em cima do .json anterior
    # JSON aberto
    with open(path_config / "dic.json", 'w') as file:
        json.dump(dic, file)

def read_json():
    # Abre o dic.json
    # Saída: dict do .json

    # JSON aberto
    with open(path_config / "dic.json", 'r') as file:
        dic = json.load(file)
    return dic

def create_pdf(dic):
    """
    Utiliza o dic para substituir no arquivo docx, cria um .pdf
    entrada: dictionary
    saída: cria arquivo
    """
    # Datetime
    now = datetime.now()

    # Documento aberto
    doc = DocxTemplate(path_config / "Contrato_template.docx")

    # Gerador de .docx e conversor para .pdf
    final_path = Path(__file__).parent / f"Contrato {now.strftime(f'%H_%M_%S')}.docx"
    doc.render(dic)
    doc.save(final_path)
    convert(final_path)
    os.remove(final_path)

if __name__ == '__main__':
    create_pdf()