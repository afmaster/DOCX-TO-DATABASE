import os
from docx import Document
import glob

import sqlite3

def add_entry(db_file: str, db: str, dic: dict) -> None:
    conn = sqlite3.connect(db_file)
    c = conn.cursor()
    columns_length = len(list(dic.items()))

    # Creating string for sql command for creating bd
    part_1_create_table_sql = f"CREATE TABLE IF NOT EXISTS {db} ("
    for r in range(0, columns_length):
        part_1_create_table_sql = part_1_create_table_sql + str(list(dic.keys())[r]) + " TEXT"
        if r < (columns_length - 1):
            part_1_create_table_sql = part_1_create_table_sql + ", "

    create_table_sql = part_1_create_table_sql + ");"

    c.execute(create_table_sql)

    sqlite_insert_row = f"INSERT INTO {db} VALUES ("
    for r in range(0, columns_length):
        sqlite_insert_row = sqlite_insert_row + "?"
        if r < (columns_length - 1):
            sqlite_insert_row = sqlite_insert_row + ", "
    sqlite_insert_row = sqlite_insert_row + ")"
    params = tuple(dic.values())
    c.execute(sqlite_insert_row, params)
    conn.commit()
    c.close()
    conn.close()


def delete_entry(db_file: str, db: str, field: str, criteria: str or int) -> None:
    conn = sqlite3.connect(db_file)
    c = conn.cursor()
    try:
        c.execute(f"DELETE from {db} where {field}= ?", (criteria,))
    except:
        pass
    conn.commit()
    c.close()
    conn.close()

def change_row(db_file: str, db: str, field: str, criteria: str, dic: dict) -> None:
    try:
        delete_entry(db_file, db, field, criteria)
    except:
        pass
    add_entry(db_file, db, dic)


caminho_pasta = os.getcwd()

arquivos_doc = glob.glob(os.path.join(caminho_pasta, '**/*.docx'), recursive=True)

for arquivo in arquivos_doc:
    print(arquivo)
    # Abre o documento usando a biblioteca python-docx
    try:
        doc = Document(arquivo)
    except:
        print('falhou')
        continue

    nome_arquivo = os.path.splitext(os.path.basename(arquivo))[0]

    texto_documento = ''
    for paragrafo in doc.paragraphs:
        texto_documento += paragrafo.text


    db_file = "yor db name"
    db = "your table name"

    dic = {
            'standard_texts_title': nome_arquivo,
            'st_content': texto_documento 
    }

    change_row(
        db_file,
        db,
        'standard_texts_title',
        nome_arquivo,
        dic)

