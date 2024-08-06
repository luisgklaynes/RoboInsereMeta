import psycopg2
from dotenv import load_dotenv
import os
import glob
import pandas as pd

load_dotenv()
def conexao(database, server=os.getenv('DB_HOST'),port=os.getenv('DB_PORT')):
    try:
        conn = psycopg2.connect(
            host=server,
            port=port,
            database=database,
            user=os.getenv('DB_USER'),
            password=os.getenv('DB_PASSWORD')
        )
    except psycopg2.Error as erro:
        print(erro)
        conn=None
    return conn

def buscarNomeArquivo():
    '''Busca o arquivo mais atual salvo/editado (xls e xlsx)'''
    directory = r'C:\arquivosExcel'
    list_of_files = glob.glob(os.path.join(directory, '*.xlsx')) + glob.glob(os.path.join(directory, '*.xls'))
    if not list_of_files:
        return None
    latest_file = max(list_of_files, key=os.path.getctime)
    return latest_file

def comparacao(doc,doc1):

    '''# Faz comparação para encontrar possível inconsistência'''
    if 'Proporção/dia' in doc.columns and 'Meta' in doc1.columns:
        doc = doc['Proporção/dia'].sum()
        doc1 = doc1['Meta'].sum()
        if doc >= (doc1+1000) or doc <= (doc1-1000):
            print(f'Possível incosistência no arquivo, valores divergêntes | Meta {doc}, curvas e volume {doc1}')
            return False
        else:
            print('Arquivo validado')
            return True
    else:
        print(f'Falha ao somar as colunas Proporção/dia e Meta')

def leituraArquivo(nomearquivo):
    '''Faz a leitura do arquivo, recebe o nome do arquivo'''
    try:
        doc = pd.read_excel(nomearquivo,sheet_name='Meta')
        doc1 = pd.read_excel(nomearquivo,header=1, sheet_name='curvas e volume')
    except FileNotFoundError:
        print(f"Erro: O arquivo '{nomearquivo}' não foi encontrado.")
        return None, None
    except ValueError as e:
        print(f"Erro de valor: {e}")
        return None, None
    except Exception as e:
        print(f"Ocorreu um erro inesperado: {e}")
        return None, None
    return doc,doc1

def insertData(dados):
    dadaini = dados['Data'].min()
    datafim = dados['Data'].max()
    dados = dados.values.tolist()
    conn = conexao('bd_planejamento')
    if conn is not None:
        consulta = conn.cursor()
        try:
            consulta.execute(
                f"DELETE FROM sc_indicadores.tb_meta_ta WHERE data BETWEEN '{dadaini}' AND '{datafim}'"
            )
            resumoDelete = (consulta.rowcount)
            consulta.executemany(
                'INSERT INTO sc_indicadores.tb_meta_ta (data,intervalo,proporcao_min,proporcao_min_dia) VALUES(%s,%s,%s,%s)',dados
            )
            resumoInsert = (consulta.rowcount)
            conn.commit()
            consulta.close()
            conn.close()
        except psycopg2.Error as erro:
            print(erro)
            return None, None
    else:
        return None, None
    return resumoDelete, resumoInsert

latest_file = buscarNomeArquivo()

if latest_file:
    doc,doc1 = leituraArquivo(latest_file)
    if comparacao(doc,doc1):
        tr1 = doc.drop(columns=['Operadores','Pausa','proporção','Unnamed: 6'])
        tr1['Intervalo'] = tr1['Intervalo'].replace('00:01:00','00:00:00')
        tr1['Proporção'] = tr1['Proporção/dia']
        print(f"O arquivo mais recente é: {latest_file}")
        resumoDelete,resumoInsert = insertData(tr1)
        print(f'Deletado {resumoDelete} registros | Inserido {resumoInsert} registros')
else:
    print("Não há arquivos no diretório especificado.")