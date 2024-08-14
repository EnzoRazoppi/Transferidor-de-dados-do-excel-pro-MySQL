import pandas as pd
import mysql.connector

# Ler planilha
sheeti_df = pd.read_excel("Teste.xlsx")

# Abrir conexão com o MySQL fora do loop
conectar = mysql.connector.connect(host='localhost', database='excel', user='root', password='')
cursor = conectar.cursor()

for i in range(len(sheeti_df)):
    nome = sheeti_df.loc[i, "nome"]
    idade = sheeti_df.loc[i, "idade"]
    email = sheeti_df.loc[i, "email"]

    # Verificar se o nome já existe na tabela
    cursor.execute("SELECT COUNT(*) FROM excel_teste WHERE nome = %s", (nome,))
    if cursor.fetchone()[0] > 0:
        print(f"Nome {nome} já existe na tabela. Dados não inseridos.")
        continue

    # Preparar comando de inserção
    dados = f"('{nome}', {idade}, '{email}')"
    comando = "INSERT INTO `excel_teste` (`nome`, `idade`, `email`) VALUES "
    sql = comando + dados

    try:
        cursor.execute(sql)
        conectar.commit()
        print(f"Dados inseridos com sucesso para: {nome}")
    except mysql.connector.Error as err:
        print(f"Erro ao inserir dados para {nome}: {err}")
        continue

# Fechar a conexão fora do loop
conectar.close()
print("Conexão com o banco de dados fechada.")
