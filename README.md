## Funções utilizadas:
#### conexao 
* Faz a conexão com o banco de dados Postgres utilizando a biblioteca "psycopg2". Recebe parâmetros de nome do servidor e banco de dados.
#### buscarNomeArquivo   
* Busca o último arquivo no diretório (salvo/editado) nos formatos xls e xlsx.
#### leituraArquivo
* Faz a leitura do arquivo validando a estrutura das abas existentes no arquivo.
#### compraracao
* Faz a leitura do arquivo, considerando diferentes abas do Excel, faz a validação do arquivo realizando soma e comparação colunas baseado na regra de negócio.
#### insertData
* Faz conexão com o banco de dados, consulta se a informação encontrada no arquivo já está presente na base de dados, caso já exista ele apenas atualiza, do contrário ele insere toda a informação.
