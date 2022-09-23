# faturamento-diario
## Relatório diário do faturamento da empresa enviado por e-mail


### View em T-SQL: (view-faturamento.sql)
O objetivo da view é criar a relação entre as notas fiscais e notas fiscais de serviço com os orçamentos faturados. Para facilitar o entendimento de cada faturamento, é exibido também o nº do serviço.


### Script Python: (faturamento_diario.py)
O script é dividido em três etapas:
#### Excel:
Abrir uma instância excel, abrir o arquivo com conexão ao banco de dados, atualizar seus vínculos e executar a macro exportpic.
Em seguida, salvar e encerrar a instância.
#### Pandas:
Filtrar apenas os faturamentos do mês, renomear algumas colunas, formatar valores e gerar um dataframe para ser exibido em um corpo de e-mail.
#### Outlook:
Abrir uma instância outlook, enviar um e-mail com a imagem gerada pela macro exportpic e o dataframe da tabela do faturamento mensal.


### VBA para exportar imagem: (vba-exportpic.bas)
Exporta o conjunto de células pré-determinados como uma imagem jpg no destino desejado.