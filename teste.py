import pandas as pd

# df = pd.read_excel("NOVA - Separação - Pessoal Eventos.xlsx")
#
# print(df.shape[0])
# print(df)
#
df = pd.read_excel('NOVA - Separação - Pessoal Eventos.xlsx', None)

table = []
for x in range(1, len(df)+1):
    table.append(pd.read_excel('NOVA - Separação - Pessoal Eventos.xlsx', +str(x)))

print(table)

pd.read_excel()
