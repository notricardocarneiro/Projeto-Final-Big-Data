import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# Passo 1: Importar as bibliotecas necessárias

# Passo 2: Carregar os dados do arquivo Excel em um DataFrame
vendas_df = pd.read_excel('resumovendas2023.xlsx')

# Convertendo a coluna 'Data' para o tipo datetime
vendas_df['Data'] = pd.to_datetime(vendas_df['Data'])

# Mostrando as primeiras linhas do DataFrame
print(vendas_df.head())

# Somando o valor total das vendas
soma_valor_bruto = vendas_df['Valor Total'].sum()

# Mostrando o valor total das vendas
print("Soma dos valores brutos:", soma_valor_bruto)

# Separando a coluna data em mês e ano
vendas_df['Ano'] = vendas_df['Data'].dt.year
vendas_df['Mes'] = vendas_df['Data'].dt.month

# Agrupando as vendas por mês do ano 2023
vendas_por_mes_df_agg = vendas_df.groupby(['Ano', 'Mes']).agg({
    'Data': 'size',        # Conta o número de linhas (vendas)
    'Valor Total': 'sum'   # Soma o valor total das vendas
}).reset_index()

# Renomeando as colunas para melhor compreensão
vendas_por_mes_df_agg.rename(columns={'Data': 'Quantidade de Vendas', 'Valor Total': 'Valor Total'}, inplace=True)

# Mostrando o valor total das vendas separado por mês e ano
print(vendas_por_mes_df_agg)

# Filtrando para o ano de 2023
vendas_2023 = vendas_por_mes_df_agg[vendas_por_mes_df_agg['Ano'] == 2023]

# Calculando a soma total das quantidades de vendas
soma_quantidade_vendas = vendas_2023['Quantidade de Vendas'].sum()

# Ajustando o estilo dos gráficos com seaborn
sns.set(style="darkgrid")

# Dicionário para mapear números dos meses para nomes
meses_nomes = {
    1: 'Janeiro', 2: 'Fevereiro', 3: 'Março', 4: 'Abril',
    5: 'Maio', 6: 'Junho', 7: 'Julho', 8: 'Agosto',
    9: 'Setembro', 10: 'Outubro', 11: 'Novembro', 12: 'Dezembro'
}

# Adicionando coluna com nomes dos meses
vendas_2023['Mes Nome'] = vendas_2023['Mes'].map(meses_nomes)

# Calculando a média das vendas mensais
media_vendas = vendas_2023['Valor Total'].mean()

# Gráfico de linha
plt.figure(figsize=(10, 6))
sns.lineplot(data=vendas_2023, x='Mes', y='Valor Total', marker='o', color='b')
plt.axhline(media_vendas, color='r', linestyle='--', label=f'Média das Vendas: R$ {media_vendas:,.2f}')
plt.title('Vendas Totais por Mês em 2023')
plt.xlabel('Mês')
plt.ylabel('Valor Total das Vendas')
plt.xticks(vendas_2023['Mes'])
plt.legend()
plt.grid(True)


# Gráfico de barras
plt.figure(figsize=(10, 6))
sns.barplot(data=vendas_2023, x='Mes Nome', y='Valor Total', hue='Valor Total', dodge=False, palette='coolwarm', legend=True)
plt.title('Vendas Totais por Mês em 2023')
plt.xlabel('Mês')
plt.ylabel('Valor Total das Vendas')
plt.xticks(rotation=45)

# Adicionando a soma total das vendas como texto no gráfico
plt.text(x=0.5, y=max(vendas_2023['Valor Total']) + 5000, s=f"Soma total das vendas: R$ {soma_valor_bruto:,.2f}", ha='center', va='bottom', fontsize=12, color='black', weight='bold')



# Gráfico de barras para quantidade de vendas
plt.figure(figsize=(10, 6))
sns.barplot(data=vendas_2023, x='Mes Nome', y='Quantidade de Vendas', hue='Mes Nome', dodge=False, palette='viridis', legend=False)
plt.title('Quantidade de Vendas por Mês em 2023')
plt.xlabel('Mês')
plt.ylabel('Quantidade de Vendas')
plt.xticks(rotation=45)

# Adicionando a soma total das quantidades de vendas como legenda
plt.text(0.5, max(vendas_2023['Quantidade de Vendas']) + 50, f"Soma total das vendas: {soma_quantidade_vendas}", ha='center', va='bottom', fontsize=12, color='black', weight='bold')

# Adicionando a quantidade de vendas como texto no gráfico
for index, row in vendas_2023.iterrows():
    plt.text(index, row['Quantidade de Vendas'] + 2, row['Quantidade de Vendas'], color='black', ha="center")

plt.show()