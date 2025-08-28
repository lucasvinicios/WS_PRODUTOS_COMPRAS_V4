import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt

# Ler o arquivo CSV
dados = pd.read_csv('../precos_supermercados.csv')

# Transformar os dados para o formato necessário (produtos como índice e supermercados como colunas)
dados.set_index('Item', inplace=True)

# Criar o heatmap
plt.figure(figsize=(10, 6))
sns.heatmap(dados, annot=True, cmap='coolwarm', fmt=".2f", linewidths=0.5)

# Título
plt.title("Heatmap de Preços por Supermercado")
plt.xlabel("Supermercado")
plt.ylabel("Produto")
plt.savefig("./Resultados/Heatmap - Preços por Supermercado.png")
# Exibir o gráfico
plt.show()
