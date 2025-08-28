import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from scipy.cluster.hierarchy import dendrogram, linkage

# Configurações iniciais para melhorar a visualização
plt.rcParams['figure.dpi'] = 120  # Aumentar resolução
plt.rcParams['font.size'] = 8     # Tamanho da fonte

# Ler o arquivo CSV
file_path = "../precos_supermercados.csv"
data = pd.read_csv(file_path)

# Preparar os dados (preços dos produtos)
X = data.iloc[:, 1:].values  # Ignorar a coluna "Item"

# Calcular a matriz de ligação
Z = linkage(X, method='ward')

# Criar figura com tamanho adequado
plt.figure(figsize=(10, 8))  # Aumentar a altura para caber os nomes

# Plotar o dendrograma com parâmetros otimizados
dendrogram(
    Z,
    labels=data['Item'].values,
    leaf_rotation=90,           # Rotacionar 90 graus
    leaf_font_size=10,           # Tamanho da fonte dos itens
    orientation='top',          # Orientação padrão (opcional: 'left')
    color_threshold=0.7*max(Z[:,2]),  # Limiar para cores diferentes
    above_threshold_color='grey'      # Cor para clusters acima do limiar
)

# Ajustes finais
plt.title("Dendrograma de Similaridade entre Produtos\n(Agrupamento Hierárquico)", 
          pad=20, fontsize=12)
plt.xlabel("Produtos", labelpad=10)
plt.ylabel("Distância (Ward)", labelpad=10)
plt.grid(False)  # Remover grid para melhor visualização

# Ajustar layout automaticamente
plt.tight_layout()

# Mostrar o gráfico
plt.show()