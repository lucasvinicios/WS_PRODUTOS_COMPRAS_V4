import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from sklearn.cluster import DBSCAN
from sklearn.preprocessing import StandardScaler

# Configurações de estilo que funcionam em qualquer versão
plt.style.use('ggplot')  # Estilo alternativo e moderno
plt.rcParams['figure.dpi'] = 100
plt.rcParams['font.size'] = 9

# Ler e preparar os dados
file_path = "../precos_supermercados.csv"
data = pd.read_csv(file_path)
X = data.iloc[:, 1:].values  # Ignorar coluna 'Item'

# Normalização importante para DBSCAN
scaler = StandardScaler()
X_normalized = scaler.fit_transform(X.T)  # Transpor para agrupar supermercados

# Parâmetros DBSCAN - ajuste conforme necessário
# dbscan = DBSCAN(eps=1.5, min_samples=2)
# dbscan = DBSCAN(eps=0.5, min_samples=3) 
dbscan = DBSCAN(eps=2.0, min_samples=2)
clusters = dbscan.fit_predict(X_normalized)

# Resultados organizados
supermercados = data.columns[1:]
resultado = pd.DataFrame({
    'Supermercado': supermercados,
    'Cluster': clusters,
    'Tipo': ['Ruído' if x == -1 else f'Cluster {x}' for x in clusters]
})

print("\n📊 Resultados da Clusterização:")
print(resultado.sort_values(by='Cluster'))

# Função de plotagem melhorada
def plot_dbscan_clear(X, labels, names):
    plt.figure(figsize=(10, 6))
    
    # Gerar cores automaticamente
    unique_labels = set(labels)
    colors = plt.cm.tab10(np.linspace(0, 1, len(unique_labels)))
    
    for i, label in enumerate(unique_labels):
        # Selecionar pontos do cluster
        cluster_mask = labels == label
        x = X[cluster_mask, 0]
        y = X[cluster_mask, 1]
        
        if label == -1:
            # Ruído
            plt.scatter(x, y, c='gray', marker='x', s=100,
                       linewidths=1, alpha=0.7, label='Ruído')
        else:
            # Clusters
            plt.scatter(x, y, c=[colors[i]], s=150, edgecolor='black',
                       linewidth=0.7, alpha=0.8, label=f'Cluster {label}')
            
            # Anotar nomes
            for j, (x_pos, y_pos) in enumerate(zip(x, y)):
                plt.annotate(names[cluster_mask][j], (x_pos, y_pos),
                           xytext=(7, 7), textcoords='offset points',
                           fontsize=8, ha='left')
    
    # Elementos gráficos
    plt.title("Análise de Similaridade entre Supermercados\n(DBSCAN)", pad=20)
    plt.xlabel("Componente Principal 1 (Preços Normalizados)")
    plt.ylabel("Componente Principal 2 (Preços Normalizados)")
    
    plt.legend(bbox_to_anchor=(1.05, 1), loc='upper left')
    plt.grid(True, linestyle=':', alpha=0.4)
    plt.tight_layout()
    plt.show()

# Visualização
plot_dbscan_clear(X_normalized, clusters, supermercados)