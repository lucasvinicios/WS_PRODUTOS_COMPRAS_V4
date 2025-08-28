import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from sklearn.cluster import KMeans
from sklearn.decomposition import PCA

# Configurações de visualização
plt.style.use('ggplot')
plt.rcParams['figure.dpi'] = 100
plt.rcParams['font.size'] = 9

# Ler o arquivo CSV
file_path = "../precos_supermercados.csv"
data = pd.read_csv(file_path)

# Preparar os dados
X = data.iloc[:, 1:].values  # Ignorar a coluna "Item"

# Reduzir dimensionalidade com PCA
pca = PCA(n_components=2)
X_reduced = pca.fit_transform(X.T)  # Transpor para agrupar supermercados

# Aplicar K-Means
n_clusters = 3
kmeans = KMeans(n_clusters=n_clusters, random_state=42)
kmeans.fit(X_reduced)

# Criar DataFrame com resultados
supermercados = data.columns[1:]
clusters_df = pd.DataFrame({
    'Supermercado': supermercados,
    'Cluster': kmeans.labels_,
    'PC1': X_reduced[:, 0],
    'PC2': X_reduced[:, 1]
})

print("Clusters encontrados pelo K-Means:")
print(clusters_df.sort_values(by='Cluster'))

# Função de plotagem melhorada
def plot_clusters_with_labels(X, labels, names):
    plt.figure(figsize=(12, 8))
    
    # Cores para cada cluster
    colors = plt.cm.tab10(np.linspace(0, 1, n_clusters))
    
    # Plotar pontos e centros
    for i in range(n_clusters):
        # Pontos do cluster
        plt.scatter(X[labels == i, 0], X[labels == i, 1], 
                   c=[colors[i]], label=f'Cluster {i}', 
                   s=150, alpha=0.7, edgecolor='black')
        
        # Anotar nomes dos supermercados
        for x, y, name in zip(X[labels == i, 0], X[labels == i, 1], names[labels == i]):
            plt.annotate(name, (x, y), 
                        textcoords="offset points", 
                        xytext=(5,5), ha='left', fontsize=8)
    
    # Centros dos clusters
    plt.scatter(kmeans.cluster_centers_[:, 0], kmeans.cluster_centers_[:, 1],
               s=250, c='red', marker='X', label='Centróides',
               edgecolor='black', linewidth=1)
    
    # Ajustes estéticos
    plt.title("Segmentação de Supermercados por Perfil de Preços\n(PCA + K-Means)", pad=20)
    plt.xlabel(f"Componente Principal 1 ({pca.explained_variance_ratio_[0]:.1%} da variância)")
    plt.ylabel(f"Componente Principal 2 ({pca.explained_variance_ratio_[1]:.1%} da variância)")
    
    plt.legend(bbox_to_anchor=(1.05, 1), loc='upper left')
    plt.grid(True, linestyle=':', alpha=0.4)
    plt.tight_layout()
    plt.show()

# Plotar com nomes
plot_clusters_with_labels(X_reduced, kmeans.labels_, supermercados)