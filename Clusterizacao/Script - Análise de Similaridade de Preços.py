import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from sklearn.cluster import KMeans

# Ler o arquivo CSV
file_path = "../precos_supermercados.csv"
data = pd.read_csv(file_path)

# Preparar os dados (preços dos produtos)
X = data.iloc[:, 1:].values  # Ignorar a coluna "Item"

# Aplicar o K-Means
n_clusters = 3  # Defina o número de clusters
kmeans = KMeans(n_clusters=n_clusters, random_state=42)
kmeans.fit(X)

# Adicionar os rótulos dos clusters ao DataFrame
data['Cluster'] = kmeans.labels_

# Função para encontrar os produtos mais próximos dos centroides
def get_closest_items_to_centroids(X, labels, centroids, item_names, n_items=3):
    closest_items = {}
    for cluster_id in range(len(centroids)):
        # Obter índices dos pontos no cluster atual
        cluster_indices = np.where(labels == cluster_id)[0]
        
        # Calcular distâncias dos pontos para o centroide do cluster
        distances = np.linalg.norm(X[cluster_indices] - centroids[cluster_id], axis=1)
        
        # Obter os índices dos pontos mais próximos
        closest_indices = cluster_indices[np.argsort(distances)[:n_items]]
        
        # Obter os nomes dos produtos mais próximos
        closest_names = item_names[closest_indices]
        
        closest_items[f'Cluster {cluster_id}'] = closest_names.tolist()
    
    return closest_items

# Obter os nomes dos produtos mais próximos aos centroides
item_names = data['Item'].values
closest_items = get_closest_items_to_centroids(X, kmeans.labels_, kmeans.cluster_centers_, item_names)

# Exibir os produtos mais próximos de cada centroide
print("\nProdutos mais próximos de cada centroide:")
for cluster, items in closest_items.items():
    print(f"{cluster}: {', '.join(items)}")

# Exibir todos os clusters
print("\nClusters encontrados pelo K-Means:")
print(data[['Item', 'Cluster']])

# Função para plotar os clusters com anotações
def plot_clusters_with_labels(X, labels, item_names):
    plt.figure(figsize=(12, 10))
    
    # Plotar os produtos
    for i in range(n_clusters):
        cluster_points = X[labels == i]
        cluster_names = item_names[labels == i]
        plt.scatter(cluster_points[:, 0], cluster_points[:, 1], label=f'Cluster {i}', alpha=0.6)
        
        # Anotar alguns pontos com os nomes dos produtos
        for j, (x, y) in enumerate(cluster_points[:10]):  # Mostrar apenas os 5 primeiros de cada cluster
            plt.annotate(cluster_names[j], (x, y), textcoords="offset points", xytext=(0,5), ha='center')
    
    # Plotar os centros dos clusters
    plt.scatter(kmeans.cluster_centers_[:, 0], kmeans.cluster_centers_[:, 1], 
                s=200, c='red', marker='X', label='Centros')
    
    # Ajustes do gráfico
    plt.title("Clusterização de Produtos por Preço (K-Means)", fontsize=14)
    plt.xlabel("Preço no Supermercado 1")
    plt.ylabel("Preço no Supermercado 2")
    plt.legend()
    plt.grid(True, alpha=0.2)
    plt.show()

# Reduzir a dimensionalidade para visualização (usando apenas 2 supermercados)
plot_clusters_with_labels(X[:, :2], kmeans.labels_, item_names)

