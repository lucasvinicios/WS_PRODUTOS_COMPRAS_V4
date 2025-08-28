import pandas as pd
import io
import matplotlib.pyplot as plt
from openpyxl.drawing.image import Image as OpenpyxlImage
import os

def realizar_analise_reversa_acumulativa(arquivos_csv):
    """
    Realiza a análise de preços de produtos, encontrando o item mais barato
    em cada mercado para a cesta de compras.
    """
    mercados_data = {}
    
    for nome_arquivo_completo, conteudo_csv in arquivos_csv.items():
        df = pd.read_csv(io.StringIO(conteudo_csv), sep=',')
        nome_mercado_curto = df.columns[1]
        df.set_index('Produto', inplace=True)
        
        frete = df.loc['Frete'].iloc[0]
        df_produtos = df.drop(['Frete', 'Valor Mínimo', 'Valor Total', 'Total Baratos + Frete'], errors='ignore')
        
        mercados_data[nome_arquivo_completo] = {
            'df': df_produtos,
            'frete': frete,
            'nome_curto': nome_mercado_curto
        }

    resultados_finais = []
    
    custos_totais = pd.DataFrame()
    for nome_arquivo_completo, data in mercados_data.items():
        df_temp = data['df'].copy()
        df_temp['Custo Total'] = df_temp.iloc[:, 0] + data['frete']
        custos_totais[nome_arquivo_completo] = df_temp['Custo Total']
    
    for produto in custos_totais.index:
        custos_ordenados = custos_totais.loc[produto].dropna().sort_values()
        
        if not custos_ordenados.empty:
            mercado_mais_barato_key = custos_ordenados.index[0]
            preco_produto = mercados_data[mercado_mais_barato_key]['df'].loc[produto].iloc[0]
            frete = mercados_data[mercado_mais_barato_key]['frete']
            custo_total = preco_produto + frete
            
            resultados_finais.append({
                'Produto': produto,
                'Supermercado Recomendado': mercado_mais_barato_key,
                'Preço do Produto': preco_produto,
                'Frete': frete,
                'Custo Total': custo_total,
            })
            
    return pd.DataFrame(resultados_finais).sort_values(by='Produto'), mercados_data


def gerar_aba_resumo(writer, analise_final_df, mercados_data):
    """
    Gera uma aba de resumo com os principais resultados da análise e um gráfico visual.
    """
    total_gasto_cesta_ideal = analise_final_df['Custo Total'].sum()
    
    custos_por_mercado = {}
    for nome_arquivo_completo, data in mercados_data.items():
        nome_mercado_curto = data['nome_curto']
        frete = data['frete']
        total_gasto = data['df'].iloc[:, 0].sum() + frete
        custos_por_mercado[nome_mercado_curto] = total_gasto
    
    df_custos = pd.DataFrame(list(custos_por_mercado.items()), columns=['Supermercado', 'Custo Total da Cesta'])
    df_custos.set_index('Supermercado', inplace=True)
    df_custos.loc['Cesta Ideal'] = total_gasto_cesta_ideal
    
    # Gera e salva o gráfico de barras
    plt.style.use('seaborn-v0_8-whitegrid')
    fig, ax = plt.subplots(figsize=(10, 6))
    
    # Cores personalizadas
    cores = ['#3CB371' for _ in df_custos.index]
    
    # Define a cor da "Cesta Ideal" para destaque
    if 'Cesta Ideal' in df_custos.index:
        cores[df_custos.index.get_loc('Cesta Ideal')] = '#00BFFF'
    
    df_custos['Custo Total da Cesta'].plot(kind='bar', ax=ax, color=cores)
    
    ax.set_title('Comparativo de Custo Total da Cesta de Compras', fontsize=16)
    ax.set_xlabel('Supermercado', fontsize=12)
    ax.set_ylabel('Custo Total (R$)', fontsize=12)
    ax.tick_params(axis='x', rotation=45)
    
    plt.tight_layout()
    grafico_path = 'comparativo_custo.png'
    plt.savefig(grafico_path)
    plt.close()

    # Escreve os dados na planilha
    df_custos.to_excel(writer, sheet_name='Resumo da Análise')
    
    # Adiciona a imagem do gráfico à planilha usando openpyxl
    workbook = writer.book
    worksheet = workbook['Resumo da Análise']
    img = OpenpyxlImage(grafico_path)
    worksheet.add_image(img, 'E2')
    
    # Remove o arquivo de imagem temporário
    os.remove(grafico_path)
    
    print("Aba 'Resumo da Análise' gerada com sucesso, incluindo o gráfico visual.")


def gerar_excel_final(analise_final_df, mercados_data, nome_arquivo='analise_supermercados_final_resumo.xlsx'):
    """
    Gera um arquivo Excel com abas para cada supermercado recomendado e um resumo.
    """
    custos_totais_com_nomes_curtos = pd.DataFrame()
    for nome_arquivo_completo, data in mercados_data.items():
        nome_mercado_curto = data['nome_curto']
        df_temp = data['df'].copy()
        df_temp['Custo Total'] = df_temp.iloc[:, 0] + data['frete']
        custos_totais_com_nomes_curtos[nome_mercado_curto] = df_temp['Custo Total']
    
    economias = {}
    for produto in custos_totais_com_nomes_curtos.index:
        custos_ordenados = custos_totais_com_nomes_curtos.loc[produto].sort_values()
        custo_melhor = custos_ordenados.iloc[0]
        try:
            custo_segundo_melhor = custos_ordenados.iloc[1]
            economia_calculada = custo_segundo_melhor - custo_melhor
            economias[produto] = economia_calculada
        except IndexError:
            economias[produto] = 0

    mercados_recomendados_keys = analise_final_df['Supermercado Recomendado'].unique()
    
    mercados_recomendados_ordenados = sorted(
        mercados_recomendados_keys,
        key=lambda x: int(x.split(' ')[1]) if 'TOP' in x else float('inf')
    )

    with pd.ExcelWriter(nome_arquivo, engine='openpyxl') as writer:
        
        gerar_aba_resumo(writer, analise_final_df, mercados_data)

        for nome_mercado_completo in mercados_recomendados_ordenados:
            data = mercados_data[nome_mercado_completo]
            nome_mercado_curto = data['nome_curto']
            df_mercado = data['df'].copy()
            frete = data['frete']
            
            df_mercado['Custo Total (c/ Frete)'] = df_mercado.iloc[:, 0] + frete
            df_mercado['Economia (vs 2º Melhor)'] = 0.00
            
            for _, row in analise_final_df[analise_final_df['Supermercado Recomendado'] == nome_mercado_completo].iterrows():
                produto = row['Produto']
                df_mercado.loc[produto, 'Economia (vs 2º Melhor)'] = economias.get(produto, 0)
            
            total_economizado = df_mercado['Economia (vs 2º Melhor)'].sum()

            df_frete = pd.DataFrame({'Frete': frete, 'Custo Total (c/ Frete)': ''}, index=['Frete'])
            df_frete.index.name = 'Produto'
            df_final = pd.concat([df_mercado.reset_index().set_index('Produto'), df_frete])

            df_total = pd.DataFrame({'Economia (vs 2º Melhor)': total_economizado}, index=['Total Economizado'])
            df_total.index.name = 'Produto'
            df_final = pd.concat([df_final, df_total])

            df_final.rename(columns={df_final.columns[0]: 'Preço Unitário'}, inplace=True)
            
            nome_aba = nome_mercado_completo.replace(' ', ' - ', 1)
            
            df_final.to_excel(writer, sheet_name=nome_aba, index=True)
            print(f"Aba '{nome_aba}' gerada com sucesso. Total economizado: R$ {total_economizado:.2f}")


arquivos_do_excel = {
    "TOP 1 TENDAATACADO": """Produto,TENDAATACADO
Arroz 5kg,17.4
Feijão 1kg,3.55
Macarrão 500g,2.59
Óleo 900ml,6.19
Açúcar 5kg,18.29
Leite Integral 1L,5.45
Pão de Forma 500g,5.59
Café 500g,21.99
Detergente 500ml,1.8
Sabão em Pó 1kg,9.25
Papel Higiênico,8.2
Creme Dental 70g,2.6
Água Sanitária,3.45
Sabonete,1.05
Fio Dental,5.15
Molho de Tomate,1.25
Azeite 500ml,32.89
Farinha de Trigo 1kg,2.99
Queijo 200g,9.49
Creme de Leite 200g,2.25
Frete,14.9
Valor Mínimo,0
Valor Total,176.32
Total Baratos + Frete,118.94
""",
    "TOP 2 BOASUPERMERCADO": """Produto,BOASUPERMERCADO
Arroz 5kg,20.9
Feijão 1kg,4.69
Macarrão 500g,3.49
Óleo 900ml,7.89
Açúcar 5kg,22.79
Leite Integral 1L,5.19
Pão de Forma 500g,6.59
Café 500g,28.79
Detergente 500ml,1.85
Sabão em Pó 1kg,4.19
Papel Higiênico,4.09
Creme Dental 70g,1.55
Água Sanitária,4.99
Sabonete,1.55
Fio Dental,5.99
Molho de Tomate,1.79
Azeite 500ml,39.9
Farinha de Trigo 1kg,3.49
Queijo 200g,10.99
Creme de Leite 200g,3.45
Frete,15
Valor Mínimo,0
Valor Total,199.16
Total Baratos + Frete,24.83
""",
    "TOP 3 TAUSTE": """Produto,TAUSTE
Arroz 5kg,19.89
Feijão 1kg,5.77
Macarrão 500g,3.98
Óleo 900ml,6.97
Açúcar 5kg,16.89
Leite Integral 1L,4.59
Pão de Forma 500g,6.86
Café 500g,27.69
Detergente 500ml,1.79
Sabão em Pó 1kg,4.59
Papel Higiênico,6.37
Creme Dental 70g,1.79
Água Sanitária,4.59
Sabonete,1.19
Fio Dental,7.69
Molho de Tomate,1.78
Azeite 500ml,38.89
Farinha de Trigo 1kg,3.29
Queijo 200g,8.77
Creme de Leite 200g,2.98
Frete,14.9
Valor Mínimo,0
Valor Total,191.26
Total Baratos + Frete,30.05
""",
    "TOP 4 CONFIANCA": """Produto,CONFIANCA
Arroz 5kg,19.89
Feijão 1kg,5.77
Macarrão 500g,2.98
Óleo 900ml,6.98
Açúcar 5kg,16.79
Leite Integral 1L,5.19
Pão de Forma 500g,6.7
Café 500g,27.59
Detergente 500ml,1.98
Sabão em Pó 1kg,9.78
Papel Higiênico,4.79
Creme Dental 70g,1.79
Água Sanitária,4.79
Sabonete,1.77
Fio Dental,9.68
Molho de Tomate,1.15
Azeite 500ml,38.9
Farinha de Trigo 1kg,3.39
Queijo 200g,8.98
Creme de Leite 200g,2.79
Frete,18.9
Valor Mínimo,0
Valor Total,200.58
Total Baratos + Frete,36.84
""",
    "TOP 5 BARBOSA": """Produto,BARBOSA
Arroz 5kg,19.99
Feijão 1kg,5.99
Macarrão 500g,2.79
Óleo 900ml,7.49
Açúcar 5kg,24.95
Leite Integral 1L,5.49
Pão de Forma 500g,6.49
Café 500g,27.99
Detergente 500ml,2.19
Sabão em Pó 1kg,7.99
Papel Higiênico,7.99
Creme Dental 70g,2.99
Água Sanitária,4.49
Sabonete,0
Fio Dental,8.99
Molho de Tomate,1.49
Azeite 500ml,37.99
Farinha de Trigo 1kg,3.99
Queijo 200g,8.99
Creme de Leite 200g,3.49
Frete,20.9
Valor Mínimo,100
Valor Total,212.67
Total Baratos + Frete,20.9
""",
    "TOP 6 COOPSUPERMERCADO": """Produto,COOPSUPERMERCADO
Arroz 5kg,25.99
Feijão 1kg,5.99
Macarrão 500g,3.59
Óleo 900ml,7.49
Açúcar 5kg,24.99
Leite Integral 1L,5.99
Pão de Forma 500g,6.79
Café 500g,23.79
Detergente 500ml,1.79
Sabão em Pó 1kg,19.29
Papel Higiênico,9.29
Creme Dental 70g,3.49
Água Sanitária,4.99
Sabonete,1.59
Fio Dental,15.19
Molho de Tomate,1.69
Azeite 500ml,42.49
Farinha de Trigo 1kg,4.29
Queijo 200g,11.99
Creme de Leite 200g,3.39
Frete,15
Valor Mínimo,0
Valor Total,239.1
Total Baratos + Frete,16.79
"""
}

# 1. Executa a análise para obter o resultado final e os dados brutos
analise_completa, mercados_dados = realizar_analise_reversa_acumulativa(arquivos_do_excel)

# 2. Gera o novo arquivo Excel
gerar_excel_final(analise_completa, mercados_dados)