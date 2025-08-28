# import pandas as pd
# import io

# def realizar_analise_reversa_acumulativa(arquivos_csv):
#     """
#     Realiza a análise reversa de preços de produtos, acumulando os itens mais baratos
#     em cada mercado e eliminando os mercados não vantajosos.

#     Args:
#         arquivos_csv (dict): Um dicionário onde a chave é o nome do arquivo
#                               e o valor é o conteúdo do arquivo como string.

#     Returns:
#         tuple: Uma tupla contendo o DataFrame com a análise final e o
#                dicionário com os dados brutos de cada mercado.
#     """
    
#     mercados_data = {}
#     produtos_alocados = set()
#     resultados_finais = []

#     # 1. Processa e carrega os dados de todos os mercados
#     for nome_arquivo, conteudo_csv in arquivos_csv.items():
#         df = pd.read_csv(io.StringIO(conteudo_csv), sep=',')
#         nome_mercado = df.columns[1]
#         df.set_index('Produto', inplace=True)
        
#         frete = df.loc['Frete'].iloc[0]
#         df_produtos = df.drop(['Frete', 'Valor Mínimo', 'Valor Total', 'Total Baratos + Frete'], errors='ignore')
        
#         mercados_data[nome_mercado] = {
#             'df': df_produtos,
#             'frete': frete,
#         }
        
#     mercados_ordenados_nomes = sorted(
#         mercados_data.keys(),
#         key=lambda x: int(x.split()[1]) if 'TOP' in x else float('inf'),
#         reverse=True
#     )
    
#     # 2. Realiza a análise acumulativa
#     for i in range(len(mercados_ordenados_nomes) - 1):
#         mercado_atual_nome = mercados_ordenados_nomes[i]
#         mercado_anterior_nome = mercados_ordenados_nomes[i+1]
        
#         data_atual = mercados_data[mercado_atual_nome]
#         data_anterior = mercados_data[mercado_anterior_nome]
        
#         produtos_para_analisar = data_atual['df'].index.difference(produtos_alocados)
        
#         if produtos_para_analisar.empty:
#             continue
            
#         frete_atual = data_atual['frete']
#         frete_anterior = data_anterior['frete']
        
#         produtos_economicos_no_atual = []
        
#         for produto in produtos_para_analisar:
#             try:
#                 preco_atual = data_atual['df'].loc[produto].iloc[0]
#                 preco_anterior = data_anterior['df'].loc[produto].iloc[0]
                
#                 custo_total_atual = preco_atual + frete_atual
#                 custo_total_anterior = preco_anterior + frete_anterior
                
#                 if custo_total_atual < custo_total_anterior:
#                     produtos_economicos_no_atual.append(produto)
#             except KeyError:
#                 pass
                
#         if produtos_economicos_no_atual:
#             print(f"Análise: {mercado_atual_nome} vs {mercado_anterior_nome}")
#             print(f"Produtos mais baratos em {mercado_atual_nome}: {', '.join(produtos_economicos_no_atual)}")
#             print("-" * 50)
            
#             for produto in produtos_economicos_no_atual:
#                 preco = data_atual['df'].loc[produto].iloc[0]
#                 custo_total = preco + frete_atual
#                 resultados_finais.append({
#                     'Produto': produto,
#                     'Supermercado Recomendado': mercado_atual_nome,
#                     'Preço do Produto': preco,
#                     'Frete': frete_atual,
#                     'Custo Total': custo_total,
#                 })
            
#             produtos_alocados.update(produtos_economicos_no_atual)
            
#     # 3. Trata o último mercado (TOP 1) para os produtos restantes
#     mercado_top1_nome = mercados_ordenados_nomes[-1]
#     data_top1 = mercados_data[mercado_top1_nome]
#     frete_top1 = data_top1['frete']
#     produtos_restantes = data_top1['df'].index.difference(produtos_alocados)
    
#     if not produtos_restantes.empty:
#         print(f"Análise: Produtos restantes no {mercado_top1_nome}")
#         print(f"Produtos restantes: {', '.join(produtos_restantes)}")
#         print("-" * 50)
#         for produto in produtos_restantes:
#             preco = data_top1['df'].loc[produto].iloc[0]
#             custo_total = preco + frete_top1
#             resultados_finais.append({
#                 'Produto': produto,
#                 'Supermercado Recomendado': mercado_top1_nome,
#                 'Preço do Produto': preco,
#                 'Frete': frete_top1,
#                 'Custo Total': custo_total,
#             })
            
#     return pd.DataFrame(resultados_finais).sort_values(by='Produto'), mercados_data


# def gerar_excel_com_economia(analise_final_df, mercados_data, nome_arquivo='analise_supermercados_final.xlsx'):
#     """
#     Gera um arquivo Excel com abas para cada supermercado, adicionando uma coluna
#     de economia em cada uma.
#     """
#     # Cria uma tabela consolidada de custos totais para facilitar a busca
#     custos_totais = pd.DataFrame()
#     for nome_mercado, data in mercados_data.items():
#         df_temp = data['df'].copy()
#         df_temp['Custo Total'] = df_temp.iloc[:, 0] + data['frete']
#         custos_totais[nome_mercado] = df_temp['Custo Total']
    
#     # Identifica o melhor e o segundo melhor custo total para cada produto
#     economias = {}
#     for produto in custos_totais.index:
#         custos_ordenados = custos_totais.loc[produto].sort_values()
#         custo_melhor = custos_ordenados.iloc[0]
#         try:
#             custo_segundo_melhor = custos_ordenados.iloc[1]
#             economia_calculada = custo_segundo_melhor - custo_melhor
#             economias[produto] = economia_calculada
#         except IndexError:
#             # Caso haja apenas um mercado, a economia é 0
#             economias[produto] = 0

#     with pd.ExcelWriter(nome_arquivo, engine='openpyxl') as writer:
#         for nome_mercado, data in mercados_data.items():
#             df_mercado = data['df'].copy()
#             df_mercado.reset_index(inplace=True)
#             df_mercado['Frete'] = data['frete']
#             df_mercado['Custo Total'] = df_mercado.iloc[:, 1] + df_mercado['Frete']
            
#             # Adiciona a coluna de economia ao DataFrame do mercado
#             df_mercado['Economia (vs 2º Melhor)'] = 0
#             for produto in df_mercado['Produto']:
#                 melhor_mercado = analise_final_df[analise_final_df['Produto'] == produto]['Supermercado Recomendado'].iloc[0]
#                 if nome_mercado == melhor_mercado:
#                     df_mercado.loc[df_mercado['Produto'] == produto, 'Economia (vs 2º Melhor)'] = economias.get(produto, 0)
            
#             # Renomeia as colunas para melhor visualização
#             df_mercado.rename(columns={df_mercado.columns[1]: 'Preço Unitário'}, inplace=True)
            
#             # Escreve a aba no arquivo Excel
#             df_mercado.to_excel(writer, sheet_name=nome_mercado, index=False)
#             print(f"Aba '{nome_mercado}' gerada com sucesso.")

# # --- Exemplo de uso com seus arquivos ---
# # (O conteúdo dos arquivos é o mesmo do exemplo anterior)
# arquivos_do_excel = {
#     "TOP 1 TENDAATACADO": """Produto,TENDAATACADO
# Arroz 5kg,17.4
# Feijão 1kg,3.55
# Macarrão 500g,2.59
# Óleo 900ml,6.19
# Açúcar 5kg,18.29
# Leite Integral 1L,5.45
# Pão de Forma 500g,5.59
# Café 500g,21.99
# Detergente 500ml,1.8
# Sabão em Pó 1kg,9.25
# Papel Higiênico,8.2
# Creme Dental 70g,2.6
# Água Sanitária,3.45
# Sabonete,1.05
# Fio Dental,5.15
# Molho de Tomate,1.25
# Azeite 500ml,32.89
# Farinha de Trigo 1kg,2.99
# Queijo 200g,9.49
# Creme de Leite 200g,2.25
# Frete,14.9
# Valor Mínimo,0
# Valor Total,176.32
# Total Baratos + Frete,118.94
# """,
#     "TOP 2 BOASUPERMERCADO": """Produto,BOASUPERMERCADO
# Arroz 5kg,20.9
# Feijão 1kg,4.69
# Macarrão 500g,3.49
# Óleo 900ml,7.89
# Açúcar 5kg,22.79
# Leite Integral 1L,5.19
# Pão de Forma 500g,6.59
# Café 500g,28.79
# Detergente 500ml,1.85
# Sabão em Pó 1kg,4.19
# Papel Higiênico,4.09
# Creme Dental 70g,1.55
# Água Sanitária,4.99
# Sabonete,1.55
# Fio Dental,5.99
# Molho de Tomate,1.79
# Azeite 500ml,39.9
# Farinha de Trigo 1kg,3.49
# Queijo 200g,10.99
# Creme de Leite 200g,3.45
# Frete,15
# Valor Mínimo,0
# Valor Total,199.16
# Total Baratos + Frete,24.83
# """,
#     "TOP 3 TAUSTE": """Produto,TAUSTE
# Arroz 5kg,19.89
# Feijão 1kg,5.77
# Macarrão 500g,3.98
# Óleo 900ml,6.97
# Açúcar 5kg,16.89
# Leite Integral 1L,4.59
# Pão de Forma 500g,6.86
# Café 500g,27.69
# Detergente 500ml,1.79
# Sabão em Pó 1kg,4.59
# Papel Higiênico,6.37
# Creme Dental 70g,1.79
# Água Sanitária,4.59
# Sabonete,1.19
# Fio Dental,7.69
# Molho de Tomate,1.78
# Azeite 500ml,38.89
# Farinha de Trigo 1kg,3.29
# Queijo 200g,8.77
# Creme de Leite 200g,2.98
# Frete,14.9
# Valor Mínimo,0
# Valor Total,191.26
# Total Baratos + Frete,30.05
# """,
#     "TOP 4 CONFIANCA": """Produto,CONFIANCA
# Arroz 5kg,19.89
# Feijão 1kg,5.77
# Macarrão 500g,2.98
# Óleo 900ml,6.98
# Açúcar 5kg,16.79
# Leite Integral 1L,5.19
# Pão de Forma 500g,6.7
# Café 500g,27.59
# Detergente 500ml,1.98
# Sabão em Pó 1kg,9.78
# Papel Higiênico,4.79
# Creme Dental 70g,1.79
# Água Sanitária,4.79
# Sabonete,1.77
# Fio Dental,9.68
# Molho de Tomate,1.15
# Azeite 500ml,38.9
# Farinha de Trigo 1kg,3.39
# Queijo 200g,8.98
# Creme de Leite 200g,2.79
# Frete,18.9
# Valor Mínimo,0
# Valor Total,200.58
# Total Baratos + Frete,36.84
# """,
#     "TOP 5 BARBOSA": """Produto,BARBOSA
# Arroz 5kg,19.99
# Feijão 1kg,5.99
# Macarrão 500g,2.79
# Óleo 900ml,7.49
# Açúcar 5kg,24.95
# Leite Integral 1L,5.49
# Pão de Forma 500g,6.49
# Café 500g,27.99
# Detergente 500ml,2.19
# Sabão em Pó 1kg,7.99
# Papel Higiênico,7.99
# Creme Dental 70g,2.99
# Água Sanitária,4.49
# Sabonete,0
# Fio Dental,8.99
# Molho de Tomate,1.49
# Azeite 500ml,37.99
# Farinha de Trigo 1kg,3.99
# Queijo 200g,8.99
# Creme de Leite 200g,3.49
# Frete,20.9
# Valor Mínimo,100
# Valor Total,212.67
# Total Baratos + Frete,20.9
# """,
#     "TOP 6 COOPSUPERMERCADO": """Produto,COOPSUPERMERCADO
# Arroz 5kg,25.99
# Feijão 1kg,5.99
# Macarrão 500g,3.59
# Óleo 900ml,7.49
# Açúcar 5kg,24.99
# Leite Integral 1L,5.99
# Pão de Forma 500g,6.79
# Café 500g,23.79
# Detergente 500ml,1.79
# Sabão em Pó 1kg,19.29
# Papel Higiênico,9.29
# Creme Dental 70g,3.49
# Água Sanitária,4.99
# Sabonete,1.59
# Fio Dental,15.19
# Molho de Tomate,1.69
# Azeite 500ml,42.49
# Farinha de Trigo 1kg,4.29
# Queijo 200g,11.99
# Creme de Leite 200g,3.39
# Frete,15
# Valor Mínimo,0
# Valor Total,239.1
# Total Baratos + Frete,16.79
# """
# }

# # Executa a análise para obter o resultado final e os dados brutos
# analise_completa, mercados_dados = realizar_analise_reversa_acumulativa(arquivos_do_excel)

# # Gera o novo arquivo Excel
# gerar_excel_com_economia(analise_completa, mercados_dados)

# import pandas as pd
# import io

# def realizar_analise_reversa_acumulativa(arquivos_csv):
#     """
#     Realiza a análise reversa de preços de produtos, acumulando os itens mais baratos
#     em cada mercado e eliminando os mercados não vantajosos.

#     Args:
#         arquivos_csv (dict): Um dicionário onde a chave é o nome do arquivo
#                               e o valor é o conteúdo do arquivo como string.

#     Returns:
#         tuple: Uma tupla contendo o DataFrame com a análise final e o
#                dicionário com os dados brutos de cada mercado.
#     """
    
#     mercados_data = {}
#     produtos_alocados = set()
#     resultados_finais = []

#     # 1. Processa e carrega os dados de todos os mercados
#     for nome_arquivo, conteudo_csv in arquivos_csv.items():
#         df = pd.read_csv(io.StringIO(conteudo_csv), sep=',')
#         nome_mercado = df.columns[1]
#         df.set_index('Produto', inplace=True)
        
#         frete = df.loc['Frete'].iloc[0]
#         df_produtos = df.drop(['Frete', 'Valor Mínimo', 'Valor Total', 'Total Baratos + Frete'], errors='ignore')
        
#         mercados_data[nome_mercado] = {
#             'df': df_produtos,
#             'frete': frete,
#         }
        
#     mercados_ordenados_nomes = sorted(
#         mercados_data.keys(),
#         key=lambda x: int(x.split()[1]) if 'TOP' in x else float('inf'),
#         reverse=True
#     )
    
#     # 2. Realiza a análise acumulativa
#     for i in range(len(mercados_ordenados_nomes) - 1):
#         mercado_atual_nome = mercados_ordenados_nomes[i]
#         mercado_anterior_nome = mercados_ordenados_nomes[i+1]
        
#         data_atual = mercados_data[mercado_atual_nome]
#         data_anterior = mercados_data[mercado_anterior_nome]
        
#         produtos_para_analisar = data_atual['df'].index.difference(produtos_alocados)
        
#         if produtos_para_analisar.empty:
#             continue
            
#         frete_atual = data_atual['frete']
#         frete_anterior = data_anterior['frete']
        
#         produtos_economicos_no_atual = []
        
#         for produto in produtos_para_analisar:
#             try:
#                 preco_atual = data_atual['df'].loc[produto].iloc[0]
#                 preco_anterior = data_anterior['df'].loc[produto].iloc[0]
                
#                 custo_total_atual = preco_atual + frete_atual
#                 custo_total_anterior = preco_anterior + frete_anterior
                
#                 if custo_total_atual < custo_total_anterior:
#                     produtos_economicos_no_atual.append(produto)
#             except KeyError:
#                 pass
                
#         if produtos_economicos_no_atual:
#             print(f"Análise: {mercado_atual_nome} vs {mercado_anterior_nome}")
#             print(f"Produtos mais baratos em {mercado_atual_nome}: {', '.join(produtos_economicos_no_atual)}")
#             print("-" * 50)
            
#             for produto in produtos_economicos_no_atual:
#                 preco = data_atual['df'].loc[produto].iloc[0]
#                 custo_total = preco + frete_atual
#                 resultados_finais.append({
#                     'Produto': produto,
#                     'Supermercado Recomendado': mercado_atual_nome,
#                     'Preço do Produto': preco,
#                     'Frete': frete_atual,
#                     'Custo Total': custo_total,
#                 })
            
#             produtos_alocados.update(produtos_economicos_no_atual)
            
#     # 3. Trata o último mercado (TOP 1) para os produtos restantes
#     mercado_top1_nome = mercados_ordenados_nomes[-1]
#     data_top1 = mercados_data[mercado_top1_nome]
#     frete_top1 = data_top1['frete']
#     produtos_restantes = data_top1['df'].index.difference(produtos_alocados)
    
#     if not produtos_restantes.empty:
#         print(f"Análise: Produtos restantes no {mercado_top1_nome}")
#         print(f"Produtos restantes: {', '.join(produtos_restantes)}")
#         print("-" * 50)
#         for produto in produtos_restantes:
#             preco = data_top1['df'].loc[produto].iloc[0]
#             custo_total = preco + frete_top1
#             resultados_finais.append({
#                 'Produto': produto,
#                 'Supermercado Recomendado': mercado_top1_nome,
#                 'Preço do Produto': preco,
#                 'Frete': frete_top1,
#                 'Custo Total': custo_total,
#             })
            
#     return pd.DataFrame(resultados_finais).sort_values(by='Produto'), mercados_data


# def gerar_excel_final(analise_final_df, mercados_data, nome_arquivo='analise_supermercados_final.xlsx'):
#     """
#     Gera um arquivo Excel com abas para cada supermercado recomendado, mantendo
#     a estrutura inicial e adicionando a coluna de economia.
#     """
#     # 1. Cria uma tabela consolidada de custos totais para todos os produtos em todos os mercados
#     custos_totais = pd.DataFrame()
#     for nome_mercado, data in mercados_data.items():
#         df_temp = data['df'].copy()
#         df_temp['Custo Total'] = df_temp.iloc[:, 0] + data['frete']
#         custos_totais[nome_mercado] = df_temp['Custo Total']
    
#     # 2. Identifica o melhor e o segundo melhor custo total para cada produto
#     economias = {}
#     for produto in custos_totais.index:
#         custos_ordenados = custos_totais.loc[produto].sort_values()
#         custo_melhor = custos_ordenados.iloc[0]
#         try:
#             custo_segundo_melhor = custos_ordenados.iloc[1]
#             economia_calculada = custo_segundo_melhor - custo_melhor
#             economias[produto] = economia_calculada
#         except IndexError:
#             economias[produto] = 0

#     # 3. Identifica os mercados que tiveram pelo menos um produto recomendado
#     mercados_recomendados = analise_final_df['Supermercado Recomendado'].unique()

#     # 4. Gera o arquivo Excel
#     with pd.ExcelWriter(nome_arquivo, engine='openpyxl') as writer:
#         for nome_mercado in mercados_recomendados:
#             data = mercados_data[nome_mercado]
#             df_mercado = data['df'].copy()
#             frete = data['frete']
            
#             # Adiciona as colunas de Custo Total e Economia
#             df_mercado['Custo Total (c/ Frete)'] = df_mercado.iloc[:, 0] + frete
#             df_mercado['Economia (vs 2º Melhor)'] = 0.00
            
#             # Preenche a coluna de economia apenas para os produtos recomendados para este mercado
#             for _, row in analise_final_df[analise_final_df['Supermercado Recomendado'] == nome_mercado].iterrows():
#                 produto = row['Produto']
#                 df_mercado.loc[produto, 'Economia (vs 2º Melhor)'] = economias.get(produto, 0)

#             # Adiciona a linha do frete ao final
#             df_frete = pd.DataFrame({'Frete': frete, 'Custo Total (c/ Frete)': ''}, index=['Frete'])
#             df_frete.index.name = 'Produto'
#             df_final = pd.concat([df_mercado.reset_index().set_index('Produto'), df_frete])

#             # Renomeia a coluna de preço unitário para ser mais genérica
#             df_final.rename(columns={df_final.columns[0]: 'Preço Unitário'}, inplace=True)
            
#             # Escreve a aba no arquivo Excel, sem o índice
#             df_final.to_excel(writer, sheet_name=nome_mercado, index=True)
#             print(f"Aba '{nome_mercado}' gerada com sucesso.")

# # --- Exemplo de uso com seus arquivos ---
# # (O conteúdo dos arquivos é o mesmo do exemplo anterior)
# arquivos_do_excel = {
#     "TOP 1 TENDAATACADO": """Produto,TENDAATACADO
# Arroz 5kg,17.4
# Feijão 1kg,3.55
# Macarrão 500g,2.59
# Óleo 900ml,6.19
# Açúcar 5kg,18.29
# Leite Integral 1L,5.45
# Pão de Forma 500g,5.59
# Café 500g,21.99
# Detergente 500ml,1.8
# Sabão em Pó 1kg,9.25
# Papel Higiênico,8.2
# Creme Dental 70g,2.6
# Água Sanitária,3.45
# Sabonete,1.05
# Fio Dental,5.15
# Molho de Tomate,1.25
# Azeite 500ml,32.89
# Farinha de Trigo 1kg,2.99
# Queijo 200g,9.49
# Creme de Leite 200g,2.25
# Frete,14.9
# Valor Mínimo,0
# Valor Total,176.32
# Total Baratos + Frete,118.94
# """,
#     "TOP 2 BOASUPERMERCADO": """Produto,BOASUPERMERCADO
# Arroz 5kg,20.9
# Feijão 1kg,4.69
# Macarrão 500g,3.49
# Óleo 900ml,7.89
# Açúcar 5kg,22.79
# Leite Integral 1L,5.19
# Pão de Forma 500g,6.59
# Café 500g,28.79
# Detergente 500ml,1.85
# Sabão em Pó 1kg,4.19
# Papel Higiênico,4.09
# Creme Dental 70g,1.55
# Água Sanitária,4.99
# Sabonete,1.55
# Fio Dental,5.99
# Molho de Tomate,1.79
# Azeite 500ml,39.9
# Farinha de Trigo 1kg,3.49
# Queijo 200g,10.99
# Creme de Leite 200g,3.45
# Frete,15
# Valor Mínimo,0
# Valor Total,199.16
# Total Baratos + Frete,24.83
# """,
#     "TOP 3 TAUSTE": """Produto,TAUSTE
# Arroz 5kg,19.89
# Feijão 1kg,5.77
# Macarrão 500g,3.98
# Óleo 900ml,6.97
# Açúcar 5kg,16.89
# Leite Integral 1L,4.59
# Pão de Forma 500g,6.86
# Café 500g,27.69
# Detergente 500ml,1.79
# Sabão em Pó 1kg,4.59
# Papel Higiênico,6.37
# Creme Dental 70g,1.79
# Água Sanitária,4.59
# Sabonete,1.19
# Fio Dental,7.69
# Molho de Tomate,1.78
# Azeite 500ml,38.89
# Farinha de Trigo 1kg,3.29
# Queijo 200g,8.77
# Creme de Leite 200g,2.98
# Frete,14.9
# Valor Mínimo,0
# Valor Total,191.26
# Total Baratos + Frete,30.05
# """,
#     "TOP 4 CONFIANCA": """Produto,CONFIANCA
# Arroz 5kg,19.89
# Feijão 1kg,5.77
# Macarrão 500g,2.98
# Óleo 900ml,6.98
# Açúcar 5kg,16.79
# Leite Integral 1L,5.19
# Pão de Forma 500g,6.7
# Café 500g,27.59
# Detergente 500ml,1.98
# Sabão em Pó 1kg,9.78
# Papel Higiênico,4.79
# Creme Dental 70g,1.79
# Água Sanitária,4.79
# Sabonete,1.77
# Fio Dental,9.68
# Molho de Tomate,1.15
# Azeite 500ml,38.9
# Farinha de Trigo 1kg,3.39
# Queijo 200g,8.98
# Creme de Leite 200g,2.79
# Frete,18.9
# Valor Mínimo,0
# Valor Total,200.58
# Total Baratos + Frete,36.84
# """,
#     "TOP 5 BARBOSA": """Produto,BARBOSA
# Arroz 5kg,19.99
# Feijão 1kg,5.99
# Macarrão 500g,2.79
# Óleo 900ml,7.49
# Açúcar 5kg,24.95
# Leite Integral 1L,5.49
# Pão de Forma 500g,6.49
# Café 500g,27.99
# Detergente 500ml,2.19
# Sabão em Pó 1kg,7.99
# Papel Higiênico,7.99
# Creme Dental 70g,2.99
# Água Sanitária,4.49
# Sabonete,0
# Fio Dental,8.99
# Molho de Tomate,1.49
# Azeite 500ml,37.99
# Farinha de Trigo 1kg,3.99
# Queijo 200g,8.99
# Creme de Leite 200g,3.49
# Frete,20.9
# Valor Mínimo,100
# Valor Total,212.67
# Total Baratos + Frete,20.9
# """,
#     "TOP 6 COOPSUPERMERCADO": """Produto,COOPSUPERMERCADO
# Arroz 5kg,25.99
# Feijão 1kg,5.99
# Macarrão 500g,3.59
# Óleo 900ml,7.49
# Açúcar 5kg,24.99
# Leite Integral 1L,5.99
# Pão de Forma 500g,6.79
# Café 500g,23.79
# Detergente 500ml,1.79
# Sabão em Pó 1kg,19.29
# Papel Higiênico,9.29
# Creme Dental 70g,3.49
# Água Sanitária,4.99
# Sabonete,1.59
# Fio Dental,15.19
# Molho de Tomate,1.69
# Azeite 500ml,42.49
# Farinha de Trigo 1kg,4.29
# Queijo 200g,11.99
# Creme de Leite 200g,3.39
# Frete,15
# Valor Mínimo,0
# Valor Total,239.1
# Total Baratos + Frete,16.79
# """
# }

# # 1. Executa a análise para obter o resultado final e os dados brutos
# analise_completa, mercados_dados = realizar_analise_reversa_acumulativa(arquivos_do_excel)

# # 2. Gera o novo arquivo Excel
# gerar_excel_final(analise_completa, mercados_dados)

# import pandas as pd
# import io

# def realizar_analise_reversa_acumulativa(arquivos_csv):
#     """
#     Realiza a análise reversa de preços de produtos, acumulando os itens mais baratos
#     em cada mercado e eliminando os mercados não vantajosos.

#     Args:
#         arquivos_csv (dict): Um dicionário onde a chave é o nome do arquivo
#                               e o valor é o conteúdo do arquivo como string.

#     Returns:
#         tuple: Uma tupla contendo o DataFrame com a análise final e o
#                dicionário com os dados brutos de cada mercado.
#     """
    
#     mercados_data = {}
#     produtos_alocados = set()
#     resultados_finais = []

#     # 1. Processa e carrega os dados de todos os mercados
#     for nome_arquivo, conteudo_csv in arquivos_csv.items():
#         df = pd.read_csv(io.StringIO(conteudo_csv), sep=',')
#         nome_mercado = df.columns[1]
#         df.set_index('Produto', inplace=True)
        
#         frete = df.loc['Frete'].iloc[0]
#         df_produtos = df.drop(['Frete', 'Valor Mínimo', 'Valor Total', 'Total Baratos + Frete'], errors='ignore')
        
#         mercados_data[nome_mercado] = {
#             'df': df_produtos,
#             'frete': frete,
#         }
        
#     mercados_ordenados_nomes = sorted(
#         mercados_data.keys(),
#         key=lambda x: int(x.split()[1]) if 'TOP' in x else float('inf'),
#         reverse=True
#     )
    
#     # 2. Realiza a análise acumulativa
#     for i in range(len(mercados_ordenados_nomes) - 1):
#         mercado_atual_nome = mercados_ordenados_nomes[i]
#         mercado_anterior_nome = mercados_ordenados_nomes[i+1]
        
#         data_atual = mercados_data[mercado_atual_nome]
#         data_anterior = mercados_data[mercado_anterior_nome]
        
#         produtos_para_analisar = data_atual['df'].index.difference(produtos_alocados)
        
#         if produtos_para_analisar.empty:
#             continue
            
#         frete_atual = data_atual['frete']
#         frete_anterior = data_anterior['frete']
        
#         produtos_economicos_no_atual = []
        
#         for produto in produtos_para_analisar:
#             try:
#                 preco_atual = data_atual['df'].loc[produto].iloc[0]
#                 preco_anterior = data_anterior['df'].loc[produto].iloc[0]
                
#                 custo_total_atual = preco_atual + frete_atual
#                 custo_total_anterior = preco_anterior + frete_anterior
                
#                 if custo_total_atual < custo_total_anterior:
#                     produtos_economicos_no_atual.append(produto)
#             except KeyError:
#                 pass
                
#         if produtos_economicos_no_atual:
#             print(f"Análise: {mercado_atual_nome} vs {mercado_anterior_nome}")
#             print(f"Produtos mais baratos em {mercado_atual_nome}: {', '.join(produtos_economicos_no_atual)}")
#             print("-" * 50)
            
#             for produto in produtos_economicos_no_atual:
#                 preco = data_atual['df'].loc[produto].iloc[0]
#                 custo_total = preco + frete_atual
#                 resultados_finais.append({
#                     'Produto': produto,
#                     'Supermercado Recomendado': mercado_atual_nome,
#                     'Preço do Produto': preco,
#                     'Frete': frete_atual,
#                     'Custo Total': custo_total,
#                 })
            
#             produtos_alocados.update(produtos_economicos_no_atual)
            
#     # 3. Trata o último mercado (TOP 1) para os produtos restantes
#     mercado_top1_nome = mercados_ordenados_nomes[-1]
#     data_top1 = mercados_data[mercado_top1_nome]
#     frete_top1 = data_top1['frete']
#     produtos_restantes = data_top1['df'].index.difference(produtos_alocados)
    
#     if not produtos_restantes.empty:
#         print(f"Análise: Produtos restantes no {mercado_top1_nome}")
#         print(f"Produtos restantes: {', '.join(produtos_restantes)}")
#         print("-" * 50)
#         for produto in produtos_restantes:
#             preco = data_top1['df'].loc[produto].iloc[0]
#             custo_total = preco + frete_top1
#             resultados_finais.append({
#                 'Produto': produto,
#                 'Supermercado Recomendado': mercado_top1_nome,
#                 'Preço do Produto': preco,
#                 'Frete': frete_top1,
#                 'Custo Total': custo_total,
#             })
            
#     return pd.DataFrame(resultados_finais).sort_values(by='Produto'), mercados_data


# def gerar_excel_final(analise_final_df, mercados_data, nome_arquivo='analise_supermercados_final.xlsx'):
#     """
#     Gera um arquivo Excel com abas para cada supermercado recomendado, mantendo
#     a estrutura inicial e adicionando a coluna de economia e o valor total economizado.
#     """
#     # 1. Cria uma tabela consolidada de custos totais
#     custos_totais = pd.DataFrame()
#     for nome_mercado, data in mercados_data.items():
#         df_temp = data['df'].copy()
#         df_temp['Custo Total'] = df_temp.iloc[:, 0] + data['frete']
#         custos_totais[nome_mercado] = df_temp['Custo Total']
    
#     # 2. Identifica a economia de cada produto
#     economias = {}
#     for produto in custos_totais.index:
#         custos_ordenados = custos_totais.loc[produto].sort_values()
#         custo_melhor = custos_ordenados.iloc[0]
#         try:
#             custo_segundo_melhor = custos_ordenados.iloc[1]
#             economia_calculada = custo_segundo_melhor - custo_melhor
#             economias[produto] = economia_calculada
#         except IndexError:
#             economias[produto] = 0

#     # 3. Identifica os mercados que tiveram pelo menos um produto recomendado
#     mercados_recomendados = analise_final_df['Supermercado Recomendado'].unique()

#     # 4. Gera o arquivo Excel
#     with pd.ExcelWriter(nome_arquivo, engine='openpyxl') as writer:
#         for nome_mercado in mercados_recomendados:
#             data = mercados_data[nome_mercado]
#             df_mercado = data['df'].copy()
#             frete = data['frete']
            
#             # Adiciona as colunas de Custo Total e Economia
#             df_mercado['Custo Total (c/ Frete)'] = df_mercado.iloc[:, 0] + frete
#             df_mercado['Economia (vs 2º Melhor)'] = 0.00
            
#             # Preenche a coluna de economia apenas para os produtos recomendados
#             for _, row in analise_final_df[analise_final_df['Supermercado Recomendado'] == nome_mercado].iterrows():
#                 produto = row['Produto']
#                 df_mercado.loc[produto, 'Economia (vs 2º Melhor)'] = economias.get(produto, 0)
            
#             # Calcula o total economizado para esta aba
#             total_economizado = df_mercado['Economia (vs 2º Melhor)'].sum()

#             # Adiciona a linha do frete ao final
#             df_frete = pd.DataFrame({'Frete': frete, 'Custo Total (c/ Frete)': ''}, index=['Frete'])
#             df_frete.index.name = 'Produto'
#             df_final = pd.concat([df_mercado.reset_index().set_index('Produto'), df_frete])

#             # Adiciona a linha com o total economizado
#             df_total = pd.DataFrame({'Economia (vs 2º Melhor)': total_economizado}, index=['Total Economizado'])
#             df_total.index.name = 'Produto'
#             df_final = pd.concat([df_final, df_total])

#             # Renomeia as colunas para melhor visualização
#             df_final.rename(columns={df_final.columns[0]: 'Preço Unitário'}, inplace=True)
            
#             # Escreve a aba no arquivo Excel, sem o índice
#             df_final.to_excel(writer, sheet_name=nome_mercado, index=True)
#             print(f"Aba '{nome_mercado}' gerada com sucesso. Total economizado: R$ {total_economizado:.2f}")

# # --- Exemplo de uso com seus arquivos ---
# arquivos_do_excel = {
#     "TOP 1 TENDAATACADO": """Produto,TENDAATACADO
# Arroz 5kg,17.4
# Feijão 1kg,3.55
# Macarrão 500g,2.59
# Óleo 900ml,6.19
# Açúcar 5kg,18.29
# Leite Integral 1L,5.45
# Pão de Forma 500g,5.59
# Café 500g,21.99
# Detergente 500ml,1.8
# Sabão em Pó 1kg,9.25
# Papel Higiênico,8.2
# Creme Dental 70g,2.6
# Água Sanitária,3.45
# Sabonete,1.05
# Fio Dental,5.15
# Molho de Tomate,1.25
# Azeite 500ml,32.89
# Farinha de Trigo 1kg,2.99
# Queijo 200g,9.49
# Creme de Leite 200g,2.25
# Frete,14.9
# Valor Mínimo,0
# Valor Total,176.32
# Total Baratos + Frete,118.94
# """,
#     "TOP 2 BOASUPERMERCADO": """Produto,BOASUPERMERCADO
# Arroz 5kg,20.9
# Feijão 1kg,4.69
# Macarrão 500g,3.49
# Óleo 900ml,7.89
# Açúcar 5kg,22.79
# Leite Integral 1L,5.19
# Pão de Forma 500g,6.59
# Café 500g,28.79
# Detergente 500ml,1.85
# Sabão em Pó 1kg,4.19
# Papel Higiênico,4.09
# Creme Dental 70g,1.55
# Água Sanitária,4.99
# Sabonete,1.55
# Fio Dental,5.99
# Molho de Tomate,1.79
# Azeite 500ml,39.9
# Farinha de Trigo 1kg,3.49
# Queijo 200g,10.99
# Creme de Leite 200g,3.45
# Frete,15
# Valor Mínimo,0
# Valor Total,199.16
# Total Baratos + Frete,24.83
# """,
#     "TOP 3 TAUSTE": """Produto,TAUSTE
# Arroz 5kg,19.89
# Feijão 1kg,5.77
# Macarrão 500g,3.98
# Óleo 900ml,6.97
# Açúcar 5kg,16.89
# Leite Integral 1L,4.59
# Pão de Forma 500g,6.86
# Café 500g,27.69
# Detergente 500ml,1.79
# Sabão em Pó 1kg,4.59
# Papel Higiênico,6.37
# Creme Dental 70g,1.79
# Água Sanitária,4.59
# Sabonete,1.19
# Fio Dental,7.69
# Molho de Tomate,1.78
# Azeite 500ml,38.89
# Farinha de Trigo 1kg,3.29
# Queijo 200g,8.77
# Creme de Leite 200g,2.98
# Frete,14.9
# Valor Mínimo,0
# Valor Total,191.26
# Total Baratos + Frete,30.05
# """,
#     "TOP 4 CONFIANCA": """Produto,CONFIANCA
# Arroz 5kg,19.89
# Feijão 1kg,5.77
# Macarrão 500g,2.98
# Óleo 900ml,6.98
# Açúcar 5kg,16.79
# Leite Integral 1L,5.19
# Pão de Forma 500g,6.7
# Café 500g,27.59
# Detergente 500ml,1.98
# Sabão em Pó 1kg,9.78
# Papel Higiênico,4.79
# Creme Dental 70g,1.79
# Água Sanitária,4.79
# Sabonete,1.77
# Fio Dental,9.68
# Molho de Tomate,1.15
# Azeite 500ml,38.9
# Farinha de Trigo 1kg,3.39
# Queijo 200g,8.98
# Creme de Leite 200g,2.79
# Frete,18.9
# Valor Mínimo,0
# Valor Total,200.58
# Total Baratos + Frete,36.84
# """,
#     "TOP 5 BARBOSA": """Produto,BARBOSA
# Arroz 5kg,19.99
# Feijão 1kg,5.99
# Macarrão 500g,2.79
# Óleo 900ml,7.49
# Açúcar 5kg,24.95
# Leite Integral 1L,5.49
# Pão de Forma 500g,6.49
# Café 500g,27.99
# Detergente 500ml,2.19
# Sabão em Pó 1kg,7.99
# Papel Higiênico,7.99
# Creme Dental 70g,2.99
# Água Sanitária,4.49
# Sabonete,0
# Fio Dental,8.99
# Molho de Tomate,1.49
# Azeite 500ml,37.99
# Farinha de Trigo 1kg,3.99
# Queijo 200g,8.99
# Creme de Leite 200g,3.49
# Frete,20.9
# Valor Mínimo,100
# Valor Total,212.67
# Total Baratos + Frete,20.9
# """,
#     "TOP 6 COOPSUPERMERCADO": """Produto,COOPSUPERMERCADO
# Arroz 5kg,25.99
# Feijão 1kg,5.99
# Macarrão 500g,3.59
# Óleo 900ml,7.49
# Açúcar 5kg,24.99
# Leite Integral 1L,5.99
# Pão de Forma 500g,6.79
# Café 500g,23.79
# Detergente 500ml,1.79
# Sabão em Pó 1kg,19.29
# Papel Higiênico,9.29
# Creme Dental 70g,3.49
# Água Sanitária,4.99
# Sabonete,1.59
# Fio Dental,15.19
# Molho de Tomate,1.69
# Azeite 500ml,42.49
# Farinha de Trigo 1kg,4.29
# Queijo 200g,11.99
# Creme de Leite 200g,3.39
# Frete,15
# Valor Mínimo,0
# Valor Total,239.1
# Total Baratos + Frete,16.79
# """
# }

# # 1. Executa a análise para obter o resultado final e os dados brutos
# analise_completa, mercados_dados = realizar_analise_reversa_acumulativa(arquivos_do_excel)

# # 2. Gera o novo arquivo Excel
# gerar_excel_final(analise_completa, mercados_dados)

import pandas as pd
import io

def realizar_analise_reversa_acumulativa(arquivos_csv):
    """
    Realiza a análise de preços de produtos, encontrando o item mais barato
    em cada mercado para a cesta de compras.

    Args:
        arquivos_csv (dict): Um dicionário onde a chave é o nome do arquivo
                              e o valor é o conteúdo do arquivo como string.

    Returns:
        tuple: Uma tupla contendo o DataFrame com a análise final e o
               dicionário com os dados brutos de cada mercado.
    """
    
    mercados_data = {}
    
    # 1. Processa e carrega os dados de todos os mercados
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
    
    # 2. Cria uma tabela consolidada de custos totais para a análise
    custos_totais = pd.DataFrame()
    for nome_arquivo_completo, data in mercados_data.items():
        df_temp = data['df'].copy()
        df_temp['Custo Total'] = df_temp.iloc[:, 0] + data['frete']
        custos_totais[nome_arquivo_completo] = df_temp['Custo Total']
    
    # 3. Identifica o supermercado mais barato para cada produto
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


def gerar_excel_final(analise_final_df, mercados_data, nome_arquivo='analise_supermercados_final.xlsx'):
    """
    Gera um arquivo Excel com abas para cada supermercado recomendado, mantendo
    a estrutura inicial e adicionando a coluna de economia e o valor total economizado.
    """
    # 1. Cria uma tabela consolidada de custos totais
    custos_totais = pd.DataFrame()
    for nome_arquivo_completo, data in mercados_data.items():
        nome_mercado_curto = data['nome_curto']
        df_temp = data['df'].copy()
        df_temp['Custo Total'] = df_temp.iloc[:, 0] + data['frete']
        custos_totais[nome_mercado_curto] = df_temp['Custo Total']
    
    # 2. Identifica a economia de cada produto
    economias = {}
    for produto in custos_totais.index:
        custos_ordenados = custos_totais.loc[produto].sort_values()
        custo_melhor = custos_ordenados.iloc[0]
        try:
            custo_segundo_melhor = custos_ordenados.iloc[1]
            economia_calculada = custo_segundo_melhor - custo_melhor
            economias[produto] = economia_calculada
        except IndexError:
            economias[produto] = 0

    # 3. Identifica os mercados que tiveram pelo menos um produto recomendado
    mercados_recomendados_keys = analise_final_df['Supermercado Recomendado'].unique()
    
    # Ordena os mercados recomendados pelo TOP X para garantir a ordem correta
    mercados_recomendados_ordenados = sorted(
        mercados_recomendados_keys,
        key=lambda x: int(x.split(' ')[1]) if 'TOP' in x else float('inf')
    )
    
    # 4. Gera o arquivo Excel
    with pd.ExcelWriter(nome_arquivo, engine='openpyxl') as writer:
        for nome_mercado_completo in mercados_recomendados_ordenados:
            data = mercados_data[nome_mercado_completo]
            nome_mercado_curto = data['nome_curto']
            df_mercado = data['df'].copy()
            frete = data['frete']
            
            # Adiciona as colunas de Custo Total e Economia
            df_mercado['Custo Total (c/ Frete)'] = df_mercado.iloc[:, 0] + frete
            df_mercado['Economia (vs 2º Melhor)'] = 0.00
            
            # Preenche a coluna de economia apenas para os produtos recomendados
            for _, row in analise_final_df[analise_final_df['Supermercado Recomendado'] == nome_mercado_completo].iterrows():
                produto = row['Produto']
                # Garante que o cálculo da economia está sendo feito em relação ao mercado correto (o que tem o melhor preço para o produto)
                df_mercado.loc[produto, 'Economia (vs 2º Melhor)'] = economias.get(produto, 0)
            
            # Calcula o total economizado para esta aba
            total_economizado = df_mercado['Economia (vs 2º Melhor)'].sum()

            # Adiciona a linha do frete ao final
            df_frete = pd.DataFrame({'Frete': frete, 'Custo Total (c/ Frete)': ''}, index=['Frete'])
            df_frete.index.name = 'Produto'
            df_final = pd.concat([df_mercado.reset_index().set_index('Produto'), df_frete])

            # Adiciona a linha com o total economizado
            df_total = pd.DataFrame({'Economia (vs 2º Melhor)': total_economizado}, index=['Total Economizado'])
            df_total.index.name = 'Produto'
            df_final = pd.concat([df_final, df_total])

            # Renomeia as colunas para melhor visualização
            df_final.rename(columns={df_final.columns[0]: 'Preço Unitário'}, inplace=True)
            
            # Define o nome da aba no formato "TOP X - NomeSupermercado"
            nome_aba = nome_mercado_completo.replace(' ', ' - ', 1)
            
            # Escreve a aba no arquivo Excel
            df_final.to_excel(writer, sheet_name=nome_aba, index=True)
            print(f"Aba '{nome_aba}' gerada com sucesso. Total economizado: R$ {total_economizado:.2f}")

# --- Exemplo de uso com seus arquivos ---
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