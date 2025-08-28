import threading
import multiprocessing
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.keys import Keys as type_key
from selenium.webdriver.chrome.options import Options
import time
import pandas as pd
# ... (outras importações)

PATH = 'C:\VIRTUALENVS\PROJETO_PRODUTOS_COMPRAS\chromedriver.exe'
CEP = '18071781'

driver_lock = threading.Lock()

class SuperMarket:
    def __init__(self, supermarkets: list[str], products: list[str]):
        # ... (código existente)
        """
        Inicializa a classe SuperMarket.
        
        :param supermarkets: Lista de supermercados a serem pesquisados.
        :param products: Lista de produtos a serem buscados.
        """
        # cService = webdriver.ChromeService(executable_path=PATH)
        # self.driver = webdriver.Chrome(service=cService)

        # Lista de URLs dos supermercados
        self.websites = ['https://tauste.com.br', 
                         'https://www.barbosasupermercados.com.br/', 
                         'https://www.confianca.com.br/marilia', 
                         'https://lista.mercadolivre.com.br', 
                         'https://www.coopsupermercado.com.br', 
                         'https://www.tendaatacado.com.br']

        self.supermarkets = supermarkets
        self.products = products
        
        #Listas para armazenar os produtos e seus respectivos preços
        self.list_arroz = []
        self.list_feijao = []
        self.list_macarrao = []
        self.list_oleo = []
        self.list_acucar = []
        self.list_leite = []
        self.list_pao = []
        self.list_cafe = []
        self.list_detergente = []
        self.list_sabao_po = []
        self.list_pagel_hig = []

    def init_driver(self) -> None:
        options = Options()
        options.add_argument("--start-maximized")
        cService = webdriver.ChromeService(executable_path=PATH)
        driver = webdriver.Chrome(service=cService, options=options)

        return driver


    def extract_data(self) -> None:
        threads = []
        for supermarket in self.supermarkets:
            thread = threading.Thread(target=self.extract_supermarket_data, args=(supermarket,))
            threads.append(thread)
            thread.start()

        # Aguarda todas as threads terminarem
        for thread in threads:
            thread.join()

        # Exportando a tabela de preços
        self.export_table()

    def extract_supermarket_data(self, supermarket: str) -> None:
        driver = self.init_driver()
        try:
            if supermarket.upper() == 'TAUSTE':
                self.extract_tauste_data(driver=driver)
            elif supermarket.upper() == 'BARBOSA':
                self.extract_barbosa_data(driver=driver)
            elif supermarket.upper() == 'CONFIANÇA':
                self.extract_confianca_data()
            elif supermarket.upper() == 'MERCADO LIVRE':
                self.extract_mercado_livre_data()
            elif supermarket.upper() == 'COOP SUPERMERCADO':
                self.extract_coop_supermercado_data()
            elif supermarket.upper() == 'TENDA ATACADO':
                self.extract_tenda_atacado_data()
        except Exception as e:
            print(f"Erro ao extrair dados do {supermarket}: {e}")

    def extract_tauste_data(self, driver):

        
        self.driver = driver
        
        with driver_lock:
            # Acessa o site do supermercado Tauste
            self.driver.get(self.websites[0])
            # self.driver.maximize_window()
            time.sleep(4)

            # Seleciona a opção de compra na loja física
            self.driver.find_element(by=By.XPATH, value='//label[@class="label-choice instore"]').click()
            time.sleep(2)

        # Seleciona a cidade da loja
        select_element = self.driver.find_element(by=By.XPATH, value='//select[@id="stores-city"]')
        select = Select(select_element)
        select.select_by_value('4')
        time.sleep(2)

        # Seleciona a loja especifica
        self.driver.find_element(by=By.XPATH, value='//*[@id="store-code-city-selected"]/div/li/div/strong[1]').click()
        time.sleep(10)

        for product in self.products:
            # Padroniza a unidade de medida dos produtos
            product = (product.replace('5kg', '5000g')
                            .replace('1kg', '1000g')
                            .replace('1L', '1000ml'))

            # Realiza a busca do produto no site do supermercado
            self.driver.find_element(by=By.XPATH, value='//input[contains(@id, "search")]').send_keys(product.replace('Sabão em ', 'Detergente '))
            self.driver.find_element(by=By.XPATH, value='//button[contains(@title, "Pesquisa")]').click()

            time.sleep(5)
            # Localiza os produtos na pagina
            contains_txt = self.set_contains(product=product.replace('Sabão em ', 'Detergente '))
            products_items = self.driver.find_elements(by=By.XPATH, value=f'//li[@class="item product product-item"][{contains_txt}]')

            # Verifica se o produto foi encontrado na pagina
            if not products_items:
                continue

            # Ordena a busca por preços crescentes
            time.sleep(7)
            select_element = self.driver.find_element(by=By.XPATH, value='//select[@id="sorter"]')
            select = Select(select_element)
            select.select_by_value('price')
            time.sleep(2)
            button = self.driver.find_element(by=By.XPATH, value='//a[contains(@class, "action sorter-action sort-desc")]')
            self.driver.execute_script("arguments[0].click();", button)
            time.sleep(5)

            # Localiza os produtos na pagina
            contains_txt = self.set_contains(product=product.replace('Sabão em ', 'Detergente '))
            products_rows = self.driver.find_elements(by=By.XPATH, value=f'//li[@class="item product product-item"][{contains_txt}]')

            # Verifica se o produto foi encontrado na pagina
            if not products_rows:
                continue

            # Usa o primeiro produto encontrado correspondente ao mais barato
            product = products_rows[0]

            # Extrai o nome e o preço do produto
            name = product.find_element(by=By.XPATH, value='.//strong[@class="product name product-item-name"]')
            print(name.text)
            price = product.find_element(by=By.XPATH, value='.//div[@class="price-box price-final_price"]//span[@class="price"]').text
            print(price)

            # Adiciona o produto e o preço na respectiva lista
            self.append_to_list(name=name, price=price, supermarket='TAUSTE')

            time.sleep(2)

    def extract_barbosa_data(self, driver):

        self.driver = driver

        with driver_lock:
            # Acessa o site do supermercado Barbosa
            website = self.websites[1]
            self.driver.get(website)

            # self.driver.maximize_window()

            time.sleep(15)
            self.driver.find_element(by=By.XPATH, value='(//div[@class="content"]//div[@class="close"])[1]').click()

        for product in self.products:
            product = product.replace('Sabão em ', 'Detergente em ')

            # Separa nome e tamanho do produto
            product_title = ' '.join(product.split(' ')[0:-1])
            product_size = product.split(' ')[-1]

            time.sleep(2)
            # Realiza a busca do produto no site do supermercado
            self.driver.get(f'{website}/buscar?texto={product_title}&ordem=Valor+crescente')

            # Localiza os produtos na página
            contains_txt = self.set_contains(product)
            time.sleep(10)
            container_products = self.driver.find_element(by=By.XPATH, value='//div[@class="items-container"]')
            products_rows = container_products.find_elements(by=By.XPATH, value=f'.//div[@class="items row"]//div[contains(@class, "row-item")][{contains_txt}]')

            without_5kg = True

            # Valida se o produto não foi encontrado e se é Açúcar Cristal
            if not products_rows and product_title == 'Açúcar Cristal':
                # Define o produto como Açúcar Cristal 1kg
                product = 'Açúcar Cristal 1kg'
                without_5kg = False

                # Realiza uma nova localização dos produtos
                contains_txt = self.set_contains(product)
                products_rows = container_products.find_elements(by=By.XPATH, value=f'.//div[@class="items row"]//div[contains(@class, "row-item")][{contains_txt}]')

            # Valida se o produto nao foi encontrado
            if not products_rows:
                continue

            # Usa o primeiro produto encontrado correspondente ao mais barato
            product = products_rows[0]

            # Extrai o nome e o preço do produto
            name = product.find_element(by=By.XPATH, value='.//div[@class="stats"]//div[@class="title"]')
            print(name.text)

            if without_5kg == False:
                price = (product.find_element(by=By.XPATH, value='.//div[@class="prices-buy"]//div[@class="prices"]//p[@class="price"]')
                                .text
                                .replace('R$ ', '')
                                .replace(',', '.'))

                price = 'R$ ' + str(round(float(price) * 5, 2))
                print(price)

            else:
                price = product.find_element(by=By.XPATH, value='.//div[@class="prices-buy"]//div[@class="prices"]//p[@class="price"]')
                price = price.text
                print(price)

            # Adiciona o produto e o preço na respectiva lista
            self.append_to_list(name, price, supermarket='BARBOSA')

        time.sleep(2)

    def extract_confianca_data(self):

        self.driver = self.init_driver()
    # Acessa o site do supermercado Confiança
        website = self.websites[2]
        self.driver.get(website)
        # self.driver.maximize_window()
        time.sleep(10)

        # Fecha o modal de boas vindas
        self.driver.find_element(by=By.XPATH, value='//div[@class="modal__close__welcome"]').click()

        # Realiza a busca dos produtos
        for product in self.products:
            # Separa nome e tamanho do produto
            final_product_name = '+'.join(product.split())

            # Realiza a busca do produto no site do supermercado
            self.driver.get(f'{website}/search?Ns=sku.activePrice%7C0&Ntt={final_product_name.lower()}')
            time.sleep(5)

            # Fecha o modal de boas vindas
            self.driver.find_element(by=By.XPATH, value='//div[@class="modal__close__welcome"]').click()

            # Localiza os produtos na página
            contains_txt = self.set_contains(product)
            time.sleep(3)
            products_list = self.driver.find_element(by=By.XPATH, value='//div[@class="product-list product-list-fill-space"]')
            products_rows = products_list.find_elements(by=By.XPATH, value=f'.//a[{contains_txt}]')

            # Valida se o produto não foi encontrado
            if not products_rows:
                continue

            # Usa o primeiro produto encontrado correspondente ao mais barato
            product = products_rows[0]

            # Extrai o nome e o preço do produto
            name = product.find_element(by=By.XPATH, value='.//article//div//h3')
            print(name.text)
            price = product.find_element(by=By.XPATH, value='.//span[@class="product-shelf__price-current"]').text
            print(price)

            # Adiciona o produto e o preço na respectiva lista
            self.append_to_list(name, price, supermarket='CONFIANÇA')

            time.sleep(2)

    def extract_mercado_livre_data(self):
        # Realiza a busca dos produtos
        for product in self.products:
            # Realiza a busca do produto no site do supermercado
            website = self.websites[3]
            product_splited = product.split()
            final_product_name = '-'.join(product_splited)

            if 'Arroz' in final_product_name:
                self.driver.get(f'{website}/alimentos-bebidas/mercearia/{final_product_name}_OrderId_PRICE_PriceRange_20-0_NoIndex_True#applied_filter_id%3Dprice%26applied_filter_name%3DPreço%26applied_filter_order%3D9%26applied_value_id%3D20-*%26applied_value_name%3D20-*%26applied_value_order%3D4%26applied_value_results%3DUNKNOWN_RESULTS%26is_custom%3Dtrue')
            else:
                self.driver.get(f'{website}/{final_product_name}_OrderId_PRICE_NoIndex_True')

            self.driver.maximize_window()

            # Localiza os produtos na página
            contains_txt = self.set_contains(product)
            time.sleep(7)
            container_products = self.driver.find_element(by=By.XPATH, value='//ol[contains(@class, "ui-search-layout ui-search-layout")]')

            if 'Detergente' in product or 'Sabão em Pó' in product or 'Papel Higiênico' in product:
                products_rows = container_products.find_elements(by=By.XPATH, value=f'.//div[contains(@class, "andes-card poly-card poly-card--grid-card andes-card--flat andes-card--padding-0 andes-card--animated")][{contains_txt}]')
            else:
                products_rows = container_products.find_elements(by=By.XPATH, value=f'.//li[@class="ui-search-layout__item shops__layout-item"][{contains_txt}]')

            # Valida se o produto não foi encontrado
            if not products_rows:
                continue

            # Usa o primeiro produto encontrado correspondente ao mais barato
            product = products_rows[0]

            # Extrai o nome e o preço do produto
            name = product.find_element(by=By.XPATH, value='.//h3')
            print(name.text)
            price = (product.find_element(by=By.XPATH, value='.//span[@class="andes-money-amount andes-money-amount--cents-superscript"]')
                            .text
                            .replace('\n', ''))
            print(price)

            # Adiciona o produto e o preço na respectiva lista
            self.append_to_list(name, price, supermarket='MERCADO LIVRE')

            time.sleep(2)

    def extract_coop_supermercado_data(self):
        # Acessa o site do supermercado Coop
        website = self.websites[4]
        self.driver.get(website)
        self.driver.maximize_window()
        time.sleep(6)

        # Define o CEP do endereço de entrega
        self.driver.find_element(by=By.XPATH, value='//input[contains(@name,"localCep")]').send_keys(CEP)
        time.sleep(3)
        self.driver.find_element(by=By.XPATH, value='//button[contains(@class, "coopsp-store-theme-YXbQXXox6msATjZ4Edfjr")]').click()
        time.sleep(10)
        self.driver.find_element(by=By.XPATH, value='(//div[contains(@class, "coopsp-store-theme-8-x-localizationDeliveryOptionsShippingValue_ctaDesc")])[2]').click()
        time.sleep(6)

        # Realiza a busca dos produtos
        for product in self.products:
            # Realiza a busca do produto no site do supermercado
            product_name = ('%20').join(product.split())
            self.driver.get(f'{website}/{product_name}?_q={product_name}&map=ft&order=OrderByPriceASC')
            time.sleep(10)

            # Localiza os produtos na página
            contains_txt = self.set_contains(product.replace(' de ', ' '))
            container = self.driver.find_element(by=By.XPATH, value='//div[contains(@class, "vtex-search-result-3-x-gallery vtex-search-result-3-x-gallery--grid flex flex-row flex-wrap items-stretch bn ph1 na4 pl9-l")]')
            products_rows = container.find_elements(by=By.XPATH, value=f'.//div[contains(@class, "vtex-search-result-3-x-galleryItem")][{contains_txt}][not(.//text()[contains(., "Indisponível")])]')

            # Valida se o produto não foi encontrado
            if not products_rows:
                continue

            # Usa o primeiro produto encontrado correspondente ao mais barato
            product = products_rows[0]

            # Extrai o nome e o preço do produto
            try:
                name = product.find_element(by=By.XPATH, value='.//div[contains(@class, "vtex-product-summary-2-x-nameContainer")]//h3')
                print(name.text)
                price = product.find_element(by=By.XPATH, value='.//span[contains(@class, "coopsp-store-theme-8-x-customProductSellingPrice_sellingPrice")]').text
                print(price)

                # Adiciona o produto e o preço na respectiva lista
                self.append_to_list(name, price, supermarket='COOP SUPERMERCADO')
            except Exception:
                # Se o produto não foi encontrado, adiciona o produto na respectiva lista
                item_name = (product_name.split()[0]
                                    .lower()
                                    .replace('ç', 'c')
                                    .replace('ú', 'u')
                                    .replace('ó', 'o')
                                    .replace('ã', 'a'))
                list = getattr(self, f'list_{item_name}', [])
                list.append({'product': item_name, 'price': None, 'supermarket': 'COOP SUPERMERCADO'})
                continue

            time.sleep(5)

    def extract_tenda_atacado_data(self):
        # Acessa o site da Tenda Atacado
        website = self.websites[5]
        self.driver.get(website)
        self.driver.maximize_window()
        time.sleep(3)

        # Define o CEP do endereço de entrega
        self.driver.find_element(by=By.XPATH, value='//input[contains(@name, "searchbarComponent")]').click()
        time.sleep(3)
        self.driver.find_element(by=By.NAME, value='zipCode').send_keys(CEP)
        time.sleep(2)
        self.driver.find_element(by=By.XPATH, value='(//img[contains(@class, "svgIcon svg-ico_close_with_circle")])[1]').click()
        time.sleep(2)

        # Realiza a busca dos produtos
        for product in self.products:
            product = product.replace('Sabão em ', 'Lava Roupas ')
            # Realiza a busca do produto no site do supermercado
            product_name = '+'.join(product.split())
            self.driver.get(f'{website}/busca?q={product_name}')
            time.sleep(5)

            # Define a ordenação dos produtos pelo preço
            select_element = self.driver.find_element(by=By.XPATH, value='//select[contains(@id, "select-sort-container")]')
            select = Select(select_element)
            select.select_by_value('ascPrice')
            time.sleep(7)

            # Localiza os produtos na página
            contains_txt = self.set_contains(product)
            container_products = self.driver.find_element(by=By.XPATH, value='//section[contains(@class, "MosaicCardContainer grid_view")]')
            products_rows = container_products.find_elements(by=By.XPATH, value=f'.//div[contains(@class, "ProductCardShowcase")][{contains_txt}]')

            # Valida se o produto não foi encontrado
            if not products_rows:
                continue

            # Usa o primeiro produto encontrado correspondente ao mais barato
            product = products_rows[0]

            # Extrai o nome e o preço do produto
            name = product.find_element(by=By.XPATH, value='.//div[contains(@class, "informations")]//h3')
            print(name.text)
            price = (product.find_element(by=By.XPATH, value='.//span[contains(@class, "showcase-price-area")]//div[contains(@class, "price-space")]')
                            .text
                            .replace(' un', ''))
            print(price)

            # Adiciona o produto e o preço na respectiva lista
            self.append_to_list(name, price, supermarket='TENDA ATACADO')

    def export_table(self) -> None:
        """
        Exporta os dados coletados em formato de tabela.
        """
        
        # Criando um dicionário para armazenar os produtos
        data = {}

        # Função auxiliar para organizar os itens corretamente
        def add_items(product_list, product_name):
            for item in product_list:
                supermarket = item["supermarket"]
                price = item["price"]

                if product_name not in data:
                    data[product_name] = {}

                data[product_name][supermarket] = float(price.replace('R$', '')
                                                             .replace(',', '.')
                                                             .replace(' ', ''))

        # Adicionando os produtos conforme as listas
        if self.list_arroz:
            add_items(self.list_arroz, 'Arroz 5kg')

        if self.list_feijao:
            add_items(self.list_feijao, 'Feijão 1kg')

        if self.list_macarrao:
            add_items(self.list_macarrao, 'Macarrão 500g')

        if self.list_oleo:
            add_items(self.list_oleo, 'Óleo 900ml')

        if self.list_acucar:
            add_items(self.list_acucar, 'Açúcar 5kg')

        if self.list_leite:
            add_items(self.list_leite, 'Leite Integral 1L')

        if self.list_pao:
            add_items(self.list_pao, 'Pão de Forma 500g')

        if self.list_cafe:
            add_items(self.list_cafe, 'Café 500g')

        if self.list_detergente:
            add_items(self.list_detergente, 'Detergente 500ml')

        if self.list_sabao_po:
            add_items(self.list_sabao_po, 'Sabão em Pó 1kg')

        if self.list_pagel_hig:
            add_items(self.list_pagel_hig, 'Papel Higiênico')

        # Criando um DataFrame
        df = pd.DataFrame.from_dict(data, orient="index").reset_index()

        # Renomeando a coluna do índice para "Item"
        df = df.rename(columns={"index": "Item"})

        # Substituindo NaN por "Não encontrado"
        df = df.fillna("Não encontrado")

        # Exportando para CSV
        df.to_csv("../precos_supermercados_7.csv", encoding="utf-8-sig", index=False)
        print("✅ Tabela exportada para 'precos_supermercados.csv'!")

        # Exibir tabela no terminal
        print(df)

        self.driver.close()

    def set_contains(self, product:str) -> str:
        """
        Cria um XPath dinâmico para buscar um produto em uma página de supermercado.

        :param product: Nome do produto a ser buscado.
        :return: XPath dinâmico para busca do produto.
        """
        # Inicializa a variável que conterá o XPath
        contains_txt = ''
        # Cria o XPath dinâmico
        for parte in product.split(' '):
                        contains_txt = contains_txt + f'contains(translate(., "ABCDEFGHIJKLMNOPQRSTUVWXYZÓ", "abcdefghijklmnopqrstuvwxyzó"), "{parte.lower()}") and '

        # Retorna o XPath
        return contains_txt[:-5]
    
    def append_to_list(self, name:str, price:str, supermarket:str):
            """
            Adiciona os produtos encontrados às listas de armazenamento correspondentes.
            
            :param name: Nome do produto encontrado.
            :param price: Preço do produto encontrado.
            :param supermarket: Nome do supermercado onde o produto foi encontrado.
            """

            print(name.text.lower())

            # Adiciona o produto e o preço na respectiva lista
            if 'arroz' in name.text.lower():
                self.list_arroz.append({'product': name.text.split(' ')[0] + ' 5kg', 'price': price, 'supermarket': supermarket})

            elif 'feijão' in name.text.lower():
                self.list_feijao.append({'product': name.text.split(' ')[0] + ' 1kg', 'price': price, 'supermarket': supermarket})

            elif 'espaguete' in name.text.lower():
                self.list_macarrao.append({'product': name.text.split(' ')[0] + ' 500g', 'price': price, 'supermarket': supermarket})

            elif 'óleo' in name.text.lower():
                self.list_oleo.append({'product': name.text.split(' ')[0] + ' 900ml', 'price': price, 'supermarket': supermarket})

            elif 'açúcar' in name.text.lower():
                self.list_acucar.append({'product': name.text.split(' ')[0] + ' 5kg', 'price': price, 'supermarket': supermarket})

            elif 'leite' in name.text.lower():
                self.list_leite.append({'product': name.text.split(' ')[0] + ' 1l', 'price': price, 'supermarket': supermarket})

            elif 'pão' in name.text.lower():
                self.list_pao.append({'product': name.text.split(' ')[0] + ' 500g', 'price': price, 'supermarket': supermarket})

            elif 'café' in name.text.lower():
                self.list_cafe.append({'product': name.text.split(' ')[0] + ' 500g', 'price': price, 'supermarket': supermarket})

            elif ('detergente em pó' in name.text.lower() or 'lava roupas em pó' in name.text.lower() or 'sabão em pó' in name.text.lower()):
                self.list_sabao_po.append({'product': name.text.split(' ')[0] + ' 1kg', 'price': price, 'supermarket': supermarket})

            elif 'detergente líquido' in name.text.lower():
                self.list_detergente.append({'product': name.text.split(' ')[0] + ' 500ml', 'price': price, 'supermarket': supermarket})

            elif 'papel' in name.text.lower():
                self.list_pagel_hig.append({'product': name.text.split(' ')[0], 'price': price, 'supermarket': supermarket})

