from numpy import empty
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
import time
import pandas as pd

# Caminho para o driver do Chrome
PATH = 'C:\VIRTUALENVS\PROJETO_PRODUTOS_COMPRAS\chromedriver.exe'
# CEP utilizado nas buscas
CEP = '18071781'

class SuperMarket:
    """
    Classe para automatizar a busca de preços de produtos em diferentes supermercados online.
    """
    def __init__(self, supermarkets:list[str], products:list[str]):
        """
        Inicializa a classe SuperMarket.
        
        :param supermarkets: Lista de supermercados a serem pesquisados.
        :param products: Lista de produtos a serem buscados.
        """

        chrome_options = Options()
        chrome_options.add_argument("--start-maximized")  # Abre em tela cheia
        chrome_options.add_argument("--disable-blink-features=AutomationControlled")  # Mascarar automação
        # chrome_options.add_argument("--headless=new")  # Executa em modo headless (sem abrir janela)
        # cService = webdriver.ChromeService(executable_path=PATH)
        # self.driver = webdriver.Chrome(service=cService)
        self.driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)

        # Lista de URLs dos supermercados
        self.websites = ['https://tauste.com.br', 
                         'https://www.barbosasupermercados.com.br/', 
                         'https://www.confianca.com.br/sorocaba', 
                         'https://lista.mercadolivre.com.br', 
                         'https://www.coopsupermercado.com.br', 
                         'https://www.tendaatacado.com.br',
                         'https://loja.boasupermercados.com.br/'
                         ]

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
        self.list_creme_dental = []
        self.list_agua_sanitaria = []
        self.list_sabonete = []
        self.list_fio_dental = []
        self.list_molho_tomate = []
        self.list_azeite = []
        self.list_farinha_trigo = []
        self.list_queijo = []
        self.list_creme_leite = []


    def extract_data(self) -> None:
        """
        Método principal para extrair os dados de preços dos produtos em cada supermercado da lista.
        """
        for supermarket in self.supermarkets:
            try:
                if supermarket.upper() == 'TAUSTE':
                    # Acessa o site do supermercado Tauste
                    self.driver.get(self.websites[0])
                    self.driver.maximize_window()
                    # time.sleep(2)

                    # Seleciona a opção de compra na loja física
                    # self.driver.find_element(by=By.XPATH, value='//label[@class="label-choice instore"]').click()
                    WebDriverWait(self.driver, 10, .1).until(EC.element_to_be_clickable((By.XPATH, '//label[@class="label-choice instore"]'))).click()
                    # time.sleep(2)

                    # Seleciona a cidade da loja
                    # select_element = self.driver.find_element(by=By.XPATH, value='//select[@id="stores-city"]')
                    select_element = WebDriverWait(self.driver, 10, .1).until(EC.presence_of_element_located((By.XPATH, '//select[@id="stores-city"]')))
                    select = Select(select_element)
                    select.select_by_value('6')
                    # time.sleep(20)

                    # Seleciona a loja especifica
                    # self.driver.find_element(by=By.XPATH, value='//*[@id="store-code-city-selected"]/div/li/div/strong[1]').click()
                    WebDriverWait(self.driver, 15, .1).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="store-code-city-selected"]/div/li/div/strong[1]'))).click()
                    # time.sleep(10)

                    for product in self.products:
                        
                        # Padroniza a unidade de medida dos produtos
                        product = (product.replace('5kg', '5000g')
                                        .replace('1kg', '1000g')
                                        .replace('1L', '1000ml')
                                        .replace('2l', '2000ml'))
                                        # .replace('Açúcar Cristal', 'Açucar'))

                        time.sleep(4)
                        # Realiza a busca do produto no site do supermercado
                        # self.driver.find_element(by=By.XPATH, value='//input[contains(@id, "search")]').send_keys(product.replace('Sabão em ', 'Detergente '))
                        WebDriverWait(self.driver, 20, .1).until(EC.element_to_be_clickable((By.XPATH, '//input[contains(@id, "search")]'))).send_keys(product.replace('Sabão em ', 'Detergente '))
                        WebDriverWait(self.driver, 15, .1).until(EC.element_to_be_clickable((By.XPATH, '//button[contains(@title, "Pesquisa")]'))).click()
                        # self.driver.find_element(by=By.XPATH, value='//button[contains(@title, "Pesquisa")]').click()

                        time.sleep(2)
                        # Localiza os produtos na pagina
                        contains_txt = self.set_contains(product=product.replace('Sabão em ', 'Detergente '))
                        # products_items = self.driver.find_elements(by=By.XPATH, value=f'//li[@class="item product product-item"][{contains_txt}]')
                        products_items = WebDriverWait(self.driver, 20, .1).until(EC.presence_of_all_elements_located((By.XPATH, f'//li[@class="item product product-item"][{contains_txt}]')))
                        
                        # Verifica se o produto foi encontrado na pagina
                        if not products_items:
                            continue

                        # Ordena a busca por preços crescentes
                        # time.sleep(7)
                        # select_element = self.driver.find_element(by=By.XPATH, value='//select[@id="sorter"]')
                        select_element = WebDriverWait(self.driver, 10, .1).until(EC.presence_of_element_located((By.XPATH, '//select[@id="sorter"]')))
                        select = Select(select_element)
                        select.select_by_value('price')
                        # time.sleep(2)
                        # button = self.driver.find_element(by=By.XPATH, value='//a[contains(@class, "action sorter-action sort-desc")]')
                        button = WebDriverWait(self.driver, 10, .1).until(EC.element_to_be_clickable((By.XPATH, '//a[contains(@class, "action sorter-action sort-desc")]')))
                        self.driver.execute_script("arguments[0].click();", button)
                        # time.sleep(5)

                        # Localiza os produtos na pagina
                        contains_txt = self.set_contains(product=product.replace('Sabão em ', 'Detergente '))
                        # products_rows = self.driver.find_elements(by=By.XPATH, value=f'//li[@class="item product product-item"][{contains_txt}]')
                        products_rows = WebDriverWait(self.driver, 10, .1).until(EC.presence_of_all_elements_located((By.XPATH, f'//li[@class="item product product-item"][{contains_txt}]')))
                        
                        # Verifica se o produto foi encontrado na pagina
                        if not products_rows:
                            continue

                        # Usa o primeiro produto encontrado correspondente ao mais barato
                        product = products_rows[0]

                        # Extrai o nome e o preço do produto
                        # name = product.find_element(by=By.XPATH, value='.//strong[@class="product name product-item-name"]')
                        name = WebDriverWait(product, 10, .1).until(EC.presence_of_element_located((By.XPATH, './/strong[@class="product name product-item-name"]')))
                        print(name.text)
                        # price = product.find_element(by=By.XPATH, value='.//div[@class="price-box price-final_price"]//span[@class="price"]').text
                        price = WebDriverWait(product, 10, .1).until(EC.presence_of_element_located((By.XPATH, './/div[@class="price-box price-final_price"]//span[@class="price"]'))).text
                        print(price)
                        
                        # Adiciona o produto e o preço na respectiva lista
                        self.append_to_list(name=name, price=price, supermarket='TAUSTE')

                        # time.sleep(2)
                
                elif supermarket.upper() == 'BARBOSA':
                    # Acessa o site do supermercado Barbosa
                    website = self.websites[1]
                    self.driver.get(website)
                    self.driver.maximize_window()

                    time.sleep(5)
                    self.driver.find_element(by=By.XPATH, value='(//div[@class="content"]//div[@class="close"])[1]').click()

                    for product in self.products:

                        print(product)

                        product = product.replace('Sabão em ', 'Detergente em ')

                        # Separa nome e tamanho do produto
                        product_title = ' '.join(product.split(' ')[0:-1])
                        product_size = product.split(' ')[-1]

                        # time.sleep(2)
                        # Realiza a busca do produto no site do supermercado
                        self.driver.get(f'{website}/buscar?texto={product_title}&ordem=Valor+crescente')
                        time.sleep(2)
                        try:
                            WebDriverWait(self.driver, 6, .2).until(EC.element_to_be_clickable((By.XPATH, '(//div[@class="content"]//div[@class="close"])[1]'))).click()
                        except:
                            pass
                        # self.driver.find_element(by=By.XPATH, value='(//div[@class="content"]//div[@class="close"])[1]').click()

                        # Localiza os produtos na página
                        contains_txt = self.set_contains(product)
                        # time.sleep(3)
                        # container_products = self.driver.find_element(by=By.XPATH, value='//div[@class="items-container"]')
                        container_products = WebDriverWait(self.driver, 10, .2).until(EC.presence_of_element_located((By.XPATH, '//div[@class="items-container"]')))
                        products_rows = container_products.find_elements(by=By.XPATH, value=f'.//div[@class="infinite-scroll-component items row"]//div[contains(@class, "row-item")][{contains_txt}]')
                        # products_rows = WebDriverWait(container_products, 20, .2).until(EC.presence_of_all_elements_located((By.XPATH, f'.//div[@class="infinite-scroll-component items row"]//div[contains(@class, "row-item")][{contains_txt}]')))
                        
                        without_5kg = True

                        # Valida se o produto não foi encontrado e se é Açúcar Cristal
                        if not products_rows and product_title == 'Açúcar Cristal':
                            # Define o produto como Açúcar Cristal 1kg
                            product = 'Açúcar Cristal 1kg'
                            without_5kg = False
                            
                            # Realiza uma nova localização dos produtos
                            contains_txt = self.set_contains(product)
                            # products_rows = container_products.find_elements(by=By.XPATH, value=f'.//div[@class="infinite-scroll-component items row"]//div[contains(@class, "row-item")][{contains_txt}]')
                            products_rows = WebDriverWait(container_products, 20, .2).until(EC.presence_of_all_elements_located((By.XPATH, f'.//div[@class="infinite-scroll-component items row"]//div[contains(@class, "row-item")][{contains_txt}]')))
                        
                        # Valida se o produto nao foi encontrado
                        if not products_rows:
                            continue

                        # Usa o primeiro produto encontrado correspondente ao mais barato
                        product = products_rows[0]
                        
                        # Extrai o nome e o preço do produto
                        # name = product.find_element(by=By.XPATH, value='.//div[@class="stats"]//div[@class="title"]')
                        name = WebDriverWait(product, 10, .2).until(EC.presence_of_element_located((By.XPATH, './/div[@class="stats"]//div[@class="title"]')))
                        print(name.text)

                        if without_5kg == False:
                            price = (product.find_element(by=By.XPATH, value='.//div[@class="prices-buy"]//div[@class="prices"]//p[@class="price"]')
                                            .text
                                            .replace('R$ ', '')
                                            .replace(',', '.'))
                            
                            price = 'R$ ' + str(round(float(price) * 5, 2))
                            print(price)
                            
                        else:
                            price = product.find_element(by=By.XPATH, value='.//div[@class="prices-buy"]//div[@class="prices"]//p[@class="price mt-1"]')
                            price = price.text
                            print(price)

                        # Adiciona o produto e o preço na respectiva lista
                        self.append_to_list(name, price, supermarket='BARBOSA')

                    time.sleep(2)

                elif supermarket.upper() == 'CONFIANÇA':
                    
                    # Acessa o site do supermercado Confiança
                    website = self.websites[2]
                    self.driver.get(website)
                    self.driver.maximize_window()
                    time.sleep(2)

                    # Fecha o modal de boas vindas
                    # self.driver.find_element(by=By.XPATH, value='//div[@class="modal__close__welcome"]').click()
                    # WebDriverWait(self.driver, 10, .1).until(EC.element_to_be_clickable((By.XPATH, '//div[@class="modal__close__welcome"]'))).click()
                    WebDriverWait(self.driver, 10, .1).until(EC.element_to_be_clickable((By.XPATH, '(//p[contains(text(), "Sorocaba")])[2]'))).click()

                    # Realiza a busca dos produtos
                    for product in self.products:

                        # Separa nome e tamanho do produto
                        final_product_name = '+'.join(product.split())

                        # Realiza a busca do produto no site do supermercado
                        self.driver.get(f'{website}/search?Ns=sku.activePrice%7C0&Ntt={final_product_name.lower()}')
                        time.sleep(2)

                        # Fecha o modal de boas vindas
                        # self.driver.find_element(by=By.XPATH, value='//div[@class="modal__close__welcome"]').click()
                        try:
                            WebDriverWait(self.driver, 10, .1).until(EC.element_to_be_clickable((By.XPATH, '//div[@class="modal__close__welcome"]'))).click()
                        except:
                            pass

                        # Localiza os produtos na página
                        contains_txt = self.set_contains(product)
                        # time.sleep(3)
                        # products_list = self.driver.find_element(by=By.XPATH, value='//div[@class="product-list product-list-fill-space"]')
                        products_list = WebDriverWait(self.driver, 10, .1).until(EC.presence_of_element_located((By.XPATH, '//div[@class="css-12dmjen e1k4it830"]')))
                        # products_rows = products_list.find_elements(by=By.XPATH, value=f'.//a[{contains_txt}]')
                        products_rows = WebDriverWait(products_list, 10, .1).until(EC.presence_of_all_elements_located((By.XPATH, f'.//a[{contains_txt}]')))

                        # Valida se o produto nao foi encontrado
                        if not products_rows:    
                            continue

                        # Usa o primeiro produto encontrado correspondente ao mais barato 
                        product = products_rows[0]
                        
                        # Extrai o nome e o preço do produto
                        # name = product.find_element(by=By.XPATH, value='.//article//div//h3')
                        # name = WebDriverWait(product, 10, .1).until(EC.presence_of_element_located((By.XPATH, './/article//div//h3')))
                        name = WebDriverWait(product, 10, .1).until(EC.presence_of_element_located((By.XPATH, './/p[@class="chakra-text css-1q1xfo8 e1k4it830"]')))
                        print(name.text)
                        # price = product.find_element(by=By.XPATH, value='.//span[@class="product-shelf__price-current"]').text
                        # price = WebDriverWait(product, 10, .1).until(EC.presence_of_element_located((By.XPATH, './/span[@class="product-shelf__price-current"]'))).text
                        price = WebDriverWait(product, 10, .1).until(EC.presence_of_element_located((By.XPATH, './/p[@class="chakra-text css-1m5u5z8 e1k4it830"]'))).text
                        print(price)

                        # Adiciona o produto e o preço na respectiva lista
                        self.append_to_list(name, price, supermarket='CONFIANÇA')

                        time.sleep(1)

                elif supermarket.upper() == 'MERCADO LIVRE':
                    
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
                        time.sleep(3)
                        container_products = WebDriverWait(self.driver, 10, .1).until(EC.presence_of_element_located((By.XPATH, '//ol[contains(@class, "ui-search-layout ui-search-layout")]')))
                        
                        if 'Detergente' in product or 'Sabão em Pó' in product or 'Papel Higiênico' in product:
                            products_rows = WebDriverWait(container_products, 10, .1).until(EC.presence_of_all_elements_located((By.XPATH, f'.//div[contains(@class, "andes-card poly-card poly-card--grid-card poly-card--large poly-card--CORE andes-card--flat andes-card--padding-0 andes-card--animated")][{contains_txt}]')))

                        else:
                            # products_rows = container_products.find_elements(by=By.XPATH, value=f'.//li[@class="ui-search-layout__item shops__layout-item"][{contains_txt}]')
                            try:
                                products_rows = WebDriverWait(container_products, 10, .1).until(EC.presence_of_all_elements_located((By.XPATH, f'.//li[@class="ui-search-layout__item"][{contains_txt}]')))
                            except:
                                continue
                        
                        # Valida se o produto nao foi encontrado
                        if not products_rows:
                            continue

                        # Usa o primeiro produto encontrado correspondente ao mais barato
                        product = products_rows[0]

                        # Extrai o nome e o preço do produto
                        # name = product.find_element(by=By.XPATH, value='.//h3')
                        name = WebDriverWait(product, 10, .1).until(EC.presence_of_element_located((By.XPATH, './/h3')))

                        print(name.text)
                        # price = (product.find_element(by=By.XPATH, value='.//span[@class="andes-money-amount andes-money-amount--cents-superscript"]')
                        #                 .text
                        #                 .replace('\n', ''))
                        price = WebDriverWait(product, 10, .1).until(EC.presence_of_element_located((By.XPATH, './/span[@class="andes-money-amount andes-money-amount--cents-superscript"]'))).text.replace('\n', '')
                        print(price)

                        # Adiciona o produto e o preço na respectiva lista
                        self.append_to_list(name, price, supermarket='MERCADO LIVRE')

                        time.sleep(2)

                elif supermarket.upper() == 'COOP SUPERMERCADO':
                    
                    # Acessa o site do supermercado Coop
                    website = self.websites[4]
                    self.driver.get(website)
                    self.driver.maximize_window()
                    time.sleep(6)

                    # Define o CEP do endereço de entrega
                    # self.driver.find_element(by=By.XPATH, value='//input[contains(@name,"localCep")]').send_keys(CEP)
                    WebDriverWait(self.driver, 10, .1).until(EC.element_to_be_clickable((By.XPATH, '//input[contains(@name,"localCep")]'))).send_keys(CEP)
                    
                    # time.sleep(3)
                    # self.driver.find_element(by=By.XPATH, value='//button[contains(@class, "coopsp-store-theme-YXbQXXox6msATjZ4Edfjr")]').click()
                    WebDriverWait(self.driver, 10, .1).until(EC.element_to_be_clickable((By.XPATH, '//button[contains(@class, "coopsp-store-theme-YXbQXXox6msATjZ4Edfjr")]'))).click()
                    # time.sleep(10)
                    # self.driver.find_element(by=By.XPATH, value='(//div[contains(@class, "coopsp-store-theme-8-x-localizationDeliveryOptionsShippingValue_ctaDesc")])[2]').click()
                    WebDriverWait(self.driver, 10, .1).until(EC.element_to_be_clickable((By.XPATH, '(//div[contains(@class, "coopsp-store-theme-8-x-localizationDeliveryOptionsShippingValue_ctaDesc")])[2]'))).click()
                    # time.sleep(6)

                    # Realiza a busca dos produtos
                    for product in self.products:
                        
                        # Realiza a busca do produto no site do supermercado
                        product_name = ('%20').join(product.split())
                        self.driver.get(f'{website}/{product_name}?_q={product_name}&map=ft&order=OrderByPriceASC')
                        print(f'{website}/{product_name}?_q={product_name}&map=ft&order=OrderByPriceASC')
                        # time.sleep(3)

                        # Localiza os produtos na página
                        contains_txt = self.set_contains(product.replace(' de ', ' '))
                        time.sleep(5)
                        # container = self.driver.find_element(by=By.XPATH, value='//div[contains(@class, "vtex-search-result-3-x-gallery vtex-search-result-3-x-gallery--grid flex flex-row flex-wrap items-stretch bn ph1 na4 pl9-l")]')
                        # products_rows = container.find_elements(by=By.XPATH, value=f'.//div[contains(@class, "vtex-search-result-3-x-galleryItem")][{contains_txt}][not(.//text()[contains(., "Indisponível")])]')
                        container = WebDriverWait(self.driver, 20, .1).until(EC.presence_of_element_located((By.XPATH, '//div[contains(@class, "vtex-search-result-3-x-gallery vtex-search-result-3-x-gallery--grid flex flex-row flex-wrap items-stretch bn ph1 na4 pl9-l")]')))
                        try:
                            products_rows = WebDriverWait(container, 20, .1).until(EC.presence_of_all_elements_located((By.XPATH, f'.//div[contains(@class, "vtex-search-result-3-x-galleryItem")][{contains_txt}][not(.//text()[contains(., "Indisponível")])]')))
                        except:
                            continue

                        # Valida se o produto nao foi encontrado
                        if not products_rows:
                            continue

                        # Usa o primeiro produto encontrado correspondente ao mais barato
                        product = products_rows[0]

                        # Extrai o nome e o preço do produto
                        try:
                            # name = product.find_element(by=By.XPATH, value='.//div[contains(@class, "vtex-product-summary-2-x-nameContainer")]//h3')
                            name = WebDriverWait(product, 10, .1).until(EC.presence_of_element_located((By.XPATH, './/div[contains(@class, "vtex-product-summary-2-x-nameContainer")]//h3')))
                            print(name.text)
                            # price = product.find_element(by=By.XPATH, value='.//span[contains(@class, "coopsp-store-theme-8-x-customProductSellingPrice_sellingPrice")]').text
                            price = WebDriverWait(product, 10, .1).until(EC.presence_of_element_located((By.XPATH, './/span[contains(@class, "coopsp-store-theme-8-x-customProductSellingPrice_sellingPrice")]'))).text
                            print(price)

                            # Adiciona o produto e o preço na respectiva lista
                            self.append_to_list(name, price, supermarket='COOP SUPERMERCADO')
                        
                        except Exception:
                            # Se o produto nao foi encontrado, adiciona o produto na respectiva lista
                            item_name = (product_name.split()[0]
                                                .lower()
                                                .replace('ç', 'c')
                                                .replace('ú', 'u')
                                                .replace('ó', 'o')
                                                .replace('ã', 'a')
                                                )
                        
                            list = getattr(self, f'list_{item_name}', [])
                            list.append({'product': item_name, 'price': None, 'supermarket': 'COOP SUPERMERCADO'})
                            continue

                        # time.sleep(5)

                elif supermarket.upper() == 'TENDA ATACADO':

                    # Acessa o site da Tenda Atacado
                    website = self.websites[5]
                    self.driver.get(website)
                    self.driver.maximize_window()
                    # time.sleep(10)

                    # Define o CEP do endereço de entrega
                    WebDriverWait(self.driver, 10, .1).until(EC.presence_of_element_located((By.XPATH, '//input[contains(@name, "searchbarComponent")]'))).click()
                    time.sleep(4)
                    WebDriverWait(self.driver, 20, .1).until(EC.presence_of_element_located((By.NAME, 'zipCode'))).send_keys(CEP)
                    WebDriverWait(self.driver, 10, .1).until(EC.element_to_be_clickable((By.XPATH, '//span[contains(text(), "Clique & Retire")]'))).click()
                    WebDriverWait(self.driver, 10, .1).until(EC.element_to_be_clickable((By.XPATH, '//span[contains(text(), "Sorocaba")]'))).click()


                    # Realiza a busca dos produtos
                    for product in self.products:
                        
                        product = product.replace('Sabão em ', 'Lava Roupas ')
                        # Realiza a busca do produto no site do supermercado
                        product_name = '+'.join(product.split())
                        self.driver.get(f'{website}/busca?q={product_name}')

                        # Define a ordenação dos produtos pelo preço
                        select_element = WebDriverWait(self.driver, 20, .1).until(EC.presence_of_element_located((By.XPATH, '//select[contains(@id, "select-sort-container")]')))
                        select = Select(select_element)
                        select.select_by_value('ascPrice')
                        time.sleep(3)

                        # Localiza os produtos na página
                        contains_txt = self.set_contains(product)
                        container_products = WebDriverWait(self.driver, 20).until(EC.presence_of_element_located((By.XPATH, '//section[contains(@class, "MosaicCardContainer grid_view")]')))
                        products_rows = WebDriverWait(container_products, 20).until(EC.presence_of_all_elements_located((By.XPATH, f'.//div[contains(@class, "ProductCardShowcase")][{contains_txt}]')))               

                        # Valida se o produto nao foi encontrado
                        if not products_rows:
                            continue
                        
                        # Usa o primeiro produto encontrado correspondente ao mais barato
                        product = products_rows[0]
                        # print(product.get_attribute('innerHTML'))
                        time.sleep(2)

                        # Extrai o nome e o preço do produto
                        name = WebDriverWait(product, 30, 1).until(EC.presence_of_element_located((By.XPATH, './/div[contains(@class, "informations")]//h3')))
                        print(name.text)
                       
                        price = WebDriverWait(product, 30, .1).until(EC.presence_of_element_located((By.XPATH, './/span[contains(@class, "showcase-price-area")]//div[contains(@class, "price-space")]'))).text.replace(' un', '')
                        print(price)

                        # Adiciona o produto e o preço na respectiva lista
                        self.append_to_list(name, price, supermarket='TENDA ATACADO')

                elif supermarket.upper() == 'BOA SUPERMERCADO':
                    
                    # Acessa o site da Boa Supermercado
                    website = self.websites[6]
                    self.driver.get(website)
                    self.driver.maximize_window()
                    # time.sleep(5)

                    # Clicar no botão de "Concordo" para os cookies
                    # self.driver.find_element(by=By.XPATH, value='//button[contains(@class, "modal-cookies-agree")]').click()
                    # WebDriverWait(self.driver, 10, .1).until(EC.element_to_be_clickable((By.XPATH, '//button[contains(@class, "modal-cookies-agree")]'))).click()

                    for product in self.products:

                        product = (product.replace('Sabão em ', 'Lava Roupas ')
                                          .replace('Mussarela', 'Muçarela'))
                        
                        # Realiza a busca do produto no site do supermercado
                        product_name = '+'.join(product.split())
                        self.driver.get(f'{website}s/?q={product_name}&sort=price_asc&page=0')
                        time.sleep(5)

                         # Seleciona a cidade da loja
                        # select_element = self.driver.find_element(by=By.XPATH, value='//select[@id="sort-select"]')
                        while len(self.driver.find_elements(by=By.XPATH, value='//a[@data-testid="show-more"]')) > 0:
                            print("QTD:", len(self.driver.find_elements(by=By.XPATH, value='//a[@data-testid="show-more"]')))
                            # button_pagination = self.driver.find_element(by=By.XPATH, value='//a[@data-testid="show-more"]')
                            time.sleep(7)
                            self.driver.execute_script("arguments[0].click();", WebDriverWait(self.driver, 20, .1).until(EC.element_to_be_clickable((By.XPATH, '//a[@data-testid="show-more"]'))))
                            # button_pagination.click()
                            #     # ActionChains(self.driver).move_to_element(button_pagination[0]).perform()
                            #     # WebDriverWait(self.driver, 10, .1).until(EC.element_to_be_clickable((By.XPATH, '//div[contains(@class, "button-pagination")]'))).click()

                        select_element = WebDriverWait(self.driver, 10, .1).until(EC.presence_of_element_located((By.XPATH, '//select[@id="sort-select"]')))
                        select = Select(select_element)
                        select.select_by_value('price_asc')
                        time.sleep(5)
                        # if not self.driver.find_elements(by=By.XPATH, value='//div[contains(@class, "button-pagination")]'):
                        #     break

                        products_container = self.driver.find_element(by=By.XPATH, value='//ul[contains(@class, "product-grid")]')

                        # print(products_container.get_attribute('innerHTML'))
                        # print(products_container)

                        contains_txt = self.set_contains(product=product)
                        # print(contains_txt)
                        # products_items = products_container.find_elements(by=By.XPATH, value=f'.//li[{contains_txt}]')
                        # print(f'{website}/s/?q={product_name}&sort=price_asc&page=0')
                        # print(f'//ul[contains(@class, "product-grid")]//li[{contains_txt}]')
                        time.sleep(5)
                        products_items = products_container.find_elements(by=By.XPATH, value=f'.//li[{contains_txt}]')
                        print(f"Quantidade: {len(products_items)}")

                         # Verifica se o produto foi encontrado na pagina
                        if not products_items:
                            continue

                         # Usa o primeiro produto encontrado correspondente ao mais barato
                        product = products_items[0]

                        # name = product.find_element(by=By.XPATH, value='.//div[contains(@class, "product-card-container")]//p[contains(@class, "product-card-name")]')
                        name = WebDriverWait(product, 15, 1).until(EC.presence_of_element_located((By.XPATH, './/p[contains(@class, "product-card-name")]')))
                        print(name.text)
                        # price = product.find_element(by=By.XPATH, value='.//div[contains(@class, "product-card-container")]//p[contains(@class, "product-card-new-price")]').text
                        price = WebDriverWait(product, 15, .1).until(EC.presence_of_element_located((By.XPATH, './/p[contains(@class, "product-card-new-price")]'))).text
                        print(price)

                        self.append_to_list(name, price, supermarket='BOA SUPERMERCADO')

                       # Adiciona o produto e o preço na respectiva lista
                        self.append_to_list(name, price, supermarket='BOA SUPERMERCADO')

            except Exception as e:
                print(f"Exceção: {e}. Linha: {e.__traceback__.tb_lineno}")
                continue
        # Exportando a tabela de preços
        self.export_table()

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

                if price is None or price is empty:
                    continue

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
            add_items(self.list_sabao_po, 'Sabão em Pó 800g')

        if self.list_pagel_hig:
            add_items(self.list_pagel_hig, 'Papel Higiênico')

        if self.list_creme_dental:
            add_items(self.list_creme_dental, 'Creme Dental 70g')

        if self.list_agua_sanitaria:
            add_items(self.list_agua_sanitaria, 'Água Sanitária')

        if self.list_sabonete:
            add_items(self.list_sabonete, 'Sabonete')

        if self.list_fio_dental:
            add_items(self.list_fio_dental, 'Fio Dental')

        if self.list_molho_tomate:
            add_items(self.list_molho_tomate, 'Molho de Tomate')
        
        if self.list_azeite: 
            add_items(self.list_azeite, 'Azeite 500ml')

        if self.list_farinha_trigo:
            add_items(self.list_farinha_trigo, 'Farinha de Trigo 1kg')

        if self.list_queijo:
            add_items(self.list_queijo, 'Queijo 150g')
        
        if self.list_creme_leite:
            add_items(self.list_creme_leite, 'Creme de Leite 200g')

        # Criando um DataFrame
        df = pd.DataFrame.from_dict(data, orient="index").reset_index()

        # Renomeando a coluna do índice para "Item"
        df = df.rename(columns={"index": "Item"})

        # Substituindo NaN por "Não encontrado"
        df = df.fillna(0.00)

        # Exportando para CSV
        df.to_csv("../precos_supermercados_3.csv", encoding="utf-8-sig", index=False)
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
                        contains_txt = contains_txt + f'contains(translate(., "ABCDEFGHIJKLMNOPQRSTUVWXYZÓÁÉ", "abcdefghijklmnopqrstuvwxyzóáé"), "{parte.lower()}") and '

        # Retorna o XPath
        return contains_txt[:-5]

    def append_to_list(self, name:str, price:str, supermarket:str):
        """
        Adiciona os produtos encontrados às listas de armazenamento correspondentes.
        
        :param name: Nome do produto encontrado.
        :param price: Preço do produto encontrado.
        :param supermarket: Nome do supermercado onde o produto foi encontrado.
        """

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

        elif 'leite' in name.text.lower().split(' ')[0]:
            self.list_leite.append({'product': name.text.split(' ')[0] + ' 1l', 'price': price, 'supermarket': supermarket})

        elif 'pão' in name.text.lower():
            self.list_pao.append({'product': name.text.split(' ')[0] + ' 500g', 'price': price, 'supermarket': supermarket})

        elif 'café' in name.text.lower():
            self.list_cafe.append({'product': name.text.split(' ')[0] + ' 500g', 'price': price, 'supermarket': supermarket})

        elif ('detergente em pó' in name.text.lower() or 'lava roupas em pó' in name.text.lower() or 'sabão em pó' in name.text.lower()) or 'lava roupas pó' in name.text.lower():
            self.list_sabao_po.append({'product': name.text.split(' ')[0] + ' 800g', 'price': price, 'supermarket': supermarket})

        elif 'detergente líquido' in name.text.lower():
            self.list_detergente.append({'product': name.text.split(' ')[0] + ' 500ml', 'price': price, 'supermarket': supermarket})

        elif 'papel higiênico' in name.text.lower():
            self.list_pagel_hig.append({'product': name.text.split(' ')[0], 'price': price, 'supermarket': supermarket})

        elif 'creme dental' in name.text.lower():
            self.list_creme_dental.append({'product': name.text.split(' ')[0] + ' 70g', 'price': price, 'supermarket': supermarket})

        elif 'água sanitária' in name.text.lower():
            self.list_agua_sanitaria.append({'product': name.text.split(' ')[0] + ' 2l', 'price': price, 'supermarket': supermarket})

        elif 'sabonete' in name.text.lower():
            self.list_sabonete.append({'product': name.text.split(' ')[0] + ' 70g', 'price': price, 'supermarket': supermarket})

        elif 'fio dental' in name.text.lower():
            self.list_fio_dental.append({'product': name.text.split(' ')[0] + ' 20cm', 'price': price, 'supermarket': supermarket})

        elif 'molho de tomate' in name.text.lower():
            self.list_molho_tomate.append({'product': name.text.split(' ')[0] + ' 300g', 'price': price, 'supermarket': supermarket})

        elif 'azeite' in name.text.lower():
            self.list_azeite.append({'product': name.text.split(' ')[0] + ' 500ml', 'price': price, 'supermarket': supermarket})

        elif 'farinha de trigo' in name.text.lower():
            self.list_farinha_trigo.append({'product': name.text.split(' ')[0] + ' 1kg', 'price': price, 'supermarket': supermarket})
        
        elif 'queijo' in name.text.lower():
            self.list_queijo.append({'product': name.text.split(' ')[0] + ' 150g', 'price': price, 'supermarket': supermarket})

        elif 'creme de leite' in name.text.lower():
            self.list_creme_leite.append({'product': name.text.split(' ')[0] + ' 200ml', 'price': price, 'supermarket': supermarket})



