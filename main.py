from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from time import sleep
import random
import openpyxl

pagina = 0

pesquisar_produto = input('Qual produto pesquisar no site Mercado Livre: ')
print('Iniciando Busca... \n')
servico = Service(ChromeDriverManager().install())
navegador = webdriver.Chrome(service=servico)

link = 'https://www.mercadolivre.com.br/'

navegador.get(link)
sleep(random.randint(2, 5))
navegador.find_element('xpath', "//input[@class='nav-search-input']").send_keys(pesquisar_produto)
sleep(random.randint(2, 5))
navegador.find_element('xpath', "//div[@aria-label='Buscar']").click()
sleep(random.randint(2, 5))

# Criação da Planilha
planilha = openpyxl.Workbook()
planilha.create_sheet('analise_de_precos',0)
analise_de_precos_page = planilha['analise_de_precos']
analise_de_precos_page.append(['Produto', 'Preço'])

while True:
    sleep(random.randint(2, 6))
    lista_de_produtos = navegador.find_elements('xpath', '//h2[@class="ui-search-item__title shops__item-title"]')
    if len(lista_de_produtos) == 0:
        lista_de_produtos = navegador.find_elements('xpath',
                                                    '//h2[@class="ui-search-item__title ui-search-item__group__element shops__items-group-details shops__item-title"]')

    lista_de_precos = navegador.find_elements('xpath',
                                              "//div[@class='ui-search-price ui-search-price--size-medium shops__price']//div[@class='ui-search-price__second-line shops__price-second-line']//span[@class='price-tag-fraction']")

    for produto, preco in zip(lista_de_produtos, lista_de_precos):
        analise_de_precos_page.append([produto.text, preco.text])
    planilha.save(f'mercado livre - {pesquisar_produto}.xlsx')

    pagina += 1
    sleep(2)
    print(f'Analisando {pagina}º pagina')

    try:
        sleep(random.randint(2, 5))

        navegador.execute_script('window.scroll(0,document.body.scrollHeight);')
        botao_proximo = navegador.find_element('xpath',
                                               '//li[@class="andes-pagination__button andes-pagination__button--next shops__pagination-button"]')
        botao_proximo.click()
    except:
        print('')
        print('Busca Finalizada!')
        print(f'Arquivo mercado livre - {pesquisar_produto}.xlsx criado com SUCESSO!')
        break
