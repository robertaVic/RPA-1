from time import sleep
from datetime import date
import gerenciadorPastas
from funcoes import padraoChrome
import shutil
import os
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains 
from openpyxl import load_workbook
# from selenium.webdriver.support.ui import WebDriverWait
# from selenium.webdriver.support import expected_conditions as EC
# from selenium.webdriver.common.by import By
# from selenium.common.exceptions import TimeoutException
from gerenciadorPlanilhas import preencher_solicitacao_pagamento_avulso


#retornar a data atual
today = date.today()
#formatar a data com pontos
data_em_texto = today.strftime("%d.%m.%Y")

#caminho da pasta macro(pasta do dia)
caminho_da_pasta = gerenciadorPastas.recuperar_diretorio_usuario() + "\\tpfe.com.br\\SGP e SGC - RPA\\Pagamento Avulso\\" 
#criar pasta do dia dentro de pagamento avulso
gerenciadorPastas.criarPastaData(caminho_da_pasta, data_em_texto)


#função para tramitar as solicitações
def pagamentoAvulso(financeiro):
    #pra uso de click
    builder = ActionChains(financeiro)
    financeiro.implicitly_wait(40)
    # sleep(10)
    # financeiro.refresh()
    financeiro.find_element_by_xpath("/html/body/div[1]/div/div[2]/main/section/div/div/div/div/section/div/div[2]/div").click()
    financeiro.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/div[2]/div/div/div[2]/div/div[2]/div[1]").click()
    url = "/runtime/44/list/190/Solicitação de Pgto Avulso"
    #clicar na parte de pagamento avulso
    financeiro.find_element_by_xpath('//a[@href="'+url+'"]').click()

    #Filtro
    #dataA = date.today().strftime("%d/%m/%Y")
   
    sleep(10)
    #limpar filtro  => financeiro.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[2]/button").click()
    filtro = financeiro.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/button[3]")
    try:
        filtro_click = builder.click(filtro)
        filtro_click.perform()
    except:
        pass
    try:
        filtro_click = builder.click(filtro)
        filtro_click.perform()
    except:
        pass    
    financeiro.find_element_by_xpath("/html/body/div[5]/div[3]/div/ul/li[3]/div/div/div/div").send_keys("\n")
    financeiro.find_element_by_xpath("/html/body/div[6]/div[3]/ul/li[3]").send_keys("\n")
    financeiro.find_element_by_xpath("/html/body/div[6]/div[1]").click()
    print(20*"=")
    # financeiro.find_element_by_xpath("/html/body/div[5]/div[3]/div/ul/li[10]/div/div/div/div/div/div[1]/div/input").send_keys(dataA)
    financeiro.find_element_by_xpath("/html/body/div[5]/div[3]/div/div[2]/button").click()#send_keys("\n")
    
    #abrir a linha (Laço para todos itens filtrados)
    sleep(3)
    #pegar o corpo da tabela de solicitaçoes
    tbody1 = financeiro.find_element_by_xpath("//*[@id='mainContent']/section/div/div/div/div[1]/div/div[3]/div/div/div/table/tbody")
    #pegar as linhas que existem dentro do corpo
    rows1 = tbody1.find_elements_by_tag_name("tr")

    global solicitaçao
    #lista que vai receber os índices de cada linha
    solicitaçao = [] 

    for row in range(len(rows1)):
        solicitaçao.append(row) #armazenando os indices de cada linha

    print(solicitaçao)

    #laço para tramitar cada solicitaçao
    for linha in solicitaçao: 
        # index1 = str(linha+1)
        #para cada linha guardar as informações dela
        global identificador
        #armazenando o id de cada solicitaçao
        identificador = financeiro.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[3]/div/div/div/table/tbody/tr[1]/td[4]/div").get_attribute("innerText")
        global razao
        #armazenando a razao social de cada solicitaçao
        razao = financeiro.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[3]/div/div/div/table/tbody/tr[1]/td[6]/div").get_attribute("innerText")
        # estado = financeiro.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[3]/div/div/div/table/tbody/tr[1]/td[7]").get_attribute("value")
        #criando um modelo de nome de pastas para serem salvas(igualmente ao modelo do financeiro)
        #criando uma condicional para saber se a pasta tem apenas id ou tem os dois(id + razao)
        if not razao:
            nome_da_pasta = (f"ID {identificador}")
        else:    
            nome_da_pasta = (f"ID {identificador} {razao}")

        print(nome_da_pasta)
        #com a funçao do outro arquivo, criar a pasta da atual solicitação de acordo com o laço
        gerenciadorPastas.criarPastasFilhas(nome_da_pasta)
      
        #tempo para salvar todas pastas
        sleep(1.5)
       
        #clicar na caixa de seleção especifica da linha
        financeiro.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[3]/div/div/div/table/tbody/tr[1]/td[2]/span/span[1]/input").click()
        #clicar no lápis de edição
        financeiro.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/div[3]/div/button[1]").click()
        #PEGAR TODAS AS INFORMAÇOES PARA ALIMENTAR A PLANILHA
        #criando lista para armazenar todos os dados do formulario
        dados_do_formulario = []
        caminho_em_comum_entre_campos_do_formulario = "/html/body/div[5]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[2]/div/div"
        #adicionando os valores à lista
        dados_do_formulario.append("Solicitação de Pgto Avulso")
        dados_do_formulario.append(identificador)
        dados_do_formulario.append(financeiro.find_element_by_xpath(caminho_em_comum_entre_campos_do_formulario + "[2]/div[1]/div/div/div/input").get_attribute("value"))
        dados_do_formulario.append(razao)
        dados_do_formulario.append(financeiro.find_element_by_xpath(caminho_em_comum_entre_campos_do_formulario + "[3]/div[1]/div/div[1]/div/div/div/input").get_attribute("value"))
        dados_do_formulario.append(financeiro.find_element_by_xpath(caminho_em_comum_entre_campos_do_formulario + "[3]/div[1]/div/div[2]/div/div/div/input").get_attribute("value"))
        dados_do_formulario.append(financeiro.find_element_by_xpath(caminho_em_comum_entre_campos_do_formulario + "[3]/div[2]/div/div[1]/div/div/div/input").get_attribute("value"))
        dados_do_formulario.append(financeiro.find_element_by_xpath(caminho_em_comum_entre_campos_do_formulario + "[3]/div[2]/div/div[2]/div/div/div/div/div/div/div[1]/input").get_attribute("value"))
        dados_do_formulario.append(financeiro.find_element_by_xpath(caminho_em_comum_entre_campos_do_formulario + "[4]/div[2]/div/div[3]/div/div/div/div/div/div/div[1]/input").get_attribute("value"))
        dados_do_formulario.append(financeiro.find_element_by_xpath(caminho_em_comum_entre_campos_do_formulario + "[4]/div[1]/div/div[1]/div/div/div/div/div/div/div[1]/input").get_attribute("value"))
        dados_do_formulario.append(financeiro.find_element_by_xpath(caminho_em_comum_entre_campos_do_formulario + "[4]/div[1]/div/div[2]/div/div/div/input").get_attribute("value"))
        dados_do_formulario.append(financeiro.find_element_by_xpath("0"))
        dados_do_formulario.append(financeiro.find_element_by_xpath(caminho_em_comum_entre_campos_do_formulario + "[5]/div[1]/div/div[1]/div/div/div/input").get_attribute("value"))
        dados_do_formulario.append(financeiro.find_element_by_xpath(caminho_em_comum_entre_campos_do_formulario + "[5]/div[1]/div/div[2]/div/div/div/input").get_attribute("value"))
        dados_do_formulario.append(financeiro.find_element_by_xpath(caminho_em_comum_entre_campos_do_formulario + "[5]/div[2]/div/div[1]/div/div/div/input").get_attribute("value"))
        dados_do_formulario.append("")                                                                             
        dados_do_formulario.append("")
        dados_do_formulario.append("Processada")
        
        #clicar em "notas fiscais"
        financeiro.find_element_by_xpath("/html/body/div[5]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[1]/div/div[2]/div/button[2]").click()

        tbody2 = financeiro.find_element_by_xpath("/html/body/div[5]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[3]/div/div/div/div[1]/div[3]/table/tbody")
        #pega todas as linhas que contem nf
        rows2 = tbody2.find_elements_by_tag_name("a") 
       
        #clicar só se houver elementos 
        try:
            if len(rows2) > 0:
                for row in rows2:
                    row.click()
            else:
                comentario_nao_possui_nota = (f"A solicitação não possui notas fiscais para serem baixadas")      
                print(comentario_nao_possui_nota)  
        except:
            comentario_nota_fiscal = (f"Não foi possível baixar a nota fiscal")     
            print(comentario_nota_fiscal) 

        sleep(4)    
        #imprimindo
        financeiro.find_element_by_xpath("/html/body/div[5]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[1]/div/div[2]/div/button[1]").click() 
        financeiro.find_element_by_xpath("/html/body/div[5]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[2]/div/div[6]/div[2]/div/div[2]/div/div/button").click()  
        #outra forma de clique diretamente no elemento
        financeiro.switch_to_frame(0)
        imprimir = financeiro.find_element_by_xpath("/html/body/div/div/div/div[2]/div/table/tbody/tr/td[1]/table/tbody/tr/td[3]/div/table/tbody/tr").click() 
        
        try:
            click_me = builder.click(imprimir)
            click_me.perform()
        except:
            print("1- Não consegui clicar")
        try:
            click_me = builder.click(imprimir)
            click_me.perform()
        except:
            print("2- Não consegui clicar")
            

        financeiro.find_element_by_xpath("/html/body/div/div/div/div[16]/div/div[1]/table/tbody/tr/td[2]").click()
        financeiro.find_element_by_xpath("/html/body/div/div/div/div[20]/div[4]/table/tbody/tr/td[1]/div/table/tbody/tr/td").click()
        financeiro.switch_to.default_content()
        # financeiro.find_element_by_xpath("/html/body/div[8]/div[3]").send_keys(Keys.ESC)
        financeiro.find_element_by_xpath("/html/body/div[8]/div[3]/div/div[1]/h2/div/div[2]/button").click()
        #financeiro.find_element_by_xpath("/html/body/div[5]/div[3]/div/div/div/div[4]/fieldset/button[2]").click()
        # financeiro.find_element_by_xpath("/html/body/div[8]/div[3]/div/div[2]/ul/div[3]").click()
        print("passou")
        # financeiro.find_element_by_xpath("/html/body/div[5]/div[3]/div/div/div/div[4]/fieldset/button[2]").click()
        financeiro.find_element_by_xpath("/html/body/div[5]/div[3]/div/div/div/div[1]/div/div[3]/button").click()

    
        # if not comentario_nota_fiscal:
        #     comentario = (f"Nenhum")
        # else:
        #     comentario = (f"2- {comentario_nota_fiscal}")   
        #     if not comentario_nao_possui_nota:
        #         comentario = (f"Nenhum")
        #     else:
        #         comentario = (f"3- {comentario_nao_possui_nota}")

        preencher_solicitacao_pagamento_avulso(dados_do_formulario)
            
        sleep(1.5)
        print("mover arquivos")
        #listando os arquivos baixados na pasta macro(pasta do dia)
        arquivos = gerenciadorPastas.listar_arquivos_em_diretorios(caminho_da_pasta)
        print(arquivos)
        #criando um laço para mover cada um para sua pasta especifica
        for arquivo in arquivos:
            print(arquivo)
            #movendo os arquivos para a pasta da sua solicitaçao
            shutil.move(caminho_da_pasta + arquivo, caminho_da_pasta + nome_da_pasta +"\\" + arquivo)
            print("moveu o arquivo!")

        #tramitação das solicitaçoes
        tramitar = financeiro.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/div[3]/div/button[2]")
        
        try:
            tramitar.send_keys("\n")
        except:
            print("1- Não consegui clicar")
        try:
            tramitar.click()
        except:
            print("1- Não consegui clicar")

        financeiro.find_element_by_xpath("/html/body/div[5]/div[3]/div/div[2]/ul/div[3]").click()
        sleep(1.5)

        
        # financeiro.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/button[2]").click()    

    # #para cada solicitação marcar a caixa de selecionar todas
    # financeiro.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[3]/div/div/div/table/thead/tr/th[2]/span/span[1]/input").click()
    # #desmarcar
    # financeiro.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[3]/div/div/div/table/thead/tr/th[2]/span/span[1]/input").click()
    # #exportando a planilha
    # financeiro.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/button[4]").click()






     # delay = 10 # seconds
    # try:
    #     myElem = WebDriverWait(financeiro, delay).until(EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/button[3]")))
    #     filtro = financeiro.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/button[3]")
    #     try:
    #         filtro_click = builder.click(filtro)
    #         filtro_click.perform()
    #     except:
    #         pass
        
   