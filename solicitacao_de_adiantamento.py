
from solicitacao_de_pagamento import encontrar_elemento_por_repeticao
from time import sleep
from datetime import date
from gerenciadorPastas import *
from funcoes import *
import shutil
import os
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains 
from openpyxl import load_workbook
from gerenciadorPlanilhas import preencher_solicitacao_na_planilha, ler_dados_da_planilha, atualizar_status_na_planilha, dados_para_tramitar_para_baixadas


def adiantamento(driver):
    #verificar se tem downloads antigos e apagar
    remover_arquivos_da_raiz(recuperar_diretorio_usuario() + "\\tpfe.com.br\\SGP e SGC - RPA\\")
    #data atual formatada
    data_em_texto = date.today().strftime("%d.%m.%Y")
    #caminho da pasta macro(pasta do dia)
    caminho_da_pasta = recuperar_diretorio_usuario() + "\\tpfe.com.br\\SGP e SGC - RPA\\Adiantamento\\" 
    #criar pasta do dia dentro de pagamento avulso
    criarPastaData(caminho_da_pasta, data_em_texto)
    tipo_de_solicitacao = "SA"
    builder = ActionChains(driver)
    #driver.implicitly_wait(2)

    #ACESSANDO SOLICITAÇAO DE ADIANTAMENTO
    espera_explicita_de_elemento(driver,"/html/body/div[1]/div/div[2]/main/section/div/div/div/div/section/div/div[2]/div","encontrar","AD",2)
    driver.get("https://tpf2.madrix.app/runtime/44/list/184/Solicitação de Adiantamento")
    # driver.implicitly_wait(10)

    #FILTRANDO AS SOLICITAÇÕES APROVADAS PELO GERENTE
    encontrar_elemento_por_repeticao(driver,"/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/div/div/div","click","Abrir filtro", 4)
    encontrar_elemento_por_repeticao(driver,"/html/body/div[4]/div[3]/ul/li[3]","click","aprovado",4)
    encontrar_elemento_por_repeticao(driver,"/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/button[3]","click","2-filtro", 4)
    encontrar_elemento_por_repeticao(driver,"/html/body/div[4]/div[3]/div/div[1]/div[1]/button","click","limpar filtro",4)
    encontrar_elemento_por_repeticao(driver, "/html/body/div[4]/div[3]/div/ul/li[3]/div/div/div/div", "click", "filtrar por",4)
    encontrar_elemento_por_repeticao(driver,"/html/body/div[5]/div[3]/ul/li[7]", "click", "2-aprovado", 4)
    encontrar_elemento_por_repeticao(driver,"/html/body/div[5]/div[1]", "click", "clique fora", 4)
    print(20*"=")
    encontrar_elemento_por_repeticao(driver,"/html/body/div[4]/div[3]/div/div[2]/button", "click", "fechando filtro", 4)
    
    #OBTER QUANTIDADE DE PAGAMENTOS
    sleep(5)
    # quantidade_de_requisicoes = int((driver.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/span[2]/div/p").get_attribute("innerText")).split(" ")[-1])
    
    #LAÇO PARA TRAMITAR TODOS AS SOLICITAÇÕES
    for linha in range(0): 
        print(linha)#voltar para antigo quantidades
        dados_do_formulario = [] 
        global identificador
        #armazenando o id de cada solicitaçao
        encontrar_elemento_por_repeticao(driver,"/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[3]/div/div/div/table/tbody/tr[1]/td[4]/div", "link", "Encontrar ID", 3)
        sleep(2)
        identificador = driver.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[3]/div/div/div/table/tbody/tr[1]/td[4]/div").get_attribute("innerText")
        solicitante = driver.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[3]/div/div/div/table/tbody/tr/td[6]/div/div[2]").get_attribute("innerText")
        #ACESSANDO A SOLICITAÇAO
        #espera_explicita_de_elemento(driver,"/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[3]/div/div/div/table/tbody/tr[1]/td[2]/span/span[1]/input", "encontrar", "clicar na caixinha", 3).click()
        encontrar_elemento_por_repeticao(driver,"/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[3]/div/div/div/table/tbody/tr[1]/td[2]/span/span[1]/input", "click", "clicar na caixinha", 3)
        sleep(3)
        #clicar no lápis de edição /html/body/div[4]/div[3]/div/div[3]/button[2]
        encontrar_elemento_por_repeticao(driver, "/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/div[3]/div/button[1]", "click", "click na linha", 0.3)
        #Acessando as informações presentes em 'Financeiro HTML'
        sleep(5)

      #  if linha == 0:
       #     encontrar_elemento_por_repeticao(driver, "/html/body/div[4]/div[3]/div/div[2]/div/div/div/div/div", "click", "escolher form", 4)
        #    filtro_click = builder.send_keys(Keys.ARROW_DOWN)
         #   filtro_click.perform()
          #  filtro_click = builder.send_keys(Keys.ARROW_DOWN)
           # filtro_click.perform()
            #filtro_click = builder.send_keys(Keys.SPACE)
            #filtro_click.perform()
            #try:
             #   encontrar_elemento_por_repeticao(driver, "/html/body/div[4]/div[3]/div/div[3]/button[2]", "click", "acessar form", 0.5)
            #except: 
             #   filtro_click = builder.send_keys(Keys.ARROW_DOWN)
              #  filtro_click.perform()
               # filtro_click = builder.send_keys(Keys.SPACE)
                #filtro_click.perform()
                #encontrar_elemento_por_repeticao(driver, "/html/body/div[4]/div[3]/div/div[3]/button[2]", "click", "acessar form", 0.5).click()
            #sleep(3)
            
    
        #Acessando as informações presentes em 'Financeiro HTML'
        encontrar_elemento_por_repeticao(driver, "/html/body/div[4]/div[3]/div/div[2]/div/div/div/div/div", "click", "escolher form", 4)
        filtro_click = builder.send_keys(Keys.DOWN)
        filtro_click = builder.send_keys(Keys.SPACE)
        filtro_click.perform()
        encontrar_elemento_por_repeticao(driver, "/html/body/div[4]/div[3]/div/div[3]/button[2]", "click", "acessar form", 4)
        sleep(3)
        #PEGAR TODAS AS INFORMAÇOES PARA ALIMENTAR A PLANILHA

        #AD
        dados_do_formulario.append(tipo_de_solicitacao)
        #ID DA SOLICITAÇAO
        dados_do_formulario.append(identificador)
        #CPF/CNPJ
        dados_do_formulario.append(driver.find_element_by_xpath("/html/body/div[4]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[2]/div/div[1]/div[2]/div/div/div/input").get_attribute("value"))
        #RAZAO SOCIAL
        dados_do_formulario.append("")
        #FORMA DE PAGAMENTO
        dados_do_formulario.append("Transferência Bancária")
        #BANCO 
        dados_do_formulario.append(driver.find_element_by_xpath("/html/body/div[4]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[2]/div/div[2]/div[2]/div/div[1]/div/div[1]/div/div/div/input").get_attribute("value"))
        #AGENCIA
        dados_do_formulario.append(driver.find_element_by_xpath("/html/body/div[4]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[2]/div/div[2]/div[2]/div/div[1]/div/div[2]/div/div/div/input").get_attribute("value"))
        #CONTA
        dados_do_formulario.append(driver.find_element_by_xpath("/html/body/div[4]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[2]/div/div[2]/div[2]/div/div[2]/div/div[1]/div/div/div/input").get_attribute("value"))
        #TIPO DE CONTA 
        dados_do_formulario.append(driver.find_element_by_xpath("/html/body/div[4]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[2]/div/div[2]/div[2]/div/div[2]/div/div[2]/div/div/div/input").get_attribute("value"))
        #NATUREZA DA CONTA
        dados_do_formulario.append("")
        #VALOR
        valor = driver.find_element_by_xpath("/html/body/div[4]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[2]/div/div[4]/div[1]/div/div[1]/div/div/div/input").get_attribute("value")
        dados_do_formulario.append(valor)
        #VALOR PAGO
        dados_do_formulario.append("")
        #DATA SOLICITADA PARA PAGAMENTO
        dados_do_formulario.append("")
        #DATA DA SOLICITAÇÃO 
        dados_do_formulario.append(driver.find_element_by_xpath("/html/body/div[4]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[2]/div/div[4]/div[1]/div/div[2]/div/div/div/input").get_attribute("value"))
        #DATA DE PAGAMENTO 
        dados_do_formulario.append("")

        valor_da_conta = valor.replace(".","")
        valor_da_conta = valor_da_conta.replace("R$","")
        valor_da_conta = valor_da_conta.replace(",",".")
        valor_da_conta = float(valor_da_conta)

        if dados_do_formulario[5] == "" or  dados_do_formulario[6] == "" or  dados_do_formulario[7] == "" or  dados_do_formulario[8] == "" or valor_da_conta == 0:# and  banco != ""
            #Comentario Robo
            dados_do_formulario.append("Dados bancários incompletos ou solicitação está com valor zerado.")
        else:
            #Comentario Robo 
            dados_do_formulario.append("")
        #Ajuste
        dados_do_formulario.append("")
        
        ####PASSO QUE DEVE SER DADO APÓS O GERENTE ADICIONAR AS NOTAS####
        #clicar em "notas fiscais"
        encontrar_elemento_por_repeticao(driver, "/html/body/div[4]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[1]/div/div[2]/div/button[2]", "click", "clicar em notas", 3)
        tbody2 = driver.find_element_by_xpath("/html/body/div[4]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[3]/div/div/div/div[1]/div[3]/table/tbody")
        # #pega todas as linhas que contem nf
        rows2 = tbody2.find_elements_by_tag_name("a") 

        #Baixando Nfs
        try:
            for row in rows2:
                maximo_tentativas = 0
                while maximo_tentativas < 40: 
                    if len(rows2) > 0:
                        row.click()
                        sleep(2)
                        maximo_tentativas = 40
                    else:
                        maximo_tentativas += 1
        except:
            comentario_nota_fiscal = (f"Não foi possível baixar a nota fiscal")     
            print(comentario_nota_fiscal)
            dados_do_formulario[15] = comentario_nota_fiscal 
 
        sleep(2)
        #IMPRIMINDO A CAPA
        encontrar_elemento_por_repeticao(driver, "/html/body/div[4]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[1]/div/div[2]/div/button[1]", "click", "VOLTAR", 5)
        encontrar_elemento_por_repeticao(driver, "/html/body/div[4]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[2]/div/div[7]/div[2]/div/div/button", "click", "imprimir", 5)
        driver.switch_to_frame(0)
        encontrar_elemento_por_repeticao(driver,"/html/body/div/div/div/div[2]/div/table/tbody/tr/td[1]/table/tbody/tr/td[3]/div/table/tbody/tr","click","filtro-impressao",2) 
        encontrar_elemento_por_repeticao(driver, "/html/body/div/div/div/div[16]/div/div[1]/table/tbody/tr/td[2]", "click", "baixar capa", 2)
        encontrar_elemento_por_repeticao(driver, "/html/body/div/div/div/div[20]/div[4]/table/tbody/tr/td[1]/div/table/tbody/tr/td", "click", "baixar capa", 2)
        driver.switch_to.default_content()
        encontrar_elemento_por_repeticao(driver, "/html/body/div[7]/div[3]/div/div[1]/h2/div/div[2]/button", "click", "baixar capa", 2)
        print("passou")
        encontrar_elemento_por_repeticao(driver, "/html/body/div[4]/div[3]/div/div/div/div[1]/div/div[3]/button", "click", "fechar", 2)

        #CRIANDO A PASTA
        solicitante1 = solicitante.replace("\\" , "")
        solicitante1 = solicitante1.replace("/", "")
        solicitante1 = solicitante1.replace(":", "")
        solicitante1 = solicitante1.replace("*", "")
        solicitante1 = solicitante1.replace("?", "")
        solicitante1 = solicitante1.replace('"', "")
        solicitante1 = solicitante1.replace("<", "")
        solicitante1 = solicitante1.replace(">", "")
        solicitante1 = solicitante1.replace("|", "")
        solicitante1 = solicitante1.replace(".", "")
        solicitante1 = solicitante1.replace(",", "")

        nome_da_pasta = (f"ID {identificador} {solicitante1}")
        print(nome_da_pasta)
        # #com a funçao do outro arquivo, criar a pasta da atual solicitação de acordo com o laço
        criarPastasFilhas("Adiantamento", nome_da_pasta)
      
        #tempo para salvar todas pastas
        sleep(3)
        # MOVENDO ARQUIVOS

        validar_download(caminho_da_pasta, data_em_texto, nome_da_pasta)
        sleep(3)
        #TRAMITAÇÃO DAS SOLICITAÇÕES
        encontrar_elemento_por_repeticao(driver,"/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/div[3]/div/button[2]","click","tramitar",4)
        encontrar_elemento_por_repeticao(driver, "/html/body/div[4]/div[3]/div/div[2]/ul/div[1]", "click", "processar", 4)
        # #status robo
        dados_do_formulario.append("Processada")

        #DATA DE EXECUÇÃO ROBO
        dados_do_formulario.append(date.today().strftime("%d/%m/%Y"))
        
        preencher_solicitacao_na_planilha(dados_do_formulario, tipo_de_solicitacao)
        sleep(5)
        driver.get("https://tpf2.madrix.app/runtime/44/list/184/Solicitação de Adiantamento")
        sleep(2)

        

   #2° PARTE: ESPERANDO DO FINANCEIRO PRA TRAMITAR PRA PAGO 
    lista_de_tramitacao = ler_dados_da_planilha(tipo_de_solicitacao)
    if len(lista_de_tramitacao) > 0:
        encontrar_elemento_por_repeticao(driver,"/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/div/div/div","click","filtro", 4)
        encontrar_elemento_por_repeticao(driver,"/html/body/div[4]/div[3]/ul/li[11]","click","processado",2)
        for solicitacao in lista_de_tramitacao:
            print(lista_de_tramitacao[0] + "" + solicitacao)
            encontrar_elemento_por_repeticao(driver,"/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/button[3]","click","2-filtro", 2)
            encontrar_elemento_por_repeticao(driver,"/html/body/div[4]/div[3]/div/div[1]/div[1]/button","click","limpar",3)
            sleep(4)
            driver.find_element_by_xpath("/html/body/div[4]/div[3]/div/ul/li[1]/div/div/div/div/input").send_keys(str(solicitacao[0]))
            encontrar_elemento_por_repeticao(driver, "/html/body/div[4]/div[3]/div/ul/li[3]/div/div/div/div", "click", "filtrar", 3)
            encontrar_elemento_por_repeticao(driver, "/html/body/div[5]/div[3]/ul/li[9]", "click", "processados", 0.4)
            encontrar_elemento_por_repeticao(driver, "/html/body/div[5]/div[1]", "click", "clique fora", 0.4 )
            encontrar_elemento_por_repeticao(driver, "/html/body/div[4]/div[3]/div/div[2]/button", "click", "fechar filtro", 4)
            sleep(5)
            encontrar_elemento_por_repeticao(driver, "/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[3]/div/div/div/table/tbody/tr/td[2]/span/span[1]/input", "click", "click na caixinha", 4 )
            encontrar_elemento_por_repeticao(driver, "/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/div[3]/div/button[2]", "click", "tramitar", 4 )
            encontrar_elemento_por_repeticao(driver, "/html/body/div[4]/div[3]/div/div[2]/ul/div", "click", "pagar", 4)
            atualizar_status_na_planilha(int(solicitacao[4]))    
    
    #3° PARTE - TRAMITAR PARA BAIXADAS
    lista_para_baixadas = dados_para_tramitar_para_baixadas(tipo_de_solicitacao)
    if len(lista_para_baixadas) > 0:
        encontrar_elemento_por_repeticao(driver,"/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/div/div/div","click","filtro", 4)
        encontrar_elemento_por_repeticao(driver,"/html/body/div[4]/div[3]/ul/li[10]","click","prestaçao",2)
        for solicitacao in lista_para_baixadas:
            print(solicitacao)
            encontrar_elemento_por_repeticao(driver,"/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/button[3]","click","2-filtro", 2)
            encontrar_elemento_por_repeticao(driver,"/html/body/div[4]/div[3]/div/div[1]/div[1]/button","click","limpar",3)
            sleep(4)
            driver.find_element_by_xpath("/html/body/div[4]/div[3]/div/ul/li[1]/div/div/div/div/input").send_keys(str(solicitacao[0]))
            encontrar_elemento_por_repeticao(driver, "/html/body/div[4]/div[3]/div/ul/li[3]/div/div/div/div", "click", "filtrar", 3)
            encontrar_elemento_por_repeticao(driver, "/html/body/div[5]/div[3]/ul/li[6]", "click", "prestacao", 0.4)
            encontrar_elemento_por_repeticao(driver, "/html/body/div[5]/div[1]", "click", "clique fora", 0.4 )
            encontrar_elemento_por_repeticao(driver, "/html/body/div[4]/div[3]/div/div[2]/button", "click", "fechar filtro", 4)
            sleep(5)
            encontrar_elemento_por_repeticao(driver,"/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[3]/div/div/div/table/tbody/tr[1]/td[2]/span/span[1]/input", "click", "clicar na caixinha", 3)
            sleep(3)
            #clicar no lápis de edição /html/body/div[4]/div[3]/div/div[3]/button[2]
            encontrar_elemento_por_repeticao(driver, "/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/div[3]/div/button[1]", "click", "click na linha", 0.3)
            #Acessando as informações presentes em 'Financeiro HTML'
            sleep(5)
            encontrar_elemento_por_repeticao(driver, "/html/body/div[4]/div[3]/div/div[2]/div/div/div/div/div", "click", "escolher form", 4)
            filtro_click = builder.send_keys(Keys.DOWN)
            filtro_click = builder.send_keys(Keys.SPACE)
            filtro_click.perform()
            encontrar_elemento_por_repeticao(driver, "/html/body/div[4]/div[3]/div/div[3]/button[2]", "click", "acessar form", 4)
            sleep(3)
             #clicar em "notas fiscais"
            encontrar_elemento_por_repeticao(driver, "/html/body/div[4]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[1]/div/div[2]/div/button[2]", "click", "clicar em notas", 3)
            tbody2 = driver.find_element_by_xpath("/html/body/div[4]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[3]/div/div/div/div[1]/div[3]/table/tbody")
            # #pega todas as linhas que contem nf
            rows2 = tbody2.find_elements_by_tag_name("a") 

            #Baixando Nfs
            try:
                for row in rows2:
                    maximo_tentativas = 0
                    while maximo_tentativas < 40: 
                        if len(rows2) > 0:
                            row.click()
                            sleep(2)
                            maximo_tentativas = 40
                        else:
                            maximo_tentativas += 1
            except:
                comentario_nota_fiscal = (f"Não foi possível baixar a nota fiscal")     
                print(comentario_nota_fiscal)
                dados_do_formulario[15] = comentario_nota_fiscal 
    
            sleep(2)
            #IMPRIMINDO A CAPA
            encontrar_elemento_por_repeticao(driver, "/html/body/div[4]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[1]/div/div[2]/div/button[1]", "click", "VOLTAR", 5)
            encontrar_elemento_por_repeticao(driver, "/html/body/div[4]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[2]/div/div[7]/div[2]/div/div/button", "click", "imprimir", 5)
            driver.switch_to_frame(0)
            encontrar_elemento_por_repeticao(driver,"/html/body/div/div/div/div[2]/div/table/tbody/tr/td[1]/table/tbody/tr/td[3]/div/table/tbody/tr","click","filtro-impressao",2) 
            encontrar_elemento_por_repeticao(driver, "/html/body/div/div/div/div[16]/div/div[1]/table/tbody/tr/td[2]", "click", "baixar capa", 2)
            encontrar_elemento_por_repeticao(driver, "/html/body/div/div/div/div[20]/div[4]/table/tbody/tr/td[1]/div/table/tbody/tr/td", "click", "baixar capa", 2)
            driver.switch_to.default_content()
            encontrar_elemento_por_repeticao(driver, "/html/body/div[7]/div[3]/div/div[1]/h2/div/div[2]/button", "click", "baixar capa", 2)
            print("passou")
            encontrar_elemento_por_repeticao(driver, "/html/body/div[4]/div[3]/div/div/div/div[1]/div/div[3]/button", "click", "fechar", 2)
            solicitante = str(solicitacao[1])
            #CRIANDO A PASTA
            if solicitante != "" or solicitante != "None":
                solicitante1 = solicitante.replace("\\" , "")
                solicitante1 = solicitante1.replace("/", "")
                solicitante1 = solicitante1.replace(":", "")
                solicitante1 = solicitante1.replace("*", "")
                solicitante1 = solicitante1.replace("?", "")
                solicitante1 = solicitante1.replace('"', "")
                solicitante1 = solicitante1.replace("<", "")
                solicitante1 = solicitante1.replace(">", "")
                solicitante1 = solicitante1.replace("|", "")
                solicitante1 = solicitante1.replace(".", "")
                solicitante1 = solicitante1.replace(",", "")

                nome_da_pasta = (f"ID {solicitacao[0]} {solicitante1}")
            else:
                nome_da_pasta = (f"ID {solicitacao[0]} {solicitante}")   
            print(nome_da_pasta)
            # #com a funçao do outro arquivo, criar a pasta da atual solicitação de acordo com o laço
            criarPastasFilhas("Adiantamento", nome_da_pasta)
        
            #tempo para salvar todas pastas
            sleep(3)
            # MOVENDO ARQUIVOS

            validar_download(caminho_da_pasta, data_em_texto, nome_da_pasta)
            sleep(3)

            #TRAMITAÇÃO DAS SOLICITAÇÕES
            encontrar_elemento_por_repeticao(driver,"/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/div[3]/div/button[2]","click","tramitar",4)
            # encontrar_elemento_por_repeticao(driver, "/html/body/div[4]/div[3]/div/div[2]/ul/div[1]", "click", "BAIXADOS", 4)
           
            atualizar_status_na_planilha(int(solicitacao[4]))
            sleep(5)
            driver.get("https://tpf2.madrix.app/runtime/44/list/184/Solicitação de Adiantamento")
            sleep(2)






    print("FIMMMMMMMMMMMMMMM ADIANTAMENTO")
    # driver.close()