from time import sleep
from datetime import date
from gerenciadorPastas import *
from funcoes import *
import shutil
import os
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains 
from openpyxl import load_workbook
from gerenciadorPlanilhas import preencher_solicitacao_na_planilha, ler_dados_da_planilha, atualizar_status_na_planilha


def aporte(driver):
    #verificar se tem downloads antigos e apagar
    remover_arquivos_da_raiz(recuperar_diretorio_usuario() + "\\tpfe.com.br\\SGP e SGC - RPA\\")
    #data atual formatada
    data_em_texto = date.today().strftime("%d.%m.%Y")
    #caminho da pasta macro(pasta do dia)
    caminho_da_pasta = recuperar_diretorio_usuario() + "\\tpfe.com.br\\SGP e SGC - RPA\\Solicitação de Aporte\\" 
    #criar pasta do dia dentro de pagamento avulso
    criarPastaData(caminho_da_pasta, data_em_texto)
    tipo_de_solicitacao = "AP"
    builder = ActionChains(driver)
    driver.implicitly_wait(2)

    #ACESSANDO SOLICITAÇAO DE APORTE
    espera_explicita_de_elemento(driver,"/html/body/div[1]/div/div[2]/main/section/div/div/div/div/section/div/div[2]/div","encontrar","Encontrar o módulo",2)
    driver.get("https://tpf2.madrix.app/runtime/44/list/220/Solicitação de Aporte")
    #driver.implicitly_wait(10)

    #Limpando Filtros
    encontrar_elemento_por_repeticao(driver,"/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/button[2]","click","1-filtro", 4)
    encontrar_elemento_por_repeticao(driver,"/html/body/div[4]/div[3]/div/div[1]/div[1]/button","click","limpar", 4)
    encontrar_elemento_por_repeticao(driver,"/html/body/div[4]/div[3]/div/div[2]/button","click","sair", 4)

    
    

    #FILTRANDO AS SOLICITAÇÕES APROVADAS PELO GERENTE
    encontrar_elemento_por_repeticao(driver,"/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/div/div/div","click","1-filtro", 4)
    encontrar_elemento_por_repeticao(driver,"/html/body/div[4]/div[3]/ul/li[3]","click","todas",4)
    encontrar_elemento_por_repeticao(driver,"//*[@id='mainContent']/section/div/div/div/div[1]/div/div[1]/button[2]","click","2-filtro", 4)
    encontrar_elemento_por_repeticao(driver,"/html/body/div[4]/div[3]/div/div[1]/div[1]/button","click","limpar filtro",4)
    encontrar_elemento_por_repeticao(driver, "/html/body/div[4]/div[3]/div/ul/li[4]/div/div/div/div", "click", "selecionar caixa",4)
    encontrar_elemento_por_repeticao(driver,"/html/body/div[5]/div[3]/ul/li[3]", "click", "solicitado", 4)
    encontrar_elemento_por_repeticao(driver,"/html/body/div[5]/div[1]", "click", "clicar fora", 4)
    sleep(4)
    encontrar_elemento_por_repeticao(driver, "/html/body/div[4]/div[3]/div/div[2]/button", "click", "aplicar", 4)
    print(20*"=")

    #OBTER QUANTIDADE DE PAGAMENTOS
    sleep(3)
    quantidade_de_requisicoes = int((driver.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/span[2]/div/p[2]").get_attribute("innerText")).split(" ")[-1])
    
    #LAÇO PARA TRAMITAR TODOS OS PAGAMENTOS
    for linha in range(3): #voltar para antigo quantidades
        dados_do_formulario = []
        global identificador
        #armazenando o id de cada solicitaçao
        identificador = driver.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[3]/div/div/div/table/tbody/tr[1]/td[4]/div").get_attribute("innerText")
        #armazenando a razao social de cada solicitaçao
        produto = driver.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[3]/div/div/div/table/tbody/tr[1]/td[6]/div").get_attribute("innerText")
        #ACESSANDO A SOLICITAÇAO
        sleep(3)
        try:
            driver.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[3]/div/div/div/table/tbody/tr[1]/td[2]/span/span[1]/input").click()
        except:
            driver.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[3]/div/div/div/table/tbody/tr[1]/td[2]/span/span[1]/input").send_keys("\n")
        #clicar no lápis de edição
        encontrar_elemento_por_repeticao(driver, "/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/div[3]/div/button[1]", "click", "editar", 4)
        
        #PEGAR TODAS AS INFORMAÇOES PARA ALIMENTAR A PLANILHA
        sleep(3)
        #AP
        dados_do_formulario.append(tipo_de_solicitacao)
        #ID DA SOLICITAÇAO
        dados_do_formulario.append(identificador)
        #CPF/CNPJ
        dados_do_formulario.append(driver.find_element_by_xpath("/html/body/div[4]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[2]/div/div[2]/div[2]/div/div/div[1]/div[1]/div/div[1]/div/div/div/input").get_attribute("value"))
        #RAZÃO SOCIAL
        dados_do_formulario.append("")
        #FORMA DE PAGAMENTO
        dados_do_formulario.append("Tranferência Bancária")
        #BANCO
        dados_do_formulario.append(driver.find_element_by_xpath("/html/body/div[4]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[2]/div/div[2]/div[2]/div/div/div[1]/div[1]/div/div[2]/div/div/div/input").get_attribute("value"))
        #AGENCIA
        dados_do_formulario.append(driver.find_element_by_xpath("/html/body/div[4]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[2]/div/div[2]/div[2]/div/div/div[1]/div[2]/div/div[1]/div/div/div/input").get_attribute("value"))
        #CONTA
        dados_do_formulario.append(driver.find_element_by_xpath("/html/body/div[4]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[2]/div/div[2]/div[2]/div/div/div[1]/div[2]/div/div[2]/div/div/div/input").get_attribute("value"))
        #TIPO DE CONTA
        dados_do_formulario.append(driver.find_element_by_xpath("/html/body/div[4]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[2]/div/div[2]/div[2]/div/div/div[1]/div[2]/div/div[3]/div/div/div/input").get_attribute("value"))
        #NATUREZA DA CONTA
        dados_do_formulario.append("")
        #VALOR
        valor = driver.find_element_by_xpath("/html/body/div[4]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[2]/div/div[2]/div[2]/div/div/div[2]/div[1]/div/div[1]/div/div/div/input").get_attribute("value")
        dados_do_formulario.append(valor)
        #VALOR PAGO
        dados_do_formulario.append("")
        #DATA SOLICITADA PARA PAGAMENTO
        dados_do_formulario.append("")
        #DATA DA SOLICITAÇÃO 
        dados_do_formulario.append(driver.find_element_by_xpath("/html/body/div[4]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[2]/div/div[2]/div[2]/div/div/div[2]/div[1]/div/div[2]/div/div/div/input").get_attribute("value"))
        #DATA DE PAGAMENTO 
        dados_do_formulario.append("")

        valor_da_conta = valor.replace(".","")
        valor_da_conta = valor_da_conta.replace("R$","")
        valor_da_conta = valor_da_conta.replace(",",".")
        valor_da_conta = float(valor_da_conta)

        if dados_do_formulario[5] == "" or  dados_do_formulario[6] == "" or  dados_do_formulario[7] == "" or  dados_do_formulario[8] == "" or valor_da_conta == 0:# and  banco != ""
            #Comentario Robo
            dados_do_formulario.append("Dados bancários incompletos ou solicitação está com valor zerado.")
            # tramitar = 1
        else:
            #Comentario Robo 
            dados_do_formulario.append("")
        #Ajuste
        dados_do_formulario.append("")

        #CRIAR A PASTA DO PAGAMENTO QUE ACABOU DE SER PROCESSADO
        produto = produto.replace("\\" , "")
        produto = produto.replace("/", "")
        produto = produto.replace(":", "")
        produto = produto.replace("*", "")
        produto = produto.replace("?", "")
        produto = produto.replace('"', "")
        produto = produto.replace("<", "")
        produto = produto.replace(">", "")
        produto = produto.replace("|", "")
        produto = produto.replace(".", "")
        produto = produto.replace(",", "")

        #nome_da_pasta = (f"ID SA{identificador}")
        nome_da_pasta = "ID SA" + str(identificador) + " " + str(produto)

        sleep(2)
        criarPastasFilhas("Solicitação de Aporte", nome_da_pasta)
        

        #BAIXAR NOTA FISCAL
        encontrar_elemento_por_repeticao(driver, "/html/body/div[4]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[1]/div/div[2]/div/button[2]", "click", "clicar em notas", 3)
        tbody2 = driver.find_element_by_xpath("/html/body/div[4]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[3]/div/div/div/div[1]/div[3]/table/tbody")
        #pega todas as linhas que contem nf 
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
        #IMPRIMINDO
        try: 
            encontrar_elemento_por_repeticao(driver, "/html/body/div[4]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[1]/div/div[2]/div/button[1]", "click", "voltar", 10)
            encontrar_elemento_por_repeticao(driver, "/html/body/div[4]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[2]/div/div[4]/div[2]/div/div/button", "click", "imprimir", 3)
            driver.switch_to_frame(0)
            encontrar_elemento_por_repeticao(driver,"/html/body/div/div/div/div[2]/div/table/tbody/tr/td[1]/table/tbody/tr/td[3]/div/table/tbody/tr","click","filtro",2) 
            encontrar_elemento_por_repeticao(driver, "/html/body/div/div/div/div[16]/div/div[1]/table/tbody/tr/td[2]", "click", "baixar capa", 2)
            encontrar_elemento_por_repeticao(driver, "/html/body/div/div/div/div[20]/div[4]/table/tbody/tr/td[1]/div/table/tbody/tr/td", "click", "baixar capa", 2)
            driver.switch_to.default_content()
            encontrar_elemento_por_repeticao(driver, "/html/body/div[5]/div[3]/div/div[1]/h2/div/div[2]/button", "click", "fcehar relatorio", 2)
            print("passou")
            encontrar_elemento_por_repeticao(driver, "/html/body/div[4]/div[3]/div/div/div/div[1]/div/div[3]/button", "click", "fechar", 2)
        except: 
            dados_do_formulario[15] += "A capa nao foi baixada"

        #MOVENDO ARQUIVOS
        print("mover arquivos")
        validar_download(caminho_da_pasta, data_em_texto, nome_da_pasta)
        sleep(1)

        #TRAMITAÇÃO
        encontrar_elemento_por_repeticao(driver,"/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/div[3]/div/button[3]","click","tramitar",2)
        encontrar_elemento_por_repeticao(driver, "/html/body/div[4]/div[3]/div/div[2]/ul/div", "click", "PROCESSAR", 2)
        #status
        dados_do_formulario.append("Processada")

        #DATA DE EXECUÇÃO ROBO
        dados_do_formulario.append(date.today().strftime("%d/%m/%Y"))

        preencher_solicitacao_na_planilha(dados_do_formulario, tipo_de_solicitacao)
        sleep(5)
        driver.get("https://tpf2.madrix.app/runtime/44/list/220/Solicitação de Aporte")
        sleep(2)

    print("Vai começar a contar")
    sleep(5)

    # 2° PARTE : ESPERANDO DO driver PRA TRAMITAR PRA PAGO 
    lista_de_tramitacao = ler_dados_da_planilha(tipo_de_solicitacao)
    if len(lista_de_tramitacao) > 0:
        #Filtro
        encontrar_elemento_por_repeticao(driver,"/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/div/div/div","click","1-filtro", 4)
        encontrar_elemento_por_repeticao(driver,"/html/body/div[4]/div[3]/ul/li[3]","click","todas",2)
        #Para cada solicitaçao que precisa ser paga
        for solicitacao in lista_de_tramitacao:
            print(lista_de_tramitacao)
            encontrar_elemento_por_repeticao(driver,"/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/button[2]","click","2-filtro", 2)
            encontrar_elemento_por_repeticao(driver,"/html/body/div[4]/div[3]/div/div[1]/div[1]/button","click","limpar",4)
            #Filtrar o ID
            driver.find_element_by_xpath("/html/body/div[4]/div[3]/div/ul/li[1]/div/div/div/div/input").send_keys(str(solicitacao[0]))
            #encontrar_elemento_por_repeticao(driver, "/html/body/div[4]/div[3]/div/ul/li[4]/div/div/div/div", "click", "filtro",0.3)
            #encontrar_elemento_por_repeticao(driver,"/html/body/div[5]/div[3]/ul/li[4]", "click", "processadoss", 0.4)
            #sleep(4)
            #encontrar_elemento_por_repeticao(driver,"/html/body/div[4]/div[1]", "click", "clicar fora", 2)
            
            encontrar_elemento_por_repeticao(driver, "/html/body/div[4]/div[3]/div/div[2]/button", "click", "aplicar", 3)
            sleep(3)
            encontrar_elemento_por_repeticao(driver, "/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[3]/div/div/div/table/tbody/tr/td[2]/span/span[1]/input", "click", "LINHA", 2 )
            encontrar_elemento_por_repeticao(driver, "/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/div[3]/div/button[3]", "click", "LINHA",2 )
            sleep(2)
            #Pago ou parcialmente pago 
            if str(solicitacao[3]) == "Pago":
                encontrar_elemento_por_repeticao(driver,"/html/body/div[4]/div[3]/div/div[2]/ul/div[1]","click","SPA",3)    
            elif str(solicitacao[3]) == "Parcialmente pago":  
                encontrar_elemento_por_repeticao(driver,"/html/body/div[4]/div[3]/div/div[2]/ul/div[2]","click","SPA",3)
            sleep(4)
            
            atualizar_status_na_planilha(int(solicitacao[4]))     
    print("FIMMMMMMMMMMMMMMM Aporte")
    # driver.close()
