from time import sleep
from datetime import date
import gerenciadorPastas
import funcoes
import shutil
import os
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains 
from openpyxl import load_workbook
from gerenciadorPlanilhas import preencher_solicitacao_na_planilha, ler_dados_da_planilha, atualizar_status_na_planilha


def aporte(driver):
    #verificar se tem downloads antigos e apagar
    gerenciadorPastas.remover_arquivos_da_raiz(gerenciadorPastas.recuperar_diretorio_usuario() + "\\tpfe.com.br\\SGP e SGC - RPA\\")
    #data atual formatada
    data_em_texto = date.today().strftime("%d.%m.%Y")
    #caminho da pasta macro(pasta do dia)
    caminho_da_pasta = gerenciadorPastas.recuperar_diretorio_usuario() + "\\tpfe.com.br\\SGP e SGC - RPA\\Solicitação de Aporte\\" 
    #criar pasta do dia dentro de pagamento avulso
    gerenciadorPastas.criarPastaData(caminho_da_pasta, data_em_texto)
    tipo_de_solicitacao = "SA"
    builder = ActionChains(driver)
    driver.implicitly_wait(2)

    #ACESSANDO SOLICITAÇAO DE APORTE
    funcoes.espera_explicita_de_elemento(driver,"/html/body/div[1]/div/div[2]/main/section/div/div/div/div/section/div/div[2]/div","encontrar","AD",2)
    driver.get("https://tpf2.madrix.app/runtime/44/list/220/Solicitação de Aporte")    #driver.implicitly_wait(10)

    #FILTRANDO AS SOLICITAÇÕES APROVADAS PELO GERENTE
    funcoes.encontrar_elemento_por_repeticao(driver,"/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/div/div/div","click","filtro", 2)
    funcoes.encontrar_elemento_por_repeticao(driver,"/html/body/div[4]/div[3]/ul/li[3]","click","filtro",0.4)
    funcoes.encontrar_elemento_por_repeticao(driver,"//*[@id='mainContent']/section/div/div/div/div[1]/div/div[1]/button[2]","click","filtro", 2)
    funcoes.encontrar_elemento_por_repeticao(driver,"/html/body/div[4]/div[3]/div/div[1]/div[1]/button","click","filtro",0.2)
    funcoes.encontrar_elemento_por_repeticao(driver, "/html/body/div[4]/div[3]/div/ul/li[4]/div/div/div/div", "click", "filtro",0.3)
    funcoes.encontrar_elemento_por_repeticao(driver,"/html/body/div[5]/div[3]/ul/li[3]", "click", "solicitado", 0.4)
    funcoes.encontrar_elemento_por_repeticao(driver,"/html/body/div[5]/div[1]", "click", "clicar fora", 0.5)
    funcoes.encontrar_elemento_por_repeticao(driver, "/html/body/div[4]/div[3]/div/div[2]/button", "click", "aplicar", 0.3)
    print(20*"=")

    #OBTER QUANTIDADE DE PAGAMENTOS
    sleep(5)
    quantidade_de_requisicoes = int((driver.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/span[2]/div/p").get_attribute("innerText")).split(" ")[-1])
    
    #LAÇO PARA TRAMITAR TODOS OS PAGAMENTOS
    for linha in range(2): #voltar para antigo quantidades
        dados_do_formulario = []
        global identificador
        #armazenando o id de cada solicitaçao
        identificador = driver.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[3]/div/div/div/table/tbody/tr[1]/td[4]/div").get_attribute("innerText")
        #armazenando a razao social de cada solicitaçao
        #estado = driver.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[3]/div/div/div/table/tbody/tr[1]/td[7]/div/span").get_attribute("innerText")
        #ACESSANDO A SOLICITAÇAO
        driver.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[3]/div/div/div/table/tbody/tr[1]/td[2]/span/span[1]/input").click()
        #clicar no lápis de edição
        funcoes.encontrar_elemento_por_repeticao(driver, "/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/div[3]/div/button[1]", "click", "click na linha", 4)
        
        #PEGAR TODAS AS INFORMAÇOES PARA ALIMENTAR A PLANILHA
        sleep(3)
        #SA
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
        dados_do_formulario.append("0")
        #DATA SOLICITADA PARA PAGAMENTO
        dados_do_formulario.append("")
        #DATA DA SOLICITAÇÃO 
        dados_do_formulario.append("")
        #DATA DE PAGAMENTO 
        dados_do_formulario.append(driver.find_element_by_xpath("/html/body/div[4]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[2]/div/div[2]/div[2]/div/div/div[2]/div[1]/div/div[2]/div/div/div/input").get_attribute("value"))


        valor_da_conta = valor.replace(".","")
        valor_da_conta = valor_da_conta.replace("R$","")
        valor_da_conta = valor_da_conta.replace(",",".")
        valor_da_conta = float(valor_da_conta)

        if dados_do_formulario[5] == "" or  dados_do_formulario[6] == "" or  dados_do_formulario[7] == "" or  dados_do_formulario[8] == "" or  dados_do_formulario[9] == "" or valor_da_conta == 0:# and  banco != ""
            #Comentario Robo
            dados_do_formulario.append("Dados bancários incompletos ou solicitação está com valor zerado.")
            # tramitar = 1
        else:
            #Comentario Robo 
            dados_do_formulario.append("")
        #Ajuste
        dados_do_formulario.append("")

        #CRIAR A PASTA DO PAGAMENTO QUE ACABOU DE SER PROCESSADO

        nome_da_pasta = (f"SA ID {identificador}")

        sleep(2)
        gerenciadorPastas.criarPastasFilhas("Solicitação de Aporte", nome_da_pasta)
        
        sleep(3)

        #BAIXAR NOTA FISCAL
        funcoes.encontrar_elemento_por_repeticao(driver, "/html/body/div[4]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[1]/div/div[2]/div/button[2]", "click", "clicar em notas", 3)
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
                        sleep(3)
                        maximo_tentativas = 40
                    else:
                        maximo_tentativas += 1
        except:
            comentario_nota_fiscal = (f"Não foi possível baixar a nota fiscal")     
            print(comentario_nota_fiscal)
            dados_do_formulario[15] = comentario_nota_fiscal 

        sleep(5)
        #IMPRIMINDO
        try: 
            funcoes.encontrar_elemento_por_repeticao(driver, "/html/body/div[4]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[1]/div/div[2]/div/button[1]", "click", "voltar", 10)
            funcoes.encontrar_elemento_por_repeticao(driver, "/html/body/div[4]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[2]/div/div[4]/div[2]/div/div/button", "click", "imprimir", 3)
            driver.switch_to_frame(0)
            funcoes.encontrar_elemento_por_repeticao(driver,"/html/body/div/div/div/div[2]/div/table/tbody/tr/td[1]/table/tbody/tr/td[3]/div/table/tbody/tr","click","filtro",2) 
            funcoes.encontrar_elemento_por_repeticao(driver, "/html/body/div/div/div/div[16]/div/div[1]/table/tbody/tr/td[2]", "click", "baixar capa", 2)
            funcoes.encontrar_elemento_por_repeticao(driver, "/html/body/div/div/div/div[20]/div[4]/table/tbody/tr/td[1]/div/table/tbody/tr/td", "click", "baixar capa", 2)
            driver.switch_to.default_content()
            funcoes.encontrar_elemento_por_repeticao(driver, "/html/body/div[5]/div[3]/div/div[1]/h2/div/div[2]/button", "click", "baixar capa", 2)
            print("passou")
            funcoes.encontrar_elemento_por_repeticao(driver, "/html/body/div[4]/div[3]/div/div/div/div[1]/div/div[3]/button", "click", "baixar capa", 2)
        except: 
            dados_do_formulario[15] += "A capa nao foi baixada"

        #MOVENDO ARQUIVOS
        print("mover arquivos")
        funcoes.validar_download(caminho_da_pasta, data_em_texto, nome_da_pasta)
        sleep(3)

        #TRAMITAÇÃO
        funcoes.encontrar_elemento_por_repeticao(driver,"/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/div[3]/div/button[3]","click","tramitar",2)
        funcoes.encontrar_elemento_por_repeticao(driver, "/html/body/div[4]/div[3]/div/div[2]/ul/div", "click", "PROCESSAR", 2)
        #status
        dados_do_formulario.append("Processada")

        #DATA DE EXECUÇÃO ROBO
        dados_do_formulario.append(date.today().strftime("%d/%m/%Y"))

        preencher_solicitacao_na_planilha(dados_do_formulario, tipo_de_solicitacao)
        sleep(5)

    sleep(1.5)
    print("Vai começar a contar")
    sleep(5)

    # 2° PARTE : ESPERANDO DO driver PRA TRAMITAR PRA PAGO 
    lista_de_tramitacao = ler_dados_da_planilha(tipo_de_solicitacao)
    if len(lista_de_tramitacao) > 0:
        #Filtro
        funcoes.encontrar_elemento_por_repeticao(driver,"/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/div/div/div","click","filtro", 4)
        funcoes.encontrar_elemento_por_repeticao(driver,"/html/body/div[4]/div[3]/ul/li[3]","click","filtro",2)
        #Para cada solicitaçao que precisa ser paga
        for solicitacao in lista_de_tramitacao:
            print(lista_de_tramitacao)
            funcoes.encontrar_elemento_por_repeticao(driver,"/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/button[3]","click","filtro", 2)
            funcoes.encontrar_elemento_por_repeticao(driver,"/html/body/div[5]/div[3]/div/div[1]/div[1]/button","click","filtro",0.2)
            #Filtrar o ID
            driver.find_element_by_xpath("/html/body/div[4]/div[3]/div/ul/li[1]/div/div/div/div/input").send_keys(str(solicitacao[0]))
            funcoes.encontrar_elemento_por_repeticao(driver, "/html/body/div[4]/div[3]/div/ul/li[4]/div/div/div/div", "click", "filtro",0.3)
            funcoes.encontrar_elemento_por_repeticao(driver,"/html/body/div[5]/div[3]/ul/li[4]", "click", "processadoss", 0.4)
            funcoes.encontrar_elemento_por_repeticao(driver,"/html/body/div[5]/div[1]", "click", "clicar fora", 0.5)
            funcoes.encontrar_elemento_por_repeticao(driver, "/html/body/div[4]/div[3]/div/div[2]/button", "click", "aplicar", 0.3)
            sleep(3)
            funcoes.encontrar_elemento_por_repeticao(driver, "/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[3]/div/div/div/table/tbody/tr/td[2]/span/span[1]/input", "click", "LINHA", 2 )
            funcoes.encontrar_elemento_por_repeticao(driver, "/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/div[3]/div/button[3]", "click", "LINHA",2 )
            sleep(2)
            #Pago ou parcialmente pago 
            if str(solicitacao[3]) == "PAGO":
                funcoes.encontrar_elemento_por_repeticao(driver,"/html/body/div[4]/div[3]/div/div[2]/ul/div[1]","click","SPA",2)    
            elif str(solicitacao[3]) == "PARCIALMENTE PAGO":  
                funcoes.encontrar_elemento_por_repeticao(driver,"/html/body/div[4]/div[3]/div/div[2]/ul/div[2]","click","SPA",2)
            sleep(2)
            
            atualizar_status_na_planilha(int(solicitacao[4]))     
    print("FIMMMMMMMMMMMMMMM AVULSO")
    # driver.close()
