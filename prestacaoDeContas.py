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



def prestacao_de_contas(driver):
    #verificar se tem downloads antigos e apagar
    remover_arquivos_da_raiz(recuperar_diretorio_usuario() + "\\tpfe.com.br\\SGP e SGC - RPA\\")
    #data atual formatada
    data_em_texto = date.today().strftime("%d.%m.%Y")
    #caminho da pasta macro(pasta do dia)
    caminho_da_pasta = recuperar_diretorio_usuario() + "\\tpfe.com.br\\SGP e SGC - RPA\\Prestação de Contas\\" 
    #criar pasta do dia dentro de pagamento avulso
    criarPastaData(caminho_da_pasta, data_em_texto)
    tipo_de_solicitacao = "PC"
    builder = ActionChains(driver)
    driver.implicitly_wait(2)

    #ACESSANDO PRESTACAO DE CONTAS - CARTAO CORPORATIVO
    # espera_explicita_de_elemento(driver,"/html/body/div[1]/div/div[2]/main/section/div/div/div/div/section/div/div[2]/div","encontrar","PC",2)
    driver.get("https://tpf2.madrix.app/runtime/44/list/221/Prestação de Contas - Cartão Corporativo")
    #driver.implicitly_wait(10)

    #FILTRANDO AS PRESTAÇÕES REALIZADAS
    encontrar_elemento_por_repeticao(driver,"/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/div/div/div","click","1-filtro", 3)
    encontrar_elemento_por_repeticao(driver,"/html/body/div[4]/div[3]/ul/li[2]","click","todas",4)
    encontrar_elemento_por_repeticao(driver,"/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/button[3]","click","2-filtro", 3)
    encontrar_elemento_por_repeticao(driver,"/html/body/div[4]/div[3]/div/div[1]/div[1]/button","click","limpar filtro",3)
    encontrar_elemento_por_repeticao(driver, "/html/body/div[4]/div[3]/div/ul/li[3]/div/div/div/div", "click", "novo filtro",3)
    encontrar_elemento_por_repeticao(driver,"/html/body/div[5]/div[3]/ul/li[3]", "click", "prestaçao realizada", 3)
    encontrar_elemento_por_repeticao(driver,"/html/body/div[5]/div[1]", "click", "clicar fora", 3)
    print(20*"=")
    encontrar_elemento_por_repeticao(driver,"/html/body/div[4]/div[3]/div/div[2]/button", "click", "fechando filtro", 0.4)

    #OBTER QUANTIDADE DE PRESTAÇÕES REALIZADAS
    sleep(5)
    # quantidade_de_requisicoes = int((driver.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/span/div/p").get_attribute("innerText")).split(" ")[-1])
    
    #LAÇO PARA TRAMITAR TODOS AS PRESTAÇÕES
    for linha in range(5):
        print(linha) #voltar para antigo quantidades
        dados_do_formulario = []
        #armazenando o id de cada prestação
        identificador = driver.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[3]/div/div/div/table/tbody/tr[1]/td[4]/div").get_attribute("innerText")
        solicitante = driver.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[3]/div/div/div/table/tbody/tr[1]/td[5]/div/div[2]").get_attribute("innerText")
        #ACESSANDO A PRESTAÇÃO
        sleep(3)
        #driver.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[3]/div/div/div/table/tbody/tr[1]/td[2]/span/span[1]/input").click()
        try:
            driver.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[3]/div/div/div/table/tbody/tr[1]/td[2]/span/span[1]/input").click()
        except:
            driver.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[3]/div/div/div/table/tbody/tr[1]/td[2]/span/span[1]/input").send_keys("\n")
        sleep(3)
        encontrar_elemento_por_repeticao(driver, "/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/div[3]/div/button[1]", "click", "click na linha", 4)
        
        sleep(5)
        #PEGAR TODAS AS INFORMAÇOES PARA ALIMENTAR A PLANILHA
        caminho_em_comum_entre_campos_do_formulario = "/html/body/div[4]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[2]/div/div"
        #PC
        dados_do_formulario.append(tipo_de_solicitacao)
        #ID DA SOLICITAÇAO
        dados_do_formulario.append(identificador)
        #CPF/CNPJ
        dados_do_formulario.append("")
        #RAZAO
        dados_do_formulario.append("")
        #FORMA DE PAGAMENTO
        dados_do_formulario.append("")
        #BANCO
        dados_do_formulario.append("")
        #AGENCIA
        dados_do_formulario.append("")
        #CONTA
        dados_do_formulario.append("")
        #TIPO DE CONTA
        dados_do_formulario.append("")
        #NATUREZA DA CONTA
        dados_do_formulario.append("")
        #VALOR
        valor = driver.find_element_by_xpath(caminho_em_comum_entre_campos_do_formulario + "[2]/div[1]/div[2]/div/div/input").get_attribute("value")
        dados_do_formulario.append(valor)
        #VALOR PAGO
        dados_do_formulario.append("")
        #DATA SOLICITADA PARA PAGAMENTO
        dados_do_formulario.append("")
        #DATA DA SOLICITAÇÃO
        dados_do_formulario.append(driver.find_element_by_xpath(caminho_em_comum_entre_campos_do_formulario + "[2]/div[2]/div[2]/div/div/input").get_attribute("value"))
        #DATA DE PAGAMENTO 
        dados_do_formulario.append("")

        valor_da_conta = valor.replace(".","")
        valor_da_conta = valor_da_conta.replace("R$","")
        valor_da_conta = valor_da_conta.replace(",",".")
        valor_da_conta = float(valor_da_conta)

        if valor_da_conta == 0:# and  banco != ""
            #Comentario Robo
            dados_do_formulario.append("Solicitação está com valor zerado.")
            # tramitar = 1
        else:
            #Comentario Robo 
            dados_do_formulario.append("")
        #Ajuste
        dados_do_formulario.append("")
        
        # #BAIXANDO AS NOTAS FISCAIS
        # #BAIXANDO AS NOTAS FISCAIS
        encontrar_elemento_por_repeticao(driver, "/html/body/div[4]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[1]/div/div[2]/div/button[2]", "click", "clicar em notas", 3)
        tbody2 = driver.find_element_by_xpath("/html/body/div[4]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[3]/div/div/div/div[1]/div[3]/table/tbody")
        # #pega todas as linhas que contem nf
        rows2 = tbody2.find_elements_by_tag_name("a") 
       
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
 
        sleep(3)
        #IMPRIMINDO E BAIXANDO CAPA
        encontrar_elemento_por_repeticao(driver, "/html/body/div[4]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[1]/div/div[2]/div/button[1]", "click", "voltar", 10)
        encontrar_elemento_por_repeticao(driver, "/html/body/div[4]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[2]/div/div[5]/div[2]/div/div/button", "click", "imprimir", 5)
        driver.switch_to_frame(0)
        encontrar_elemento_por_repeticao(driver,"/html/body/div/div/div/div[2]/div/table/tbody/tr/td[1]/table/tbody/tr/td[3]/div/table/tbody/tr","click","filtro-impressao",2) 
        encontrar_elemento_por_repeticao(driver, "/html/body/div/div/div/div[16]/div/div[1]/table/tbody/tr/td[2]", "click", "baixar capa", 2)
        encontrar_elemento_por_repeticao(driver, "/html/body/div/div/div/div[20]/div[4]/table/tbody/tr/td[1]/div/table/tbody/tr/td", "click", "baixar capa", 2)
        driver.switch_to.default_content()
        encontrar_elemento_por_repeticao(driver, "/html/body/div[7]/div[3]/div/div[1]/h2/div/div[2]/button", "click", "baixar capa", 2)
        print("passou")
        encontrar_elemento_por_repeticao(driver, "/html/body/div[4]/div[3]/div/div/div/div[1]/div/div[3]/button", "click", "fechar", 2)
    
        # CRIAR A PASTA DO PAGAMENTO QUE ACABOU DE SER PROCESSADO
    
        nome_da_pasta = "ID PC" + str(identificador) + " " + str(solicitante)

        print(nome_da_pasta)
        criarPastasFilhas("Prestação de Contas", nome_da_pasta) 

        print("mover arquivos")
        sleep(3)
        # MOVENDO A CAPA PARA SUA RESPECTIVA PASTA
        validar_download(caminho_da_pasta, data_em_texto, nome_da_pasta)
        sleep(3)

        #TRAMITAR PARA "PROCESSADA"
        encontrar_elemento_por_repeticao(driver,"/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/div[3]/div/button[4]","click","tramitar",2)
        encontrar_elemento_por_repeticao(driver, "/html/body/div[4]/div[3]/div/div[2]/ul/div", "click", "processar", 2)
        #status
        dados_do_formulario.append("Processada")
        #DATA EXECUÇAO ROBO
        dados_do_formulario.append(date.today().strftime("%d/%m/%Y"))
        
        preencher_solicitacao_na_planilha(dados_do_formulario, tipo_de_solicitacao)
        sleep(5)
        driver.get("https://tpf2.madrix.app/runtime/44/list/221/Prestação de Contas - Cartão Corporativo")
        sleep(2)

    sleep(5)
    print("FIMMMMMMMMMMMMMMM PC-CARTAO")
    # driver.close()

    # lista_de_tramitacao = ler_dados_da_planilha(tipo_de_solicitacao)
    # if len(lista_de_tramitacao) > 0:
    #     encontrar_elemento_por_repeticao(driver,"/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/div/div/div","click","filtro", 4)
    #     encontrar_elemento_por_repeticao(driver,"/html/body/div[5]/div[3]/ul/li[6]","click","filtro",2)
    #     for solicitacao in lista_de_tramitacao:
    #         print(solicitacao)
    #         encontrar_elemento_por_repeticao(driver,"/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/button[3]","click","filtro", 2)
    #         encontrar_elemento_por_repeticao(driver,"/html/body/div[5]/div[3]/div/div[1]/div[1]/button","click","filtro",0.2)
    #         driver.find_element_by_xpath("/html/body/div[5]/div[3]/div/ul/li[1]/div/div/div/div/input").send_keys(ler_dados_da_planilha(str(solicitacao[0]))
    #         # encontrar_elemento_por_repeticao(driver, "/html/body/div[5]/div[3]/div/ul/li[3]/div/div/div/div", "click", "filtro", 0.4)
    #         # encontrar_elemento_por_repeticao(driver, "/html/body/div[6]/div[3]/ul/li[7]", "click", "filtro", 0.4)
    #         encontrar_elemento_por_repeticao(driver, "/html/body/div[6]/div[1]", "click", "filtro", 0.4 )
    #         print(20*"=")
    #         encontrar_elemento_por_repeticao(driver, "/html/body/div[5]/div[3]/div/div[2]/button", "click", "filtro", 0.4 )
    #         sleep(3)
    #         encontrar_elemento_por_repeticao(driver, "/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[3]/div/div/div/table/tbody/tr/td[2]/PDCn/PDCn[1]/input", "click", "LINHA", 2 )
    #         encontrar_elemento_por_repeticao(driver, "/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/div[3]/div/button[1]", "click", "LINHA",2 )
    #         encontrar_elemento_por_repeticao(driver, "/html/body/div[5]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[1]/div/div[2]/div/button[3]", "click", "LINHA", 2 )
    #         sleep(2)
    #         encontrar_elemento_por_repeticao(driver,"/html/body/div[5]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[4]/div/div/div/div[1]/div[1]/div[1]/div/div/PDCn/div/button[1]","click","PDC",2)
    #         sleep(3)
    #         driver.find_element_by_xpath("/html/body/div[8]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div/div[1]/div[1]/div/div/div/input").send_keys(ler_dados_da_planilha(tipo_de_solicitacao)[0][2])
    #         driver.find_element_by_xpath("/html/body/div[8]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div/div[1]/div[2]/div/div/div/input").send_keys(ler_dados_da_planilha(tipo_de_solicitacao)[0][1])
    #         encontrar_elemento_por_repeticao(driver,"/html/body/div[8]/div[3]/div/div/div/div[4]/fieldset/button[2]","click","PDC",3)
    #         encontrar_elemento_por_repeticao(driver,"/html/body/div[5]/div[3]/div/div/div/div[1]/div/div[3]/button","click","PDC",2)
    #         encontrar_elemento_por_repeticao(driver,"/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/div[3]/div/button[2]","click","PDC",3)
    #         encontrar_elemento_por_repeticao(driver,"/html/body/div[5]/div[3]/div/div[2]/ul/div[1]","click","PDC",2) 
    #         sleep(3)    
    #         if ler_dados_da_planilha(tipo_de_solicitacao)[0][3] == "PAGO":
    #             encontrar_elemento_por_repeticao(driver,"/html/body/div[5]/div[3]/div/div[2]/ul/div[1]","click","tramitar pago",2)    
    #         elif ler_dados_da_planilha(tipo_de_solicitacao)[0][3] == "PARCIALMENTE PAGO":  
    #             encontrar_elemento_por_repeticao(driver,"/html/body/div[5]/div[3]/div/div[2]/ul/div[2]","click","tramitar parcialmente",2)
    #         sleep(2)
    #         atualizar_status_na_planilha(ler_dados_da_planilha(tipo_de_solicitacao)[0][4])    
    # driver.close()    
    print("FIMMMMMMMMMMMMMMM")
    
