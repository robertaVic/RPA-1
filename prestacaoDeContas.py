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



def prestacao_de_contas(driver):
    #verificar se tem downloads antigos e apagar
    gerenciadorPastas.remover_arquivos_da_raiz(gerenciadorPastas.recuperar_diretorio_usuario() + "\\tpfe.com.br\\SGP e SGC - RPA\\")
    #data atual formatada
    data_em_texto = date.today().strftime("%d.%m.%Y")
    #caminho da pasta macro(pasta do dia)
    caminho_da_pasta = gerenciadorPastas.recuperar_diretorio_usuario() + "\\tpfe.com.br\\SGP e SGC - RPA\\Prestação de Contas\\" 
    #criar pasta do dia dentro de pagamento avulso
    #gerenciadorPastas.criarPastaData(caminho_da_pasta, data_em_texto)
    tipo_de_solicitacao = "PC"
    builder = ActionChains(driver)
    driver.implicitly_wait(2)

    #ACESSANDO PRESTACAO DE CONTAS - CARTAO CORPORATIVO
    funcoes.espera_explicita_de_elemento(driver,"/html/body/div[1]/div/div[2]/main/section/div/div/div/div/section/div/div[2]/div","encontrar","PC",2)
    driver.get("https://tpf2.madrix.app/runtime/44/list/221/Prestação de Contas - Cartão Corporativo")
    driver.implicitly_wait(10)

    #FILTRANDO AS PRESTAÇÕES REALIZADAS
    funcoes.encontrar_elemento_por_repeticao(driver,"/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/div/div/div","click","filtro", 3)
    funcoes.encontrar_elemento_por_repeticao(driver,"/html/body/div[4]/div[3]/ul/li[6]","click","filtro",2)
    funcoes.encontrar_elemento_por_repeticao(driver,"/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/button[3]","click","filtro", 3)
    funcoes.encontrar_elemento_por_repeticao(driver,"/html/body/div[4]/div[3]/div/div[1]/div[1]/button","click","filtro",0.2)
    funcoes.encontrar_elemento_por_repeticao(driver, "/html/body/div[4]/div[3]/div/ul/li[3]/div/div/div/div", "click", "filtro",2)
    funcoes.encontrar_elemento_por_repeticao(driver,"/html/body/div[5]/div[3]/ul/li[3]", "click", "filtro", 0.4)
    funcoes.encontrar_elemento_por_repeticao(driver,"/html/body/div[5]/div[1]", "click", "filtro", 0.3)
    print(20*"=")
    funcoes.encontrar_elemento_por_repeticao(driver,"/html/body/div[4]/div[3]/div/div[2]/button", "click", "fechando filtro", 0.4)

    #OBTER QUANTIDADE DE PRESTAÇÕES REALIZADAS
    sleep(5)
    quantidade_de_requisicoes = int((driver.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/span/div/p").get_attribute("innerText")).split(" ")[-1])
    
    #LAÇO PARA TRAMITAR TODOS AS PRESTAÇÕES
    for linha in range(2): #voltar para antigo quantidades
        dados_do_formulario = []
        #armazenando o id de cada prestação
        identificador = driver.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[3]/div/div/div/table/tbody/tr[1]/td[4]/div").get_attribute("innerText")
        #estado = driver.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[3]/div/div/div/table/tbody/tr[1]/td[7]/div/span").get_attribute("innerText")
        #ACESSANDO A PRESTAÇÃO
        driver.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[3]/div/div/div/table/tbody/tr[1]/td[2]/span/span[1]/input").click()
        funcoes.encontrar_elemento_por_repeticao(driver, "/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/div[3]/div/button[1]", "click", "click na linha", 4)
        
        sleep(3)
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
        dados_do_formulario.append("0")
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
        # funcoes.encontrar_elemento_por_repeticao(driver, "/html/body/div[5]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[1]/div/div[2]/div/button[2]", "click", "clicar em notas", 3)
        # tbody2 = driver.find_element_by_xpath("/html/body/div[5]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[3]/div/div/div/div[1]/div[3]/table/tbody")
        # #pega todas as linhas que contem nf
        # rows2 = tbody2.find_elements_by_tag_name("a") 
       
        # try:
        #     if len(rows2) > 0:
        #         for row in rows2:
        #             row.click()
        #             sleep(5)
        #     # else:
        #     #     comentario_nao_possui_nota = (f"A prestação não possui notas fiscais para serem baixadas")      
        #     #     print(comentario_nao_possui_nota)  
        # except:
        #     comentario_nota_fiscal = (f"Não foi possível baixar a nota fiscal")     
        #     print(comentario_nota_fiscal) 
 
        # sleep(10)
        # #IMPRIMINDO E BAIXANDO CAPA
        # funcoes.encontrar_elemento_por_repeticao(driver, "/html/body/div[5]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[1]/div/div[2]/div/button[1]", "click", "voltar", 10)
        # funcoes.encontrar_elemento_por_repeticao(driver, "/html/body/div[5]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[2]/div/div[6]/div[2]/div/div[2]/div/div/button", "click", "imprimir", 10)
        # driver.switch_to_frame(0)
        # funcoes.encontrar_elemento_por_repeticao(driver,"/html/body/div/div/div/div[2]/div/table/tbody/tr/td[1]/table/tbody/tr/td[3]/div/table/tbody/tr","click","filtro",2) 
        # funcoes.encontrar_elemento_por_repeticao(driver, "/html/body/div/div/div/div[16]/div/div[1]/table/tbody/tr/td[2]", "click", "baixar capa", 2)
        # funcoes.encontrar_elemento_por_repeticao(driver, "/html/body/div/div/div/div[20]/div[4]/table/tbody/tr/td[1]/div/table/tbody/tr/td", "click", "baixar capa", 2)
        # driver.switch_to.default_content()
        # funcoes.encontrar_elemento_por_repeticao(driver, "/html/body/div[8]/div[3]/div/div[1]/h2/div/div[2]/button", "click", "baixar capa", 2)
        # print("passou")
        funcoes.encontrar_elemento_por_repeticao(driver, "/html/body/div[4]/div[3]/div/div/div/div[1]/div/div[3]/button", "click", "baixar capa", 2)
    
        # # if not comentario_nota_fiscal:
        # #     comentario = ("Nenhum")
        # # else:
        # #     comentario = (f"2- {comentario_nota_fiscal}")   
        # #     if not comentario_nao_possui_nota:
        # #         comentario = (f"Nenhum")
        # #     else:
        # #         comentario = (f"3- {comentario_nao_possui_nota}")
        # #         comentario+= comentario, 

        # # print(comentario) 
        
        # #CRIAR A PASTA DO PAGAMENTO QUE ACABOU DE SER PROCESSADO
    
        # nome_da_pasta = (f"ID {identificador}")

        # print(nome_da_pasta)
        # gerenciadorPastas.criarPastasFilhas("Prestação de Contas", nome_da_pasta) 

        # sleep(3)
        # print("mover arquivos")

        # #LISTANDO OS ARQUIVOS BAIXADOS E MOVENDO PARA SUA RESPECTIVA PASTA
        # arquivos = gerenciadorPastas.listar_arquivos_em_diretorios(gerenciadorPastas.recuperar_diretorio_usuario() + "\\tpfe.com.br\\SGP e SGC - RPA")
        # print(arquivos)
        # #criando um laço para mover cada um para sua pasta especifica
        # for arquivo in arquivos:
        #     print(arquivo)
        #     #movendo os arquivos para a pasta da sua solicitaçao
        #     down = os.path.splitext(arquivo)[-1].lower()
        #     maximo_tentativas = 0
        #     while maximo_tentativas <= 20:
        #         if down == ".crdownload":
        #             maximo_tentativas+= 1
        #             sleep(2)
        #         else:
        #             try:
        #                 shutil.move(gerenciadorPastas.recuperar_diretorio_usuario() + "\\tpfe.com.br\\SGP e SGC - RPA\\" + arquivo, caminho_da_pasta + data_em_texto +"\\"+ nome_da_pasta + "\\" + arquivo)
        #                 maximo_tentativas = 11
        #             except:
        #                 maximo_tentativas+= 1
        # sleep(3)

        #TRAMITAR PARA "PROCESSADA"
        funcoes.encontrar_elemento_por_repeticao(driver,"/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/div[3]/div/button[4]","click","tramitar",2)
        funcoes.encontrar_elemento_por_repeticao(driver, "/html/body/div[4]/div[3]/div/div[2]/ul/div", "click", "tramitar", 2)
        #status
        dados_do_formulario.append("Processada")
        #DATA EXECUÇAO ROBO
        dados_do_formulario.append(date.today().strftime("%d/%m/%Y"))
        
        preencher_solicitacao_na_planilha(dados_do_formulario, tipo_de_solicitacao)
        sleep(5)

    sleep(5)
    print("FIMMMMMMMMMMMMMMM PC-CARTAO")
    # driver.close()
    # funcoes.encontrar_elemento_por_repeticao(driver,"/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/div/div/div","click","filtro", 4)
    # funcoes.encontrar_elemento_por_repeticao(driver,"/html/body/div[5]/div[3]/ul/li[6]","click","filtro",2)
    # for i in range(len(ler_dados_da_planilha(tipo_de_solicitacao))):
    #     print(ler_dados_da_planilha(tipo_de_solicitacao))
    #     funcoes.encontrar_elemento_por_repeticao(driver,"/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/button[3]","click","filtro", 2)
    #     funcoes.encontrar_elemento_por_repeticao(driver,"/html/body/div[5]/div[3]/div/div[1]/div[1]/button","click","filtro",0.2)
    #     driver.find_element_by_xpath("/html/body/div[5]/div[3]/div/ul/li[1]/div/div/div/div/input").send_keys(ler_dados_da_planilha(tipo_de_solicitacao)[0][0])
    #     # funcoes.encontrar_elemento_por_repeticao(driver, "/html/body/div[5]/div[3]/div/ul/li[3]/div/div/div/div", "click", "filtro", 0.4)
    #     # funcoes.encontrar_elemento_por_repeticao(driver, "/html/body/div[6]/div[3]/ul/li[7]", "click", "filtro", 0.4)
    #     funcoes.encontrar_elemento_por_repeticao(driver, "/html/body/div[6]/div[1]", "click", "filtro", 0.4 )
    #     print(20*"=")
    #     funcoes.encontrar_elemento_por_repeticao(driver, "/html/body/div[5]/div[3]/div/div[2]/button", "click", "filtro", 0.4 )
    #     sleep(3)
    #     funcoes.encontrar_elemento_por_repeticao(driver, "/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[3]/div/div/div/table/tbody/tr/td[2]/PDCn/PDCn[1]/input", "click", "LINHA", 2 )
    #     funcoes.encontrar_elemento_por_repeticao(driver, "/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/div[3]/div/button[1]", "click", "LINHA",2 )
    #     funcoes.encontrar_elemento_por_repeticao(driver, "/html/body/div[5]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[1]/div/div[2]/div/button[3]", "click", "LINHA", 2 )
    #     sleep(2)
    #     funcoes.encontrar_elemento_por_repeticao(driver,"/html/body/div[5]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[4]/div/div/div/div[1]/div[1]/div[1]/div/div/PDCn/div/button[1]","click","PDC",2)
    #     sleep(3)
    #     driver.find_element_by_xpath("/html/body/div[8]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div/div[1]/div[1]/div/div/div/input").send_keys(ler_dados_da_planilha(tipo_de_solicitacao)[0][2])
    #     driver.find_element_by_xpath("/html/body/div[8]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div/div[1]/div[2]/div/div/div/input").send_keys(ler_dados_da_planilha(tipo_de_solicitacao)[0][1])
    #     funcoes.encontrar_elemento_por_repeticao(driver,"/html/body/div[8]/div[3]/div/div/div/div[4]/fieldset/button[2]","click","PDC",3)
    #     funcoes.encontrar_elemento_por_repeticao(driver,"/html/body/div[5]/div[3]/div/div/div/div[1]/div/div[3]/button","click","PDC",2)
    #     funcoes.encontrar_elemento_por_repeticao(driver,"/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/div[3]/div/button[2]","click","PDC",3)
    #     funcoes.encontrar_elemento_por_repeticao(driver,"/html/body/div[5]/div[3]/div/div[2]/ul/div[1]","click","PDC",2) 
    #     sleep(3)    
    #     if ler_dados_da_planilha(tipo_de_solicitacao)[0][3] == "PAGO":
    #         funcoes.encontrar_elemento_por_repeticao(driver,"/html/body/div[5]/div[3]/div/div[2]/ul/div[1]","click","tramitar pago",2)    
    #     elif ler_dados_da_planilha(tipo_de_solicitacao)[0][3] == "PARCIALMENTE PAGO":  
    #         funcoes.encontrar_elemento_por_repeticao(driver,"/html/body/div[5]/div[3]/div/div[2]/ul/div[2]","click","tramitar parcialmente",2)
    #     sleep(2)
    #     atualizar_status_na_planilha(ler_dados_da_planilha(tipo_de_solicitacao)[0][4])    
    # driver.close()    
    # print("FIMMMMMMMMMMMMMMM")
    
