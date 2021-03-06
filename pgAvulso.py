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





#função para tramitar as solicitações
def pagamentoAvulso(financeiro):
    #verificar se tem downloads antigos e apagar
    remover_arquivos_da_raiz(recuperar_diretorio_usuario() + "\\tpfe.com.br\\SGP e SGC - RPA\\")
    #data atual formatada
    data_em_texto = date.today().strftime("%d.%m.%Y")
    #caminho da pasta macro(pasta do dia)
    caminho_da_pasta = recuperar_diretorio_usuario() + "\\tpfe.com.br\\SGP e SGC - RPA\\Pagamento Avulso\\" 
    #criar pasta do dia dentro de pagamento avulso
    criarPastaData(caminho_da_pasta, data_em_texto)
    tipo_de_solicitacao = "PA"
    builder = ActionChains(financeiro)
    #financeiro.implicitly_wait(10)

    #ACESSANDO PAGAMENTO AVULSO
    espera_explicita_de_elemento(financeiro,"/html/body/div[1]/div/div[2]/main/section/div/div/div/div/section/div/div[2]/div","encontrar","SRB1",2)
    financeiro.get("https://tpf2.madrix.app/runtime/44/list/190/Solicitação de Pgto Avulso")
    #financeiro.implicitly_wait(10)
    #FILTRANDO OS PAGAMENTOS SOLICITADOS
    encontrar_elemento_por_repeticao(financeiro,"/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/div/div/div","click","filtro", 3)
    encontrar_elemento_por_repeticao(financeiro,"/html/body/div[4]/div[3]/ul/li[2]","click","filtro",2)
    
    encontrar_elemento_por_repeticao(financeiro,"/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/button[3]","click","filtro", 3)
    encontrar_elemento_por_repeticao(financeiro,"/html/body/div[4]/div[3]/div/div[1]/div[1]/button","click","limpar",0.2)
    encontrar_elemento_por_repeticao(financeiro, "/html/body/div[4]/div[3]/div/ul/li[3]/div/div/div", "click", "filtro",2)
    encontrar_elemento_por_repeticao(financeiro,"/html/body/div[5]/div[3]/ul/li[3]", "click", "filtro", 0.4)
    encontrar_elemento_por_repeticao(financeiro,"/html/body/div[5]/div[1]", "click", "filtro", 0.3)
    print(20*"=")
    encontrar_elemento_por_repeticao(financeiro,"/html/body/div[4]/div[3]/div/div[2]/button", "click", "fechando filtro", 0.4)
    
    #OBTER QUANTIDADE DE PAGAMENTOS
    sleep(5)
    quantidade_de_requisicoes = int((financeiro.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/span[2]/div/p[2]").get_attribute("innerText")).split(" ")[-1])
    #LAÇO PARA TRAMITAR TODOS OS PAGAMENTOS
    for linha in range(1): #voltar para antigo quantidades
        dados_do_formulario = []
        global identificador
        #armazenando o id de cada solicitaçao
        identificador = financeiro.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[3]/div/div/div/table/tbody/tr[1]/td[4]/div").get_attribute("innerText")
        global razao
        #armazenando a razao social de cada solicitaçao
        #/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[3]/div/div/div/table/tbody/tr[1]
        #/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[3]/div/div/div/table/tbody/tr[1]/td[2]/span/span[1]/input
        razao = financeiro.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[3]/div/div/div/table/tbody/tr[1]/td[6]/div").get_attribute("innerText")
        #estado = financeiro.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[3]/div/div/div/table/tbody/tr[1]/td[7]/div/span").get_attribute("innerText")
        #ACESSANDO A SOLICITAÇAO
        #/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/span[2]/div/p[2]
        encontrar_elemento_por_repeticao(financeiro, "/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[3]/div/div/div/table/tbody/tr[1]/td[2]/span/span[1]/input", "click", "click na checkbox", 4)
        #financeiro.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[3]/div/div/div/table/tbody/tr[1]/td[2]/span/span[1]/input").click()
        #clicar no lápis de edição
        encontrar_elemento_por_repeticao(financeiro, "/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/div[3]/div/button[1]", "click", "click na linha", 4)
        
        #PEGAR TODAS AS INFORMAÇOES PARA ALIMENTAR A PLANILHA
        caminho_em_comum_entre_campos_do_formulario = "/html/body/div[4]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[2]/div/div"
       
        #PA
        dados_do_formulario.append(tipo_de_solicitacao)
        #ID DA SOLICITAÇAO
        dados_do_formulario.append(identificador)
        #CPF/CNPJ
        encontrar_elemento_por_repeticao(financeiro, "/html/body/div[4]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[2]/div/div[2]/div[1]/div/div/div/input", "click", "cnpj", 4)
        cnpj = financeiro.find_element_by_xpath("/html/body/div[4]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[2]/div/div[2]/div[1]/div/div/div/input").get_attribute("value")
        cpf = financeiro.find_element_by_xpath(caminho_em_comum_entre_campos_do_formulario + "[2]/div[2]/div/div/div/input").get_attribute("value")
        try:
            validacao_cpf = 0
            validacao_cpf = cpf.replace(".","")
            validacao_cpf = validacao_cpf.replace("-","")
            validacao_cpf = int(validacao_cpf)
        except:
            validacao_cpf = 0
        if len(cpf)> 0 and validacao_cpf > 0:
            dados_do_formulario.append(cpf)
        else:
            dados_do_formulario.append(cnpj)
        #RAZÃO SOCIAL
        dados_do_formulario.append(razao) 
       
        try:
            #FORMA DE PAGAMENTO
            forma_de_pagamento = financeiro.find_element_by_xpath("/html/body/div[4]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[2]/div/div[4]/div[1]/div/div[1]/div/div/div/div/div/div/div[1]/div").get_attribute("innerText")
            dados_do_formulario.append(forma_de_pagamento)
        except:
            forma_de_pagamento = ""
            dados_do_formulario.append("")
        #BANCO
        dados_do_formulario.append(financeiro.find_element_by_xpath(caminho_em_comum_entre_campos_do_formulario + "[3]/div[1]/div/div[1]/div/div/div/input").get_attribute("value"))
        #AGENCIA
        dados_do_formulario.append(financeiro.find_element_by_xpath("/html/body/div[4]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[2]/div/div[3]/div[1]/div/div[2]/div/div/div/input").get_attribute("value"))
        #CONTA
        dados_do_formulario.append(financeiro.find_element_by_xpath(caminho_em_comum_entre_campos_do_formulario + "[3]/div[2]/div/div[1]/div/div/div/input").get_attribute("value"))
        #TIPO DE CONTA
        try:
            dados_do_formulario.append(financeiro.find_element_by_xpath("/html/body/div[4]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[2]/div/div[3]/div[2]/div/div[2]/div/div/div/div/div/div/div[1]/div").get_attribute("innerText"))
        except:
            dados_do_formulario.append("")
        #NATUREZA DA CONTA
        try:
            dados_do_formulario.append(financeiro.find_element_by_xpath(caminho_em_comum_entre_campos_do_formulario + "[4]/div[2]/div/div[3]/div/div/div/div/div/div/div[1]/div").get_attribute("innerText"))
        except:
            dados_do_formulario.append("")
        #VALOR
        valor = financeiro.find_element_by_xpath(caminho_em_comum_entre_campos_do_formulario + "[4]/div[1]/div/div[2]/div/div/div/input").get_attribute("value")
        dados_do_formulario.append(valor)
        #VALOR PAGO
        dados_do_formulario.append("") 
        #DATA SOLICITADA PARA PAGAMENTO
        dados_do_formulario.append(financeiro.find_element_by_xpath(caminho_em_comum_entre_campos_do_formulario + "[5]/div[1]/div/div[1]/div/div/div/input").get_attribute("value"))
        #DATA DA SOLICITAÇÃO 
        dados_do_formulario.append(financeiro.find_element_by_xpath(caminho_em_comum_entre_campos_do_formulario + "[5]/div[1]/div/div[2]/div/div/div/input").get_attribute("value"))
        #DATA DE PAGAMENTO 
        dados_do_formulario.append(financeiro.find_element_by_xpath(caminho_em_comum_entre_campos_do_formulario + "[5]/div[2]/div/div[1]/div/div/div/input").get_attribute("value"))


        valor_da_conta = valor.replace(".","")
        valor_da_conta = valor_da_conta.replace("R$","")
        valor_da_conta = valor_da_conta.replace(",",".")
        valor_da_conta = float(valor_da_conta)

        if forma_de_pagamento == "Transferência Bancária":
            if dados_do_formulario[5] == "" or  dados_do_formulario[6] == "" or  dados_do_formulario[7] == "" or  dados_do_formulario[8] == "" or  dados_do_formulario[9] == "" or valor_da_conta == 0:# and  banco != ""
                #Comentario Robo
                dados_do_formulario.append("Dados bancários incompletos ou solicitação está com valor zerado.")
                # tramitar = 1
            else:
                dados_do_formulario.append("")    
        elif forma_de_pagamento == "Boleto Bancário":
            if valor_da_conta == 0:
                dados_do_formulario.append("Boleto com valor zerado")
            else:
                #Comentario Robo 
                dados_do_formulario.append("")
        else:
            dados_do_formulario.append("")                
        #Ajuste
        dados_do_formulario.append("")
        
        sleep(3)
        #CRIAR A PASTA DO PAGAMENTO QUE ACABOU DE SER PROCESSADO
        
        razao = razao.replace("\\" , "")
        razao = razao.replace("/", "")
        razao = razao.replace(":", "")
        razao = razao.replace("*", "")
        razao = razao.replace("?", "")
        razao = razao.replace('"', "")
        razao = razao.replace("<", "")
        razao = razao.replace(">", "")
        razao = razao.replace("|", "")
        razao = razao.replace(".", "")
        razao = razao.replace(",", "")
        
        
        nome_da_pasta = "ID " + str(identificador)  + " " + str(razao)
        # else:    
        #     nome_da_pasta = (f"PA ID {identificador} {razao}")

        print(nome_da_pasta)
        criarPastasFilhas("Pagamento Avulso", nome_da_pasta) 

        #BAIXAR AS NF
        encontrar_elemento_por_repeticao(financeiro, "/html/body/div[4]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[1]/div/div[2]/div/button[2]", "click", "clicar em notas", 3)
        tbody2 = financeiro.find_element_by_xpath("/html/body/div[4]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[3]/div/div/div/div[1]/div[3]/table/tbody")
        #pega todas as linhas que contem nf
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
                
            # else:
            #     comentario_nao_possui_nota = (f"A solicitação não possui notas fiscais para serem baixadas")      
            #     print(comentario_nao_possui_nota)  
            #     dados_do_formulario[15] = comentario_nao_possui_nota
        sleep(2)
        #IMPRIMINDO
        try:
            encontrar_elemento_por_repeticao(financeiro, "/html/body/div[4]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[1]/div/div[2]/div/button[1]", "click", "voltar", 10)
            encontrar_elemento_por_repeticao(financeiro, "/html/body/div[4]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[2]/div/div[6]/div[2]/div/div[2]/div/div/button", "click", "imprimir", 3)
            financeiro.switch_to_frame(0)
            encontrar_elemento_por_repeticao(financeiro,"/html/body/div/div/div/div[2]/div/table/tbody/tr/td[1]/table/tbody/tr/td[3]/div/table/tbody/tr","click","filtro",2) 
            encontrar_elemento_por_repeticao(financeiro, "/html/body/div/div/div/div[16]/div/div[1]/table/tbody/tr/td[2]", "click", "baixar capa", 2)
            encontrar_elemento_por_repeticao(financeiro, "/html/body/div/div/div/div[20]/div[4]/table/tbody/tr/td[1]/div/table/tbody/tr/td", "click", "baixar capa", 2)
            financeiro.switch_to.default_content()
            encontrar_elemento_por_repeticao(financeiro, "/html/body/div[7]/div[3]/div/div[1]/h2/div/div[2]/button", "click", "baixar capa", 2)
            print("passou")
            encontrar_elemento_por_repeticao(financeiro, "/html/body/div[4]/div[3]/div/div/div/div[1]/div/div[3]/button", "click", "baixar capa", 2)
        except:
            dados_do_formulario[15] += "A capa nao foi baixada"
            print(dados_do_formulario[15])

        #MOVENDO ARQUIVOS
        print("mover arquivos")
        validar_download(caminho_da_pasta, data_em_texto, nome_da_pasta)
        
        #TRAMITAÇÃO
        encontrar_elemento_por_repeticao(financeiro,"/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/div[3]/div/button[2]","click","tramitar",2)
        encontrar_elemento_por_repeticao(financeiro, "/html/body/div[4]/div[3]/div/div[2]/ul/div[3]", "click", "tramitar", 2)
        #Status 
        dados_do_formulario.append("Processada")

        #DATA DE EXECUÇÃO ROBO
        dados_do_formulario.append(date.today().strftime("%d/%m/%Y"))
     
        preencher_solicitacao_na_planilha(dados_do_formulario, tipo_de_solicitacao)
        sleep(5)
        financeiro.get("https://tpf2.madrix.app/runtime/44/list/190/Solicitação de Pgto Avulso")

    sleep(1.5)
    print("Vai começar a contar")
    sleep(5)

     # 2° PARTE : ESPERANDO DO FINANCEIRO PRA TRAMITAR PRA PAGO 
    lista_de_tramitacao = ler_dados_da_planilha(tipo_de_solicitacao)
    #Filtro
    if len(lista_de_tramitacao) > 0:
        encontrar_elemento_por_repeticao(financeiro,"/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/div/div/div","click","filtro", 4)
        encontrar_elemento_por_repeticao(financeiro,"/html/body/div[4]/div[3]/ul/li[6]","click","filtro",2)
        #Para cada solicitaçao que precisa ser paga
        for solicitacao in lista_de_tramitacao:
            print(lista_de_tramitacao)
            encontrar_elemento_por_repeticao(financeiro,"/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/button[3]","click","filtro", 2)
            encontrar_elemento_por_repeticao(financeiro,"/html/body/div[4]/div[3]/div/div[1]/div[1]/button","click","filtro",0.2)
            #Filtrar o ID
            solicitacao2 = str(solicitacao[0])
            solicitacao2 = solicitacao2[2:len(solicitacao2)]
            financeiro.find_element_by_xpath("/html/body/div[4]/div[3]/div/ul/li[1]/div/div/div/div/input").send_keys(solicitacao2)
            # encontrar_elemento_por_repeticao(financeiro, "/html/body/div[6]/div[1]", "click", "filtro", 0.4 )
            print(20*"=")
            encontrar_elemento_por_repeticao(financeiro, "/html/body/div[4]/div[3]/div/div[2]/button", "click", "filtro", 0.4 )
            sleep(3)
            encontrar_elemento_por_repeticao(financeiro, "/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[3]/div/div/div/table/tbody/tr/td[2]/span/span[1]/input", "click", "LINHA", 2 )
            encontrar_elemento_por_repeticao(financeiro, "/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/div[3]/div/button[1]", "click", "editar",2 )
            encontrar_elemento_por_repeticao(financeiro, "/html/body/div[4]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[1]/div/div[2]/div/button[3]", "click", "pagamentos", 2 )
            sleep(2)

            encontrar_elemento_por_repeticao(financeiro,"/html/body/div[4]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[4]/div/div/div/div[1]/div[1]/div[1]/div/div/span/div/button[1]","click","add um novo pg",2)
            sleep(3)
            #Data
            financeiro.find_element_by_xpath("/html/body/div[7]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div/div[1]/div[1]/div/div/div/input").send_keys(str(solicitacao[2]))
            #Valor
            valor = str(solicitacao[1]).replace(".",",")

            teste_casas_decimais_virgula = valor.count(",")
            teste_casas_decimais = valor.split(",")

            if teste_casas_decimais_virgula > 0 and len(teste_casas_decimais[1]) == 1:
                valor+="0"
            elif teste_casas_decimais_virgula == 0:
                valor+=",00"
            financeiro.find_element_by_xpath("/html/body/div[7]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div/div[1]/div[2]/div/div/div/input").send_keys(valor)
            #Salvar
            encontrar_elemento_por_repeticao(financeiro,"/html/body/div[7]/div[3]/div/div/div/div[4]/fieldset/button[2]","click","SPA",3)
            #Voltar
            encontrar_elemento_por_repeticao(financeiro,"/html/body/div[4]/div[3]/div/div/div/div[1]/div/div[3]/button","click","SPA",2)
            #Tramitar
            encontrar_elemento_por_repeticao(financeiro,"/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/div[3]/div/button[2]","click","SPA",3) 
            sleep(3)   
            #Pago ou parcialmente pago 
            #Erro no Maiusculo
            if str(solicitacao[3]).lower() == "pago":
                encontrar_elemento_por_repeticao(financeiro,"/html/body/div[4]/div[3]/div/div[2]/ul/div[1]","click","pago",3)
                #/html/body/div[4]/div[3]/div/div[2]/ul/div[1]
            elif str(solicitacao[3]).lower() == "parcialmente pago":  
                encontrar_elemento_por_repeticao(financeiro,"/html/body/div[4]/div[3]/div/div[2]/ul/div[2]","click","SPA",3)
            sleep(4)
            
            atualizar_status_na_planilha(int(solicitacao[4]))     
    print("FIMMMMMMMMMMMMMMM AVULSO")
    #financeiro.close()
            











    # #para cada solicitação marcar a caixa de selecionar todas
    # financeiro.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[3]/div/div/div/table/thead/tr/th[2]/span/span[1]/input").click()
    # #desmarcar
    # financeiro.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[3]/div/div/div/table/thead/tr/th[2]/span/span[1]/input").click()
    # #exportando a planilha
    # financeiro.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/button[4]").click()


     # delay = 10 # seconds
    # try:
    #     myElem = WebfinanceiroWait(financeiro, delay).until(EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/button[3]")))
    #     filtro = financeiro.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/button[3]")
    #     try:
    #         filtro_click = builder.click(filtro)
    #         filtro_click.perform()
    #     except:
    #         pass
        
   # #pegar o corpo da tabela de solicitaçoes
    # tbody1 = financeiro.find_element_by_xpath("//*[@id='mainContent']/section/div/div/div/div[1]/div/div[3]/div/div/div/table/tbody")
    # #pegar as linhas que existem dentro do corpo
    # rows1 = tbody1.find_elements_by_tag_name("tr")

    # global solicitaçao
    # #lista que vai receber os índices de cada linha
    # solicitaçao = [] 

    # for row in range(len(rows1)):
    #     solicitaçao.append(row) #armazenando os indices de cada linha

    # print(solicitaçao)