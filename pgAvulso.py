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





#função para tramitar as solicitações
def pagamentoAvulso(financeiro):
    #verificar se tem downloads antigos e apagar
    gerenciadorPastas.remover_arquivos_da_raiz(gerenciadorPastas.recuperar_diretorio_usuario() + "\\tpfe.com.br\\SGP e SGC - RPA\\")
    #data atual formatada
    data_em_texto = date.today().strftime("%d.%m.%Y")
    #caminho da pasta macro(pasta do dia)
    caminho_da_pasta = gerenciadorPastas.recuperar_diretorio_usuario() + "\\tpfe.com.br\\SGP e SGC - RPA\\Pagamento Avulso\\" 
    #criar pasta do dia dentro de pagamento avulso
    gerenciadorPastas.criarPastaData(caminho_da_pasta, data_em_texto)
    tipo_de_solicitacao = "PA"
    builder = ActionChains(financeiro)
    financeiro.implicitly_wait(10)

    #ACESSANDO PAGAMENTO AVULSO
    # funcoes.espera_explicita_de_elemento(financeiro,"/html/body/div[1]/div/div[2]/main/section/div/div/div/div/section/div/div[2]/div","encontrar","SRB1",2)
    financeiro.get("https://tpf2.madrix.app/runtime/44/list/190/Solicitação de Pgto Avulso")
    #financeiro.implicitly_wait(10)
    #FILTRANDO OS PAGAMENTOS SOLICITADOS
    funcoes.encontrar_elemento_por_repeticao(financeiro,"/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/div/div/div","click","filtro", 3)
    funcoes.encontrar_elemento_por_repeticao(financeiro,"/html/body/div[4]/div[3]/ul/li[2]","click","filtro",2)
    
    funcoes.encontrar_elemento_por_repeticao(financeiro,"/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/button[3]","click","filtro", 3)
    funcoes.encontrar_elemento_por_repeticao(financeiro,"/html/body/div[4]/div[3]/div/div[1]/div[1]/button","click","limpar",0.2)
    funcoes.encontrar_elemento_por_repeticao(financeiro, "/html/body/div[4]/div[3]/div/ul/li[3]/div/div/div", "click", "filtro",2)
    funcoes.encontrar_elemento_por_repeticao(financeiro,"/html/body/div[5]/div[3]/ul/li[3]", "click", "filtro", 0.4)
    funcoes.encontrar_elemento_por_repeticao(financeiro,"/html/body/div[5]/div[1]", "click", "filtro", 0.3)
    print(20*"=")
    funcoes.encontrar_elemento_por_repeticao(financeiro,"/html/body/div[4]/div[3]/div/div[2]/button", "click", "fechando filtro", 0.4)
    
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
        razao = financeiro.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[3]/div/div/div/table/tbody/tr[1]/td[6]/div").get_attribute("innerText")
        #estado = financeiro.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[3]/div/div/div/table/tbody/tr[1]/td[7]/div/span").get_attribute("innerText")
        #ACESSANDO A SOLICITAÇAO
        funcoes.encontrar_elemento_por_repeticao(financeiro, "/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[3]/div/div/div/table/tbody/tr[1]/td[2]/span/span[1]/input", "click", "click na checkbox", 4)
        #financeiro.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[3]/div/div/div/table/tbody/tr[1]/td[2]/span/span[1]/input").click()
        #clicar no lápis de edição
        funcoes.encontrar_elemento_por_repeticao(financeiro, "/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/div[3]/div/button[1]", "click", "click na linha", 4)
        
        #PEGAR TODAS AS INFORMAÇOES PARA ALIMENTAR A PLANILHA
        caminho_em_comum_entre_campos_do_formulario = "/html/body/div[4]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[2]/div/div"

        #SA
        dados_do_formulario.append(tipo_de_solicitacao)
        #ID DA SOLICITAÇAO
        dados_do_formulario.append(identificador)
        #CPF/CNPJ
        cnpj = financeiro.find_element_by_xpath(caminho_em_comum_entre_campos_do_formulario + "[2]/div[1]/div/div/div/input").get_attribute("value")
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
        #FORMA DE PAGAMENTO
        try:
            dados_do_formulario.append(financeiro.find_element_by_xpath("/html/body/div[4]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[2]/div/div[4]/div[1]/div/div[1]/div/div/div/div/div/div/div[1]/div").get_attribute("innerText"))
        except:
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
        dados_do_formulario.append("0") 
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

        if dados_do_formulario[5] == "" or  dados_do_formulario[6] == "" or  dados_do_formulario[7] == "" or  dados_do_formulario[8] == "" or  dados_do_formulario[9] == "" or valor_da_conta == 0:# and  banco != ""
            #Comentario Robo
            dados_do_formulario.append("Dados bancários incompletos ou solicitação está com valor zerado.")
            # tramitar = 1
        else:
            #Comentario Robo 
            dados_do_formulario.append("")
        #Ajuste
        dados_do_formulario.append("")
        
        sleep(3)
        #CRIAR A PASTA DO PAGAMENTO QUE ACABOU DE SER PROCESSADO
        
        nome_da_pasta = (f"PA ID {identificador}")
        # else:    
        #     nome_da_pasta = (f"PA ID {identificador} {razao}")

        print(nome_da_pasta)
        gerenciadorPastas.criarPastasFilhas("Pagamento Avulso", nome_da_pasta) 

        #BAIXAR AS NF
        funcoes.encontrar_elemento_por_repeticao(financeiro, "/html/body/div[4]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[1]/div/div[2]/div/button[2]", "click", "clicar em notas", 3)
        tbody2 = financeiro.find_element_by_xpath("/html/body/div[4]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[3]/div/div/div/div[1]/div[3]/table/tbody")
        #pega todas as linhas que contem nf
        rows2 = tbody2.find_elements_by_tag_name("a") 
       
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
                
            # else:
            #     comentario_nao_possui_nota = (f"A solicitação não possui notas fiscais para serem baixadas")      
            #     print(comentario_nao_possui_nota)  
            #     dados_do_formulario[15] = comentario_nao_possui_nota
        sleep(5)
        #IMPRIMINDO
        try:
            funcoes.encontrar_elemento_por_repeticao(financeiro, "/html/body/div[4]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[1]/div/div[2]/div/button[1]", "click", "voltar", 10)
            funcoes.encontrar_elemento_por_repeticao(financeiro, "/html/body/div[4]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[2]/div/div[6]/div[2]/div/div[2]/div/div/button", "click", "imprimir", 3)
            financeiro.switch_to_frame(0)
            funcoes.encontrar_elemento_por_repeticao(financeiro,"/html/body/div/div/div/div[2]/div/table/tbody/tr/td[1]/table/tbody/tr/td[3]/div/table/tbody/tr","click","filtro",2) 
            funcoes.encontrar_elemento_por_repeticao(financeiro, "/html/body/div/div/div/div[16]/div/div[1]/table/tbody/tr/td[2]", "click", "baixar capa", 2)
            funcoes.encontrar_elemento_por_repeticao(financeiro, "/html/body/div/div/div/div[20]/div[4]/table/tbody/tr/td[1]/div/table/tbody/tr/td", "click", "baixar capa", 2)
            financeiro.switch_to.default_content()
            funcoes.encontrar_elemento_por_repeticao(financeiro, "/html/body/div[7]/div[3]/div/div[1]/h2/div/div[2]/button", "click", "baixar capa", 2)
            print("passou")
            funcoes.encontrar_elemento_por_repeticao(financeiro, "/html/body/div[4]/div[3]/div/div/div/div[1]/div/div[3]/button", "click", "baixar capa", 2)
        except:
            dados_do_formulario[15] += "A capa nao foi baixada"
            print(dados_do_formulario[15])

        #MOVENDO ARQUIVOS
        print("mover arquivos")
        funcoes.validar_download(caminho_da_pasta, data_em_texto, nome_da_pasta)
        
        #TRAMITAÇÃO
        funcoes.encontrar_elemento_por_repeticao(financeiro,"/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/div[3]/div/button[2]","click","tramitar",2)
        funcoes.encontrar_elemento_por_repeticao(financeiro, "/html/body/div[4]/div[3]/div/div[2]/ul/div[3]", "click", "tramitar", 2)
        #Status 
        dados_do_formulario.append("Processada")

        #DATA DE EXECUÇÃO ROBO
        dados_do_formulario.append(date.today().strftime("%d/%m/%Y"))
     
        preencher_solicitacao_na_planilha(dados_do_formulario, tipo_de_solicitacao)
        sleep(5)

    sleep(1.5)
    print("Vai começar a contar")
    sleep(5)

    # 2° PARTE : ESPERANDO DO FINANCEIRO PRA TRAMITAR PRA PAGO 
    lista_de_tramitacao = ler_dados_da_planilha(tipo_de_solicitacao)
    #Filtro
    if len(lista_de_tramitacao) > 0:
        funcoes.encontrar_elemento_por_repeticao(financeiro,"/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/div/div/div","click","filtro", 4)
        funcoes.encontrar_elemento_por_repeticao(financeiro,"/html/body/div[4]/div[3]/ul/li[6]","click","filtro",2)
        #Para cada solicitaçao que precisa ser paga
        for solicitacao in lista_de_tramitacao:
            print(lista_de_tramitacao)
            funcoes.encontrar_elemento_por_repeticao(financeiro,"/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/button[3]","click","filtro", 2)
            funcoes.encontrar_elemento_por_repeticao(financeiro,"/html/body/div[4]/div[3]/div/div[1]/div[1]/button","click","filtro",0.2)
            #Filtrar o ID
            financeiro.find_element_by_xpath("/html/body/div[4]/div[3]/div/ul/li[1]/div/div/div/div/input").send_keys(str(solicitacao[0]))
            funcoes.encontrar_elemento_por_repeticao(financeiro, "/html/body/div[6]/div[1]", "click", "filtro", 0.4 )
            print(20*"=")
            funcoes.encontrar_elemento_por_repeticao(financeiro, "/html/body/div[5]/div[3]/div/div[2]/button", "click", "filtro", 0.4 )
            sleep(3)
            funcoes.encontrar_elemento_por_repeticao(financeiro, "/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[3]/div/div/div/table/tbody/tr/td[2]/span/span[1]/input", "click", "LINHA", 2 )
            funcoes.encontrar_elemento_por_repeticao(financeiro, "/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/div[3]/div/button[1]", "click", "LINHA",2 )
            funcoes.encontrar_elemento_por_repeticao(financeiro, "/html/body/div[5]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[1]/div/div[2]/div/button[3]", "click", "LINHA", 2 )
            sleep(2)

            funcoes.encontrar_elemento_por_repeticao(financeiro,"/html/body/div[4]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[4]/div/div/div/div[1]/div[1]/div[1]/div/div/span/div/button[1]","click","SPA",2)
            sleep(3)
            #Data
            financeiro.find_element_by_xpath("/html/body/div[8]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div/div[1]/div[1]/div/div/div/input").send_keys(str(solicitacao[2]))
            #Valor
            valor = str(solicitacao[1]).replace(".",",")

            teste_casas_decimais_virgula = valor.count(",")
            teste_casas_decimais = valor.split(",")

            if teste_casas_decimais_virgula > 0 and len(teste_casas_decimais[1]) == 1:
                valor+="0"
            elif teste_casas_decimais_virgula == 0:
                valor+=",00"
            financeiro.find_element_by_xpath("/html/body/div[8]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div/div[1]/div[2]/div/div/div/input").send_keys(valor)
            #Salvar
            funcoes.encontrar_elemento_por_repeticao(financeiro,"/html/body/div[8]/div[3]/div/div/div/div[4]/fieldset/button[2]","click","SPA",3)
            #Voltar
            funcoes.encontrar_elemento_por_repeticao(financeiro,"/html/body/div[4]/div[3]/div/div/div/div[1]/div/div[3]/button","click","SPA",2)
            #Tramitar
            funcoes.encontrar_elemento_por_repeticao(financeiro,"/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/div[3]/div/button[2]","click","SPA",3) 
            sleep(3)   
            #Pago ou parcialmente pago 
            if str(solicitacao[3]) == "PAGO":
                funcoes.encontrar_elemento_por_repeticao(financeiro,"/html/body/div[5]/div[3]/div/div[2]/ul/div[1]","click","SPA",2)    
            elif str(solicitacao[3]) == "PARCIALMENTE PAGO":  
                funcoes.encontrar_elemento_por_repeticao(financeiro,"/html/body/div[5]/div[3]/div/div[2]/ul/div[2]","click","SPA",2)
            sleep(2)
            
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