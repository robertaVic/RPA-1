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


#data atual formatada
data_em_texto = date.today().strftime("%d.%m.%Y")

#caminho da pasta macro(pasta do dia)
caminho_da_pasta = gerenciadorPastas.recuperar_diretorio_usuario() + "\\tpfe.com.br\\SGP e SGC - RPA\\Pagamento Avulso\\" 
#criar pasta do dia dentro de pagamento avulso
gerenciadorPastas.criarPastaData(caminho_da_pasta, data_em_texto)


#função para tramitar as solicitações
def pagamentoAvulso(financeiro):
    tipo_de_solicitacao = "SPA"
    builder = ActionChains(financeiro)
    financeiro.implicitly_wait(10)

    #ACESSANDO ANDIANTAMENTO
    funcoes.espera_explicita_de_elemento(financeiro,"/html/body/div[1]/div/div[2]/main/section/div/div/div/div/section/div/div[2]/div","encontrar","SRB1",2)
    financeiro.get("https://tpf.madrix.app/runtime/44/list/190/Solicitação de Pgto Avulso")
    financeiro.implicitly_wait(10)

    #FILTRANDO OS PAGAMENTOS SOLICITADOS
    funcoes.espera_explicita_de_elemento(financeiro,"/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/div/div/div","click","filtro", 15)
    funcoes.espera_explicita_de_elemento(financeiro,"/html/body/div[5]/div[3]/ul/li[2]","click","filtro",3)
    funcoes.espera_explicita_de_elemento(financeiro,"/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/button[3]","click","filtro", 20)
    funcoes.espera_explicita_de_elemento(financeiro, "/html/body/div[5]/div[3]/div/ul/li[3]/div/div/div/div", "click", "filtro",0.3)
    funcoes.espera_explicita_de_elemento(financeiro,"/html/body/div[6]/div[3]/ul/li[3]", "click", "filtro", 0.4)
    funcoes.espera_explicita_de_elemento(financeiro,"/html/body/div[6]/div[1]", "click", "filtro", 0.3)
    print(20*"=")
    funcoes.espera_explicita_de_elemento(financeiro,"/html/body/div[5]/div[3]/div/div[2]/button", "click", "fechando filtro", 0.4)
    
    #OBTER QUANTIDADE DE PAGAMENTOS
    sleep(5)
    quantidade_de_requisicoes = int((financeiro.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/span[2]/div/p").get_attribute("innerText")).split(" ")[-1])
    
    #LAÇO PARA TRAMITAR TODOS OS PAGAMENTOS
    for linha in range(2): #voltar para antigo quantidades
        dados_do_formulario = []
        global identificador
        #armazenando o id de cada solicitaçao
        identificador = financeiro.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[3]/div/div/div/table/tbody/tr[1]/td[4]/div").get_attribute("innerText")
        global razao
        #armazenando a razao social de cada solicitaçao
        razao = financeiro.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[3]/div/div/div/table/tbody/tr[1]/td[6]/div").get_attribute("innerText")
        
        #ACESSANDO A SOLICITAÇAO
        financeiro.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[3]/div/div/div/table/tbody/tr[1]/td[2]/span/span[1]/input").click()
        #clicar no lápis de edição
        funcoes.espera_explicita_de_elemento(financeiro, "/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/div[3]/div/button[1]", "click", "click na linha", 4)
        
        #PEGAR TODAS AS INFORMAÇOES PARA ALIMENTAR A PLANILHA
        caminho_em_comum_entre_campos_do_formulario = "/html/body/div[5]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[2]/div/div"

        #SPA
        dados_do_formulario.append(tipo_de_solicitacao)
        #ID DA SOLICITAÇAO
        dados_do_formulario.append(identificador)
        #CPF/CNPJ
        dados_do_formulario.append(financeiro.find_element_by_xpath(caminho_em_comum_entre_campos_do_formulario + "[2]/div[1]/div/div/div/input").get_attribute("value"))
        #RAZÃO SOCIAL
        dados_do_formulario.append(razao)
        #FORMA DE PAGAMENTO
        dados_do_formulario.append(financeiro.find_element_by_xpath(caminho_em_comum_entre_campos_do_formulario + "[3]/div[1]/div/div[1]/div/div/div/input").get_attribute("value"))
        #BANCO
        dados_do_formulario.append(financeiro.find_element_by_xpath(caminho_em_comum_entre_campos_do_formulario + "[3]/div[1]/div/div[2]/div/div/div/input").get_attribute("value"))
        #AGENCIA
        dados_do_formulario.append(financeiro.find_element_by_xpath(caminho_em_comum_entre_campos_do_formulario + "[3]/div[2]/div/div[1]/div/div/div/input").get_attribute("value"))
        #CONTA
        dados_do_formulario.append(financeiro.find_element_by_xpath(caminho_em_comum_entre_campos_do_formulario + "[3]/div[2]/div/div[2]/div/div/div/div/div/div/div[1]/input").get_attribute("value"))
        #TIPO DE CONTA
        dados_do_formulario.append(financeiro.find_element_by_xpath(caminho_em_comum_entre_campos_do_formulario + "[4]/div[2]/div/div[3]/div/div/div/div/div/div/div[1]/input").get_attribute("value"))
        #NATUREZA DA CONTA
        dados_do_formulario.append(financeiro.find_element_by_xpath(caminho_em_comum_entre_campos_do_formulario + "[4]/div[1]/div/div[1]/div/div/div/div/div/div/div[1]/input").get_attribute("value"))
        #VALOR
        dados_do_formulario.append(financeiro.find_element_by_xpath(caminho_em_comum_entre_campos_do_formulario + "[4]/div[1]/div/div[2]/div/div/div/input").get_attribute("value"))
        #VALOR PAGO
        dados_do_formulario.append("0")
        #DATA SOLICITADA PARA PAGAMENTO
        dados_do_formulario.append(financeiro.find_element_by_xpath(caminho_em_comum_entre_campos_do_formulario + "[5]/div[1]/div/div[1]/div/div/div/input").get_attribute("value"))
        #DATA DA SOLICITAÇÃO 
        dados_do_formulario.append(financeiro.find_element_by_xpath(caminho_em_comum_entre_campos_do_formulario + "[5]/div[1]/div/div[2]/div/div/div/input").get_attribute("value"))
        #DATA DE PAGAMENTO 
        dados_do_formulario.append(financeiro.find_element_by_xpath(caminho_em_comum_entre_campos_do_formulario + "[5]/div[2]/div/div[1]/div/div/div/input").get_attribute("value"))
        #COMENTÁRIO ROBO
        dados_do_formulario.append("")                                                                             
        #AJUSTE FINANCEIRO
        dados_do_formulario.append("")
        #STATUS ROBO
        dados_do_formulario.append("Processada")
        
        #clicar em "notas fiscais"
        funcoes.espera_explicita_de_elemento(financeiro, "/html/body/div[5]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[1]/div/div[2]/div/button[2]", "click", "clicar em notas", 3)
        tbody2 = financeiro.find_element_by_xpath("/html/body/div[5]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[3]/div/div/div/div[1]/div[3]/table/tbody")
        #pega todas as linhas que contem nf
        rows2 = tbody2.find_elements_by_tag_name("a") 
       
        #Baixando Nfs
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
        funcoes.espera_explicita_de_elemento(financeiro, "/html/body/div[5]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[1]/div/div[2]/div/button[1]", "click", "imprimir", 1)
        funcoes.espera_explicita_de_elemento(financeiro, "/html/body/div[5]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[2]/div/div[6]/div[2]/div/div[2]/div/div/button", "click", "imprimir", 2)
        financeiro.switch_to_frame(0)
        funcoes.espera_explicita_de_elemento(financeiro,"/html/body/div/div/div/div[2]/div/table/tbody/tr/td[1]/table/tbody/tr/td[3]/div/table/tbody/tr","click","filtro",2) 
        funcoes.espera_explicita_de_elemento(financeiro, "/html/body/div/div/div/div[16]/div/div[1]/table/tbody/tr/td[2]", "click", "baixar capa", 2)
        funcoes.espera_explicita_de_elemento(financeiro, "/html/body/div/div/div/div[20]/div[4]/table/tbody/tr/td[1]/div/table/tbody/tr/td", "click", "baixar capa", 2)
        financeiro.switch_to.default_content()
        funcoes.espera_explicita_de_elemento(financeiro, "/html/body/div[8]/div[3]/div/div[1]/h2/div/div[2]/button", "click", "baixar capa", 2)
        print("passou")
        funcoes.espera_explicita_de_elemento(financeiro, "/html/body/div[5]/div[3]/div/div/div/div[1]/div/div[3]/button", "click", "baixar capa", 2)
    
        # if not comentario_nota_fiscal:
        #     comentario = ("Nenhum")
        # else:
        #     comentario = (f"2- {comentario_nota_fiscal}")   
        #     if not comentario_nao_possui_nota:
        #         comentario = (f"Nenhum")
        #     else:
        #         comentario = (f"3- {comentario_nao_possui_nota}")
        #         comentario+= comentario, 

        # print(comentario) 

         #criando um modelo de nome de pastas para serem salvas(igualmente ao modelo do financeiro)
        #criando uma condicional para saber se a pasta tem apenas id ou tem os dois(id + razao)
        if not razao:
            nome_da_pasta = (f"ID {identificador}")
        else:    
            nome_da_pasta = (f"ID {identificador} {razao}")

        print(nome_da_pasta)
        #CRIAR A PASTA DO PAGAMENTO QUE ACABOU DE SER PROCESSADO
        gerenciadorPastas.criarPastasFilhas("Pagamento Avulso", nome_da_pasta) 

        sleep(3)
        sleep(1.5)
        print("mover arquivos")

        #listando os arquivos baixados na pasta macro(pasta do dia)
        arquivos = gerenciadorPastas.listar_arquivos_em_diretorios(gerenciadorPastas.recuperar_diretorio_usuario() + "\\tpfe.com.br\\SGP e SGC - RPA")
        print(arquivos)

        #criando um laço para mover cada um para sua pasta especifica
        for arquivo in arquivos:
            print(arquivo)
            #movendo os arquivos para a pasta da sua solicitaçao
            try:
                shutil.move(gerenciadorPastas.recuperar_diretorio_usuario() + "\\tpfe.com.br\\SGP e SGC - RPA\\" + arquivo, caminho_da_pasta + data_em_texto +"\\"+ nome_da_pasta + "\\" + arquivo)
            except:
                print("não moveu o arquivo!")
        sleep(3)

        #tramitação das solicitaçoes
        ler_dados_da_planilha(tipo_de_solicitacao)
        funcoes.espera_explicita_de_elemento(financeiro,"/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/div[3]/div/button[2]","click","tramitar",2)
        funcoes.espera_explicita_de_elemento(financeiro, "/html/body/div[5]/div[3]/div/div[2]/ul/div[3]", "click", "tramitar", 2)
        preencher_solicitacao_na_planilha(dados_do_formulario, tipo_de_solicitacao)
        sleep(3)

    sleep(1.5)
    financeiro.quit()
        
    print("Vai começar a contar")
    #selecionar_ids_do_tipo_de_solicitacao(tipo_de_solicitacao)
    for i in range(0,60):
        print(i)
        sleep(1)
    #2° parte: ESPERANDO DO FINANCEIRO PRA TRAMITAR PRA PAGO 
    # #parte do sgp
def tramitar_para_pago_no_sgp(financeiro):
    funcoes.chamarDriver(financeiro)
    funcoes.fazerLogin(financeiro)
    funcoes.espera_explicita_de_elemento(financeiro,"/html/body/div[1]/div/div[2]/main/section/div/div/div/div/section/div/div[2]/div","encontrar","SPA",2)
    financeiro.get("https://tpf.madrix.app/runtime/44/list/190/Solicitação de Pgto Avulso")
    for i in range(len(ler_dados_da_planilha(tipo_de_solicitacao))):
        funcoes.espera_explicita_de_elemento(financeiro,"/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/div/div/div","click","filtro",0.2)
        funcoes.espera_explicita_de_elemento(financeiro,"/html/body/div[5]/div[3]/ul/li[6]","click","filtro",0.2)
        funcoes.espera_explicita_de_elemento(financeiro,"/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/button[3]","click","filtro",0.2)
        funcoes.espera_explicita_de_elemento(financeiro,"/html/body/div[5]/div[3]/div/div[1]/div[1]/button","click","filtro",0.2)
        financeiro.find_element_by_xpath("/html/body/div[5]/div[3]/div/ul/li[1]/div/div/div/div/input").send_keys(ler_dados_da_planilha(tipo_de_solicitacao)[i][0])
        funcoes.espera_explicita_de_elemento(financeiro, "/html/body/div[5]/div[3]/div/ul/li[3]/div/div/div/div", "click", "filtro", 0.4)
        funcoes.espera_explicita_de_elemento(financeiro, "/html/body/div[6]/div[3]/ul/li[7]", "click", "filtro", 0.4)
        funcoes.espera_explicita_de_elemento(financeiro, "/html/body/div[6]/div[1]", "click", "filtro", 0.4 )
        print(20*"=")
        funcoes.espera_explicita_de_elemento(financeiro, "/html/body/div[5]/div[3]/div/div[2]/button", "click", "filtro", 0.4 )
        sleep(5)
        funcoes.espera_explicita_de_elemento(financeiro, "/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[3]/div/div/div/table/tbody/tr/td[2]/span/span[1]/input", "click", "filtro", 0.4 )
        funcoes.espera_explicita_de_elemento(financeiro, "/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/div[3]/div/button[1]", "click", "filtro", 0.4 )
        funcoes.espera_explicita_de_elemento(financeiro, "/html/body/div[5]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[1]/div/div[2]/div/button[3]", "click", "filtro", 0.4 )
        sleep(3)
        funcoes.espera_explicita_de_elemento(financeiro,"/html/body/div[5]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[4]/div/div/div/div[1]/div[1]/div[1]/div/div/span/div/button[1]","click","SPA",0.2)
        financeiro.find_element_by_xpath("/html/body/div[8]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div/div[1]/div[1]/div/div/div/input").send_keys(ler_dados_da_planilha(tipo_de_solicitacao)[i][2])
        financeiro.find_element_by_xpath("/html/body/div[8]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div/div[1]/div[2]/div/div/div/input").send_keys(ler_dados_da_planilha(tipo_de_solicitacao)[i][1])
        funcoes.espera_explicita_de_elemento(financeiro,"/html/body/div[8]/div[3]/div/div/div/div[4]/fieldset/button[2]","click","SPA",0.4)
        funcoes.espera_explicita_de_elemento(financeiro,"/html/body/div[5]/div[3]/div/div/div/div[1]/div/div[3]/button","click","SPA",0.4)
        funcoes.espera_explicita_de_elemento(financeiro,"/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/div[3]/div/button[2]","click","SPA",0.4)
        funcoes.espera_explicita_de_elemento(financeiro,"/html/body/div[5]/div[3]/div/div[2]/ul/div[1]","click","SPA",0.4) 
        sleep(3)    
        if ler_dados_da_planilha(tipo_de_solicitacao)[i][3] == "PAGO":
            funcoes.espera_explicita_de_elemento(financeiro,"/html/body/div[5]/div[3]/div/div[2]/ul/div[1]","click","SPA",0.4)    
        elif ler_dados_da_planilha(tipo_de_solicitacao)[i][3] == "PARCIALMENTE PAGO":  
            funcoes.espera_explicita_de_elemento(financeiro,"/html/body/div[5]/div[3]/div/div[2]/ul/div[2]","click","SPA",0.4)
        sleep(4)
        atualizar_status_na_planilha(ler_dados_da_planilha(tipo_de_solicitacao)[i][4])    
    print("FIMMMMMMMMMMMMMMM")
            











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