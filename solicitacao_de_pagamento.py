from datetime import date

from selenium.webdriver.common import keys
from gerenciadorPlanilhas import atualizar_status_na_planilha, ler_dados_da_planilha, preencher_solicitacao_na_planilha
import time
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
import gerenciadorPastas
import shutil
from funcoes import espera_explicita_de_elemento


'''Função responsavel por fazer buscas repetitivas em um mesmo elemento da página'''
def encontrar_elemento_por_repeticao(drive, element_path, acao, informacao_acao, tempo_espera):
    maximo_tentativas = 0
    while maximo_tentativas <= 20:
        print(informacao_acao, maximo_tentativas)
        try:
            drive.find_element_by_xpath(element_path)
            if acao == "click":
                drive.find_element_by_xpath(element_path).click()
                maximo_tentativas = 21
            elif acao == "link":
                maximo_tentativas = 21
                pass
        except:
            maximo_tentativas+=1
            time.sleep(tempo_espera)
    if maximo_tentativas > 20:
        return("#Erro " + informacao_acao)


'''Realizar coleta de todas as solicitações  de pagemento com status 'Pagamento Solicitado'
    e realizar o download de seus arquivos anexos na pasta do destinada a solicitação'''
def pagamentos(drive):
    data_em_texto = date.today().strftime("%d.%m.%Y")
    caminho_da_pasta = gerenciadorPastas.recuperar_diretorio_usuario() + "\\tpfe.com.br\\SGP e SGC - RPA\\Pagamentos\\"
    gerenciadorPastas.criarPastaData(caminho_da_pasta, data_em_texto)

    builder = ActionChains(drive)
    #drive.implicitly_wait(70)

    #Aceesando o menu de pagamento
    espera_explicita_de_elemento(drive,"/html/body/div[1]/div/div[2]/main/section/div/div/div/div/section/div/div[2]/div","encontrar","SRB1",100)

    drive.get("https://tpf2.madrix.app/runtime/44/list/186/Solicitação de Pagamento")
    time.sleep(8)
    #Filtrando as solicitações com status pagamento solicitado
    espera_explicita_de_elemento(drive,"/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/div/div/div","encontrar","SRB2",120)
    espera_explicita_de_elemento(drive,"/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/div/div/div","click","SRB3",120)
    espera_explicita_de_elemento(drive,"/html/body/div[4]/div[3]/ul/li[3]","click","SRB4",120)
    time.sleep(6)


    quantidade_de_requisicoes = int((drive.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/span[2]/div/p").get_attribute("innerText")).split(" ")[-1])

    #Percorrento por todas as solicitações filtradas com o status definido no sistema
    for qtd_solicitacoes in range(10):
        tramitar = 0
        #Lista para coleta das informações que serão enviadas para a planilha
        dados_do_formulario = []
        #Tipo
        dados_do_formulario.append("SP")
        path_comum = "/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[3]/div/div/div/table/tbody/tr[1]"
        #ID da Solicitação
        espera_explicita_de_elemento(drive, path_comum + "/td[4]/div","encontrar","id_solicitacao",120)
        id_solicitacao = drive.find_element_by_xpath(path_comum + "/td[4]/div").get_attribute("innerText")
        dados_do_formulario.append(id_solicitacao)
        #Razão social
        razao_social = drive.find_element_by_xpath(path_comum + "/td[5]/div").get_attribute("innerText")
        #estado
        estado = drive.find_element_by_xpath(path_comum + "/td[7]/div/span").get_attribute("innerText")
        #estado
        valor = drive.find_element_by_xpath(path_comum + "/td[10]").get_attribute("innerText")
        #data de previsao do pgt
        try:
            data_previsao = drive.find_element_by_xpath(path_comum + "/td[11]/div/time").get_attribute("innerText")
        except:
            data_previsao = ""
        try:
            #data de solicitacao
            data_solicitacao = drive.find_element_by_xpath(path_comum + "/td[12]/div/time").get_attribute("innerText")
        except:
            data_solicitacao = ""
        try:
            #data de pagamento
            data_pagamento = drive.find_element_by_xpath(path_comum + "/td[13]/div/time").get_attribute("innerText")
        except:
            data_pagamento = ""
        
        #Acessando informações dentro de uma solicitação
        drive.find_element_by_xpath(path_comum).click()
        time.sleep(3)
        #Banco
        banco = drive.find_element_by_xpath("/html/body/div[4]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[2]/div/div[4]/div[1]/div/div[1]/div[1]/div/div/input").get_attribute("value")
        #Agencia
        agencia = drive.find_element_by_xpath("/html/body/div[4]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[2]/div/div[4]/div[1]/div/div[2]/div/div/div/input").get_attribute("value")
        #Conta
        conta = drive.find_element_by_xpath("/html/body/div[4]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[2]/div/div[4]/div[2]/div/div[1]/div/div/div/input").get_attribute("value")
        #Tipo de Conta
        try:
            tipo_de_conta = drive.find_element_by_xpath("/html/body/div[4]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[2]/div/div[4]/div[2]/div/div[2]/div/div/div/div/div/div/div[1]/div").get_attribute("innerText")
        except:
            tipo_de_conta = ""
        
        #Natureza da conta
        try:
            natureza_conta = drive.find_element_by_xpath("/html/body/div[4]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[2]/div/div[4]/div[1]/div/div[1]/div[2]/div/div/div/div/div/div[1]/div").get_attribute("innerText")
        except:
            natureza_conta = ""
        
        #cnpj
        cnpj = drive.find_element_by_xpath("/html/body/div[4]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[2]/div/div[3]/div[1]/div/div/div/input").get_attribute("value")
        #cpf
        cpf = drive.find_element_by_xpath("/html/body/div[4]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[2]/div/div[3]/div[2]/div/div/div/input").get_attribute("value")
        


        #Imprimindo a Capa
        #drive.find_element_by_xpath("/html/body/div[5]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[1]/div/div[2]/div/button[1]").click()
        drive.find_element_by_xpath("/html/body/div[4]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[2]/div/div[8]/div[2]/div/div/button").click()
        time.sleep(15)
        drive.switch_to_frame(0)
        
        espera_explicita_de_elemento(drive,"/html/body/div/div/div/div[2]/div/table/tbody/tr/td[1]/table/tbody/tr/td[3]/div/table/tbody/tr/td[2]","click","filtro",60) 
        time.sleep(3)
        drive.find_element_by_xpath("/html/body/div/div/div/div[16]/div/div[1]/table/tbody/tr/td[2]").click()
        drive.find_element_by_xpath("/html/body/div/div/div/div[20]/div[4]/table/tbody/tr/td[1]/div/table/tbody/tr/td").click()
        drive.switch_to.default_content()
        drive.find_element_by_xpath("/html/body/div[7]/div[3]/div/div[1]/h2/div/div[2]/button").click()


        time.sleep(4)

        drive.find_element_by_xpath("/html/body/div[4]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[1]/div/div[2]/div/button[2]").click()
        time.sleep(10)

        #/html/body/div[5]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[3]/div
        div_externa = drive.find_element_by_xpath("/html/body/div[4]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[3]/div")
        trs = div_externa.find_elements_by_tag_name("a")
        for certidao in trs:
            certidao.click()
            time.sleep(5)
        
        #Download de Nota Fiscal
        drive.find_element_by_xpath("/html/body/div[4]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[1]/div/div[2]/div/button[3]").click()
        div_externa = drive.find_element_by_xpath("/html/body/div[4]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[4]/div/div/div/div[1]")
        trs = div_externa.find_elements_by_tag_name("a")
        for nfs in trs:
            nfs.click()
            time.sleep(10)

        #Criando pasta para o ID da solicitação no diretorio de reembolso
        nome_da_pasta = "SP ID " + id_solicitacao
        gerenciadorPastas.criarPastasFilhas('Pagamentos', nome_da_pasta)

        #Preencher lista com as informações que serão enviadas para a planilha
        #CPF
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
        #Razão Social
        dados_do_formulario.append(razao_social)
        #Forma de Pagt
        dados_do_formulario.append("")
        #Banco
        dados_do_formulario.append(banco)
        #Agencia
        dados_do_formulario.append(agencia)
        #Conta
        dados_do_formulario.append(conta)
        #Tipo de Conta
        dados_do_formulario.append(tipo_de_conta)
        #Natureza da conta
        dados_do_formulario.append(natureza_conta)
        #Valor
        dados_do_formulario.append(valor)
        #Valor Pg
        dados_do_formulario.append("")
        #Data Solicitada p/ pgto
        dados_do_formulario.append(data_previsao)
        #Data Solicitação
        dados_do_formulario.append(data_solicitacao)
        #Data Pgto
        dados_do_formulario.append(data_pagamento)
        
        valor_da_conta = drive.find_element_by_xpath("/html/body/div[4]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[2]/div/div[5]/div[1]/div[1]/div/div/input").get_attribute("value")
        valor_da_conta = valor_da_conta.replace(".","")
        valor_da_conta = valor_da_conta.replace("R$","")
        valor_da_conta = valor_da_conta.replace(",",".")
        valor_da_conta = float(valor_da_conta)
        
        if banco == "" or  agencia == "" or  conta == "" or  tipo_de_conta == "" or  natureza_conta == "" or valor_da_conta == 0:# and  banco != ""
            #Comentario Robo
            dados_do_formulario.append("Atenção: Dados bancarios inconpletos ou solicitação esta com valor da conta com zero.")
            tramitar = 1
        else:
            #Comentario Robo 
            dados_do_formulario.append("")
        #Ajuste
        dados_do_formulario.append("")
        


        #Movendo os Arquivos para a pasta da solicitacao
        arquivos = gerenciadorPastas.listar_arquivos_em_diretorios(gerenciadorPastas.recuperar_diretorio_usuario() + "\\tpfe.com.br\\SGP e SGC - RPA")
        for arquivo in arquivos:
            try:
                shutil.move(gerenciadorPastas.recuperar_diretorio_usuario() + "\\tpfe.com.br\\SGP e SGC - RPA\\" + arquivo, caminho_da_pasta + data_em_texto +"\\"+ nome_da_pasta + "\\" + arquivo)
            except:
                print("Não moveu o arquivo!")

        drive.find_element_by_xpath("/html/body/div[4]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[1]/div/div[2]/div/button[1]").click()
        drive.find_element_by_xpath("/html/body/div[4]/div[3]/div/div/div/div[4]/fieldset/button[3]").click()
        drive.find_element_by_xpath("/html/body/div[7]/div[3]/div/div[2]/ul/div[4]/div").click()
        time.sleep(10)
        #Status
        dados_do_formulario.append("Processada")
        #DATA DE EXECUÇÃO ROBO
        dados_do_formulario.append(date.today().strftime("%d/%m/%Y"))

        #Preencher Planilha
        preencher_solicitacao_na_planilha(dados_do_formulario,'SP')

        filtro_click = builder.send_keys(Keys.ESCAPE)
        filtro_click.perform()
        time.sleep(2)

        drive.get("https://tpf2.madrix.app/runtime/44/list/186/Solicitação de Pagamento")
        espera_explicita_de_elemento(drive,"/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/div/div/div","escontrar","SRB5",120)
        time.sleep(4)
        
    time.sleep(4)

    tramitar_para_pago(drive)
    drive.close()
    

def tramitar_para_pago(drive):

    drive.get("https://tpf2.madrix.app/runtime/44/list/186/Solicitação de Pagamento")
    #Filtrando as solicitações com status processado
    espera_explicita_de_elemento(drive,"/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/div/div/div","encontrar","SRB2",120)
    espera_explicita_de_elemento(drive,"/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/div/div/div","click","SRB3",120)
    espera_explicita_de_elemento(drive,"/html/body/div[4]/div[3]/ul/li[8]","click","SRB4",120)
    time.sleep(3)

    lista_de_tramitacao = ler_dados_da_planilha("SP")
    
    for solicitacao in lista_de_tramitacao:
        #Acessando o botao do filtro
        espera_explicita_de_elemento(drive,"/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/button[2]","click","SRB5",120)
        espera_explicita_de_elemento(drive,"/html/body/div[4]/div[3]/div/div[1]/div[1]/button","click","SRB6",120)
        drive.find_element_by_xpath("/html/body/div[4]/div[3]/div/ul/li[1]/div/div/div/div/input").send_keys(str(solicitacao[0]))
        drive.find_element_by_xpath("/html/body/div[4]/div[3]/div/div[2]/button").click()
        time.sleep(2)

        drive.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[3]/div/div/div/table/tbody/tr").click()
        #/html/body/div[4]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[2]/div/div[8]/div[1]/div/button
        time.sleep(5)
        drive.find_element_by_xpath("/html/body/div[4]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[1]/div/div[2]/div/button[4]").click()
        espera_explicita_de_elemento(drive,"/html/body/div[4]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[5]/div/div/div/div[1]/div[1]/div[1]/div/div/span/div/button[1]","click","SRB7",120)
        time.sleep(10)
        #Data
        drive.find_element_by_xpath("/html/body/div[7]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div/div[1]/div[1]/div/div/div/input").send_keys(solicitacao[2])
        #Valor
        valor = str(solicitacao[1]).replace(".",",")

        teste_casas_decimais_virgula = valor.count(",")
        teste_casas_decimais = valor.split(",")

        if teste_casas_decimais_virgula > 0 and len(teste_casas_decimais[1]) == 1:
            valor+="0"
        elif teste_casas_decimais_virgula == 0:
            valor+=",00"
        drive.find_element_by_xpath("/html/body/div[7]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div/div[1]/div[2]/div/div/div/input").send_keys(valor)
        #Salvar
        drive.find_element_by_xpath("/html/body/div[7]/div[3]/div/div/div/div[4]/fieldset/button[2]").click()
        #/html/body/div[7]/div[3]/div/div/div[1]/div[4]/fieldset/button[2]
        #/html/body/div[7]/div[3]/div/div/div/div[4]/fieldset/button[2]
        #Voltar
        #drive.find_element_by_xpath("/html/body/div[7]/div[3]/div/div/div/div[1]/button").click()
        #Dados
        drive.find_element_by_xpath("/html/body/div[4]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[1]/div/div[2]/div/button[1]").click()
        #Valor Liquido
        drive.find_element_by_xpath("/html/body/div[4]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[2]/div/div[5]/div[1]/div[2]/div/div/input").send_keys(valor)
        #Pagar
        drive.find_element_by_xpath("/html/body/div[4]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[2]/div/div[8]/div[1]/div/button").click()
        time.sleep(5)
        
        atualizar_status_na_planilha(int(solicitacao[4]))
        print("p")

#tramitar_para_pago()