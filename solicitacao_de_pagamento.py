from datetime import date
from gerenciadorPlanilhas import preencher_solicitacao_na_planilha
import time
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
import gerenciadorPastas
import shutil
from funcoes import espera_explicita_de_elemento

data_em_texto = date.today().strftime("%d.%m.%Y")
caminho_da_pasta = gerenciadorPastas.recuperar_diretorio_usuario() + "\\tpfe.com.br\\SGP e SGC - RPA\\Pagamentos\\"
gerenciadorPastas.criarPastaData(caminho_da_pasta, data_em_texto)

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
    builder = ActionChains(drive)
    #drive.implicitly_wait(70)

    #Aceesando o menu de pagamento
    espera_explicita_de_elemento(drive,"/html/body/div[1]/div/div[2]/main/section/div/div/div/div/section/div/div[2]/div","encontrar","SRB1",100)

    drive.get("https://tpf.madrix.app/runtime/44/list/186/Solicitação de Pagamento")
    time.sleep(8)
    #Filtrando as solicitações com status pagamento solicitado
    espera_explicita_de_elemento(drive,"/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/div/div/div","encontrar","SRB2",120)
    espera_explicita_de_elemento(drive,"/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/div/div/div","click","SRB3",120)
    espera_explicita_de_elemento(drive,"/html/body/div[5]/div[3]/ul/li[3]","click","SRB4",60)
    time.sleep(3)


    quantidade_de_requisicoes = int((drive.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/span[2]/div/p").get_attribute("innerText")).split(" ")[-1])

    #Percorrento por todas as solicitações filtradas com o status definido no sistema
    for qtd_solicitacoes in range(2):
        #Lista para coleta das informações que serão enviadas para a planilha
        dados_do_formulario = []
        #Tipo
        dados_do_formulario.append("SPG")
        path_comum = "/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[3]/div/div/div/table/tbody/tr[1]"
        #ID da Solicitação
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
        banco = drive.find_element_by_xpath("/html/body/div[5]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[2]/div/div[4]/div[1]/div/div[1]/div[1]/div/div/input").get_attribute("value")
        #Agencia
        agencia = drive.find_element_by_xpath("/html/body/div[5]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[2]/div/div[4]/div[1]/div/div[2]/div/div/div/input").get_attribute("value")
        #Conta
        conta = drive.find_element_by_xpath("/html/body/div[5]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[2]/div/div[4]/div[2]/div/div[1]/div/div/div/input").get_attribute("value")
        #Tipo de Conta
        tipo_de_conta = drive.find_element_by_xpath("/html/body/div[5]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[2]/div/div[4]/div[2]/div/div[2]/div/div/div/div/div/div/div[1]/div").get_attribute("innerText")
        
        #cnpj
        cnpj = drive.find_element_by_xpath("/html/body/div[5]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[2]/div/div[3]/div[1]/div/div/div/input").get_attribute("value")
        #cpf
        cpf = drive.find_element_by_xpath("/html/body/div[5]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[2]/div/div[3]/div[2]/div/div/div/input").get_attribute("value")
        


        #Imprimindo a Capa
        #drive.find_element_by_xpath("/html/body/div[5]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[1]/div/div[2]/div/button[1]").click()
        drive.find_element_by_xpath("/html/body/div[5]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[2]/div/div[8]/div[2]/div/div/button").click()
        time.sleep(15)
        drive.switch_to_frame(0)
        
        espera_explicita_de_elemento(drive,"/html/body/div/div/div/div[2]/div/table/tbody/tr/td[1]/table/tbody/tr/td[3]/div/table/tbody/tr/td[2]","click","filtro",60) 
        time.sleep(3)
        drive.find_element_by_xpath("/html/body/div/div/div/div[16]/div/div[1]/table/tbody/tr/td[2]").click()
        drive.find_element_by_xpath("/html/body/div/div/div/div[20]/div[4]/table/tbody/tr/td[1]/div/table/tbody/tr/td").click()
        drive.switch_to.default_content()
        drive.find_element_by_xpath("/html/body/div[8]/div[3]/div/div[1]/h2/div/div[2]/button").click()
        time.sleep(4)

        drive.find_element_by_xpath("/html/body/div[5]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[1]/div/div[2]/div/button[2]").click()
        time.sleep(10)

        #/html/body/div[5]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[3]/div
        div_externa = drive.find_element_by_xpath("/html/body/div[5]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[3]/div")
        trs = div_externa.find_elements_by_tag_name("a")
        for certidao in trs:
            certidao.click()
            time.sleep(5)
        
        #Download de Nota Fiscal
        drive.find_element_by_xpath("/html/body/div[5]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[1]/div/div[2]/div/button[3]").click()
        div_externa = drive.find_element_by_xpath("/html/body/div[5]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[4]/div/div/div/div[1]")
        trs = div_externa.find_elements_by_tag_name("a")
        for nfs in trs:
            nfs.click()
            time.sleep(5)

        #Criando pasta para o ID da solicitação no diretorio de reembolso
        nome_da_pasta = "ID " + id_solicitacao
        gerenciadorPastas.criarPastasFilhas('Pagamentos', nome_da_pasta)

        #Preencher lista com as informações que serão enviadas para a planilha
        #CPF
        if len(cpf)>0:
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
        dados_do_formulario.append("")
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

        
        #Tramitar para processado
        drive.find_element_by_xpath("/html/body/div[5]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[1]/div/div[2]/div/button[1]").click()
        drive.find_element_by_xpath("/html/body/div[5]/div[3]/div/div/div/div[4]/fieldset/button[3]").click()
        drive.find_element_by_xpath("/html/body/div[8]/div[3]/div/div[2]/ul/div[4]/div").click()
        time.sleep(5)
        dados_do_formulario.append("Processada")

        #Preencher Planilha
        preencher_solicitacao_na_planilha(dados_do_formulario,'SPG')

        filtro_click = builder.send_keys(Keys.ESCAPE)
        filtro_click.perform()
        time.sleep(2)

        drive.get("https://tpf.madrix.app/runtime/44/list/186/Solicitação de Pagamento")
        espera_explicita_de_elemento(drive,"/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/div/div/div","escontrar","SRB5",120)
        time.sleep(4)
        
    time.sleep(4)
    drive.close()