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
    #espera_explicita_de_elemento(drive,"/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/div/div/div","encontrar","SRB2",20)
    espera_explicita_de_elemento(drive,"/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/div/div/div","click","SRB3",120)
    espera_explicita_de_elemento(drive,"/html/body/div[5]/div[3]/ul/li[3]","click","SRB4",60)

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
        
        drive.find_element_by_xpath("/html/body/div[5]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[1]/div/div[2]/div/button[2]").click()


        #/html/body/div[5]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[3]/div
        div_externa = drive.find_element_by_xpath("/html/body/div[5]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[3]/div")
        trs = div_externa.find_elements_by_tag_name("a")
        for certidao in trs:
            certidao.click()
            time.sleep(5)
        
        print(2)
        '''
        #Download de Certidão negativa
        estruta_externas =  drive.find_elements_by_css_selector("tr.MuiTableRow-root")
        for row in estruta_externas:
            try:
                link = row.find_element_by_tag_name("a")
                nome_anexo = link.get_attribute("innerText")
                nome_anexo = nome_anexo.split(".")[-1]
                if len(nome_anexo) == 3:
                   link.click()
            except:
                link = ""
                continue
        '''
        
        #Download de Nota Fiscal
        drive.find_element_by_xpath("/html/body/div[5]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[1]/div/div[2]/div/button[3]").click()
        div_externa = drive.find_element_by_xpath("/html/body/div[5]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[4]/div/div/div/div[1]")
        trs = div_externa.find_elements_by_tag_name("a")
        for nfs in trs:
            nfs.click()
            time.sleep(5)
        
        a = 2
        print(a)
        
     