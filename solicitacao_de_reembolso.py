from datetime import date
from funcoes import espera_explicita_de_elemento, validar_download
from gerenciadorPlanilhas import ler_dados_da_planilha, preencher_solicitacao_na_planilha
import time
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
import gerenciadorPastas
import shutil

  

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


'''Realizar coleta de todas as solicitações com status 'Aprovado pelo gerente'
    e realizar o download de seus arquivos anexos na pasta do destinada a solicitação'''
def reembolso(drive):
    #verificar se tem downloads antigos e apagar
    gerenciadorPastas.remover_arquivos_da_raiz(gerenciadorPastas.recuperar_diretorio_usuario() + "\\tpfe.com.br\\SGP e SGC - RPA\\")
    data_em_texto = date.today().strftime("%d.%m.%Y")
    caminho_da_pasta = gerenciadorPastas.recuperar_diretorio_usuario() + "\\tpfe.com.br\\SGP e SGC - RPA\\Reembolso\\"
    gerenciadorPastas.criarPastaData(caminho_da_pasta, data_em_texto)
    builder = ActionChains(drive)
    #drive.implicitly_wait(70)

    #Aceesando o menu de reembolso
    espera_explicita_de_elemento(drive,"/html/body/div[1]/div/div[2]/main/section/div/div/div/div/section/div/div[2]/div","encontrar","SRB1",120)
    drive.get("https://tpf2.madrix.app/runtime/44/list/176/Solicitação de Reembolso")
    time.sleep(8)

    #Filtrando as solicitações com status aprovado pelo gerente
    espera_explicita_de_elemento(drive,"/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/div/div/div","encontrar","SRB2",120)
    espera_explicita_de_elemento(drive,"/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/div/div/div","click","SRB3",120)
    espera_explicita_de_elemento(drive,"/html/body/div[4]/div[3]/ul/li[3]","click","SRB4",120)

    quantidade_de_requisicoes = int((drive.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/span[2]/div/p").get_attribute("innerText")).split(" ")[-1])

    #Percorrento por todas as solicitações filtradas com o status definido no sistema
    for qtd_solicitacoes in range(10):
        #Lista para coleta das informações que serão enviadas para a planilha
        dados_do_formulario = []
        #Tipo
        dados_do_formulario.append("SR")
        path_comum = "/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[3]/div/div/div/table/tbody/tr[1]"
        #ID da Solicitação
        id_solicitacao = drive.find_element_by_xpath(path_comum + "/td[4]/div").get_attribute("innerText")
        dados_do_formulario.append(id_solicitacao)
        projeto = drive.find_element_by_xpath(path_comum + "/td[8]/div").get_attribute("innerText")
        estado = drive.find_element_by_xpath(path_comum + "/td[7]/div/span").get_attribute("innerText")
        valor_solicitado = drive.find_element_by_xpath(path_comum + "/td[9]").get_attribute("innerText")
        data_da_solicitacao = drive.find_element_by_xpath(path_comum + "/td[10]").get_attribute("innerText")

        #Acessando informações dentro de uma solicitação
        drive.find_element_by_xpath(path_comum).click()
        time.sleep(3)
        drive.find_element_by_xpath("/html/body/div[4]/div[3]/div/div[2]/div/div/div/div/div").click()
        time.sleep(3)

        #Acessando as informações presentes em 'Financeiro HTML'
        filtro_click = builder.send_keys(Keys.SPACE)
        filtro_click.perform()
        try:
            drive.find_element_by_xpath("/html/body/div[4]/div[3]/div/div[3]/button[2]").click()
        except:
            filtro_click = builder.send_keys(Keys.SPACE)
            filtro_click.perform()
            drive.find_element_by_xpath("/html/body/div[4]/div[3]/div/div[3]/button[2]").click()
        time.sleep(3)

        #Recuperando as demais informações da solicitação que irão para a planilha
        path_comum = "/html/body/div[4]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[2]/div/"
        cpf = drive.find_element_by_xpath(path_comum + "div[1]/div[2]/div/div/div/input").get_attribute("value")
        banco = drive.find_element_by_xpath(path_comum + "div[2]/div[2]/div/div[1]/div/div[1]/div/div/div/input").get_attribute("value")
        agencia = drive.find_element_by_xpath(path_comum + "div[2]/div[2]/div/div[1]/div/div[2]/div/div/div/input").get_attribute("value")
        conta = drive.find_element_by_xpath(path_comum + "div[2]/div[2]/div/div[2]/div/div[1]/div/div/div/input").get_attribute("value")
        tipo_de_conta = drive.find_element_by_xpath(path_comum + "div[2]/div[2]/div/div[2]/div/div[2]/div/div/div/input").get_attribute("value")

        #Acesando a aba de Anexo
        drive.find_element_by_xpath("/html/body/div[4]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[1]/div/div[2]/div/button[2]").click()
        qtd_anexos = (drive.find_element_by_xpath("/html/body/div[4]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[3]/div/div/div/div[2]/span").get_attribute("innerText")).split(" ")[-1]
        

        #Preencher lista com as informações que serão enviadas para a planilha
        #CPF
        dados_do_formulario.append(cpf)
        #Razão Social
        dados_do_formulario.append("")
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
        dados_do_formulario.append(valor_solicitado)
        #Valor Pg
        dados_do_formulario.append("")
        #Data Solicitada p/ pgto
        dados_do_formulario.append("")
        #Data Solicitação
        dados_do_formulario.append(data_da_solicitacao)
        #Data Pgto
        dados_do_formulario.append("")

        if banco == "" or  agencia == "" or  conta == "" or  tipo_de_conta == "":
            #Comentario Robo
            dados_do_formulario.append("Dados bancários incompletos")
        else:
            #Comentario Robo
            dados_do_formulario.append("")
        #Ajuste
        dados_do_formulario.append("")

        #Baixando os anexos
        for anexo in range(int(qtd_anexos)):
            drive.find_element_by_xpath("/html/body/div[4]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[3]/div/div/div/div[1]/div[3]/table/tbody/tr[" + str(anexo+1) + "]/td[2]/div[2]/div/a").click()
            time.sleep(10)
        
        #Imprimindo a Capa
        drive.find_element_by_xpath("/html/body/div[4]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[1]/div/div[2]/div/button[1]").click()
        drive.find_element_by_xpath("/html/body/div[4]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[2]/div/div[7]/div[2]/div/div/button").click()
        time.sleep(15)
        drive.switch_to_frame(0)
        
        espera_explicita_de_elemento(drive,"/html/body/div/div/div/div[2]/div/table/tbody/tr/td[1]/table/tbody/tr/td[3]/div/table/tbody/tr/td[2]","click","filtro",160) 
        espera_explicita_de_elemento(drive,"/html/body/div/div/div/div[16]/div/div[1]/table/tbody/tr/td[2]","click","filtro2",160) 
        #drive.find_element_by_xpath("/html/body/div/div/div/div[16]/div/div[1]/table/tbody/tr/td[2]").click()
        drive.find_element_by_xpath("/html/body/div/div/div/div[20]/div[4]/table/tbody/tr/td[1]/div/table/tbody/tr/td").click()
        drive.switch_to.default_content()
        drive.find_element_by_xpath("/html/body/div[7]/div[3]/div/div[1]/h2/div/div[2]/button").click()
        time.sleep(3)

        nome_da_pasta = id_solicitacao[0:1] + " ID " + id_solicitacao[2:len(id_solicitacao)]
        
        #Criando pasta para o ID da solicitação no diretorio de reembolso
        gerenciadorPastas.criarPastasFilhas('Reembolso', nome_da_pasta)
        
        #Movendo os Arquivos para a pasta da solicitacao
        validar_download(caminho_da_pasta, data_em_texto, nome_da_pasta)

        #Tramitando uma solicitação
        try:
            drive.find_element_by_xpath("/html/body/div[4]/div[3]/div/div/div/div[4]/fieldset/button[3]").click()
            drive.find_element_by_xpath("/html/body/div[7]/div[3]/div/div[2]/ul/div[1]").click()
            time.sleep(6)
            dados_do_formulario.append("Processada")
        except:
            dados_do_formulario.append("")
            dados_do_formulario[15] = "Falha na tramitação"
        
        #Preencher a planilha
        preencher_solicitacao_na_planilha(dados_do_formulario,'SR')

        time.sleep(5)

        filtro_click = builder.send_keys(Keys.ESCAPE)
        filtro_click.perform()

        drive.get("https://tpf2.madrix.app/runtime/44/list/176/Solicitação de Reembolso")
        espera_explicita_de_elemento(drive,"/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/div/div/div","link","SRB5",160)
        time.sleep(4)
    
    time.sleep(4)
    drive.close()


def tramitar_para_pago(drive):
    drive.get("https://tpf2.madrix.app/runtime/44/list/176/Solicitação de Reembolso")
    #Filtrando as solicitações com status processado
    
    espera_explicita_de_elemento(drive,"/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/div/div/div","encontrar","SRB2",120)
    espera_explicita_de_elemento(drive,"/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/div/div/div","click","SRB3",120)
    espera_explicita_de_elemento(drive,"/html/body/div[4]/div[3]/ul/li[9]","click","SRB4",120)
    
    time.sleep(3)

    lista_de_tramitacao = ler_dados_da_planilha("RB")
    
    for solicitacao in lista_de_tramitacao:
        #Acessando o botao do filtro
        espera_explicita_de_elemento(drive,"/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/button[3]","click","SRB5",120)
        espera_explicita_de_elemento(drive,"/html/body/div[4]/div[3]/div/div[1]/div[1]/button","click","SRB6",120)
        drive.find_element_by_xpath("/html/body/div[4]/div[3]/div/ul/li[1]/div/div/div/div/input").send_keys(str(solicitacao[0][]))
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

        

        
        
        
        

