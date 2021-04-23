from datetime import date
import time
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
import gerenciadorPastas
import shutil
#from funcoes import repeticao_de_acao
data_em_texto = date.today().strftime("%d.%m.%Y")
caminho_da_pasta = gerenciadorPastas.recuperar_diretorio_usuario() + "\\tpfe.com.br\\SGP e SGC - RPA\\Reembolso\\"
gerenciadorPastas.criarPastaData(caminho_da_pasta, data_em_texto)

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

def reembolso(drive):
    builder = ActionChains(drive)
    #drive.implicitly_wait(70)

    encontrar_elemento_por_repeticao(drive,"/html/body/div[1]/div/div[2]/main/section/div/div/div/div/section/div/div[2]/div","link","SRB1",0.2)
    drive.get("https://tpf.madrix.app/runtime/44/list/176/Solicitação de Reembolso")
    encontrar_elemento_por_repeticao(drive,"/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/div/div/div","link","SRB2",0.4)
    time.sleep(10)
    encontrar_elemento_por_repeticao(drive,"/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/div/div/div","click","SRB3",0.3)
    encontrar_elemento_por_repeticao(drive,"/html/body/div[5]/div[3]/ul/li[3]","click","SRB4",0.3)

    #drive.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/div/div/div").click()
    #drive.find_element_by_xpath("/html/body/div[5]/div[3]/ul/li[3]").click()

    quantidade_de_requisicoes = int((drive.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/span[2]/div/p").get_attribute("innerText")).split(" ")[-1])
    #TESTE
    quantidade_de_requisicoes-= 99

    for qtd_solicitacoes in range(2):

        #ID da Solicitação
        tipo = "SRB"
        path_comum = "/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[3]/div/div/div/table/tbody/tr[1]"
        id_solicitacao = drive.find_element_by_xpath(path_comum + "/td[4]/div").get_attribute("innerText")
        projeto = drive.find_element_by_xpath(path_comum + "/td[8]/div").get_attribute("innerText")
        estado = drive.find_element_by_xpath(path_comum + "/td[7]/div/span").get_attribute("innerText")
        valor_solicitado = drive.find_element_by_xpath(path_comum + "/td[9]").get_attribute("innerText")
        data_da_solicitacao = drive.find_element_by_xpath(path_comum + "/td[10]").get_attribute("innerText")
        drive.find_element_by_xpath(path_comum).click()
        time.sleep(3)
        drive.find_element_by_xpath("/html/body/div[5]/div[3]/div/div[2]/div/div/div/div/div").click()
        time.sleep(3)

        #filtro = drive.find_element_by_xpath("/html/body/div[5]/div[3]/div/div[2]/div/div/div/div/div")

        filtro_click = builder.send_keys(Keys.SPACE)
        filtro_click.perform()

        #filtro_click = builder.send_keys(Keys.SPACE)
        #filtro_click.perform()

        #drive.find_element_by_xpath("/html/body/div[5]/div[3]/div/div[2]/div/div/div/div/div/div[1]/div[1]").send_keys(Keys.SPACE)
        #drive.find_element_by_xpath("/html/body/div[5]/div[3]/div/div[2]/div/div/div/div/div/div[1]/div[1]").send_keys("Form Financeiro HTML")
        
        try:
            drive.find_element_by_xpath("/html/body/div[5]/div[3]/div/div[3]/button[2]").click()
        except:
            filtro_click = builder.send_keys(Keys.SPACE)
            filtro_click.perform()
            drive.find_element_by_xpath("/html/body/div[5]/div[3]/div/div[3]/button[2]").click()

        time.sleep(3)

        path_comum = "/html/body/div[5]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[2]/div/"
        cpf = drive.find_element_by_xpath(path_comum + "div[1]/div[2]/div/div/div/input").get_attribute("value")
        banco = drive.find_element_by_xpath(path_comum + "div[2]/div[2]/div/div[1]/div/div[1]/div/div/div/input").get_attribute("value")
        agencia = drive.find_element_by_xpath(path_comum + "div[2]/div[2]/div/div[1]/div/div[2]/div/div/div/input").get_attribute("value")
        conta = drive.find_element_by_xpath(path_comum + "div[2]/div[2]/div/div[2]/div/div[1]/div/div/div/input").get_attribute("value")
        tipo_de_conta = drive.find_element_by_xpath(path_comum + "div[2]/div[2]/div/div[2]/div/div[2]/div/div/div/input").get_attribute("value")

        drive.find_element_by_xpath("/html/body/div[5]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[1]/div/div[2]/div/button[2]").click()
        
        qtd_anexos = (drive.find_element_by_xpath("/html/body/div[5]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[3]/div/div/div/div[2]/span").get_attribute("innerText")).split(" ")[-1]
        
        for anexo in range(int(qtd_anexos)):
            drive.find_element_by_xpath("/html/body/div[5]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[3]/div/div/div/div[1]/div[3]/table/tbody/tr[" + str(anexo+1) + "]/td[2]/div[2]/div/a").click()

        time.sleep(10)
        nome_da_pasta = "ID " + id_solicitacao
        gerenciadorPastas.criarPastasFilhas('Reembolso', nome_da_pasta)

        arquivos = gerenciadorPastas.listar_arquivos_em_diretorios(gerenciadorPastas.recuperar_diretorio_usuario() + "\\tpfe.com.br\\SGP e SGC - RPA")

        for arquivo in arquivos:
            #print(arquivo)
            #movendo os arquivos para a pasta da sua solicitaçao
            try:
                shutil.move(gerenciadorPastas.recuperar_diretorio_usuario() + "\\tpfe.com.br\\SGP e SGC - RPA\\" + arquivo, caminho_da_pasta + data_em_texto +"\\"+ nome_da_pasta + "\\" + arquivo)
            except:
                print("Não moveu o arquivo!")

        
        
        drive.find_element_by_xpath("/html/body/div[5]/div[3]/div/div/div/div[4]/fieldset/button[3]").click()
        drive.find_element_by_xpath("/html/body/div[8]/div[3]/div/div[2]/ul/div[1]").click()

        time.sleep(5)

        filtro_click = builder.send_keys(Keys.ESCAPE)
        filtro_click.perform()

        drive.get("https://tpf.madrix.app/runtime/44/list/176/Solicitação de Reembolso")
        encontrar_elemento_por_repeticao(drive,"/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/div/div/div","link","SRB5",0.4)
        time.sleep(4)

#print("Fim")
        

        
        
        
        

