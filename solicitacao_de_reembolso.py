from datetime import date
import time
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
import gerenciadorPastas
import shutil

data_em_texto = date.today().strftime("%d.%m.%Y")
caminho_da_pasta = gerenciadorPastas.recuperar_diretorio_usuario() + "\\tpfe.com.br\\SGP e SGC - RPA\\Reembolso\\"
gerenciadorPastas.criarPastaData(caminho_da_pasta, data_em_texto)


def reembolso(drive):
    builder = ActionChains(drive)
    drive.implicitly_wait(70)
    drive.get("https://tpf.madrix.app/runtime/44/list/176/Solicitação de Reembolso")
    drive.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/div/div/div").click()
    drive.find_element_by_xpath("/html/body/div[5]/div[3]/ul/li[3]").click()

    quantidade_de_requisicoes = int((drive.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/span[2]/div/p").get_attribute("innerText")).split(" ")[-1])
    #TESTE
    quantidade_de_requisicoes-= 114

    #ID da Solicitação
    tipo = "SRB"
    path_comum = "/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[3]/div/div/div/table/tbody/tr[1]"
    id_solicitacao = drive.find_element_by_xpath(path_comum + "/td[4]/div").get_attribute("innerText")
    projeto = drive.find_element_by_xpath(path_comum + "/td[8]/div").get_attribute("innerText")
    estado = drive.find_element_by_xpath(path_comum + "/td[7]/div/span").get_attribute("innerText")
    valor_solicitado = drive.find_element_by_xpath(path_comum + "/td[9]").get_attribute("innerText")
    data_da_solicitacao = drive.find_element_by_xpath(path_comum + "/td[10]").get_attribute("innerText")
    drive.find_element_by_xpath(path_comum).click()

    drive.find_element_by_xpath("/html/body/div[5]/div[3]/div/div[2]/div/div/div/div/div").click()

    #filtro = drive.find_element_by_xpath("/html/body/div[5]/div[3]/div/div[2]/div/div/div/div/div")

    filtro_click = builder.send_keys(Keys.SPACE)
    filtro_click.perform()

    #drive.find_element_by_xpath("/html/body/div[5]/div[3]/div/div[2]/div/div/div/div/div/div[1]/div[1]").send_keys(Keys.SPACE)
    #drive.find_element_by_xpath("/html/body/div[5]/div[3]/div/div[2]/div/div/div/div/div/div[1]/div[1]").send_keys("Form Financeiro HTML")
    drive.find_element_by_xpath("/html/body/div[5]/div[3]/div/div[3]/button[2]").click()

    path_comum = "/html/body/div[5]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[2]/div/"
    cpf = drive.find_element_by_xpath(path_comum + "div[1]/div[2]/div/div/div/input").get_attribute("value")
    banco = drive.find_element_by_xpath(path_comum + "div[2]/div[2]/div/div[1]/div/div[1]/div/div/div/input").get_attribute("value")
    agencia = drive.find_element_by_xpath(path_comum + "div[2]/div[2]/div/div[1]/div/div[2]/div/div/div/input").get_attribute("value")
    conta = drive.find_element_by_xpath(path_comum + "div[2]/div[2]/div/div[2]/div/div[1]/div/div/div/input").get_attribute("value")
    tipo_de_conta = drive.find_element_by_xpath(path_comum + "div[2]/div[2]/div/div[2]/div/div[2]/div/div/div/input").get_attribute("value")

    drive.find_element_by_xpath("/html/body/div[5]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[1]/div/div[2]/div/button[2]").click()
    
    tabela_anexos = drive.find_element_by_xpath("/html/body/div[5]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[3]/div/div/div/div[1]/div[3]/table")

    for anexo in tabela_anexos.find_elements_by_xpath(".//tr"):
        filtro_click = builder.click(anexo)
        filtro_click.perform()
    
    nome_da_pasta = "ID " + id_solicitacao
    gerenciadorPastas.criarPastasFilhas('Reembolso', nome_da_pasta)

    arquivos = gerenciadorPastas.listar_arquivos_em_diretorios(gerenciadorPastas.recuperar_diretorio_usuario() + "\\tpfe.com.br\\SGP e SGC - RPA")

    for arquivo in arquivos:
        #print(arquivo)
        #movendo os arquivos para a pasta da sua solicitaçao
        try:
            shutil.move(caminho_da_pasta + arquivo, caminho_da_pasta + data_em_texto +"\\"+ nome_da_pasta + "\\" + arquivo)
        except:
            print("Não moveu o arquivo!")
    
    

