from time import sleep
from datetime import date
import gerenciadorPastas
import funcoes
import shutil
import os
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains 
from openpyxl import load_workbook
from gerenciadorPlanilhas import preencher_solicitacao_na_planilha, tramitar_para_pago


#data atual formatada
data_em_texto = date.today().strftime("%d.%m.%Y")

#caminho da pasta macro(pasta do dia)
caminho_da_pasta = gerenciadorPastas.recuperar_diretorio_usuario() + "\\tpfe.com.br\\SGP e SGC - RPA\\Adiantamento\\" 
#criar pasta do dia dentro de pagamento avulso
gerenciadorPastas.criarPastaData(caminho_da_pasta, data_em_texto)

def adiantamento(driver):
    tipo_de_solicitacao = "SAT"
    builder = ActionChains(driver)
    #drive.implicitly_wait(70)
    dados_do_formulario = []
    #Aceesando o menu de reembolso
    driver.implicitly_wait(2)
    funcoes.espera_explicita_de_elemento(driver,"/html/body/div[1]/div/div[2]/main/section/div/div/div/div/section/div/div[2]/div","encontrar","SAT",2)
    driver.get("https://tpf.madrix.app/runtime/44/list/176/Solicitação de Adiantamento")
    driver.implicitly_wait(10)
    funcoes.espera_explicita_de_elemento(driver,"/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/div/div/div","click","filtro", 15)
    funcoes.espera_explicita_de_elemento(driver,"/html/body/div[5]/div[3]/ul/li[3]","click","filtro",3)
    funcoes.espera_explicita_de_elemento(driver,"/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/button[3]","click","filtro", 20)
    funcoes.espera_explicita_de_elemento(driver, "/html/body/div[5]/div[3]/div/ul/li[4]/div/div/div/div", "click", "filtro",0.3)
    funcoes.espera_explicita_de_elemento(driver,"/html/body/div[6]/div[3]/ul/li[7]", "click", "filtro", 0.4)
    funcoes.espera_explicita_de_elemento(driver,"/html/body/div[6]/div[1]", "click", "filtro", 0.3)
    print(20*"=")
    funcoes.espera_explicita_de_elemento(driver,"/html/body/div[5]/div[3]/div/div[2]/button", "click", "fechando filtro", 0.4)
    
    #abrir a linha (Laço para todos itens filtrados)
    sleep(5)
    quantidade_de_requisicoes = int((driver.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/span[2]/div/p").get_attribute("innerText")).split(" ")[-1])
    
    #laço para tramitar cada solicitaçao
    for linha in range(2): #voltar para antigo quantidades
        global identificador
        #armazenando o id de cada solicitaçao
        #colocar um int?
        identificador = driver.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[3]/div/div/div/table/tbody/tr[1]/td[4]/div").get_attribute("innerText")

        #ACESSANDO A SOLICITAÇAO
        driver.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[3]/div/div/div/table/tbody/tr[1]/td[2]/span/span[1]/input").click()
        #clicar no lápis de edição
        funcoes.espera_explicita_de_elemento(driver, "/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/div[3]/div/button[1]", "click", "click na linha", 4)
        #Acessando as informações presentes em 'driver HTML'
        sleep(5)
        filtro_click = builder.send_keys(Keys.ARROW_DOWN)
        filtro_click.perform()
        filtro_click = builder.send_keys(Keys.SPACE)
        filtro_click.perform()
        try:
            driver.find_element_by_xpath("/html/body/div[5]/div[3]/div/div[3]/button[2]").click()
        except:
            filtro_click = builder.send_keys(Keys.ARROW_DOWN)
            filtro_click.perform()
            filtro_click = builder.send_keys(Keys.SPACE)
            filtro_click.perform()
            driver.find_element_by_xpath("/html/body/div[5]/div[3]/div/div[3]/button[2]").click()
        sleep(3)

        #PEGAR TODAS AS INFORMAÇOES PARA ALIMENTAR A PLANILHA
        caminho_em_comum_entre_campos_do_formulario = "/html/body/div[5]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[2]/div/div"

        #SPA
        dados_do_formulario.append(tipo_de_solicitacao)
        #ID DA SOLICITAÇAO
        dados_do_formulario.append(identificador)
        #CPF/CNPJ
        dados_do_formulario.append(driver.find_element_by_xpath(caminho_em_comum_entre_campos_do_formulario + "[1]/div[2]/div/div/div/input").get_attribute("value"))
        #FORMA DE PAGAMENTO
        dados_do_formulario.append("Transferência Bancária")
        #BANCO 
        dados_do_formulario.append(driver.find_element_by_xpath(caminho_em_comum_entre_campos_do_formulario + "[2]/div[2]/div/div[1]/div/div[1]/div/div/div/input").get_attribute("value"))
        #AGENCIA
        dados_do_formulario.append(driver.find_element_by_xpath(caminho_em_comum_entre_campos_do_formulario + "[2]/div[2]/div/div[1]/div/div[2]/div/div/div/input").get_attribute("value"))
        #CONTA
        dados_do_formulario.append(driver.find_element_by_xpath(caminho_em_comum_entre_campos_do_formulario + "[2]/div[2]/div/div[2]/div/div[1]/div/div/div/input").get_attribute("value"))
        #TIPO DE CONTA 
        dados_do_formulario.append(driver.find_element_by_xpath(caminho_em_comum_entre_campos_do_formulario + "[2]/div[2]/div/div[2]/div/div[2]/div/div/div/input").get_attribute("value"))
        #NATUREZA DA CONTA
        dados_do_formulario.append("")
        #VALOR
        dados_do_formulario.append(driver.find_element_by_xpath(caminho_em_comum_entre_campos_do_formulario + "[4]/div[1]/div/div[1]/div/div/div/input").get_attribute("value"))
        #VALOR PAGO
        dados_do_formulario.append("0")
        #DATA SOLICITADA PARA PAGAMENTO
        dados_do_formulario.append("")
        #DATA DA SOLICITAÇÃO 
        dados_do_formulario.append(driver.find_element_by_xpath(caminho_em_comum_entre_campos_do_formulario + "[4]/div[1]/div/div[2]/div/div/div/input").get_attribute("value"))
        #DATA DE PAGAMENTO 
        dados_do_formulario.append("")
        #COMENTÁRIO ROBO
        dados_do_formulario.append("")                                                                             
        #AJUSTE Finaceiro
        dados_do_formulario.append("")
        #STATUS ROBO
        dados_do_formulario.append("Processada")
        
        #clicar em "notas fiscais"
        funcoes.espera_explicita_de_elemento(driver, "/html/body/div[5]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[1]/div/div[2]/div/button[2]", "click", "clicar em notas", 3)
        tbody2 = driver.find_element_by_xpath("/html/body/div[5]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[3]/div/div/div/div[1]/div[3]/table/tbody")
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
        funcoes.espera_explicita_de_elemento(driver, "/html/body/div[5]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[1]/div/div[2]/div/button[1]", "click", "imprimir", 1)
        funcoes.espera_explicita_de_elemento(driver, "/html/body/div[5]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[2]/div/div[6]/div[2]/div/div[2]/div/div/button", "click", "imprimir", 2)
        driver.switch_to_frame(0)
        funcoes.espera_explicita_de_elemento(driver,"/html/body/div/div/div/div[2]/div/table/tbody/tr/td[1]/table/tbody/tr/td[3]/div/table/tbody/tr","click","filtro",2) 
        funcoes.espera_explicita_de_elemento(driver, "/html/body/div/div/div/div[16]/div/div[1]/table/tbody/tr/td[2]", "click", "baixar capa", 2)
        funcoes.espera_explicita_de_elemento(driver, "/html/body/div/div/div/div[20]/div[4]/table/tbody/tr/td[1]/div/table/tbody/tr/td", "click", "baixar capa", 2)
        driver.switch_to.default_content()
        funcoes.espera_explicita_de_elemento(driver, "/html/body/div[8]/div[3]/div/div[1]/h2/div/div[2]/button", "click", "baixar capa", 2)
        print("passou")
        funcoes.espera_explicita_de_elemento(driver, "/html/body/div[5]/div[3]/div/div/div/div[1]/div/div[3]/button", "click", "baixar capa", 2)

        #CRIANDO A PASTA
        nome_da_pasta = (f"ID {identificador}")
        print(nome_da_pasta)
        #com a funçao do outro arquivo, criar a pasta da atual solicitação de acordo com o laço
        gerenciadorPastas.criarPastasFilhas("Adiantamento", nome_da_pasta)
      
        #tempo para salvar todas pastas
        sleep(3)
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
        funcoes.espera_explicita_de_elemento(driver,"/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/div[3]/div/button[2]","click","tramitar",2)
        funcoes.espera_explicita_de_elemento(driver, "/html/body/div[5]/div[3]/div/div[2]/ul/div[3]", "click", "tramitar", 2)
        preencher_solicitacao_na_planilha(dados_do_formulario, tipo_de_solicitacao)
        sleep(3)
