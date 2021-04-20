from datetime import date
import time

from selenium.webdriver.common.action_chains import ActionChains
from gerenciadorPastas import criarPastaData, recuperar_diretorio_usuario

data_em_texto = date.today().strftime("%d.%m.%Y")
caminho_da_pasta = recuperar_diretorio_usuario() + "\\tpfe.com.br\\SGP e SGC - RPA\\Reembolso\\"
criarPastaData(caminho_da_pasta, data_em_texto)


def reembolso(drive):
    builder = ActionChains(drive)
    drive.implicitly_wait(70)
    drive.get("https://tpf.madrix.app/runtime/44/list/176/Solicitação de Reembolso")
    drive.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/div/div/div").click()
    drive.find_element_by_xpath("/html/body/div[5]/div[3]/ul/li[3]").click()
    drive.find_element_b

