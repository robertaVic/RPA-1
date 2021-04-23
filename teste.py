from selenium import webdriver
from selenium.webdriver import Chrome
from gerenciadorPlanilhas import tramitar_para_pago

driver = Chrome()

tramitar_para_pago("SPA", driver)