
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
from time import sleep 


numero_oab = 133864

#1 - Entrar no site consultado: https://pje-consulta-publica.tjmg.jus.br/
driver = webdriver.Chrome()
driver.get ('https://pje-consulta-publica.tjmg.jus.br/')
sleep(10)

#2 - Inserir número da OAB e estado: //tag[atributo='valor]
campo_oab = driver.find_element(By.XPATH,"//input[@id='fPP:Decoration:numeroOAB']")
campo_oab.send_keys(numero_oab)
#Estado 
selecao_estado = driver.find_element(By.XPATH, "//select[@id='fPP:Decoration:estadoComboOAB']")
opcao_estado = Select(selecao_estado)
opcao_estado.select_by_visible_text('SP')

#3 - Pesquisar
pesquisar = driver.find_element(By.XPATH, "//input[@id='fPP:searchProcessos']")
pesquisar.click()
sleep(10)

#4 - Entrar em cada processo
processos = driver.find_elements(By.XPATH, "//b[@class='btn-block']")
for processo in processos:
    processo.click()
sleep(15)
#Ajustar tamanho da janela
janela = driver.window_handles
driver.switch_to.window(janela[-1])
driver.set_window_size(1920,1080)
#Coletar número dos processos
numero_processo = driver.find_elements(By.XPATH, "//div[@class='col-sm-12']")
numero_processo = numero_processo[0]
numero_processo = numero_processo.text
#Data
data_processo = driver.find_elements(By.XPATH, "//div[@class='col-sm-12']")
data_processo = data_processo[0]
data_processo = data_processo.text