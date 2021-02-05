from selenium import webdriver
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from os import path
from shutil import copyfile, move
from datetime import datetime
import openpyxl

dir = path.dirname(__file__)
novo = ''


def abrir_site():
    driver = webdriver.Chrome(dir +'/chromedriver')
    driver.get('https://cert.pmi.org/registry.aspx')
    elemento = aguardar_elemento(driver, 'dph_RegistryContent_Button1')
    return driver

def verificar_elemento(elemento):
    resultado = driver.find_elemnets_by_xpath(elemento)
    return resultado


def aguardar_elemento(driver, elemento):
    element = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, elemento)))


def copiar_arquivo(arquivo):
    now = datetime.now()
    now = str(now).replace(':','_')
    novo = 'Lista de certificados PMP_' + now + '.xlsx'
    copyfile(arquivo, 'temp.xlsx')
    move('temp.xlsx', dir + '/arquivo_saida/' + novo)
    return novo



def limpar_Template_Saida(arquivo):
    wb = openpyxl.load_workbook(arquivo)
    ws = wb.active
    iniPla_Saida = 3
    ultima_linha = ws.max_row
    ws.delete_rows(iniPla_Saida, ultima_linha)
    wb.save('Template_arquivo_saida.xlsx')
    wb.close()



def verificar_tam_nome(nb, nome):
    c = nome.index(nb)
    if c == len(nome) - 1:  # se tiver 2 caracetres ou 3 e for a ultima posição da lista utiliza como ultimo nome
        nb = nome[c]
    elif c + 1 < len(nome) - 1:
        nb = nome[c] + ' ' + nome[c + 1]
        c = c + 1
    else:  # se tiver dois ou tres letras e não for o ultimo indice da lista de nome concateno com o proximo indice.
        nb = nome[c] + ' ' + nome[c + 1]
    return nb








