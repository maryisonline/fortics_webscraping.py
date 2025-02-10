from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
import time
import os
import schedule
import pandas as pd
from datetime import datetime, timedelta, date
import calendar

# definindo nome de arquivos e caminhos utilizados
oldName = 'relatorio-de-tabulacao-'
newName = 'BASE_FORTICS_ATUAL.xlsx'
originalName = 'BASE_FORTICS.xlsx'
download_path = r'C:\Users\Benvista\Downloads'
final_path = r'C:\Users\Benvista\Desktop\BASES\TABULAÇÃO'

# excluindo antiga extração se existir
for file in os.listdir(final_path):
    if file.startswith(newName):
        antigaBase = os.path.join(final_path, newName)
        os.remove(antigaBase)

def webscraping():

    options = webdriver.ChromeOptions()
    service = Service(ChromeDriverManager().install())
    web = webdriver.Chrome(service = service, options = options)

    web.get('https://benvista.sz.chat/static/signin?action=session_expired')
    
    # login e senha
    wait = WebDriverWait(web, 100)
    email = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '[name="email"]')))
    email.send_keys('mary.castro@benvista.com.br')
    web.find_element(By.CSS_SELECTOR, '[name="password"]').send_keys('Ma@2805203')
    web.find_element(By.CSS_SELECTOR, '#q-app > div > div > div.fullscreen.form-signin > div.login-card.q-card > div.card-form.q-card__section.q-card__section--vert.sz-card-section.bg-white > form > button > span.q-btn__wrapper.col.row.q-anchor--skip > span').click()

    ####### ACIMA OK

    # interagindo com elementos da pagina para fazer o caminho e baixar o arquivo sempre com a opcao de semana passada
    wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '#navbar > header > nav.navbar.navbar-fixed-top > a'))).click()
    wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '#navbar > aside > section > ul > li:nth-child(2) > a'))).click()
    wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '#navbar > aside > section > ul > li:nth-child(2) > ul > li:nth-child(2) > a'))).click()
    wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '#Periods-input'))).click()
    wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '#Periods-picker-container-DatePicker > div.shortcuts-container > button:nth-child(2) > span.custom-button-content.flex.align-center.justify-content-center > span'))).click()
    wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '#Periods-wrapper > div.datetimepicker.flex.visible > div > div.datepicker-buttons-container.flex.justify-content-right.button-validate.flex-fixed > button'))).click()
    wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '#reportAttendance > div > div.ui.row > div:nth-child(2) > div > div > div.text'))).click()
    wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '#reportAttendance > div > div.ui.row > div:nth-child(2) > div > div > div.menu.ft-scroll.transition.visible > div:nth-child(6)'))).click()
    # acima seleciona ate o campo de tabulação

    wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '#reportAttendance > div > div.ui.row > div:nth-child(6) > span > div'))).click()
    wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '#reportAttendance > div > div.ui.row > div:nth-child(6) > span > div > div.menu.transition > div:nth-child(2)'))).click()
    print("Inserido no input a opção: Call Center")
    wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '#reportAttendance > div > div.ui.row > div:nth-child(6) > span > div'))).click()
    
    confirmacao = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '#reportAttendance > div > div.ui.row > div:nth-child(6) > span > div > div.menu.transition > div:nth-child(4)')))
    
    web.execute_script("arguments[0].click()", confirmacao)
    print("Inserido no input a opção: Confirmação")
    
    botao_pesquisa = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '#reportAttendance > div > div.right.aligned.sixteen.wide.column.adjust-top-4 > button')))
    web.execute_script("arguments[0].click()", botao_pesquisa)
    print('Opções inseridas!')
    time.sleep(60)

    botao_download = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#reportAttendance > div > div.right.aligned.sixteen.wide.column.adjust-top-4 > button.icon.mini.brown.ui.button")))
    web.execute_script("arguments[0].click()", botao_download)
    # acima aperta até o botao de exportar
    wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#navbar > aside > section > ul > li.treeview.active.menu-open > ul > li:nth-child(6) > a"))).click()
    wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#app > div.row.adjust-top-6 > div > div.table-responsive.ft-scroll > table > tbody > tr:nth-child(1) > td:nth-child(7) > div"))).click()
    print("Mudando para a página de Extrações...")

    wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#app > div.row.adjust-top-6 > div > div.table-responsive.ft-scroll > table > tbody > tr:nth-child(1) > td:nth-child(7) > div"))).click()
    wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#app > div.row.adjust-top-6 > div > div.table-responsive.ft-scroll > table > tbody > tr:nth-child(1) > td:nth-child(7) > div > div > a"))).click()
    print("Processo de extração bem-sucedido. Fim da função scrap.")
    time.sleep(60)

    web.close()

webscraping()

from datetime import date, timedelta

def semana_do_mes(data):
    primeiro_dia = data.replace(day=1)
    primeiro_domingo = primeiro_dia + timedelta(days=(6 - primeiro_dia.weekday()) % 7)  # Primeiro domingo do mês

    if data < primeiro_domingo:
        return 1  # Se a data for antes do primeiro domingo, está na primeira semana

    return ((data - primeiro_domingo).days // 7) + 2  # Contagem de semanas

# Definir a data de referência e calcular a semana anterior
data_atual = date.today()
semana_atual = semana_do_mes(data_atual)

# Subtrair 7 dias para pegar a semana anterior - peguei do chatgpt slc mo bagulho dificil de entender vsf
data_anterior = data_atual- timedelta(days=7)
semana_anterior = semana_do_mes(data_anterior)

print(f"Data atual: {data_atual} - Semana do mês: {semana_atual}")
print(f"Data anterior: {data_anterior} - Semana do mês: {semana_anterior}")

def filepath(caminho_antigo, caminho_novo, novo_nome, arquivo_original):
    for file in os.listdir(caminho_antigo):
        if file.startswith(oldName):
            oldPath = os.path.join(caminho_antigo, file)
            newPath = os.path.join(caminho_novo, novo_nome)
            originalPath = os.path.join(caminho_novo, arquivo_original)

            # renomeia os bang ai
            os.rename(oldPath, newPath)

            if os.path.exists(newPath):
                df_existente = pd.read_excel(originalPath, engine='openpyxl')
            else:
                df_existente = pd.DataFrame() # cria um df vazio
            print(df_existente.head())

            df_novo = pd.read_excel(newPath, engine='openpyxl')
            data_ref = semana_anterior
            df_novo.loc[:, df_existente.columns[-1]] = data_ref

            df_final = pd.concat([ df_novo, df_existente], ignore_index=True)

            df_final.to_excel(originalPath, index=False, engine='openpyxl')

filepath(download_path, final_path, newName, originalName)
