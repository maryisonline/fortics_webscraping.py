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

def webscraping():

    oldName = 'initial_file'
    newName = 'BASE_FORTICS'

    download_path = r'C:\Users\Benvista\Desktop\bases_download'
    file_path = r'C:\Users\Benvista\Desktop\BASES\TABULAÇÃO'

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
    
    wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '#reportAttendance > div > div.right.aligned.sixteen.wide.column.adjust-top-4 > button'))).click()

webscraping()
