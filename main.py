import time
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from openpyxl import Workbook, load_workbook

# Jorge Orellana - 0909-17-7161
# https://github.com/L2Ernestoo

driver = webdriver.Chrome('webdriver/chromedriver_91.exe')

url = 'https://www.demoblaze.com/index.html'

driver.get(url)
time.sleep(3)
articulo = driver.find_element_by_link_text('Samsung galaxy s6')
articulo.click()

try:
    #Agregar Carrito
    element = WebDriverWait(driver, 5).until(
        EC.presence_of_element_located((By.LINK_TEXT, "Add to cart"))
    )
    element.click()
    time.sleep(3) #Espera

    #Aceptar Alert()
    WebDriverWait(driver, 10).until(EC.alert_is_present())
    driver.switch_to.alert.accept()

    #Ir A Carrito
    carrito = WebDriverWait(driver, 5).until(
        EC.presence_of_element_located((By.LINK_TEXT, "Cart"))
    )
    carrito.click()
    time.sleep(3)

    #Click en Boton Poner Orden
    orden = WebDriverWait(driver, 5).until(
        EC.presence_of_element_located((By.XPATH, '//button[text()="Place Order"]'))
    )
    orden.click()

    time.sleep(2)

    data = load_workbook('info.xlsx')
    de = data.active

    name = driver.find_element_by_id('name')
    name.send_keys(de['A2'].value)

    country = driver.find_element_by_id('country')
    country.send_keys(de['B2'].value)

    city = driver.find_element_by_id('city')
    city.send_keys(de['C2'].value)

    card = driver.find_element_by_id('card')
    card.send_keys(de['D2'].value)

    month = driver.find_element_by_id('month')
    month.send_keys(de['E2'].value)

    year = driver.find_element_by_id('year')
    year.send_keys(de['F2'].value)

    time.sleep(3)

    comprar = WebDriverWait(driver, 5).until(
        EC.presence_of_element_located((By.XPATH, '//button[text()="Purchase"]'))
    )
    comprar.click()
    titulo = WebDriverWait(driver, 2).until(
        EC.presence_of_element_located((By.XPATH, '//div[@class="sweet-alert  showSweetAlert visible"]/h2'))
    )
    driver.execute_script("arguments[0].innerText = 'Inteligencia Artificial - Ing. Erick Alvarez'", titulo)

    time.sleep(10)
except:
    print("Error Reinicie y compruebe su velocidad de internet.")

