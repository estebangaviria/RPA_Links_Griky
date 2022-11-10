from lib2to3.pgen2 import driver
from operator import index
from tabnanny import check
from typing import final
from unicodedata import name
from urllib import response
import pandas
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
import requests
#import lxml.html as html
from selenium.webdriver.support.events import EventFiringWebDriver, AbstractEventListener
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from datetime import date 


def robot(url, instancia):
        
    #Lee usuario y contraseña:
    credenciales_excel = r"D:\PROYECTOS\RPA-LOGIN-GRIKY\credenciales.xlsx"
    df_credenciales = pandas.read_excel(credenciales_excel)
    user = df_credenciales["username"][0]
    psw = df_credenciales["password"][0]


    driver = webdriver.Chrome()

    # Maximizar pantalla
    driver.maximize_window()
    time.sleep(2)
    driver.get(url)
    time.sleep(5)

    #Quitar widget eanx
    try:
        driver.switch_to.frame(1)
        driver.find_element(By.CLASS_NAME, "widget")
        driver.find_element(By.XPATH, "/html/body/div/div[1]/div/div/button").click()
        driver.switch_to.default_content()
    except:
        driver.switch_to.default_content()
        pass

    #Quitar widget griky
    try:
        driver.switch_to.frame(2)
        driver.find_element(By.CLASS_NAME, "widget")
        driver.find_element(By.XPATH, "/html/body/div/div[1]/div/div/button").click()
        driver.switch_to.default_content()
    except:
        driver.switch_to.default_content()
        pass

    # Acciones:
    #driver.find_element(By.CSS_SELECTOR, "#login-form > article > div > div.email-login-link > a").click()
    try:
        driver.find_element(By.CSS_SELECTOR, "#login-form > article > div > div.email-login-link > a").click()
    except Exception:
        pass

    # Login:
    driver.find_element(By.CSS_SELECTOR, "#login-email").send_keys(user)
    # time.sleep(2)
    driver.find_element(By.CSS_SELECTOR, "#login-password").send_keys(psw)
    time.sleep(1)

    mensaje = ""
    
    
    date = time.strftime('%m-%d-%Y')
    current_time = time.strftime('%H:%M:%S')
    
    try:
        inicio = time.time()
        driver.find_element(By.CSS_SELECTOR, "#login > button").click()
        time.sleep(1)
        driver.find_element(By.CLASS_NAME, "secondary")
        fin = time.time()
        tiempo_carga = fin-inicio
        mensaje = "Inicio de sesión exitoso" + ","+ instancia +  "," +date + "," +current_time + "," +str(tiempo_carga)
    except: 
        mensaje = "Inicio de sesión fallido" + ","+ instancia +  "," +date + "," +current_time + "," 

    #time.sleep(2)

    # Cerrar las acciones:
    # driver.quit()
    return mensaje

def ciclo():
    links_excel = r"D:\PROYECTOS\RPA-LOGIN-GRIKY\links.xlsx"
    df_links = pandas.read_excel(links_excel)
    mensaje =""
    file = open("textfile.txt","w+")
    file.write("Mensaje, Instancia, Fecha(mm:dd:aaaa), Hora(h:m:s), Tiempo de carga(seg)")
    today = time.strftime('%m-%d-%Y_%H-%M-%S')
    for i in df_links.index:
        #print(str(df_links["Instance"][i])+ str(df_links["url"][i]))
        mensaje = robot(url= df_links["url"][i], instancia=df_links["Instance"][i])
        print(mensaje)
        file.write('\n' + mensaje)
        
        #file.writelines('\n'.join(mensaje))
    
    file.close()
    #file.to_csv("D:\PROYECTOS\RPA-LOGIN-GRIKY\Status"+today+".xlsx", index=False)
    read_file = pandas.read_csv(r'D:\PROYECTOS\RPA-LOGIN-GRIKY\textfile.txt', encoding='latin-1',index_col=0)
    read_file.to_csv('D:\PROYECTOS\RPA-LOGIN-GRIKY\\test_'+today+".csv", index=False)

ciclo()
#robot(url='https://umaplus.uma.edu.pe/', instancia = "UMA")


