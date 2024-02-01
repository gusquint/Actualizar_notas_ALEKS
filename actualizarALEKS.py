import gspread,time,os
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import date
import pandas as pd
from json import dumps
from httplib2 import Http
from decouple import config


def open_ALEKS(lugar,ussername,password):
    if lugar=="colegio":
        PATH="C:/Users/gquintero/Desktop/python/chromedriver.exe"
    else:
        PATH="C:/Users/PC/Desktop/python/chromedriver.exe" 
    driver=webdriver.Chrome(PATH)   
    driver.get("https://latam.aleks.com/login")
    search=driver.find_element(By.ID,value="login_name_alone")
    search.send_keys(ussername)
    search=driver.find_element(By.ID,value="login_pass_alone")
    search.send_keys(password)
    search.send_keys(Keys.RETURN)
    return driver



def close_ALEKS(driver,seconds):
    time.sleep(seconds)
    driver.close()


def download_report(driver,lugar):    
    if lugar=="colegio":
        download_folder="C:/Users/gquintero/Downloads"
    else:
        download_folder="C:/Users/PC/Downloads"
    reports = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "navigation_report")))
    reports.click()
    search=WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,"//span[@style='display: inline-block;']")))
    tag_a=search.find_element(By.TAG_NAME,"a")
    tag_span=tag_a.find_element(By.TAG_NAME,"span")
    next_tag_span=tag_span.find_elements(By.TAG_NAME,"span")[1] 
    tag_div=next_tag_span.find_element(By.TAG_NAME,"div")
    custom_reports=tag_div.find_element(By.TAG_NAME,"span")
    time.sleep(1)
    custom_reports.click()
    templates= WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.LINK_TEXT, "Templates")))
    templates.click()
    acumulado=driver.find_element(By.XPATH, value="//td[@class='cr_table_class left td_1']")
    schedule_report=acumulado.find_elements(By.TAG_NAME, "a")[1]
    schedule_report.click()
    report_name = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "schedule_name")))
    today=date.today().strftime("%d/%m/%Y")
    report_name.send_keys(today)
    minutos_tick=driver.find_element(By.ID,"cr_duration_format_num_0_m_3")
    minutos_tick.click()
    nov_a=driver.find_element(By.ID,"cbx_class_35")
    nov_a.click()
    nov_b=driver.find_element(By.ID,"cbx_class_27")
    nov_b.click()
    nov_c=driver.find_element(By.ID,"cbx_class_18")
    nov_c.click()
    nov_d=driver.find_element(By.ID,"cbx_class_10")
    nov_d.click()
    nov_e=driver.find_element(By.ID,"cbx_class_5")
    nov_e.click()
    
    schedule_report_button=driver.find_element(By.ID,"button_3")
    schedule_report_button.click()
    
    #revisa la carpeta de descargas y encuentra el archivo descargado
    entries_old=os.listdir(download_folder)
    while True:
        tr_tag=WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID,"div_row_1")))
        try:
            download_report=tr_tag.find_element(By.LINK_TEXT,"Download Report")
            break
        except:
            refresh_table=WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.LINK_TEXT, "Refresh Table")))
            refresh_table.click()
    download_report.click()
    time.sleep(5)
    entries_new=os.listdir(download_folder)
    for entrie in entries_new:
        if entrie not in entries_old:
            return entrie

def fix_data(group,dataframe,google_sheet):   
    hoja=google_sheet.worksheet(group)
    progress=dataframe[7:]["Unnamed: 2"]
    minutes=dataframe[7:]["Unnamed: 3"]
    row=9
    for i in progress:
        if i=="-":
            hoja.update_cell(row,3,0)
        else:
            hoja.update_cell(row,3,round(float(i),2))
        row+=1
    row=9
    for i in minutes:
        hoja.update_cell(row,4,round(float(i)))
        row+=1
    print(f"actualizado {group}")


def update_driver(file,lugar):
    if lugar=="colegio":
        sa=gspread.service_account(filename="C:/Users/gquintero/Desktop/python/pythontest-361720-6e048b13eddb.json")   
    else:
        sa=gspread.service_account(filename="C:/Users/PC/Desktop/python/pythontest-361720-6e048b13eddb.json")

    with pd.ExcelFile(file) as xls:
        df1 = pd.read_excel(xls, '9a 3rd')
        df2 = pd.read_excel(xls, '9b 3rd')
        df3 = pd.read_excel(xls, '9c 3rd')
        df4 = pd.read_excel(xls, '9d 3rd')
        df5 = pd.read_excel(xls, '9e 3rd')
    
    #tengo que esperar 60 segundos para resetar los usos de escritura por minuto de gspread
    google_sheet=sa.open("reporte ALEKS (no editar)")     
    fix_data("9a",df1,google_sheet)
    print("Esperando 1 minuto")   
    time.sleep(60) 
    fix_data("9b",df2,google_sheet)
    print("Esperando 1 minuto")   
    time.sleep(60) 
    fix_data("9c",df3,google_sheet)
    print("Esperando 1 minuto")   
    time.sleep(60) 
    fix_data("9d",df4,google_sheet)
    print("Esperando 1 minuto")   
    time.sleep(60) 
    fix_data("9e",df5,google_sheet)


def send_message():
    #antes de ejecutar el comando, se debe crear el bot.
    #para mas detalles, visitar el link
    #https://developers.google.com/chat/how-tos/webhooks
    #El icono del bot se puede obtener de esta pagina
    #https://www.google.com/s2/favicons?domain_url=@@@@@
    #remplazar los @@@@@ por el url de la pagina con el icono a elegir
    url = "https://chat.googleapis.com/v1/spaces/AAAAAzaTnxw/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=teoYH11HNJgV2U1IMh1GjQjPsYW54ZvxX9p_DqKv95A%3D"
    bot_message = {'text': f'ALEKS ha sido actualizado con datos del {date.today()}'}
    message_headers = {'Content-Type': 'application/json; charset=UTF-8'}
    http_obj = Http()
    response = http_obj.request(uri=url,method='POST',headers=message_headers,body=dumps(bot_message),)
    #print(response)


def main():
    lugar="colegio"
    driver=open_ALEKS(lugar,config("aleks_usuario"),config("aleks_password"))
    file_name=download_report(driver,lugar)    
    close_ALEKS(driver,3)
    file=f"C:/Users/gquintero/Downloads/{file_name}"
    update_driver(file,lugar)    
    send_message()



if __name__ == "__main__":
    main()