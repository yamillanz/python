import gc
from openpyxl import Workbook
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException, ElementClickInterceptedException

PATH = 'C:\Program Files (x86)\chromedriver.exe'

driver = webdriver.Chrome(PATH)

driver.get('https://www.empleate.com/empresarial/portal/')

username = driver.find_element_by_id('user')
username.clear()
username.send_keys('USER')
password = driver.find_element_by_id('pass')
password.clear()
password.send_keys('PASS')
driver.find_element_by_class_name('btn-loading').click()

#Boton de ventana emergente
while(True):
    try:
        if driver.find_element_by_id('nothanks'):
            driver.find_element_by_id('nothanks').click()
            break
    except:
        continue

#Link a pagina "Buscar talentos"
driver.find_element_by_xpath('//a[contains(@href, "/talentos/buscador_talentos/")]').click()

#Click en formulario input de "Area"
while(True):
    try:
        if driver.find_element_by_xpath("//li[contains(@class, 'search-field')]"):
            driver.find_element_by_xpath("//li[contains(@class, 'search-field')]").click()
            break
    except:
        continue

driver.find_element_by_xpath("//li[contains(text(), 'Administración y Gerencia')]").click()
driver.find_element_by_xpath("//li[contains(@class, 'search-field')]").click()
driver.find_element_by_xpath("//li[contains(text(), 'Salud')]").click()

#Boton de buscar
btn = driver.find_element_by_id('button-publicar')
driver.execute_script("arguments[0].click();", btn)

wb = Workbook()
destExcelFilename = 'data.xlsx'

ws1 = wb.create_sheet(title='Datos')

ws1['A1'] = 'Nombre'
ws1['B1'] = 'Correo'
ws1['C1'] = 'Telefono 1'
ws1['D1'] = 'Telefono 2'
ws1['E1'] = 'Telefono 3'

rowNumber = 2
try:
    while(True):
        #Obtencion de todos los links a empleados en la pagina
        names = driver.find_elements_by_class_name('titulo_postulados')

        cvCodes = []

        # Se llama al Garbage Collector
        gc.collect()

        url = 'https://empresas.empleate.com'
        #Se quita el dominio del link
        for name in names:
            link = name.get_attribute("href")
            link = link.replace(url, '')
            cvCodes.append(link)

        #Recorrido por todos los resultados de una pagina
        for code in cvCodes:
            try:
                emp = driver.find_element_by_xpath('//a[contains(@href, "'+code+'") and contains(@class, "titulo_postulados")]')
                emp.click()
            except (NoSuchElementException, ElementClickInterceptedException):
                print("No se encontro el CV:"+code)
                continue
            try:
                name = driver.find_element_by_xpath("//h4[contains(@style, '#0c7bb2')]").get_attribute("innerText")
                #Proceso para quitar caracteres no requeridos
                param = '\n'
                if param in name:
                    index = name.index(param) -1
                    name = name[0:index]
            except NoSuchElementException:
                name = 'No encontrado'
            try:
                email = driver.find_element_by_class_name('na').get_attribute("innerHTML")
            except NoSuchElementException:
                email = 'No encontrado'
            try:
                phones = driver.find_elements_by_class_name('blue')
            except NoSuchElementException:
                phones = []

            ws1['A' + str(rowNumber)] = name
            ws1['B' + str(rowNumber)] = email
            ws1['C' + str(rowNumber)] = phones[1].get_attribute("innerHTML") if (phones != None) and (len(phones) >= 2) else 0
            ws1['D' + str(rowNumber)] = phones[3].get_attribute("innerHTML") if (phones != None) and (len(phones) >= 4) else 0
            ws1['E' + str(rowNumber)] = phones[5].get_attribute("innerHTML") if (phones != None) and (len(phones) >= 6) else 0

            driver.back()

            rowNumber = rowNumber + 1

            print('Registro número : ' + str(rowNumber - 1))
            print('pulsa ctrl+c para terminar el proceso...')

        #Click en siguiente pagina
        try:
            nextPage = driver.find_element_by_class_name('icon-angle-right')
            nextPage.find_element_by_xpath('..').click()
        except (NoSuchElementException, ElementClickInterceptedException):
            print("Llegada a la ultima pagina")
            break

except KeyboardInterrupt:
    wb.save(filename=destExcelFilename)

else:
    wb.save(filename=destExcelFilename)