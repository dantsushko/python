import time
from selenium import webdriver
import chromedriver_binary
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import xlwt

link = "https://widget.profitbase.ru/?pbApiKey=aed5ba22faea5061b5664ca46671aab7&pbSubdomain=pb1382&pbBaseDomain=profitbase.ru&referer=https%3A%2F%2Fsolnechnyy.sinara-development.ru&v=1#/houses/13632/presentation/small/"

options = webdriver.ChromeOptions()
options.add_argument('headless')
driver = webdriver.Chrome(chrome_options=options)
driver.set_window_position(0, 0)
driver.set_window_size(1600, 1200)

driver.get(link)
print("Connecting to server...")
time.sleep(5)
cells = driver.find_elements(By.CLASS_NAME, 'js-cell-property')

wb = xlwt.Workbook()
ws = wb.add_sheet('output')
outfile_name = 'out.xls'


ws.write(0, 0, "Flat number")
ws.write(0, 1, "Floor number")
ws.write(0, 2, "Room number")
i = 1
print ("Processing data... Wait for 2 minutes")

for cell in cells:
    if ("apartment-cell_empty" not in cell.get_attribute("class")):
        cell.click()
        table = driver.find_element(By.CLASS_NAME, 'apartment-info__table')
        rows = table.find_elements(By.TAG_NAME, 'tr')
        apartment_number = rows[0].find_element(By.TAG_NAME, 'td')
        ws.write(i,0,apartment_number.get_attribute("textContent").strip())
        floor_number= rows[1].find_element(By.TAG_NAME, 'td')
        ws.write(i,1,floor_number.get_attribute("textContent").strip())
        if cell.get_attribute("textContent").strip().isdigit():
            room_number= rows[2].find_element(By.TAG_NAME, 'td')
            ws.write(i,2, int(room_number.get_attribute("textContent").strip()))
        else:
            ws.write(i,2, 0)
        i += 1
        if i == 10:
            break


print("Data is written!")
wb.save(outfile_name)
print("Data is saved in", outfile_name)
driver.close()
quit()
