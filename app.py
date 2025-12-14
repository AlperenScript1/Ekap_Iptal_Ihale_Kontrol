from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
import chromedriver_autoinstaller
from selenium import webdriver 
from time import sleep, time
from openpyxl import load_workbook


driver = webdriver.Chrome(); 
url = "https://ekapv2.kik.gov.tr/ekap/search";
#option = option() 
#option.add_argument("--headless")
driver.maximize_window()
driver.get(url);

path = "excel/iptal.xlsx"
wb = load_workbook(path)
ws = wb.active



try:
    iptal = [] #iptal ihalelerin tutulacağı array

    for row in ws.iter_rows(min_row=1, max_col=1, values_only=True):
        sleep(1)
        deger = row[0]
        textBox = sonucilanText = WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.CSS_SELECTOR, "dx-text-box[id='search-by-word'] input[placeholder='Ara']")))
        textBox.clear()
        textBox.send_keys(deger)
        print("find element: textBox")
        
        searchButton = sonucilanText = WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.CLASS_NAME, "search-button")))        
        searchButton.click()
        print("find element: searchButton")





except ValueError as e :
    print("Hata kodu: " + str(e))




sleep(5);
driver.quit(); 