from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import time

url = 'http://quanly.bgo.edu.vn/class-attendances'
chrome_options = Options()
chrome_options.add_argument('--start-maximized')
driver = webdriver.Chrome(options=chrome_options)
driver.get(url)

time.sleep(10)  
driver.execute_script("document.body.style.zoom='67%'")
time.sleep(5)  

wait = WebDriverWait(driver, 20)

try:
    # Use the original working XPath for search
    find = wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/bgo-root/div/div/bgo-main-layout/bgo-class-attendances/div/div[4]/bgo-grid/div/div[1]/div[6]/div/input")))
    find.send_keys('053')
    time.sleep(5)  # Wait for search results

    # First dropdown
    select_1 = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/bgo-root/div/div/bgo-main-layout/bgo-class-attendances/div/div[4]/bgo-grid/div/div[2]/div/ag-grid-angular/div/div[1]/div/div[3]/div[2]/div/div/div[1]/div[3]/bgo-grid-dropdown/select/option[2]")))
    select_1.click()
    print('Xong 1')
    time.sleep(5)

    # Second dropdown
    select_2 = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/bgo-root/div/div/bgo-main-layout/bgo-class-attendances/div/div[4]/bgo-grid/div/div[2]/div/ag-grid-angular/div/div[1]/div/div[3]/div[2]/div/div/div[1]/div[5]/bgo-grid-dropdown/select/option[2]")))
    select_2.click()
    print('Xong 2')
    time.sleep(5)

    # Third dropdown
    select_3 = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/bgo-root/div/div/bgo-main-layout/bgo-class-attendances/div/div[4]/bgo-grid/div/div[2]/div/ag-grid-angular/div/div[1]/div/div[3]/div[2]/div/div/div[1]/div[3]/bgo-grid-dropdown/select/option[2]")))
    select_3.click()
    print('Xong 3')
    time.sleep(5)

except Exception as e:
    print(f"An error occurred: {str(e)}")
    driver.save_screenshot("error_screenshot.png")
    print("Screenshot saved as error_screenshot.png")



