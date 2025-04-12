from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
import time
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
# Initialize the Chrome WebDriver
driver = webdriver.Chrome(service=Service(), options=Options())


# Open the BGO website
driver.get('http://quanly.bgo.edu.vn/')

# Wait for user input to proceed
input("Press Enter after you have navigated to the 'Điểm danh' tab...")

try:
    # Wait for and click the thumbs up button
    thumbs_up = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "bgo-grid-cell-icon i.far.fa-thumbs-up.text-success"))
    )
    thumbs_up.click()
    print("Successfully clicked the thumbs up button!")

    # Ask for the grade value
    grade = input("Enter the grade value: ")

    # Wait for and find the grade input field
    grade_input = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "input.form-control[formcontrolname='btvnGrade']"))
    )
    
    # Clear existing value and enter new grade
    grade_input.clear()
    grade_input.send_keys(grade)
    grade_input.send_keys(Keys.TAB)  # Tab out to trigger any validation
    print(f"Successfully entered grade: {grade}")

    # Wait for and click the Save button
    save_button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, "button.btn.btn-info[type='submit']"))
    )
    save_button.click()

    print("Successfully clicked the Save button!")

    # Wait for and accept the alert (if it appears) driver.switchTo().alert().accept(); driver.switchTo().alert().dismiss();
    try:
        alert = WebDriverWait(driver, 10).until(EC.alert_is_present())
        alert_text = alert.text
        print(f"Alert text: {alert_text}")
        alert.accept() # Or alert.dismiss() if you need to cancel instead of accept
        print("Successfully handled the alert!")
    except Exception as alert_error:
        print(f"No alert was present or error handling alert: {alert_error}")


except Exception as e:
    print(f"An error occurred: {e}")

# Keep the browser open
input("Press Enter to close the browser...")
driver.quit()