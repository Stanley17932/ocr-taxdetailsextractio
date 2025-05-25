from selenium import webdriver
from selenium.webdriver.chrome.service import Service
import os
import time
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


service = Service(executable_path='chromedriver.exe')
driver = webdriver.Chrome(service=service)



driver.get('https://efiskalizimi-app.tatime.gov.al/invoice-check/#/verify?iic=062B5EF16CEFB69661516E0C56CECF20&tin=K92423004V&crtd=2025-02-12T08:10:38%2001:00&prc=3720.00')
    
# Wait for the page to load completely
WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.XPATH, "/html/body/app-root/app-navigation/div/nav/ul/li[2]"))
)

# Method 1: Try clicking the dropdown using the provided XPath
try:
    language_dropdown = driver.find_element(By.XPATH, "/html/body/app-root/app-navigation/div/nav/ul/li[2]")
    language_dropdown.click()
    time.sleep(1)  # Wait for dropdown to open
    
    # Look for "Anglisht" (English in Albanian) in the dropdown options
    english_option = WebDriverWait(driver, 5).until(
        EC.element_to_be_clickable((By.XPATH, "//*[contains(text(), 'Anglisht')]"))
    )
    english_option.click()
    print("Successfully changed language to English")
    
except Exception as e1:
    print(f"Method 1 failed: {e1}")

time.sleep(20)
driver.quit()