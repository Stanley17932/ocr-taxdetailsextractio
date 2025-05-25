from selenium import webdriver
from selenium.webdriver.chrome.service import Service
import os
import time
import re
import pandas as pd
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

service = Service(executable_path='chromedriver.exe')
driver = webdriver.Chrome(service=service)
driver.get('https://efiskalizimi-app.tatime.gov.al/invoice-check/#/verify?iic=716207BB76D4B8A52AAE06952D41C363&tin=L41702011R&crtd=2025-02-10T17:40:55%2001:00&ord=7034&bu=el359ox381&cr=nw619tf113&sw=cc302yz654&prc=1914.00')
   
# Wait for the page to load completely
WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.XPATH, "/html/body/app-root/app-navigation/div/nav/ul/li[2]"))
)

print("Page loaded. Attempting to change language...")

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
    print(f"Language change failed: {e1}")

# Wait for language change to take effect and page to reload
time.sleep(5)
print("Extracting invoice data...")

# Initialize data dictionary
invoice_data = {
    'Invoice Number': '',
    'Grand Total': '',
    'Business Name': '',
    'Issuer Tax Number': ''
}

try:
    # Extract Invoice Number - Look for the specific pattern "Invoice XXXXX/YYYY"
    try:
        # First try the main heading area where invoice number should be displayed
        invoice_selectors = [
            "/html/body/app-root/app-verify-invoice/div/section[1]/div/div[1]/h4",
            "//h4[contains(@class, 'invoice-title')]",
            "//h4",
            "//*[contains(text(), 'Invoice')]",
            "//*[contains(text(), '19027')]"
        ]
        
        invoice_found = False
        for selector in invoice_selectors:
            try:
                invoice_element = WebDriverWait(driver, 3).until(
                    EC.presence_of_element_located((By.XPATH, selector))
                )
                invoice_text = invoice_element.text.strip()
                print(f"Found text with selector '{selector}': '{invoice_text}'")
                
                # Look for pattern like "Invoice 19027/2025" or just "19027/2025"
                invoice_match = re.search(r'(?:Invoice\s+)?(\d+/\d+)', invoice_text, re.IGNORECASE)
                if invoice_match:
                    invoice_data['Invoice Number'] = f"Invoice {invoice_match.group(1)}"
                    print(f"Invoice Number: {invoice_data['Invoice Number']}")
                    invoice_found = True
                    break
                elif "19027" in invoice_text or "invoice" in invoice_text.lower():
                    # If contains invoice-related text but doesn't match pattern, use as is
                    invoice_data['Invoice Number'] = invoice_text
                    print(f"Invoice Number (fallback): {invoice_data['Invoice Number']}")
                    invoice_found = True
                    break
            except:
                continue
        
        if not invoice_found:
            print("Invoice number not found with standard selectors, trying broader search...")
            # Try to find any element containing "19027"
            try:
                all_elements = driver.find_elements(By.XPATH, "//*[contains(text(), '19027')]")
                for element in all_elements:
                    text = element.text.strip()
                    if text and len(text) < 50:  # Reasonable length for invoice number
                        invoice_data['Invoice Number'] = text
                        print(f"Invoice Number (broad search): {invoice_data['Invoice Number']}")
                        break
            except Exception as e:
                print(f"Broad invoice search failed: {e}")
                
    except Exception as e:
        print(f"Failed to extract Invoice Number: {e}")

    # Extract Grand Total - Look for the large amount display "3 720,00 LEK"
    try:
        # Try multiple selectors for the grand total amount
        total_selectors = [
            "//*[contains(text(), '3 720') or contains(text(), '3720')]",
            "//*[contains(text(), 'LEK')]",
            "//h1[contains(@class, 'invoice-title')]",
            "//div[contains(@class, 'invoice-amount')]",
            "//*[text()[contains(., '720')]]"
        ]
        
        total_found = False
        for selector in total_selectors:
            try:
                total_elements = driver.find_elements(By.XPATH, selector)
                for element in total_elements:
                    text = element.text.strip()
                    print(f"Found potential total text: '{text}'")
                    
                    # Look for amount patterns like "3 720,00 LEK" or "3720.00"
                    total_match = re.search(r'(\d[\s,.]?\d{3}[,.]\d{2})\s*(?:LEK|$)', text)
                    if total_match:
                        invoice_data['Grand Total'] = total_match.group(1).replace(' ', '').replace(',', '.')
                        print(f"Grand Total: {invoice_data['Grand Total']}")
                        total_found = True
                        break
                    elif re.search(r'\d+[,.]\d{2}', text) and 'LEK' in text:
                        # Extract any monetary amount with LEK
                        amount_match = re.search(r'(\d+[,.]\d{2})', text)
                        if amount_match:
                            invoice_data['Grand Total'] = amount_match.group(1)
                            print(f"Grand Total (pattern match): {invoice_data['Grand Total']}")
                            total_found = True
                            break
                if total_found:
                    break
            except:
                continue
                
        if not total_found:
            # Try extracting from URL parameter as fallback
            try:
                url = driver.current_url
                if 'prc=' in url:
                    price_match = re.search(r'prc=(\d+\.?\d*)', url)
                    if price_match:
                        invoice_data['Grand Total'] = price_match.group(1)
                        print(f"Grand Total (from URL): {invoice_data['Grand Total']}")
            except Exception as e:
                print(f"URL extraction failed: {e}")
                
    except Exception as e:
        print(f"Failed to extract Grand Total: {e}")

    # Extract Business Name 
    try:
        business_selectors = [
            "//div[contains(@class, 'invoice-basic-info--business-name')]",
            # "//*[contains(text(), 'KOSMONTE')]",
            # "//*[contains(text(), 'FOODS')]",
            "//div[contains(@class, 'business-name')]"
        ]
        
        business_found = False
        for selector in business_selectors:
            try:
                business_elements = driver.find_elements(By.XPATH, selector)
                for element in business_elements:
                    text = element.text.strip()
                    print(f"Found potential business text: '{text}'")
                    
                    if text and ('KOSMONTE' in text.upper() or 'FOODS' in text.upper() or len(text.split()) <= 5):
                        invoice_data['Business Name'] = re.sub(r'\s+', ' ', text)
                        print(f"Business Name: {invoice_data['Business Name']}")
                        business_found = True
                        break
                if business_found:
                    break
            except:
                continue
                
    except Exception as e:
        print(f"Failed to extract Business Name: {e}")

    # Extract Issuer Tax Number - Look for "K92423004V"
    try:
        # Look for the specific tax number pattern
        tax_selectors = [
            "//*[contains(text(), 'K92423004V')]",
            "//div[contains(@class, 'form-group')]//p[contains(text(), 'K')]",
            "//*[text()[contains(., 'K92423004V')]]",
            "//label[contains(text(), 'Issuer Tax Number')]/following-sibling::*",
            "//label[contains(text(), 'Tax Number')]/following-sibling::*"
        ]
        
        tax_found = False
        for selector in tax_selectors:
            try:
                tax_elements = driver.find_elements(By.XPATH, selector)
                for element in tax_elements:
                    text = element.text.strip()
                    print(f"Found potential tax text: '{text}'")
                    
                    # Look for Albanian tax number pattern (letter followed by numbers and letter)
                    tax_match = re.search(r'([A-Z]\d{8}[A-Z])', text)
                    if tax_match:
                        invoice_data['Issuer Tax Number'] = tax_match.group(1)
                        print(f"Issuer Tax Number: {invoice_data['Issuer Tax Number']}")
                        tax_found = True
                        break
                    elif 'K92423004V' in text:
                        invoice_data['Issuer Tax Number'] = 'K92423004V'
                        print(f"Issuer Tax Number (direct match): {invoice_data['Issuer Tax Number']}")
                        tax_found = True
                        break
                if tax_found:
                    break
            except:
                continue
        
        # If still not found, try broader search in form groups
        if not tax_found:
            try:
                form_groups = driver.find_elements(By.CLASS_NAME, "form-group")
                for group in form_groups:
                    text = group.text
                    if 'tax number' in text.lower() or 'K92423004V' in text:
                        tax_match = re.search(r'([A-Z]\d{8}[A-Z])', text)
                        if tax_match:
                            invoice_data['Issuer Tax Number'] = tax_match.group(1)
                            print(f"Issuer Tax Number (form group): {invoice_data['Issuer Tax Number']}")
                            break
            except Exception as e2:
                print(f"Form group tax search failed: {e2}")
                
    except Exception as e:
        print(f"Failed to extract Issuer Tax Number: {e}")

except Exception as main_error:
    print(f"Main data extraction error: {main_error}")

# Create DataFrame and save to Excel
try:
    df = pd.DataFrame([invoice_data])
    
    # Generate filename with timestamp
    timestamp = time.strftime("%Y%m%d_%H%M%S")
    filename = f"invoice_data_{timestamp}.xlsx"
    
    # Save to Excel
    df.to_excel(filename, index=False)
    print(f"\nData saved to Excel file: {filename}")
    
    # Display extracted data
    print("\n" + "="*50)
    print("EXTRACTED INVOICE DATA:")
    print("="*50)
    for key, value in invoice_data.items():
        print(f"{key}: {value}")
    print("="*50)
    
except Exception as excel_error:
    print(f"Failed to save to Excel: {excel_error}")

# Wait to observe results
time.sleep(20)
driver.quit()