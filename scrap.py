import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

def accept_popup(driver):
    try:
        # Check if the popup exists
        popup_button = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, "//input[@class='btn' and @value='Accept']")))
        
        # If the popup exists, click on the accept button
        popup_button.click()
        
    except Exception as e:
        print("Popup not found or error occurred while accepting popup:", e)

def get_zoning_info(driver, parcel_id):
    try:
        # Find the search bar by its name attribute
        search_bar = driver.find_element(By.NAME, "search")
        search_bar.clear()  # Clear any existing value in the search bar
        search_bar.send_keys(parcel_id)  # Enter the parcel ID
        time.sleep(2)  # Wait for autocomplete suggestions to appear

        # Click on the autocomplete suggestion
        autocomplete_item = driver.find_element(By.XPATH, "//li[@class='autocomplete-item']")
        autocomplete_item.click()

        # Wait for results to load
        time.sleep(5)

        # Extract zoning information
        zoning_element = driver.find_element(By.CLASS_NAME, "zone-link")
        zoning_info = zoning_element.text

        return zoning_info

    except Exception as e:
        print("An error occurred:", e)
        return None

# Read parcel IDs from Excel file, get zoning info, and write back to Excel file
try:
    # Initialize Chrome driver
    driver = webdriver.Chrome()
    driver.get("https://map.gridics.com/us/fl/miami")
    time.sleep(2)
    
    # Accept popup only the first time
    accept_popup(driver)

    # Read parcel IDs from Excel file
    df = pd.read_excel(r'C:\Users\Himu\Desktop\New folder (3)\book2.xlsx')
    parcel_ids = df['Parcel ID'].tolist()
    zoning_ids = []  # Initialize zoning_ids list

    # Loop through parcel IDs
    for parcel_id in parcel_ids:
        zoning_id = get_zoning_info(driver, parcel_id)
        zoning_ids.append(zoning_id) if zoning_id is not None else zoning_ids.append("")  # Append zoning_id to zoning_ids list

    # Add zoning_ids list as a new column named "Zoning ID" to the DataFrame
    df['Zoning ID'] = zoning_ids

    # Write the DataFrame back to the Excel file
    df.to_excel(r'C:\Users\Himu\Desktop\New folder (3)\book2.xlsx', index=False)

    print("Zoning IDs saved to Excel file.")

except FileNotFoundError:
    print("Error occurred while reading parcel IDs from Excel: No such file or directory.")

finally:
    # Quit the driver
    driver.quit()
