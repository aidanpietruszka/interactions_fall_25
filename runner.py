from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import pandas as pd
import time

building = "North Avenue East" # Replace if in different building (ex. North Avenue North/West/South)

def wait_for_sign_in():
    while True:
        response = input("Sign in, then press y and enter:\n").strip().lower()
        if response == 'y':
            print("Beginning form submissions...")
            break
        elif response == 'n':
            print("Please sign in and then type 'y' to continue.")
        else:
            print("Invalid input. Please type 'y' or 'n'.")

# Load resident data from Excel
def load_data(file_path):
    """
    Load resident names from an Excel file.
    Ensure the Excel file has a column named correctly, with no errors in resident names.
    """
    return pd.read_excel(file_path)

# Automate form filling using Selenium
def automate_form(data, form_url):
    """
    Automate navigating to a form URL and filling in resident names.
    """
    driver = webdriver.Chrome()  # Make sure 'chromedriver' is in your PATH with the correct version
    driver.get(form_url)  # Open the form URL
    time.sleep(2) # Allow page load
    wait_for_sign_in()


    for index, row in data.iterrows():
        try:
            driver.get(form_url)
            time.sleep(2) # Allow page load
            res_name = row['Resident Name']
            #date = (str(row['Date'])[6:10] + '/' + str(row['Date'])[0:4]).replace('-', '/')
            date = (str(row['Date'])[0:5] + '/' + str(row['Date'])[6:10]).replace('-', '/')
            description = row["Description"]
            theme = row['Theme']
            
            search_fields = driver.find_elements(By.ID, "forms-tag-search-input-div")
            resident_name_field = search_fields[0]
            building_field = search_fields[1]
            student_population_field = search_fields[2]
            themes_field = search_fields[3]

            resident_name_field.send_keys(res_name)
            time.sleep(4)
            search_names = driver.find_elements(By.CLASS_NAME, "forms-subscriptions-search-result-row")
            if len(search_names) == 0:
                continue
            search_names[0].click()

            building_field.send_keys(building)
            time.sleep(2)
            search_buildings = driver.find_elements(By.CLASS_NAME, "forms-subscriptions-search-result-row")
            if len(search_names) == 0:
                continue
            search_buildings[0].click()

            student_population_field.send_keys("Upperclass")
            time.sleep(2)
            search_populations = driver.find_elements(By.CLASS_NAME, "forms-subscriptions-search-result-row")
            if len(search_populations) == 0:
                continue
            search_populations[0].click()
            time.sleep(1)

            date_field = driver.find_element(By.CLASS_NAME, "elm-datepicker--input")
            date_field.send_keys(date)
            time.sleep(1)

            description_field = driver.find_element(By.ID, "desc_resp_sub_YmoBEE")
            description_field.send_keys(description)
            time.sleep(1)

            themes_field.send_keys(theme)
            time.sleep(2)
            search_themes = driver.find_elements(By.CLASS_NAME, "forms-subscriptions-search-result-row")
            if len(search_themes) == 0:
                continue
            search_themes[0].click()
            
            time.sleep(2)

            if 'Resident opted out' in description:
                optout_button = driver.find_element(By.XPATH, "//input[contains(@aria-label, 'Yes')]")
            else:
                optout_button = driver.find_element(By.XPATH, "//input[@aria-label='No']")
            optout_button.click()

            time.sleep(2)

            submit_button = driver.find_element(By.XPATH, "//button[contains(@class, 'forms-submit-btn')]")
            submit_button.click()

            print(f"Submitted: {row['Resident Name']}")
            time.sleep(5)  # Allow form to process

        except Exception as e:
            print(f"Error processing {row['Resident Name']}: {e}")
    
    # Close the browser
    driver.quit()

# Main function
def main():
    file_path = "residents.xlsx"  # Path to your Excel file
    form_url = "https://roompact.com/forms/#/form/7odwkY"  # Replace with the actual form URL
    
    data = load_data(file_path)
    print("Loaded data. Beginning interaction submission...")
    
    automate_form(data, form_url)
    print("Submission complete!")

if __name__ == "__main__":
    main()
