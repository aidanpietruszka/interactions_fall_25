import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

BUILDING = "North Avenue East"  # Replace if in a different building

def wait_for_sign_in():
    """Wait for the user to sign in by prompting until 'y' is entered."""
    while True:
        response = input("Sign in, then press y and enter:\n").strip().lower()
        if response == "y":
            print("Beginning form submissions...")
            break
        elif response == "n":
            print("Please sign in and then type 'y' to continue.")
        else:
            print("Invalid input. Please type 'y' or 'n'.")


def clear_field(field):
    """
    Clear a given input field.
    
    First, it attempts to use the built-in clear() method.
    If that fails, it sends backspace keystrokes to remove the text.
    """
    try:
        field.clear()
    except Exception:
        current_value = field.get_attribute("value")
        if current_value:
            field.send_keys(Keys.END)
            for _ in range(len(current_value)):
                field.send_keys(Keys.BACKSPACE)


def clear_submission(driver):
    """
    Clears all fields in the submission form.
    
    It locates each of the input fields (search fields, date field,
    and description field) and clears them.
    """
    try:
        # Clear search fields
        search_fields = driver.find_elements(By.ID, "forms-tag-search-input-div")
        for field in search_fields:
            clear_field(field)
        # Clear the date field
        date_field = driver.find_element(By.CLASS_NAME, "elm-datepicker--input")
        clear_field(date_field)
        # Clear the description field
        description_field = driver.find_element(By.ID, "desc_resp_sub_YmoBEE")
        clear_field(description_field)
    except Exception as error:
        print("Error clearing submission fields:", error)


def load_data(file_path):
    """
    Load resident data from an Excel file.
    
    The Excel file must contain the columns 'Resident Name', 'Date',
    'Theme', and 'Description'.
    """
    return pd.read_excel(file_path)


def automate_form(data, form_url):
    """
    Automate form submissions using Selenium.
    
    Iterates over the data rows and fills in the form fields.
    If a submission fails (e.g., the confirmation popup is not detected),
    the resident name is added to the error list.
    """
    driver = webdriver.Chrome()  # Ensure 'chromedriver' is in your PATH
    driver.get(form_url)
    time.sleep(2)  # Allow page load
    wait_for_sign_in()

    error_list = []  # Track any names that failed submission

    for _, row in data.iterrows():
        try:
            driver.get(form_url)
            time.sleep(2)  # Allow page load
            clear_submission(driver)
            time.sleep(2)

            res_name = row["Resident Name"]
            # Format date as MM/DD/YYYY
            date_str = (str(row["Date"])[5:10] + "/" + str(row["Date"])[0:4]).replace("-", "/")
            theme = row["Theme"]
            description = row["Description"]

            # Locate the search fields (resident, building, population, themes)
            search_fields = driver.find_elements(By.ID, "forms-tag-search-input-div")
            resident_name_field = search_fields[0]
            building_field = search_fields[1]
            student_population_field = search_fields[2]
            themes_field = search_fields[3]

            # Fill out the resident name and select the first result
            resident_name_field.send_keys(res_name)
            time.sleep(3)
            search_names = driver.find_elements(By.CLASS_NAME, "forms-subscriptions-search-result-row")
            if not search_names:
                continue
            search_names[0].click()
            time.sleep(1)

            # Fill out the building field and select the first result
            building_field.send_keys(BUILDING)
            time.sleep(2)
            search_buildings = driver.find_elements(By.CLASS_NAME, "forms-subscriptions-search-result-row")
            if not search_buildings:
                continue
            search_buildings[0].click()

            # Fill out the student population field and select the first result
            student_population_field.send_keys("Upperclass")
            time.sleep(2)
            search_populations = driver.find_elements(By.CLASS_NAME, "forms-subscriptions-search-result-row")
            if not search_populations:
                continue
            search_populations[0].click()
            time.sleep(1)

            # Fill out the date field
            date_field = driver.find_element(By.CLASS_NAME, "elm-datepicker--input")
            date_field.send_keys(date_str)
            time.sleep(1)

            # Fill out the description field
            description_field = driver.find_element(By.ID, "desc_resp_sub_YmoBEE")
            description_field.send_keys(description)
            time.sleep(1)

            # Fill out the themes field and select the first result
            themes_field.send_keys(theme)
            time.sleep(2)
            search_themes = driver.find_elements(By.CLASS_NAME, "forms-subscriptions-search-result-row")
            if not search_themes:
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

            # Detect the confirmation popup using an explicit wait
            try:
                WebDriverWait(driver, 3).until(
                    EC.presence_of_element_located(
                        (By.XPATH, "//*[contains(text(), 'Form submitted successfully!')]")
                    )
                )
                print(f"Submitted: {res_name}")
            except TimeoutException:
                print("Submission confirmation popup element was not detected.")
                print(f"NOT SUBMITTED: {res_name}")
                error_list.append(res_name)
                clear_submission(driver)
            time.sleep(1)  # Allow form processing

        except Exception as error:
            print(f"Error processing {res_name}: {error}")
            error_list.append(res_name)
            clear_submission(driver)

    # Write errors to file if any
    if error_list:
        with open("submission_failure.txt", "w") as file:
            for error in error_list:
                file.write(f"{error}\n")
        print("Errors have been written to submission_failure.txt")

    print("Script complete. Errors:", error_list)
    print("Error Count:", len(error_list))
    driver.quit()


def main():
    file_path = "updated_residents.xlsx"  # Path to your Excel file
    form_url = "https://roompact.com/forms/#/form/7odwkY"  # Replace with the actual form URL

    data = load_data(file_path)
    print("Loaded data. Beginning interaction submission...")
    automate_form(data, form_url)
    print("Submission complete!")


if __name__ == "__main__":
    main()
