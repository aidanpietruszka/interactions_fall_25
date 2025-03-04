import math
import time
import threading
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

BUILDING = "North Avenue East"  # Replace if needed

def wait_for_sign_in():
    """
    Wait for the user to sign in manually.
    """
    while True:
        response = input("Sign in in the opened browser, then press y and enter:\n").strip().lower()
        if response == "y":
            print("Sign in complete.")
            break
        elif response == "n":
            print("Please sign in and then type 'y' to continue.")
        else:
            print("Invalid input. Please type 'y' or 'n'.")

def clear_field(field):
    """
    Clear a given input field.
    First, attempt using field.clear(); if that fails, send backspace keystrokes.
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
    The Excel file must contain 'Resident Name', 'Date', 'Theme', and 'Description' columns.
    """
    return pd.read_excel(file_path)

def fill_form_for_chunk(thread_id, data_chunk, form_url, cookies, error_list, error_list_lock):
    """
    Worker function for each thread.
    Creates its own browser (driver), loads the saved cookies to maintain the session,
    and processes its data chunk.
    """
    driver = webdriver.Chrome()  # Ensure chromedriver is in your PATH
    driver.get(form_url)
    time.sleep(2)
    # Load cookies into the new driver instance
    driver.delete_all_cookies()
    for cookie in cookies:
        try:
            driver.add_cookie(cookie)
        except Exception as e:
            print(f"[Thread {thread_id}] Failed to add cookie: {cookie} Error: {e}")
    driver.get(form_url)  # Refresh so cookies are applied
    # Extra refresh to beat any load issues
    time.sleep(2)
    driver.get(form_url)
    time.sleep(2)
    
    for _, row in data_chunk.iterrows():
        try:
            driver.get(form_url)
            time.sleep(2)
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

            # Fill out resident name and select the first result
            resident_name_field.send_keys(res_name)
            time.sleep(3)
            search_names = driver.find_elements(By.CLASS_NAME, "forms-subscriptions-search-result-row")
            if not search_names:
                continue
            search_names[0].click()
            time.sleep(1)

            # Fill out building field and select the first result
            building_field.send_keys(BUILDING)
            time.sleep(2)
            search_buildings = driver.find_elements(By.CLASS_NAME, "forms-subscriptions-search-result-row")
            if not search_buildings:
                continue
            search_buildings[0].click()

            # Fill out student population field and select the first result
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

            # Handle opt-out scenario based on description
            if "Resident opted out" in description:
                optout_button = driver.find_element(By.XPATH, "//input[contains(@aria-label, 'Yes')]")
            else:
                optout_button = driver.find_element(By.XPATH, "//input[@aria-label='No']")
            optout_button.click()
            time.sleep(2)

            submit_button = driver.find_element(By.XPATH, "//button[contains(@class, 'forms-submit-btn')]")
            submit_button.click()

            # Check for submission confirmation popup
            try:
                WebDriverWait(driver, 3).until(
                    EC.presence_of_element_located(
                        (By.XPATH, "//*[contains(text(), 'Form submitted successfully!')]")
                    )
                )
                print(f"[Thread {thread_id}] Submitted: {res_name}")
            except TimeoutException:
                print(f"[Thread {thread_id}] Submission confirmation not detected for: {res_name}")
                with error_list_lock:
                    error_list.append(res_name)
                clear_submission(driver)
            time.sleep(1)

        except Exception as e:
            print(f"[Thread {thread_id}] Error processing {res_name}: {e}")
            with error_list_lock:
                error_list.append(res_name)
            clear_submission(driver)
    
    driver.quit()
    print(f"[Thread {thread_id}] Finished processing its chunk.")

def run_submission_pass(data_df, form_url, cookies, n):
    """
    Splits the given dataframe into n chunks and processes them concurrently.
    Returns a list of resident names that failed submission.
    """
    n = n if n else 1
    
    chunk_size = math.ceil(len(data_df) / n)
    threads = []
    error_list = []
    error_list_lock = threading.Lock()
    
    for i in range(n):
        start_idx = i * chunk_size
        end_idx = (i + 1) * chunk_size
        data_chunk = data_df.iloc[start_idx:end_idx]
        if data_chunk.empty:
            continue
        t = threading.Thread(
            target=fill_form_for_chunk,
            args=(i, data_chunk, form_url, cookies, error_list, error_list_lock)
        )
        t.start()
        threads.append(t)
        
    for t in threads:
        t.join()
    
    return error_list

def main():
    file_path = "updated_residents.xlsx"  # Path to your Excel file
    form_url = "https://roompact.com/forms/#/form/7odwkY"  # Replace with the actual form URL

    data = load_data(file_path)
    print("Loaded data.")

    # Step 1: Open a master driver and sign in manually, then extract cookies.
    master_driver = webdriver.Chrome()
    master_driver.get(form_url)
    time.sleep(2)
    wait_for_sign_in()
    cookies = master_driver.get_cookies()
    print("Extracted cookies from signed-in session.")
    master_driver.quit()

    # Set number of parallel windows (driver instances)
    n = 24

    # --- Pass 1: Process entire dataset ---
    print("Running Pass 1...")
    error_list = run_submission_pass(data, form_url, cookies, n)
    print("Pass 1 errors:", error_list)

    error_size = 2 ** 31 - 1  # Initialize to a large number
    pass_num = 2
    while error_size != len(error_list):
        error_size = len(error_list)
        retry_data = data[data["Resident Name"].isin(error_list)]
        print("Running Pass 2 for these errors:", error_list)
        threads = min(n, len(retry_data))
        error_list = run_submission_pass(retry_data, form_url, cookies, threads)
        print(f'Pass {pass_num} errors: ', error_list)
        pass_num += 1

    # Write remaining errors to file (those that failed all three passes)
    if error_list:
        with open("submission_failure.txt", "w") as f:
            for err in error_list:
                f.write(f"{err}\n")
        print("Final errors have been written to submission_failure.txt")
        print("Script complete. Final error list:", error_list)
        print("Check the excel sheet for any errors with these names. Does this student show up in roompact?")
    else:
        print(f'All submissions succeeded after {pass_num} passes.')
        print('100% ACCURACY!!!')



if __name__ == "__main__":
    main()