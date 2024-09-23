from selenium import webdriver
from selenium.webdriver.common.by import By
import time
from selenium.webdriver.common.keys import Keys
from datetime import datetime, timedelta
import requests
from selenium.webdriver.common.action_chains import ActionChains

# Initialize attempt count for retry mechanism
attempt = 0
while attempt <= 3:
    try:
        driver = webdriver.Chrome()
        driver.maximize_window()

        # ------------------------------EMAIL DATA EXTRACTION-----------------------------------

        driver.get("URL_PLACEHOLDER")  # Replace with the actual URL
        time.sleep(18)
        iframe = driver.find_element(By.CSS_SELECTOR, "iframe_selector_placeholder")
        driver.switch_to.frame(iframe)

        actions = ActionChains(driver)
        actions.move_by_offset("500", "500").click().perform()

        # Scroll to load email data
        for _ in range(7):
            actions.send_keys(Keys.PAGE_DOWN).perform()
            time.sleep(1)

        table_tag = []
        all_elements_email = driver.find_elements(By.XPATH, "//*")

        # Extract data from the table where RESOURCE_NAME_PLACEHOLDER is found
        for element in all_elements_email:
            if element.tag_name == "table" and "RESOURCE_NAME_PLACEHOLDER" in element.text:
                rows = element.find_elements(By.TAG_NAME, "tr")
                for row in rows:
                    row_tag = []
                    columns = row.find_elements(By.TAG_NAME, "td")
                    for column in columns:
                        row_tag.append(column.text)
                    table_tag.append(row_tag)
                break

        # Creating a dictionary for storing and reversing data (e.g., emails to names)
        data_dict = {}
        reverse_data_dict = {}
        for sublist in table_tag[1:]:
            third_element = sublist[2]
            fifth_element = sublist[4]
            data_dict[third_element] = fifth_element
        for key, value in data_dict.items():
            reverse_data_dict[value] = key

        # ------------------------------INCIDENTS DATA EXTRACTION-----------------------------------

        driver.get("URL_PLACEHOLDER")  # Replace with the actual URL
        time.sleep(15)
        iframe = driver.find_element(By.CSS_SELECTOR, "iframe_selector_placeholder")
        driver.switch_to.frame(iframe)

        actions = ActionChains(driver)
        actions.move_by_offset("500", "500").click().perform()

        # Scroll through the incidents page
        for _ in range(20):
            actions.send_keys(Keys.PAGE_DOWN).perform()
            time.sleep(1)

        table_data = []
        all_elements = driver.find_elements(By.XPATH, "//*")

        # Extract table data containing incident information
        for element in all_elements:
            if element.tag_name == "table" and "INCIDENT_IDENTIFIER_PLACEHOLDER" in element.text:
                rows = element.find_elements(By.TAG_NAME, "tr")
                for row in rows:
                    row_data = []
                    columns = row.find_elements(By.TAG_NAME, "td")
                    for column in columns:
                        row_data.append(column.text)
                    table_data.append(row_data)
                break

        driver.quit()

        flow_url = "URL_PLACEHOLDER"

        # Process and send data to API if the incident meets the criteria
        for row in table_data[1:]:
            date_col = row[4]
            if date_col == '':
                date_col = '01/01/2010'
            date = datetime.strptime(date_col, "%d/%m/%Y")
            date_difference = date - datetime.now() + timedelta(days=1)

            if date_difference.days <= 5:
                tag = row[5]
                inc_no = row[2]
                data = {"Inc no.": row[2],
                        "Difference": date_difference.days,
                        "Customer": row[6],
                        "Inc subject": row[7],
                        "Person responsible": data_dict.get(row[5]),
                        "Test name": data_dict[tag],
                        }
                response = requests.post(flow_url, json=data)
                time.sleep(1)

        break

    except Exception as e:
        attempt += 1
        error_message = str(e)
        with open("error_log_reminder.txt", "w") as f:
            f.write(error_message)
        print("Failed to connect, retrying....")
        time.sleep(20)
