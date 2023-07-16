# Import the necessary modules
import pandas as pd
import pytesseract
import re
import time
import os
from csv import writer, reader
import tempfile
from pdf2image import convert_from_path
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.remote.webelement import WebElement

# Define variables
file_type = 'intake'  # 'intake' or 'vf'
batch_number = 1
path_to_batch = r"C:\Users\shtey\Downloads\Batch-{bn}.pdf".format(bn=batch_number)
default_null_date = '10/10/1903'
default_null_phone_number = '(102) 301-2309'

# Define pandas options
pd.set_option('display.max_columns', None)
pd.set_option('display.max_rows', None)

# Get the pages of the pdf
pages = convert_from_path(path_to_batch)
num_pages = len(pages)
print(f"Number of pages in the batch: {num_pages}")

# Define empty lists to hold the information
pages_text = [None for i in range(num_pages)]
document_dates = [default_null_date for i in range(num_pages)]


def extract_text(page) -> str:
    """
    Extract the text from a page of a pdf
    """
    return pytesseract.image_to_string(page, lang='eng', config='--psm 6')


def extract_document_date(page) -> str:
    """
    Extract the document date from the bottom of a page of a pdf
    """
    text = extract_text(page)
    regex = r"(\d{2}/\d{2}/\d{4})"
    matches = re.findall(regex, text)
    if matches:
        return matches[-1]
    else:
        return default_null_date


def extract_information_from_text(text: str) -> list:
    """
    Extract the relevant information from the text of a page
    """
    # Intake forms and VF forms have different information to extract
    data_regexes = {}
    if file_type == 'intake':
        data_regexes = {
            'First name': r'First:\s([A-Za-z]+)',
            'Last name': r'Last:\s([A-Za-z]+)',
            'DOB': r'DOB:\s(\d{2}/\d{2}/\d{4})',
            'Sex': r'Sex:\s([A-Za-z]+)',
            'Preferred Phone': r'Preferred:\sCell:\s(\(\d{3}\)\s\d{3}-\d{4})',
            'Address': r'Address:\s(.+?)\n',
            'Provider': r'Provider:\s(.+?)\n',
        }
    elif file_type == 'vf':
        data_regexes = {
            'Full Name': r"NAME:\s+(\w+\s*,\s*\w+)\s+",
            'DOB': r"DOB:\s*(\d{2}-\d{2}-\d{4})",
            'Screening Date': r"Screening DATE:\s*(\d{2}-\d{2}-\d{4})"
        }

    info = []

    for key, regex in data_regexes.items():
        matches = re.findall(regex, text)
        if matches:
            if key == 'Sex':
                info.append(
                    str(matches[-1]).replace('i', 'l').replace('t', 'l'))  # Sometimes the l is misread as an i or t
            elif key == 'DOB' or key == 'Screening Date':
                year = str(matches[-1])[-4:]
                if int(year) < 1900:
                    info.append(default_null_date)
                else:
                    info.append(str(matches[-1]).replace('-',
                                                         '/'))  # The date is formatted with dashes instead of slashes in the visual fields
            else:
                info.append(str(matches[-1]))
        else:
            if key == 'DOB' or key == 'Screening Date':
                info.append(default_null_date)  # Set the date to default null date if it is not found
            elif key == 'Preferred Phone':
                info.append(
                    default_null_phone_number)  # Set the phone number to default null phone number if it is not found
            else:
                info.append('***')  # Set the other fields to *** if they are not found

    if file_type == 'intake':
        # Set Document Date to default null date if it is the same as the DOB
        if info[-1] == info[1]:
            info[-1] = default_null_date
    elif file_type == 'vf':
        if info[0] != '***':
            # Split the name into first and last name
            first_name = info[0].split(',')[1].strip()
            last_name = info[0].split(',')[0].strip()
            info[0] = first_name
            info.insert(1, last_name)
        else:
            info.insert(1, '***')

    return info


if file_type == 'intake':
    pages_bottom = [None for i in range(len(pages))]  # The bottom of the page holds the document date
    pages_cropped = [None for i in range(len(pages))]  # The top of the page holds the patient information

    for i in range(num_pages):
        h = pages[i].height
        w = pages[i].width
        # Crop function works as follows: crop((left, top, right, bottom))
        pages_bottom[i] = pages[i].crop((0, h - (h / 10), w, h))  # Bottom 10% of the page
        pages_cropped[i] = pages[i].crop((0, 0, w, h / 3))  # Top 33% of the page

    pages_text = [extract_text(page) for page in pages_cropped]  # Extract the text from the top of the page
    document_dates = [extract_document_date(page) for page in
                      pages_bottom]  # Extract the document date from the bottom of the page

elif file_type == 'vf':
    pages_cropped = [None for i in range(len(pages))]  # The information is all in the top 1/7 of the page
    for i in range(num_pages):
        h = pages[i].height
        w = pages[i].width
        pages_cropped[i] = pages[i].crop((0, 0, w, h / 7))

    pages_text = [extract_text(page) for page in pages_cropped]

# Create a dataframe that will hold the info for each patient
df = pd.DataFrame(
    columns=['First Name', 'Last Name', 'Date of Birth', 'Sex', 'Preferred Phone', 'Address', 'Provider',
             'Document Date', 'Screening Date'])

for i in range(num_pages):
    # Extract the information from the text of the page
    data = extract_information_from_text(pages_text[i])
    if file_type == 'intake':
        # Initialize the screening date and document date to the default null date
        data.append(default_null_date)
        data.append(default_null_date)
    elif file_type == 'vf':
        # Set all the intake form variables to ***
        for j in range(3, 8):
            data.insert(j, '***')
    # Add the formatted data list to the dataframe
    df.loc[i] = data

# Set the document dates to the dates extracted from the bottom of the pages
df['Document Date'] = document_dates

# Save the dataframe to a csv file for debugging purposes
df.to_csv('data.csv', index=False)

# Set pd options
pd.set_option('display.max_columns', None)
pd.set_option('display.max_rows', None)

# Using Selenium to Automate the Process of Inputting the Data into the Database
# Open Chrome Browser
driver = webdriver.Chrome()

# Open the URL
url = "https://revolutionehr.com/static/#/"
driver.get(url)

# Define wait condition
wait = WebDriverWait(driver, 10)


def ensure_click(location: str) -> WebElement:
    """
    Ensures that the element at the given location is clickable and clicks it.
    Sometimes, the element clicks are intercepted by other elements, so this function keeps trying until the element is clickable.
    :param location:
    :return:
    """
    while True:  # Keep trying until the element is clickable
        try:
            field = fetch_element(location, EC.element_to_be_clickable)
            field.click()
            break
        except:
            continue  # Try again

    return field


def fetch_element(location: str, condition=EC.presence_of_element_located, locator=By.XPATH) -> WebElement or None:
    """
    Fetches an element from the page using the given xpath and condition.
    :param location: describes the path to the element
    :param condition: the condition that the element must meet
    :param locator: the type of locator to use (e.g. By.XPATH, By.ID, By.CSS_SELECTOR)
    :return: The element if it is found, None otherwise
    """
    try:
        return wait.until(condition((locator, location)))
    except Exception as e:
        max_retries = 3
        retries = 0
        while retries < max_retries:
            try:
                return wait.until(condition((locator, location)))
            except Exception as e:
                retries += 1
                time.sleep(1)
        return None


# Login with username and password.
username_field = fetch_element(
    '/html/body/div[2]/div/div/div[1]/div/rev-login-page/div/div[2]/div/rev-login-form/div[2]/div/form/div[1]/div/input',
    EC.element_to_be_clickable)
username_field.click()
username_field.send_keys('kshteyn')
password_field = fetch_element(
    '/html/body/div[2]/div/div/div[1]/div/rev-login-page/div/div[2]/div/rev-login-form/div[2]/div/form/div[2]/div/input',
    EC.element_to_be_clickable)
password_field.click()
password_field.send_keys('katushaa')
login_button = fetch_element(
    '/html/body/div[2]/div/div/div[1]/div/rev-login-page/div/div[2]/div/rev-login-form/div[2]/div/form/button',
    EC.element_to_be_clickable)
login_button.click()


def search_patient(field_path: str, field_value: str, patient_number: int) -> bool:
    """
    Search for a patient using the given fields and values.
    :param field_path: the path to the field
    :param field_value: the field's name
    :param patient_number: the patient number
    :return: True if the search was successful, False otherwise
    """
    # Fetch the field, click on it, and enter the value
    field = ensure_click(field_path)
    field.send_keys(df[field_value][patient_number])
    field.send_keys(u'\ue007')  # Press enter

    # Check if the search is finished
    try:  # Sometimes stale element reference exception is thrown
        search_finished_indicator = fetch_element(
            '/html/body/div[2]/div/div/pms-root/pms-patients/div/div/div/div/pms-search-patients/div[2]/div/h4',
            EC.presence_of_element_located)
        while 'Search Results' not in search_finished_indicator.text:
            search_finished_indicator = fetch_element(
                '/html/body/div[2]/div/div/pms-root/pms-patients/div/div/div/div/pms-search-patients/div[2]/div/h4',
                EC.presence_of_element_located)
    except:
        return False

    # Check if there are any results
    results_label = fetch_element(
        '/html/body/div[2]/div/div/pms-root/pms-patients/div/div/div/div/pms-search-patients/div[2]/div/ejs-grid/div[5]/div[4]/span[2]',
        EC.presence_of_element_located)

    # Extract the number of results. If there are no results, then try again with the next search method, if any.
    number_of_results = int(results_label.text[1])

    local_success = number_of_results > 0

    if local_success:
        local_success = find_patient(patient_data, number_of_results)

        if local_success:
            if file_type == 'intake':
                upload_form(patient_data[7])
            elif file_type == 'vf':
                upload_form(patient_data[8])

    return local_success


def find_patient(patient_info: list, n: int) -> bool:
    """
    Find the patient using the given patient info
    :param patient_info: all the patient info
    :param n: number of results
    :return: True if patient found, False otherwise
    """
    # Convert the results table to a dataframe
    table = fetch_element(
        '/html/body/div[2]/div/div/pms-root/pms-patients/div/div/div/div/pms-search-patients/div[2]/div/ejs-grid/div[3]/div/table',
        EC.presence_of_element_located)
    table_html = table.get_attribute('outerHTML')
    table = pd.read_html(table_html)[0]

    # Go through each patient in table to find valid patient
    valid_row_index = None
    for i, patient in table.iterrows():
        patient_info = [str(x).lower() for x in
                        patient_info]  # convert all patient info to lowercase for easier comparison

        date_of_birth_valid = patient_info[2] in table[3][i].lower()  # shared between intake forms and vf forms

        if file_type == 'intake':
            full_name_valid = patient_info[0] in table[2][i].lower() and patient_info[1] in table[2][0].lower()
            sex_valid = patient_info[3] in table[4][i].lower()
            phone_number_valid = patient_info[4] in str(table[5][i]).lower()
            address_valid = patient_info[5] in table[6][i].lower()
            provider_valid = patient_info[6] in table[7][i].lower()
            if (full_name_valid and date_of_birth_valid) or (
                    (full_name_valid or date_of_birth_valid) and (sex_valid or provider_valid) and (
                    phone_number_valid or address_valid)):
                valid_row_index = i  # if any of the above conditions are true, then the patient is valid
        elif file_type == 'vf':
            first_name_valid = patient_info[0] in table[2][i].lower()
            last_name_valid = patient_info[1] in table[2][i].lower()
            if (first_name_valid or last_name_valid) and date_of_birth_valid:
                valid_row_index = i  # if any of the above conditions are true, then the patient is valid

    # If a valid patient was found, click on the patient
    if valid_row_index is not None:
        # Click on the patient
        patient = fetch_element(
            f'/html/body/div[2]/div/div/pms-root/pms-patients/div/div/div/div/pms-search-patients/div[2]/div/ejs-grid/div[3]/div/table/tbody/tr[{valid_row_index + 1}]',
            EC.element_to_be_clickable)
        patient.click()
        return True
    else:  # No valid patient was found
        return False


def reset_search() -> None:
    # Clears the search fields
    ensure_click(
        '/html/body/div[2]/div/div/pms-root/pms-patients/div/div/div/div/pms-search-patients/div[1]/pms-patients-advanced-search/form/div[2]/div/button[2]')


def get_table() -> pd.DataFrame:
    """
    Get the table from the given table path
    :return: the table as a dataframe
    """
    table = fetch_element(
        '/html/body/div[2]/div/div/pms-root/pms-patients/div/div/div/div/pms-patient/div[2]/div/div/pms-patient-files/div/div[2]/div/div[2]/pms-folder-file-list/div/ejs-grid/div[4]/div/table',
        EC.presence_of_element_located)
    table_html = table.get_attribute('outerHTML')
    table = pd.read_html(table_html)[0]
    while table[1][0] == 'Documents':
        table = fetch_element(
            '/html/body/div[2]/div/div/pms-root/pms-patients/div/div/div/div/pms-patient/div[2]/div/div/pms-patient-files/div/div[2]/div/div[2]/pms-folder-file-list/div/ejs-grid/div[4]/div/table',
            EC.presence_of_element_located)
        table_html = table.get_attribute('outerHTML')
        table = pd.read_html(table_html)[0]
    return table


def upload_form(date: str) -> None:
    time.sleep(.5)  # Give the alert time to pop up
    # Check for alert pop-up and close it if it is present
    try:
        alert_cancel_btn = driver.find_element(By.CSS_SELECTOR,
                                               'button[data-test-id="alertHistoryModalCloseButton"]')
        alert_cancel_btn.click()
    except:
        pass

    # Navigate to the documents tab
    ensure_click(
        '/html/body/div[2]/div/div/pms-root/pms-patients/div/div/div/div/pms-patient/div[2]/div/pms-patient-navigation-bar/ejs-sidebar/div/a[20]/span')

    # Navigate to the correct folder
    if file_type == 'intake':
        ensure_click("//li[@title='IntakeForms']")
    elif file_type == 'vf':
        ensure_click("//li[@title='Visual Fields']")

    # Get the table that contains the patient's documents
    table = get_table()

    if date == default_null_date:
        filename = f'Unknown-Document-Date-{file_type}.pdf'  # Name the file as an unknown document date if the date is null
    else:
        # Get the month and year from the date and include it in the filename
        months = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October',
                  'November', 'December']
        try:
            month = months[int(date.split('/')[0]) - 1]
        except:
            month = 'UnknownMonth'
        year = date.split('/')[2]
        filename = f'{month}-{year}-{file_type}.pdf'

    # Check if the file has already been uploaded by comparing it to the files in the table
    files = table[1].tolist()
    if filename not in files:
        # Create a temporary path for the form and save it there
        form_path = os.path.join(tempfile.gettempdir(), filename)
        pages[i].save(form_path)

        # Click on the upload button
        ensure_click(
            '/html/body/div[2]/div/div/pms-root/pms-patients/div/div/div/div/pms-patient/div[2]/div/div/pms-patient-files/div/div[2]/div/div[2]/pms-folder-file-list/div/ejs-grid/div[2]/rev-table-action-menu/div/div/div[1]/div/button[1]')

        # Click on the upload file button and input the path to the form
        file_input = fetch_element(
            "input[type='file']",
            locator=By.CSS_SELECTOR)
        file_input.send_keys(form_path)

        # Wait until the file is uploaded into the patient's documents
        while True:
            table = get_table()
            files = table[1].tolist()
            if filename in files:
                break

        os.remove(form_path)  # Delete the temporary file to avoid cluttering the computer's storage and path errors


# Create a list to store patient's that could not be found/uploaded
error_patients = []

# For each patient in the dataframe
for index, row in df.iterrows():
    time.sleep(.5)  # Give the search page time to load
    # Retrieve the patient's data from the dataframe
    patient_data = df.iloc[index].tolist()

    # Click Patient Tab
    ensure_click('/html/body/div[1]/header/div/div[2]/ul/li[1]/a')

    # Click Advanced Search Button if we are searching for the first time
    if index == 0:
        ensure_click(
            '/html/body/div[2]/div/div/pms-root/pms-patients/div/div/div/div/pms-search-patients/div[1]/pms-patients-simple-search/form/div[2]/div/div[2]/button[3]')

    # Reset the search to avoid errors
    reset_search()

    # Search for the patient using last name
    success = search_patient(
        '/html/body/div[2]/div/div/pms-root/pms-patients/div/div/div/div/pms-search-patients/div[1]/pms-patients-advanced-search/form/div[1]/div[2]/div[1]/div/input',
        'Last Name', int(index))

    if not success:  # If no results are found using last name search
        reset_search()
        # Search for the patient using first name
        success = search_patient(
            '/html/body/div[2]/div/div/pms-root/pms-patients/div/div/div/div/pms-search-patients/div[1]/pms-patients-advanced-search/form/div[1]/div[2]/div[2]/div/input',
            'First Name', int(index))

    if not success:  # If no results are found using first name search
        reset_search()
        # Search for the patient using date of birth
        success = search_patient(
            '/html/body/div[2]/div/div/pms-root/pms-patients/div/div/div/div/pms-search-patients/div[1]/pms-patients-advanced-search/form/div[1]/div[2]/div[3]/div/ejs-datepicker/span/input',
            'Date of Birth', int(index))

    if not success and file_type == 'intake':  # If no results are found using date of birth search and we are working with an intake form
        reset_search()

        # Search for the patient using phone number
        success = search_patient(
            '/html/body/div[2]/div/div/pms-root/pms-patients/div/div/div/div/pms-search-patients/div[1]/pms-patients-advanced-search/form/div[1]/div[3]/div[2]/div/ejs-maskedtextbox/span/input',
            'Preferred Phone', int(index))

    if not success and file_type == 'intake':  # If no results are found using phone number search
        reset_search()

        # Final attempt: Search for the patient using address
        success = search_patient(
            '/html/body/div[2]/div/div/pms-root/pms-patients/div/div/div/div/pms-search-patients/div[1]/pms-patients-advanced-search/form/div[1]/div[3]/div[1]/div/input',
            'Address', int(index))

    if not success:
        # Add the patients location in their batch so that the user can easily locate their file. Also their information
        error_patients.append([str(batch_number), str(index + 1), patient_data[0], patient_data[1], patient_data[2]])

    # Display the patient's information and whether the upload was successful
    print("Patient " + str(index + 1) + ": " + str(patient_data[0]) + " " + str(patient_data[1]) + " " + str(patient_data[2]))
    print("Patient Upload Successful: " + str(success) + "\n")

# Get the corresponding csv path
csv_path = f'error_{file_type}.csv'

# Make a list of the patients that have already been uploaded as errors
with open(csv_path) as file:
    existing_patients = []
    reader_object = reader(file)
    for row in reader_object:
        existing_patients.append(row)

# Add the error patients into a dataframe containing the patients that could not be found/uploaded
with open(csv_path, 'a') as error_df:
    writer_object = writer(error_df)
    for error_patient in error_patients:
        if error_patient not in existing_patients:
            writer_object.writerow(error_patient)
    error_df.close()
