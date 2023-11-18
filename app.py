import requests
from bs4 import BeautifulSoup
from configparser import ConfigParser
import pandas as pd

# Read config.ini file
config_object = ConfigParser()
config_object.read("./config/config.ini")

# Get the password
company = config_object["company"]

# URL of the login form
login_url = company["login_url"]
target_url = company["target_url"]  # Replace with the desired URL

# Start a session to persist cookies, including the CSRF token
session = requests.Session()

# Send a GET request to the login page to obtain the CSRF token
response = session.get(login_url)
soup = BeautifulSoup(response.text, 'html.parser')

# Find the CSRF token in the form
csrf_token = soup.find('input', {'name': '_token'}).get('value')

# User and password data to be submitted
payload = {
    'email': company["email"],
    'password': company["password"],
    '_token': csrf_token  # Include the CSRF token in the payload
}

# Send a POST request with the data
response = session.post(login_url, data=payload)

# Check for "Page Expired" in the response content
if 'Page Expired' in response.text:
    # Fetch a new CSRF token and retry the login
    print("Page Expired. Fetching a new CSRF token and retrying...")

    # Send another GET request to obtain a new CSRF token
    response = session.get(login_url)
    soup = BeautifulSoup(response.text, 'html.parser')

    # Find the new CSRF token in the form
    new_csrf_token = soup.find('input', {'name': '_token'}).get('value')

    # Update the payload with the new CSRF token
    payload['_token'] = new_csrf_token

    # Retry the login with the new CSRF token
    response = session.post(login_url, data=payload)

# Check if the login was successful
if response.status_code == 200:
    print("Login successful")

    # List to store information for each student
    students_list = []
    for i in range(3, 27):

        # Now, you can access the target URL
        response = session.get(target_url + str(i))

        # Check the response for the target URL
        if response.status_code == 200:

            # Parse the HTML content of the target page
            soup_target = BeautifulSoup(response.text, 'html.parser')

            # Find the element with the class 'description__title title' and get its text content
            description_title = soup_target.find('h1', {'class': 'description__title title'}).text
            description_company = soup_target.find('a', {'class': 'description__company'}).text

            # List to store information for each student

            student_items = soup_target.find_all('div', {'class': 'student__item'})

            # Iterate over each 'student__item'
            for student_item in student_items:
                # Extract relevant information

                student_name = student_item.find('div', class_='student__title').text.strip()

                total_earned_coin_element = student_item.find('span', {'title': 'total earned coin'})
                total_earned_coin = total_earned_coin_element.text.strip() if total_earned_coin_element else None

                student_options = student_item.find_all('div', class_='student__option')

                student_counters = [option.find('div', class_='student__counter').text.strip() for option in
                                    student_options]

                year = student_counters[0]
                month = student_counters[1]
                day = student_counters[2]

                # Create a dictionary for the current student
                student_info = {
                    'unit': description_title,
                    'name': student_name,
                    'total_earned_coin': int(total_earned_coin.replace(',', '')),
                    'year': int(year),
                    'month': int(month),
                    'day': int(day)
                }

                # Append the student dictionary to the list
                students_list.append(student_info)

        else:
            print("Failed to access the target URL. Status code:", response.status_code)

    # # Convert the list of dictionaries to a DataFrame
    df = pd.DataFrame(students_list)

    # # Export the DataFrame to Excel
    df.to_excel('./output/teamvibe.xlsx', index=False, sheet_name='teamvibe')

else:
    print("Login failed. Status code:", response.status_code)
