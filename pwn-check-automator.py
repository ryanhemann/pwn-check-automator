import os
import requests
import pandas as pd
import time
from datetime import datetime

# Specify the directory path you want the script to run in 
desired_directory = '/path/to/directory'

# Change the working directory to the desired location
os.chdir(desired_directory)

# API Key (Replace [your key] with your actual API key)
api_key = 'your key'

# Read email addresses from Excel
df = pd.read_excel('input_file_name.xlsx')
email_addresses = df['Email'].tolist()

# Create a list to store breach data
breach_data_list = []

# Define rate limiting parameters
rate_limit_delay = 6  # Delay in seconds (adjust as needed)
max_requests_per_minute = 10  # Adjust based on the API rate limits

# Initialize a timestamp for rate limiting
last_request_time = time.time()

# Generate a timestamp for the output file name
timestamp = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')

# Construct the output file name with the timestamp
output_file_name = f'output_file_{timestamp}.xlsx'

# Loop through each email address
for email in email_addresses:
    # Construct the URL with the email address (ensure it's URL-encoded)
    url = f'https://haveibeenpwned.com/api/v3/breachedaccount/{email}'

    # Create headers with the API key
    headers = {
        'hibp-api-key': api_key
    }

    # Check the time since the last request
    time_since_last_request = time.time() - last_request_time
    if time_since_last_request < 60 / max_requests_per_minute:
        # If the rate limit is reached, wait to stay within the limit
        time_to_wait = (60 / max_requests_per_minute) - time_since_last_request
        print(f'Rate limited. Waiting for {time_to_wait:.2f} seconds...')
        time.sleep(time_to_wait)

    # Make a request to the API
    response = requests.get(url, headers=headers)

    # Update the last request time
    last_request_time = time.time()

    if response.status_code == 200:
        # If the account is found in a breach, add the breach data to the list
        breaches = response.json()
        for breach in breaches:
            breach_data_list.append({
                'Email': email,
                'BreachName': breach['Name']
            })
    elif response.status_code == 404:
        # Handle cases where the email was not found in any breaches
        breach_data_list.append({
            'Email': email,
            'BreachName': 'null'  # Set the value to 'null'
        })
    else:
        # Handle other status codes or errors
        print(f'Error for {email}: Status Code {response.status_code}')

# Create a DataFrame from the breach data list
breach_df = pd.DataFrame(breach_data_list)

# Write the results to the Excel file with the timestamp in the name
breach_df.to_excel(output_file_name, index=False)
