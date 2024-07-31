import pandas as pd
import pywhatkit as kit
import pickle
import os
from datetime import datetime

FORCE_RESEND = False

# Load the Excel file
file_path = 'invitation_list.xlsx'

# Function to read template from file
def read_template(file_path):
    with open(file_path, 'r') as file:
        return file.read()

# URLs for different events
event_urls = {
    'dinner': 'https://amarwedrachmalia.viding.co/?event=The+Reception+Soiree',
    'lunch': 'https://amarwedrachmalia.viding.co/?event=Wedding+Lunch+Reception',
    'ceremony': 'https://amarwedrachmalia.viding.co/?event=Wedding+Ceremony+%26+Lunch+Reception',
    'general': 'https://amarwedrachmalia.viding.co/'
}

if __name__ == "__main__":
    data = pd.read_excel(file_path)

    # Load sent and error data
    sent_file = 'sent_messages.pkl'
    if os.path.exists(sent_file):
        with open(sent_file, 'rb') as f:
            sent_data = pickle.load(f)
    else:
        sent_data = {'sent': [], 'errors': []}

    # Send messages
    for index, row in data.iterrows():
        try:
            name = row['name'].title()
        except:
            continue
        number = row['wa']
        if number:
            number = str(number).split('.')[0]
        else:
            continue
        event = row['event_link'].lower()
        if event not in event_urls:
            sent_data['errors'].append(number)
        try:
            language = row['language'].lower()  # Assuming there is a column for language preference
        except:
            language = 'english'
        
        # Normalize phone number
        if not number.startswith('+'):
            number = '+' + number
        
        # Prepare the message
        try:
            message = read_template(f'templates/{language}.txt').format(name=name, url=event_urls[event])
        except:
            message = read_template('templates/default.txt').format(name=name, url=event_urls[event])
        
        # Check if the message was already sent or had errors
        if number in sent_data['sent']:
            print(f"Message to {number} already sent.")
            if not FORCE_RESEND:
                continue
        elif number in sent_data['errors']:
            print(f"Retrying message to {number}.")
        
        try:
            # Send the message
            kit.sendwhatmsg_instantly(number, message, wait_time=20, tab_close=True)
            sent_data['sent'].append(number)
            print(f"Message sent to {number}.")
        except Exception as e:
            sent_data['errors'].append(number)
            print(f"Error sending message to {number}: {e}")

    # Save sent and error data
    with open(sent_file, 'wb') as f:
        pickle.dump(sent_data, f)