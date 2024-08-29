import win32com.client  # Importing the win32com.client library to interact with Outlook
import time  # Importing time library for sleep functionality
import pandas as pd  # Importing pandas for data manipulation
import datetime  # Importing datetime for date operations

# Read data from the Excel file 'Template.xlsx'
df = pd.read_excel('Template.xlsx')

# Get the current date
current_date = datetime.datetime.now().date()

# Initialize an empty string for the acronym
ACRONYM = ""

# Loop through the DataFrame to find the member name based on the current date
for i in range(len(df)):
    if pd.to_datetime(df.loc[i, 'jour']).date() == current_date:
        member = df.loc[i, 'nom']  # Get the member name
        words = member.split()  # Split the name into words
        ACRONYM = ''.join(word[0] for word in words).upper()  # Create acronym from the initials
        print(f"Acronym: {ACRONYM}")  # Print the acronym
        break  # Exit the loop once the acronym is found

# Connect to Outlook application
outlook = win32com.client.Dispatch("Outlook.Application")
mapi = outlook.GetNamespace("MAPI")  # Get the MAPI namespace

# Define the HTML email template
template = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Email Template</title>
    <style>
        body {
            font-family: Arial, sans-serif;
        }
        table {
            width: auto; /* Set the table width to auto to fit content */
            border-collapse: collapse;
            border: 2px solid orange; /* Orange border for the table */
            margin: 0 auto; /* Center the table horizontally */
        }
        td {
            border: 2px solid orange; /* Orange border for table cells */
            padding: 15px;
            text-align: left;
        }
        th {
            border: 2px solid orange; /* Orange border for table cells */
            padding: 15px;
            text-align: center;
        }
        .logo {
            text-align: center;
        }
        .text-block {
            background-color: white;
        }
        /* CSS to control image size */
        .logo img, .text-block img {
            max-width: 100px; /* Set maximum width for logos */
            height: auto; /* Maintain aspect ratio */
        }
        /* Align the French flag image with the middle of the text */
        .text-block img.french-logo {
            margin-right: 10px; /* Adjust the space between image and text */
            vertical-align: middle; /* Align the image with the middle of the text */
        }
    </style>
</head>
<body>

<table>
    <tr class="logo">
        <th>
            <img src="D:\Stage\Infosquare_logo.png" alt="Infosquare Logo"> <!-- Replace with actual path to logo image -->
        </th>
    </tr>
    <tr class="text-block">
        <td>
            <img src="D:\Stage\OIPfr.jpg" alt="French Logo" class="french-logo">
            &nbsp;Bonjour Madame, Monsieur,<br><br>
            Nous accusons la bonne réception de votre ticket et reviendrons vers vous dans les meilleurs délais.<br><br>
            Bien cordialement,<br><br>
            <strong>L'équipe TMA Infosquare ({ACRONYM})</strong>
        </td> <!-- French Text Block -->
    </tr>
    <tr class="text-block">
        <td>
            <img src="D:\Stage\eng.png" alt="English Logo">
            &nbsp;Dear Madam, Sir,<br><br>
            We acknowledge the successful reception of your ticket and will get back to you as soon as possible.<br><br>
            Kind regards,<br><br>
            <strong>The Infosquare AMS team ({ACRONYM})</strong>
        </td> <!-- English Text Block -->
    </tr>
</table>

</body>
</html>
"""

# Define the category for replied emails and keywords to search for
replied_category = "Replied"  
keywords = [ "project", "meeting", "task"]  

# Infinite loop to continuously check for new emails
while True:
    inbox = mapi.GetDefaultFolder(6)  # Access the inbox folder
    items = inbox.Items  # Get the items in the inbox
    items.Sort("[ReceivedTime]", True)  # Sort items by received time in descending order
    latest_email = items.GetFirst()  # Get the latest email
    
    print("Checking for unreplied emails...")
    
    # Loop through the emails
    while latest_email:
        # Check if the email has not been replied to, is not a reply itself, and contains specified keywords
        if (replied_category not in latest_email.Categories and 
            "RE" not in latest_email.Subject.lower() and 
            any(keyword.lower() in latest_email.Subject.lower() for keyword in keywords)):
            
            print(f"Keyword found in email subject: '{latest_email.Subject}'.")
            
            # Replace the acronym placeholder in the template with the actual acronym
            dynamic_template = template.replace("{ACRONYM}", ACRONYM)
            reply = latest_email.Reply()  # Create a reply to the email              
            reply.HTMLBody = dynamic_template + "<br><br><br>" + reply.HTMLBody  # Insert the template into the reply
            
            reply.Send()  # Send the reply
            print(f"Reply sent to message: '{latest_email.Subject}'.")
            
            # Mark the email as replied and update its status
            latest_email.Categories = replied_category  
            latest_email.Save()  
            latest_email.UnRead = False  # Mark the email as read
        
        # Get the next email in the inbox
        latest_email = items.GetNext()  
    
    time.sleep(15)  # Wait for 15 seconds before checking for new emails again