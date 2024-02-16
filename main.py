import pandas as pd
from jinja2 import Environment, FileSystemLoader
import random
import os

# Load the Excel file
excel_data = pd.read_excel('responses2.xlsx')  # Replace with your actual Excel file name
rows = excel_data.to_dict(orient='records')

# Set up Jinja2 environment
env = Environment(loader=FileSystemLoader('.'))
template = env.get_template('template.html')  # Ensure this is in your current directory

# Check if the cards directory exists, if not, create it
if not os.path.exists('cards'):
    os.makedirs('cards')

# Generate HTML files
for row in rows:
    random_digit = random.randint(0, 9)
    file_name = f"{str(row['Name of the Recipient']).replace(' ', '')}Card{random_digit}.html"

    # Determine the background based on the user's choice
    background_choice = row['What background image pattern would you like for your e-card?']
    if background_choice == 'Red Hearts':
        background_color = 'white'
        background_image = 'https://www.icegif.com/wp-content/uploads/2023/10/icegif-498.gif'
    elif background_choice == 'Pink Hearts':
        background_color = 'white'
        background_image = 'https://imgur.com/0z2kvaf.jpg'
    elif background_choice == 'Solid Red':
        background_color = 'red'
        background_image = ''  # No image for solid color backgrounds
    elif background_choice == 'Solid Pink':
        background_color = 'pink'
        background_image = ''  # No image for solid color backgrounds
    else:
        background_color = 'white'  # Default background color
        background_image = ''  # Default background image if no option matches

    # Render the template with the row data
    html_content = template.render(
        recipient_name=row['How would you like the card to be addressed? Ex. "Dear Jane,"'],
        message=row['Write the body of your card... what message would you like to send?'],
        sender_name=row['How would you like to sign your card?  Ex. "From John", "From Anonymous", etc.'],
        background_color=background_color,
        background_image=background_image
    )

    # Write the HTML content to a file
    with open(os.path.join('cards', file_name), 'w', encoding='utf-8') as f:
        f.write(html_content)

print("HTML cards have been generated.")

# Directory containing the card files
cards_directory = './cards'

# Base URL to attach filenames to
base_url = 'https://valentinescardtemplate.vercel.app/'

# Output file to save the URLs
output_file = 'urls.txt'

# Open the output file in write mode
with open(output_file, 'w') as file:
    # Traverse the cards directory
    for filename in os.listdir(cards_directory):
        # Check if the file is an HTML file
        if filename.endswith('.html'):
            # Construct the full URL
            full_url = f'{base_url}{filename}'
            # Write the URL to the file, followed by a newline
            file.write(full_url + '\n')

print(f'URLs have been saved to {output_file}.')