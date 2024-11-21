import requests
from bs4 import BeautifulSoup
import pandas as pd
import logging

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler("details_extraction.log", mode="w", encoding="utf-8"),
    ],
)

# Function to extract details from the response
def extract_details(html_content):
    soup = BeautifulSoup(html_content, 'html.parser')
    
    # Extract company name
    company_name = soup.select_one(".ce_head").text.strip() if soup.select_one(".ce_head") else "N/A"
    
    # Extract address
    address = soup.select_one(".ce_addr").text.strip() if soup.select_one(".ce_addr") else "N/A"
    
    # Extract phone
    phone = soup.select_one(".ce_phone .ce_cTxt a").text.strip() if soup.select_one(".ce_phone .ce_cTxt a") else "N/A"
    
    # Extract email
    email = soup.select_one(".ce_email .ce_cTxt a").text.strip() if soup.select_one(".ce_email .ce_cTxt a") else "N/A"
    
    # Extract website
    website = soup.select_one(".ce_website a").get("href").strip() if soup.select_one(".ce_website a") else "N/A"

    facebook = soup.select_one(".ce_smch.ce_Facebook a").get("href").strip() if soup.select_one(".ce_smch.ce_Facebook a") else "N/A"
    linkedin = soup.select_one(".ce_smch.ce_LinkedIn a").get("href").strip() if soup.select_one(".ce_smch.ce_LinkedIn a") else "N/A"
    instagram = soup.select_one(".ce_smch.ce_Instagram a").get("href").strip() if soup.select_one(".ce_smch.ce_Instagram a") else "N/A"

    
    
    # Combine the data into a dictionary
    data = {
        "Company Name": company_name,
        "Address": address,
        "Phone": phone,
        "Email": email,
        "Website": website,
        "Facebook":facebook,
        "LinkedIn":linkedin,
        "Instagram":instagram
    }
    
    return data

# Load the Excel file
input_file = "exhibitors.xlsx"
logging.info(f"Loading file: {input_file}")
df = pd.read_excel(input_file)

# Iterate through the links in column B and extract details
extracted_data = []
for index, row in df.iterrows():
    url = row["Link"]  # Assuming column B is named "Link"
    if pd.isna(url):  # Skip rows with missing URLs
        logging.warning(f"Skipping row {index}: No URL found")
        extracted_data.append({"Company Name": "N/A", "Address": "N/A", "Phone": "N/A", "Email": "N/A", "Website": "N/A","Facebook":"N/A","LinkedIn":"N/A","Instagram":"N/A"})
        continue
    
    logging.info(f"Fetching details for: {url}")
    try:
        response = requests.get(url, timeout=10)  # Set a timeout for requests
        if response.status_code == 200:
            details = extract_details(response.text)
            extracted_data.append(details)
        else:
            logging.error(f"Failed to fetch URL: {url}, Status Code: {response.status_code}")
            extracted_data.append({"Company Name": "N/A", "Address": "N/A", "Phone": "N/A", "Email": "N/A", "Website": "N/A","Facebook":"N/A","LinkedIn":"N/A","Instagram":"N/A"})
    except requests.exceptions.RequestException as e:
        logging.error(f"Error fetching URL: {url}, Error: {e}")
        extracted_data.append({"Company Name": "N/A", "Address": "N/A", "Phone": "N/A", "Email": "N/A", "Website": "N/A","Facebook":"N/A","LinkedIn":"N/A","Instagram":"N/A"})

# Add extracted data to the DataFrame
details_df = pd.DataFrame(extracted_data)
output_df = pd.concat([df, details_df], axis=1)

# Save the updated DataFrame back to Excel
output_file = "exhibitor_updated_remaining.xlsx"
logging.info(f"Saving updated data to: {output_file}")
output_df.to_excel(output_file, index=False)

logging.info("Details extraction process completed.")
