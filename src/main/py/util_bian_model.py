import requests
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.worksheet.table import Table, TableStyleInfo
import logging

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Define the API endpoints and headers
base_url = "https://api-v3.bian.org"
service_domains_url = f"{base_url}/ServiceDomainsBasic"
service_domain_details_url = f"{base_url}/ServiceDomainsByBianId"
headers = {
    "Accept": "application/json",
    "Authorization": "Bearer YOUR_ACCESS_TOKEN"  # Replace with your actual access token
}

# Log the start of the process
logging.info("Starting to fetch service domains")

# Send a GET request to the /ServiceDomainsBasic API
response = requests.get(service_domains_url, headers=headers)

# Check if the request was successful
if response.status_code == 200:
    logging.info("Successfully fetched service domains")
    # Parse the JSON response
    service_domains = response.json()

    # Create a list of dictionaries with the relevant data
    data = []
    for index, domain in enumerate(service_domains):
        logging.info(f"Processing domain {index + 1} of {len(service_domains)}")
        bian_id = domain.get('bianId')
        service_domain_name = domain.get('name')
        description = domain.get('roleDefinition')

        # Send a GET request to the /serviceDomainByBianId/{bianId} API
        details_response = requests.get(f"{service_domain_details_url}/{bian_id}", headers=headers)
        functional_pattern: str = "N/A"
        asset_type: str = "N/A"
        generic_artefact: str = "N/A"

        # Check if the request was successful
        if details_response.status_code == 200:
            logging.info(f"Successfully fetched details for BIAN ID {bian_id}")
            # Parse the JSON response for the functional pattern
            domain_details = details_response.json()

            logging.info(f"Service Domain: {service_domain_name} || Functional Pattern: {domain_details[0].get('characteristics').get('functionalPattern')} || Asset Type: {domain_details[0].get('characteristics').get('assetType')} || Generic Artefact Type: {domain_details[0].get('characteristics').get('genericArtefactType')}")
            functional_pattern = domain_details[0].get('characteristics').get('functionalPattern')
            asset_type = domain_details[0].get('characteristics').get('assetType')
            generic_artefact = domain_details[0].get('characteristics').get('genericArtefactType')

        # Append the data to the list
        data.append({
            "BIAN ID": bian_id,
            "Service Domain": service_domain_name,
            "Description": description,
            "Functional Pattern": functional_pattern,
            "Asset Type" : asset_type,
            "Generic Artefact" : generic_artefact
        })

    # Create a DataFrame from the data
    df = pd.DataFrame(data)

    # Save the DataFrame to an Excel file
    file_name = "../../../output/service_domains_with_functional_patterns.xlsx"
    df.to_excel(file_name, index=False)

    logging.info("Saved data to Excel file")

    # Load the workbook and select the active worksheet
    wb = load_workbook(file_name)
    ws = wb.active

    # Define the table
    tab = Table(displayName="ServiceDomains", ref=f"A1:F{len(df) + 1}")

    # Add a default style with striped rows and banded columns
    style = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False,
                           showLastColumn=False, showRowStripes=True, showColumnStripes=True)
    tab.tableStyleInfo = style

    # Add the table to the worksheet
    ws.add_table(tab)

    # Save the workbook
    wb.save(file_name)

    logging.info(f"Spreadsheet created successfully: {file_name}")
else:
    logging.error(f"Failed to retrieve service domains. Status code: {response.status_code}")
    logging.error(f"Response: {response.text}")
