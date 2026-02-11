import pandas as pd
import requests as rq
import openpyxl as op
import time
import os
from dotenv import load_dotenv

load_dotenv()
API_KEY = os.getenv('GOOGLE_API_KEY')

def get_address_details(address, city, state, postal_code):
    url = f"https://addressvalidation.googleapis.com/v1:validateAddress?key={API_KEY}"
    payload = {
        "address": {
            "regionCode": "US",
            "locality": city,
            "administrativeArea": state,
            "postalCode": postal_code,
            "addressLines": [address],
        },
        "enableUspsCass": True 
    }
    response = rq.post(url, json=payload)

    if response.status_code == 200:
        return response.json()

    else:
        print(f"Error: {response.status_code} - {response.text}")

def extract_address_components(address):
    components = {}
    for component in address.get('address', {}).get('addressComponents', []):
        component_type = component.get('componentType', '')
        components[component_type] = {
            'text': component.get('componentName', {}).get('text', ''),
            'languageCode': component.get('componentName', {}).get('languageCode', ''),
            'confirmation': component.get('confirmationLevel', ''),
            'inferred': component.get('inferred', False),
            'spellCorrected': component.get('spellCorrected', False),
            'replaced': component.get('replaced', False),
            'unexpected': component.get('unexpected', False)
        }
    return components, len(components)

def classify_address(address):
    address_results = address.get('result', {})

    # Address Components
    address_components, component_count = extract_address_components(
                    address_results.get('address', {}).get('addressComponents', [])
                )
    
    # Verdict details
    verdict = {
        'inputGranularity': address_results.get('verdict', {}).get('inputGranularity', ''),
        'validationGranularity': address_results.get('verdict', {}).get('validationGranularity', ''),
        'geocodeGranularity': address_results.get('verdict', {}).get('geocodeGranularity', ''),
        'addressComplete': address_results.get('verdict', {}).get('addressComplete', False),
        'hasUnconfirmedComponents': address_results.get('verdict', {}).get('hasUnconfirmedComponents', False),
        'hasInferredComponents': address_results.get('verdict', {}).get('hasInferredComponents', False),
        'hasReplacedComponents': address_results.get('verdict', {}).get('hasReplacedComponents', False),
        'possibleNextActions': address_results.get('verdict', {}).get('possibleNextActions', []),
        'hasSpellCorrectedComponents': address_results.get('verdict', {}).get('hasSpellCorrectedComponents', False)
    }

    if verdict['validationGranularity'] == 'OTHER':
        print("Address is classified as OTHER. No further processing.")
    elif verdict['validationGranularity'] == 'PREMISE':
        # Postal Suffix Extraction Logic
        final_zip = 'N/A'
        current_zip = address_results.get('address', {}).get('postalCode', '')
        inferred_suffix = address_components.get('postalCodeSuffix', {}).get('inferred', False)
        if inferred_suffix and component_count == 7:
            final_zip = current_zip + address_components.get('postalCodeSuffix', {}).get('text', '')
        else: final_zip = current_zip

    return final_zip #, other things to return later

def get_addresses_from_excel(file_path):
    df = pd.read_excel(file_path, dtype={"ZIP": str})
    return df

def main():
    file_path = 'test.xlsx'  # Path to your Excel file
    addresses_df = get_addresses_from_excel(file_path)

    for index, row in addresses_df.iterrows():
        address = row['Address 1'] + ' ' + row['Address 2'] if pd.notna(row['Address 2']) else row['Address 1']
        city = row['City']
        state = row['State']
        postal_code = row['ZIP']
        postal_suffix = 'N/A'
        
        print(f"Validating: {address}, {city}, {state} {postal_code}")
        postal_suffix = get_address_details(address, city, state, postal_code)
        time.sleep(0.1)  # To avoid hitting API rate limits

if __name__ == "__main__":
    main()
