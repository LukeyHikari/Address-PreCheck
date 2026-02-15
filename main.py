import pandas as pd
import requests as rq
import openpyxl as op
import time
import os
from dotenv import load_dotenv

load_dotenv()
API_KEY = os.getenv('GOOGLE_API_KEY')

# TODO: Add error handling for addresss lines that are invalid for the API.

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
        with open('api_response.json', 'w') as f:
            f.write(response.text)
        return response.json()
    else:
        print(f"Error: {response.status_code} - {response.text}")

def extract_address_components(address):
    components = {}
    for component in address.get('addressComponents', []):
        component_type = component.get('componentType', '')
        components[component_type] = {
            'text': component.get('componentName', {}).get('text', ''),
            'confirmation': component.get('confirmationLevel', ''),
            'inferred': component.get('inferred', False),
        }
    return components, len(components)

def classify_address(address):
    address_results = address.get('result', {})

    # Address Components
    address_components, component_count = extract_address_components(address_results.get('address', {}))
    final_zip = 'N/A'
    current_zip = address_components.get('postal_code', {}).get('text', '')
    response_address_line = address_results.get('address', {}).get('postalAddress', {}).get('addressLines', [])

    # Verdict details
    verdict = {
        'inputGranularity': address_results.get('verdict', {}).get('inputGranularity', ''),
        'validationGranularity': address_results.get('verdict', {}).get('validationGranularity', ''),
        'geocodeGranularity': address_results.get('verdict', {}).get('geocodeGranularity', ''),
        'addressComplete': address_results.get('verdict', {}).get('addressComplete', False),
        'hasUnconfirmedComponents': address_results.get('verdict', {}).get('hasUnconfirmedComponents', False),
        'hasInferredComponents': address_results.get('verdict', {}).get('hasInferredComponents', False),
        'hasReplacedComponents': address_results.get('verdict', {}).get('hasReplacedComponents', False),
        'possibleNextAction': address_results.get('verdict', {}).get('possibleNextAction', ''),
        'hasSpellCorrectedComponents': address_results.get('verdict', {}).get('hasSpellCorrectedComponents', False)
    }

    # Classification Logic (Granularity Levels)
    # Other: Address probably nonexistent
    # Route: Street/Route found but no premise-level details, street number might be wrong
    # Premise: Full address with premise-level details
    # Sub_Premise: Has apartment, etc.
    if verdict['validationGranularity'] == 'OTHER' and verdict['geocodeGranularity'] == 'OTHER':
        print("Address is classified as OTHER. No further processing.")
    elif verdict['validationGranularity'] == 'ROUTE' and verdict['geocodeGranularity'] in ['ROUTE', 'PREMISE']: 
        print("Address is classified as ROUTE. Street/Route found but no premise-level details.")   
    elif verdict['validationGranularity'] in ['PREMISE', 'SUB_PREMISE'] and verdict['geocodeGranularity'] in ['PREMISE', 'SUB_PREMISE']:
        print("Address is classified as PREMISE. Found on both validation and geocode levels, might not have +4 details.")
        # Postal Suffix Extraction Logic
        inferred_suffix = address_components.get('postal_code_suffix', {}).get('inferred', False)
        if inferred_suffix and component_count > 6:
            final_zip = current_zip + "-" + address_components.get('postal_code_suffix', {}).get('text', '')

    return {
        "final_zip": current_zip if final_zip == 'N/A' else final_zip,
        "response_address_line": response_address_line,
        "validation_granularity": verdict['validationGranularity'],
        "geocode_granularity": verdict['geocodeGranularity'],
        "possible_next_action": verdict['possibleNextAction']
    }

def get_addresses_from_excel(file_path):
    df = pd.read_excel(file_path, dtype={"ZIP": str})
    return df

def main():
    file_path = 'test.xlsx'  # Path to your Excel file
    addresses_df = get_addresses_from_excel(file_path)
    output_df = pd.DataFrame(columns=
                             ['Input Address', 'City', 'State', 'Postal Code', 'Final ZIP', 'Response Address Line', 
                              'Validation Granularity', 'Geocode Granularity', 'Possible Next Action']
                             )

    for index, row in addresses_df.iterrows():
        address = row['Address 1'] + ' ' + row['Address 2'] if pd.notna(row['Address 2']) else row['Address 1']
        city = row['City']
        state = row['State']
        postal_code = row['ZIP']
        postal_suffix = 'N/A'
        
        print(f"Validating: {address}, {city}, {state} {postal_code}")
        api_response = get_address_details(address, city, state, postal_code)
        address_classification = classify_address(api_response)

        # Append the classification results to the output DataFrame
        output_df = pd.concat([output_df, pd.DataFrame({
            'Input Address': [address],
            'City': [city],
            'State': [state],
            'Postal Code': [postal_code],
            'Final ZIP': [address_classification['final_zip']],
            'Response Address Line': [', '.join(address_classification['response_address_line'])],
            'Validation Granularity': [address_classification['validation_granularity']],
            'Geocode Granularity': [address_classification['geocode_granularity']],
            'Possible Next Action': [address_classification['possible_next_action']]
        })], ignore_index=True)

        time.sleep(0.1)  # To avoid hitting API rate limits

        with pd.ExcelWriter('output.xlsx', engine='openpyxl') as writer:
            output_df.to_excel(writer, index=False, sheet_name='Address Validation Results')

if __name__ == "__main__":
    main()
