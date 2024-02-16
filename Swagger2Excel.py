import sys
import requests
import json
from openpyxl import Workbook

def get_swagger_endpoints(swagger_url_or_file):
    try:
        if swagger_url_or_file.startswith('http'):
            response = requests.get(swagger_url_or_file)
            if response.status_code == 200:
                swagger_spec = response.json()
            else:
                print(f"Failed to retrieve Swagger spec. Status code: {response.status_code}")
                return []
        else:
            with open(swagger_url_or_file, 'r') as file:
                swagger_spec = json.load(file)

        paths = swagger_spec.get('paths', {})
        endpoints = []
        for path, methods in paths.items():
            for method, _ in methods.items():
                endpoints.append((method.upper(), path))
        return endpoints
    except Exception as e:
        print(f"An error occurred: {e}")
        return []

def write_to_excel(endpoints, excel_file):
    wb = Workbook()
    ws = wb.active
    ws.append(["Method", "Endpoint"])
    for endpoint in endpoints:
        ws.append(endpoint)
    wb.save(excel_file)
    print(f"API endpoints have been written to {excel_file}")

def print_help():
    print("Usage: python script.py  <swagger_url_or_file>  <output_fileName>.xlsx")
    print("  swagger_url_or_file:  URL or file path to the Swagger JSON file")
    print("  output_excel_file:  Path to the output Excel file")

if __name__ == "__main__":
    if len(sys.argv) != 3 or sys.argv[1] in ['-h', '--help']:
        print_help()
        sys.exit(1)

    swagger_url_or_file = sys.argv[1]
    output_excel_file = sys.argv[2]
    endpoints = get_swagger_endpoints(swagger_url_or_file)
    if endpoints:
        write_to_excel(endpoints, output_excel_file)
