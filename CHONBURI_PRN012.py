import requests
import xml.etree.ElementTree as ET
import datetime
import csv
import os

# Set CSV path
CSV_FILE_PATH = 'C:\\Users\\cboonsangiem\\OneDrive - Sarens\\BUSINESS UNIT THAILAND - IT\\71 Documents\\Sarens Thailand - CHOBURI_PRN012.csv'

def get_printer_status():

    # Global Const
    global CSV_FILE_PATH

    # Check CSV_FILE_PATH
    if len(CSV_FILE_PATH) == 0:
        print('Enter the CSV path.')
        return

    # Send SOAP request
    url = 'http://10.56.3.12/ssm/Management/Anonymous/StatusConfig'
    headers = {
        'Soapaction': '"http://www.fujixerox.co.jp/2003/12/ssm/management/statusConfig#GetAttribute"',
        'Content-Type': 'text/xml; charset="UTF-8"',
    }
    payload = """<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"><soap:Header><msg:MessageInformation xmlns:msg="http://www.fujixerox.co.jp/2014/08/ssm/management/message"><msg:MessageExchangeType>RequestResponse</msg:MessageExchangeType><msg:MessageType>Request</msg:MessageType><msg:Action>http://www.fujixerox.co.jp/2003/12/ssm/management/statusConfig#GetAttribute</msg:Action><msg:From><msg:Address>http://www.fujixerox.co.jp/2014/08/ssm/management/soap/epr/client</msg:Address><msg:ReferenceParameters/></msg:From></msg:MessageInformation></soap:Header><soap:Body><cfg:GetAttribute xmlns:cfg="http://www.fujixerox.co.jp/2003/12/ssm/management/statusConfig"><cfg:Object name="urn:fujixerox:names:ssm:1.0:management:root" offset="0"/><cfg:Object name="urn:fujixerox:names:ssm:1.0:management:WirelessLANConfig" offset="0"/><cfg:Object name="urn:fujixerox:names:ssm:1.0:management:MarkerSupply" offset="0"/><cfg:Object name="urn:fujixerox:names:ssm:1.0:management:DeviceStatus" offset="0"/></cfg:GetAttribute></soap:Body></soap:Envelope>"""

    # Get SOAP response
    response = requests.post(url, headers=headers, data=payload, verify=False)


    if response.status_code == 200:
        soap_response = response.text.replace('<?xmlversion="1.0"encoding="UTF-8"?>', '')

        # Parse the XML
        root = ET.fromstring(soap_response)

        # Define namespaces
        namespaces = {
            'env': 'http://schemas.xmlsoap.org/soap/envelope/',
            'ns': 'http://www.fujixerox.co.jp/2003/12/ssm/management/statusConfig'
        }

        # Find the GetAttributeResponse element
        get_attribute_response = root.find('.//ns:GetAttributeResponse', namespaces)

        # Convert XML elements to dictionary
        def xml_element_to_dict(element):
            result = {}
            for child in element:
                result[child.get('name')] = child.text
            return result

        # Convert each Object element to a dictionary
        objects = get_attribute_response.findall('.//ns:Object', namespaces)
        data = [xml_element_to_dict(obj) for obj in objects]

        # Get the current date
        current_date = datetime.datetime.now()

        csv_data = {
            'Device': data[0]['Name'],
            'Date': current_date,
            'Toner_C_Remaining': data[12]['Remaining'],
            'Toner_C_PageRemaining': data[12]['PageRemaining'],
            'Toner_C_DateInstalled': data[12]['DateInstalled'],
            'Toner_M_Remaining': data[11]['Remaining'],
            'Toner_M_PageRemaining': data[11]['PageRemaining'],
            'Toner_M_DateInstalled': data[11]['DateInstalled'],
            'Toner_Y_Remaining': data[10]['Remaining'],
            'Toner_Y_PageRemaining': data[10]['PageRemaining'],
            'Toner_Y_DateInstalled': data[10]['DateInstalled'],
            'Toner_K_Remaining': data[13]['Remaining'],
            'Toner_K_PageRemaining': data[13]['PageRemaining'],
            'Toner_K_DateInstalled': data[13]['DateInstalled'],
            'Drum_C_Remaining': data[8]['Remaining'],
            'Drum_C_PageRemaining': data[8]['PageRemaining'],
            'Drum_M_Remaining': data[7]['Remaining'],
            'Drum_M_PageRemaining': data[7]['PageRemaining'],
            'Drum_Y_Remaining': data[6]['Remaining'],
            'Drum_Y_PageRemaining': data[6]['PageRemaining'],
            'Drum_K_Remaining': data[9]['Remaining'],
            'Drum_K_PageRemaining': data[9]['PageRemaining'],
            'Dev_C_LifeState': data[4]['LifeState'],
            'Dev_M_LifeState': data[3]['LifeState'],
            'Dev_Y_LifeState': data[2]['LifeState'],
            'Dev_K_LifeState': data[5]['LifeState'],
            'Waste_Toner_LifeState': data[14]['LifeState']
        }


        # Check if the CSV file exists
        file_exists = os.path.exists(CSV_FILE_PATH)


        # Append the dictionary data to the CSV file
        with open(CSV_FILE_PATH, mode="a", newline="") as csv_file:
            writer = csv.DictWriter(csv_file, fieldnames=csv_data.keys())

            # Write headers if the file doesn't exist
            if not file_exists:
                writer.writeheader()
            
            # Write data
            writer.writerow(csv_data)

        print("Data appended to CSV file successfully.")
    else:
        print('Cannot get SOAP response.')

if __name__ == '__main__':
    get_printer_status()
