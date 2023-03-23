from dnacentersdk import DNACenterAPI, ApiError
from requests.auth import HTTPBasicAuth
import pandas as pd
import requests
import urllib3
import textfsm
import json
import time
import netaddr
import config


# DNA-Center Instance
base_url = config.base_url
username = config.dnac_username
password = config.dnac_password

# Intent-Based API Version
version = '1.3.3'
urllib3.disable_warnings()

# DNA sdk Session creation
dnac = DNACenterAPI(username=username, password=password, base_url=base_url, version=version, verify=False)


# Define SPECIFIC MAC address grammar
class Mac_Custom(netaddr.mac_unix):
    pass


Mac_Custom.word_fmt = '%.2X'


# ---------- DNA INTENT API FUNCTIONS ----------
def getAuthToken():
    """
    Intent-based Authentication API call
    The token obtained using this API is required to be set as value to the X-Auth-Token HTTP Header
    for all API calls to Cisco DNA Center.
    :return: Token STRING
    """
    url = '{}/dna/system/api/v1/auth/token'.format(base_url)
    # Make the POST Request
    resp = requests.post(url, auth=HTTPBasicAuth(username, password), verify=False)

    # Validate Response
    if 'error' in resp.json():
        print('ERROR: Failed to retrieve Access Token!')
        print('REASON: {}'.format(resp.json()['error']))

    else:
        print("Response", resp.json())
        return resp.json()['Token']


def getTopology(token):
    """
    Intent-based Topology API call
    Authentication Token for DNAC access
    :return: Return a LIST of Network Sites
    """
    url = '{}/dna/intent/api/v1/topology/site-topology'.format(base_url)
    hdr = {
        'x-auth-token': token,
        'Content-Type': 'application/json',
        'Accept': 'application/json'
    }
    # Make the GET Request
    resp = requests.get(url, headers=hdr, verify=False)

    # Validate Response
    if 'error' in resp.json():
        print('ERROR: Failed to retrieve Site Topology!')
        print('REASON: {}'.format(resp.json()['error']))
        return []
    elif 'message' in resp.json():
        print('ERROR: Unable to retrieve Site Topology!')
        print('REASON: {}'.format(resp.json()['message']))
        return []
    else:
        return resp.json()['response']['sites']


def getSiteByType(token, type):
    """
    Intent-based Site API call
    :param token: Authentication Token for DNAC access
    :param type: Category of Site (Area/Building/Floor)
    :return: Return a LIST of Sites of a certain Type
    """
    url = '{}/dna/intent/api/v1/site?type={}'.format(base_url, type)
    hdr = {
        'x-auth-token': token,
        'Content-Type': 'application/json',
        'Accept': 'application/json'
    }

    # Make the GET Request
    resp = requests.get(url, headers=hdr, verify=False)

    # Validate Response
    if 'error' in resp.json():
        print('ERROR: Failed to retrieve Sites of a Certain type!')
        print('REASON: {}'.format(resp.json()['error']))
        return []
    elif 'message' in resp.json():
        print('ERROR: Unable to retrieve Sites of a Certain type!')
        print('REASON: {}'.format(resp.json()['message']))
        return []
    else:
        return resp.json()['response']


def getSiteHealth(token):
    """
    Intent-based Site Health API call
    :param token: Authentication Token for DNAC access
    :return: Return a LIST of Sites and their Overall Health metrics
    """
    url = '{}/dna/intent/api/v1/site-health'.format(base_url)
    hdr = {
        'x-auth-token': token,
        'Content-Type': 'application/json',
        'Accept': 'application/json'
    }
    # Make the GET Request
    resp = requests.get(url, headers=hdr, verify=False)

    # Validate Response
    if 'error' in resp.json():
        print('ERROR: Failed to retrieve Overall Site Health!')
        print('REASON: {}'.format(resp.json()['error']))
        return []
    elif 'message' in resp.json():
        print('ERROR: Unauthorized to retrieve Overall Site Health!')
        print('REASON: {}'.format(resp.json()['message']))
        return []
    else:
        print("Response in 138", resp.json())
        return resp.json() #['response']


def getSiteDevices(token, siteId):
    """
    Intent-based Site Membership API call
    :param token: Authentication Token for DNAC access
    :param siteId: Parent site ID
    :return: Return a DICTIONARY of Site
    """
    url = '{}/dna/intent/api/v1/membership/{}'.format(base_url, siteId)
    hdr = {
        'x-auth-token': token,
        'Content-Type': 'application/json',
        'Accept': 'application/json'
    }
    # Make the GET Request
    resp = requests.get(url, headers=hdr, verify=False)

    # Validate Response
    if 'error' in resp.json():
        print('ERROR: Failed to retrieve Site Membership!')
        print('REASON: {}'.format(resp.json()['error']))
        return []
    elif 'message' in resp.json():
        print('ERROR: Unauthorized to Site Membership!')
        print('REASON: {}'.format(resp.json()['message']))
        return []
    else:
        if 'errorCode' in resp.json()['site']['response']:
            print('ERROR: Invalid SiteID used!')
            print('REASON: {}'.format(resp.json()['site']['response']['message']))
            return []
        # IF Successful
        return resp.json()


def getSiteDeviceDetails(token, macAddress):
    """
    Intent-based Device Details API call
    Using DNAC SDK to handle rate limit, as script queries Site Devices
    :param token: Authentication Token for DNAC access
    :param macAddress: Device MAC Address
    :return:
    """

    mac = netaddr.EUI(macAddress, dialect=Mac_Custom)
    hdr = {'x-auth-token': token,
           'Content-Type': 'application/json',
           'Accept': 'application/json'
           }
    # Make the GET Request
    resp = dnac.devices.get_device_detail(identifier='macAddress', search_by=str(mac))
    if 'errorCode' in resp:
        print('ERROR: Failed to retrieve Device Details!')
        print('RESULT: {}'.format(resp.json()))
        return resp
    if 'response' in resp:
        return resp['response']
    else:
        print('WARNING: Failed to retrieve Device Details!')
        print('RESULT: {}'.format(resp))


def getNetworkWirelessControllers(token):
    """
    Intent-based Device List API call w/ filter criteria:
    family = Wireless Controller
    :param token: Authentication Token for DNAC access
    :return:
    """
    url = '{}/dna/intent/api/v1/network-device?family=Wireless Controller'.format(base_url)
    hdr = {
        'x-auth-token': token,
        'Content-Type': 'application/json',
        'Accept': 'application/json'
    }
    # Make the GET Request
    resp = requests.get(url, headers=hdr, verify=False)

    # Validate Response
    if 'error' in resp.json():
        print('ERROR: Failed to retrieve Overall Site Health!')
        print('REASON: {}'.format(resp.json()['error']))
        return []
    elif 'message' in resp.json():
        print('ERROR: Unauthorized to retrieve Overall Site Health!')
        print('REASON: {}'.format(resp.json()['message']))
        return []
    else:
        # Tailor output
        wlcs = []
        for wirelessControllers in resp.json()['response']:
            controllerInfo = [wirelessControllers['id'], wirelessControllers['series'], wirelessControllers['hostname']]
            wlcs.append(controllerInfo)
        return wlcs


def getEndDeviceDetails(token, macAddress, floor):
    """
    Intent-based GET Client details API call
    Using DNAC SDK to handle rate limit, as script queries Site User Devices
    :param token: Authentication Token for DNAC access
    :param macAddress: Associated AP Mac Address
    :param floor: Site Floor number
    :return:
    """
    mac = netaddr.EUI(macAddress, dialect=Mac_Custom)
    hdr = {
        'entity_type': 'mac_address',
        'entity_value': str(mac),
        'x-auth-token': token,
        'Content-Type': 'application/json',
        'Accept': 'application/json'
    }
    # Make the GET Request
    resp = dnac.users.get_user_enrichment_details(headers=hdr)
    if 'errorCode' in resp:
        # Tailor output
        endDevice = {
            'macAddress': str(mac),
            'hostName': '',
            'hostOs': '',
            'healthScore': 0,
            'ssid': '',
            'frequency': '',
            'channel': '',
            'floor': floor
        }
        print('WARNING: Invalid response from User Device!')
        return endDevice
    response = resp[0]
    if 'userDetails' in response:
        userDetails = response['userDetails']
        # Tailor output
        endDevice = {
            'macAddress': userDetails['hostMac'],
            'hostName': userDetails['hostName'],
            'hostOs': userDetails['hostOs'],
            'healthScore': userDetails['healthScore'][0]['score'],
            'ssid': userDetails['ssid'],
            'frequency': userDetails['frequency'],
            'channel': userDetails['channel'],
            'floor': floor
        }
        print('|------------------> {} ({}), OS = {}, Health = {}'.format(endDevice['hostName'], endDevice['macAddress'], endDevice['hostOs'], endDevice['healthScore']))
        return endDevice

    else:
        print('WARNING: Failed to retrieve User Details!')
        print('RESULT: {}'.format(resp))


def postCommandRunnerSession(token, device_id, device_name, command):
    """
    Intent-based GET Command Runner API call
    :param token: Authentication Token for DNAC access
    :param device_id: Device for Command Runner session
    :param device_name: Hostname of WLC
    :param command: Command String to execute
    :return:
    """
    ios_cmd = command
    body = {
        "name": "{}".format(ios_cmd),
        "commands": [ios_cmd],
        "deviceUuids": [device_id]
    }
    url = "{}/api/v1/network-device-poller/cli/read-request".format(base_url)
    header = {
        'x-auth-token': token,
        'Content-Type': 'application/json',
        'Accept': 'application/json'
    }

    # Make the POST Request
    resp = requests.post(url, data=json.dumps(body), headers=header, verify=False)

    if 'message' in resp.json():
        print('ERROR: Unauthorized to run Command Runner!')
        print('REASON: {}'.format(resp.json()['message']))
        return []
    elif 'errorCode' in resp.json()['response']:
        print('ERROR: Invalid input to run Command Runner!')
        print('REASON: {}'.format(resp.json()['detail']))
        return
    else:
        task_id = resp.json()['response']['taskId']
        getTaskInfo(token, device_name, ios_cmd, task_id)
        return task_id


def getTaskInfo(token, device_name, command, task_id):
    """
    Intent-based GET Task Info API call
    :param token: Authentication Token for DNAC access
    :param device_name: Hostname for Device
    :param command: Command String to execute
    :param task_id: Create Command Runner Task ID
    :return:
    """
    url = "{}/api/v1/task/{}".format(base_url, task_id)
    hdr = {
        'x-auth-token': token,
        'Content-Type': 'application/json',
        'Accept': 'application/json'
    }

    # Make the GET Request
    resp = requests.get(url, headers=hdr, verify=False)

    if 'message' in resp.json():
        print('ERROR: Unauthorized to run Command Runner!')
        print('REASON: {}'.format(resp.json()['message']))
        return []
    elif 'errorCode' in resp.json()['response']:
        print('ERROR: Invalid input to run Command Runner!')
        print('REASON: {}'.format(resp.json()['detail']))
        return
    else:
        file_id = resp.json()['response']['progress']
        # IF Successful command file made
        if "fileId" in file_id:
            # Parse Task file id
            unwanted_chars = '{"}'
            for char in unwanted_chars:
                file_id = file_id.replace(char, '')
            file_id = file_id.split(':')
            file_id = file_id[1]
            getCmdOutput(token, device_name, command, file_id)
        # ELSE Try again (can be a lag for task file to be created)
        else:
            getTaskInfo(token, device_name, command, task_id)


def getCmdOutput(token, device_name, command, file_id):
    """
    Intent-based GET File API call
    :param token: Authentication Token for DNAC access
    :param device_name: Hostname of Device
    :param command: Command String to execute
    :param file_id: File Id
    :return:
    """
    url = "{}/api/v1/file/{}".format(base_url, file_id)
    hdr = {
        'x-auth-token': token,
        'Content-Type': 'application/json',
        'Accept': 'application/json'
    }

    # Make the GET Request
    resp = requests.get(url, headers=hdr, verify=False)

    if 'message' in resp.json():
        print('ERROR: Unauthorized to run Command Runner!')
        print('REASON: {}'.format(resp.json()['message']))
        return []
    elif 'response' in resp.json()[0]:
        print('ERROR: Invalid input to run Command Runner!')
        print('REASON: {}'.format(resp.json()[0]['response']['detail']))
        return
    else:
        commandResult = resp.json()[0]['commandResponses']
        # Ensure Success
        if command in commandResult['SUCCESS']:
            output = commandResult['SUCCESS'][command]
            with open('../DNAC_Assurance_Reporting/Wireless_Controller_Reports/{}.txt'.format(device_name),
                      'w') as outfile:
                outfile.write(output)
        else:
            print('ERROR: Invalid response from Command Runner!')
            print('RESPONSE: {}'.format(commandResult['FAILURE']))


# ---------- HELPER FUNCTIONS ----------
def retrieveWLCClients(device_name):
    """
    Retrieve the Command Runner output stored in Wireless_Controller_Reports/{{device_id}}.txt
    Grammar of retrieved output defined in TextFSM file
    :param device_name: Controller Hostname used to identify Report Files
    :return:
    """
    # Pull Data
    input_file = open('../DNAC_Assurance_Reporting/Wireless_Controller_Reports/{}.txt'.format(device_name),
                      encoding='utf-8')
    cli_data = input_file.read()
    input_file.close()

    # Parse CLI output based off of defined Grammar
    with open("sh_client_sum_template.textfsm") as f:
        template = textfsm.TextFSM(f)
    fsm_results = template.ParseText(cli_data)

    lineNumber = 0
    associatedClients = []

    # From CLI output Tailor returned value
    for row in fsm_results:
        if lineNumber > 5 and row != ['(Cisco', 'Controller)']:
            wirelessClient = {'macAddress': row[0], 'ap': row[1]}
            associatedClients.append(wirelessClient)
        else:
            lineNumber += 1

    return associatedClients


def exportData(networkSite, floorList, floorDevices, clientDetails):
    """
    Helper function used to Export script output into Excel document for further use
    :param networkSite: Network Site for Report
    :param floorList: Floors associated to Network Site
    :param floorDevices: Received per floor metrics
    :param clientDetails: Received per client metrics
    :return:
    """
    # Create Report
    print()
    print("Floor List", floorList)
    print("Floor Devices", floorDevices)
    print()
    print("Client Details", clientDetails)
    print()
    with pd.ExcelWriter('{}.xlsx'.format(networkSite['name'])) as writer:
        # Write DATA SET 1 - Worksheet = Overall_Health
        siteDataFrame = pd.DataFrame(networkSite, index=[0])
        siteDataFrame.to_excel(excel_writer=writer, sheet_name='Overall_Health', startrow=0, startcol=0)

        # Write DATA SET 2 - Worksheet = Floor_Health
        rowCounter = 0
        for floor in floorList:
            floorDataFrame = pd.DataFrame(floor, index=[0])
            if rowCounter == 0:
                floorDataFrame.to_excel(excel_writer=writer, sheet_name='Floor_Health', startrow=rowCounter, startcol=0)
                rowCounter += 2
            else:
                floorDataFrame.to_excel(excel_writer=writer, sheet_name='Floor_Health',
                                        startrow=rowCounter, startcol=0, header=False)
                rowCounter += 1

        # Write DATA SET 3 -  Worksheet = Client_Count
        rowCounter = 0
        for dev in floorDevices:
            devDataFrame = pd.DataFrame(dev, index=[0])
            if rowCounter == 0:
                devDataFrame.to_excel(excel_writer=writer, sheet_name='Client_Count', startrow=rowCounter, startcol=0)
                rowCounter += 2
            else:
                devDataFrame.to_excel(excel_writer=writer, sheet_name='Client_Count',
                                      startrow=rowCounter, startcol=0, header=False)
                rowCounter += 1

        # Write DATA SET 4 -  Worksheet = Client_Details
        rowCounter = 0
        for client in clientDetails:
            clientDataFrame = pd.DataFrame(client, index=[0])
            if rowCounter == 0:
                clientDataFrame.to_excel(excel_writer=writer, sheet_name='Client_Details', startrow=rowCounter,
                                         startcol=0)
                rowCounter += 2
            else:
                clientDataFrame.to_excel(excel_writer=writer, sheet_name='Client_Details',
                                         startrow=rowCounter, startcol=0, header=False)
                rowCounter += 1


# ----- Main Script -----
# TRY-EXCEPT block to catch API Errors
try:
    # Retrieve valid Authentication Token with DNA Center instance
    authToken = getAuthToken()
    print(authToken)

    # Retrieve ALL Network Sites (Buildings & Geo-Areas)
    site_healthList = getSiteHealth(authToken)
    print("SITE HEALTH", site_healthList)

    # Setup for Command Runner Session
    wireless_Controller = "device_id"
    device_CLI_Command1 = 'show client summary'
    device_CLI_Command2 = 'show wireless client summary'
    complete_Client_List = []

    # Retrieve ALL Network Wireless Controllers
    wireless_Network_Controllers = getNetworkWirelessControllers(authToken)
    for wireless_Controller in wireless_Network_Controllers:
        # Tailor WLC CLI command based off of Device Model
        if '9800' in wireless_Controller[1]:
            postCommandRunnerSession(authToken, wireless_Controller[0], wireless_Controller[2], device_CLI_Command2)
        else:
            postCommandRunnerSession(authToken, wireless_Controller[0], wireless_Controller[2], device_CLI_Command1)

        # Retrieve ALL Network Wireless Clients
        controller_Client_List = retrieveWLCClients(wireless_Controller[2])
        for client in controller_Client_List:
            complete_Client_List.append(client)

    print('DNA Network Sites: ')

    # Retrieve ALL Network Sites (Buildings & Floors)
    site_BuildingList = getSiteByType(authToken, 'building')
    site_floorList = getSiteByType(authToken, 'floor')

    print("SITE BUILDING LIST", site_BuildingList)
    for site in site_BuildingList:
        print(site['name'])

    print()
    input1 = input('Select the site that you would like to report on [name/all]: ')

    if input1 == 'all':
        # Print ALL sites
        for site in site_BuildingList:

            for location in site_healthList:
                if location['siteName'] == site['name']:
                    site = location
                    break

            siteFloors = []
            siteDevices = []
            siteClientDetails = []

            # REFRESH TOKEN HERE IF NEEDED!
            # authToken = getAuthToken()

            print(site['name'])

            for floor in site_floorList:
                if site['name'] in floor['siteNameHierarchy']:

                    # Creating custom Data set for each Floor
                    customFloorDict = {
                        'id': floor['id'],
                        'name': floor['name']
                    }
                    floorMembership = getSiteDevices(authToken, floor['id'])
                    floorDevices = floorMembership['device'][0]['response']
                    deviceCount = 0
                    floorClientCount = 0
                    floor_Health_Total = 0

                    print('|---> {}'.format(floor['name']))

                    for device in floorDevices:

                        # DO ONLY IF DEVICE IS AN AP
                        # if device['family'] == 'Unified AP':
                        # REFRESH TOKEN HERE IF NEEDED

                        if 'macAddress' in device:
                            details = {}
                            device_Details = getSiteDeviceDetails(authToken, device['macAddress'])

                            print('|----------> {}'.format(device['hostname']))

                            details['name'] = device_Details['nwDeviceName']
                            details['floor'] = floor['name']

                            if 'clientCount' in device_Details:
                                details['clientCount'] = device_Details['clientCount']

                            deviceCount += 1
                            if 'overallHealth' in device_Details:
                                floor_Health_Total += device_Details['overallHealth']
                            clientCount = 0

                            for client in complete_Client_List:
                                # Retrieve all Clients associated to the Device
                                if client['ap'] == device_Details['nwDeviceName']:
                                    clientCount += 1
                                    # if clientCount % 5 == 0:
                                        # time.sleep(30)
                                    siteClientDetails.append(
                                        getEndDeviceDetails(authToken, client['macAddress'], floor['name']))
                            details['clientCount'] = clientCount
                            siteDevices.append(details)

                    customFloorDict['nwDeviceCount'] = deviceCount
                    if deviceCount != 0:
                        customFloorDict['floorHealth'] = floor_Health_Total / deviceCount
                    siteFloors.append(customFloorDict)

                # EXPORT
                exportData(site, siteFloors, siteDevices, siteClientDetails)

    else:
        for site in site_BuildingList:
            if input1 == site['name']:

                for location in site_healthList:
                    if location['siteName'] == site['name']:
                        site = location
                        break

                siteFloors = []
                siteDevices = []
                siteClientDetails = []

                # REFRESH TOKEN HERE IF NEEDED!
                # authToken = getAuthToken()

                print(site['name'])

                for floor in site_floorList:
                    if site['name'] in floor['siteNameHierarchy']:

                        # Creating custom Data set for each Floor
                        customFloorDict = {
                            'id': floor['id'],
                            'name': floor['name']
                        }
                        floorMembership = getSiteDevices(authToken, floor['id'])
                        floorDevices = floorMembership['device'][0]['response']
                        deviceCount = 0
                        floorClientCount = 0
                        floor_Health_Total = 0

                        print('|---> {}'.format(floor['name']))

                        for device in floorDevices:

                            # DO ONLY IF DEVICE IS AN AP
                            # if device['family'] == 'Unified AP':
                            # REFRESH TOKEN HERE IF NEEDED
                            if 'macAddress' in device:
                                details = {}

                                device_Details = getSiteDeviceDetails(authToken, device['macAddress'])

                                print('|----------> {}'.format(device['hostname']))

                                details['name'] = device_Details['nwDeviceName']
                                details['floor'] = floor['name']

                                if 'clientCount' in device_Details:
                                    details['clientCount'] = device_Details['clientCount']

                                deviceCount += 1
                                if 'overallHealth' in device_Details:
                                    floor_Health_Total += device_Details['overallHealth']
                                clientCount = 0

                                for client in complete_Client_List:
                                    # Retrieve all Clients associated to the Device
                                    if client['ap'] == device_Details['nwDeviceName']:
                                        clientCount += 1
                                        # if clientCount % 5 == 0:
                                            # time.sleep(30)
                                        siteClientDetails.append(
                                            getEndDeviceDetails(authToken, client['macAddress'], floor['name']))
                                details['clientCount'] = clientCount
                                siteDevices.append(details)

                        customFloorDict['nwDeviceCount'] = deviceCount
                        if deviceCount != 0:
                            customFloorDict['floorHealth'] = floor_Health_Total / deviceCount
                        siteFloors.append(customFloorDict)

                    # EXPORT
                    exportData(site, siteFloors, siteDevices, siteClientDetails)

# If any API Errors are caught during execution
except ApiError as e:
    print(e)
