import os
import time
import meraki
import config
import datetime
import time
import os
import schedule
from datetime import timedelta
from openpyxl import load_workbook
import pandas as pd
from time import sleep
import statistics
import pprint
import json
import pandas as pd

api_key =  config.api_key           
dashboard = meraki.DashboardAPI(api_key)
today=datetime.datetime.today().strftime("%Y-%M-%D")
#print_console=True,output_log=False


def getOrgs():
    response = dashboard.organizations.getOrganizations()
    return response



def get_networks(organizationid):
    networks = []
    response = dashboard.organizations.getOrganizationNetworks(organizationId=organizationid)  

    for network in response:
        item = {}
        
        item["id"] = network["id"]
        item["name"] = network["name"]
        
        networks.append(item)

    return networks



def excel(network_name,AP_Status,AP_Bandwidth,Client_Performance,Wireless_Health,client,switch_statuses):

    sheet_names=['AP_Status','AP_Bandwidth','Client_Performance','Wireless_Health']
    file=network_name+'.xlsx'
    if os.path.isfile(file):
        
        writer = pd.ExcelWriter(file, engine="openpyxl", mode="a", if_sheet_exists="replace")
        fileb = open(file, "rb")
        writer.book = load_workbook(fileb)
        writer.sheets = dict((ws.title, ws) for ws in writer.book.worksheets)
        print(writer.sheets)
        if 'Wireless_Health' in writer.sheets:
            print("Yes present")
            converts = {
                "assoc": str,
                "auth": str,
                "dhcp": str,
                "dns": str,
                "success": str,
                "failedClients": str
            }
            reader = pd.read_excel(file, sheet_name='Wireless_Health', converters=converts)
            data = pd.DataFrame(Wireless_Health)
            df = pd.concat([reader, data])
            df.to_excel(writer, sheet_name='Wireless_Health', index=False)

        if 'AP_Status' in writer.sheets:
            converts = {
                "name": str,
                "status": str
            }
            reader = pd.read_excel(file, sheet_name='AP_Status', converters=converts)
            data = pd.DataFrame(AP_Status)
            df = pd.concat([reader, data])
            df.to_excel(writer, sheet_name='AP_Status', index=False)
            
            
        if 'AP_Bandwidth' in writer.sheets:
            converts = {
                "startTs": str,
                "endTs": str,
                "totalKbps": str,
                "sentKbps": str,
                "receivedKbps": str
                }
            reader = pd.read_excel(file, sheet_name='AP_Bandwidth', converters=converts)
            data = pd.DataFrame(AP_Bandwidth)
            df = pd.concat([reader, data])
            df.to_excel(writer, sheet_name='AP_Bandwidth', index=False)

        if 'Client_Performance' in writer.sheets:
            converts = {
                "mac": str,
                "id": str,
                "name": str,
                "ip": str,
                "ap": str,
                "ssid": str,
                "snr": str,
                "rssi": str
            }
            reader = pd.read_excel(file, sheet_name='Client_Performance', converters=converts)
            data = pd.DataFrame(data=Client_Performance,index=[0])
            print("Client Performance", data)
            df = pd.concat([reader, data])
            df.to_excel(writer, sheet_name='Client_Performance', index=False)

        if 'Client_Count' in writer.sheets:
            converts = {
                "Network": str,
                "count": str,
                "Date": str
                }
            reader = pd.read_excel(file, sheet_name='Client_Count', converters=converts)
            data = pd.DataFrame(client)
            df = pd.concat([reader, data])
            df.to_excel(writer, sheet_name='Client_Count', index=False)

        if "Switch_Status" in writer.sheets:
            converts = {
                "serial": str,
                "status": str,
                "name": str,
                "networkId": str
            }
            reader = pd.read_excel(file, sheet_name='Switch_Status', converters=converts)
            data = pd.DataFrame(switch_statuses)
            df = pd.concat([reader, data])
            df.to_excel(writer, sheet_name='Switch_Status', index=False)

        
        writer.close()

    else:
        writer = pd.ExcelWriter(file, engine="xlsxwriter")
        df = pd.DataFrame(Wireless_Health)
        print(df)
        df.to_excel(writer, sheet_name='Wireless_Health', startrow=0, startcol=0)

        # Data for AP_Status
        df = pd.DataFrame(AP_Status)
        print(df)
        df.to_excel(writer, sheet_name='AP_Status', startrow=0, startcol=0)

        # Data for AP Bandwidth
        df = pd.DataFrame(AP_Bandwidth)
        print(df)
        df.to_excel(writer, sheet_name='AP_Bandwidth', startrow=0, startcol=0)

        # Data for Client Performance
        df = pd.DataFrame(Client_Performance,index=[0])
        print(df)
        df.to_excel(writer, sheet_name='Client_Performance', startrow=0, startcol=0)

        # Data for Client Count
        df = pd.DataFrame(client,index=[0])
        print(df)
        df.to_excel(writer, sheet_name='Client_Count', startrow=0, startcol=0)

        #Data for switch status
        df = pd.DataFrame(switch_statuses,index=[0])
        print(df)
        df.to_excel(writer, sheet_name='Switch_Status', startrow=0, startcol=0)
        writer.close()




def days_between(d1, d2):
    d1 = datetime.datetime.strptime(d1, "%Y-%m-%d")
    d2 = datetime.datetime.strptime(d2, "%Y-%m-%d")
    return (d2 - d1).days

def plus_days(d, p):
    d = datetime.datetime.strptime(d, "%Y-%m-%d") + datetime.timedelta(days=p)
    return d.strftime("%Y-%m-%d")


def get_wireless_health(net_id, t0, t1):

    try:
        if days_between(t0, t1) <= 7:
            print("Yes")
            headers = {
                "Content-Type": "application/json",
                "Accept": "application/json",
                "X-Cisco-Meraki-API-KEY": config.api_key
                }

            connStats = dashboard.wireless.getNetworkWirelessConnectionStats(
                networkId=net_id, t0=t0, t1=t1
            )
            failedConns = dashboard.wireless.getNetworkWirelessFailedConnections(
                networkId=net_id, t0=t0, t1=t1)

            failedClients = []
            for failedConn in failedConns:
                failedClients.append((failedConn['clientMac'], failedConn['failureStep'], failedConn['type']))
            connStats['failedClients'] = failedClients

            if connStats == None:
                connStats = {
                    'assoc': 0,
                    'auth': 0,
                    'dhcp': 0,
                    'dns': 0,
                    'success': 0,
                    'failedClients': []
                }

            return connStats

        else:
            aggConnStats = dict()
            aggConnStats['assoc'] = 0
            aggConnStats['auth'] = 0
            aggConnStats['dhcp'] = 0
            aggConnStats['dns'] = 0
            aggConnStats['success'] = 0
            aggConnStats['failedClients'] = []
            while days_between(t0, t1) > 0:
                if days_between(t0, t1) <= 6:
                    t = t1
                else:
                    t = plus_days(t0, 6)
                print(f"t0 {t0} t {t}")
                connStats = dashboard.wireless.getNetworkWirelessConnectionStats(
                networkId=net_id, t0=t0, t1=t1
                )
                failedConns = dashboard.wireless.getNetworkWirelessFailedConnections(
                networkId=net_id, t0=t0, t1=t1)
                print(f"connStats {connStats}")
                print(f"failedConns {failedConns}")

                if connStats == None:
                    connStats = {
                        'assoc': 0,
                        'auth': 0,
                        'dhcp': 0,
                        'dns': 0,
                        'success': 0
                    }

                aggConnStats['assoc'] += connStats['assoc']
                aggConnStats['auth'] += connStats['auth']
                aggConnStats['dhcp'] += connStats['dhcp']
                aggConnStats['dns'] += connStats['dns']
                aggConnStats['success'] += connStats['success']
                for failedConn in failedConns:
                    aggConnStats['failedClients'].append((failedConn['clientMac'], failedConn['failureStep'], failedConn['type']))

                t0 = t

        return aggConnStats

    except Exception as e:
        print(e)



def get_all_aps(org_id):
    macs = []
    names = []
    serials = []
    try:
        devices = dashboard.organizations.getOrganizationDevices(
        org_id, total_pages='all'
        )
        if devices != None:
            for device in devices:
                if "MR" in device['model']:
                    macs.append(device['mac'])
                    names.append(device['name'])
                    serials.append(device['serial'])
    except Exception as e:
        print(f"An error has errored during getting AP from {org['name']}. Error: {e}")
    return [names, macs, serials]



def get_all_network_aps(networkId):
    macs = []
    serials = []
    try:
        devices = dashboard.networks.getNetworkDevices(
            networkId
        )

        if devices != None:
            for device in devices:
                if "MR" in device['model']:
                    macs.append(device['mac'])
                    serials.append(device['serial'])
            
    except Exception as e:
        print(f"An error has errored during getting AP from {networkId}. Error: {e}")
    return [macs, serials]


    print()





def poll_ap_status(org_id,networkId):
    orgs = getOrgs()
    aps = get_all_network_aps(networkId)
    print("Start polling AP Status...")
    down_aps = dict()
    APStatuses={
            'name':[],
            'status': []
        }
    try:
        device_statuses = dashboard.organizations.getOrganizationDevicesStatuses(
                org_id, total_pages='all')
        print()
        if device_statuses != None:
            #device_statuses = json.loads(device_statuses.text)
            for device_status in device_statuses:
                if device_status['mac'] in aps[0]:
                    ap_name = device_status['name']
                    ap_status = device_status['status']

                    APStatuses['name'].append(ap_name)
                    APStatuses['status'].append(ap_status)
            APStatuses['Date']=today
            return APStatuses
    except Exception as e:
        print(f"An error has errored during AP status polling. Error: {e}")



def poll_ap_bandwidth(networkId):

    global polled_APBandwidth
    AP_BANDWIDTH_POLLING_INTERVAL=4
    
    aps = get_all_network_aps(networkId)
    macs = aps[0]
    serials = aps[1]
    print("Start polling AP Bandwidth...")

    try:
        for i in range(len(serials)):
            get_device = dashboard.devices.getDevice(
                serials[i]
            )
            net_id = get_device['networkId']
            print()
            get_usage = dashboard.wireless.getNetworkWirelessUsageHistory(
                networkId= get_device['networkId'],deviceSerial=serials[i],timespan=300, resolution=300
            )
        return get_usage
            
    except Exception as e:
        print(f"An error has errored during AP bandwidth polling. Error: {e}")



def numberofclients(network_id, network_name):
    try:
        clientCount={
            "Network":[],
            "count": []   
        }
        count=[]
        client_history = dashboard.wireless.getNetworkWirelessClientCountHistory(networkId=network_id, timespan=900,resolution=300)
        for j in range(len(client_history)):
            if client_history[j]['clientCount']==None:
                count.append(0)
            elif client_history[j]['clientCount'] == '':
                count.append(0)           
                continue
            else:
               count.append(client_history[j]['clientCount'])
    
        if len(count) != 0:
            clientCount['count'].append(max(count))
            clientCount['Network'].append(network_name)
        
        return clientCount

    except:
        print("There is no AP in the")
        clientCount['count'].append("0")
        clientCount['Network'].append(network_name)
        return clientCount         




def switchStatus(organization_id, network_id):

    switch_statuses = { "serial" : [],
                        "status" : [],
                        "name": [],
                        "networkId": []
                    }
    device_statuses=[]

    try:
        response = dashboard.organizations.getOrganizationDevicesStatuses(
        organization_id, total_pages='all'
            )

        response1 = dashboard.networks.getNetworkDevices(
                network_id
        )

        for device in response:
            for device1 in response1:
                if device['serial']==device1['serial']:
                    device_statuses.append(device)
            else:
                continue

 
        if device_statuses != None:
            for device_status in device_statuses:
                if "MS" in device_status['model']:
                    switch_statuses["serial"].append(device_status["serial"])
                    switch_statuses["status"].append(device_status["status"])
                    switch_statuses["name"].append(device_status["name"])
                    switch_statuses["networkId"].append(device_status["networkId"])
                    
            return switch_statuses
    except Exception as e:
        print(f"An error has errored during switch status polling. Error: {e}")




def poll_client_performance(net_id):
    global polled_Client
    orgs = getOrgs()
    print("Start polling Client performance...")

    try:
        ap_clients = dict()
        clients = dashboard.networks.getNetworkClients(
        net_id
        )
        if clients != None:
            for client in clients:
                if not client['ssid'] == None and client['status'] == "Online":
                    client_id = client['id']
                    get_signal = dashboard.wireless.getNetworkWirelessSignalQualityHistory(
                        networkId=net_id,clientId=client_id,timespan=3600,resolution=3600
                    )
                    if get_signal != None:
                        ap_clients = {
                            'mac':client['mac'],
                            "id": client['id'],
                            "name": client['description'],
                            "ip": client['ip'],
                            "ap": client['recentDeviceName'],
                            "ssid": client['ssid'],
                            "snr": get_signal[0]['snr'],
                            "rssi": get_signal[0]['rssi']
                        }
                                    
                    else:
                        ap_clients = {
                        'mac':None,
                        "id":None,
                        "name":None,
                        "ip": None,
                        "ap": None,
                        "ssid": None,
                        "snr": None,
                        "rssi": None
                        }
                
                else:
                    ap_clients = {
                    'mac':None,
                    "id":None,
                    "name":None,
                    "ip": None,
                    "ap": None,
                    "ssid": None,
                    "snr": None,
                    "rssi": None
                    }

        else:
            ap_clients = {
                'mac':None,
                "id":None,
                "name":None,
                "ip": None,
                "ap": None,
                "ssid": None,
                "snr": None,
                "rssi": None
            }
        ap_clients['Date']=today
        return ap_clients

    except Exception as e:
        print(f"An error has errored during Client performance polling. Error: {e}")

    polled_Client = True



try:
    orgs=getOrgs()
    for org in orgs:
        networks=get_networks(org['id'])
        print("################## Organization #########",org['name'])
        for network in networks:
            print("############# Network ############", network['name'])
            print()
            if poll_ap_status(org['id'], network['id']) != None:
                ap_status=poll_ap_status(org['id'], network['id'])
                #Poll AP Bandwidth 
                ap_bandwidth=poll_ap_bandwidth(network['id'])
                #Polling Client Performance from the APs
                Client_Performance=poll_client_performance(network['id']) 
                
                #Polling wireless health 
                #Wireless_Health=get_wireless_health(network['id'],'2022-2-10','2022-2-15')
                Wireless_Health=get_wireless_health(network['id'],'2022-2-10','2022-2-15')
                Wireless_Health['Date']=today

                client=numberofclients(network['id'],network['name'])
                client['Date']=today
                
                switch_statuses=switchStatus(org['id'],network['id'])
                switch_statuses['Date']=today
                

                excel(network['name'],ap_status,ap_bandwidth,Client_Performance,Wireless_Health, client,switch_statuses)
            else:
                continue

except e:
    print("Error",e)
