import requests
import json
import configparser
import xlsxwriter
import time
from datetime import datetime

# Custom exception class for handling API or session errors
class MyError(Exception):
    def __init__(self,args):
        Exception.__init__(self,"my exception was raised with arguments {0}".format(args))
        self.args = args

# -----------------------------
# Space-Track API endpoints
# -----------------------------
uriBase = "https://www.space-track.org"
requestLogin = "/ajaxauth/login"                # Login endpoint
requestCmdAction = "/basicspacedata/query"     # Main query endpoint

# ✅ Get only debris updated in last 30 days (faster + relevant)
requestFindDebris = "/class/tle_latest/OBJECT_TYPE/DEBRIS/EPOCH/>now-30/orderby/NORAD_CAT_ID/format/json"

# ✅ Get only the latest OMM record for each debris object
requestOMMDebris1 = "/class/omm/NORAD_CAT_ID/"   # Base URL to get OMM data
requestOMMDebris2 = "/orderby/EPOCH%20desc/limit/1/format/json"

# -----------------------------
# Orbital constants
# -----------------------------
GM = 398600441800000.0
GM13 = GM ** (1.0/3.0)
MRAD = 6378.137
PI = 3.14159265358979
TPI86 = 2.0 * PI / 86400.0

# -----------------------------
# Read config file for credentials and output Excel file
# -----------------------------
config = configparser.ConfigParser()
config.read("./SLTrack.ini")
configUsr = config.get("configuration","username")
configPwd = config.get("configuration","password")
configOut = config.get("configuration","output")  # Example: debris.xlsx
siteCred = {'identity': configUsr, 'password': configPwd}

# -----------------------------
# Create Excel workbook and worksheet
# -----------------------------
workbook = xlsxwriter.Workbook(configOut)
worksheet = workbook.add_worksheet()

# Number formatting for Excel columns
z0_format = workbook.add_format({'num_format': '#,##0'})
z1_format = workbook.add_format({'num_format': '#,##0.0'})
z2_format = workbook.add_format({'num_format': '#,##0.00'})
z3_format = workbook.add_format({'num_format': '#,##0.000'})

# Write headers in Excel
now = datetime.now()
nowStr = now.strftime("%m/%d/%Y %H:%M:%S")
worksheet.write('A1', 'Space Debris data from ' + uriBase + " on " + nowStr)
worksheet.write('A3','NORAD_CAT_ID')
worksheet.write('B3','OBJECT_NAME')
worksheet.write('C3','EPOCH')
worksheet.write('D3','Orb')
worksheet.write('E3','Inc')
worksheet.write('F3','Ecc')
worksheet.write('G3','MnM')
worksheet.write('H3','ApA')
worksheet.write('I3','PeA')
worksheet.write('J3','AvA')
worksheet.write('K3','LAN')
worksheet.write('L3','AgP')
worksheet.write('M3','MnA')
worksheet.write('N3','SMa')
worksheet.write('O3','T')
worksheet.write('P3','Vel')
wsline = 3  # Start writing data from row 4

# Auto-adjust column widths
for i in range(16):
    worksheet.set_column(i, i, 15)

# -----------------------------
# Start a session and login
# -----------------------------
with requests.Session() as session:
    resp = session.post(uriBase + requestLogin, data=siteCred)
    if resp.status_code != 200:
        raise MyError(resp, "POST fail on login")

    # Fetch debris list updated in last 30 days
    resp = session.get(uriBase + requestCmdAction + requestFindDebris)
    if resp.status_code != 200:
        print(resp)
        raise MyError(resp, "GET fail when requesting debris data")

    retData = json.loads(resp.text)
    # ✅ Only include debris with valid NORAD_CAT_ID
    debrisIds = [e.get('NORAD_CAT_ID') for e in retData if e.get('NORAD_CAT_ID') is not None]

    # Loop through each debris object and get the latest OMM record
    maxs = 1
    for s in debrisIds:
        resp = session.get(uriBase + requestCmdAction + requestOMMDebris1 + str(s) + requestOMMDebris2)
        if resp.status_code != 200:
            print(resp)
            raise MyError(resp, "GET fail for debris object " + str(s))

        retData = json.loads(resp.text)
        for e in retData:
            # Use get() with default to avoid KeyError
            obj_name = e.get('OBJECT_NAME', 'UNKNOWN')
            epoch = e.get('EPOCH', 'UNKNOWN')
            print(f"Scanning debris {obj_name} at epoch {epoch}")

            mmoti = float(e.get('MEAN_MOTION', 0))
            ecc = float(e.get('ECCENTRICITY', 0))

            # Write raw orbital data into Excel
            worksheet.write(wsline, 0, int(e.get('NORAD_CAT_ID', 0)))
            worksheet.write(wsline, 1, obj_name)
            worksheet.write(wsline, 2, epoch)
            worksheet.write(wsline, 3, float(e.get('REV_AT_EPOCH', 0)))
            worksheet.write(wsline, 4, float(e.get('INCLINATION', 0)), z1_format)
            worksheet.write(wsline, 5, ecc, z3_format)
            worksheet.write(wsline, 6, mmoti, z1_format)

            # Calculate orbital parameters
            sma = GM13 / ((TPI86 * mmoti) ** (2.0 / 3.0)) / 1000.0 if mmoti else 0
            apo = sma * (1.0 + ecc) - MRAD
            per = sma * (1.0 - ecc) - MRAD
            smak = sma * 1000.0
            orbT = 2.0 * PI * ((smak ** 3.0) / GM) ** 0.5 if sma else 0
            orbV = (GM / smak) ** 0.5 if sma else 0

            # Write computed values to Excel
            worksheet.write(wsline, 7, apo, z1_format)
            worksheet.write(wsline, 8, per, z1_format)
            worksheet.write(wsline, 9, (apo + per)/2.0, z1_format)
            worksheet.write(wsline, 10, float(e.get('RA_OF_ASC_NODE', 0)), z1_format)
            worksheet.write(wsline, 11, float(e.get('ARG_OF_PERICENTER', 0)), z1_format)
            worksheet.write(wsline, 12, float(e.get('MEAN_ANOMALY', 0)), z1_format)
            worksheet.write(wsline, 13, sma, z1_format)
            worksheet.write(wsline, 14, orbT, z0_format)
            worksheet.write(wsline, 15, orbV, z0_format)

            wsline += 1

        maxs += 1
        # Pause every 18 objects to avoid API rate limits
        if maxs > 18:
            print("Sleeping 60 seconds to avoid rate limits...")
            time.sleep(60)
            maxs = 1

# Close session and save Excel file
session.close()
workbook.close()
print("✅ Completed collecting debris data")
