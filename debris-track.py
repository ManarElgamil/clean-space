import requests
import json
import configparser
import xlsxwriter
import time
from datetime import datetime

class MyError(Exception):
    def __init__(self,args):
        Exception.__init__(self,"my exception was raised with arguments {0}".format(args))
        self.args = args

uriBase                = "https://www.space-track.org"
requestLogin           = "/ajaxauth/login"
requestCmdAction       = "/basicspacedata/query" 

# ✅ Replace Starlink search with debris search
requestFindDebris      = "/class/tle_latest/OBJECT_TYPE/DEBRIS/ORDINAL/1/format/json/orderby/NORAD_CAT_ID%20asc"
requestOMMDebris1      = "/class/omm/NORAD_CAT_ID/"
requestOMMDebris2      = "/orderby/EPOCH%20asc/format/json"

GM = 398600441800000.0
GM13 = GM ** (1.0/3.0)
MRAD = 6378.137
PI = 3.14159265358979
TPI86 = 2.0 * PI / 86400.0

config = configparser.ConfigParser()
config.read("./SLTrack.ini")
configUsr = config.get("configuration","username")
configPwd = config.get("configuration","password")
configOut = config.get("configuration","output")
siteCred = {'identity': configUsr, 'password': configPwd}

workbook = xlsxwriter.Workbook(configOut)
worksheet = workbook.add_worksheet()
z0_format = workbook.add_format({'num_format': '#,##0'})
z1_format = workbook.add_format({'num_format': '#,##0.0'})
z2_format = workbook.add_format({'num_format': '#,##0.00'})
z3_format = workbook.add_format({'num_format': '#,##0.000'})

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
wsline = 3

with requests.Session() as session:
    resp = session.post(uriBase + requestLogin, data = siteCred)
    if resp.status_code != 200:
        raise MyError(resp, "POST fail on login")

    # ✅ Get debris list instead of Starlink
    resp = session.get(uriBase + requestCmdAction + requestFindDebris)
    if resp.status_code != 200:
        print(resp)
        raise MyError(resp, "GET fail when requesting debris data")

    retData = json.loads(resp.text)
    debrisIds = [e['NORAD_CAT_ID'] for e in retData]

    maxs = 1
    for s in debrisIds:
        resp = session.get(uriBase + requestCmdAction + requestOMMDebris1 + s + requestOMMDebris2)
        if resp.status_code != 200:
            print(resp)
            raise MyError(resp, "GET fail for debris object " + s)

        retData = json.loads(resp.text)
        for e in retData:
            print("Scanning debris " + e['OBJECT_NAME'] + " at epoch " + e['EPOCH'])
            mmoti = float(e['MEAN_MOTION'])
            ecc = float(e['ECCENTRICITY'])
            
            worksheet.write(wsline, 0, int(e['NORAD_CAT_ID']))
            worksheet.write(wsline, 1, e['OBJECT_NAME'])
            worksheet.write(wsline, 2, e['EPOCH'])
            worksheet.write(wsline, 3, float(e['REV_AT_EPOCH']))
            worksheet.write(wsline, 4, float(e['INCLINATION']),z1_format)
            worksheet.write(wsline, 5, ecc,z3_format)
            worksheet.write(wsline, 6, mmoti,z1_format)

            sma = GM13 / ((TPI86 * mmoti) ** (2.0 / 3.0)) / 1000.0
            apo = sma * (1.0 + ecc) - MRAD
            per = sma * (1.0 - ecc) - MRAD
            smak = sma * 1000.0
            orbT = 2.0 * PI * ((smak ** 3.0) / GM) ** (0.5)
            orbV = (GM / smak) ** (0.5)

            worksheet.write(wsline, 7, apo,z1_format)
            worksheet.write(wsline, 8, per,z1_format)
            worksheet.write(wsline, 9, (apo + per)/2.0,z1_format)
            worksheet.write(wsline, 10, float(e['RA_OF_ASC_NODE']),z1_format)
            worksheet.write(wsline, 11, float(e['ARG_OF_PERICENTER']),z1_format)
            worksheet.write(wsline, 12, float(e['MEAN_ANOMALY']),z1_format)
            worksheet.write(wsline, 13, sma,z1_format)
            worksheet.write(wsline, 14, orbT,z0_format)
            worksheet.write(wsline, 15, orbV,z0_format)
            
            wsline += 1

        maxs += 1
        if maxs > 18:
            print("Sleeping 60 seconds to avoid rate limits...")
            time.sleep(60)
            maxs = 1

session.close()
workbook.close()
print("✅ Completed collecting debris data")
