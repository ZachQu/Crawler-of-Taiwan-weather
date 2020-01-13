'''
Copyright <2019> <Jinhan Wu>

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
'''

import requests
import numpy as np
import pandas as pd
from openpyxl.workbook import Workbook
months = ["01", '02', '03', '04', '05', '06', '07', '08', '09', '10', '11', '12']
days = ["01", '02', '03', '04', '05', '06', '07', '08', '09', '10', '11', '12', '13', '14', '15', '16', '17', '18', '19', '20', '21', '22', '23', '24', '25', '26', '27', '28', '29', '30', '31']

holidays = [
'2016-01-01', '2016-01-02', '2016-01-03', '2016-01-09', '2016-01-10', '2016-01-16', '2016-01-17', '2016-01-23', '2016-01-24', '2016-01-31',
'2016-02-06', '2016-02-07', '2016-02-08', '2016-02-09', '2016-02-10', '2016-02-11', '2016-02-12', '2016-02-13', '2016-02-14', '2016-02-20', '2016-02-21', '2016-02-27', '2016-02-28', '2016-02-29',
'2016-03-05', '2016-03-06', '2016-03-12', '2016-03-13', '2016-03-19', '2016-03-20', '2016-03-26', '2016-03-27',
'2016-04-02', '2016-04-03', '2016-04-04', '2016-04-05', '2016-04-09', '2016-04-10', '2016-04-16', '2016-04-17', '2016-04-23', '2016-04-24', '2016-04-30',
'2016-05-01', '2016-05-02', '2016-05-07', '2016-05-08', '2016-05-14', '2016-05-15', '2016-05-21', '2016-05-22', '2016-05-28', '2016-05-29',
'2016-06-02', '2016-06-05', '2016-06-09', '2016-06-10', '2016-06-11', '2016-06-12', '2016-06-18', '2016-06-19', '2016-06-25', '2016-06-26',
'2016-07-02', '2016-07-03', '2016-07-09', '2016-07-10', '2016-07-16', '2016-07-17', '2016-07-23', '2016-07-24', '2016-07-30', '2016-07-31',
'2016-08-06', '2016-08-07', '2016-08-13', '2016-08-14', '2016-08-20', '2016-08-21', '2016-08-27', '2016-08-28',
'2016-09-03', '2016-09-04', '2016-09-11', '2016-09-15', '2016-09-16', '2016-09-17', '2016-09-18', '2016-09-24', '2016-09-25',
'2016-10-01', '2016-10-02', '2016-10-08', '2016-10-09', '2016-10-10', '2016-10-15', '2016-10-16', '2016-10-22', '2016-10-23', '2016-10-29', '2016-10-30',
'2016-11-05', '2016-11-06', '2016-11-12', '2016-11-13', '2016-11-19', '2016-11-20', '2016-11-26', '2016-11-27',
'2016-12-03', '2016-12-04', '2016-12-10', '2016-12-11', '2016-12-17', '2016-12-18', '2016-12-24', '2016-12-25', '2016-12-31',
'2017-01-01', '2017-01-02', '2017-01-07', '2017-01-08', '2017-01-14', '2017-01-15', '2017-01-21', '2017-01-22', '2017-01-27', '2017-01-28', '2017-01-29', '2017-01-30', '2017-01-31',
'2017-02-01', '2017-02-04', '2017-02-05', '2017-02-11', '2017-02-12', '2017-02-18', '2017-02-19', '2017-02-25', '2017-02-26', '2017-02-27', '2017-02-28',
'2017-03-04', '2017-03-05', '2017-03-11', '2017-03-12', '2017-03-18', '2017-03-19', '2017-03-25', '2017-03-26',
'2017-04-01', '2017-04-02', '2017-04-03', '2017-04-04', '2017-04-08', '2017-04-09', '2017-04-15', '2017-04-16', '2017-04-22', '2017-04-23', '2017-04-29', '2017-04-30',
'2017-05-01', '2017-05-06', '2017-05-07', '2017-05-13', '2017-05-14', '2017-05-20', '2017-05-21', '2017-05-27', '2017-05-28', '2017-05-29', '2017-05-30',
'2017-06-04', '2017-06-10', '2017-06-11', '2017-06-17', '2017-06-18', '2017-06-24', '2017-06-25',
'2017-07-01', '2017-07-02', '2017-07-08', '2017-07-09', '2017-07-15', '2017-07-16', '2017-07-22', '2017-07-23', '2017-07-29', '2017-07-30',
'2017-08-05', '2017-08-06', '2017-08-12', '2017-08-13', '2017-08-19', '2017-08-20', '2017-08-26', '2017-08-27',
'2017-09-02', '2017-09-03', '2017-09-09', '2017-09-10', '2017-09-16', '2017-09-17', '2017-09-23', '2017-09-24',
'2017-10-01', '2017-10-04', '2017-10-07', '2017-10-08', '2017-10-09', '2017-10-10', '2017-10-14', '2017-10-15', '2017-10-21', '2017-10-22', '2017-10-28', '2017-10-29',
'2017-11-04', '2017-11-05', '2017-11-11', '2017-11-12', '2017-11-18', '2017-11-19', '2017-11-25', '2017-11-26',
'2017-12-02', '2017-12-03', '2017-12-09', '2017-12-10', '2017-12-16', '2017-12-17', '2017-12-23', '2017-12-24', '2017-12-30', '2017-12-31',
'2018-01-01', '2018-01-06', '2018-01-07', '2018-01-13', '2018-01-14', '2018-01-20', '2018-01-21', '2018-01-27', '2018-01-28'
]

# holidays = []

date_start = [2016,1]
date_end = [2018,1]
time_range = [8,18] # working time as 8:00 to 18:00

target_url = 'https://e-service.cwb.gov.tw/HistoryDataQuery/DayDataController.do?command=viewMain&station=467650&stname=%25E6%2597%25A5%25E6%259C%2588%25E6%25BD%25AD&datepicker=2016-01-03#'

def wet_bulb_temp(T, RH):
    return T*np.arctan(0.151977*(RH+8.313659)**0.5)+np.arctan(T+RH)-np.arctan(RH-1.676331)+0.00391838*(RH)**1.5*np.arctan(0.023101*RH)-4.686035

def download_data(weather_url):
    response = requests.get(weather_url)
    csv = str(response.text)
    lines = csv.split('\n')
    pres_list = []
    Dtemp_list=[]
    Wtemp_list=[]
    RH_list = []
    wind_list = []
    rain_list = []
    sun_list = []
    sunP_list = []
    CDH18d_list = []
    CDH18w_list = []
    for i in range(time_range[0]-1, time_range[1]-1):
        pres_list.append(float(lines[828+i*21][lines[828+i*21].find('<td>')+len('<td>'):lines[828+i*21].find('&nbsp')]))
        Dtemp_list.append(float(lines[830+i*21][lines[830+i*21].find('<td>')+len('<td>'):lines[830+i*21].find('&nbsp')]))
        RH_list.append(float(lines[832+i*21][lines[832+i*21].find('<td>')+len('<td>'):lines[832+i*21].find('&nbsp')]))
        wind_list.append(float(lines[833+i*21][lines[833+i*21].find('<td>')+len('<td>'):lines[833+i*21].find('&nbsp')]))
        rain_list.append(float(lines[838+i*21][lines[838+i*21].find('<td>')+len('<td>'):lines[838+i*21].find('&nbsp')]))
        try:
            sun_list.append(float(lines[839+i*21][lines[839+i*21].find('<td>')+len('<td>'):lines[839+i*21].find('&nbsp')]))
        except:
            sun_list.append(0.0)
        try:
            sunP_list.append(float(lines[840+i*21][lines[840+i*21].find('<td>')+len('<td>'):lines[840+i*21].find('&nbsp')]))
        except:
            sunP_list.append(0.0)
        Wtemp_list.append(wet_bulb_temp(Dtemp_list[-1], RH_list[-1]))
        if Dtemp_list[-1]>18.0:
            CDH18d_list.append(Dtemp_list[-1]-18.0)
        if Wtemp_list[-1]>18.0:
            CDH18w_list.append(Wtemp_list[-1]-18.0)

    print ("pressure", pres_list)
    print ("Dry-bulb temperature", Dtemp_list)
    print ("Wet-bulb temperature", Wtemp_list)
    print ("RH", RH_list)
    print ("wind", wind_list)
    print ("rain", rain_list)
    print ("sun", sun_list)
    print ("sunP", sunP_list)
    print ("CDHd18", CDH18d_list)
    print ("CDHw18", CDH18w_list)
    return pres_list, Dtemp_list, Wtemp_list, RH_list, wind_list, rain_list, sun_list, sunP_list, CDH18d_list, CDH18w_list

pres_ave_month_list = []
Dtemp_ave_month_list = []
Wtemp_ave_month_list = []
RH_ave_month_list = []
wind_ave_month_list = []
rain_total_month_list = []
sun_total_month_list = []
sunP_total_month_list = []
pres_std_month_list = []
Dtemp_std_month_list = []
Wtemp_std_month_list = []
RH_std_month_list = []
wind_std_month_list = []
CDH18d_total_month_list = []
CDD18d_total_month_list = []
CDH18w_total_month_list = []
CDD18w_total_month_list = []
time_stamp_list = []

for year in range(date_start[0], date_end[0]+1):
    for month in months:
        if year == date_end[0] and int(date_end[1])<int(month):
            break
        else:
            pres_day_list = []
            Dtemp_day_list = []
            Wtemp_day_list = []
            RH_day_list = []
            wind_day_list = []
            rain_day_list = []
            sun_day_list = []
            sunP_day_list = []
            CDH18d_day_list = []
            CDD18d_day_list = []
            CDH18w_day_list = []
            CDD18w_day_list = []
            for day in days:
                time_stamp = '{}-{}-{}#'.format(year, month, day)
                if time_stamp[:-1] not in holidays:
                    print ("Starting getting data on {}".format(time_stamp))
                    url= target_url[:-11]+time_stamp
                    try:
                        pres_total, Dtemp_total, Wtemp_total, RH_total, wind_total, rain_total, sun_total, sunP_total, CDH18d_total, CDH18w_total = download_data(url)
                        print ("Finish getting data on {}".format(time_stamp))
                        if np.mean(Dtemp_total) > 18.0:
                            CDD18d_day_list.append(np.mean(Dtemp_total)-18.0)
                        if np.mean(Wtemp_total) > 18.0:
                            CDD18w_day_list.append(np.mean(Wtemp_total)-18.0)

                        pres_day_list.extend(pres_total)
                        Dtemp_day_list.extend(Dtemp_total)
                        Wtemp_day_list.extend(Wtemp_total)
                        RH_day_list.extend(RH_total)
                        wind_day_list.extend(wind_total)
                        rain_day_list.extend(rain_total)
                        sun_day_list.extend(sun_total)
                        sunP_day_list.extend(sunP_total)
                        CDH18d_day_list.extend(CDH18d_total)
                        CDH18w_day_list.extend(CDH18w_total)
                    except:
                        pass
        pres_ave_month = np.mean(pres_day_list)
        Dtemp_ave_month = np.mean(Dtemp_day_list)
        Wtemp_ave_month = np.mean(Wtemp_day_list)
        RH_ave_month = np.mean(RH_day_list)
        wind_ave_month = np.mean(wind_day_list)
        rain_total_month = np.sum(rain_day_list)
        sun_total_month = np.sum(sun_day_list)
        sunP_total_month = np.sum(sunP_day_list)
        CDH18d_total_month = np.sum(CDH18d_day_list)
        CDD18d_total_month = np.sum(CDD18d_day_list)
        CDH18w_total_month = np.sum(CDH18w_day_list)
        CDD18w_total_month = np.sum(CDD18w_day_list)

        pres_std_month = np.std(pres_day_list)
        Dtemp_std_month = np.std(Dtemp_day_list)
        Wtemp_std_month = np.std(Wtemp_day_list)
        RH_std_month = np.std(RH_day_list)
        wind_std_month = np.std(wind_day_list)

        pres_ave_month_list.append(pres_ave_month)
        Dtemp_ave_month_list.append(Dtemp_ave_month)
        Wtemp_ave_month_list.append(Wtemp_ave_month)
        RH_ave_month_list.append(RH_ave_month)
        wind_ave_month_list.append(wind_ave_month)
        rain_total_month_list.append(rain_total_month)
        sun_total_month_list.append(sun_total_month)
        sunP_total_month_list.append(sunP_total_month)
        CDH18d_total_month_list.append(CDH18d_total_month)
        CDD18d_total_month_list.append(CDD18d_total_month)
        CDH18w_total_month_list.append(CDH18w_total_month)
        CDD18w_total_month_list.append(CDD18w_total_month)

        pres_std_month_list.append(pres_std_month)
        Dtemp_std_month_list.append(Dtemp_std_month)
        Wtemp_std_month_list.append(Wtemp_std_month)
        RH_std_month_list.append(RH_std_month)
        wind_std_month_list.append(wind_std_month)

        time_stamp_list.append(time_stamp[:-4])

np.savetxt('dataV4_during_{}_to_{}_as_working_{}_to_{}.csv'.format(time_stamp_list[0], time_stamp_list[-1], time_range[0], time_range[1]), list(zip(time_stamp_list, pres_ave_month_list, pres_std_month_list, Dtemp_ave_month_list, Dtemp_std_month_list, Wtemp_ave_month_list, Wtemp_std_month_list, RH_ave_month_list, RH_std_month_list, wind_ave_month_list, wind_std_month_list, rain_total_month_list, sun_total_month_list, sunP_total_month_list, CDH18d_total_month_list, CDD18d_total_month_list, CDH18w_total_month_list, CDD18w_total_month_list)), header = "Month,Pressure,Pressure_std,Dry_Bulb_Temperature,Dry_Bulb_Temperature_std,Wet_Bulb_Temperature,Wet_Bulb_Temperature_std,RH,RH_std,Wind,Wind_std,Raintime,Suntime,SunPower,CDHd18,CDDd18,CDHw18,CDDw18", delimiter=',', comments="", fmt="%s")

data_csv = pd.read_csv('dataV4_during_{}_to_{}_as_working_{}_to_{}.csv'.format(time_stamp_list[0], time_stamp_list[-1], time_range[0], time_range[1]))
data_excel = pd.ExcelWriter('dataV4_during_{}_to_{}_as_working_{}_to_{}.xlsx'.format(time_stamp_list[0], time_stamp_list[-1], time_range[0], time_range[1]))
data_csv.to_excel(data_excel, index = False)
data_excel.save()

print ("done")
