#Credit
#Rupak Kannan 5/18/2023 Programmer
#Website API used: https://github.com/open-meteo/open-meteo/discussions
#ALl Credit goes to https://open-meteo.com/


#!pip install openpyxl
#!pip install pandas

import requests
import json
#import pprint
import openpyxl
import pandas as pd
import xlsxwriter
import datetime

#subprocess.check_call([sys.executable, '-m', 'pip', 'install', ',modulename'])
#pp = pprint.PrettyPrinter(indent=4)
#--------------------------Starting program--------------------------------------------
data_choice = input("Would you like to recieve historical data or a forecast? (1 for forecast, any key for historical): ")
if(data_choice == "1"):
    try:
      latitude = float(input("enter the latitude of the specificed area you want forecasted: "))
      longtitude = float(input("enter the longtitude of the specificed area you want forecasted: "))
      forecast_hours = int(input("Enter the amount of hours you want forecasted: "))
    except:
       print("defaulting to Rolla, Mo for 24 hours")
       longtitude = -91.7715
       latitude = 37.9485
       forecast_hours = 24
    #--------------------------------------------------------------------------------------

    response = requests.get("https://api.open-meteo.com/v1/forecast?latitude=%f&longitude=%f&hourly=temperature_2m,relativehumidity_2m,dewpoint_2m,precipitation_probability,surface_pressure,windspeed_180m,winddirection_180m,uv_index,is_day,direct_radiation,direct_normal_irradiance" % (latitude, longtitude)).text
    response_info = json.loads(response)

    weather_info = []
    block = []
    i=0

    for i in range(0, forecast_hours):
        for hour in response_info['hourly']:

          block.append(response_info['hourly'][hour][i])
        weather_info.append(block)
        block = []

    print("Forecast for the next " + str(forecast_hours) + " specified hours, starting from start of the current hour")
    weather_forecast_for_today = pd.DataFrame(data=weather_info, columns=['Time', 'Temperature 2M (C)', 'Humidity 2M (%) ', 'Dew Point 2M (C)', 'Precipitation Probability (%)', 'Surface Pressure (hpa)', 'Wind Speed 180M (km/h)', 'Wind Direction 180M (Degrees)', 'UV Index', 'Daytime (1 = True, 0 = False)', 'Direct radiation (w/m^2)', 'Direct normal irradiance (w/m^2)'])
    weather_forecast_for_today.head(forecast_hours)




    path = '%f_%f_forecast_weather_data.xlsx' % (longtitude,latitude)
    writer = pd.ExcelWriter(path)
    weather_forecast_for_today.to_excel(writer, sheet_name='info', index=False, na_rep='NaN')

    for column in weather_forecast_for_today:
        column_width = max(weather_forecast_for_today[column].astype(str).map(len).max(), len(column))
        col_idx = weather_forecast_for_today.columns.get_loc(column)
        writer.sheets['info'].set_column(col_idx, col_idx, column_width)

    writer.close()

else:
    try:
        latitude = float(input("enter the latitude of the specificed area you want forecasted: "))
        longtitude = float(input("enter the longtitude of the specificed area you want forecasted: "))
        days_back = int(input("Enter the amount of days back you want historical data of: "))
    except:
        print("defaulting to Rolla, Mo for 365 days")
        longtitude = -91.7715
        latitude = 37.9485
        days_back = 14
    # ------------------------
    start_date = datetime.datetime.today().date()-datetime.timedelta(days=days_back)
    end_date = datetime.datetime.today().date()

    response = requests.get(
        "https://archive-api.open-meteo.com/v1/archive?latitude=%f&longitude=%f&start_date=%s&end_date=%s&hourly=temperature_2m,relativehumidity_2m,dewpoint_2m,surface_pressure,precipitation,cloudcover,direct_radiation,direct_normal_irradiance,windspeed_10m,winddirection_10m" %
        (latitude, longtitude, start_date,  end_date)).text
    response_info = json.loads(response)

    weather_info = []
    block = []
    i = 0

    for i in range(0, days_back*24):
        for hour in response_info['hourly']:
            block.append(response_info['hourly'][hour][i])

        weather_info.append(block)
        block = []




    print("Historical weather data recorded")
    weather_forecast_for_today = pd.DataFrame(data=weather_info, columns=['Time', 'Temperature 2M (C)', 'Humidity 2M (%) ', 'Dew Point 2M (C)', 'Surface Pressure (hpa)', 'Precipitation (mm)', 'Cloud cover (%)', 'Direct radiation (w/m^2)', 'Direct normal irradiance (w/m^2)', 'Wind Speed 10M (km/h)', 'Wind Direction 10M (Degrees)'])
    weather_forecast_for_today.head(days_back*24)

    path = '%f_%f_historical_weather_data.xlsx' % (longtitude, latitude)
    writer = pd.ExcelWriter(path)
    weather_forecast_for_today.to_excel(writer, sheet_name='info', index=False, na_rep='NaN')

    for column in weather_forecast_for_today:
        column_width = max(weather_forecast_for_today[column].astype(str).map(len).max(), len(column))
        col_idx = weather_forecast_for_today.columns.get_loc(column)
        writer.sheets['info'].set_column(col_idx, col_idx, column_width)

    writer.close()