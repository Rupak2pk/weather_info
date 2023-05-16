
import requests
import json
#import pprint
import openpyxl
import pandas as pd
import xlsxwriter

#pp = pprint.PrettyPrinter(indent=4)
#--------------------------Starting program--------------------------------------------
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

response = requests.get("https://api.weather.gov/points/%f,%f" % (latitude, longtitude)).text
response_info = json.loads(response)

x = response_info['properties']['gridX']
y = response_info['properties']['gridY']
Id = response_info['properties']['gridId']
city_name = response_info['properties']['relativeLocation']['properties']['city']
state_name = response_info['properties']['relativeLocation']['properties']['state']
print("retrieving values for %s, %s" % (city_name, state_name))



response = requests.get('https://api.weather.gov/gridpoints/%s/%d,%d/forecast/hourly' % (Id, x, y)).text
response_info =  json.loads(response)


peroids = response_info['properties']['periods']
weather_info = []
#weather_info = [hour #, humidity %, short forecast, temperature F, wind direction, wind speed, daytime (True or False), dew point C, precipitation chance %,]
for hour in peroids:
  if hour['number'] == forecast_hours:
    break
  weather_info.append([hour['startTime'][0:10], hour['startTime'][11:],  hour['relativeHumidity']['value'], hour['shortForecast'], hour['temperature'], hour['windDirection'], int(hour['windSpeed'].translate({ord(i): None for i in 'mph'})), hour['isDaytime'], hour['dewpoint']['value'], hour['probabilityOfPrecipitation']['value']])


print("Forecast for the next " + str(forecast_hours) + " specified hours, starting from start of the current hour")
weather_forecast_for_today = pd.DataFrame(data=weather_info, columns=['Date', 'Time', 'Humidity (%)', 'Forecast', 'Temperature (F)', 'Wind Direction', 'Wind Speed (mph)', 'Daytime', 'Dew Point (C)', 'Precipitation Chance (%)',])
weather_forecast_for_today.head(forecast_hours)




path = '%s_%s_%d_hour(s)_weather_data.xlsx' % (city_name, state_name, forecast_hours)
writer = pd.ExcelWriter(path)
weather_forecast_for_today.to_excel(writer, sheet_name='info', index=False, na_rep='NaN')

for column in weather_forecast_for_today:
    column_width = max(weather_forecast_for_today[column].astype(str).map(len).max(), len(column))
    col_idx = weather_forecast_for_today.columns.get_loc(column)
    writer.sheets['info'].set_column(col_idx, col_idx, column_width)

writer.close()
