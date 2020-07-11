import requests
from time import sleep
import xlwings as xw
import pywintypes
import tkinter  as tk
from tkinter import messagebox

def get_data(city):
	res=requests.get('http://api.openweathermap.org/data/2.5/weather?q='+city+'&appid=03001eadc2c46c7ccf002a53f48c0b15')
	response = res.json()
	data={
		'city' : {
			'id' : response['weather'][0]['id'],
			'name' : response['name'],
			'lat' : response['coord']['lat'],
			'lon' : response['coord']['lon'],
		},
		'temp' : {
			'K' : response['main']['temp'],
			'C' : round(response['main']['temp'] - 273.15, 2),
			'F' : round((( response['main']['temp'] - 273.15) * 9/5) + 32, 2)
		}
	}
	return(data)


if __name__=='__main__':
	wb = xw.Book('weather.xlsx')
	sheet1 = wb.sheets['Sheet1']
	sheet2 = wb.sheets['Sheet2']
	total_cities = wb.sheets[0].range('A' + str(wb.sheets[0].cells.last_cell.row)).end('up').row
	while True:
		for i in range(2, total_cities + 2):
			try:
				city = sheet1.range('A' + str(i)).value
				data = get_data(city)
				sheet1.range('B'+str(i)).value = data['temp'][sheet1.range('C'+str(i)).value]
				sheet2.range('A'+str(i)).value = data['city']['id']
				sheet2.range('B'+str(i)).value = data['city']['name']
				sheet2.range('C'+str(i)).value = data['city']['lat']
				sheet2.range('D'+str(i)).value = data['city']['lon']
			except KeyError:
				sheet1.range('B'+str(i)).value = 'Not found'
				sheet2.range('B'+str(i)).value = city
				sheet2.range('C'+str(i)).value = 'No data found'
			except TypeError:
				sheet1.range('B'+str(i)).value = 'no city name'
			except pywintypes.com_error:
				root = tk.Tk()
				root.withdraw()
				messagebox.showerror(title='Error', message='Application is closed.\nDo not edit while the application is running.\nRun the application again after editing and saving the sheet.')
				quit()