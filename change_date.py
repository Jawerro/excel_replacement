import os, win32com.client, time
basepath = os.getcwd()
print(f'Корневой каталог: {basepath}\n')
date = input('введите дату: ')

Exel = win32com.client.Dispatch('Excel.Application')

for root, dirs, files in os.walk(basepath):
	for file in files:
		if file.endswith('.xlsx') and not(file.startswith('ЦЕНЫ') or file.endswith('2021.xlsx') or file.startswith('Расчет')):
			try:
				wb = Exel.Workbooks.Open(os.path.join(root, file))
				sheet = wb.ActiveSheet
				old_date = sheet.Range('AL14').value
				sheet.Range('AL14').value = date
				print(f'"{file[:-5]}" дата изменена с {old_date}г. на {date}г.')
				wb.Save()
				wb.Close()
			except:
				print(f' не удалось открыть книгу {file}')
				continue

time.sleep(2)
Exel.Quit()
