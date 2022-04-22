import os, win32com.client, time
basepath = os.getcwd()
print(f'Корневой каталог: {basepath}\n')


Excel = win32com.client.Dispatch('Excel.Application')

while True:
	choice = input('Нажмите "1" чтобы поменять дату\nНажмите "2" чтобы найти и поменять значение\nНажмите "3" чтобы вывести меню и среднее значение\nНажмите "q" чтобы выйти   ')
	os.system('cls')
	if choice == '1':
		# поиск и замена всех дат на указанное значение
		date = input('введите дату: ')
		for root, dirs, files in os.walk(basepath):
			for file in files:
				if file.endswith('.xlsx') and not(file.startswith('ЦЕНЫ') or file.endswith('2021.xlsx') or file.startswith('Расчет')):
					try:
						wb = Excel.Workbooks.Open(os.path.join(root, file))
						sheet = wb.ActiveSheet
						old_date = sheet.Range('AL14').value
						sheet.Range('AL14').value = date
						print(f'"{file[:-5]}" дата изменена с {old_date}г. на {date}г.')
						wb.Save()
						wb.Close()
					except:
						print(f' не удалось открыть книгу {file}')
						continue			
		
	elif choice == '2':
		# поиск и замена случайного значения
		print('что-то')

	elif choice == '3':
		# вывод цены меню и среднего значения
		for i in range(9):
			wb = Excel.Workbooks.Open(os.path.join(basepath, str(i+1) + ' Меню Колледж2021.xlsx'))
			sheet = wb.ActiveSheet
			print(sheet.name)
			for j in range(10, 20):
			# тут ошибка разобраться
				if sheet.Cell(j, 11).value == 'итого:':
					print(f'{i+1} Меню Колледж2021:         цена: {sheet.Cell(j+1, 11).value}')	
			wb.Close()
		wb = Excel.Workbooks.Open(os.path.join(basepath, 'Расчет среднего.xlsx'))
		sheet = wb.ActiveSheet
		print(f"средняя цена: {sheet.Range('B14').value}")

	elif choice == 'q':
		Excel.Quit()
		exit()

	else:
		print('Вы нажали неверную клавишу!!')
		os.system('cls')

Excel.Quit()
