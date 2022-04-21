import datetime
from pathlib import Path
import pandas as pd
from data_puller import data_pull
import os
import xlsxwriter

# defining paths
here = str(Path.cwd())
path = Path('files to process/')
input_folder = path.iterdir()
current_time = datetime.datetime.now().strftime('%m_%d_%y')
temp_files = here + '\\Temp Files\\'

# import battery locations, name, and serial correlation
batteries = pd.read_excel('Batterylocations.xlsx')

# for i in range(len(batteries)):
#         name = batteries.iloc[i,1]
#         print(name)
#         worksheet = workbook.add_worksheet(name)
# writer.save()

# create excel to be written to
writer = pd.ExcelWriter('Charger Data ' + current_time + '.xlsx', engine='xlsxwriter')
workbook = writer.book

# print(list(batteries['Name']))
for name in list(batteries['Name']):
        print(name)
        worksheet = workbook.add_worksheet(name)
writer.save()

writer = pd.ExcelWriter('Charger Data ' + current_time + '.xlsx', engine='xlsxwriter')
workbook = writer.book
# for each file: process, print data and plot in Excel sheet
for file in input_folder:
        print('')
        print('')
        print('Processing ' + str(file))
        run = data_pull(file)
        name = run[0]
        location = run[1]
        serial = run[2]
        fig = run[3]
        df = run[4]
        time_dif = run[5]
        EQs = run[6]

        # push dataframe to excel, format, add plot and text
        df.to_excel(writer, sheet_name=name)
        workbook = writer.book
        worksheet = writer.sheets[name]
        header = workbook.add_format({'bold': True})
        header.set_align('right')
        worksheet.set_column('A:D', 10)
        worksheet.set_column('E:F', 20)
        worksheet.write('F2', 'Days Recorded', header)
        worksheet.write('F3', time_dif)
        worksheet.write('F5', 'Equalizations', header)
        worksheet.write('F6', EQs)
        worksheet.write('F8', 'EQs Per Week', header)
        worksheet.write_formula('F9', '=F6 * 7 / F3')

        figpath = temp_files + name + '.png'
        worksheet.insert_image('H1', figpath)
        print(name + ' ' + serial + ' ' + location + ' success')
# summary page
workbook = writer.book
worksheet = workbook.add_worksheet('Images Summary')
row = 0
for img in Path(temp_files).iterdir():
        worksheet.insert_image(row, 0, img)
        row += 23

writer.save()

#delete all png in Temp Files