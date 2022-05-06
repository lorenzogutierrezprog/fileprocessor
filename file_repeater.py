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

# import charger locations, name, and serial correlation
chargers = pd.read_excel('Chargerlocations.xlsx')


# create excel to be written to
writer = pd.ExcelWriter('Charger Data ' + current_time + '.xlsx', engine='xlsxwriter')
workbook = writer.book

for name in list(chargers['Name']):
        print(name)
        worksheet = workbook.add_worksheet(name)
writer.save()

writer = pd.ExcelWriter('Charger Data ' + current_time + '.xlsx', engine='xlsxwriter')
workbook = writer.book
# for each file: process, print data and plot in Excel sheet
for file in input_folder:
        print('')
        print('')
        print('Processing ' + str(file)[17:])
        run = data_pull(file)
        name = run[0]
        location = run[1]
        serial = run[2]
        fig = run[3]
        df = run[4]
        time_dif = run[5]
        EQs = run[6]
        DCs= run[7]
        DC_percent = run[8]

        # push dataframe to excel, format, add plot and text
        if name == "Error":
                df.to_excel(writer, sheet_name=str(file)[17:-4])
                workbook = writer.book
                worksheet = writer.sheets[str(file)[17:-4]]
        else:
                df.to_excel(writer, sheet_name=name)
                workbook = writer.book
                worksheet = writer.sheets[name]

        header = workbook.add_format({'bold': True})
        header.set_align('right')
        cell_format = workbook.add_format()
        cell_format.set_align('right')
        worksheet.set_column('A:D', 10)
        worksheet.set_column('E:F', 20)
        worksheet.write('F2', 'Days Recorded', header)
        worksheet.write('F3', time_dif)
        worksheet.write('F5', 'Equalizations', header)
        worksheet.write('F6', EQs)
        worksheet.write('F8', 'EQs Per Week', header)
        worksheet.write_formula('F9', '=F6 * 7 / F3')
        if DC_percent ==0:
                pass
        else:
                worksheet.write('F11', "Invalid Disconnects", header)
                worksheet.write('F12', DCs, cell_format)
                worksheet.write('F13', str(round(DC_percent)) + "%", cell_format)


        if name == "Error":
                figpath = temp_files + "Error " + str(file)[17:-4] + '.png'
                worksheet.insert_image('H1', figpath)
                print("Found error in file " + str(file)[17:-4] + " but completed")
        else:
                figpath = temp_files + name + '.png'
                worksheet.insert_image('H1', figpath)
                print(name + ' ' + serial + ' ' + location + ' - Success')


# summary page
workbook = writer.book
worksheet = workbook.add_worksheet('Images Summary')
row = 0
for img in Path(temp_files).iterdir():
        worksheet.insert_image(row, 0, img)
        row += 23

writer.save()

# TODO delete png in Temp Files
# TODO move processed files to folder
