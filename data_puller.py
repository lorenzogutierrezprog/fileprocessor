import pandas as pd
import matplotlib.pyplot as plt
import matplotlib as mpl
from pathlib import Path
import matplotlib.patches as mpatches
from pandas.api.types import CategoricalDtype

pd.set_option('display.max_columns', None)
pd.set_option('display.max_rows', None)
here = str(Path.cwd())
temp_files = here + '\\Temp Files\\'

def data_pull(file):

    # define working directory
    # parsing first 6 lines of csv due to manufacturer formatting
    first_two_rows = pd.read_csv(file, nrows=1, sep = '[;,]', engine = 'python')
    # second_two_rows = pd.read_csv(file, nrows=1, skiprows=2)
    # third_two_rows = pd.read_csv(file, nrows=1, skiprows=4)

    # the bulk of data for charger use
    df = pd.read_csv(file, skiprows=6, sep = '[;,]', engine = 'python')

    # pulling serial number for charger identification
    try:
        serial_num = first_two_rows.iloc[-1]['S/N                '].replace(' ', '')
    except:
        serial_num = first_two_rows.iloc[-1]['S/N EEP            '].replace(' ', '')

    # convert Start of charge to time, adding day of week and time of day columns
    df['Charge Start'] = pd.to_datetime(df['SoC                '])
    df['Week Day'] = df['Charge Start'].dt.day_name()
    df['Day Num'] = df['Charge Start'].dt.dayofweek
    df['Time of Day'] = df['Charge Start'].dt.hour

    # defining charge profile as number for color scaling

    # legend handles
    opp = mpatches.Patch(color= 'blue', label= 'OPP')
    cold = mpatches.Patch(color= 'darkblue', label= 'COLD')
    ionic = mpatches.Patch(color= '#800080', label= 'IONIC')
    cmp_ch = mpatches.Patch(color= 'green', label= 'CMP CH')
    equal = mpatches.Patch(color= 'darkred', label= 'EQUAL')
    plt.legend(handles=[opp, cold, ionic, cmp_ch, equal], bbox_to_anchor=(1.05, 1.0), loc='upper left')

    # print(profile_names)
    color_dictionary = {'OPP    ': 1, 'COLD   ': 2, 'IONIC  ': 3, 'CMP CH ': 4, 'EQUAL  ': 5}
    color_map = mpl.colors.ListedColormap(['blue', 'darkblue', '#800080', 'green', 'darkred'])
    df['Color Code'] = df['Profile'].replace(color_dictionary)
    # norm = mpl.colors.Normalize(vmin=1, vmax=5)

    # plot time, day, weighted charge
    plt.scatter(df['Day Num'], df['Time of Day'], s = df['Chg Tim'] * 1.6, c = df['Color Code'], alpha = .6,
                cmap = color_map)
    plt.clim(vmin=1, vmax=5)

    plt.xticks(ticks = [0, 1, 2, 3, 4, 5, 6], labels = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday',
                                                        'Saturday', 'Sunday'], rotation=15)
    plt.ylabel('Time of Day')
    # plt.show()


    # filter to relevant data, sort longest charge times
    df = df.filter(items=['Week Day', 'Time of Day', 'Chg Tim', 'Profile', 'Charge Start'])\
        .sort_values('Chg Tim', ascending=False).set_index('Chg Tim')

    # import battery locations, name, and serial correlation
    batteries = pd.read_excel('Batterylocations.xlsx')

    row = batteries.loc[batteries['Serial Number'] == serial_num]
    row = row[['Name', 'Location', 'Serial Number']]

    # Title from battery locations file
    try:
        name = str(row.iloc[0, 0])
        location = str(row.iloc[0, 1])
        serial = str(row.iloc[0, 2])
        plt.title(name + ' ' + location + ' ' + serial)
        plt.tight_layout()
        fig = plt.savefig(temp_files + name + '.png')
    except:
        name = "error"
        location = "error"
        serial = "error"
        plt.title("Identification Error")
        plt.tight_layout()
        fig = plt.savefig(temp_files + "IDError" + '.png')


    # save_path = Path('C:\\Users\\lgutierrez.MYSMITHFIELD\\PycharmProjects\\Chargerdataprocessor\\Temp Files')


    plt.clf()

    # Number of equalizes
    try:
        EQs = df['Profile'].value_counts()['EQUAL  ']
    except:
        EQs = 0

    # Data date range
    time_dif = (max(df['Charge Start']) - min(df['Charge Start']))
    print('Data pulled.')

    return name, location, serial, fig, df, time_dif, EQs
