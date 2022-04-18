import pandas as pd
import matplotlib.pyplot as plt
import matplotlib as mpl
from pathlib import Path
# import seaborn as sns
import matplotlib.patches as mpatches
from pandas.api.types import CategoricalDtype

pd.set_option('display.max_columns', None)
pd.set_option('display.max_rows', None)
here = str(Path.cwd())
temp_files = here + '\\Temp Files\\'

def data_pull(file):

    # define working directory
    # parsing first 6 lines of csv due to manufacturer formatting
    first_two_rows = pd.read_csv(file, nrows=1)
    # second_two_rows = pd.read_csv(file, nrows=1, skiprows=2)
    # third_two_rows = pd.read_csv(file, nrows=1, skiprows=4)

    # the bulk of data for charger use
    df = pd.read_csv(file, skiprows=6)

    # pulling serial number for charger identification
    serial_num = first_two_rows.iloc[-1]['S/N                '].replace(' ', '')

    # convert Start of charge to time, adding day of week and time of day columns
    # id = pd.Index(, ordered=True)
    df['Charge Start'] = pd.to_datetime(df['SoC                '])
    df['Week Day'] = df['Charge Start'].dt.day_name()
    df['Day Num'] = df['Charge Start'].dt.dayofweek
    df['Time of Day'] = df['Charge Start'].dt.hour

    # defining charge profile as number for color scaling
    color_dictionary2 = {'OPP    ': 'blue',  'COLD   ' :'darkblue','CMP CH ': 'green', 'EQUAL  ': 'darkred'}
    color_steps = 1, 2, 3, 4

    # legend handles
    opp = mpatches.Patch(color='blue', label='OPP')
    cold = mpatches.Patch(color='darkblue', label='COLD')
    cmp_ch = mpatches.Patch(color='green', label='CMP CH')
    equal = mpatches.Patch(color='darkred', label='EQUAL')

    # print(profile_names)
    color_dictionary = {'OPP    ': 1, 'COLD   ': 2, 'CMP CH ': 3, 'EQUAL  ': 4}
    color = df['Profile'].replace(color_dictionary)
    color_map = mpl.colors.ListedColormap(['blue','darkblue', 'green', 'darkred'])


    # plot time, day, weighted charge
    scatter = plt.scatter(df['Day Num'], df['Time of Day'], s=df['Chg Tim'] * 1.6, c=color, alpha=.6, cmap=color_map)
    plt.legend(handles=[opp, cold, cmp_ch, equal],bbox_to_anchor=(1.05, 1.0), loc='upper left')

    plt.xticks(ticks=[0, 1, 2, 3, 4, 5, 6], labels=['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'], rotation=15)
    plt.ylabel('Time of Day')
    # plt.show()


    # filter to relevant data, sort longest charge times
    df = df.filter(items=['Week Day', 'Time of Day', 'Chg Tim', 'Profile', 'Charge Start'])\
        .sort_values('Chg Tim', ascending=False).set_index('Chg Tim')

    # import battery locations, name, and serial correlation
    batteries = pd.read_excel('Batterylocations.xlsx')

    row = batteries.loc[batteries['Serial Number'] == serial_num]
    row = row[['Name', 'Location', 'Serial Number']]

    name = str(row.iloc[0, 0])
    location = str(row.iloc[0, 1])
    serial = str(row.iloc[0, 2])

    # title and extract figure
    plt.title(name + ' ' + location + ' ' + serial)
    plt.tight_layout()
    # save_path = Path('C:\\Users\\lgutierrez.MYSMITHFIELD\\PycharmProjects\\Chargerdataprocessor\\Temp Files')
    fig = plt.savefig(temp_files + name + '.png')

    plt.clf()

    # Number of equalizes
    if df.nunique()['Profile'] == 1:
        EQs = 0
    else:
        EQs = df['Profile'].value_counts()['EQUAL  ']

    # Data date range
    time_dif = (max(df['Charge Start']) - min(df['Charge Start']))
    print('.')

    return name, location, serial, fig, df, time_dif, EQs
