import sys
print("Python executable:", sys.executable)
print("Python version:", sys.version)
print("\nSys.path (where Python looks for modules):")
for path in sys.path:
    print(path)

import xlwings as xlw
from tkinter import Tk
from tkinter.filedialog import askopenfilename
from tkinter.filedialog import asksaveasfilename
from collections import defaultdict
import pandas as pd
from copy import deepcopy
import sys

# To Use:
# Download the Block Office Hours Schedule as a .xlsx file
# Run this program
# Input the Block Office Hours Schedule .xlsx into this program
# Wait
# Select the Save File
# Done!

# Notes:
# The Block Office Hours Schedule must have the sheets
# 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Sunday', and 'Name List'
# Otherwise, the program will crash
# Additionally, ThreeSheet expects 3 PTs and FiveSheet expects 5 PTs per timerange, and that it is blocked row-by-row, starting at row 2
# with the first column being the timeranges
# It also expects Sunday to be blocked Column-by-column and start at row 6, with row 5 being the timeranges

def ThreeSheet(sheet, OHTimes):
    prevNames = []
    # goes through every timerange in the sheet
    # +1+1 for final time range (4:30-5:00)
    for row in range(2, len(sheet.used_range.rows)+1+1):
        names = sheet['B'+str(row)+':D'+str(row)]
        for name in names.value:
            # bad names
            if (name != 'Admin Block' and name is not None):
                # this is for consolidating times
                if (name not in prevNames):
                    OHTimes[name][sheet.name] += [sheet['A'+str(row)].value.split()[0]]
                OHTimes[name]['hours'] += 0.5
        
        for name in prevNames:
            # bad names
            if (name == 'Admin Block' or name is None):
                continue
            # end of time consolidation
            if (name not in names.value):
                startTime = OHTimes[name][sheet.name][-1]
                endTime = sheet['A'+str(row-1)].value.split()[2]
                startHour = int(startTime.split(':')[0])
                endHour = int(endTime.split(':')[0])
                # for am/pm splits
                if ((startHour >= 8 and startHour < 12) and (endHour >= 8 and endHour < 12)):
                    OHTimes[name][sheet.name][-1] = [startTime+' - '+endTime+' am']
                elif ((startHour >= 8 and startHour < 12) and (endHour >= 12 or endHour <= 5)):
                    OHTimes[name][sheet.name][-1] = [startTime+' am - '+endTime+' pm']
                else:
                    OHTimes[name][sheet.name][-1] = [startTime+' - '+endTime+' pm']
        prevNames = names.value

def FiveSheet(sheet, OHTimes):
    prevNames = []
    # goes through every timerange in the sheet
    # +1+1 for final time range (4:30-5:00)
    for row in range(2, len(sheet.used_range.rows)+1+1):
        # 5 PTs, so increase range to B:F
        names = sheet['B'+str(row)+':F'+str(row)]
        for name in names.value:
            # bad names
            if (name != 'Admin Block' and name is not None):
                # this is for consolidating times
                if (name not in prevNames):
                    OHTimes[name][sheet.name] += [sheet['A'+str(row)].value.split()[0]]
                OHTimes[name]['hours'] += 0.5
        
        for name in prevNames:
            # bad names
            if (name == 'Admin Block' or name is None):
                continue
            # end of time consolidation
            if (name not in names.value):
                startTime = OHTimes[name][sheet.name][-1]
                endTime = sheet['A'+str(row-1)].value.split()[2]
                startHour = int(startTime.split(':')[0])
                endHour = int(endTime.split(':')[0])
                # for am/pm splits
                if ((startHour >= 8 and startHour < 12) and (endHour >= 8 and endHour < 12)):
                    OHTimes[name][sheet.name][-1] = [startTime+' - '+endTime+' am']
                elif ((startHour >= 8 and startHour < 12) and (endHour >= 12 or endHour <= 5)):
                    OHTimes[name][sheet.name][-1] = [startTime+' am - '+endTime+' pm']
                else:
                    OHTimes[name][sheet.name][-1] = [startTime+' - '+endTime+' pm']
        prevNames = names.value

def OHTimesFactory():
    d = defaultdict(list)
    d['hours'] = 0
    return d

print('Select the Office Hours Excel File (Block Office Hours Spreadsheet)')
Tk().withdraw()
officeHoursFile = askopenfilename(filetypes=[('Excel Files', '*.xlsx')], title='Office Hours Spreadsheet')
if (officeHoursFile == '' or officeHoursFile is None):
    print("No file selected")
    sys.exit()
print('Selected', officeHoursFile)

with xlw.App(visible=False) as APP:
    with xlw.Book(officeHoursFile) as officeHours:
        NAMES = officeHours.sheets['Name List']
        MONDAY = officeHours.sheets['Monday']
        TUESDAY = officeHours.sheets['Tuesday']
        WEDNESDAY = officeHours.sheets['Wednesday']
        THURSDAY = officeHours.sheets['Thursday']
        FRIDAY = officeHours.sheets['Friday']
        SUNDAY = officeHours.sheets['Sunday']
        ThreeDates = [MONDAY, TUESDAY, WEDNESDAY]
        FiveDates = [THURSDAY, FRIDAY]
        # Start from C2
        nameList = []

        OHTimes = defaultdict(OHTimesFactory)
        for end in range(2, len(NAMES.used_range.rows)+1):
            if (NAMES['A'+str(end)].value is None) or (NAMES['A'+str(end)].value == 'Admin'):
                continue
            nameList += [NAMES['A'+str(end)].value + " " + NAMES['B'+str(end)].value]
            name = NAMES['A'+str(end)].value + " " + NAMES['B'+str(end)].value
            # this line is here to ensure that the csv exports
            # with the correct name order
            OHTimes[name]
            OHTimes[name]['hours'] = 0
        print("Found names")
        
        # abuse Python's weird pass-by-id-or-smth to imitate a C++ pass-by-reference
        # to update OHTimes using a function and not a copy paste of code
        ThreeSheet(MONDAY, OHTimes)
        ThreeSheet(TUESDAY, OHTimes)
        ThreeSheet(WEDNESDAY, OHTimes)
        FiveSheet(THURSDAY, OHTimes)
        FiveSheet(FRIDAY, OHTimes)
                    
        # similar to the other 5 days
        # except the times are columns and not rows
        # however, each block is 1 hour and cannot overlap, so we can ignore column by column checking
        # and just use a 2 column range
        prevNames = []
        for row in SUNDAY['A6:B30'].value:
            # row is a list of two values
            if (row[0] is not None):
                name = row[0]
            else:
                name = row[1]
            if (name is not None):
                OHTimes[name]['Sunday'] += ['2:00 - 3:00 pm']
                prevNames += [name]
                OHTimes[name]['hours'] += 1
                
        currNames = []
        for row in SUNDAY['C6:D30'].value:
            if (row[0] is not None):
                name = row[0]
            else:
                name = row[1]
            if (name is not None):
                if (name not in prevNames):
                    OHTimes[name]['Sunday'] += ['3:00 - 4:00 pm']
                currNames += [name]
                OHTimes[name]['hours'] += 1
                
        for name in prevNames:
            if (name not in currNames):
                OHTimes[name]['Sunday'][-1] = ['2:00 - 3:00 pm']
        # python always does what is basically a shallow copy with default =
        # so make deepcopy here because currNames will be overwritten
        prevNames = deepcopy(currNames)
        currNames = []
        for row in SUNDAY['E6:F30'].value:
            if (row[0] is not None):
                name = row[0]
            else:
                name = row[1]
            if (name is not None):
                if (name not in prevNames):
                    # nested list because Python will just add the string to the list instead of a list of strings
                    # this is for the string conversion later
                    # it is only here becuase this is the only time where the string is not overwritten by a list later
                    OHTimes[name]['Sunday'] += [['4:00 - 5:00 pm']]
                currNames += [name]
                OHTimes[name]['hours'] += 1
                
        for name in prevNames:
            startTime = OHTimes[name]['Sunday'][-1].split()[0]
            if (name not in currNames):
                OHTimes[name]['Sunday'][-1] = [startTime+' - 4:00 pm']
            else:
                OHTimes[name]['Sunday'][-1] = [startTime+' - 5:00 pm']
        
        print("Found Office Hours")
        # print(OHTimes)
        # for key in OHTimes:
        #     print(key)
        #     for day in OHTimes[key]:
        #         print(day,end=': ')
        #         print(OHTimes[key][day])

# changing the values to a more human readable format
# and flattens the day lists to a single combined string
OHTimesStrings = defaultdict(lambda: str())
for name in OHTimes:
    for day in OHTimes[name]:
        if day == 'hours':
            continue
        for timeRange in OHTimes[name][day]:
            if (len(timeRange) > 0):
                if (day != 'Thursday'):
                    OHTimesStrings[name] += day[0]+' '+timeRange[0]+', '
                else:
                    OHTimesStrings[name] += 'R '+timeRange[0]+', '
    # delete the last ', '
    OHTimesStrings[name] = [OHTimesStrings[name][:-2], OHTimes[name]['hours']]
print("Converted Office Hours to Readable Format")
# print(OHTimesStrings)
for name in OHTimesStrings:
    print(name,end=': ')
    print(OHTimesStrings[name])

print("Select the Save File")
saveFile = asksaveasfilename(defaultextension='.csv', filetypes=[("CSV File","*.csv")], title='Save File')
if (saveFile == '' or saveFile is None):
    print("File not saved")
    sys.exit()
# one-liner to convert dict to a csv
pd.DataFrame.from_dict(data=OHTimesStrings, orient='index').to_csv(saveFile, header=False)
print('Saved in', saveFile)