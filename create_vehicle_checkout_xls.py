
import datetime
import calendar
from openpyxl import Workbook

#Writes a spreadsheet to your hard drive with the dates and times for "currentyear" in our vehicle checkout format.
#Import to Google Sheets for a good time!  you'll need pip install datetime calendar and openpyxl


#load months with hard data out of laziness, I don't think the months will change anytime soon though.
months = [
    'January', 'February', 'March', 'April', 'May', 'June', 'July',
    'August', 'September', 'October', 'November', 'December'
    ]
#This is the list of times, you can safely add values here and everything should still work
times = [
    '6:00', '6:30', '7:00', '7:30', '8:00', '8:30', '9:00', '9:30', '10:00',
    '10:30', '11:00', '11:30', '12:00', '12:30', '13:00', '13:30', '14:00',
    '14:30', '15:00', '15:30', '16:00', '16:30', '17:00', '17:30', '18:00',
    '18:30', 'Overnight'
    ]
#Starting row is usually 4 for a 3 row header.  Can change that and the year safely.
sheet_startingrow = 4
currentyear = 2019

#Function that returns an array of datetime objects in the month specified for each day
def loadDays(year, month):
    num_days = calendar.monthrange(year, month)[1]
    days = [datetime.date(year, month, day) for day in range(1, num_days+1)]
    return days
#Create a workbook in memory:
wb = Workbook()
#Loop through all months, using enumerate as we need to reference the mindex for the loaddays function call
for mindex, m in enumerate(months):
    #crete a worksheet with name = to current month
    wb.create_sheet(m)
    #grab sheet to make things easier to read
    current_sheet = wb[m]
    #load our days list using our function loaddays
    days_list = loadDays(currentyear, (mindex + 1))
    #we will need to track what row we are on during our loop through each day, reset to 0 for each month.
    row_counter = 0
    #now we are ready to loop through our days list.  again using enumerate since we need to track where we are during the loop.
    for dindex, d in enumerate(days_list):
            #we start on cell A4: row counter will increase to the next day, and we format using strftime
            current_sheet['A' + str(row_counter + sheet_startingrow)].value = days_list[dindex].strftime('%m/%d/%Y')
            #right below A4, A5 (4 + 1).  This is our day of the week.  rowcounter will save us
            current_sheet['A' + str(row_counter + sheet_startingrow + 1 )].value = calendar.day_name[days_list[dindex].weekday()]
            #just as with days, we need to track what time we're at in the time loop, reset each day
            tcounter = 0
            #time loop - gets the values into each time slot.  reference tcounter to offset correctly, and tindex for the loop
            for tindex, t in enumerate(times): 
                #this will run ~14 times per day, setting our times into their respective cells
                current_sheet['B' + str(row_counter + sheet_startingrow + tcounter + tindex)].value = t
            #increment by however many rows are in times but add one to get some white space
            row_counter += (1 + len(times))
#save everything
wb.save('C:/tmp/vehicle-checkout-'+str(currentyear)+'.xlsx')