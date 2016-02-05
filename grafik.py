#!/usr/bin/python

#Usage:
#./grafik.py 11/11/2015
#./grafik.py



import xlsxwriter
from datetime import datetime, date, timedelta
import calendar
import sys

def get_month_range( start_date = None ):
    if start_date is None:
        start_date = date.today().replace( day = 1 )
    _, days_in_month = calendar.monthrange( start_date.year, start_date.month )
    end_date = start_date + timedelta( days = days_in_month )
    return ( start_date, end_date )

workbook = xlsxwriter.Workbook( 'grafik.xlsx' )
worksheet = workbook.add_worksheet()
highlight_free = workbook.add_format( {'bg_color': '#FFFF66'} )
highlight_working = workbook.add_format( {'bg_color': '#66FF66'} )
highlight_header = workbook.add_format( {'bg_color': '#FFCCFF', 'bold': True} )

worksheet.write( 'A1', 'Data', highlight_header )
worksheet.write( 'B1', 'Godziny', highlight_header )
worksheet.set_column( 'A:A', 10 )

a_day = timedelta( days = 1 )
day, last_day = get_month_range()

if len(sys.argv) >= 2:
    holiday = sys.argv
else:
    holiday = None

counter = 1
last_empty_day = 0
my_hours = 0
working_days_list = []

while day < last_day:
    weekday = day.weekday()
    date_string = day.strftime( "%d/%m/%Y" )

    today_is_holiday = False
    if holiday is not None:
        if day.strftime( "%d/%m/%Y" ) in holiday:
                today_is_holiday = True

    if ( 5 <= weekday <= 6 ) or today_is_holiday:
        worksheet.write( counter, 0, date_string, highlight_free )
        worksheet.write( counter, 1, 0, highlight_free )
    else:
        working_days_list.append( counter )
        if( 4 == weekday ):
            worksheet.write( counter, 0, date_string, highlight_working )
            worksheet.write( counter, 1, 0, highlight_free )
            last_empty_day = counter
        else:
            worksheet.write( counter, 0, date_string, highlight_working )
            worksheet.write( counter, 1, 8, highlight_working )
            my_hours += 8

    day += a_day
    counter += 1

hours_to_fill = round( len( working_days_list ) * 8 * 4 / 5, 2 )
missing = round( hours_to_fill - my_hours, 2 )
print( hours_to_fill, missing )
worksheet.write( counter, 0, 'suma:', highlight_header )
worksheet.write( counter, 1, hours_to_fill, highlight_header )
if missing > 0:
        worksheet.write( last_empty_day, 1, missing, highlight_working )
if missing < 0:
        missing = abs( missing )
        offset = int( missing / 8 )
        i = 0
        working_days_list.reverse()
        while i < offset:
            worksheet.write( working_days[ i ], 1, 0, highlight_free )
            i += 1
        leftovers = 8 - ( missing % 8 )
        if leftovers < 4:
            worksheet.write( working_days_list[ i + 1 ], 1, leftovers + 4, highlight_working )
            worksheet.write( working_days_list[ i ], 1, 4, highlight_working )
        else:
            worksheet.write( working_days_list[ i ], 1, leftovers, highlight_working )


workbook.close()

