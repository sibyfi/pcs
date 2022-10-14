#!/usr/bin/env python3

# Instructions:  Run the following commands download source code and install required packages.
# 
# git clone https://github.com/sibyfi/pcs
# pip3 install xlsxwriter
# pip3 install argparse
# 


from nls.classreport import *
from datetime import datetime
import csv
import xlsxwriter
import re
import sys
import argparse

# background true/false
print_console = False

# command line arguments
parser = argparse.ArgumentParser()
parser.add_argument('-i', '--input', help="Input file", required=True)
args = parser.parse_args()

# inputfile stuff
inputfile = args.input
file1     = open(inputfile, mode = 'r', newline = '')
reader    = csv.reader(file1)

# outputfile stuff
now        = datetime.now().strftime("%Y-%m-%d-%H%M%S")
outputfile = 'Class_Schedule_' + now + ".xlsx"

# call nls
formatcsv = BuildCsv(reader)


outWorkbook = xlsxwriter.Workbook(outputfile)
outSheet    = outWorkbook.add_worksheet()

# Format Definitions
header_format = outWorkbook.add_format()
header_format.set_bold()
header_format.set_align('center')
header_format.set_align('vcenter')
header_format.set_font_color('#FFFFFF')
header_format.set_bg_color('#0067C5')
header_format.set_border()
header_format.set_border_color('#0067C5')
header_format.set_text_wrap()

normal_format = outWorkbook.add_format()
normal_format.set_border()
normal_format.set_border_color('#0067C5')
# normal_format.set_align('center')
normal_format.set_align('vcenter')

enroll_format = outWorkbook.add_format()
enroll_format.set_border()
enroll_format.set_border_color('#0067C5')
enroll_format.set_align('center')
enroll_format.set_align('vcenter')
enroll_format.set_font_color('#0067C5')

date_format = outWorkbook.add_format()
date_format.set_border()
date_format.set_border_color('#0067C5')
date_format.set_align('center')
date_format.set_align('vcenter')

number_format = outWorkbook.add_format()
number_format.set_num_format('0')
number_format.set_border()
number_format.set_align('center')
number_format.set_align('vcenter')
number_format.set_border_color('#0067C5')

caution_format = outWorkbook.add_format()
caution_format.set_border()
caution_format.set_align('center')
caution_format.set_align('vcenter')
caution_format.set_bg_color('yellow')
caution_format.set_border_color('#0067C5')

alert_format = outWorkbook.add_format()
alert_format.set_border()
alert_format.set_align('center')
alert_format.set_align('vcenter')
alert_format.set_bg_color('red')
alert_format.set_font_color('white')
alert_format.set_border_color('#0067C5')

full_format = outWorkbook.add_format()
full_format.set_border()
full_format.set_align('center')
full_format.set_align('vcenter')
full_format.set_bg_color('green')
full_format.set_font_color('white')
full_format.set_border_color('#0067C5')



head_row = 1
head_col = "A"
for x in formatcsv.create_csv_header():
    outSheet.write(head_col + str(head_row), x, header_format)
    head_col = chr(ord(head_col) + 1)


data_row = 2
data_col = "A"
x = 0
len_data = formatcsv.len_data()
counter = 1
for bigdata in formatcsv.create_csv_data():
    course_name            = bigdata[0]
    course_number          = bigdata[1]
    offering_start_date    = bigdata[2]
    offering_end_date      = bigdata[3]
    course_duration        = bigdata[4]
    offering_location      = bigdata[5]
    offering_region        = bigdata[6]
    max_student_count      = bigdata[7]
    currently_enrolled     = bigdata[8]
    open_seats             = bigdata[9]
    customer_service_rep   = bigdata[10]
    offering_number        = bigdata[11]
    enroll                 = bigdata[12]
    catalog_domain_Name    = bigdata[13]
    offering_domain        = bigdata[14]
    display_for_learner    = bigdata[15]
    class_type             = bigdata[16]
    offering_status        = bigdata[17]
    content_version_number = bigdata[18]
    offering_instructor    = bigdata[19]
    check_alert            = CheckDatesAlert(offering_start_date)
    date_warning           = check_alert.alert()

    for write_row in bigdata:
        if x == 8:
            if currently_enrolled < 6 and date_warning >= 14 and date_warning < 31:
                outSheet.write_number(data_col + str(data_row), int(write_row), caution_format)
                data_col = chr(ord(data_col) + 1)
            elif currently_enrolled < 6 and date_warning <= 14:
                outSheet.write_number(data_col + str(data_row), int(write_row), alert_format)
                data_col = chr(ord(data_col) + 1)
            elif open_seats == 0:
                outSheet.write_number(data_col + str(data_row), int(write_row), full_format)
                data_col = chr(ord(data_col) + 1)
            else:
                outSheet.write_number(data_col + str(data_row), int(write_row), number_format)
                data_col = chr(ord(data_col) + 1)
        elif x == 4 or x == 7 or x == 9 or x == 15:
            outSheet.write_number(data_col + str(data_row), int(write_row), number_format)
            data_col = chr(ord(data_col) + 1)
        elif x == 18:
            outSheet.write(data_col + str(data_row), write_row, normal_format)
            data_col = chr(ord(data_col) + 1)
        elif x == 2 or x == 3:
            outSheet.write(data_col + str(data_row), write_row, date_format)
            data_col = chr(ord(data_col) + 1)
        elif x == 12:
            fix_url   = StripUrl(bigdata[12], offering_number)
            write_row = fix_url.fixurl()
            outSheet.write_url(data_col + str(data_row), str(write_row), enroll_format, string='Enroll')
            data_col = chr(ord(data_col) + 1)
        else:
            try:
                outSheet.write(data_col + str(data_row), str(write_row), normal_format)
                data_col = chr(ord(data_col) + 1)
            except:
                pass

        if x == 19:
            x = 0
        else:
            x += 1

    if print_console: print(str(counter) + " - " + str(offering_number))
    counter += 1
    data_col = "A"
    data_row += 1

# set column widths
outSheet.set_column('A:A', 57.25)
outSheet.set_column('B:B', 23.25)
outSheet.set_column('C:D', 12)
outSheet.set_column('E:E', 10)
outSheet.set_column('F:F', 55.75)
outSheet.set_column('G:G', 11)
outSheet.set_column('H:J', 10)
outSheet.set_column('K:K', 16.5)
outSheet.set_column('L:L', 11.75)
outSheet.set_column('M:M', 9.75)
outSheet.set_column('N:N', 11.33)
outSheet.set_column('O:O', 18.83)
outSheet.set_column('P:P', 11.83)
outSheet.set_column('Q:Q', 13)
outSheet.set_column('R:R', 11.5)
outSheet.set_column('S:S', 11.33)
outSheet.set_column('T:T', 19.5)

# ignore 'number stored as text' stupid warning
outSheet.ignore_errors({'number_stored_as_text': 'L2:L100000'})

# hide gridlines
outSheet.hide_gridlines(option=2)

# set global xlsx params
outSheet.autofilter('A1:T1')

# close workbook
outWorkbook.close()

# close files
file1.close()

if print_console: print("\nOutput file: " + outputfile)






