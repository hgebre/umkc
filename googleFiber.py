#!/usr/bin/python
# This script takes in xls files
# and converts it to csv and take three columns from the csv file
# to send request to google fiber website and read the result and store 
# in dictionary by category and write it out to a file
# by HGA 2015


import os
import os.path
import sys 
import xlrd
import csv
import re
import urllib
import urllib2
import pandas
from collections import defaultdict

bln_append_filename = 1

if len(sys.argv) == 1:
    print "Please input xls - file name to covert to csv"
    print "usage: python googleFiber.py (required inputfilenamd) (optional outputfilename)"
    sys.exit(0)

infile = sys.argv[1]


def check_ext(input):
    """
    #**********************************************************************************************
    # This taks in input path and returns the extension of the file
    #**********************************************************************************************
    """
    extension = os.path.splitext(input)[1]
    return extension


def addDict(cat):
    """
    #***********************************************************************************************
    # takes in the category checks to see if it exists, if not adds it
    # if it exists increment the value by 1
    #***********************************************************************************************
    """
    global catDict
    if catDict.has_key(cat) == 1:
        if( cat > ''):
            catDict[cat] += 1
    else:
        if( cat > ''):
            catDict[cat]= 1
    return catDict


def onTracking():
    """
    #***********************************************************************************************
    # this function turns tracking on
    #***********************************************************************************************
    """
    global track
    track = 1


def offTracking():
    """
    #***********************************************************************************************
    # this function resets tracking
    #***********************************************************************************************
    """

    global track
    track = 0


"""
# remove extention of the input files for output name
"""
output = infile.replace(check_ext(infile),'')
output = output.replace(" ","")

detailfile = open(output+'_detail.csv','w')
sumfile = open(output+'_summary.txt','w')
errorfile = open(output+'_error.txt','w')
csvfile = output + ".csv"
 
fp_out = open(csvfile, 'w')
fp_out.truncate()
csv_writer_out = csv.writer(fp_out)

date_out = re.sub( r'^(\d\d)(\d\d)(\d\d).*', r'20\3-\2-\1', infile)

book = xlrd.open_workbook(infile)
sheet = book.sheet_by_index(0)
for row_index in xrange(sheet.nrows):
        excel_row = sheet.row_values(row_index)
        if bln_append_filename:
                excel_row.append(infile)
                excel_row.append( date_out)
        csv_writer_out.writerow(excel_row)
track = 0
fp_out.close()

catDict = {}


url = 'https://fiber.google.com/cities/kansascity/'

from collections import defaultdict
columns = defaultdict(list)
with open(csvfile) as f:
    reader = csv.reader(f)
    reader.next()
    s = 0
    tmpDict = {}
    for row in reader:
        for(col,val) in enumerate(row):
            columns[col].append(val)
            if(col==1):
	        address = val.rstrip(' ') if ' ' in val else val
            elif (col == 1):
	        unit = val.rstrip(' ') if ' ' in val else val
            elif (col == 3):
	        zip = val.rstrip('0').rstrip('.') if '.' in val else val 
	        
	data = {}
        data['street_address'] = address 
	data['unit_number'] = unit
	data['zip_code'] = zip
        url_values = urllib.urlencode(data)
        offTracking()  
        #print url_values
        full_url = url + '?' + url_values
        try:
            data = urllib2.urlopen(full_url).readlines()
            keystr = '<div class="status-icon'
            found = re.compile(keystr, re.IGNORECASE)
            for j in range(len(data)):
                start = data[j].find('<div class="status-icon') + 24
                end = data[j].find('ng-class="">')
                if (found.search(data[j])):
                    pattern = str(data[j][start:end]).strip('=')
                    pattern = re.sub('"','',pattern).strip()
                    tmpDict = addDict(pattern)
                    #print tmpDict
                if(track == 0 and found.search(data[j])):
                    detailfile.write(address + ',' + zip + ',' + re.sub('"','',str(data[j][start:end]).strip('=')) + '\n')
                    s +=1
                    onTracking()   
        except urllib2.HTTPError, error:
            errmsg = error.read()
            errorfile.write(errmsg)
            pass
 
        if(track != 1 ):
            detailfile.write(address + ', ' + zip + '->****Incomplete Request****\n')
            s +=1
            tmpDict = addDict('incomplete-request')
    sumfile.write('File Name : ' + infile +' \n')
    sumfile.write('Total addresses in file : ' + str(s) +' \n')
    sumfile.write('--------BREAKDOWN BY CATEGORY ----------------'+' \n') 
    if (s > 0):
        for key, val in catDict.items():
            sumfile.write(str(key) + ': ' + str(val) + ' \n')
#     print a + ', ' + z + '-> Address IS'
detailfile.close()
sumfile.close()
errorfile.close()
