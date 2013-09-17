__author__ = 'Dave Dyer'
__version__ = '1.0'

import os
import xlrd
from matplotlib import pyplot as plt
import operator
import numpy as np
import datetime
from operator import itemgetter

local_path = "C:\\Users\\f2i3j07\\Documents\\dev\\py\\SEMA_monthlyMetrics\\CTACdats\\"

def get_files(path, file_count):
    """
    path = accepts a string 
    file_count = number of files you'd like to include in the plot
    """
    filenames = sorted(os.listdir(local_path))
    n = int(file_count)
    files_to_open = [local_path + y for y in filenames[-n:]]
    workbooks = [xlrd.open_workbook(x,'r') for x in files_to_open]
    return workbooks
    
def sheets_from_books(workbooks):
    """
    returns sheets and cell_values once you've input books
    
    """
    sheetlist = []

    for book in workbooks:
        sheetlist.append(book.sheet_by_name("Totals for Plot"))    
    return sheetlist
    
def sheets_to_dict(sheets):
    """
    Takes an xlrd workbook and returns the appropriate fields for monthly
    SEMA metric reporting:
    
    ---- inputs ---- 
    accepts an xlrd workbook 
    
    ---- outputs ----
    returns a dictionary that contains values
    """
    ndates, ntotal_iss, ntotal, ntotal_sw, ntotal_win,ntotal_unix, ntotal_watchlist_fw \
    ntotal_watchlist_bc, nevents_esc_sw, nevents_anal_sw, nevents_esc_internal, nevents_anal_internal,\
    nevents_esc_external, nevents_anal_external ,nevents_esc_internal, nevents_anal_internal= []
    for sheet in ss:
        ndates.extend(sheet.row_values(0)[1:])
        
    
books = get_files(local_path, 9)
sheets = sheets_from_books(books)
totals_dict = sheets_to_dict(sheets)
print 'datadict is ' + str(totals_dict['date'])

#fields = pull_fields(books[1])

#print "fields are " + str(fields)
            
