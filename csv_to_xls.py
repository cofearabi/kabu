# -*- coding: utf-8 -*-
import jsm
import sys, csv
import pyExcelerator
# About csv module, see "http://www.python.jp/doc/2.4/lib/module-csv.html"
# pyExcelerator can be redistributed with BSD Lisence and all rights reserved Roman V. Kiseliov
# The project page of pyExcelerator is "http://sourceforge.net/projects/pyexcelerator/"

def convert( csv_fobj, excel_file_name ):
    """
    convert( csv_fobj, excel_file_name )
        Converts a csv file into MS-Excel file.
        First argument: csv_fobj is a file object which reference to csv file.
        Second argument: excel_file_name is a file name.
    """
    record_from_csv = csv.reader( csv_fobj )
    
    excel_workbook  = pyExcelerator.Workbook()
    excel_sheets    = [sheet for sheet in map( excel_workbook.add_sheet, [u"sheet" + str( index ) for index in range( 3 )] )]
    
    (row, column) = (0, 0)
    for fields in record_from_csv:
        for cell_val in fields:
            excel_sheets[0].write( row, column, label=cell_val )
            column += 1
        column = 0
        row   += 1
    
    excel_workbook.save( excel_file_name )


if __name__ == "__main__":
    if( len(sys.argv) < 2 ) or ( sys.argv[1] in ("-h", "/h", "-?", "/?", "--help") ) or ( len( sys.argv ) < 2 ):
        print "Usage: python csv_to_excel.py ccode (4689)"
    else:
        c = jsm.QuotesCsv()
        c.save_historical_prices( sys.argv[1] + ".csv",sys.argv[1])
#   csv_fobj = open( sys.argv[1], "rb" )
        csv_fobj = open( sys.argv[1] + ".csv" , "rb" )
        convert( csv_fobj, sys.argv[1] + ".xls" )

# on BSD License.
# See "http://www.opensource.org/licenses/bsd-license.php"
#
# Finally I need your constructive feedbacks. Thank you!
#
# Copyright (C) 2006 Fomalhaut Weisszwerg
# mail to: weisszwerg@gmail.com
# All rights reserved.
