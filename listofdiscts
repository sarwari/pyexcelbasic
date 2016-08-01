from openpyxl import load_workbook
from openpyxl import Workbook


def readrolefile(filename, activesheet='Sheet1', commandline=False):
    '''Read in the role file and return the workbook and active sheet 
       Returns:  wb ->  Workbook object
                 ws ->  Active worksheet in workbook
       If called from the command line, then print to stdout and exit, otherwise, raise          
    '''             
    try:
    	wb = load_workbook(filename=filename)
    	ws = wb[activesheet] 
    except:
        if commandline:
            print "Error loading worksheet - File not found or unavailable"
            sys.exit(1)
        else:
            raise
    return wb, ws

def getsheetdata(activesheet, addindex=False):    
    '''Get the data from the sheet.  
       Returns: cleanedheader -> Header in a list format 
                sheetdata -> Sheet data in a list of dicts, similar to csvDictReader
    '''
    cleanedheader = [ str(cell.value).strip() for cell in activesheet.rows[0] ]  
    #Create a list of dicts, similar to a csv.DictReader function
    sheetdata = [ dict(zip(cleanedheader, [str(cell.value) for cell in line ])) for line in activesheet.rows[1:] ]
    if addindex:
        cleanedheader.append('Original Row')
        for index, line in enumerate(sheetdata):
            line['Original Row'] = index + 1
    return cleanedheader, sheetdata
