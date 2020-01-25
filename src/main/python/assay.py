import pandas as pd
import ntpath
import xlwings as xw
import numpy as np
from datetime import datetime
from error_handling import RangeNotFoundError
import time


class Assay:
    """ Class containing all details from F007:
        Date, technician, sponsor, study, samples and plate IDs """

    def __init__(self, f007, qc_file, curve_file, xl_id, files=[]):
        """ f007 and files as full paths from file picker dialog """

        self.f007 = f007  # File path
        self.f007_ref = get_file_from_path(self.f007)  # Standard F007 reference ID

        self.files = files  # List of elisa plate file paths
        self.qc_file = qc_file  # File containing QC ranges
        self.curve_file = curve_file  # File containing Curve Info
        self.xl_id = xl_id  # PID of Excel application
        self.wb = None  # Assay workbook

        # Get assay details
        self.tech, self.date, self.sponsor, self.study = self.get_assay_details()

        # List of first run and repeat plates and samples
        self.first_list, self.repeats_list = self.get_samplelist()
        
        # Determine whether first run, mixed or repeats
        self.run_type = self.get_run_type()

        self.qc_limits = self.get_qc_limits()  # Get QC limits
        self.curve_vals = self.get_curve_vals()  # Get IgG curve concentrations
        
    def get_assay_details(self):
        """ Get technician, date, sponsor and study details"""

        app = self.get_xl_app()  # Get working Excel process by pID
        self.wb = app.books.open(self.f007)  # Open F007 workbook

        # Defined range names to check
        details = {'AnalystName': '', 'AssayStart': '', 'Sponsor': '', 'StudyName': ''}
        
        # getting the details if range exists in workbook
        for d in details.keys():
            try:
                rng = self.wb.names(d).refers_to_range
            except:
                raise RangeNotFoundError("F007 Range (" + d + ") Not Found")
                return
                
            # Check that ranges contain a value - if not, return all as none
            val = rng.value
            if val:
                details[d] = val
            else:
                raise RangeNotFoundError("Range in F007 is empty: " + d)
                wb.close()
                return
                
        # Check technician (Empty, upper case, space)
        details["AnalystName"] = get_tech_initials(details["AnalystName"])
    
        # Check and re-format Assay Date
        details["AssayStart"] = change_assaydate(details["AssayStart"])

        # Get final details
        tech = details['AnalystName']
        date = details['AssayStart']
        sponsor = details['Sponsor']
        study = details['StudyName']

        return tech, date, sponsor, study


    def get_samplelist(self):
        """ Get list of plate IDs and samples"""
        
        # Get sample table worksheet - can't use dynamic named range
        ws = self.wb.sheets['Sample Table']

        # Check empty
        if ws.range('A2').value is None:
            print("Sample list not found where expected. Please check that sample table is up to date.")
        
        # Sample table
        tbl = ws.range('A2').options(numbers=int).current_region
        
        # Empty dictionaries
        first_run = {}
        repeats = {}
        
        # Loop through sample list in F007
        for idx, row in enumerate(tbl.rows):
            
            # Get plates as keys and samples as values
            if idx != 0:
                plate = row[0].value  # Plate ID
                samples = list(row[1:].value)  # List of samples
                # If sample ID numeric - round or return as string or EMPTY
                samples = [round_to0(s) if s else "EMPTY" for s in samples]
                
                # If only a single letter - first run block
                if len(plate) == 1 and plate.isalpha():    
                    first_run[plate] = samples
                else:
                    repeats[plate] = samples

        return first_run, repeats

    def get_run_type(self):
        """ Determine whether the assay is first run, repeats or mixed """
        
        if self.first_list and self.repeats_list:
            return "mixed"
        elif self.first_list:
            return "first run"
        else:
            return "repeats"

    def get_qc_limits(self):
        """ Import QC Limits"""
        
        try:
            qc_lims = pd.read_csv(self.qc_file, index_col=0)
            return qc_lims
        except pd.errors.EmptyDataError:
            print("no data in QC limits file")
            return        
    
    def get_curve_vals(self):
        """ Import QC Limits"""
        
        try:
            curve_vals = pd.read_csv(self.curve_file, index_col=0)
            return curve_vals
        except pd.errors.EmptyDataError:
            print("no data in 007 Curve IgG reference file")
            return

    def get_xl_app(self):
        """ Find Excel app by id and return """

        for app in xw.apps:
            if app.pid == self.xl_id:
                return app

def change_assaydate(datestring):
    """ Convert the date string into clinical date format"""

    # Convert assay date to dd-mmm-yy format
    try:
        mydate = datetime.strptime(datestring, '%Y-%m-%d %H:%M:%S')
        mydate = mydate.strftime('%d-%b-%y')
    except ValueError:
        print("Can't parse the date - please check study start date in F007")    
        return
    except TypeError:
        mydate = datestring.strftime('%d-%b-%y')
    
    return mydate


def get_tech_initials(name):
    """ Convert technician name into technician initials"""

    '''
        Will require user input to select suggestion if required 
    '''
    initials = ''.join([c for c in name if c.isupper()])
    if not initials or len(initials) == 1:
        splitname = name.split(' ')

        if len(splitname) > 1:

            name1 = splitname[0]
            name2 = splitname[1]

            if not name2:
                print("no second name")
                return

            initials = name1[0] + name2[0]
            print("Expected two capitalised names. Use " + initials.upper() + "?")
        else:
            print("Please check technician name")
            return

    return initials.upper()


def get_file_from_path(path):

    if isinstance(path, list):
        path = path[0]

    head, tail = ntpath.split(path)
    f = tail or ntpath.basename(head)
    f = f.replace(".xlsm","")
    return f

def check_dup_blockid(block_ids, all_ids):
    """ Check that the block ID for a repeat plate is not the same as 
        a first run """
        
    # Loop through blocks and see if in plate list
    for b in block_ids:
        
        for a in all_ids:
            if b == a[-1] and b != a:
                return True
        
    return False


def round_to0(val):
    """ Returns number as string, unless not a number, in which case returns
        as string """
                
    try: # If number - ok
        new_val = "%.0f" % val
        
    except TypeError:  # If a string instead of a number
        new_val = np.array(val)
        
        try:
            new_val = "%.0f" % new_val
        except ValueError:  # If text that cannot be converted
            new_val = val
        
    return new_val



