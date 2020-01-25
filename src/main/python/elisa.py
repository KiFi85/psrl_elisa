import pandas as pd
from datetime import datetime
import numpy as np
import os


class ELISA:
    """ Class to contain all information about ELISA plate """

    def __init__(self, file, first_list, repeats_list, qc_limits, curve_vals, savedir,
                 cut_high_ods, cut_low_ods, apply_lloq):

        # Get elisa data
        self.file = file  # File path of plate csv
        self.plate_fail = None
        self.qc_limits = qc_limits
        self.curve_vals = curve_vals
        self.savedir = savedir  # Directory

        # Get curve/reader parameters and data
        self.parameters, self.data = self.get_data()
        if self.parameters is None or self.data is None:
            return

        self.warnings = []
        
        # Get r squared value for curve
        self.rsquared = self.get_rsquared()

        # Barcode details
        self.barcode, self.reader_id = self.get_ids()
        # Extract barcode details
        self.barc_tech, self.barc_date, self.barc_id = self.get_barcode_details()

        # Check template - will return True if applied
        self.template = self.template_applied()
        if not self.template:
            return

        # Testing details
        self.protocol, self.read_date, self.read_time = self.get_testdetails()

        # Get serotype and save name
        self.serotype = self.barc_id[:-1]
        self.save_name, self.pdf_path = self.get_save_name()

        # Test for R4
        if self.serotype != self.protocol:
            self.plate_fail = "R4"
            warn_str = "Plate " + self.barc_id + ": Wrong protocol applied (" + self.protocol + ")"
            self.warnings.append(warn_str)
            return

        # Reader temperature from last column and row
        self.reader_temp = self.data.iloc[-1, -1]

        # Results
        self.blank = self.get_blank()
        
        # Get samples - specify sample number 1:4
        self.sample_ids = self.get_sample_ids(first_list, repeats_list)
        smps = self.sample_ids
        
        # Create Sample Object with data, column number and sample ID
        self.Samples = [Sample(self.data, c, smps[c-1], self.curve_vals,
                               cut_high_ods, cut_low_ods, apply_lloq) for c in range(1, 5)]

        # Check for sample warnings
        self.get_sample_warnings()

        # Create Curve object
        self.Curve = Curve(self.data, sample_number=0, sample_id="Curve", 
                           serotype=self.serotype, curve_vals=self.curve_vals)

        # Create QCs
        self.High_QC = QC(self.data, sample_number=5, sample_id="HI")
        self.Low_QC = QC(self.data, sample_number=5, sample_id="LO")

        # Check plate fail
        self.plate_fail = self.get_plate_fail()

        # If curve, blank or protocol fail - Don't trend QCs (and don't check for OOR)
        if self.plate_fail in ["R16", "R11"]:
            self.High_QC.fail = True
            self.High_QC.result_recalc = "NR"
            self.Low_QC.fail = True
            self.Low_QC.result_recalc = "NR"
            return

        # Check for QC OOR fails
        if not self.plate_fail:
            self.plate_fail = self.check_qc_oor()

    def get_data(self):
        """ Import mars data file from csv"""

        # Import ELISA MARS CSV
        col_names = ['Row', 'Col', 'Ref', 'Group', 'Raw_405', 'Raw_620',
                     'Raw_405-620', 'BlankCorrect', 'Conc', 'RangeCheck', 'Temp']

        df = pd.read_csv(self.file, encoding="ISO-8859-1", names=col_names)

        # Remove units from curve concentrations
        df.replace(to_replace=r" ug/mL$", value="", regex=True, inplace=True)

        # Get row where data starts (if not found, not valid plate - return)
        try:
            pltidx = df.index[df['Row'] == "Well Row"].tolist()[0]
        except IndexError:
            return None, None

        # Get test Details
        parms = df.iloc[:pltidx, :3]
        parms.columns = ['ProtocolID', 'DateReader', 'Time']

        # Get plate Details
        plate = df.iloc[pltidx + 1:, :].copy()

        # Check data is present
        missing = plate['Conc'].isnull().all() | plate['BlankCorrect'].isnull().all()

        if missing:
            plate = None

        else:
            # Convert column reference to numeric and set multi-index
            plate.Col = pd.to_numeric(plate.Col)
            plate.set_index(['Row', 'Col'], inplace=True)
            plate.sort_index(inplace=True)

        return parms, plate

    def find_parameter(self, text, col):
        """ Find parameter in dataframe """

        df = self.parameters
        findtext = df[df[col].str.contains(text, regex=True) == True].values
        return findtext

    def get_ids(self):
        """ Get barcode and reader ID """

        ids = self.find_parameter("ID1:", "ProtocolID")
        barcode = ids[0][0].replace("ID1: ", "")
        reader = ids[0][1].replace("ID2: ", "")
        
        reader = "PSRLR3" if reader == "415-2020" else "PSRLR4"

        return barcode, reader

    def template_applied(self):
        """ Check that ICH template is applied. Return True if it has or False """

        # Check that the file contains rsquared value
        if self.rsquared is None:
            warn_str1 = 'Plate ' + self.barc_id
            warn_str2 = ': r squared value not found. ICH Template possibly not applied'

            # Add warning message
            self.warnings.append(warn_str1 + warn_str2)

        # Check that there is data
        if self.data is None:
            warn_str1 = 'Plate ' + self.barc_id
            warn_str2 = ': Fewer columns than expected in data. ICH Template possibly not applied'

            # Add warning message
            self.warnings.append(warn_str1 + warn_str2)

        # If one is not present - template not applied
        if self.rsquared is None or self.data is None:
            return False
        else:
            return True

    def get_testdetails(self):
        """ Return Protocol, Date and Time of read plate """

        ids = self.find_parameter("Test name:", "ProtocolID")
        serotype = ids[0][0].replace("Test name: ", "")
        read_date = ids[0][1].replace("Date: ", "")
        read_time = ids[0][2].replace("Time: ", "")

        return serotype, read_date, read_time

    def get_rsquared(self):
        """ Get r^2 value from data file """

        ids = self.find_parameter("^[r].$","DateReader")
        
        # If ids not empty array
        if ids.size:
            rsquared = round(pd.to_numeric(ids[0][2]), 3)  # Round to 3dp
        else:
            rsquared = None

        return rsquared

    def get_barcode_details(self):
        """ Get technician, date and plate ID from barcode """

        # Check if last character is alpga
        if self.barcode[0].isalpha():
            bstring = self.barcode[1:]

        # Get details - extract portion of barcode
        if bstring[-1] == "R":
            bdate = bstring[-7:-1]
            btech = bstring[-9:-7]
            bplate = bstring[:-9]
        else:
            bdate = bstring[-6:]
            btech = bstring[-8:-6]
            bplate = bstring[:-8]

        bdate = datetime.strptime(bdate, '%d%m%y')
        bdate = bdate.strftime('%d-%b-%y')

        return btech, bdate, bplate

    def get_save_name(self):
        """ Get the file name and pdf path for saved pdfs"""

        # Filename for pdf
        save_name = self.serotype + "_" + self.barcode

        # Full file path for pdf
        pdf_path = os.path.join(os.path.abspath(self.savedir), save_name + ".pdf")

        return save_name, pdf_path

    
    def get_blank(self):
        """ Get average blank value """

        df = self.data
        blankdf = df[df['Ref'].str.contains("Blank B")]['Raw_405-620']
        blankvals = round(pd.to_numeric(blankdf).mean(), 3)
        
        return blankvals

    def get_sample_ids(self, first_list, repeats_list):
        """ Get the list of samples associated with this plate """
    
        plate_id = self.barc_id
        sample_ids = ""
        
        # If "R" at end of barcode - get samples from repeat plate ID
        if self.barcode[-1] == "R":
            plate_list = list(repeats_list.keys())
            
            if plate_id in plate_list:
                sample_ids = repeats_list[plate_id]
        
        else:            
            plate_list = list(first_list.keys())
            plate_id = plate_id[-1]
            
            if plate_id in plate_list:
                sample_ids = first_list[plate_id]
        
        if sample_ids:            
            return sample_ids
        else:
            return

    def get_plate_fail(self):
        """ Check if plate has failed """

        # Check blank value
        if self.blank >= 0.1:
            return "R11"

        # Check r squared
        if not 0.9 <= self.rsquared <= 1.1:
            return "R16"

        # If curve fail
        if self.Curve.fail:
            return "R16"

        # If QC fail (NR)
        if self.High_QC.fail & self.Low_QC.fail:
            return "R2+R3"
        elif self.High_QC.fail:
            return "R2"
        elif self.Low_QC.fail:
            return "R3"
            
        else:
            return None

    def check_qc_oor(self):
        """ Check to see whether QC is out of range """

        # If high/low QC recalculated
        h_recalc = self.High_QC.result_recalc
        l_recalc = self.Low_QC.result_recalc

        # If recalculated value exists and isn't a string
        # Use recalculated value, if not, use result
        if h_recalc and not isinstance(h_recalc, str):
            hi_result = h_recalc
        else:
            hi_result = float(self.High_QC.result)

        if l_recalc and not isinstance(l_recalc, str):
            lo_result = l_recalc
        else:
            lo_result = float(self.Low_QC.result)

        # Get QC limits
        hi_1 = self.qc_limits.loc[self.serotype]['Hi_Lower']
        hi_2 = self.qc_limits.loc[self.serotype]['Hi_Upper']
        lo_1 = self.qc_limits.loc[self.serotype]['Lo_Lower']
        lo_2 = self.qc_limits.loc[self.serotype]['Lo_Upper']

        # If high control out of range
        if not hi_1 <= hi_result <= hi_2:
            r2 = True
        else:
            r2 = False

        # If low control out of range
        if not lo_1 <= lo_result <= lo_2:
            r3 = True
        else:
            r3 = False

        # If both out of range
        if r2 and r3:
            return "R2+R3"
        elif r2:
            return "R2"
        elif r3:
            return "R3"
        else:
            return None

    def get_sample_warnings(self):
        """ Get sample warning and append to plate details """

        # Loop through samples
        for s in self.Samples:
            s_id = str(s.sample_id)
            plate_id = str(self.barc_id)

            # Check if sample warning exists
            if s.warning:
                warn_str = 'Sample ' + s_id + \
                           ' on Plate ' + plate_id + \
                           ' is EMPTY & ' + s.warning
                self.warnings.append(warn_str)
            
class Sample:
    """ Class containing all sample details and checks """

    def __init__(self, data, sample_number, sample_id, curve_vals,
                 cut_high_ods, cut_low_ods, apply_lloq):
        
        self.data = data  # Sample data
        self.sample_number = sample_number  # Position on plate
        self.curve_vals = curve_vals  # Curve concentrations
        self.sample_id = sample_id  # Sample ID
        self.fail = False
        self.warning = ''
        self.cut_high_ods = cut_high_ods  # Upper OD limit
        self.cut_low_ods = cut_low_ods  # Lower OD limit
        self.apply_lloq = apply_lloq  # Apply LLOQ yes/no
        idx_list = ['A','B','C','D','E','F','G','H']

        # If sample ID == Empty then return empty series and values
        if self.sample_id.upper() == "EMPTY":
            self.replabels = pd.Series('', index=idx_list)
            self.average_concs = pd.Series('', index=idx_list)
            self.cvs = pd.Series('', index=idx_list)
            self.result_recalc = ''
            self.result = ''
            return

        # Get ODs and concentrations as arrays
        self.ods, self.concs = self.get_data()  # ODs and concs from data
        self.ods_orig = self.ods.copy()  # ALL ODs (as copy)
        self.concs_orig = self.concs.copy()  # ALL Concs (as copy)
        
        # Remove 2/0.1 ODs if required
        self.apply_od_cutoff()

        # Get replicates
        self.replicates, self.replabels = self.get_replicates()
        
        # Get average concentration now ODs and replicates have been removed
        self.average_concs = np.mean(self.concs, axis=1)
        self.result = self.get_result()

        # Get CVs between dilutions
        self.cvs = self.get_cvs()

        # Check if LLOQ - ignore if validation
        self.lloq = self.check_lloq() if self.apply_lloq else False

        # Check if repeat if not lloq
        if not self.lloq:
            self.result_recalc = self.check_recalc()
        else:
            self.result_recalc = "<0.15"

        # Format values
        self.format_values()

    def apply_od_cutoff(self):
        """ Remove ODs if necessary """

        if not self.cut_high_ods and not self.cut_low_ods:
            return
        
        # Get ODs within range
        mask = self.get_od_mask()

        # Mask ODs and concs outside limits
        self.ods = self.ods[mask]
        self.concs = self.concs[mask]

    def get_od_mask(self):
        """ Create a mask to include values between OD limits """
        
        upper_od = self.cut_high_ods
        lower_od = self.cut_low_ods
        
        if lower_od and upper_od:
            mask = (self.ods <= upper_od) & (self.ods >= lower_od)
        elif lower_od:
            mask = self.ods >= lower_od
        elif upper_od:
            mask = self.ods <= upper_od
            
        return mask

    def get_data(self):
        """ Get sample data """

        # Get data and determine column on plate
        plate = self.data
        n = self.sample_number
        cols = ((n*2)+1, (n*2)+2)  # e.g. Sample 1 = cols 3,4

        # Find samples from columns on plate
        samples = plate.loc[plate.index.get_level_values('Col').isin(cols)]

        # Create arrays of ODs and concentrations
        ods = samples['BlankCorrect'].unstack()
        concs = samples['Conc'].unstack()
        
        ods = ods.apply(pd.to_numeric)
        
        return ods, concs

    def get_replicates(self):
        """ Calculate %CV between replicate values """

        # Get concs and calculate replicates
        concs = self.concs.apply(pd.to_numeric, errors='coerce')
        # Replace values with nan if only one replicate value obtained
        concs.loc[concs.isna().any(axis=1), :] = np.nan

        # Calculate replicates
        replicates = np.std(concs, axis=1, ddof=1) / np.mean(concs, axis=1) * 100

        # Create mask for replicates >= 15
        mask = (replicates >= 15)

        # Replace concentrations with poor replicates as NaN. Update concs
        concs[mask] = float('NaN')
        self.concs = concs

        # Create replicate labels for poor replicates        
        idx = concs.index.tolist()
        replabels = pd.Series('', index=idx)
        replabels[mask] = '>15%'

        return replicates, replabels

    def get_cvs(self):
        """ Calculate %CV (% difference) between adjacent rows """

        # Get average concs - get index for each element (A:H)
        c = self.average_concs
        all_idx = c.index.tolist()

        # Get only present concs and index reference
        if c.isna().all():
            cvs = c.copy()
            return cvs

        c = c[c.notnull()]
        sub_idx = c.index.tolist()

        # Calculate difference between each row and max element
        d = np.absolute(np.diff(c, axis=0))
        m = [(max(z)) for z in zip(c, c[1:])]

        # Create a list of %CVs and add NaN at start
        cv = list(np.round((d / m * 100), decimals=3))
        cv.insert(0, np.nan)

        # Create series with index A:H inputting CVs only for relevant concentrations
        cvs = pd.Series(index=all_idx).astype(float)
        cvs.loc[sub_idx] = cv

        return cvs

    def get_result(self):
        """ Calculate average and return as ug/ml """

        # Average the concentration
        r = np.mean(self.average_concs)/1000
        if np.isnan(r):
            return ''
        else:
            return round_to3(r)  # Round to 3dp

    def check_lloq(self):
        """ Check if the sample is <LLOQ """
        
        # First use the result
        if self.result and float(self.result) < 0.15:
            return True
        elif self.result and float(self.result) >= 0.15:
            return False
        
        # Then check for poor replicates
        # Get list of replicates
        r = self.replicates
        # Get those over 15
        reprefs = r[r > 15].index
        # Subset the concentration
        c = self.concs.loc[reprefs]
        
        # If lloq then return
        # If not, one more check
        lloq = get_lloq_mean(c)

        if lloq:
            return True
        
        # Look for ANY values <0.15
        c = self.concs_orig.apply(pd.to_numeric, errors='coerce')
        
        # Only ODs >= 0.1        
        mask = (self.ods >= 0.1)
        c = c[mask]
        lloq = get_lloq_mean(c)

        if lloq:
            return lloq
        else:
            return False
    
    def check_recalc(self):
        """ Check if sample needs to be repeated or recalculate values"""

        # Number of concentrations returned
        c = self.average_concs.count()
        rpt = None
        high_low = None

        # If < 2 concentration - check for poor replicates
        # If c = 0 check for high and low concentrations
        if c < 2:
            rpt = self.check_repeats()
        if c == 0:
            high_low = self.check_empty_concs()

        # If repeat then return as fail
        if rpt:
            self.fail = True
            return rpt
        elif high_low and not rpt:
            self.fail = True
            return high_low

        # If nothing found - check for NP
        non_parallel = self.check_np()
        if non_parallel:
            self.fail = True
            return non_parallel

        # If no NP check for recalculation
        new_result = self.get_recalc()
        return round_to3(new_result)

    def check_repeats(self):
        """ Check if sample should be repeated """

        # Get number of replicates
        r = self.replicates
        n_reps = r[r > 15].count()

        # If n_reps > 1 then at least one poor replicate
        # If c<=1 then repeat
        if n_reps:
            return "RPT"
        else:
            return None

    def check_empty_concs(self):
        """ Check samples that have no values after OD limits """

        # If c = 0 then check for high (bottom row of plate)
        high_od = np.mean(np.array(self.ods_orig.loc['H'])) > 2
        high_col1 = np.sum(self.concs_orig.iloc[:, 0].str.contains(">")) > 1
        high_col2 = np.sum(self.concs_orig.iloc[:, 1].str.contains(">")) > 1

        if high_od | high_col1 | high_col2:
            self.warning = 'HIGH: Check repeat 1:500'
            return "RPT 1:500"

        # If c = 0 then check for low (top row of plate)
        low_od = np.mean(np.array(self.ods_orig.loc['A'])) < 0.1
        low_col1 = np.sum(self.concs_orig.iloc[:, 0].str.contains("<")) > 1
        low_col2 = np.sum(self.concs_orig.iloc[:, 1].str.contains("<")) > 1

        if low_od | low_col1 | low_col2:
            self.warning = 'LOW: Check QNS or <0.15'            
            return "Check \nLow"

        return None

    def check_np(self):
        """ Check if sample is non-parallel """

        # Check for RPTNP
        # If first CV > 20% Then NP
        cv_vals = self.cvs.notnull()

        # Check if RPT NP or >20% RPT
        if any(cv_vals) and self.cvs[cv_vals][0] > 20:

            # Index of CV position
            idx_cvs = self.cvs[cv_vals].index.tolist()[0]
            # Get row number (location) of
            idx_loc = self.cvs.index.get_loc(idx_cvs)
            # Check replicate above - if >15% --> >20% RPT
            rep_above = self.replicates.iloc[idx_loc - 1]

            # If non-parallel and replicate above then repeat
            if rep_above > 15:
                return '>20% \n RPT'
            else:
                return "RPT NP"
        else:
            return None

    def get_recalc(self):
        """ Get a recalculated value for sample if necessary """

        # List of CVs
        cv_vals = self.cvs.notnull()

        # Check if RPT NP or >20% RPT
        if any(cv_vals) and any(self.cvs[cv_vals] > 20):
            # Index of CV position
            idx_cvs = self.cvs[self.cvs > 20].index.tolist()[0]
            # Get row number (location) of
            idx_loc = self.cvs.index.get_loc(idx_cvs)
            result = np.round(np.mean(self.average_concs[:idx_loc]) / 1000, decimals=3)

            # Check whether new result is now <0.15 (unless validation assay)
            if result < 0.15 and self.apply_lloq:
                result = "<0.15"

        else:
            result = ''

        return result

    def format_values(self):
        """ Return all values as 3dp or empty strings """

        # Average concs
        self.average_concs.fillna(value='', inplace=True)
        self.average_concs = self.average_concs.apply(round_to3)

        # Replicates
        self.replicates.fillna(value='', inplace=True)
        self.replicates = self.replicates.apply(round_to3)

        # CVs
        self.cvs.fillna(value='', inplace=True)
        self.cvs = self.cvs.apply(round_to3)

class QC(Sample):
    """ QC Sample subclassed from sample """

    def __init__(self, data, sample_number, sample_id, curve_vals=None):
        super().__init__(data, sample_number, sample_id,
                         curve_vals=None, cut_low_ods=0.1, cut_high_ods=2, apply_lloq=False)


    def get_data(self):
        """ Get QC data - overridden function """

        # Get data and determine column on plate
        plate = self.data
        n = self.sample_number
        cols = ((n * 2) + 1, (n * 2) + 2)

        # Find samples from columns on plate
        if self.sample_id == "HI":
            rows = ['A', 'B', 'C', 'D']
        else:
            rows = ['E', 'F', 'G', 'H']

        samples = plate.loc[plate.index.get_level_values('Col').isin(cols)]
        samples = samples.loc[samples.index.get_level_values('Row').isin(rows)]

        # Create arrays of ODs and concentrations
        ods = samples['BlankCorrect'].unstack()
        concs = samples['Conc'].unstack()

        ods = ods.apply(pd.to_numeric)

        return ods, concs

    def check_recalc(self):
        """ Check if QC needs to be repeated or recalculate values"""

        # Number of concentrations returned
        c = self.average_concs.count()

        # If <= 1 concentration - check for poor replicates
        # If c = 0 check for high and low concentrations
        if c <= 1:
            self.fail = True
            return "NR"

        # If not NR by replicate - check for NP
        non_parallel = self.check_np()
        if non_parallel:
            self.fail = True
            return "NR"

        # If no NP check for recalculation
        new_result = self.get_recalc()
        return new_result

    def format_values(self):
        """ Return all values as 3dp or empty strings """

        # Average concs
        self.average_concs.fillna(value='', inplace=True)
        self.average_concs = self.average_concs.apply(round_to3)

        # Replicates
        self.replicates.fillna(value='', inplace=True)
        self.replicates = self.replicates.apply(round_to3)

        # cvs
        self.cvs.fillna(value='', inplace=True)
        self.cvs = self.cvs.apply(round_to3)

class Curve(Sample):
    """ Curve Class subclassed from sample """

    def __init__(self, data, sample_number, sample_id, serotype, curve_vals):

        self.serotype = serotype
        self.curve_vals = curve_vals
        # Get curve top point
        self.top_point = self.get_top_point()

        super().__init__(data, sample_number, sample_id, curve_vals,
                         cut_low_ods=None, cut_high_ods=None, apply_lloq=False)

        # Calculate curve replicates
        self.replicates, self.replabels = self.get_replicates()

        # Check for curve fail due to poor replicates
        self.fail = self.check_fail()

        self.format_values()

    def get_replicates(self):
        """ Calculate %CV between replicate values """

        # Get concs and calculate replicates
        concs = self.concs.apply(pd.to_numeric, errors='coerce')
        av_ods = np.mean(self.ods, axis=1)

        # Replace values with nan if only one replicate value obtained
        concs.loc[concs.isna().any(axis=1), :] = np.nan
        self.concs = concs

        # calculate replicates
        replicates = np.std(concs, axis=1, ddof=1) / np.mean(concs, axis=1) * 100

        # Create mask for replicates >= 15 and OD < 0.1
        mask_reps = (replicates >= 15) & (av_ods < 0.1)
        av_concs = np.mean(self.concs, axis=1)
        mask_top_point = av_concs > self.top_point
        av_concs[mask_top_point] = np.nan
        self.average_concs = av_concs

        # Create labels for curve (>max or <0.1)
        idx = concs.index.tolist()
        replabels = pd.Series('', index=idx)
        replabels[mask_reps] = '<0.1'
        replabels[mask_top_point] = '>max'
        replicates[mask_top_point] = np.nan
        return replicates, replabels

    def check_fail(self):
        """ Check for a curve fail based on replicates """
        
        # Get average ODs and check for poor replicates
        av_ods = np.mean(self.ods, axis=1)
        poor_reps = (self.replicates > 15) & (av_ods >= 0.1)
        
        # If only one poor replicate and occurs in top row:
            # Check average ods for top two points
        if sum(poor_reps) == 1 and poor_reps['A']:
            
            # If both cal1 and cal2 are >= 2
            top2_check = av_ods[['A', 'B']] >= 2
            
            # If both above 2, change cal labels and return no fail
            if top2_check.all():
                self.replabels[['A', 'B']] = '>2.0'
                return False
            else:
                return True
        
        # Else if any poor replicates
        elif poor_reps.any():
            return True
        else:
            False

    def get_top_point(self):
        """ Get the IgG concentration assigned to 007sp 
            i.e. the top point on the curve """
        
        top_point = self.curve_vals.loc[str(self.serotype)]['cal1_IgG']
        return top_point

    def format_values(self):
        """ Return all values as 3dp or empty strings """

        # Average concs
        self.average_concs.fillna(value='', inplace=True)
        self.average_concs = self.average_concs.apply(round_to3)

        # Replicates
        self.replicates.fillna(value='', inplace=True)
        self.replicates = self.replicates.apply(round_to3)


def get_lloq_mean(concs):
    """ Get means from conc array for checking LLOQ """
    
    rowmeans = concs.mean(axis=1, skipna=False)
    grandmean = np.round(rowmeans.mean()/1000,decimals=3)

    # If NaN from ignoring NaN - try again (will look for any valid value)
    if np.isnan(grandmean):
        rowmeans = concs.mean(axis=1, skipna=True)
        grandmean = np.round(rowmeans.mean()/1000,decimals=3)
    
    if np.isnan(grandmean):
        return False

    if grandmean and grandmean < 0.15:
        return True
    elif grandmean and grandmean >= 0.15:
        return False


def round_to3(val):
    """ Returns number as 3dp string, unless not a number, in which case returns
        as string """

    try:  # If number - ok
        new_val = "%.3f" % val

    except TypeError:  # If a string instead of a number
        new_val = np.array(val)

        try:
            new_val = "%.3f" % new_val
        except ValueError:  # If text that cannot be converted
            new_val = val

    return new_val

