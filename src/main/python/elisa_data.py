import time
import numpy as np
import pdfkit
import jinja2
import csv
import xlwings as xw
import ntpath
import pandas as pd
import win32print
import win32api
from pathlib import Path
import os

# PDF OPTIONS
pdf_options = {
            'quiet': '',
            'page-size': 'A4',
            'dpi': 300,
            'disable-smart-shrinking': ''
        }


class ELISAData:
    """ Class containing functions to process elisa data """

    def __init__(self, assay, savedir, trend_file, f093_file, master_file,
                 xl_id, parms_dict, ctx):

        self.assay = assay
        self.savedir = savedir  # Dir when ELISA data is stored

        self.trend_file = trend_file  # QC Trending file
        self.f093 = f093_file  # F093 Excel template
        self.master_file = os.path.abspath(master_file)  # Study master file
        self.xl_id = xl_id  # ID of working Excel process
        self.printer = win32print.GetDefaultPrinter()  # Default printer
        self.parms_dict = parms_dict  # The parameters used for processing (OD/LLOQ)
        self.ctx = ctx  # Application Context (for resources)
        # A summary table of plate fails to check for R35s
        self.df_plates = pd.DataFrame(
            columns=['Plate', 'Sample1', 'Sample2', 'Sample3', 'Sample4', 'Fail'])

        # PDF configuration (wkhtmltopdf)
        self.pdf_config = pdfkit.configuration(wkhtmltopdf=ctx.pdf_exe)

        self.warnings = []  # List of warnings
        self.plate_fails = 0  # Plate fail counter
        self.plate_list = []  # List of plate IDs
        self.results_list = []  # List of sample results for every plate

        self.trend_data = []  # List of trending data

        # Check that a F093 has been created
        # Create F093 file name
        f093_save_name = self.assay.f007_ref + " F093"
        self.f093_save_name = os.path.join(os.path.abspath(self.savedir),
                                 f093_save_name)

        # If json file exists - import - else create empty dataframe        
        if os.path.isfile(self.f093_save_name + ".json"):
            self.f093_df = import_f093_json(self.f093_save_name)
            self.df_names = self.f093_df.columns.tolist()

        else:
            self.f093_df = pd.DataFrame()
            self.df_names = []

    def get_html_template(self, template_file):
        """ Get configurations for pdfkit and jinja2 """

        # Get the parent path of the template file
        template = self.ctx.template
        searchpath = os.path.join(Path(template).parent).replace("\\", "/")

        # CREATE TEMPLATE ENVIRONMENT FOR JINJA2
        templateLoader = jinja2.FileSystemLoader(searchpath=searchpath)
        templateEnv = jinja2.Environment(loader=templateLoader)
        template = templateEnv.get_template(template_file)

        return template

    def input_plate_data(self, elisa, to_pdf=True):
        """ Inputs data to html template and saves to pdf (if to_pdf = True) """

        # If there are warning messages - append
        if elisa.warnings:
            for w in elisa.warnings:
                self.warnings.append(w)

        # If ELISA ICH template not applied - stop
        if not elisa.template:
            return

        # If read on wrong protocol and exporting to Excel
        r4 = True if elisa.plate_fail == "R4" else False

        # HTML template to use - gather data if not R4
        if r4:
            template_file = 'r4template.html'
        else:
            template_file = 'template.html'

            # Add plate details to plate list
            plate_details = get_plate_details(elisa)
            self.plate_list.append(plate_details)

            # Input as summary to plate dataframe
            row = len(self.df_plates.index)
            df_plates = [plate_details[0]]
            for r in range(2, 7):
                df_plates.append(plate_details[r])

            self.df_plates.loc[row] = df_plates

            # Add sample results to results list
            results = self.get_result_details(elisa)
            self.results_list.append(results)

            # Increase fail if plate fail
            self.plate_fails += 1 if elisa.plate_fail else 0

        # Create pdf
        if to_pdf:
            self.create_pdf(template_file, elisa)

    def create_pdf(self, template_file, elisa):
        """ Create a pdf from an html template"""

        # Get html template file
        template = self.get_html_template(template_file)

        # Get save name
        pdf_path = elisa.pdf_path

        # Get ods and concs tables
        ods, concs = get_table_details(elisa.data)
        od_html = ods.to_html(classes='od_table', table_id='od_tbl')
        concs_html = concs.to_html(classes='concs_table', table_id='concs_tbl')

        # Render html template to string
        version = self.ctx.build_settings['version']
        rendered = template.render(elisa=elisa, assay=self.assay, version=version,
                                   tables=[od_html, concs_html],
                                   titles=ods.columns.values)

        # Save rendered template as PDF
        pdfkit.from_string(rendered, pdf_path, configuration=self.pdf_config, css=self.ctx.css, options=pdf_options)

    def data_to_table(self, elisa):
        """ Creates f093 dataframe if first plate or updates dataframe. """

        # If no ELISA template - move to next plate
        if not elisa.template:
            return

        # If repeat - don't add to table
        if elisa.barcode[-1] == "R":
            return

        # If first file then initialise dataframe
        # Else add to dataframe
        if self.f093_df.empty:
            self.create_f093_df(elisa.serotype)

        # Check for serotype in existing dataframe
        serotype_check = "_" + elisa.serotype + "$"
        checklist = self.f093_df.filter(regex=serotype_check).columns.tolist()

        # If serotype not shown in dataframe, get new columns and input dates
        if not checklist:
            get_new_colnames(self.df_names, elisa.serotype)
            self.f093_df = self.f093_df.reindex(columns=self.df_names)

        # Input results
        self.input_f093_results(elisa)

    def create_f093_df(self, serotype):
        """ Initialises the f093 dataframe based on study, samples and plate IDs """

        # Get list of samples and the plates on which they were run
        sample_list = list(self.assay.first_list.values())

        # Plate ID*4
        key_list = [list(k)*4 for k in self.assay.first_list.keys()]

        # Convert to string
        sample_list = [str(item) for sublist in sample_list for item in sublist]
        key_list = [str(item) for sublist in key_list for item in sublist]

        # Get the column names for the data table
        self.df_names = get_f093_colnames(serotype, init=True)
        # Create empty dataframe based on n samples
        self.f093_df = pd.DataFrame(index=range(0, len(sample_list)),
                                    columns=['Sample ID', 'Plate ID'])

        # Input general testing details (study, sample plates, dates)
        self.f093_df['Sample ID'] = sample_list
        self.f093_df['Plate ID'] = key_list

    def input_f093_results(self, elisa):
        """ Input results to f093 dataframe """

        # Get index of result column from dataframe for serotype
        result_col = "Result_" + elisa.serotype
        col_ref = self.f093_df.columns.get_loc(result_col)

        # IF plate fail, can just remove results, lab date and technician from plate
        if elisa.plate_fail:

            # Plate fail code for formatting
            self.f093_df.loc[self.f093_df['Plate ID'] == elisa.barc_id[-1],
                             self.f093_df.columns[col_ref]] = elisa.plate_fail

            return

        # Get ids, values and sample fails
        ids, results, fails = zip(*[get_sample_info(s) for s in elisa.Samples])

        # Boolean array where sample IDs match
        sample_bool = self.f093_df['Sample ID'].isin(ids)

        # Row index of samples
        sample_rows = self.f093_df.index[sample_bool].tolist()

        # Input results
        self.f093_df.loc[sample_rows, self.f093_df.columns[col_ref]] = results

    def f093_to_excel(self):
        """ Save dataframe as table in F093 and format """

        # Rename dataframe columns
        self.f093_df.columns = self.f093_df.columns.str.replace(
                "Result_", "PnC-IgG-ELISA type ")

        # Save dataframe to json
        self.f093_df.to_json(self.f093_save_name + ".json", orient='table')

        # Get working Excel process based on pid
        app = self.get_xl_app()
        # Open F093 workbook
        wb = app.books.open(self.f093)

        # Add dataframe to sheet
        ws = wb.sheets['Results']
        ws.range('A6').options(index=False).value = self.f093_df

        # Run macro in workbook
        formatting = wb.macro('format_page')
        formatting(self.assay.f007_ref)

        # Save F093        
        wb.save(self.f093_save_name + ".xlsm")
        wb.close()

    def create_summary(self):
        """ Create a summary file containing assay details, errors/warnings
            and plate read time/sample data """

        # Get list of testing details
        test_list = self.get_testing_summary()

        # File name to save
        file_name = "run_details " + self.assay.f007_ref + ".csv"
        file_name = os.path.join(os.path.abspath(self.savedir),
                                 file_name)

        # Write lists and plate dataframe to CSV
        with open(file_name, 'a', newline='') as csvFile:

            writer = csv.writer(csvFile)

            # Testing details
            for r in test_list:
                writer.writerow(r)

            # Warnings
            self.warnings.append("")
            self.warnings.append("")

            # Loop through warnings and write
            for idx, w in enumerate(self.warnings):

                if idx == 0:
                    writer.writerow(["Warnings:", w])
                else:
                    writer.writerow(["", w])
#
            # plate list column names
            writer.writerow([
                    "Plate",
                    "Read Time",
                    "Sample 1",
                    "Sample 2",
                    "Sample 3",
                    "Sample 4",
                    "Plate Fail"])

            # plate details
            for r in self.plate_list:
                writer.writerow(r)

            csvFile.close()

    def create_master(self):
        """ Create a master study file if doesn't exist. Create headers and
            write sample results """

        time_ctr = 0
        while time_ctr < 20:
            time.sleep(1)
            time_ctr += 1

            # Open master_details so it will be locked for editing
            try:
                with open(self.master_file, 'w', newline='') as csvFile:

                    writer = csv.writer(csvFile)
                    # headers
                    headers = ['Sample ID',	'PnC Serotype', 'Plate ID',
                               'Result', 'Amended Result', 'LABDT', 'Technician']
                    writer.writerow(headers)

                    # Write results
                    for l in self.results_list:
                        for row in l:
                            writer.writerow(row)

                    csvFile.close()
                    break
            except PermissionError:
                print("Already open")

    def update_master(self):
        """ Update the master study file. Remove duplicate entries.
            Detect NRs """

        # Get dataframe with new data (will remove duplicates and format)
        df = self.fill_master_details()

        time_ctr = 0

        # check each second whether the master csv file is open and write when it becomes free (timeout at 20 seconds)
        while time_ctr < 20:
            time.sleep(1)
            time_ctr += 1

            # Open master_details so it will be locked for editing
            try:
                with open(self.master_file, 'w', newline='') as csvFile:

                    # Save dataframe to csv
                    df.to_csv(self.master_file, index=False)
                    csvFile.close()
                    break

            except PermissionError:
                print("Already open")

    def fill_master_details(self):
        """ Fill master details with sample results. 
            Create a dataframe as easier to find duplicates """

        # Import master_details file
        df = self.get_master_df()

        # Write results to dataframe if not there
        for l in self.results_list:
            for row in l:
                # If row already in master - continue
                if not df[(df == row).all(axis=1)].empty:
                    continue
                else:
                    df.loc[len(df)] = row

        # Drop Duplicates (sample plate will have been re-printed)
        df.drop_duplicates(keep='first', inplace=True)

        # Sort by date if more than one date
        df['LABDT'] = pd.to_datetime(df['LABDT'])

        df.sort_values(by=['LABDT','PnC Serotype', 'Plate ID', 'Sample ID'], inplace=True)
        df['LABDT'] = df['LABDT'].dt.strftime('%d-%b-%y')

        # Get NP duplicates and assign second one to 'NR'
        dups = df[df['Result'].eq('NP')].duplicated(subset=[
            'Sample ID', 'PnC Serotype', 'Result'], keep='first')

        nr_idx = dups[dups == True].index
        df.loc[nr_idx, 'Amended Result'] = 'NR'

        return df

    def get_master_df(self):
        """ Get the master details dataframe. 
            If exists - import, if not - create """

        try:
            # Import master_details file
            df = pd.read_csv(self.master_file,
                             dtype={'Sample ID': 'str', 'PnC Serotype': 'str'},
                             parse_dates=['LABDT'], infer_datetime_format=True)

        except pd.errors.EmptyDataError:  # If no data found - create

            headers = ['Sample ID', 'PnC Serotype', 'Plate ID',
                       'Result', 'Amended Result', 'LABDT', 'Technician']

            df = pd.DataFrame(columns=headers)

        df['Sample ID'] = df['Sample ID'].astype(str)
        df['PnC Serotype'] = df['PnC Serotype'].astype(str)
        df['LABDT'] = pd.to_datetime(df['LABDT'])

        # Format date column
        df['LABDT'] = df['LABDT'].dt.strftime('%d-%b-%y')
        df.fillna(value='', inplace=True)

        return df

    def update_summary(self):
        """ Update the run_details summary file when printing extra plates 
            or re-printing plates """

        # Import run_details as dataframe
        df, save_name = self.import_summary()

        # Update data frame with new plates, plate counts and fail counts
        df = update_plate_summary(df, self.plate_list)

        # Add new warnings to dataframe and return
        df = get_warnings_df(df, self.warnings)

        # Check that all warnings are required
        # if warning relates to a plate that is in the plate list, remove warning
        df = check_warnings(df)

        # Reset index and Save
        df.reset_index(drop=True, inplace=True)
        df.to_csv(save_name, header=False, index=False)

    def import_summary(self):
        """ Import the run_details file for the assay """

        # Get file name of run_details based on data directory
        file_name = "run_details " + self.assay.f007_ref + ".csv"
        file_name = os.path.join(os.path.abspath(self.savedir),
                                 file_name)

        # Import existing run_details file
        headers = ['ref', 'val', 'smp1', 'smp2', 'smp3', 'smp4', 'fail']
        df = pd.read_csv(file_name, header=None, names=headers)

        return df, file_name

    def get_testing_summary(self):
        """ Create a summary list of testing details """

        assay = self.assay
        n_plates = len(self.plate_list)

        test_list = [["Master File Study Path:", self.master_file, "", "", "", "", ""],
                     ["", "", "", "", "", "", ""],
                     ["", "","", "", "", "", ""],
                     ["Run Date:", assay.date, "", "", "", "", ""],
                     ["Technician:", assay.tech, "", "", "", "", ""],
                     ["Sponsor:", assay.sponsor, "", "", "", "", ""],
                     ["Study:", assay.study, "", "", "", "", ""],
                     ["OD Upper Limit:", self.parms_dict['OD_Upper'], "", "", "", "", ""],
                     ["OD Lower Limit:", self.parms_dict['OD_Lower'], "", "", "", "", ""],
                     ["Apply LLOQ::", self.parms_dict['LLOQ'], "", "", "", "", ""],
                     ["", "", "", "", "", "", ""],
                     ["", "", "", "", "", "", ""],
                     ["Number of Plates:", n_plates, "", "", "", "", ""],
                     ["Number of Fails:", self.plate_fails, "", "", "", "", ""],
                     ["", "", "", "", "", "", ""],
                     ["", "", "", "", "", "", ""]]

        return test_list

    def get_trend_data(self, elisa):
        """ Save the trending details to a list of lists """

        t_data = []
        t_data.append(elisa.barc_date)  # Assay date
        t_data.append(elisa.barc_tech)  # Technician
        t_data.append(self.assay.sponsor)  # Sponsor
        t_data.append(self.assay.study)  # Study
        t_data.append(elisa.barc_id)  # Plate ID
        t_data.append(elisa.serotype)  # Serotype

        # Try round the result - if attribute error must be NR
        try:
            hi = elisa.High_QC.result if not elisa.High_QC.result_recalc else elisa.High_QC.result_recalc
            lo = elisa.Low_QC.result if not elisa.Low_QC.result_recalc else elisa.Low_QC.result_recalc
            hi = round_to3(hi)
            lo = round_to3(lo)

            t_data.append(hi)
            t_data.append(lo)
        except AttributeError:
            t_data.append("NR")
            t_data.append("NR")

        t_data.append(elisa.plate_fail)
        self.trend_data.append(t_data)  # Append to master trending data list

    def update_trending(self):
        """ Update the master trending file. Remove duplicate entries. """

        # Get dataframe with new data (will remove duplicates and format)
        df = self.fill_trending_details()

        time_ctr = 0

        # check each second whether the trending csv file is open and write when it becomes free (timeout at 20 seconds)
        while time_ctr < 20:
            time.sleep(1)
            time_ctr += 1

            # Open trending file so it will be locked for editing
            try:
                with open(self.trend_file, 'w', newline='') as csvFile:

                    # Save dataframe to csv
                    df.to_csv(self.trend_file, index=False)
                    csvFile.close()
                    break

            except PermissionError:
                print("Already open")

    def fill_trending_details(self):
        """ Fill trending details with sample QC results. 
            Create a dataframe as easier to find duplicates """

        # Import trending file
        df = self.get_trend_df()

        # Write results
        for row in self.trend_data:
            df.loc[len(df)] = row

        # Sort by date
        df['Lab Date'] = pd.to_datetime(df['Lab Date'])
        df.sort_values(by=['Lab Date'], inplace=True)
        df['Lab Date'] = df['Lab Date'].dt.strftime('%d-%b-%y')

        # Drop Duplicates (will have been re-printed)
        df.drop_duplicates(keep='first', inplace=True)
        return df

    def get_trend_df(self):
        """ Import the trending CSV file """

        # Import trending (only get here if found at startup)        
        df = pd.read_csv(self.trend_file,
                         parse_dates=['Lab Date'],
                         infer_datetime_format=True)

        return df

    def get_xl_app(self):
        """ Find Excel app by id and return """

        for app in xw.apps:
            if app.pid == self.xl_id:
                return app

    def get_result_details(self, elisa):
        """ Get sample ID, serotype, result, labdate and technician as list """

        # Empty list to store results
        result_list = []
        # Empty amendment
        # amendment = ""

        # Get plate and block ID
        plate_id = elisa.barc_id
        block_id = elisa.barc_id[-1]

        # Loop through samples and get details
        for s in elisa.Samples:

            if s.sample_id.upper() != "EMPTY":

                # If plate fail - just report fail, else get the sample result
                if elisa.plate_fail:
                    result = elisa.plate_fail
                else:
                    result = get_sample_result(s)

                # Check sample amendment
                # if not self.amendments.empty:
                #     amendment = self.get_amendments(plate_id, s)

                # Overwrite result with amendment if not empty
                # result = amendment if amendment else result

                # Create a list of results to input to master
                smp_list = [s.sample_id,
                            elisa.serotype,
                            block_id,
                            result,
                            "",
                            elisa.barc_date,
                            elisa.barc_tech]

                # Append list to result_list
                result_list.append(smp_list)

        return result_list

    # def get_amendments(self, plate_id, sample_id):
    #     """ Check for plate or sample amendments in amendments dataframe """
    #
    #     amendment = ""
    #
    #     # Check sample is in dataframe
    #     if sample_id in self.amendments.Sample.tolist():
    #
    #         # Get row
    #         row = self.amendments[self.amendments['Sample'] == sample_id]
    #         # Get amendment
    #         amendment = row['Amendment'].tolist()[0]
    #
    #     # Else check if plate is in dataframe
    #     elif plate_id in self.amendments.Plate.tolist():
    #
    #         # Get row
    #         row = self.amendments[self.amendments['Plate'] == plate_id]
    #         # Get amendment
    #         amendment = row['Amendment'].tolist()[0]
    #
    #     return amendment


def get_file_from_path(path):

    if isinstance(path, list):
        path = path[0]

    head, tail = ntpath.split(path)
    f = tail or ntpath.basename(head)
    f = f.replace(".CSV", "")
    f = f.replace(".csv", "")
    return f


def get_f093_colnames(serotype, init=False):
    
    """ Return the column names for the f093 dataframe for result input.
    If initialising, will contain study, sample and plate ID columns """

    serotype = str(serotype)
    colnames = ['Result_' + serotype]

    # If initialising the F093
    if init:

        init_colnames = ['Sample ID',
                         'Plate ID']

        colnames = init_colnames + colnames

    return colnames


def get_new_colnames(current_names, new_serotype):
    ''' Get the column names for a new serotype for entry into f093 dataframe.
        Returns a new list of columns names, after sorting the serotypes by
        numerical order '''

    idx = 0
    # Get new serotype as string and integer for sorting
    new_serotype = str(new_serotype)
    new_sero_num = int(new_serotype[:-1]) if new_serotype[-1].isalpha() else int(new_serotype)
    
    # Loop through existing names to get a list of existing serotypes
    # If result column - get the serotype as a number
    for idx, s in enumerate(current_names):
        if "Result_" in s:
            serotype = s[7:]
            sero_num = serotype[:-1] if serotype[-1].isalpha() else serotype
            sero_num = int(sero_num)

            # If the serotype is already in the list - return
            # If the serotype found > new serotype then break as should be
            # inserted at this index
            if new_serotype == serotype:
                return
            elif sero_num > new_sero_num:
                idx -= 1
                break

    # Get the column names for the new serotype
    new_column = "Result_" + new_serotype

    current_names.insert(idx+1,new_column)
        
    return current_names


def get_sample_info(sample):

    """ Get list of ids and their results from Sample objects.
        Fail is a boolean: true if repeat or np """
    
    sample_id = sample.sample_id
    
    if sample_id.upper() == "EMPTY":
        result = ""
        fail = True
    else:
        result = sample.result_recalc if sample.result_recalc else sample.result
        result = round_to3(result)
        fail = sample.fail

    return sample_id, result, fail


def create_table(ax, cell_text, columns, rows):

    col_width = [0.3] * len(columns)
    the_table = ax.table(cellText=cell_text, colLabels=columns,
                         rowLabels=rows, cellLoc='center', colWidths=col_width)
    
    cell_height = 1.5
    
    for i in range(1, 9):  # Loop through rows

        # A to H as grey with black edges
        the_table[(i, -1)].set_facecolor("#E8E8E8")
        the_table[(i, -1)].set_edgecolor("k")
        the_table[(i, -1)].set_width(1.5)
        the_table[(i, 0)].set_edgecolor("k")
        the_table[(i, 0)].visible_edges = 'L'
        the_table[(i, -1)].set_height(cell_height)
        the_table[(i, 0)].set_height(cell_height)
        
        for j in range(0,12):

            # Row 0 is column headers - grey with black outline
            the_table[(0, j)].set_facecolor("#E8E8E8")
            the_table[(0, j)].set_edgecolor("k")
            the_table[(0, j)].set_height(cell_height)

            # Row 1 is the top row of numbers
            # Black edges at top
            the_table[(1, j)].set_edgecolor("k")
            the_table[(1, j)].visible_edges = 'T'

            # All other cells - no edges
            the_table[(i, j)].visible_edges = 'open'
            the_table[(i, j)].set_height(cell_height)
            # the_table[(i, j)].set_edgecolor("w")
            
    the_table.auto_set_font_size(False)
    the_table.set_fontsize(13)

    return ax


def get_table_details(data):
    """ Return array of values to be used to create OD or Conc table
        Input data as column from dataframe """

    # Create array of 8x12
    ods = data['BlankCorrect'].unstack()
    concs = data['Conc'].unstack()

    # Get row and column names
    columns = ods.columns.tolist()
    rows = ods.index.tolist()

    # Replace NaN with 0
    ods.fillna(0, inplace=True)
    concs.fillna(0, inplace=True)

    # Round and report to 3dp
    od_array = ods.apply(np.vectorize(round_to3)).values
    conc_array = concs.apply(np.vectorize(round_to3)).values

    od_df = pd.DataFrame(data=od_array, index=rows, columns=columns)
    conc_df = pd.DataFrame(data=conc_array, index=rows, columns=columns)

    return od_df, conc_df


def round_to3(val):
    """ Returns number as 3dp string, unless not a number, in which case returns
        as string """
                
    try: # If number - ok
        new_val = "%.3f" % val
        
    except TypeError:  # If a string instead of a number
        new_val = np.array(val)
        
        try:
            new_val = "%.3f" % new_val
        except ValueError:  # If text that cannot be converted
            new_val = val
        
    return new_val


def get_plate_details(elisa):
    """ Get plate details - ID, read time and samples """

    plate_list = [s.sample_id for s in elisa.Samples]
    plate_id = elisa.barc_id
    read_time = elisa.read_time
    plate_list.insert(0, read_time)
    plate_list.insert(0, plate_id)
    plate_list.append(elisa.plate_fail)

    return plate_list


def get_sample_result(sample):
    """ Determine how to report a sample result based on it's recalculated
        result """
        
    # Replacement values
    result_dict = {'RPT NP': 'NP',
                   '>20% \n RPT': '20% RPT',
                   'Check \nLow': 'Empty: Low'} 
    
    # Get result if not been recalculated
    result = sample.result if not sample.result_recalc else sample.result_recalc
    
    # If result is in the dictionary - replace
    if result in result_dict:
        result = result_dict[result]
        
    # Round to 3dp and return as string (unless string to begin with)
    result = str(round_to3(result))
    
    return result


def update_plate_summary(df, new_plates):
    """ Input dataframe of existing run_details csv. Update the number of
        plates based on the new number being re-printed and the existing
        list of plates. new_plates is a list of plate details """

    old_plates = get_plate_ids(df)

    # Loop through rows in the list of new_plates
    for r in new_plates:
        if r[0] not in old_plates:

            # Append plate list to dataframe
            df = df.append(pd.DataFrame([r], columns=df.columns.tolist()),
                           ignore_index=True)
        else:
            # Look up index of plate ID and replace
            df.loc[df['ref'] == r[0], :] = r

    # Get new plate and fail numbers and update
    n_plates = len(get_plate_ids(df))
    df.loc[df['ref'] == 'Number of Plates:', 'val'] = n_plates
    n_fails = get_fail_count(df)
    df.loc[df['ref'] == 'Number of Fails:', 'val'] = n_fails

    return df


def warnings_to_list(df):
    """ Get the list of the current warnings, including the start and end
        of the list """

    # Get row where warnings start
    start = df.loc[df['ref'] == 'Warnings:'].index.values[0]

    # Two rows above where the plate appears
    end = df.loc[df['ref'] == 'Plate'].index.values[0] - 2

    # Get existing warnings
    warnings = df.loc[range(start, end), 'val'].tolist()

    return start, end, warnings


def insert_warnings(df, warnings, warn_end):
    """ Insert any new warnings if present, remove unnecessary warnings """

    # Number of new warnings
    n_warns = len(warnings)

    # Index to insert new warnings
    insert_idx = list(range(warn_end, warn_end + n_warns))

    # Create a new dataframe to concatenate
    new_warnings = pd.DataFrame({
        'ref': [''] * n_warns,
        'val': warnings,
        'smp1': [''] * n_warns,
        'smp2': [''] * n_warns,
        'smp3': [''] * n_warns,
        'smp4': [''] * n_warns,
        'fail': ['']}, index=insert_idx)

    # Create a new dataframe
    df = pd.concat([
        df.iloc[:warn_end], new_warnings, df.iloc[warn_end:]
    ]).reset_index(drop=True)

    return df


def get_warnings_df(df, warnings):
    """ Return the new dataframe with updated and deleted warnings """

    # New list of warnings (any that need to added)
    new_warn_list = []

    # Get details on current warnings in list
    warn_start, warn_end, old_warnings = warnings_to_list(df)

    # Find any warnings that aren't present in the file already
    for w in warnings:
        if w and w not in old_warnings:
            new_warn_list.append(w)

    # If there are new warnings to be added - insert in dataframe
    if new_warn_list:
        df = insert_warnings(df, new_warn_list, warn_end)

    return df


def get_plate_ids(df):
    """ Get a list of plate IDs from summary table"""

    # Get a list of plates already in run_details
    plt_idx = df.loc[df['ref'] == 'Plate'].index.values[0] + 1
    plate_ids = df.iloc[plt_idx:, :]['ref'].tolist()

    return plate_ids


def get_fail_count(df):
    """ Get a count of fail codes from summary table"""

    # Get a list of failed plates in run_details
    plt_idx = df.loc[df['ref'] == 'Plate'].index.values[0] + 1
    fail_list = df.iloc[plt_idx:, :]['fail'].notna().tolist()
    fail_count = sum(fail_list)

    return fail_count


def update_warnings(df, warnings, warn_start):
    """ During the summary file update, those warnings that relate to plates not in the list
        are retained and passed here as a list for re-adding to the table. """

    # Get number of warnings to write
    n_warnings = len(warnings)

    # Rewrite warnings
    new_warn_end = warn_start + n_warnings
    warn_range = range(warn_start, new_warn_end)
    for idx, row in enumerate(warn_range):
        df.loc[row, 'val'] = warnings[idx]

    return df


def check_warnings(df):
    """ Check that the warnings in the dataframe are required.
        Compare the plate IDs printed to plate IDs in the warnings list.
        If the plate ID is present, the warning can be removed """

    # Get plate IDs from dataframe
    p_ids = get_plate_ids(df)

    # Get the existing warnings as a list from the dataframe
    warn_start, warn_end, ex_warnings = warnings_to_list(df)
    warnings_to_add = []

    # Loop through existing warnings and remove if plate ID is found
    for w in ex_warnings:

        # If na - continue
        if pd.isnull(w):
            continue

        # Ignore R35 warnings
        if 'R35' in w:
            continue

        # Get plate ID and see if it's in list
        # if plate ID from warning is not present in list - add to new list
        try:
            plate_id = w.split(':')[0].replace("Plate ", "")
            if plate_id not in p_ids:
                warnings_to_add.append(w)
        except AttributeError:
            continue

    # Remove all existing warnings
    df.loc[range(warn_start, warn_end), 'val'] = np.nan

    # If there are no warnings left - remove empty rows
    if not warnings_to_add:
        df.drop(range(warn_start + 1, warn_end), inplace=True)
    else:
        df = update_warnings(df, warnings_to_add, warn_start)

    return df

def import_f093_json(file_path):
    """ Import an existing F093 json file. """
    
    # Import the file
    df = pd.read_json(file_path + ".json", orient='table')

    # Rename dataframe columns
    df.columns = df.columns.str.replace(
            "PnC-IgG-ELISA type ", "Result_")
    
    return df

