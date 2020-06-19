# Important to PIP install openpyxl to be able to import it
import os
import sys
import shutil
import re
import csv
import openpyxl
from openpyxl.styles import NamedStyle, Font, Border, Side, PatternFill, Alignment

import shutil # DELTE WHEN DELETED FILE REPOSITORY IS NO LONGER NEEDED

###################################################################################################
##############     Get the project directory, project number, and all assemblies     ##############
###################################################################################################

print('Documentation Automator v0.2.2\n')
# The number of assemblies found
assembly_count = 0
# All the assemblies in the project
assemblies = []

# Get the path of the Altium project from user
project_path = input('Enter the full path of your Altium project:\n')

# Strip quotes if present from copying path in windows
if project_path[0] == '"' and project_path[-1] == '"':
    project_path = project_path.strip('"')

# Verify that it is an Altium project file given ie ending in .PrjPcb
if not re.search('\.PrjPcb$', project_path, re.IGNORECASE):
    input('File linked to is not an Altium project. Press enter to exit')
    exit()

# Verify the file is at the location of the user input
if not os.path.isfile(project_path):
    input('Project not found. Press enter to exit')
    exit()

# Get the board number. Take the Altium project path, rsplit the last \ and take the second string of the 2 new ones
# Then rsplit again by underscore to come up with board name in the end.
pcb_number = project_path.rsplit('\\', 1)[1].rsplit('_', 1)[0]

# Verifies that the board number matches the template of 1234B4657A
if not re.search('^\d\d\d\dB46\d\d[A-Z]$', pcb_number):
    input('Error determining the project number, project number found ', + pcb_number + '. Press enter to exit')
    exit()
    
# Get the project Directory. Take the Altium project path, rsplit the last \ and take the first string of the 2 new ones
project_dir = project_path.rsplit('\\', 1)[0]
# Store the original working directory to restore later
owd = os.getcwd()
# Change the current working directory to that of the project. Make the \ a \\ to allow python to take the slashes literally.
os.chdir(project_dir.replace('\\', '\\\\'))

# Get assemblies from project file
# Open the project, read its lines, and if it find [ProjectVariantX] verify 2 lines ahead is the description and store its value
with open(project_path, 'r', encoding = 'UTF-8') as project_file:
    lines = project_file.readlines()
    for i in range(0, len(lines)):
        line = lines[i]
        if line == f"[ProjectVariant{assembly_count + 1}]\n":
            if lines[i+2].split('=')[0] == 'Description':
                assembly_count += 1
                assemblies.append(lines[i+2].split('=')[1].rstrip())

print('\nProject Information')
print(  '*******************')
print(str(assembly_count) + ' Assemblie(s) Found In Project File:')
for i in assemblies: print(i)
print()

###################################################################################################
####################     Declare arrays and dictionaries for needed files     #####################
###################################################################################################

# List of file we are going to be manipulating. Will allow ability for user to choose what to keep in future updates
###################################################################################################

# Assembly BOMs found
assembly_boms = []
# SAP BOMs found
sap_boms = []
# Aegis Sync BOMs found
aegis_boms = []

# Lists of files that we want to keep with known names. Allows easy way to adjust what should be kept
###################################################################################################

# List of files for the reports folder that we want to keep
reports_keep = ['^' + pcb_number + '[ ,_]Order[ ,_]Information.xls(x)?$',
                '^' + pcb_number + '[ ,_]Build[ ,_]Request.doc(x)?$',
                '^' + pcb_number + '[ ,_]EE[ ,_]Review.xls(x)?$'
                ]

# List of files for Source folder we want to keep
source_keep = ['^Assy[ ,_]' + pcb_number + '_v[0-9]+.PCBDwf$',
                '^PCB[ ,_]' + pcb_number + '_v[0-9]+.PcbDoc$',
                '.SchDoc$'
                ]

# List of Mfg-Data files to keep, adds aegis sync later depending on if excel or text needed.
mfgdata_keep = ['^ODB[ ,_]' + pcb_number + '.zip$'
                 ]

# Left empty, filled in when Gerbers are found. Will allow ability for user to choose what to keep in future updates
gerbers_keep = []

# List of files / folders for CAM folder that are valid
cam_keep = ['^Gerber and Drill$',
              '^' + pcb_number + '_[A-Z][0-9]+.pdf$',
              '^' + pcb_number + '_[A-Z][0-9]+_(RoHS)?(R)?(FLEX)?(PILLAR)?(MCPCB)?(VIPPO)?.zip$',
              '^Spec[ ,_]' + pcb_number + '_[A-Z][0-9]+.dwg$',
             ]

# Dictionary of Altiums Gerber layer file extensions
gerber_ext = {  pcb_number + '.G[0-9]+' : '   -  Mid Layer ',              # .G(int) =  internal layer (int)
              pcb_number + '.GBL' : '  -  Bottom Layer',
              pcb_number + '.GBO' : '  -  Bottom Overlay',
              pcb_number + '.GBP' : '  -  Bottom Paste Mask',
              pcb_number + '.GBS' : '  -  Bottom Solder Mask',
#              pcb_number + '.GD[0-9]+' : '   -  Drill Drawing ',        # .GD(int) = Drill Drawing (int)
#              pcb_number + '.GG[0-9]+' : '   -  Drill Guide ',          # .GG(int) = Drill Guide (int)
#             pcb_number + '.GKO' : '  -  Keep Out Layer',
#              pcb_number + '.GM[0-9]+' : '  -  Mechanical Layer ',      # .GM(int) = Mechanical Layer (int)
               pcb_number + '.GP[0-9]+' : '  -  Internal Plane Layer ', # .GP(int) = Internal Plane Layer (int)
#             pcb_number + '.GPB' : '  -  Pad Master Bottom',
#             pcb_number + '.GPT' : '  -  Pad Master Top',
              pcb_number + '.GTL' : '  -  Top Layer',
              pcb_number + '.GTO' : '  -  Top Overlay',
              pcb_number + '.GTP' : '  -  Top Paste Mask',
              pcb_number + '.GTS' : '  -  Top Solder Mask',
#              pcb_number + '.P[0-9]+' : '  -  Gerber Panels ',        # .P0(int) = Ger Panel (int)
              pcb_number + '.DRR' : '  -  NC Drill Report',
              pcb_number + '.TXT' : '  -  Drill File',
              pcb_number + '-SlotHoles.TXT' : '  -  Slot Drill File',
              pcb_number + '-RoundHoles.TXT' : '  -  Hole Drill File'
              }

###################################################################################################
######################     Find all the relavent project info and files     #######################
###################################################################################################

# Find the assembly BOMs
###################################################################################################

print('Assembly BOMs:')
# For each assembly, see if the its BOM is in the reports folder
for assembly in assemblies:
    found = False
    print(assembly, end = '')
    for file in os.listdir('.\\Reports'):
        if re.search(assembly + '[ ,_]Assembly[ ,_]BOM.xls(x)?', file, re.IGNORECASE):
            found = True
            assembly_boms.append(file)
            reports_keep.append(file)
            print(' Assembly BOM Found')
    if found == False:
        print(' MISSING')
print()

# Find the SAP BOMs
###################################################################################################

print('SAP BOMs:')
# For each assembly, see if the its BOM is in the reports folder
for assembly in assemblies:
    found = False
    print(assembly, end = '')
    for file in os.listdir('.\\Reports'):
        if re.search(assembly + '[ ,_]SAP[ ,_]Import[ ,_]File.xls(x)?', file, re.IGNORECASE):
            found = True
            sap_boms.append(file)
            reports_keep.append(file)
            print(' SAP Import File Found')
    if found == False:
        print(' MISSING')
print()

# Find Gerber layer data from Gerber and Drill folder
###################################################################################################

print('Gerber Files Found:')
# For all files/folders in the Gerber and Drill folder, if it is a file continue
for file in [file for file in os.listdir('..\\Cam\\Gerber and Drill\\') if os.path.isfile('..\\Cam\\Gerber and Drill\\' + file)]:
    # For the key and values in the Gerber dictionary
    for gerber, label in gerber_ext.items():
        # If that file matches the key, keep it and print the label (value)
        if re.search(gerber, file, re.IGNORECASE):
            gerbers_keep.append(file)
            print(file + label, end = '')
            # If its a numbered gerber, print what number to the user after its description
            if file[-1].isdigit():
                # Strip everything except the number from the key, then strip that from the file name 
                print(file.strip(gerber.strip('[0-9]+')))
            else:
                # Prints end line that was withheld until number check
                print()
print()

# Find the Aegis Sync BOMs
###################################################################################################

print('Aegis Sync Excel BOMs:')
# For each assembly, see if the its BOM is in the reports folder
for assembly in assemblies:
    excel = ''
    text = ''
    # Find if the excel and text files exist, if so remember them for later
    for file in os.listdir('..\\Mfg-Data\\'):
        if re.search('Aegis[ ,_]Sync[ ,_]' + assembly + '.xls(x)?', file, re.IGNORECASE):
            excel = file
        elif re.search('Aegis[ ,_]Sync[ ,_]' + assembly + '.txt', file, re.IGNORECASE):
            text = file
            
    # If both txt and excel files are there, see which one is newer
    if excel and text:
        # If the excel is newer, alert user and save only excel to be saved
        if os.path.getmtime('..\\Mfg-Data\\' + excel) >= os.path.getmtime('..\\Mfg-Data\\' + text):
            print(text + ' File Found - OUT OF DATE')
            print(excel + ' File Found')
            aegis_boms.append(excel)
            mfgdata_keep.append(excel)
        # If the text is newer, alert user and save only excel to be saved
        else:
            print(text + ' File Found')
            print(text + ' File Found - Marked For Deletion')
            mfgdata_keep.append(text)
    # If just the excel file is there add it to be saved
    elif os.path.isfile('..\\Mfg-Data\\' + excel):
        print(excel + ' File Found')
        aegis_boms.append(excel)
        mfgdata_keep.append(excel)
    # Just the text BOM was found
    elif os.path.isfile('..\\Mfg-Data\\' + text):
        print(text + ' File Found - Excel missing')
        mfgdata_keep.append(text)
    else:
        print(assembly + ' MISSING')

###################################################################################################
#################################     Sort the assembly BOMs     ##################################
###################################################################################################

print('\n\nSorting Assembly BOMs')
print('*********************')

# Ask the user if they would like to sort all assembly BOM
response_all = ''
if len(assembly_boms) > 1:
    while response_all not in ('y', 'yes', 'n', 'no'):
        response_all = input('Would you like to sort all assembly BOMs?: ').lower()
        print()
else:
    response_all = 'no'

# Do it for all assembly BOMs found
for assembly_bom in assembly_boms:

    # Ask the user if they would like to sort each assembly BOM if they didnt want to do all
    response2 = ''
    if re.search('^n(o)?$', response_all, re.IGNORECASE):
            while response2 not in ('y', 'yes', 'n', 'no'):
                response2 = input(f'Would you like to sort {assembly_bom}?: ').lower()
    if re.search('^n(o)?$', response2, re.IGNORECASE) and re.search('^n(o)?$', response_all, re.IGNORECASE):
        print(assembly_bom + ' Skipped\n')
        continue
    
    print('Sorting ' + assembly_bom + '...')
    # Open the excel file and go in the first sheet
    wb = openpyxl.load_workbook('.\\Reports\\' + assembly_bom)
    sheet = wb['Sheet1']

    # Set up fonts
    # Font for component section declaration 
    comp_row_font  = NamedStyle(name = 'comp_row_font')
    comp_row_font.font = Font(name = 'Verdana', bold = True, size = 14)
    # Font for the section headers
    comp_header_font  = NamedStyle(name = 'comp_header_font')
    comp_header_font.font = Font(name = 'Verdana', bold = True, size = 12, underline = 'single')
    # Font for default lines
    default_font  = NamedStyle(name = 'default_font')
    default_font.font = Font(name = 'Verdana', size = 12)
    default_font.alignment = Alignment(horizontal = 'left')
    # Something to tell when the BOM starts
    header_row = 0
    # A list that will contain dictionaries of the parts
    bom_content = []

# Read the cells
###################################################################

    # Scan through each row in the BOM excel from 1 to end
    for row in range(1, sheet.max_row + 1):
        # Column 1 is part number, 2 is quantity etc. Store the values in these variables
        part_number = sheet.cell(row, 1).value
        quantity    = sheet.cell(row, 2).value
        description = sheet.cell(row, 3).value
        designator  = sheet.cell(row, 4).value
        layer       = sheet.cell(row, 5).value
        fitted      = sheet.cell(row, 6).value
        # If its previously found the header, and the part number is not None, add the part
        if header_row != 0 and part_number is not None:
            # Handle layer None components, ask for input until valid response
            if layer == 'None':
                while 1:
                    response = input(f'{part_number} | {description} | is layer None, does it belong on top or bottom side?: ')
                    if re.search('^t(op)?( )?(side)?$', response, re.IGNORECASE):
                        layer = 'Top'
                        break
                    if re.search('^b(ottom)?( )?(side)?$', response, re.IGNORECASE):
                         layer = 'Bottom'
                         break
            # Add the part as a dictionary to the list of parts. Make them strings to allow sorting to work.
            bom_content.append({'part_number': str(part_number),
                                'quantity'   : quantity,
                                'description': str(description),
                                'designator' : str(designator),
                                'layer'      : str(layer),
                                'fitted'     : str(fitted)})
        # If the row equals the header Altium uses, singal that the header has been found
        if str(part_number) + str(quantity) + str(description) + str(designator) + str(layer) + str(fitted) == 'LibRefQuantityDescriptionDesignatorLayerFitted':
            header_row = row
    # If the header was not found, it was probably already messed with. Alert the user
    if header_row == 0:
        print(f'No Altium header row found for {assembly_bom}. Unable to sort...\n')
        continue
    
    # Sort the BOMs, Least priority to most. Sorting doesnt change order of those who match sort key
    bom_content.sort(key = lambda k: k['part_number'])
    bom_content.sort(key = lambda k: k['fitted'])
    bom_content.sort(key = lambda k: k['layer'], reverse = True)

    # Delete the old rows of data 
    sheet.delete_rows(header_row, sheet.max_row)
    
# Write the data to the cells
###################################################################

    current_section = 'None'
    pattern_fill = False
    # Start printing 14 rows up from where the header row was found
    row = header_row - 14
    for part in bom_content:
        # Check if you are in a new section of the BOM, if so print the header
        if current_section != part['layer'] + part['fitted']:
            row += 2
            current_section = part['layer'] + part['fitted']
            # Determine what section of BOM is starting
            if part['layer'] == 'Top': header_string = 'TOP SIDE COMPONENTS'
            elif part['layer'] == 'Bottom': header_string = 'BOTTOM SIDE COMPONENTS'
            else: header_string = 'None Side Components'
            if part['fitted'] == 'Not Fitted': header_string += ' (not installed)'
            # Print the component section row and format it
            sheet.cell(row = row, column = 1).value = header_string
            sheet['A' + str(row)].style = comp_row_font
            # Print the component headers and format it
            sheet.cell(row = row + 1, column = 1).value = 'PART#'
            sheet['A' + str(row + 1)].style = comp_header_font
            sheet.cell(row = row + 1, column = 2).value = 'QTY'
            sheet['B' + str(row + 1)].style = comp_header_font
            sheet.cell(row = row + 1, column = 3).value = 'DESCRIPTION'
            sheet['C' + str(row + 1)].style = comp_header_font
            sheet.cell(row = row + 1, column = 4).value = 'SYMBOL'
            sheet['D' + str(row + 1)].style = comp_header_font
            row += 2
            pattern_fill = False
        # Print the row data for the component, give them default 
        sheet.cell(row = row, column = 1).value = part['part_number']
        sheet['A' + str(row)].style = default_font
        if pattern_fill == True: sheet['A' + str(row)].fill = PatternFill('solid', fgColor='D9D9D9')
        sheet.cell(row = row, column = 2).value = str(part['quantity'])
        sheet['B' + str(row)].style = default_font
        if pattern_fill == True: sheet['B' + str(row)].fill = PatternFill('solid', fgColor='D9D9D9')
        sheet.cell(row = row, column = 3).value = part['description']
        sheet['C' + str(row)].style = default_font
        if pattern_fill == True: sheet['C' + str(row)].fill = PatternFill('solid', fgColor='D9D9D9')
        sheet.cell(row = row, column = 4).value = part['designator']  
        sheet['D' + str(row)].style = default_font
        if pattern_fill == True: sheet['D' + str(row)].fill = PatternFill('solid', fgColor='D9D9D9')
        sheet.cell(row = row, column = 5).value = part['layer']
        sheet['E' + str(row)].style = default_font
        sheet.cell(row = row, column = 6).value = part['fitted']
        sheet['F' + str(row)].style = default_font
        row += 1
        pattern_fill = not pattern_fill

    # Save the excel
    wb.save('.\\Reports\\' + assembly_bom)
    print(assembly_bom + ' sorting complete\n')

###################################################################################################
#################################     Clean SAP Import Files     ##################################
###################################################################################################

print('Cleaning SAP Import Files')
print('*************************')

# Ask the user if they would like to sort all assembly BOM
response = ''
if len(sap_boms) > 1:
    while response not in ('y', 'yes', 'n', 'no'):
        response = input('Would you like to clean all SAP Import files?: ').lower()
        print()
else:
    response = 'no'

# Do it for all assembly BOMs found
for sap_bom in sap_boms:

    # List to store the dictionary of data
    sap_content = []

    # Ask the user if they would like to sort each assembly BOM if they didnt want to do all
    response2 = ''
    if re.search('^n(o)?$', response, re.IGNORECASE):
            while response2 not in ('y', 'yes', 'n', 'no'):
                response2 = input(f'Would you like to clean {sap_bom}?: ').lower()
    if re.search('^n(o)?$', response2, re.IGNORECASE) and re.search('^n(o)?$', response, re.IGNORECASE):
        print(sap_bom + ' Skipped\n')
        continue
    
    print('Cleaning ' + sap_bom + '...')
    # Open the excel file and go in the first sheet
    wb = openpyxl.load_workbook('.\\Reports\\' + sap_bom)
    sheet = wb['Sheet1']
    
    # Read the data
    # Verify the header row is present
    if sheet.cell(1, 2).value != 'LibRef' or sheet.cell(1, 4).value != 'Quantity':
        print(f'No Altium header row found for {sap_bom}. Unable to sort...\n')
        continue
    # For each row of data
    for row in range(2, sheet.max_row + 1):
        # Read the data for the rows and store temporarily
        part_number    = sheet.cell(row, 2).value
        quantity       = sheet.cell(row, 4).value
        if part_number is not None:
            sap_content.append({'part_number' : str(part_number),
                                'quantity'    : quantity
                                })
    # Sort the BOM
    sap_content.sort(key = lambda k: k['part_number'])

    # Delete the old rows of data 
    sheet.delete_rows(1, sheet.max_row)
    
    # Write the data
    row = 1
    for part in sap_content:
        sheet.cell(row = row, column = 1).value = 'L'
        sheet.cell(row = row, column = 2).value = part['part_number']
        sheet.cell(row = row, column = 4).value = part['quantity']
        row += 1

    # Save the data
    wb.save('.\\Reports\\' + sap_bom)
    print(sap_bom + ' cleaning complete\n')
    
###################################################################################################
#################################     Tab Delimit Aegis Sync     ##################################
###################################################################################################

print('Tab Delimit Aegis Sync Files')
print('****************************') 

# Ask the user if they would like to sort all assembly BOM
response = ''
if len(aegis_boms) > 1:
    while response not in ('y', 'yes', 'n', 'no'):
        response = input('Would you like to create all Aegis sync files?: ').lower()
        print()
else:
    response = 'no'

# Do it for all assembly BOMs found
for excel in aegis_boms:

    # Ask the user if they would like to sort each assembly BOM if they didn't want to do all
    response2 = ''
    if re.search('^n(o)?$', response, re.IGNORECASE):
            while response2 not in ('y', 'yes', 'n', 'no'):
                response2 = input(f'Would you like to make text file for {excel}?: ').lower()
    if re.search('^n(o)?$', response2, re.IGNORECASE) and re.search('^n(o)?$', response, re.IGNORECASE):
        print(sap_bom + ' Skipped\n')
        continue

    # If the text file already exists, delete it
    if os.path.isfile('..\\Mfg-Data\\' + excel.rsplit('.', 1)[0] + '.txt'):
        try:
            os.remove('..\\Mfg-Data\\' + excel.rsplit('.', 1)[0] + '.txt')
        except OSError as e: 
            print ("Error: %s - %s." % (e.filename, e.strerror))
            
    # Open the excel file and go in the first sheet
    print('Generating text file for ' + excel)
    wb = openpyxl.load_workbook('..\\Mfg-Data\\' + excel)
    sheet = wb['Sheet1']
    # Save each row to text file, tab delimited
    with open('..\\Mfg-Data\\' + excel.rsplit('.', 1)[0] + '.txt', 'w', newline = '') as aegis_text:
        file = csv.writer(aegis_text, delimiter = '\t')
        for row in sheet.rows:
            file.writerow([cell.value for cell in row])

    # Finished with conversion, no need to save excel
    print(excel + ' converted to text\n\n')

    # If the excel is in mfgdata_keep, replace it with the text
    mfgdata_keep = [excel.rsplit('.', 1)[0] + '.txt' if x == (excel) else x for x in mfgdata_keep]

# Print if none edited
if not aegis_boms:
    print('No files to convert\n\n')    
    
###################################################################################################
###################################     Remove Junk Files     #####################################
###################################################################################################

print('***********************')
print('Deleting Unneeded Files')
print('***********************\n')

# Reports Folder
###################################################################################################

print('Reports Folder')
print('**************')

# Array to store files that it finds unneeded. Ensures whats deleted is what user saw
reports_unneeded = []

# Scan through each file in the reports folder
for report_file in os.listdir('.\\Reports\\'):
    valid = False
    # Scan through each file in valid reports list, if it matches the file found mark it valid, if its not valid add it to delete list
    for unneeded_file in [report_file for file in reports_keep if re.search(file, report_file, re.IGNORECASE)]:
        valid = True
    if valid == False:
        reports_unneeded.append(report_file)
        print(report_file)

# Ask if ok to delete all unneeded file from Reports
response = ''
if reports_unneeded:
    while response not in ('y', 'yes', 'n', 'no'):
        response = input('OK to delete all unneeded files from Reports?').lower()
    if re.search('^y(es)?$', response, re.IGNORECASE):
        # Delete all the files
        for unneeded_file in reports_unneeded:
            # Verify its a file and if so try to delete, if not report back
            if os.path.isfile('.\\Reports\\' + unneeded_file):
                try:
                    if not os.path.exists('..\\Deleted\\'): # DELTE WHEN DELETED FILE REPOSITORY IS NO LONGER NEEDED
                        os.mkdir('..\\Deleted\\') # DELTE WHEN DELETED FILE REPOSITORY IS NO LONGER NEEDED
                    shutil.move('.\\Reports\\' + unneeded_file, '..\\Deleted\\' + unneeded_file) # DELTE WHEN DELETED FILE REPOSITORY IS NO LONGER NEEDED
#                    os.remove('.\\Reports\\' + unneeded_file)
                except OSError as e: 
                    print ("Error: %s - %s." % (e.filename, e.strerror))
            else:
                try:
                    if not os.path.exists('..\\Deleted\\'): # DELTE WHEN DELETED FILE REPOSITORY IS NO LONGER NEEDED
                        os.mkdir('..\\Deleted\\') # DELTE WHEN DELETED FILE REPOSITORY IS NO LONGER NEEDED
                    shutil.move('.\\Reports\\' + unneeded_file, '..\\Deleted\\' + unneeded_file) # DELTE WHEN DELETED FILE REPOSITORY IS NO LONGER NEEDED
#                    shutil.rmtree('.\\Reports\\' + unneeded_file)
                except OSError as e:
                    print ("Error: %s - %s." % (e.filename, e.strerror))
print('Reports folder done\n')

# Source Folder
###################################################################################################

print('Source Folder')
print('*************')

# Array to store files that it finds unneeded. Ensures whats deleted is what user saw
source_unneeded = []

# Scan through each file in the source folder
for source_file in os.listdir('.\\Source\\'):
    valid = False
    # Scan through each file in valid source list, if it matches the file found mark it valid, if its not valid add it to delete list
    for unneeded_file in [source_file for file in source_keep if re.search(file, source_file, re.IGNORECASE)]:
        valid = True
    if valid == False:
        source_unneeded.append(source_file)
        print(source_file)

# Ask if OK to delete all unneeded file from Reports
response = ''
if source_unneeded:
    while response not in ('y', 'yes', 'n', 'no'):
        response = input('OK to delete all unneeded files from Source?').lower()
    if re.search('^y(es)?$', response, re.IGNORECASE):
        # Delete all the files
        for unneeded_file in source_unneeded:
            # Verify its a file and if so try to delete, if not report back
            if os.path.isfile('.\\Source\\' + unneeded_file):
                try:
                    if not os.path.exists('..\\Deleted\\'): # DELTE WHEN DELETED FILE REPOSITORY IS NO LONGER NEEDED
                        os.mkdir('..\\Deleted\\') # DELTE WHEN DELETED FILE REPOSITORY IS NO LONGER NEEDED
                    shutil.move('.\\Source\\' + unneeded_file, '..\\Deleted\\' + unneeded_file) # DELTE WHEN DELETED FILE REPOSITORY IS NO LONGER NEEDED
#                    os.remove('.\\Source\\' + unneeded_file)
                except OSError as e: 
                    print ("Error: %s - %s." % (e.filename, e.strerror))
            else:
                try:
                    if not os.path.exists('..\\Deleted\\'): # DELTE WHEN DELETED FILE REPOSITORY IS NO LONGER NEEDED
                        os.mkdir('..\\Deleted\\') # DELTE WHEN DELETED FILE REPOSITORY IS NO LONGER NEEDED
                    shutil.move('.\\Source\\' + unneeded_file, '..\\Deleted\\' + unneeded_file) # DELTE WHEN DELETED FILE REPOSITORY IS NO LONGER NEEDED
#                    shutil.rmtree('.\\Source\\' + unneeded_file)
                except OSError as e:
                    print ("Error: %s - %s." % (e.filename, e.strerror))
print('Source folder done\n')

# CAM Folder
###################################################################################################

print('CAM Folder')
print('**************')

# Array to store files that it finds unneeded. Ensures whats deleted is what user saw
cam_unneeded = []

# Scan through each file in the Cam folder
for cam_file in os.listdir('..\\Cam\\'):
    valid = False
    # Scan through each file in valid cam list, if it matches the file found mark it valid, if its not valid add it to delete list
    for unneeded_file in [cam_file for file in cam_keep if re.search(file, cam_file, re.IGNORECASE)]:
        valid = True
    if valid == False:
        cam_unneeded.append(cam_file)
        print(cam_file)

# Ask if OK to delete all unneeded file from Reports
response = ''
if cam_unneeded:
    while response not in ('y', 'yes', 'n', 'no'):
        response = input('OK to delete all unneeded files from Reports?').lower()
    if re.search('^y(es)?$', response, re.IGNORECASE):
        # Delete all the files
        for unneeded_file in cam_unneeded:
            # Verify its a file and if so try to delete, if not report back
            if os.path.isfile('..\\Cam\\' + unneeded_file):
                try:
                    if not os.path.exists('..\\Deleted\\'): # DELTE WHEN DELETED FILE REPOSITORY IS NO LONGER NEEDED
                        os.mkdir('..\\Deleted\\') # DELTE WHEN DELETED FILE REPOSITORY IS NO LONGER NEEDED
                    shutil.move('..\\Cam\\' + unneeded_file, '..\\Deleted\\' + unneeded_file) # DELTE WHEN DELETED FILE REPOSITORY IS NO LONGER NEEDED
#                    os.remove('..\\Cam\\' + unneeded_file)
                except OSError as e: 
                    print ("Error: %s - %s." % (e.filename, e.strerror))
            else:
                try:
                    if not os.path.exists('..\\Deleted\\'): # DELTE WHEN DELETED FILE REPOSITORY IS NO LONGER NEEDED
                        os.mkdir('..\\Deleted\\') # DELTE WHEN DELETED FILE REPOSITORY IS NO LONGER NEEDED
                    shutil.move('..\\Cam\\' + unneeded_file, '..\\Deleted\\' + unneeded_file) # DELTE WHEN DELETED FILE REPOSITORY IS NO LONGER NEEDED
#                    shutil.rmtree('..\\Cam\\' + unneeded_file)
                except OSError as e:
                    print ("Error: %s - %s." % (e.filename, e.strerror))
print('Cam folder done\n')

# Gerber and Drill Folder
###################################################################################################

print('Gerber and Drill Folder')
print('***********************')

# Array to store files that it finds unneeded. Ensures whats deleted is what user saw
gerber_unneeded = []

# Prints all the files in Reports folder that are not in the list to keep
for unneeded_file in [file for file in os.listdir('..\\Cam\\Gerber and Drill') if file not in gerbers_keep]:
    gerber_unneeded.append(unneeded_file)
    print(unneeded_file)

# Ask if OK to delete all unneeded file from Reports
response = ''
if gerber_unneeded:
    while response not in ('y', 'yes', 'n', 'no'):
        response = input('OK to delete all unneeded files from Gerber and Drill?').lower()
    if re.search('^y(es)?$', response, re.IGNORECASE):
        # Delete all the files
        for unneeded_file in gerber_unneeded:
            # Verify its a file and if so try to delete, if not report back
            if os.path.isfile('..\\Cam\\Gerber and Drill\\' + unneeded_file):
                try:
                    if not os.path.exists('..\\Deleted\\'): # DELTE WHEN DELETED FILE REPOSITORY IS NO LONGER NEEDED
                        os.mkdir('..\\Deleted\\') # DELTE WHEN DELETED FILE REPOSITORY IS NO LONGER NEEDED
                    shutil.move('..\\Cam\\Gerber and Drill\\' + unneeded_file, '..\\Deleted\\' + unneeded_file) # DELTE WHEN DELETED FILE REPOSITORY IS NO LONGER NEEDED
#                    os.remove('..\\Cam\\Gerber and Drill\\' + unneeded_file)
                except OSError as e: 
                    print ("Error: %s - %s." % (e.filename, e.strerror))
            else:
                try:
                    if not os.path.exists('..\\Deleted\\'): # DELTE WHEN DELETED FILE REPOSITORY IS NO LONGER NEEDED
                        os.mkdir('..\\Deleted\\') # DELTE WHEN DELETED FILE REPOSITORY IS NO LONGER NEEDED
                    shutil.move('..\\Cam\\Gerber and Drill\\' + unneeded_file, '..\\Deleted\\' + unneeded_file) # DELTE WHEN DELETED FILE REPOSITORY IS NO LONGER NEEDED
#                    shutil.rmtree('..\\Cam\\Gerber and Drill\\' + unneeded_file)
                except OSError as e:
                    print ("Error: %s - %s." % (e.filename, e.strerror))
print('Gerber and Drill folder done\n')

# Manufacturing and Data Folder
###################################################################################################

print('Mfg-Data Folder')
print('***************')

# Array to store files that it finds unneeded. Ensures whats deleted is what user saw
mfg_unneeded = []

# Scan through each file in the Mfg-Data folder
for mfgdata_file in os.listdir('..\\Mfg-Data\\'):
    valid = False
    # Scan through each file in valid source list, if it matches the file found mark it valid, if its not valid add it to delete list
    for unneeded_file in [mfgdata_file for file in mfgdata_keep if re.search(file, mfgdata_file, re.IGNORECASE)]:
        valid = True
    if valid == False:
        mfg_unneeded.append(mfgdata_file)
        print(mfgdata_file)
        

# Ask if ok to delete all unneeded file from Reports
response = ''
if mfg_unneeded:
    while response not in ('y', 'yes', 'n', 'no'):
        response = input('OK to delete all unneeded files from Mfg-Data?').lower()
    if re.search('^y(es)?$', response, re.IGNORECASE):
        # Delete all the files
        for unneeded_file in mfg_unneeded:
            # Verify its a file and if so try to delete, if not report back
            if os.path.isfile('..\\Mfg-Data\\' + unneeded_file):
                try:
                    if not os.path.exists('..\\Deleted\\'): # DELTE WHEN DELETED FILE REPOSITORY IS NO LONGER NEEDED
                        os.mkdir('..\\Deleted\\') # DELTE WHEN DELETED FILE REPOSITORY IS NO LONGER NEEDED
                    shutil.move('..\\Mfg-Data\\' + unneeded_file, '..\\Deleted\\' + unneeded_file) # DELTE WHEN DELETED FILE REPOSITORY IS NO LONGER NEEDED
#                    os.remove('..\\Mfg-Data\\' + unneeded_file)
                except OSError as e: 
                    print ("Error: %s - %s." % (e.filename, e.strerror))
            else:
                try:
                    if not os.path.exists('..\\Deleted\\'): # DELTE WHEN DELETED FILE REPOSITORY IS NO LONGER NEEDED
                        os.mkdir('..\\Deleted\\') # DELTE WHEN DELETED FILE REPOSITORY IS NO LONGER NEEDED
                    shutil.move('..\\Mfg-Data\\' + unneeded_file, '..\\Deleted\\' + unneeded_file) # DELTE WHEN DELETED FILE REPOSITORY IS NO LONGER NEEDED
#                    shutil.rmtree('..\\Mfg-Data\\' + unneeded_file)
                except OSError as e:
                    print ("Error: %s - %s." % (e.filename, e.strerror))
print('Mfg-Data folder done\n')

# Change working directory back to default to prevent program from preventing deleting project file
os.chdir(owd)
input('Documentation Cleanup Complete!! Press Enter To Exit')
