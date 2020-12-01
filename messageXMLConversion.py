#Import modules
import openpyxl
from openpyxl import load_workbook
import tkinter
from tkinter import filedialog
from tkinter import Tk
import xml.etree.ElementTree as ET
import logging
import getpass
import os
import sys

#Define functions
def getXML(root, currdir):
    origXML = currdir + "/XML/sal_interfaces"
    root.xmlFilename = filedialog.askopenfilename(initialdir=origXML, title='Please select the GitHub XML file', filetypes=[("XML Files", "*.xml")])
    xmlFileLocation = root.xmlFilename
    if len(xmlFileLocation) > 0:
        #Temporarily set path to same directory as python file
        path, filename = os.path.split(xmlFileLocation)
        path = currdir
        fileBase = filename.split('.')[0]
        messageType = fileBase.split('_')[1]
        cscName = fileBase.split('_')[0]
        logger.info("You chose %s" % xmlFileLocation)
        getExcelFiles(path, fileBase, messageType, cscName, xmlFileLocation)
    else:
        logger.warning("You must select an XML file.")
        getXML(root, currdir)

def getExcelFiles(path, fileBase, messageType, cscName, xmlFileLocation):
    excelFileLocation_Message = path + "/" + cscName + "/" + fileBase + ".xlsx"
    excelFileLocation_Parameter = path + "/" + cscName + "/"  + fileBase + "_Properties.xlsx"
    excelFileLocation_Enumeration = path + "/" + cscName + "/"  + fileBase + "_Enumerations.xlsx"
    excelFileLocation_Literal = path + "/" + cscName + "/"  + fileBase + "_Enumerations_Options.xlsx"

    #Get message Excel file
    if os.path.isfile(excelFileLocation_Message):
        wb = load_workbook(filename = excelFileLocation_Message)

        #Verifiy message Excel ws is in correct format
        if 'Sheet1' in wb.sheetnames:
            ws = wb['Sheet1']


            if (ws['A1'].value == 'Name') and (ws['B1'].value == 'Alias') and (ws['C1'].value == 'Documentation') and (ws['D1'].value == 'Subsystem') and (ws['E1'].value == 'Version') and (ws['F1'].value == 'Author') and (ws['G1'].value == 'Device') and (ws['H1'].value == 'Property') and (ws['I1'].value == 'Action') and (ws['J1'].value == 'Value') and (ws['K1'].value == 'Explanation') and (ws['L1'].value == 'Order') and (ws['M1'].value == 'Element ID'):
                #Get message parameter Excel file
                if os.path.isfile(excelFileLocation_Parameter):
                    wb_p = load_workbook(filename = excelFileLocation_Parameter)

                    #Verifiy message parameter Excel ws is in correct format
                    if 'Sheet1' in wb_p.sheetnames:
                        ws_p = wb_p['Sheet1']

                        if (ws_p['A1'].value == 'Owner') and (ws_p['B1'].value == 'Name') and (ws_p['C1'].value == 'Documentation') and (ws_p['D1'].value == 'Type') and (ws_p['E1'].value == 'Size') and (ws_p['F1'].value == 'Type Modifier') and (ws_p['G1'].value == 'Multiplicity') and (ws_p['H1'].value == 'Order') and (ws_p['I1'].value == 'Element ID'):
                            #Get enumeration Excel file
                            if os.path.isfile(excelFileLocation_Enumeration):
                                wb_e = load_workbook(filename = excelFileLocation_Enumeration)

                                #Verifiy enumeration Excel ws is in correct format
                                if 'Sheet1' in wb_e.sheetnames:
                                    ws_e = wb_e['Sheet1']

                                    if (ws_e['A1'].value == 'Name') and (ws_e['B1'].value == 'Order') and (ws_e['C1'].value == 'Element ID'):
                                        #Get enumeration literal Excel file
                                        if os.path.isfile(excelFileLocation_Literal):
                                            wb_l = load_workbook(filename = excelFileLocation_Literal)

                                            #Verifiy enumeration literal Excel ws is in correct format
                                            if 'Sheet1' in wb_l.sheetnames:
                                                ws_l = wb_l['Sheet1']

                                                if (ws_l['A1'].value == 'Owner') and (ws_l['B1'].value == 'Name') and (ws_l['C1'].value == 'Order') and (ws_l['D1'].value == 'Element ID'):
                                                    updateExcelFiles(xmlFileLocation, excelFileLocation_Message, excelFileLocation_Parameter, excelFileLocation_Enumeration, excelFileLocation_Literal, wb, wb_p, wb_e, wb_l, ws, ws_p, ws_e, ws_l, messageType)
                                                else:
                                                    logger.error("The ENUMERATION OPTIONS Excel file columns do not follow the required format of [Owner, Name, Order]. Please fix the format and run again.")
                                                    sys.exit()
                                            else:
                                                logger.error("Could not find 'Sheet1' in ENUMERATION OPTIONS Excel file.")
                                                sys.exit()
                                        else:
                                            logger.error("Could not find the Excel file for the ENUMERATION OPTIONS at the following location: %s." % excelFileLocation_Enumeration)
                                            sys.exit()
                                    else:
                                        logger.error("The ENUMERATION Excel file columns do not follow the required format of [Name, Order]. Please fix the format and run again.")
                                        sys.exit()
                                else:
                                    logger.error("Could not find 'Sheet1' in ENUMERATION Excel file.")
                                    sys.exit()
                            else:
                                logger.error("Could not find the Excel file for the ENUMERATION at the following location: %s." % excelFileLocation_Enumeration)
                                sys.exit()
                        else:
                            logger.error("The ENUMERATION Excel file columns do not follow the required format of [Owner, Name, Documentation, Type, Size, Type Modifier, Multiplicity, Order]. Please fix the format and run again.")
                            sys.exit()
                    else:
                        logger.error("Could not find 'Sheet1' in MESSAGE PARAMETER Excel file.")
                        sys.exit()
                else:
                    logger.error("Could not find the Excel file for the MESSAGE PARAMETERS at the following location: %s." % excelFileLocation_Parameter)
                    sys.exit()
            else:
                logger.error("The MESSAGE Excel file columns do not follow the required format of [Name, Alias, Documentation, Subsystem, Version, Author, Device, Property, Action, Value, Explanation, Order]. Please fix the format and run again.")
                sys.exit()
        else:
            logger.error("Could not find 'Sheet1' in MESSAGE Excel file.")
            sys.exit()
    else:
        logger.error("Could not find the Excel file for the MESSAGE at the following location: %s." % excelFileLocation_Message)
        sys.exit()

def updateExcelFiles(xmlFileLocation, excelFileLocation_Message, excelFileLocation_Parameter, excelFileLocation_Enumeration, excelFileLocation_Literal, wb, wb_p, wb_e, wb_l, ws, ws_p, ws_e, ws_l, messageType):
    tree = ET.parse(xmlFileLocation)
    treeRoot = tree.getroot()

    salMessages = []
    salAlias = []
    messageParams = []
    messageTypeEnums = []
    enumLiterals = []
    messageOrder = 0
    messagePropOrder = 0
    enumOrder = 0
    enumLitOrder = 0

   #Clearing the SyncAction Column in Signals
    ws.delete_cols(14)
    ws.insert_cols(14)
    ws.cell(row=1, column =14, value = 'SyncAction')
    logger.info("Clearing the SyncAction column")

    #Clearing the SyncAction Column in signal properties
    ws_p.delete_cols(10)
    ws_p.insert_cols(10)
    ws_p.cell(row=1, column =10, value = 'SyncAction')
    logger.info("Clearing the SyncAction column")

    #Clearing the SyncAction Column in enumerations
    ws_e.delete_cols(4)
    ws_e.insert_cols(4)
    ws_e.cell(row=1, column =4, value = 'SyncAction')
    logger.info("Clearing the SyncAction column")

    #Clearing the SyncAction Column in enemuration literals
    ws_l.delete_cols(5)
    ws_l.insert_cols(5)
    ws_l.cell(row=1, column =5, value = 'SyncAction')
    logger.info("Clearing the SyncAction column")


    if messageType == "Commands":
        messageRoot = "SALCommand"
    elif messageType == "Events":
        messageRoot = "SALEvent"
    elif messageType == "Telemetry":
        messageRoot = "SALTelemetry"
    else:
        logger.error("Unknown message type of %s." % messageType)

    #Check for enumerations
    for enum in treeRoot.findall('Enumeration'):
        enumOrder = enumOrder + 1
        count_e = 0
        enumNm = enum.text.split('_')[0]
        enumName = enumNm.lstrip()
        enumNameLower = enumName[0].lower() + enumName[1:]
        messageTypeEnums.append(enumNameLower)

        #Check if existing enumeration row is found
        for row in ws_e.iter_rows(min_row=2, max_col=1):
            for cell in row:
                if cell.value == enumNameLower:
                    count_e = count_e + 1
                    thisRow_e = cell.row
                    ws_e.cell(row=thisRow_e, column=2, value=str(enumOrder))
                    ws_e.cell(row=thisRow_e, column=4, value="Update")

        #If no existing enumeration row found, create a new row
        if count_e == 0:
            filteredNumRows_e = list(filter(None, ws_e['A']))
            totalRows_e = len(filteredNumRows_e)
            newRow_e = totalRows_e + 1
            ws_e.cell(row=newRow_e, column=1, value=enumNameLower)
            ws_e.cell(row=newRow_e, column=2, value=str(enumOrder))
            ws_e.cell(row=newRow_e, column=4, value="Add")
        elif count_e > 1:
            logger.error("There are duplicate rows for enumeration {%s}" % enumNameLower)

        #Now start on literals
        enumLits = enum.text.replace(" ", "").split(',')
        for lit in enumLits:
            enumLitOrder = enumLitOrder + 1
            count_l = 0
            litOwn = lit.split('_')[0]
            litOwner = litOwn.lstrip()
            litOwnerLower = litOwner[0].lower() + litOwner[1:]
            lst = lit.lstrip()
            litName = lst.split('_',1)[1]
            litTup = (litOwnerLower, litName)
            enumLiterals.append(litTup)

            #Update literal row if existing literal is found
            for row in ws_l.iter_rows(min_row=2, max_col=2):
                rowList = []
                for cell in row:
                    rowList.append(cell.value)
                    thisRow = cell.row
                tupRow = tuple(rowList)

                if tupRow == litTup:
                    count_l = count_l + 1
                    ws_l.cell(row=thisRow, column=3, value=str(enumLitOrder))
                    ws_l.cell(row=thisRow, column=5, value="Update")


            #If no existing literal row found, create a new row
            if count_l == 0:
                filteredNumRows_l = list(filter(None, ws_l['A']))
                totalRows_l = len(filteredNumRows_l)
                newRow_l = totalRows_l + 1
                ws_l.cell(row=newRow_l, column=1, value=litOwnerLower)
                ws_l.cell(row=newRow_l, column=2, value=litName)
                ws_l.cell(row=newRow_l, column=3, value=str(enumLitOrder))
                ws_l.cell(row=newRow_l, column=5, value="Add")
            elif count_l > 1:
                errorMessage = litTup
                logger.error("There are duplicate rows for enumeration option {%s}" % (errorMessage,))
            del litTup



    #Remove any enumeration rows not in XML
    i=2
    while i <= ws_e.max_row:
        cellVal = ws_e.cell(row=i, column=1).value
        if cellVal == None:
            ws_e.delete_rows(i)
            continue
        elif (cellVal not in messageTypeEnums) :
            ws_e.cell(row=i, column=4, value="Delete")
            logger.info("The row for enumeration with the name {%s} is not found in the XML and has been marked for deletion." % cellVal)
        i += 1

    #Save enumeration file
    wb_e.save(excelFileLocation_Enumeration)

    #Remove any literal rows not in XML
    i=2
    while i <= ws_l.max_row:
        tupVal = (ws_l.cell(row=i, column=1).value, ws_l.cell(row=i, column=2).value)
        if (ws_l.cell(row=i, column=1).value == None) and (ws_l.cell(row=i, column=2).value == None) :
            ws_l.delete_rows(i)
            continue
        elif (tupVal not in enumLiterals) :
            ws_l.cell(row=i, column=5, value="Delete")
            logger.info("The row for enumeration literal with the name {%s} is not found in the XML and has been marked for deletion." % (tupVal,))
        i += 1


    #Save literals file
    wb_l.save(excelFileLocation_Literal)

    for member in treeRoot.findall(messageRoot):
        messageOrder = messageOrder + 1
        count = 0
        alias = ""
        topic = ""
        subsystem = ""
        version = ""
        author = ""
        explanation = ""
        items = ""
        device = ""
        prop = ""
        action = ""
        value = ""
        desc = ""

        #Get values
        if member.find('Alias') != None:
            alias = member.find('Alias').text
        else:
            if member.find('EFDB_Topic') != None:
                alias = member.find('EFDB_Topic').text.split('_')[-1]
            else:
                logger.error("The message has no Alias or EFDB_Topic.")
                sys.exit()
        if member.find('EFDB_Topic') != None:
            topic = member.find('EFDB_Topic').text
        else:
            logger.error("The message has no EFDB_Topic.")
            sys.exit()
        if member.find('Subsystem') != None:
            subsystem = member.find('Subsystem').text
        if member.find('Version') != None:
            version = member.find('Version').text
        if member.find('Author') != None:
            author = member.find('Author').text
        if member.find('Explanation') != None:
            explanation = member.find('Explanation').text
        if member.find('item') != None:
            items = member.findall('item')
        if member.find('Device') != None:
            device = member.find('Device').text
        if member.find('Property') != None:
            prop = member.find('Property').text
        if member.find('Action') != None:
            action = member.find('Action').text
        if member.find('Value') != None:
            value = member.find('Value').text
        if member.find('Description') != None:
            desc = member.find('Description').text

        #Add message topic and alias to lists
        salMessages.append(topic)
        salAlias.append(alias)


        #Update message row if existing message is found
        for row in ws.iter_rows(min_row=2, max_col=1):
            for cell in row:
                if cell.value == topic:
                    count += 1
                    thisRow = cell.row
                    ws.cell(row=thisRow, column=2, value=alias)
                    ws.cell(row=thisRow, column=3, value=desc)
                    ws.cell(row=thisRow, column=4, value=subsystem)
                    ws.cell(row=thisRow, column=5, value=version)
                    ws.cell(row=thisRow, column=6, value=author)
                    ws.cell(row=thisRow, column=7, value=device)
                    ws.cell(row=thisRow, column=8, value=prop)
                    ws.cell(row=thisRow, column=9, value=action)
                    ws.cell(row=thisRow, column=10, value=value)
                    ws.cell(row=thisRow, column=11, value=explanation)
                    ws.cell(row=thisRow, column=12, value=str(messageOrder))
                    ws.cell(row=thisRow, column=14, value="Update")

        #If no existing message row found, create a new row
        if count == 0:
            filteredNumRows = list(filter(None, ws['A']))
            totalRows = len(filteredNumRows)
            newRow = totalRows + 1
            ws.cell(row=newRow, column=1, value=topic)
            ws.cell(row=newRow, column=2, value=alias)
            ws.cell(row=newRow, column=3, value=desc)
            ws.cell(row=newRow, column=4, value=subsystem)
            ws.cell(row=newRow, column=5, value=version)
            ws.cell(row=newRow, column=6, value=author)
            ws.cell(row=newRow, column=7, value=device)
            ws.cell(row=newRow, column=8, value=prop)
            ws.cell(row=newRow, column=9, value=action)
            ws.cell(row=newRow, column=10, value=value)
            ws.cell(row=newRow, column=11, value=explanation)
            ws.cell(row=newRow, column=12, value=(messageOrder))
            ws.cell(row=newRow, column=14, value="Add")
        elif count > 1:
            logger.error("There are duplicate rows for message {%s}" % alias)

        #Now start on the parameters
        for item in items:
            messagePropOrder = messagePropOrder + 1
            count_p = 0
            param_name = ""
            param_desc = ""
            idl_type = ""
            idl_size = ""
            unit = ""
            param_count = ""


            #Get values
            if item.find('EFDB_Name') != None:
                param_name = item.find('EFDB_Name').text
            else:
                logger.error("The message has no EFDB_Name.")
                sys.exit()
            if item.find('Description') != None:
                param_desc = item.find('Description').text
            if item.find('IDL_Type') != None:
                if alias in messageTypeEnums:
                    idl_type = alias
                elif param_name in messageTypeEnums:
                    idl_type = param_name
                else:
                    idl_type = item.find('IDL_Type').text
                    if idl_type == "string":
                        idl_type = idl_type[0].upper() + idl_type[1:]
            if item.find('IDL_Size') != None:
                idl_size = int(item.find('IDL_Size').text)
            if item.find('Units') != None:
                unit = item.find('Units').text
            if item.find('Count') != None:
                param_count = item.find('Count').text

            #Create tuple of owner and name
            tup = (topic, param_name)

            #Add param tuple to list
            messageParams.append(tup)

            #Update parameter row if existing message is found
            for row in ws_p.iter_rows(min_row=2, max_col=2):
                rowList = []
                for cell in row:
                    rowList.append(cell.value)
                    thisRow = cell.row
                tupRow = tuple(rowList)

                if tupRow == tup:
                    count_p += 1
                    ws_p.cell(row=thisRow, column=3, value=param_desc)
                    ws_p.cell(row=thisRow, column=4, value=idl_type)
                    ws_p.cell(row=thisRow, column=5, value=str(idl_size))
                    ws_p.cell(row=thisRow, column=6, value=unit)
                    ws_p.cell(row=thisRow, column=7, value=param_count)
                    ws_p.cell(row=thisRow, column=8, value=str(messagePropOrder))
                    ws_p.cell(row=thisRow, column=10, value="Update")

            #If no existing parameter row found, create a new row
            if count_p == 0:
                filteredNumRows_p = list(filter(None, ws_p['A']))
                totalRows_p = len(filteredNumRows_p)
                newRow_p = totalRows_p + 1
                ws_p.cell(row=newRow_p, column=1, value=topic)
                ws_p.cell(row=newRow_p, column=2, value=param_name)
                ws_p.cell(row=newRow_p, column=3, value=param_desc)
                ws_p.cell(row=newRow_p, column=4, value=idl_type)
                ws_p.cell(row=newRow_p, column=5, value=str(idl_size))
                ws_p.cell(row=newRow_p, column=6, value=unit)
                ws_p.cell(row=newRow_p, column=7, value=param_count)
                ws_p.cell(row=newRow_p, column=8, value=str(messagePropOrder))
                ws_p.cell(row=newRow_p, column=10, value="Add")

            elif count_p > 1:
                errorMessage = tup
                logger.error("There are duplicate rows for message parameters {%s}" % (errorMessage,))
            del tup

    #Remove any message rows not in XML
    i=2
    while i <= ws.max_row:
        cellVal = ws.cell(row=i, column=1).value
        if cellVal == None :
           ws.delete_rows(i)
           continue
        elif (cellVal not in salMessages) :
            ws.cell(row=i, column=14, value="Delete")
            logger.info("The row for message with the name {%s} is not found in the XML and has been marked for deletion." % cellVal)
        i += 1

    #Save messages file
    wb.save(excelFileLocation_Message)

    #Remove any parameter rows not in XML
    i=2
    while i <= ws_p.max_row:
        tupVal = (ws_p.cell(row=i, column=1).value, ws_p.cell(row=i, column=2).value)
        if (ws_p.cell(row=i, column=1).value == None) and (ws_p.cell(row=i, column=2).value == None) :
            ws.delete_rows(i)
            continue
        elif (tupVal not in messageParams) :
            ws_p.cell(row=i, column=10, value="Delete")
            logger.info("The row for message parameter with {%s} is not found in the XML and has been marked for deletion." % (tupVal,))
        i += 1

    #Save parameters file
    wb_p.save(excelFileLocation_Parameter)

    #Check for unused enumerations
    for e in messageTypeEnums:
        if e not in salAlias:
            logger.error("The enumeration {%s} is defined but not found as any message alias." % e)


#Define global variables
root = Tk()
root.withdraw() #use to hide tkinter window

currdir = os.getcwd()
xmlFileLocation = ""
excelFileLocation_Message = ""
excelFileLocation_Parameter = ""
excelFileLocation_Enumeration = ""
excelFileLocation_Literal = ""

#Setup logging
user = getpass.getuser()

logger = logging.getLogger('LSST XML Message Converter')

file_log_handler = logging.FileHandler('logfile.log')
logger.addHandler(file_log_handler)

stderr_log_handler = logging.StreamHandler()
logger.addHandler(stderr_log_handler)

class ContextFilter(logging.Filter):
    """
    This is a filter which injects contextual information into the log.
    """
    def filter(self, record):
        record.user = user
        return True

logger.addFilter(ContextFilter())

#Format logging output
formatter = logging.Formatter('%(asctime)s - %(name)s - %(user)s - %(levelname)s - %(message)s')
file_log_handler.setFormatter(formatter)
stderr_log_handler.setFormatter(formatter)

#Set the log level
logger.setLevel('DEBUG')

#Call initial function
getXML(root, currdir)
