#Import modules
import lxml
from lxml import etree
import xmldiff
from xmldiff import main, formatting
import pandas as pd
import openpyxl
from openpyxl import load_workbook
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from tkinter import scrolledtext
from tkinter import messagebox
import xml.etree.ElementTree as ET
import logging
import getpass
import os
import sys
import subprocess
from pathlib import Path
import time
#Define functions
def needExport():
    response = input("Do you need to export the XML files? [Y/N]: ")
    #If exports are needed get login credentials
    if response in questionYes:
        logger.info("You chose to export the XML files.")
        branch = input("\nDo you need to export from a branch? This is typically required if there is a new CSC package for import. [Y/N]: ")
        if branch in questionYes:
            global branchName
            branchName = input("Please enter the branch name: ")
            logger.info("You chose to export from the branch: %s" % branchName)

        #Prompt to select generate script if not on mac or can't find in normal install location
        if sys.platform == "darwin":
            generateFile = "/Applications/MagicDraw/plugins/com.nomagic.magicdraw.reportwizard/generate.sh"
            if os.path.isfile(generateFile):
                getXML(generateFile)
            else:
                logger.info("Unable to find the generate.sh file at the standard location: %s" % generateFile)
                selectGenerateFile()
        else:
            selectGenerateFile()
    elif response in questionNo:
        logger.info("You chose to skip exporting the XML files. Comparing the model generated and GitHub XML files.")
        generateFile = 0
        xmlFileList = None
        compareXML(xmlFileList)
    else:
        print("Please enter 'yes' or 'no'...")
        needExport()

def selectGenerateFile():
    root.scriptFilename = filedialog.askopenfilename(initialdir=home, title='Please select the MagicDraw generate.bat (Windows) or generate.sh (macOS) file', filetypes=[("generate.bat", "*.bat"), ("generate.sh", "*.sh")])
    generateFile = root.scriptFilename
    path, filename = os.path.split(generateFile)
    if filename == "generate.sh" or filename == "generate.bat":
        getXML(generateFile)
    else:
        logger.warning("You must select either a generate.bat or generate.sh file. This should be located at <yourMagicDrawInstallLocation>/plugins/com.nomagic.magicdraw.reportwizard/.")
        selectGenerateFile()

def getXML(generateFile):
    origXML = currdir + "/XML/sal_interfaces"
    root.xmlFilename = filedialog.askopenfilenames(initialdir=origXML, title='Please select the GitHub XML file(s)', filetypes=[("XML Files", "*.xml")])
    selectedFiles = root.xmlFilename
    xmlFileList = list(selectedFiles)
    #If we need to export the XML Files
    if len(xmlFileList) > 0:
        exportResult = 0
        if generateFile != 0:
            exportResult = exportXML(xmlFileList, generateFile)
        if exportResult == 0 :
            logger.info("Model XML files are exported. Comparing the model generated and GitHub XML files.")
            compareXML(xmlFileList)
        else:
            messagebox.showerror("XML Export Errors", exportResult)
            #sys.exit()
    else:
        logger.warning("You must select an XML file.")
        getXML(generateFile)

def exportXML(xmlFileList, generateFile):
    propArray = []
    for xmlFileLocation in xmlFileList:
        #Temporarily set path to same directory as python file
        path, filename = os.path.split(xmlFileLocation)
        path = currdir
        fileName = filename
        fileBase = filename.split('.')[0]
        messageType = fileBase.split('_')[1]
        cscName = fileBase.split('_')[0]
        #Set the file locations
        fileLocation = path + "/XML_Model_Output/" + cscName + "/"
        xmlFileLocation = fileLocation + fileBase + ".xml"
        propFileLocation = fileLocation + fileBase + ".properties"

        #Check if file directory exists
        if os.path.isdir(fileLocation) == False:
            os.mkdir(fileLocation)
            logger.info("The folder does not exist and will be created: %s" % fileLocation)

        #Check if XML file exists and delete
        if os.path.isfile(xmlFileLocation):
            os.remove(xmlFileLocation)
            logger.info("Deleting the existing XML file: %s" % xmlFileLocation)

        #Create the property file
        project = 'project = Telescope & Site Software Components\n'
        package = 'package = ' + cscName + '\n'
        template = 'template = '+ signalType + ' XML Export\n'
        output = 'output = ' + xmlFileLocation

        propFile = open(propFileLocation, "w")
        propFile.write(project)
        if branchName != "N/A":
            branch = 'branch = '+ branchName +'\n'
            propFile.write(branch)
        propFile.write(package)
        propFile.write(template)
        propFile.write(output)
        propFile.close()
        propArray.append(propFileLocation)

    propList = "' '".join(propArray)
    #Run the command line generate script to export the files
    exportResult = subprocess.run([generateFile, "-server 'twcloud.lsst.org'", "-servertype 'twcloud'", "-ssl true", "-login 'rgenerator'", "-spassword '4e0e16bb44e1c62fe007b776d088c899c585425cbf6f424161867c13735f7f67e683f676e4eb811e4ad464e04c822cb14c1c8a385bc7461b8e94ae7292a31ee9e2758a20124d77b7c79546c6be4878db10f021f91f147833cc63fb93b8aefb6f13110dd959128c49a1da292c7ee8f98640982e1e75b40abc397514d4712d471c'", "-leaveprojectopen true", "-properties '" + propList + "'"], check=True)
    result = exportResult.returncode
    return result

def compareXML(xmlFileList):
    #Ask if they want to delete previous results
    deleteResponse = input("Do you want to delete the previous comparison results files? [Y/N]: ")
    if deleteResponse in questionYes:
        logger.info("You chose to delete the previous comparison results files.")

    compareArray = []
    if xmlFileList == None:
         #Get the GitHub XML file
         origXML = currdir + "/XML/sal_interfaces"
         root.xmlFilename = filedialog.askopenfilenames(initialdir=origXML, title='Please select the GitHub XML file(s)', filetypes=[("XML Files", "*.xml")])
         selectedFiles = root.xmlFilename
         xmlFileList = list(selectedFiles)
    if len(xmlFileList) > 0:
       for gitXMLLocation in xmlFileList:
           #Temporarily set path to same directory as python file
           path, filename = os.path.split(gitXMLLocation)
           path = currdir
           fileName = filename
           fileBase = filename.split('.')[0]
           messageType = fileBase.split('_')[1]
           cscName = fileBase.split('_')[0]
           #Set the file locations
           fileLocation = path + "/XML_Model_Output/" + cscName + "/"
           modelXMLLocation = fileLocation + fileBase + ".xml"

           #Delete previous files if selected
           if deleteResponse in questionYes:
               if os.path.isdir(fileLocation):
                   files = os.listdir(fileLocation)
                   for f in files:
                       if ".txt" in str(f) and fileBase in str(f):
                           os.remove(os.path.join(fileLocation, f))
               else:
                   logger.error("Unable to find the directory: %s" % fileLocation)

           if os.path.isfile(modelXMLLocation):
               diff = main.diff_files(modelXMLLocation, gitXMLLocation, diff_options={'F': 0.1, 'ratio_mode': 'accurate', 'fast_match': True}, formatter=None)
               if len(diff) == 0:
                   logger.info("There were no differences found!")
               else:
                   #Write output to text file
                   timestr = time.strftime("%Y%m%d-%H%M%S")
                   outputFile = os.path.join(currdir + "/XML_Model_Output/" + cscName, "xmlCompareOutput_" + fileBase + "_" + timestr + ".txt")
                   f = open(outputFile, 'w')
                   f.write(str(diff).replace('),', '),\n'))
                   f.close()
                   #Log differences
                   logger.warning("%s: There are %s actions required to make the {%s} Model exported XML the same as the GitHub XML file. Please check the compare results window.", cscName, len(diff), str(fileBase))
                   compareResult = cscName + ": There are "+ str(len(diff)) +" actions required to make the {"+ fileBase +"} Model exported XML the same as the GitHub XML file. Please check the compare results file: "+ outputFile
                   compareArray.append(compareResult)
           else:
               logger.error(cscName + ": Unable to find the model XML file {%s}" % modelXMLLocation)
               compareResult = cscName + ": Unable to find the model XML file {" + modelXMLLocation + "}"
               compareArray.append(compareResult)

       if len(compareArray) > 0:
            compareList = "\n\n".join(compareArray)
            showCompareResults(compareList)
    else:
        logger.warning("You must select an XML file.")
        compareXML(xmlFileList)

def showCompareResults(compareList):
    windowTitle = "XML Comparison Results"
    topLabel = "The XML files had the following differences.\n\n"
    bottomLabel = "\n\nPlease review the full comparison results.\n\n"
    compareWindow = tk.Toplevel()
    compareWindow.title(windowTitle)

    #Center and size the window
    myWidth = int(compareWindow.winfo_screenwidth() / 2)
    myHeight = int(compareWindow.winfo_screenheight() / 2)
    myLeftPos = (compareWindow.winfo_screenwidth() - myWidth) / 2
    myTopPos = (compareWindow.winfo_screenheight() - myHeight) / 2
    compareAreaHeight = int((myHeight / 20) - 1)
    compareWindow.geometry( "%dx%d+%d+%d" % (myWidth, myHeight, myLeftPos, myTopPos))

    #Frame for compare message
    compareFrame = ttk.Frame(compareWindow)
    compareFrame.grid(row=0, column=0, columnspan=3, ipadx=10, ipady=10, sticky="N, W, E, S")
    compareWindow.columnconfigure(0, weight=1)
    compareWindow.rowconfigure(1, weight=1)
    compareFrame.columnconfigure(0, weight=1)
    compareFrame.rowconfigure(0, weight=1)

    beforeCompare = ttk.Label(compareFrame, text=topLabel).grid(row=0, column=0, columnspan=3)
    compareArea = scrolledtext.ScrolledText(compareFrame, wrap="word", width=myWidth, height=compareAreaHeight)
    compareArea.grid(row=1, column=0, sticky="N, W, E")
    compareArea.insert(tk.INSERT, compareList)
    afterCompare = ttk.Label(compareFrame, text=bottomLabel).grid(row=2, column=0, columnspan=3)

    #Button frame
    buttonFrame = ttk.Frame(compareWindow)
    buttonFrame.grid(row=1, column=0, columnspan=2, ipadx=10, ipady=10, sticky="W, E")
    buttonFrame.columnconfigure(0, weight=1)
    buttonFrame.rowconfigure(1, weight=1)
    closeButton = ttk.Button(buttonFrame, text="Close", default="active", command=quitProcess).grid(row=1, column=1, padx=10, pady=10, sticky="SE")

    #Require a button click
    compareWindow.protocol("WM_DELETE_WINDOW", disableEvent)
    compareArea.configure(state ='disabled')
    compareWindow.resizable(False, False)

    #Open the window
    compareWindow.mainloop()

def quitProcess():
    sys.exit()

def disableEvent():
    pass
#Define global variables
root = tk.Tk()
root.withdraw() #use to hide tkinter window

currdir = os.getcwd()
home = str(Path.home())
xmlFileLocation = ""
xmlFileLocation = ""
xmlFileLocation = ""
questionYes = {"Y", "y", "Yes", "yes"}
questionNo = {"N", "n", "No", "no"}
branchName = "N/A"

#Setup logging
user = getpass.getuser()

logger = logging.getLogger('Rubin Observatory XML Message Converter')

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

#Ask if XML exports are required
#showExportQuestion(root)
needExport = needExport()
