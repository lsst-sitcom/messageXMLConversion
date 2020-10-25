#Import modules
import lxml
from lxml import etree
import xmldiff
from xmldiff import main, formatting
import tkinter
from tkinter import filedialog
from tkinter import Tk
import logging
import getpass
import os
import sys
import time

#Request XML files from user
def getXML1(root, currdir):
    origXML = currdir + "/XML/sal_interfaces"
    root.xmlFilename = filedialog.askopenfilename(initialdir=origXML, title='Please select the GitHub XML file', filetypes=[("XML Files", "*.xml")])
    xmlFileLocation1 = root.xmlFilename
    if len(xmlFileLocation1) > 0:
        logger.info("You chose %s for the GitHub XML file." % xmlFileLocation1)
        getXML2(root, currdir, xmlFileLocation1)
    else:
        logger.warning("You must select an XML file.")
        getXML1(root, currdir)

def getXML2(root, currdir, xmlFileLocation1):
    #Get the second XML file
    modelXML = currdir + "/XML_Model_Output"
    root.xmlFilename = filedialog.askopenfilename(initialdir=modelXML, title='Please select the Model generated XML file', filetypes=[("XML Files", "*.xml")])
    xmlFileLocation2 = root.xmlFilename
    if len(xmlFileLocation2) > 0:
        logger.info("You chose %s for the Model generated XML file." % xmlFileLocation2)
        #Run the xmldiff
        diff = main.diff_files(xmlFileLocation2, xmlFileLocation1, diff_options={'F': 0.1, 'ratio_mode': 'accurate', 'fast_match': True}, formatter=None)
        if len(diff) == 0:
            logger.info("There were no differences found!")
        else:
            logger.warning("There are %s actions required to make the Model exported XML the same as the GitHub XML file. Please check the output file." % len(diff))
            #Write output to text file
            path, filename = os.path.split(xmlFileLocation1)
            fileBase = filename.split('.')[0]
            timestr = time.strftime("%Y%m%d-%H%M%S")
            outputFile = os.path.join(currdir + "/XML_Compare_Output", "xmlCompareOutput_" + fileBase + "_" + timestr + ".txt")
            f = open(outputFile, 'w')
            f.write(str(diff).replace('),', '),\n'))
            f.close()
            os.system("start " + outputFile)
    else:
        logger.warning("You must select an XML file.")
        getXML2(root, currdir, xmlFileLocation1)

#Define global variables
root = Tk()
root.withdraw() #use to hide tkinter window

currdir = os.getcwd()
xmlFileLocation1 = ""
xmlFileLocation2 = ""

#Setup logging
user = getpass.getuser()

logger = logging.getLogger('LSST XML Message Differencer')

file_log_handler = logging.FileHandler('logfile_diff.log')
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
getXML1(root, currdir)