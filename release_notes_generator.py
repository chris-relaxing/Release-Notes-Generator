#-------------------------------------------------------------------------------
# Name:             Release Notes Generator
# Purpose:          Create Word.docx Release Notes on the fly, with minimal input.
#
# Author:           Chris Nielsen
# Last Updated:     June 15, 2015

# Note:             Some of the code used for editing .docx files with Python comes from:
#                   https://github.com/mikemaccana/python-docx/blob/master/docx.py

#-------------------------------------------------------------------------------

import os
import re
import time
import shutil
import zipfile
from xml.etree import ElementTree as etree

from os.path import abspath, basename, join
from exceptions import PendingDeprecationWarning
from warnings import warn

from Tkinter import *
import tkMessageBox
import tkFileDialog


regionList = ['APAC', 'AUNZ', 'NA', 'SAM', 'India', 'EEU', 'WEU', 'MEA', 'TWN', 'EU', 'KOR', 'HK']
dvnList = ['151E0','15105','15109','151F0','15118','15122','151G0','15131','15135','151H0','15144','15148', '151J0','161E0','161F0','161G0','161H0']
monthList = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']
yearList = ['2013', '2014', '2015', '2016', '2017']
productList = []
versionList = ['1.0', '2.1', '3.0', '4.0']

selected_region = ''
selected_initDVN = ''
selected_product = ''
selected_month = ''
selected_year = ''
selected_version = ''

# Inputs for testing
##region = "SAM"
##qtr = "Q2"
##year = "2014"
##month = "April"
##product = "2D Signs"

# All Word prefixes / namespace matches used in document.xml & core.xml.
# LXML doesn't actually use prefixes (just the real namespace) , but these # make it easier to copy Word output more easily.
nsprefixes = {
    'mo': 'http://schemas.microsoft.com/office/mac/office/2008/main',
    'o':  'urn:schemas-microsoft-com:office:office',
    've': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
    # Text Content
    'w':   'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'w10': 'urn:schemas-microsoft-com:office:word',
    'wne': 'http://schemas.microsoft.com/office/word/2006/wordml',
    # Drawing
    'a':   'http://schemas.openxmlformats.org/drawingml/2006/main',
    'm':   'http://schemas.openxmlformats.org/officeDocument/2006/math',
    'mv':  'urn:schemas-microsoft-com:mac:vml',
    'pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture',
    'v':   'urn:schemas-microsoft-com:vml',
    'wp':  ('http://schemas.openxmlformats.org/drawingml/2006/wordprocessing'
            'Drawing'),
    # Properties (core and extended)
    'cp':  ('http://schemas.openxmlformats.org/package/2006/metadata/core-pr'
            'operties'),
    'dc':  'http://purl.org/dc/elements/1.1/',
    'ep':  ('http://schemas.openxmlformats.org/officeDocument/2006/extended-'
            'properties'),
    'xsi': 'http://www.w3.org/2001/XMLSchema-instance',
    # Content Types
    'ct':  'http://schemas.openxmlformats.org/package/2006/content-types',
    # Package Relationships
    'r':  ('http://schemas.openxmlformats.org/officeDocument/2006/relationsh'
           'ips'),
    'pr':  'http://schemas.openxmlformats.org/package/2006/relationships',
    # Dublin Core document properties
    'dcmitype': 'http://purl.org/dc/dcmitype/',
    'dcterms':  'http://purl.org/dc/terms/'}

#------------------------------------------------------------------------
class Page(Frame):                                                               # A tk Frame widget
    def __init__(self, parent, page, *args, **kwargs):
        Frame.__init__(self, *args, borderwidth=0, **kwargs)
        self.parent = parent
        self.pack(fill=BOTH, expand=1)
        self.columnconfigure(0, weight = 1)
        self.centerWindow()

        if page == "p1":
            self.initUI(page)
        else:
            self.initPage2(page)


    def initUI(self, page):

        root.title("Release Notes Generator")

        windowBorder = LabelFrame(self, text=" Release Notes Inputs: ", padx=0, pady=0, width=740,height=260)
        windowBorder.grid(row = 0, column = 0, pady=10, padx=10, columnspan = 3, rowspan = 4, sticky='NW')

        region = StringVar()
        initDVN = StringVar()
        month = StringVar()
        year = StringVar()
        product = StringVar()
        version = StringVar()

        select_width = 48

        product.set('Select Product:') # default value
        S = OptionMenu(self,  product, *productList)
        S.config(width=select_width)
        S.pack( side = LEFT)
        S.grid(row = 1, column = 0, pady=10, padx=20, sticky='NW')

        region.set('Select Region:') # default value
        O = OptionMenu(self, region, *regionList)
        O.config(width=select_width)
        O.pack( side = LEFT)
        O.grid(row = 1, column = 1, pady=10, padx=20, columnspan = 2, sticky='NW')


        month.set('Select Month:') # default value
        Q = OptionMenu(self,  month, *monthList)
        Q.config(width=select_width)
        Q.pack( side = LEFT)
        Q.grid(row = 2, column = 0, pady=10, padx=20, sticky='NW')


        year.set('Select Year:') # default value
        R = OptionMenu(self,  year, *yearList)
        R.config(width=select_width)
        R.pack( side = LEFT)
        R.grid(row = 2, column = 1, pady=10, padx=20, columnspan = 2, sticky='NW')


        initDVN.set('Select the initial release DVN:') # default value
        P = OptionMenu(self, initDVN, *dvnList)
        P.config(width=select_width)
        P.pack( side = LEFT)
        P.grid(row = 3, column = 0, pady=10, padx=20, sticky='NW')


        DVN = StringVar()
        Label(self, text = 'DVN:').grid(row = 3, column = 1, pady=15, padx=0, sticky='NE')
        Entry(self, width=6, textvariable = DVN).grid(row = 3, column = 2, pady=15, padx=0, sticky='NW')


        submitButton = LabelFrame(self, text="", padx=0, pady=0, width=740,height=80)
        submitButton.grid(row = 4, column = 0, pady=10, padx=10, columnspan = 3, sticky='NW')

        Button(self, text = '     Generate Release Notes     ', command = lambda: multCommands(region, initDVN, product, month, year, DVN)).grid(row = 4, columnspan = 3, pady=35, padx=15, sticky='N')

        def multCommands(region, initDVN, product, month, year, DVN):

            global selected_region
            global selected_initDVN
            global selected_product
            global selected_month
            global selected_year
            global selected_DVN

            region = str(region.get())
            initDVN = str(initDVN.get())
            month = str(month.get())
            year = str(year.get())
            product = str(product.get())
            DVN = str(DVN.get())

            selected_region = region
            selected_initDVN = initDVN
            selected_product = product
            selected_month = month
            selected_year = year
            if DVN <> '':
                selected_DVN = DVN
            else:
                selected_DVN = initDVN

            printInputs(region, initDVN, product, month, year, DVN)


            # This is the logic that determines whether or not to go on to a second
            # page of inputs. A second page of inputs will appear (asking for version (placeholder))
            # if the product selected is "Hypothetical". Otherwise, the root window will close after
            # one page of inputs.
            # -------------------------------
            if selected_product == "Hypothetical":
                self.callback()
            else:
                try:
                    self.close_window()
                except:
                    pass
            # -------------------------------

        def printInputs(region, initDVN, product, month, year, DVN):

            print "The selected region is:", region
            print "The selected initial release DVN is:", initDVN
            print "The selected month is:", month
            print "The selected year is:", year
            print "The selected product is:", product
            print "The selected DVN is:", DVN

            missing_selections = ["Select Region:", "Select Quarter:", "Select Month:", "Select Year:", "Select Product:"]
            e = "Error"
            if product == missing_selections[4]:
                m = "Error. Please select product to continue."
                ThrowError(e, m, "", "")
            elif region == missing_selections[0]:
                m = "Error. Please select region to continue."
                ThrowError(e, m, "", "")
            elif initDVN == missing_selections[1]:
                m = "Error. Please select initial release DVN to continue."
                ThrowError(e, m, "", "")
            elif year == missing_selections[3]:
                m = "Error. Please select year to continue."
                ThrowError(e, m, "", "")
            elif month == missing_selections[2]:
                m = "Error. Please select month to continue."
                ThrowError(e, m, "", "")
            else:
                pass


    def initPage2(self, page):

        windowBorder = LabelFrame(self, text=" More release notes inputs needed: ", padx=0, pady=0, width=740,height=260)
        windowBorder.grid(row = 0, column = 0, pady=10, padx=10, columnspan = 2, rowspan = 4, sticky='NW')

        version = StringVar()
        select_width = 46
        version.set('Select Version:') # default value
        t = OptionMenu(self,  version, *versionList)
        t.config(width=select_width)
        t.pack( side = TOP)
        t.grid(row = 1, column = 0, pady=0, padx=20, sticky='NW')

        submitButton = LabelFrame(self, text="", padx=0, pady=0, width=600,height=80)
        submitButton.grid(row = 4, column = 0, pady=10, padx=10, columnspan = 2, sticky='NW')

        Button(self, text = '     Generate Release Notes     ', command = lambda: multCommands2(version)).grid(row = 4, columnspan = 2, pady=35, padx=15, sticky='N')

        def multCommands2(version):
            self.callback()
            printInputs2(version)

        def printInputs2(version):

            global selected_version
            version = str(version.get())
            selected_version = version
            print "The selected version is:", version

    def centerWindow(self):
        w = 760
        h = 380

        sw = root.winfo_screenwidth()
        sh = root.winfo_screenheight()

        x = (sw - w)/2
        y = (sh - h)/2
        root.geometry('%dx%d+%d+%d' % (w,h, x, y))

    def onlift(self):
        self.lift()

    def close_window(self):
        root.destroy()

#------------------------------------------------------------------------
class App(Frame):                                               # A tk Frame widget app, allowing for switching between multiple frames
    def __init__(self, *args, **kwargs):
        Frame.__init__(self, *args, **kwargs)

        root.protocol("WM_DELETE_WINDOW", self.handler)

        p1 = Page(self, 'p1')                                   # Create two Page instances
        p2 = Page(self, 'p2')                                   #

        p1.callback = p2.onlift                                 # Switch to the second window
        p2.callback = p2.close_window                           # close the second window

        p1.place(x=0, y=0, relwidth=1, relheight=1)             # both frames stacked on top of each other
        p2.place(x=0, y=0, relwidth=1, relheight=1)             # both frames stacked on top of each other

        p1.onlift()

    def handler(self):
        if tkMessageBox.askokcancel("Quit?", "Are you sure you want to quit?"):
            root.destroy()
            print "Destoy root window."
            self.master.quit()
            print "Quit main loop."
            sys.exit()

#------------------------------------------------------------------------

#taken from http://stackoverflow.com/questions/458436/adding-folders-to-a-zip-file-using-python
def zipdir(dirPath=None, zipFilePath=None, includeDirInZip=False):

    if not zipFilePath:
        zipFilePath = dirPath + ".zip"
    if not os.path.isdir(dirPath):
        raise OSError("dirPath argument must point to a directory.'%s' does not." % dirPath)
    parentDir, dirToZip = os.path.split(dirPath)

    ##---------------------------------
    def trimPath(path):
        try:
            archivePath = path.replace(parentDir, "", 1)
            if parentDir:
                archivePath = archivePath.replace(os.path.sep, "", 1)
            if not includeDirInZip:
                archivePath = archivePath.replace(dirToZip + os.path.sep, "", 1)
            return os.path.normcase(archivePath)
        except:
            print "trimPath failure, exiting.."
            sys.exit()
    ##---------------------------------

    try:
        outFile = zipfile.ZipFile(zipFilePath, "w", compression=zipfile.ZIP_DEFLATED)
    except:
        e = "Error"
        m = "Error. The Release Notes generator is looking for a \"new_rn\" folder in the same directory where the script is running. \nThis folder needs to be created and is where your generated release notes will be stored."
        ThrowError(e, m, generated_folder, "")
        sys.exit()
    for (archiveDirPath, dirNames, fileNames) in os.walk(dirPath):
        for fileName in fileNames:
            filePath = os.path.join(archiveDirPath, fileName)
            outFile.write(filePath, trimPath(filePath))
        # Make sure we get empty directories as well
        if not fileNames and not dirNames:
            zipInfo = zipfile.ZipInfo(trimPath(archiveDirPath) + "/")

    outFile.close()

#------------------------------------------------------------------------
def createSecondaries():
    global yyyy_q
    global qqyy
    global qq_yy
    global full_region

    region = selected_region
##    qtr = selected_qtr
    product = selected_product
    month = selected_month
    year = selected_year
    version = selected_version

    regionHash = {
        'TWN' : 'Taiwan',
        'APAC' : 'Asia Pacific',
        'WEU' : 'Western Europe',
        'EEU' : 'Eastern Europe',
        'NA' : 'North America',
        'RN' : 'India',
        'India' : 'India',
        'SAM' : 'South America',
        'MEA' : 'Middle East/Africa',
        'AUNZ' : 'Australia/New Zealand',
        'EU' : 'Europe',
        'KOR' : 'South Korea',
        'HK' : 'Hong Kong-China'
        }

##    q = qtr[1:]
##    yy = year[2:]
##    yyyy_q = year+'.'+q
##
##    qqyy = qtr+yy
##    qq_yy = qtr+'/'+yy

    try:
        full_region = regionHash[region]+' ('+region+')'
        if region == "AUNZ":
            full_region = regionHash[region]+' (AU)'
        if region == "India":
            full_region = region
    except:
        sys.exit()

#------------------------------------------------------------------------
def getReplacements():

    global selected_region
    if selected_region == 'AUNZ':
        adjusted_region = 'AU'
    else:
        adjusted_region = selected_region

    # Populated with globals
    replacementHash = {
        '==YEAR==' : selected_year,                         # eg. 2014
        '==INITDVN==' : selected_initDVN,                   # eg. 151F0,15135
        '==REGION==' : adjusted_region,                     # eg. TWN
        '==MONTH==' : selected_month,                       # eg. February
        '==FULL_REGION==' : full_region,                    # eg. Taiwan (TWN)
        '==DVN==' : selected_DVN                            # eg. 151F0
##        '==YYYY.Q==' : yyyy_q,                              # eg. 2014.2
##        '==QQYY==' : qqyy,                                  # eg. Q214
##        '==QQ/YY==' : qq_yy,                                # eg. Q2/14
        }

    return replacementHash
#------------------------------------------------------------------------
def readDocument(theDirectory):
    xmlDataFile = open(theDirectory)
    xmlData = file.read(xmlDataFile)
    document = etree.fromstring(xmlData)
    return document

#------------------------------------------------------------------------
# Unzip an OpenXML Document and pass the directory back
def unpackTheOpenXMLFile(theOpenXMLFile, uncompressedDirectoryName):
    theFile = zipfile.ZipFile(theOpenXMLFile)
    theFile.extractall(path=uncompressedDirectoryName)
    return uncompressedDirectoryName

#------------------------------------------------------------------------
# The AdvSearch and AdvReplace were based off of https://github.com/mikemaccana/python-docx/blob/master/docx.py
def findTypeParent(element, tag):
    """ Finds fist parent of element of the given type

    @param object element: etree element
    @param string the tag parent to search for

    @return object element: the found parent or None when not found
    """

    p = element
    while True:
        p = p.getparent()
        if p.tag == tag:
            return p

    # Not found
    return None

#------------------------------------------------------------------------
def AdvReplace(document, search, replace, bs=3):
    # Change this function so that search and replace are arrays instead of strings
    """
    Replace all occurences of string with a different string, return updated
    document

    This is a modified version of python-docx.replace() that takes into
    account blocks of <bs> elements at a time. The replace element can also
    be a string or an xml etree element.

    What it does:
    It searches the entire document body for text blocks.
    Then scan thos text blocks for replace.
    Since the text to search could be spawned across multiple text blocks,
    we need to adopt some sort of algorithm to handle this situation.
    The smaller matching group of blocks (up to bs) is then adopted.
    If the matching group has more than one block, blocks other than first
    are cleared and all the replacement text is put on first block.

    Examples:
    original text blocks : [ 'Hel', 'lo,', ' world!' ]
    search / replace: 'Hello,' / 'Hi!'
    output blocks : [ 'Hi!', '', ' world!' ]

    original text blocks : [ 'Hel', 'lo,', ' world!' ]
    search / replace: 'Hello, world' / 'Hi!'
    output blocks : [ 'Hi!!', '', '' ]

    original text blocks : [ 'Hel', 'lo,', ' world!' ]
    search / replace: 'Hel' / 'Hal'
    output blocks : [ 'Hal', 'lo,', ' world!' ]

    @param instance  document: The original document
    @param str       search: The text to search for (regexp)
    @param mixed     replace: The replacement text or lxml.etree element to
                         append, or a list of etree elements
    @param int       bs: See above

    @return instance The document with replacement applied

    """
    # Enables debug output
    DEBUG = False

    newdocument = document

    # Compile the search regexp

    for k, v in replacementHash.iteritems():
        #print k, v
        search = k
        replace = v

        searchre = re.compile(search)

        # Will match against searchels. Searchels is a list that contains last
        # n text elements found in the document. 1 < n < bs
        searchels = []

        # If using Python 2.6, use newdocument.getiterator() instead of newdocument.iter():
        for element in newdocument.iter():
            if element.tag == '{%s}t' % nsprefixes['w']:  # t (text) elements
                if element.text:
                    # Add this element to searchels
                    searchels.append(element)
                    if len(searchels) > bs:
                        # Is searchels is too long, remove first elements
                        searchels.pop(0)

                    # Search all combinations, of searchels, starting from
                    # smaller up to bigger ones
                    # l = search lenght
                    # s = search start
                    # e = element IDs to merge
                    found = False
                    for l in range(1, len(searchels)+1):
                        if found:
                            break
                        #print "slen:", l
                        for s in range(len(searchels)):
                            if found:
                                break
                            if s+l <= len(searchels):
                                e = range(s, s+l)
                                #print "elems:", e
                                txtsearch = ''
                                for k in e:
                                    txtsearch += searchels[k].text

                                # Searcs for the text in the whole txtsearch
                                match = searchre.search(txtsearch)
                                if match:
                                    found = True


                                    curlen = 0
                                    replaced = False
                                    for i in e:
                                        curlen += len(searchels[i].text)
                                        if curlen > match.start() and not replaced:
                                            # The match occurred in THIS element.
                                            # Puth in the whole replaced text
                                            if isinstance(replace, etree._Element):
                                                # Convert to a list and process
                                                # it later
                                                replace = [replace]
                                            if isinstance(replace, (list, tuple)):
                                                # I'm replacing with a list of
                                                # etree elements
                                                # clear the text in the tag and
                                                # append the element after the
                                                # parent paragraph
                                                # (because t elements cannot have
                                                # childs)
                                                p = findTypeParent(
                                                    searchels[i],
                                                    '{%s}p' % nsprefixes['w'])
                                                searchels[i].text = re.sub(
                                                    search, '', txtsearch)
                                                insindex = p.getparent().index(p)+1
                                                for r in replace:
                                                    p.getparent().insert(
                                                        insindex, r)
                                                    insindex += 1
                                            else:
                                                # Replacing with pure text
                                                searchels[i].text = re.sub(
                                                    search, replace, txtsearch)
                                            replaced = True

                                        else:
                                            # Clears the other text elements
                                            searchels[i].text = ''
    return newdocument

#------------------------------------------------------------------------
def saveElements(document, docName):

    if 'footer' in docName:
        theData = etree.tostring(document)
        outputPath = extraction_dir+'\\'+docName
        theOutputFile = open(outputPath, 'w')
        theOutputFile.write(theData)
    elif 'header' in docName:
        theData = etree.tostring(document)
        outputPath = extraction_dir+'\\'+docName
        theOutputFile = open(outputPath, 'w')
        theOutputFile.write(theData)
    elif docName == 'document.xml':
        theData = etree.tostring(document)
        outputPath = extraction_dir+'\\'+docName
        theOutputFile = open(outputPath, 'w')
        theOutputFile.write(theData)
    else:
        pass

#------------------------------------------------------------------------

def ThrowError(title, message, path, special_note):
    root = Tk()
    root.title(title)

    w = 1000
    h = 200

    sw = root.winfo_screenwidth()
    sh = root.winfo_screenheight()

    x = (sw - w)/2
    y = (sh - h)/2
    root.geometry('%dx%d+%d+%d' % (w,h, x, y))

    m = message
    m += '\n'
    m += path
    m += special_note
    w = Label(root, text=m, width=240, height=10)
    w.pack()
    b = Button(root, text="OK", command=root.destroy, width=10)
    b.pack()
    mainloop()

#------------------------------------------------------------------------
# Placeholder function to be used if config files are implemented
def readConfig():
    script_dir = os.path.dirname(os.path.realpath(__file__))
    print script_dir
#------------------------------------------------------------------------
def getScriptPath():
    return os.path.dirname(os.path.realpath(sys.argv[0]))
#------------------------------------------------------------------------
def setupEnvironment():

    global script_dir
    global scratch_folder
    global generated_folder
    global template_folder
    global new_rn
    global specific_template
    global extraction_dir

    script_dir = getScriptPath()
    print script_dir
    scratch_folder = script_dir+'\scratch'
    print scratch_folder
    template_folder = script_dir+'\\templates'
    print template_folder
    generated_folder = script_dir+'\\new_rn'
    print generated_folder


    extraction_dir = scratch_folder+'\\word'
    print extraction_dir
#------------------------------------------------------------------------
def replaceALL(theDocumentData, replacementHash):
    newXMLobject = AdvReplace(theDocumentData, replacementHash, '')
    return newXMLobject

#------------------------------------------------------------------------
def getReleaseNotesName():
    underscore = "_"
    quarterly_release = 0
    global selected_initDVN

    if selected_initDVN.isdigit():
        quarterly_release = 0
    else:
        quarterly_release = 1

    if selected_initDVN:
        selected_initDVN = " "+selected_initDVN
    else:
        selected_initDVN = ""

    if underscore in selected_product:
        u = selected_product.split(underscore)
        deduced_product = u[0]
        deduced_region = u[1]
        print "deduced product and region are:", deduced_product, deduced_region
        rn_name = deduced_product+" "+deduced_region+selected_initDVN+" Release Notes.docx"
    else:
        rn_name = selected_product+" "+selected_region+selected_initDVN+" Release Notes.docx"

    print "rn_name", rn_name
    return rn_name

#------------------------------------------------------------------------
def loadProductTemplates():
    global productList
    print template_folder
    dot = "."
    temp_file = "~"
    try:
        for file_name in os.listdir(template_folder):
            p = file_name.split(dot)
            product_name = p[0]
            print product_name
            if temp_file not in product_name:
                productList.append(product_name)
        print productList
    except:
        e = "Error"
        m = "Error. The Release Notes generator is looking for a \"templates\" folder in the same directory where the script is running. \nThis folder needs to be created and is where your release notes templates will be stored."
        ThrowError(e, m, template_folder, "")
        sys.exit()


#------------------------------------------------------------------------
if __name__ == '__main__':


    setupEnvironment()
    loadProductTemplates()

    root = Tk()
    root.resizable(0, 0)
    app = App(root)
    root.mainloop()

    createSecondaries()
    replacementHash = getReplacements()

    specific_template = template_folder+'\\'+selected_product+'.docx'
    print "specific_template", specific_template

    theDirectory = unpackTheOpenXMLFile(specific_template, scratch_folder)

    filePath = extraction_dir+'\\'+'document.xml'
    headerPath = extraction_dir+'\\'+'header1.xml'
    headerPath2 = extraction_dir+'\\'+'header2.xml'
    headerPath3 = extraction_dir+'\\'+'header3.xml'
    footerPath = extraction_dir+'\\'+'footer1.xml'
    footerPath2 = extraction_dir+'\\'+'footer2.xml'

    theDocumentData = readDocument(filePath)
    theHeaderData = readDocument(headerPath)
    theHeaderData2 = readDocument(headerPath2)
    theHeaderData3 = readDocument(headerPath3)
    theFooterData = readDocument(footerPath)
    theFooterData2 = readDocument(footerPath2)

    documentBody = replaceALL(theDocumentData, replacementHash)
    documentHeader = replaceALL(theHeaderData, replacementHash)
    documentHeader2 = replaceALL(theHeaderData2, replacementHash)
    documentHeader3 = replaceALL(theHeaderData3, replacementHash)
    documentFooter = replaceALL(theFooterData, replacementHash)
    documentFooter2 = replaceALL(theFooterData2, replacementHash)

    saveElements(documentBody, 'document.xml')
    saveElements(documentHeader, 'header1.xml')
    saveElements(documentHeader2, 'header2.xml')
    saveElements(documentHeader3, 'header3.xml')
    saveElements(documentFooter, 'footer1.xml')
    saveElements(documentFooter2, 'footer2.xml')

    rn_name = getReleaseNotesName()
    new_rn = generated_folder+'\\'+rn_name
    print new_rn

    zipdir(scratch_folder, new_rn)

    ThrowError("Process Complete", "Process complete. New Release Notes were generated and can be found here:", new_rn, "\n\n Note: The new Release Notes must be opened and saved before they will be usable.\n\n")
