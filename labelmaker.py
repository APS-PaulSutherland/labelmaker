'''
a script for generating labels from a spreadsheet
'''

from configparser import ConfigParser
from os.path import realpath, dirname, splitext
from pandas import read_excel
from re import match, search
from math import ceil
from reportlab.lib.units import inch
from reportlab.lib.pagesizes import LETTER
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase.pdfmetrics import registerFontFamily
from reportlab.pdfgen.canvas import Canvas
from reportlab.rl_config import TTFSearchPath
from reportlab.lib.utils import simpleSplit

CONTENT_INI = 'config-content.ini'
LAYOUT_INI = 'config-layout.ini'
CURRENTFOLDER = dirname(realpath(__file__))
LOGOFILE = 'logo.png'
ELLIPSIS = "…"
QUOTE_VERTICAL = "\""
QUOTE_DOUBLECOMMA_OPEN = "“"
QUOTE_DOUBLECOMMA_CLOSE = "”"

'''
parse .ini files
'''

# load content
contentIniParser = ConfigParser()
contentIniParser.read(CONTENT_INI)
# load layout
layoutIniParser = ConfigParser()
layoutIniParser.read(LAYOUT_INI)
# make specs dictionary
specs = dict()
# load main settings
specs.update(contentIniParser['main settings'])
labelType = specs['label-type-to-use']
dataFilename = specs['spreadsheet']
outputFilename = splitext(dataFilename)[0] + '_OUTPUT_' + labelType + '.pdf'
# load remainder based on labelType
if not labelType in contentIniParser.sections():
    raise ValueError('Label type ' + labelType + ' not found in ' + CONTENT_INI)
specs.update(contentIniParser[labelType])
specs.update(layoutIniParser[contentIniParser[labelType]['layout']])
print('Creating label type: ' + labelType)
print('Spreadsheet to use: ' + dataFilename)

# load spreadsheet to dictionary, each column is a list, column headers are 0 onwards
data = read_excel(dataFilename, header=None, na_filter=False)
totalLabels = data.shape[0]
dataColumns = data.shape[1]
data = data.to_dict(orient="list")
print('Number of labels to make: ' + str(totalLabels))
print('Columns in spreadsheet: ' + str(dataColumns))

# we know what our variables are, but they are in the wrong type (if not string). correct the types.
# this will also tell us if we're missing something.
FONTFACE = specs['font-face']
FONTPADDING = float(specs['font-padding'])
FONTSIZENORMAL = int(specs['font-normal'])
FONTSIZESMALL = int(specs['font-small'])
if specs['trailing-ellipsis'] == 'no':
    TRAILINGELLIPSIS = False
elif specs['trailing-ellipsis'] == 'yes':
    TRAILINGELLIPSIS = True
if specs['overflow-into-spaces'] == 'no':
    OVERFLOWALLOWED = False
elif specs['overflow-into-spaces'] == 'yes':
    OVERFLOWALLOWED = True
ROWS = int(specs['rows'])
COLUMNS = int(specs['columns'])
LABELHEIGHT = float(specs['label-height']) * inch
LABELWIDTH = float(specs['label-width']) * inch
LABELINTERNALPADDING = float(specs['label-internal-padding']) * inch
GAPPRINTERRORADJUST = float(specs['gap-print-error-adjust']) * inch
GAPX = float(specs['gap-x']) * inch
GAPXMID = float(specs['gap-x-mid']) * inch
GAPY = float(specs['gap-y']) * inch
# add error adjustment to GAPX, GAPXMID, GAPY
GAPX += GAPPRINTERRORADJUST
GAPXMID += GAPPRINTERRORADJUST
GAPY += GAPPRINTERRORADJUST
if specs['use-logo'] == 'yes':
    USELOGO = True
    LOGOWIDTH = float(specs['logo-width']) * inch
    LOGOHEIGHT = float(specs['logo-height']) * inch
elif specs['use-logo'] == 'no':
    USELOGO = False
# lines - a little more calculation
# find the lines programmatically by searching for the string ^line.*
linesInSpecs = [k for k, v in specs.items() if search(r'^line.*', k)]
LINES = [specs[x] for x in linesInSpecs]
NUMBEROFLINES = len(LINES)
# find out which are blank, and not, and some useful numbers
LINESUSEDINDICES = [index for index, item in enumerate(LINES) if item] # indices of the used lines. thanks https://stackoverflow.com/questions/41346403/return-index-of-last-non-zero-element-in-list
LINESUNUSEDINDICES = [index for index, item in enumerate(LINES) if not item]
LASTUSEDLINEINDEX = LINESUSEDINDICES[-1]
LINESUNUSEDCOUNTFINAL = NUMBEROFLINES - 1 - LASTUSEDLINEINDEX
LINESUNUSEDCOUNTMIDDLE = len([x for x in LINES[:LASTUSEDLINEINDEX] if not x])

# check and register fonts
ALLOWEDFONTS = ('Garamond', 'Arial')
if FONTFACE not in ALLOWEDFONTS:
    raise ValueError('only these fonts are allowed: ' + str(ALLOWEDFONTS))
TTFSearchPath.append(CURRENTFOLDER)
pdfmetrics.registerFont(TTFont('Garamond', 'font_garamond.ttf'))
pdfmetrics.registerFont(TTFont('Garamond_b', 'font_garamond_b.ttf'))
pdfmetrics.registerFont(TTFont('Garamond_i', 'font_garamond_i.ttf'))
pdfmetrics.registerFont(TTFont('Garamond_bi', 'font_garamond_bi.ttf'))
pdfmetrics.registerFont(TTFont('Arial', 'font_arial.ttf'))
pdfmetrics.registerFont(TTFont('Arial_b', 'font_arial_b.ttf'))
pdfmetrics.registerFont(TTFont('Arial_i', 'font_arial_i.ttf'))
pdfmetrics.registerFont(TTFont('Arial_bi', 'font_arial_bi.ttf'))
registerFontFamily('Garamond', normal='Garamond', bold='Garamond_b', italic='Garamond_i', boldItalic='Garamond_bi')
registerFontFamily('Arial', normal='Arial', bold='Arial_b', italic='Arial_i', boldItalic='Arial_bi')
             
# check the contents of LINES to make sure they only contain valid characters
for line in LINES:
    # simple validation - it's an int or a valid character
    for char in line:
        if char not in 'bis':
            try:
                char = int(char)
            except:
                raise ValueError('line ' + line + ' contains an invalid character: ' + str(char))
    # check we only have one pipe
    if line.count('|') > 1:
        raise ValueError('line ' + line + ' contains too many pipes (INCLUDE NONE, THESE ARE NOT SUPPORTED)')
    # check the data columns being referenced are within range
    line = line.split('|')
    for li in line:
        if li:
            dig = int(match('^\d+', li).group()) # first 1 or more digits
            if dig > dataColumns:
                raise ValueError('column ' + str(dig) + ' in line ' + li + ' is beyond the columns in your xlsx sheet')


print("Parsed .ini files ...")

'''
    SIZE CALCULATIONS WITHIN A LABEL
    these will give us some warnings if, say, our font size is too big to fit the number of lines we want
'''

totalPages = ceil(totalLabels/(ROWS*COLUMNS))
print('PDF pages to make: ' + str(totalPages))
pageWidth, pageHeight = LETTER

# some variables that exist independent of specific label:

# WRITEABLE AREA
textMaxHeight = LABELHEIGHT - LABELINTERNALPADDING - LABELINTERNALPADDING
textMaxWidth = LABELWIDTH - LABELINTERNALPADDING - LABELINTERNALPADDING
if USELOGO:
    textMaxWidth = textMaxWidth - LOGOWIDTH - LABELINTERNALPADDING

# LINE HEIGHT
fontFaceNormal = pdfmetrics.getFont(FONTFACE + "_b").face # use pdfmetrics to get info about the font we chose, using bold because bold big
lineHeightPx = (fontFaceNormal.ascent - fontFaceNormal.descent) / 1000 * FONTSIZENORMAL # line height is based on the biggest line - normal size - to avoid complexity
# calc max lines assuming no padding at bottom
textMaxLines = int(textMaxHeight/(lineHeightPx * FONTPADDING)) # get the minimum number, assuming all padded lines
r = textMaxHeight % (textMaxLines * lineHeightPx * FONTPADDING) # get the remainder
if lineHeightPx <= r:
    textMaxLines += 1 # add 1 if the remainder could fit a line
print('Maximum lines per label: ' + str(textMaxLines) + ' and lines provided in .ini file: ' + str(NUMBEROFLINES))
if NUMBEROFLINES > textMaxLines:
    raise ValueError('The number of lines in the .ini file is more than can be represented on the page. Label creation canceled.')

'''
    LABEL POSITIONING
'''
class Label:
    '''
    class to define the position of each label
    '''

    def __init__(self, page, row, column):
        self.page = page
        self.row = row
        self.column = column
        if USELOGO:
            self.logoTopLeft()
        self.textTopLeft()

    def logoTopLeft(self):
        self.logoX = GAPX + max(LABELWIDTH * self.column, 0) + (self.column * GAPXMID) + LABELINTERNALPADDING
        self.logoY = GAPY + max(LABELHEIGHT * self.row, 0) + LABELINTERNALPADDING
        # flip Y, because canvas writes from bottom-left for some absurd reason
        self.logoY += LOGOHEIGHT
        self.logoY = pageHeight - self.logoY

    def textTopLeft(self):
        self.textX = GAPX + max(LABELWIDTH * self.column, 0) + (self.column * GAPXMID) + LABELINTERNALPADDING
        if USELOGO:
            self.textX += LOGOWIDTH + LABELINTERNALPADDING
        self.textY = GAPY + max(LABELHEIGHT * self.row, 0) + LABELINTERNALPADDING
        # flip Y
        self.textY = pageHeight - self.textY
 
# make my labels
labels = []
c = 0
for page in range(totalPages):
    for row in range(ROWS):
        for column in range(COLUMNS):
            c += 1
            if c <= totalLabels:
                labels.append(Label(page, row, column))
            else:
                break

print('... label positions calculated ...')

'''
    LABEL TEXT
'''

class Line():
    '''
    class to define a list of the maximum contents of a specific line
    receives the line contents (e.g. 1bu), creates list of split text according to the data source and formatting instructions in that line
    '''

    def __init__(self, lineRaw):
        # put variables into a nicer place
        # text gets extracted here. column-=1 so that people aren't confused with lists that start at 0
        self.lineRaw = lineRaw
        self.parseVariables()
        self.pickFonts()
        self.splitText()
        self.linesToAttempt = 1
        
    def parseVariables(self):
        self.column = int(match('^\d+', self.lineRaw).group())
        self.text = data[self.column-1]
        self.bold = True if 'b' in self.lineRaw else False
        self.italic = True if 'i' in self.lineRaw else False
        self.small = True if 's' in self.lineRaw else False
        # unused so far
        self.underline = True if 'u' in self.lineRaw else False
        self.right = True if 'r' in self.lineRaw else False
    
    def pickFonts(self):
        # bold, italic
        if self.bold and self.italic:
            self.fontface = FONTFACE + '_bi'
        elif self.bold:
            self.fontface = FONTFACE + '_b'
        elif self.italic:
            self.fontface = FONTFACE + '_i'
        else:
            self.fontface = FONTFACE
        # small
        self.fontsize = FONTSIZESMALL if self.small else FONTSIZENORMAL
        # right, underline are irrelevant here

    def splitText(self):
        # split the text
        self.textSplit = [simpleSplit(str(text), self.fontface, self.fontsize, textMaxWidth) for text in self.text]
        # find out how long text spilt over
        self.overflowLines = len(max(self.textSplit, key=len))
        
# make my lines
lines = []
for x in LINES:
    if x:
        lines.append(Line(x))
    else:
        lines.append(None)

'''
    DETERMINE CONTENT OF EACH LINE
    nightmare for loop
'''
                    
# calculate the list of actual text to use in each line, passed to the label
for labelIndex, label in enumerate(labels):
    label.textToWrite = [None for i in range(NUMBEROFLINES)]
    label.fontFaceToUse = [None for i in range(NUMBEROFLINES)]
    label.fontSizeToUse = [None for i in range(NUMBEROFLINES)]
    for lineIndex, line in enumerate(lines):
        if line:
            for textIndex, text in enumerate(line.textSplit[labelIndex]):
                # fill the first one, in place, no questions
                if textIndex == 0:
                    label.textToWrite[lineIndex] = text
                    label.fontFaceToUse[lineIndex] = line.fontface
                    label.fontSizeToUse[lineIndex] = line.fontsize
                else:
                    # if we've reached the end of our lines, check about applying ellipsis (& quote), then break out
                    if lineIndex + textIndex >= NUMBEROFLINES:
                        if TRAILINGELLIPSIS and len(line.textSplit[labelIndex]) > textIndex:
                            label.textToWrite[lineIndex + textIndex - 1] += ELLIPSIS
                            # add the quote if needed, by checking if the very last character was a quote mark
                            if line.textSplit[labelIndex][-1][-1] == QUOTE_VERTICAL:
                                label.textToWrite[lineIndex + textIndex - 1] += QUOTE_VERTICAL
                            elif line.textSplit[labelIndex][-1][-1] == QUOTE_DOUBLECOMMA_CLOSE:
                                label.textToWrite[lineIndex + textIndex - 1] += QUOTE_DOUBLECOMMA_CLOSE
                        break
                    # if the next one is a line, AND that line HAS CONTENT, check about applying ellipsis (& quote), then break out
                    if lines[lineIndex + textIndex] and lines[lineIndex + textIndex].text[labelIndex]:
                        if TRAILINGELLIPSIS and len(line.textSplit[labelIndex]) > textIndex:
                            label.textToWrite[lineIndex + textIndex - 1] += ELLIPSIS
                            # add the quote if needed, by checking if the very last character was a quote mark
                            if line.textSplit[labelIndex][-1][-1] == QUOTE_VERTICAL:
                                label.textToWrite[lineIndex + textIndex - 1] += QUOTE_VERTICAL
                            elif line.textSplit[labelIndex][-1][-1] == QUOTE_DOUBLECOMMA_CLOSE:
                                label.textToWrite[lineIndex + textIndex - 1] += QUOTE_DOUBLECOMMA_CLOSE
                        break
                    # if the next one is not a line, check if overflow is allowed. then write the next line and fill down to end.
                    else:
                        if OVERFLOWALLOWED:
                            label.textToWrite[lineIndex + textIndex] = text
                            label.fontFaceToUse[lineIndex + textIndex] = line.fontface
                            label.fontSizeToUse[lineIndex + textIndex] = line.fontsize
                            # fill rest down to end, to avoid a situation where the very very long text of one cell continues after the shorter text of another
                            if textIndex - 1 == len(line.textSplit[labelIndex]):
                                try:
                                    for i in range(lineIndex + textIndex + 1, len(label.textToWrite) - 1):
                                        label.textToWrite[i] = None
                                        label.fontFaceToUse[i] = None
                                        label.fontSizeToUse[i] = None
                                except:
                                    pass

print("... label text calculated ...")
'''
    PUT LABELS ON THE PAGE
'''
# create the blank canvas
canvas = Canvas(outputFilename, pagesize=LETTER)

# add each label's text
pc = 0 # page count
for label in labels:
    if label.page > pc:
        pc += 1
        canvas.showPage() # showPage = move to the next page
    # if logo, draw logo
    if USELOGO:
        canvas.drawImage(LOGOFILE, label.logoX, label.logoY, LOGOWIDTH, LOGOHEIGHT)
    # write the text
    for index, text in enumerate(label.textToWrite):
        if text:
            canvas.setFont(label.fontFaceToUse[index], size=label.fontSizeToUse[index])
            canvas.drawString(label.textX, label.textY - (index * lineHeightPx * FONTPADDING), text=text)
    
# save
canvas.save()
print("Labels saved. Your output file is in this directory and saved as: " + outputFilename)
# keep window open so that .exe terminal will not disappear shortly after creating file
input("Press enter to exit.")