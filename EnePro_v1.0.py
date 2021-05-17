# EnePro v1.0
# Energy profile diagram generator, powered by Python 3.9
# Last Update: 2021-05-15
# Author: Zhe Wang
# Homepage: https://www.wangzhe95.net/program-enepro

print("*******************************************************************************")
print("*                                                                             *")
print("*                                                                             *")
print("*                                 E n e P r o                                 *")
print("*                                                                             *")
print("*                                                                             *")
print("*     ====================== Version 1.0 for macOS ======================     *")
#print("*     ====================== Version 1.0 for Linux ======================     *")
#print("*     ================ Version 1.0 for Microsoft Windows ================     *")
#print("*     =================== Version 1.0 for Source Code ===================     *")
print("*                           Last update: 2021-05-15                           *")
print("*                                                                             *")
print("*       An energy profile generator, developed by Zhe Wang. Online document   *")
print("*     is available from https://www.wangzhe95.net/program-enepro .            *")
print("*                                                                             *")
print("*                             -- Catch me with --                             *")
print("*                         E-mail  wongzit@yahoo.co.jp                         *")
print("*                       Homepage  https://www.wangzhe95.net                   *")
print("*                         GitHub  https://github.com/wongzit                  *")
print("*                                                                             *")
print("*******************************************************************************")
print("\nPRESS Ctrl+c to exit the program.\n")

import openpyxl
import time

# ============================ Get time stamp ============================
currentTime = int(time.time())

# ============================ Create ChemDraw .cdxml file ============================
eneproFile = open(f"EnePro_{currentTime}.cdxml", "w")

# ============================ Write general section to .cdxml file ============================
eneproFile.write("<?xml version=\"1.0\" encoding=\"UTF-8\" ?>\n")
eneproFile.write("<!DOCTYPE CDXML SYSTEM \"http://www.cambridgesoft.com/xml/cdxml.dtd\" >\n")
eneproFile.write("<CDXML\n")
eneproFile.write(" CreationProgram=\"ChemDraw 20.0.0.38\"\n")
eneproFile.write(" Name=\"EnePro.cdxml\"\n")
eneproFile.write(" BoundingBox=\"0 0 0 0\"\n")
eneproFile.write(" WindowPosition=\"0 0\"\n")
eneproFile.write(" WindowSize=\"0 0\"\n")
eneproFile.write(" FractionalWidths=\"yes\"\n")
eneproFile.write(" InterpretChemically=\"yes\"\n")
eneproFile.write(" ShowAtomQuery=\"yes\"\n")
eneproFile.write(" ShowAtomStereo=\"no\"\n")
eneproFile.write(" ShowAtomEnhancedStereo=\"yes\"\n")
eneproFile.write(" ShowAtomNumber=\"no\"\n")
eneproFile.write(" ShowResidueID=\"no\"\n")
eneproFile.write(" ShowBondQuery=\"yes\"\n")
eneproFile.write(" ShowBondRxn=\"yes\"\n")
eneproFile.write(" ShowBondStereo=\"no\"\n")
eneproFile.write(" ShowTerminalCarbonLabels=\"no\"\n")
eneproFile.write(" ShowNonTerminalCarbonLabels=\"no\"\n")
eneproFile.write(" HideImplicitHydrogens=\"no\"\n")
eneproFile.write(" LabelFont=\"21\"\n")
eneproFile.write(" LabelSize=\"10\"\n")
eneproFile.write(" LabelFace=\"96\"\n")
eneproFile.write(" CaptionFont=\"21\"\n")
eneproFile.write(" CaptionSize=\"10\"\n")
eneproFile.write(" HashSpacing=\"2.50\"\n")
eneproFile.write(" MarginWidth=\"1.60\"\n")
eneproFile.write(" LineWidth=\"0.60\"\n")
eneproFile.write(" BoldWidth=\"2\"\n")
eneproFile.write(" BondLength=\"14.40\"\n")
eneproFile.write(" BondSpacing=\"18\"\n")
eneproFile.write(" ChainAngle=\"120\"\n")
eneproFile.write(" LabelJustification=\"Auto\"\n")
eneproFile.write(" CaptionJustification=\"Left\"\n")
eneproFile.write(" AminoAcidTermini=\"HOH\"\n")
eneproFile.write(" ShowSequenceTermini=\"yes\"\n")
eneproFile.write(" ShowSequenceBonds=\"yes\"\n")
eneproFile.write(" ShowSequenceUnlinkedBranches=\"no\"\n")
eneproFile.write(" ResidueWrapCount=\"40\"\n")
eneproFile.write(" ResidueBlockCount=\"10\"\n")
eneproFile.write(" ResidueZigZag=\"yes\"\n")
eneproFile.write(" NumberResidueBlocks=\"no\"\n")
eneproFile.write(" PrintMargins=\"36 36 36 36\"\n")
eneproFile.write(" MacPrintInfo=\"000300000048004800000000023B0331FFF4FFF40247033E0365057B03E0000200000048004800000000023B0331000100000064000000010001010100000001270F000100010000000000000000000000000002001901900000000000400000000000000000000100000000000000000000000000000000\"\n")
eneproFile.write(" ChemPropName=\"\"\n")
eneproFile.write(" ChemPropFormula=\"Chemical Formula: \"\n")
eneproFile.write(" ChemPropExactMass=\"Exact Mass: \"\n")
eneproFile.write(" ChemPropMolWt=\"Molecular Weight: \"\n")
eneproFile.write(" ChemPropMOverZ=\"m/z: \"\n")
eneproFile.write(" ChemPropAnalysis=\"Elemental Analysis: \"\n")
eneproFile.write(" ChemPropBoilingPt=\"Boiling Point: \"\n")
eneproFile.write(" ChemPropMeltingPt=\"Melting Point: \"\n")
eneproFile.write(" ChemPropCritTemp=\"Critical Temp: \"\n")
eneproFile.write(" ChemPropCritPres=\"Critical Pres: \"\n")
eneproFile.write(" ChemPropCritVol=\"Critical Vol: \"\n")
eneproFile.write(" ChemPropGibbs=\"Gibbs Energy: \"\n")
eneproFile.write(" ChemPropLogP=\"Log P: \"\n")
eneproFile.write(" ChemPropMR=\"MR: \"\n")
eneproFile.write(" ChemPropHenry=\"Henry&apos;s Law: \"\n")
eneproFile.write(" ChemPropEForm=\"Heat of Form: \"\n")
eneproFile.write(" ChemProptPSA=\"tPSA: \"\n")
eneproFile.write(" ChemPropID=\"\"\n")
eneproFile.write(" ChemPropFragmentLabel=\"\"\n")
eneproFile.write(" color=\"0\"\n")
eneproFile.write(" bgcolor=\"1\"\n")
eneproFile.write(" RxnAutonumberStart=\"1\"\n")
eneproFile.write(" RxnAutonumberConditions=\"no\"\n")
eneproFile.write(" RxnAutonumberStyle=\"Roman\"\n")
eneproFile.write(" RxnAutonumberFormat=\"(#)\"\n")

# Color table
eneproFile.write("><colortable>\n")
eneproFile.write("<color r=\"1\" g=\"1\" b=\"1\"/>\n")   # id = 2, white, w
eneproFile.write("<color r=\"0\" g=\"0\" b=\"0\"/>\n")   # id = 3, black, b
eneproFile.write("<color r=\"1\" g=\"0\" b=\"0\"/>\n")   # id = 4, red, r
eneproFile.write("<color r=\"1\" g=\"1\" b=\"0\"/>\n")   # id = 5, yellow, y
eneproFile.write("<color r=\"0\" g=\"1\" b=\"0\"/>\n")   # id = 6, green, g
eneproFile.write("<color r=\"0\" g=\"1\" b=\"1\"/>\n")   # id = 7, light blue, lb
eneproFile.write("<color r=\"0\" g=\"0\" b=\"1\"/>\n")   # id = 8, blue, bl
eneproFile.write("<color r=\"1\" g=\"0\" b=\"1\"/>\n")   # id = 9, purple, p
eneproFile.write("<color r=\"0.0392\" g=\"0.2196\" b=\"0.3882\"/>\n")   # Dark blue, id = 10, db
eneproFile.write("<color r=\"0.2314\" g=\"0.5961\" b=\"0.7294\"/>\n")   # Sky blue, id = 11, sb
eneproFile.write("<color r=\"0.9098\" g=\"0.8510\" b=\"0.0471\"/>\n")   # Lemon, id = 12, le
eneproFile.write("<color r=\"0.5274\" g=\"0.8307\" b=\"0.0253\"/>\n")   # Lime, id = 13, l
eneproFile.write("<color r=\"0.8078\" g=\"0\" b=\"0.4431\"/>\n")   # Fuchsia, id = 14, f
eneproFile.write("<color r=\"0.5294\" g=\"0.0353\" b=\"0.3216\"/>\n")   # Plum, id = 15, pl
eneproFile.write("<color r=\"0.9922\" g=\"0.5490\" b=\"0.0902\"/>\n")   # Tangerine, id = 16, t
eneproFile.write("<color r=\"0.9804\" g=\"0.2078\" b=\"0.0275\"/>\n")   # Burnt orange, id = 17, bo
eneproFile.write("<color r=\"0.7255\" g=\"0\" b=\"0.1333\"/>\n")   # Scarlet, id = 18, s
eneproFile.write("<color r=\"0.0471\" g=\"0.3686\" b=\"0.3020\"/>\n")   # Jade, id = 19, j
eneproFile.write("<color r=\"0.0667\" g=\"0.2980\" b=\"0.3255\"/>\n")   # Forest, id = 20, fo
eneproFile.write("<color r=\"0.3804\" g=\"0.8392\" b=\"0.7922\"/>\n")   # Turquoise, id = 21, tu
eneproFile.write("<color r=\"0.2588\" g=\"0.2745\" b=\"0.2863\"/>\n")   # Graphite, id = 22, gra
eneproFile.write("</colortable><fonttable>\n")
eneproFile.write("<font id=\"21\" charset=\"x-mac-roman\" name=\"Helvetica\"/>\n")
eneproFile.write("</fonttable><page\n")
eneproFile.write(" id=\"999999999\"\n")
eneproFile.write(" BoundingBox=\"0 0 769.89 523.28\"\n")
eneproFile.write(" HeaderPosition=\"36\"\n")
eneproFile.write(" FooterPosition=\"36\"\n")
eneproFile.write(" PrintTrimMarks=\"yes\"\n")
eneproFile.write(" HeightPages=\"1\"\n")
eneproFile.write(" WidthPages=\"1\"\n")
eneproFile.write(">\n")

# ============================ Function Section ============================
# Function for printing bold lines of energy state
def boldLinePrint(id, ZOrder, orderOfSpecie, height, boldLineColor, boldLen, dashLen):
	eneproFile.write("<graphic\n")
	eneproFile.write(f" id=\"100000{id}\"\n")
	boldLineStartX = 80 + (orderOfSpecie - 1) * boldLen + (orderOfSpecie - 1) * dashLen
	boldLineEndX = 80 + orderOfSpecie * boldLen + (orderOfSpecie - 1) * dashLen
	eneproFile.write(f" BoundingBox=\"{boldLineEndX} {height} {boldLineStartX} {height}\"\n")
	eneproFile.write(f" Z=\"100000{ZOrder}\"\n")
	eneproFile.write(f" color=\"{boldLineColor}\"\n")
	eneproFile.write(" LineType=\"Bold\"\n")
	eneproFile.write(" GraphicType=\"Line\"\n")
	eneproFile.write("/>\n")

# Function for printing dash lines connecting two bold lines
def dashLinePrint(id, ZOrder, startX, startY, endX, endY, boldLineColor = 0, boldLen = 30, dashLen = 20):
	eneproFile.write("<graphic\n")
	eneproFile.write(f" id=\"{id}\"\n")
	eneproFile.write(f" BoundingBox=\"{endX} {endY} {startX} {startY}\"\n")
	eneproFile.write(f" Z=\"{ZOrder}\"\n")
	eneproFile.write(f" color=\"{boldLineColor}\"\n")
	eneproFile.write(" LineType=\"Dashed\"\n")
	eneproFile.write(" GraphicType=\"Line\"\n")
	eneproFile.write("/>\n")

# Function for printing energy values
def energyEditor(id, ZOrder, positionX1, positionX2, positionY, boldLineColor, eneData):
	eneproFile.write("<t\n")
	eneproFile.write(f" id=\"200000{id}\"\n")
	positionX = positionX1 / 2.0 + positionX2 / 2.0
	eneproFile.write(f" p=\"{positionX} {positionY - 5}\"\n")
	eneproFile.write(f" BoundingBox=\"{positionX} {positionY - 20} {positionX + 20} {positionY - 10}\"\n")
	eneproFile.write(f" Z=\"200000{ZOrder}\"\n")
	eneproFile.write(f" color=\"{boldLineColor}\"\n")
	eneproFile.write(" CaptionJustification=\"Center\"\n")
	eneproFile.write(" Justification=\"Center\"\n")
	eneproFile.write(f"><s font=\"21\" size=\"10\" color=\"{boldLineColor}\">{eneData}</s></t>\n")

# Function for printing state labels
def nameEditor(id, ZOrder, positionX1, positionX2, positionY, boldLineColor, eneName, boldFont):
	eneproFile.write("<t\n")
	eneproFile.write(f" id=\"300000{id}\"\n")
	positionX = positionX1 / 2.0 + positionX2 / 2.0
	eneproFile.write(f" p=\"{positionX} {positionY + 12}\"\n")
	eneproFile.write(f" BoundingBox=\"{positionX} {positionY - 20} {positionX + 20} {positionY - 10}\"\n")
	eneproFile.write(f" Z=\"300000{ZOrder}\"\n")
	eneproFile.write(f" color=\"{boldLineColor}\"\n")
	eneproFile.write(" CaptionJustification=\"Center\"\n")
	eneproFile.write(" Justification=\"Center\"\n")
	eneproFile.write(f"><s font=\"21\" size=\"10\" color=\"{boldLineColor}\" face=\"{boldFont}\">{eneName}</s></t>\n")

# Function for reading user-determined color
def colorMeter(colorAlpha):
	colorNumber = 0
	if colorAlpha == 'w':
		colorNumber = 2
	elif colorAlpha == 'b':
		colorNumber = 3
	elif colorAlpha == 'r':
		colorNumber = 4
	elif colorAlpha == 'y':
		colorNumber = 5
	elif colorAlpha == 'g':
		colorNumber = 6
	elif colorAlpha == 'lb':
		colorNumber = 7
	elif colorAlpha == 'bl':
		colorNumber = 8
	elif colorAlpha == 'p':
		colorNumber = 9
	elif colorAlpha == 'db':
		colorNumber = 10
	elif colorAlpha == 'sb':
		colorNumber = 11
	elif colorAlpha == 'le':
		colorNumber = 12
	elif colorAlpha == 'l':
		colorNumber = 13
	elif colorAlpha == 'f':
		colorNumber = 14
	elif colorAlpha == 'pl':
		colorNumber = 15
	elif colorAlpha == 't':
		colorNumber = 16
	elif colorAlpha == 'bo':
		colorNumber = 17
	elif colorAlpha == 's':
		colorNumber = 18
	elif colorAlpha == 'j':
		colorNumber = 19
	elif colorAlpha == 'fo':
		colorNumber = 20
	elif colorAlpha == 'tu':
		colorNumber = 21
	elif colorAlpha == 'gra':
		colorNumber = 22
	return colorNumber

# Function for calculating axis scale values
def axisDevideValue(maxEne, minEne):
	diff = maxEne - minEne
	value1 = format(minEne, '.1f')
	value2 = format(minEne + diff / 4.0 * 1, '.1f')
	value3 = format(minEne + diff / 4.0 * 2, '.1f')
	value4 = format(minEne + diff / 4.0 * 3, '.1f')
	value5 = format(maxEne, '.1f')
	return value1, value2, value3, value4, value5

# Function for priting axis scale
def axisDevideLine(id, ZOrder, height):
	eneproFile.write("<graphic\n")
	eneproFile.write(f" id=\"500000{id}\"\n")
	eneproFile.write(f" BoundingBox=\"55 {height} 50 {height}\"\n")
	eneproFile.write(f" Z=\"500000{ZOrder}\"\n")
	eneproFile.write(" color=\"0\"\n")
	eneproFile.write(" GraphicType=\"Line\"\n")
	eneproFile.write("/>\n")

# Function for printing axis scale values
def axisDevideName(id, ZOrder, height, axisEne):
	eneproFile.write("<t\n")
	eneproFile.write(f" id=\"600000{id}\"\n")
	eneproFile.write(f" p=\"45 {height + 2}\"\n")
	eneproFile.write(f" BoundingBox=\"45 {height - 20} 25 {height - 10}\"\n")
	eneproFile.write(f" Z=\"600000{ZOrder}\"\n")
	eneproFile.write(" color=\"0\"\n")
	eneproFile.write(" CaptionJustification=\"Right\"\n")
	eneproFile.write(" Justification=\"Right\"\n")
	eneproFile.write(f"><s font=\"21\" size=\"10\" color=\"0\">{axisEne}</s></t>\n")

# ============================ Reading input file ============================
print("Please specify the EnePro input file path:")

# For macOS/Linux
fileName = input("(e.g.: /EnePro/example/ChemSci2021.xlsx)\n")
if fileName.strip()[0] == '\'' and fileName.strip()[-1] == '\'':
    fileName = fileName[1:-2]

# For Microsift Windows
#fileName = input("(e.g.: C:\\EnePro\\example\\ChemSci2021.xlsx)\n")

# ============================ Default parameters ============================
digitNumber = 2      # decimal digit number
sheetNumber = 0       # Number of work sheet containing plot data
rowNumbers = 20    # maximum 10 energy potential surfaces
columnNumbers = 25      # maximum 25 species on potential surfaces
boldLineLength = 30      # length of bold energy state lines
dashLineSpan = 0.6      # ratio: bold line length / (bold + dash line length)
energyUnit = 'kJ/mol'      # unit
fontBold = 1      # whether use bold font for state labels

# ============================ User determined parameter ============================
while True:
	print("\n===========================================================================")
	print("                    Energy profile plotting parameters")
	print("---------------------------------------------------------------------------")
	print(f"   1 - Decimal digit number for energy value ({digitNumber})")
	print(f"   2 - Length of energy line ({boldLineLength})")
	print(f"   3 - State span ({dashLineSpan})")
	print(f"   4 - Energy unit ({energyUnit})")
	if fontBold == 1:
		print("   5 - Use bold font for state label (yes)")
	elif fontBold == 0:
		print("   5 - Use bold font for state label (no)")
	print("===========================================================================")
	paraInp = input("Press ENTER to use current settings, input number to modify the parameters.\n")

	if paraInp == '':
		break

	elif paraInp == '1':
		while True:
			try:
				digitNumber = int(input("Specify the decimal digit number for energy value:"))
				break
			except ValueError:
				print("\nInput error, please input a number!")
				continue
		print(f"Decimal digit number for energy value: {digitNumber}")

	elif paraInp == '2':
		while True:
			try:
				boldLineLength = int(input("Specify the length of energy line:"))
				break
			except ValueError:
				print("\nInput error, please input a number!")
				continue
		print(f"Length of energy line: {boldLineLength}")

	elif paraInp == '3':
		while True:
			try:
				dashLineSpan = float(input("Specify the state span:"))
				break
			except ValueError:
				print("\nInput error, please input a number!")
				continue
		print(f"State span: {dashLineSpan}")

	elif paraInp == '4':
		energyUnit = input("Specify the energy unit:")
		print(f"Energy unit will use {energyUnit}.")

	elif paraInp == '5':
		print("*****************************************")
		print("1 - Use bold font for energy label")
		print("2 - DO NOT use bold font for energy label")
		fontBoldUser = input("Input the menu number:")
		if fontBoldUser == '1':
			fontBold = 1
		elif fontBoldUser == '2':
			fontBold = 0
		else:
			print("Input error, EnePro will use default value.")
			fontBold = 1

	else:
		print("Input error, EnePro will use default value.\n")
		break

# ============================ Calculate length of dash line ============================
dashLineLength = int((1 - dashLineSpan) / dashLineSpan * boldLineLength)

# ============================ Read Excel input file ============================
excelDataFile = openpyxl.load_workbook(fileName.strip())
excelDataSheet = excelDataFile.worksheets[sheetNumber]

fullData = []
dataRow = 0
energyMax = -999999999.0
energyMin = 999999999.0
boldLineCoorsX = []
boldLineCoorsY = []

# Save Excel data to list: fullData
for rowNumber in range(1, rowNumbers + 1):
	rowData = []
	for columnNumber in range(1, columnNumbers + 2):
		data = excelDataSheet.cell(rowNumber, columnNumber).value
		if data == None:
			rowData.append('')
		elif isinstance(data, int) and rowNumber%2 == 1:
			rowData.append(format(data, f'.{digitNumber}f'))
		elif isinstance(data, float) and rowNumber%2 == 1:
			rowData.append(format(data, f'.{digitNumber}f'))
		else:
			rowData.append(data)
	fullData.append(rowData)

for energyRowNumber in range(1, rowNumbers + 1, 2):
	fullData[energyRowNumber - 1][0] = fullData[energyRowNumber][0]

for dataLineNumber in range(len(fullData)):
	if fullData[dataLineNumber].count('') != len(fullData[dataLineNumber]):
		dataRow += 1

fullData = fullData[:dataRow]

# Creat a simple data list, without empty value
simpleData = fullData[:]

for i in range(len(fullData)):
	simpleData[i] = [noNone for noNone in simpleData[i] if noNone != '']

# ============================ Displaying process information ============================
print("Loading information from input file...")
print(f"{int(len(simpleData)/2)} energy surfaces detected!\n")
print("Saving energy profile data to .cdxml file...")

# ============================ Find the energy maximum and minimum ============================
for lineSimple in range(0, len(simpleData), 2):
	for noSimple in range(1, len(simpleData[lineSimple])):
		if float(simpleData[lineSimple][noSimple]) >= energyMax:
			energyMax = float(simpleData[lineSimple][noSimple])
		if float(simpleData[lineSimple][noSimple]) <= energyMin:
			energyMin = float(simpleData[lineSimple][noSimple])

# Calculate energy difference
energyDiff = energyMax - energyMin

# ============================ ChemDraw file writting section ============================
# Print bold state lines
for energyRowNumber_2 in range(0, len(fullData), 2):
	boldLineCoorX = []
	boldLineCoorY = []
	for energyColumnNumber in range(1, len(fullData[energyRowNumber_2])):
		if fullData[energyRowNumber_2][energyColumnNumber] != '':
			boldLineY = round(150 + 200 * (energyMax - float(fullData[energyRowNumber_2][energyColumnNumber])) / energyDiff, 3)
			idOrder = energyRowNumber_2 + energyColumnNumber
			boldLineCoorStart = 80 + (energyColumnNumber - 1) * boldLineLength + (energyColumnNumber - 1) * dashLineLength
			boldLineCoorEnd = 80 + energyColumnNumber * boldLineLength + (energyColumnNumber - 1) * dashLineLength
			boldLinePrint(idOrder, idOrder, energyColumnNumber, boldLineY, colorMeter(fullData[energyRowNumber_2][0]), boldLineLength, dashLineLength)
			boldLineCoorX.append(boldLineCoorStart)
			boldLineCoorX.append(boldLineCoorEnd)
			boldLineCoorY.append(boldLineY)
			boldLineCoorY.append(boldLineY)
	boldLineCoorsX.append(boldLineCoorX)
	boldLineCoorsY.append(boldLineCoorY)

# Print dash connecting lines
for boldLineCoorsNo in range(len(boldLineCoorsX)):
	for numberOrder in range(1, len(boldLineCoorsX[boldLineCoorsNo]) - 1, 2):
		idOrder_2 = numberOrder + boldLineCoorsNo
		startX_2 = boldLineCoorsX[boldLineCoorsNo][numberOrder]
		endX_2 = boldLineCoorsX[boldLineCoorsNo][numberOrder + 1]
		startY_2 = boldLineCoorsY[boldLineCoorsNo][numberOrder]
		endY_2 = boldLineCoorsY[boldLineCoorsNo][numberOrder + 1]
		dashLinePrint(idOrder_2, idOrder_2, startX_2, startY_2, endX_2, endY_2, colorMeter(fullData[boldLineCoorsNo * 2][0]), boldLineLength, dashLineLength)

# Print energy values and state labels
for boldLineCoorsNo2 in range(len(boldLineCoorsX)):
	for numberOrder_2 in range(0, len(boldLineCoorsX[boldLineCoorsNo2]), 2):
		textStartX = boldLineCoorsX[boldLineCoorsNo2][numberOrder_2]
		textEndX = boldLineCoorsX[boldLineCoorsNo2][numberOrder_2 + 1]
		textY = boldLineCoorsY[boldLineCoorsNo2][numberOrder_2]
		energyEditor(boldLineCoorsNo2, numberOrder_2, textStartX, textEndX, textY, colorMeter(simpleData[boldLineCoorsNo2 * 2][0]), simpleData[boldLineCoorsNo2 * 2][int(numberOrder_2 / 2.0) + 1])
		nameEditor(boldLineCoorsNo2, numberOrder_2, textStartX, textEndX, textY, colorMeter(simpleData[boldLineCoorsNo2 * 2][0]), simpleData[boldLineCoorsNo2 * 2 + 1][int(numberOrder_2 / 2.0) + 1], fontBold)

# ============================ Frame printing section ============================
xMax = 60

for a in range(len(boldLineCoorsX)):
	if boldLineCoorsX[a][-1] > xMax:
		xMax = boldLineCoorsX[a][-1]
xMax = xMax + 30

# Print y_left axis
eneproFile.write("<graphic\n")
eneproFile.write(" id=\"99999998\"\n")
eneproFile.write(" BoundingBox=\"50 110 50 390\"\n")
eneproFile.write(" Z=\"99999998\"\n")
eneproFile.write(" color=\"0\"\n")
eneproFile.write(" GraphicType=\"Line\"\n")
eneproFile.write("/>\n")

# Print y_right axis
eneproFile.write("<graphic\n")
eneproFile.write(f" id=\"99999997\"\n")
eneproFile.write(f" BoundingBox=\"{xMax} 110 {xMax} 390\"\n")
eneproFile.write(f" Z=\"99999997\"\n")
eneproFile.write(f" color=\"0\"\n")
eneproFile.write(" GraphicType=\"Line\"\n")
eneproFile.write("/>\n")

# Print x_down axis
eneproFile.write("<graphic\n")
eneproFile.write(f" id=\"99999996\"\n")
eneproFile.write(f" BoundingBox=\"50 390 {xMax} 390\"\n")
eneproFile.write(f" Z=\"99999996\"\n")
eneproFile.write(f" color=\"0\"\n")
eneproFile.write(" GraphicType=\"Line\"\n")
eneproFile.write("/>\n")

# Print x_up axis
eneproFile.write("<graphic\n")
eneproFile.write(f" id=\"99999995\"\n")
eneproFile.write(f" BoundingBox=\"50 110 {xMax} 110\"\n")
eneproFile.write(f" Z=\"99999995\"\n")
eneproFile.write(f" color=\"0\"\n")
eneproFile.write(" GraphicType=\"Line\"\n")
eneproFile.write("/>\n")

# Print y_title
eneproFile.write("<t\n")
eneproFile.write(" id=\"99999993\"\n")
eneproFile.write(" p=\"13 250\"\n")
eneproFile.write(" BoundingBox=\"15 100 25 400\"\n")
eneproFile.write(" Z=\"99999993\"\n")
eneproFile.write(" color=\"0\"\n")
eneproFile.write(" CaptionJustification=\"Center\"\n")
eneproFile.write(" Justification=\"Center\"\n")
eneproFile.write(" RotationAngle=\"17694720\"\n")
eneproFile.write(f"><s font=\"21\" size=\"11\" color=\"0\">Relative Energy ({energyUnit})</s></t>\n")

# Print x_title
eneproFile.write("<t\n")
eneproFile.write(" id=\"99999992\"\n")
centerX = xMax / 2 + 25
eneproFile.write(f" p=\"{centerX} 410\"\n")
eneproFile.write(" BoundingBox=\"700 400 30 420\"\n")
eneproFile.write(" Z=\"99999992\"\n")
eneproFile.write(" color=\"0\"\n")
eneproFile.write(" CaptionJustification=\"Center\"\n")
eneproFile.write(" Justification=\"Center\"\n")
eneproFile.write(f"><s font=\"21\" size=\"11\" color=\"0\">Reaction Coordinates</s></t>\n")

# Print y_right axis scale bar
yAxiss = list(axisDevideValue(energyMax, energyMin))
for yAxisNo in range(len(yAxiss)):
	yHeight = round(150 + 200 * (energyMax - float(yAxiss[yAxisNo])) / energyDiff, 3)
	axisDevideLine(yAxisNo, yAxisNo, yHeight)
	axisDevideName(yAxisNo, yAxisNo, yHeight, yAxiss[yAxisNo])

# ============================ End of ChemDraw file ============================
eneproFile.write("</page></CDXML>\n")
eneproFile.close()

# ============================ Result information ============================
print("\n*******************************************************************************")
print("")
print("                   Energy diagram is saved in .cdxml file.")
print("                        Normal termination of EnePro.")
print("")
print("*******************************************************************************\n")
