'''
Brad Barakat
Made for AME-341b

This script is meant to be a Mac workaround for PLX-DAQ since PLX-DAQ uses COM ports (Macs do not have these).
However, this has worked on a Windows computer, so this is not exclusive to Mac.
There are two parts to this script: the port finder, and the serial reader and publisher (to Excel or CSV file).

NOTE: This script does not replicate all of the features of PLX-DAQ! This script was originally made to read data
  serially from an Arduino (see Appendix B in the README for the specific Arduino file), plot the data, and write
  to Excel. Replications for commands like "RESETTIMER" and "CLEARDATA" were added over a year later as an
  afterthought.
'''

# If serial is not installed, type "python3 -m pip install pyserial" into a Terminal window
# Note that if you have another serial library installed, it may interfere with this one
import serial
from serial.tools import list_ports
# If xlsxwriter is not installed, type "pip3 install xlsxwriter" into a Terminal window
import xlsxwriter
# If matplotlib is not installed, type "pip3 install matplotlib" into a Terminal window
from matplotlib import pyplot as plt
# Python has a built-in enum library
from enum import Enum
# Python has a built-in os library
import os
# Python has a built-in datetime library
from datetime import datetime
# Python has a built-in time library
import time
# Python has a built-in traceback library
import traceback


class GraphChoice(Enum):
    LIVE = 0
    EXCEL_ONLY = 1
    NONE = 2


# This function gets a valid integer input from the user
def getValidIntInput(prompt, lBnd=None, hBnd=None):
    valid = False
    while (not valid):
        x = input(prompt).strip()
        try:
            x = int(x)
            if (lBnd is None): lBnd = x
            if (hBnd is None): hBnd = x
            valid = (lBnd <= x) and (x <= hBnd)
            if (not valid): print(f"Error: Integer out of range [{lBnd},{hBnd}]")
        except:
            print("Error: Numeric input not an integer")
    return x


# This function prompts the user to either overwrite the specified file or enter another file name
# It will return the final file name (as a string)
def resolveDupFile(filepath, ext):
    fileOverwrite = False
    while (os.path.exists(filepath)) and (not fileOverwrite):
        print("This file already exists:", filepath)
        fileOverwrite = input("Do you wish to overwrite it? ('y'/'n'): ").upper() == "Y"
        if not fileOverwrite:
            rawFile = input(f"Enter another workbook/file name or path (without the '{ext}' at the end): ")
            filepath = os.path.normpath(rawFile + ext)
    # Make sure the directory exists
    fileDir = os.path.split(filepath)[0]
    if (fileDir != ""):
        # Using the following os methods on "" would throw errors
        if (not os.path.exists(fileDir)):
            os.makedirs(fileDir)
    return filepath


# This function gets a valid file name from the user 
def getValidFileName(saveAsXLSX):
    rawFile = input("Enter workbook/file name or path (without the file-specific extension): ")
    fileName = os.path.normpath(rawFile)
    ext = ".xlsx" if (saveAsXLSX) else ".csv"
    fileName += ext
    fileName = resolveDupFile(fileName, ext)
    return fileName


# This function finds the header of the data and the Arduino delay time between data lines
# It also finds the time the program should pause for after updating the graph
def getHeaderAndDelay(ser, DATA_START_AFTER, user_GC):
    dataStarted = readHeader = time0Found = False
    runHeaderLoop = True
    headerTxt = ""
    print("\nMeasuring delay between Arduino data packets...")
    ser.open()
    while runHeaderLoop:
        dataIn = (ser.readline()).decode().rstrip('\r\n')
        if (not dataStarted):
            if (dataIn.upper() == DATA_START_AFTER):
                dataStarted = True
        elif (not readHeader):
            headerTxt = dataIn # Save the header
            readHeader = True
        elif (not time0Found):
            t0 = time.time()
            time0Found = True
        else:
            t1 = time.time()
            runHeaderLoop = False
    ser.close()
    delayArd = round(t1 - t0, 3)
    # Determine pause time for graph
    graphPause = 0.5*delayArd # Account for data processing time (the 0.5 is arbitrary)
    if (delayArd == 0) and (user_GC == GraphChoice.LIVE):
        # This is unlikely to happen, but I must account for it
        print("No notable Arduino delay between messages. There should be some sort of delay on the order of at least milliseconds.")
        print("If you want to see the live graph, there must be some pause for it to update.")
        print("If a delay is introduced, the graph and data-writing will lag behind, but there will be no gaps in the data stream.")
        addDelay = input("Add a delay of 1 ms? ('y'/'n'): ").upper() == "Y"
        if (addDelay):
            delayArd = 0.001
            graphPause = 0.001
    else:
        print("Delay:", delayArd, "s.")
    return [headerTxt, delayArd, graphPause]


# This function determines what a row is (i.e., LABEL or DATA)
# None is returned if the row is blank
# If the first value is not a row type or directive, then the row is assumed to be the given default
def getRowTypeAndNumCols(rowArr, KEY_WORDS, defaultRow):
    numCols = len(rowArr)
    missingLabel = False
    if (numCols == 0):
        rowType = None
    else:
        col1 = rowArr[0].strip().upper()
        if (col1 == ""):
            rowType = None
        elif (col1 in KEY_WORDS):
            rowType = col1
        else:
            missingLabel = True
            rowType = defaultRow # Set to default
    return [rowType, numCols, missingLabel]


# This function writes text to a file
# This function will allow for more concise if-else statements
def writeToTextFile(file, text, append=True):
    with open(file, "at" if (append) else "wt") as fOut: fOut.write(text)


# This function will be used to overwrite the buffers with None
def fillTwoBufsWithNone(bufSize):
    buf1 = [None]*bufSize
    buf2 = [None]*bufSize # This way, the lists are independent of each other
    return [buf1, buf2]


# This function will get a valid spreadsheet name from the user
def getValidSheetName():
    print("\nRefer to the following website for sheet-naming rules:")
    print("https://support.microsoft.com/en-us/office/rename-a-worksheet-3f1f7148-ee83-404d-8ef0-9ff99fbad1f9\n")
    sheetPrompt = "Enter valid sheet name: "
    isGoodName = False
    BAD_STARTS = {"'"}
    BAD_ENDS = {"'"}
    BAD_CHARS = {"/", "\\", "?", "*", ":", "[", "]"}
    while (not isGoodName):
        sheetName = input(sheetPrompt).strip()
        lenName = len(sheetName)
        # Run checks
        if (lenName > 32) or (lenName == 0): continue
        if (sheetName.lower() == "history"): continue
        if (sheetName[0] in BAD_STARTS): continue
        if (sheetName[-1] in BAD_ENDS): continue
        charSet = {ch for ch in sheetName}
        isGoodName = charSet.isdisjoint(BAD_CHARS)
    return sheetName


# This function adds a sheet to the workbook and formats it
def addAndFormatSheet(workbook, sheetName):
    sheet = None
    while (sheet is None):
        try:
            sheet = workbook.add_worksheet(sheetName)
        except:
            print(f"\nSomething went wrong:\n{traceback.format_exc()}\n")
            sheetName = getValidSheetName()
    # Make columns 1 and 3 (0-indexed) wider this formatting is for BB-DAQ's original purpose, so feel free to change it
    sheet.set_column(1, 1, 15)
    sheet.set_column(3, 3, 15)
    return [sheet, sheetName]


# This function processes a data row, which entails checking for key words, writing to file, and graphing
def processDataRow(rowNum, numCols, row, dataColInd, timeColInd, saveAsXLSX, bufXPlot, bufYPlot, bufInd, bufSize, ax, \
                   graphPause, timerT0, fileName, sheet, format_time, format_timer, format_date, user_GC):
    # Key words
    TIME_WORD = "TIME"
    TIMER_WORD = "TIMER"
    DATE_WORD = "DATE"
    # Begin data processing
    currTime = currData = None # Reset the time and data values
    for col in range(numCols):
        cellData = row[col]
        isTime = isTimer = isDate = False # Used in later if-else statements since cellData will be overwritten
        # Swap out key words with the values
        cellDataUpper = cellData.upper()
        if (cellDataUpper == TIME_WORD):
            isTime = True
            cellData = datetime.now().time()
            cellFormat = format_time
        elif (cellDataUpper == TIMER_WORD):
            isTimer = True
            cellData = round(time.time() - timerT0, 3)
            cellFormat = format_timer
        elif (cellDataUpper == DATE_WORD):
            isDate = True
            cellData = datetime.now().date()
            cellFormat = format_date
        else:
            cellFormat = None
        # Check if the time or date is a graphed value
        isDateTime = (isTime) or (isDate)
        if (isDateTime) and (user_GC != GraphChoice.NONE):
            strData = str(cellData) # Datetime objects can't be plotted
            if (col == dataColInd): currData = strData
            if (col == timeColInd): currTime = strData
        # Write to file accordingly
        if (saveAsXLSX):
            try:
                if (not isDateTime) and (user_GC != GraphChoice.NONE):
                    cellData = float(cellData)
                    if (col == dataColInd): currData = cellData
                    if (col == timeColInd): currTime = cellData
            except:
                pass
            finally:
                sheet.write(rowNum, col, cellData, cellFormat)
        else:
            try:
                if (not isDateTime) and (user_GC != GraphChoice.NONE):
                    # Don't make cellData a float since it will be converted back to string for CSV
                    if (col == dataColInd): currData = float(cellData)
                    if (col == timeColInd): currTime = float(cellData)
            except:
                pass
            finally:
                # The row array is unused after the column iteration, so it can be reused for holding CSV values
                if (isDateTime) or (isTimer): row[col] = str(cellData) # All values in CSV are strings
    if (user_GC == GraphChoice.LIVE):
        # Add data from row to buffer and plot it if buffers are full
        bufXPlot[bufInd] = currTime
        bufYPlot[bufInd] = currData
        bufInd += 1
        if (bufInd == bufSize):
            ax.plot(bufXPlot, bufYPlot, '-b') # Plot values
            bufXPlot[0] = currTime # Put the last plotted values in (for graph continuity)
            bufYPlot[0] = currData
            bufInd = 1 # Reset the index
            if (plt.waitforbuttonpress(graphPause)): raise KeyboardInterrupt # This will wait for your keypress
    # Increment row/line
    if (saveAsXLSX): rowNum += 1
    else: writeToTextFile(fileName, f"{','.join(row)}\n")
    return [sheet, rowNum, bufXPlot, bufYPlot, bufInd, ax]


# This function processes a label row
def processLabelRow(saveAsXLSX, sheet, rowNum, row, fileName, dataIn):
    if (saveAsXLSX):
        sheet.write_row(rowNum, 0, row)
        rowNum += 1
    else:
        writeToTextFile(fileName, f"{dataIn}\n")
    return [sheet, rowNum]


# This function processes a message row
# (Nothing is to be done in this case)
def processMsgRow():
    pass


# This function processes the reset timer directive
def processResetTimer():
    return time.time() # New timerT0


# This function processes the clear data directive
def processClearData(saveAsXLSX, workbook, sheet, sheetName, fileName, header, headerTxt, ax, xLabel, yLabel, bufSize, user_GC):
    if (saveAsXLSX):
        # Overwrite all rows after header then reset rowNum
        # The below is inspired by https://github.com/jmcnamara/XlsxWriter/pull/432/commits/613f2ca7a60018337222f6a07d602e3f28595a36
        workbook.worksheets().remove(sheet)
        [sheet, sheetName] = addAndFormatSheet(workbook, sheetName)
        sheet.write_row(0, 0, header) # Preserve the header to replicate PLX-DAQ's "CLEARDATA"
        rowNum = 1 # Reset row counter to after header
    else:
        workbook = sheet = rowNum = None
        writeToTextFile(fileName, f"{headerTxt}\n", False)
    if (user_GC == GraphChoice.LIVE):
        # Reset graph
        plt.cla()
        ax.set_xlabel(xLabel)
        ax.set_ylabel(yLabel)
        # Reset buffers and index
        [bufXPlot, bufYPlot] = fillTwoBufsWithNone(bufSize)
        bufInd = 0 # Restart the index count (afterwards, reserve 0 for the last plotted value)
    else:
        bufXPlot = bufYPlot = bufInd = None
    return [workbook, sheet, rowNum, bufXPlot, bufYPlot, bufInd, ax]


# This function does the reading of serial data and writing of the output file (the optional parameters are populated internally)
def getAndWriteData(saveAsXLSX, fileName, headerTxt, DATA_DELIM, INTERVAL_PLOT, delayArd, timeColInd, dataColInd, ser, \
                    DATA_START_AFTER, graphPause, user_GC, workbook=None, formatList=None):
    # Directives
    RESET_TIMER = "RESETTIMER"
    CLEAR_DATA = "CLEARDATA"
    #DIRECTIVES = {RESET_TIMER, CLEAR_DATA}
    # Row types
    DATA_ROW = "DATA"
    LABEL_ROW = "LABEL"
    MSG_ROW = "MSG"
    #ROW_TYPES = {DATA_ROW, LABEL_ROW, MSG_ROW}
    # All key words
    #KEY_WORDS = DIRECTIVES.union(ROW_TYPES)
    KEY_WORDS = {RESET_TIMER, CLEAR_DATA, DATA_ROW, LABEL_ROW, MSG_ROW}
    
    # Find how many columns the header has
    header = headerTxt.split(DATA_DELIM)
    numHeaderCols = len(header)

    # Prepare the output files (and related variables)
    if (saveAsXLSX):
        sheetName = getValidSheetName()
        if (workbook is None):
            workbook = xlsxwriter.Workbook(fileName, {'constant_memory': True})
            format_time = workbook.add_format({'num_format': 'hh:mm:ss.000'})
            format_timer = workbook.add_format({'num_format': '0.00'})
            format_date = workbook.add_format({'num_format': 'mm-dd-yyyy'})
        else:
            format_time, format_timer, format_date = formatList
        [sheet, sheetName] = addAndFormatSheet(workbook, sheetName)
        rowNum = 0
    else:
        workbook = sheet = sheetName = rowNum = None
        format_time = format_timer = format_date = None
        writeToTextFile(fileName, "", False)

    if (user_GC == GraphChoice.LIVE):
        # Prepare the list that will be filled with values before being plotted
        # Index 0 will be reserved for last value of the previous plot, so add 1 to the buffer size for that
        bufSize = int(INTERVAL_PLOT/delayArd) + 1 + 1 # The other +1 is to make sure at least 1 new value is plotted (if delayArd > INTERVAL_PLOT)
        [bufXPlot, bufYPlot] = fillTwoBufsWithNone(bufSize)
        bufInd = 0 # Start the index count (afterwards, reserve 0 for the last plotted value)
        # Prepare the figure
        fig, ax = plt.subplots(1,1)
        plt.ion()
        xLabel = header[timeColInd]
        yLabel = header[dataColInd]
        ax.set_xlabel(xLabel)
        ax.set_ylabel(yLabel)
    else:
        bufSize = bufXPlot = bufYPlot = bufInd = fig = ax = xLabel = yLabel = None

    dataStarted = False
    timerT0 = time.time()
    ser.open()
    try:
        print("\nThere are three ways to stop the program:")
        print("  Press any key while the graph window is selected.")
        print("  Press the Reset button on the Arduino.")
        print("  Press Ctrl+C (use as last resort).\n")
        while True:
            # The rows are iterated by the while loop, but columns will be iterated by the for loop
            # Read in a line of data and parse it
            dataIn = (ser.readline()).decode().strip()
            if (not dataStarted):
                dataStarted = (dataIn.upper() == DATA_START_AFTER)
            else:
                print(dataIn)
                row = dataIn.split(DATA_DELIM)
                [rowType, numCols, missingLabel] = getRowTypeAndNumCols(row, KEY_WORDS, DATA_ROW)
                rowIsData = (rowType == DATA_ROW)
                rowIsMsg = (rowType == MSG_ROW)
                # Check if the data stopped coming in
                if (rowType is None):
                    print("\nSerial timed out.")
                    raise KeyboardInterrupt
                # If the label is missing, add it
                if (missingLabel) and (not rowIsMsg):
                    if (saveAsXLSX):
                        row = [rowType] + row
                        numCols += 1
                    else:
                        writeToTextFile(fileName, f"{rowType},")
                # Perform actions depending on the row type
                if (rowIsData):
                    [sheet, rowNum, bufXPlot, bufYPlot, bufInd, ax] = processDataRow(rowNum, numCols, row, dataColInd, \
                        timeColInd, saveAsXLSX, bufXPlot, bufYPlot, bufInd, bufSize, ax, graphPause, timerT0, fileName, \
                        sheet, format_time, format_timer, format_date, user_GC)
                elif (rowType == RESET_TIMER):
                    timerT0 = processResetTimer()
                elif (rowType == CLEAR_DATA):
                    [workbook, sheet, rowNum, bufXPlot, bufYPlot, bufInd, ax] = processClearData(saveAsXLSX, workbook, \
                        sheet, sheetName, fileName, header, headerTxt, ax, xLabel, yLabel, bufSize, user_GC)
                elif (rowIsMsg):
                    processMsgRow()
                elif (rowType == LABEL_ROW):
                    [sheet, rowNum] = processLabelRow(saveAsXLSX, sheet, rowNum, row, fileName, dataIn)
                else:
                    print(f"Unexpected row type: {rowType}") # This line should not be reached, so it's good for troubleshooting
    except KeyboardInterrupt:
        print("\nExiting...")
    except:
        print(f"\nSomething went wrong:\n{traceback.format_exc()}\n")
    finally:
        ser.close()
        if (user_GC == GraphChoice.LIVE): plt.close(fig)
        
    # Give the user the option to run BB-DAQ again with the same settings (but in a new file/worksheet)
    print("\nWould you like to run BB-DAQ again with the same settings, but with the output in a new file/worksheet?")
    rerunPrompt = "Enter 0 to exit, or enter 1 to run again: "
    runAgain = (getValidIntInput(rerunPrompt, 0, 1) == 1)

    if (saveAsXLSX):
        if (user_GC != GraphChoice.NONE):
            # Create a new chart object before closing the workbook
            capitalA_Int = ord("A")
            timeCol = chr(capitalA_Int + timeColInd)
            dataCol = chr(capitalA_Int + dataColInd)
            chartCol = chr(capitalA_Int + numHeaderCols + 1)
            chart = workbook.add_chart({'type': 'line'})
            finalRowStr = str(rowNum)
            chart.add_series({
                'categories': f'={sheetName}!${timeCol}$2:${timeCol}${finalRowStr}',
                'values':     f'={sheetName}!${dataCol}$2:${dataCol}${finalRowStr}',
            })
            chart.set_x_axis({'name': f'={sheetName}!${timeCol}$1'})
            chart.set_y_axis({'name': f'={sheetName}!${dataCol}$1'})
            chart.set_legend({'none': True})
            # Insert the chart into the worksheet
            sheet.insert_chart(chartCol + '2', chart)
        if (runAgain):
            rerunPromptXlsx = "Enter 0 to make a new worksheet in the same workbook, or enter 1 to make a new workbook: "
            newWkbk = (getValidIntInput(rerunPromptXlsx, 0, 1) == 1)
            if (newWkbk):
                workbook.close()
                fileName = getValidFileName(saveAsXLSX)
                workbook = None
            formatList = [format_time, format_timer, format_date]
        else:
            workbook.close()
    else:
        if (runAgain):
            fileName = getValidFileName(saveAsXLSX)
            formatList = None # Prepare for function call (workbook is already None from beginning for CSV)
    
    # Run again (generalized for both cases)
    if (runAgain):
        # Now we can use the optional parameters in the function call
        getAndWriteData(saveAsXLSX, fileName, headerTxt, DATA_DELIM, INTERVAL_PLOT, delayArd, timeColInd, dataColInd, \
            ser, DATA_START_AFTER, graphPause, user_GC, workbook, formatList)


# main()
def main():
    # Constants
    INTERVAL_PLOT = 0.5 # Minimum number of seconds before plot is updated (semi-arbitrary)
    DATA_START_AFTER = "CLEARDATA"
    DATA_DELIM = ","
    
    # This first part will find the available serial ports. The Arduino should be a USB port.
    # To be sure of the Arduino's port, run this part before and after plugging in the Arduino, and compare the
    # output. To minimize confusion, make sure no other devices are also being plugged in between the two runs.
    portList = list_ports.comports() # Outputs a list
    p_ind = 0
    print("Ports:")
    for p in portList:
        print(f"{p_ind}: {p.device}")
        p_ind += 1
    portPrompt = "Enter the index of the port you want to use, or -1 to exit: "
    portChoice = getValidIntInput(portPrompt, -1, p_ind-1) # At this point, p_ind = len(portList)

    if (portChoice == -1):
        print("Exiting...")
        return
    
    # This second part will actually read the serial data from the Arduino and write it to a file.
    # A live graph of the numerical data will also be generated.

    # Check to see if the user wants the live graph
    print("")
    graphPrompt = "Enter 0 to see the live graph, 1 to see the graph only in the Excel output, or 2 to not see the graph at all: "
    user_GC = GraphChoice(getValidIntInput(graphPrompt, 0, 2))

    # Get port info from user
    print("")
    port = portList[portChoice].device
    buad = getValidIntInput("Enter the buad rate: ", 1)
    # See the rest of serial.Serial()'s parameters here:
    # https://pyserial.readthedocs.io/en/latest/pyserial_api.html#serial.Serial.__init__
    ser = serial.Serial(port, buad)
    # Close the port in case it is already open (this can happen when a serial connection isn't closed gracefully)
    ser.close()

    # Find the header and delay time between data (and for graph)
    [headerTxt, delayArd, graphPause] = getHeaderAndDelay(ser, DATA_START_AFTER, user_GC)
    TO_time = 1.25*delayArd # Amount of time before timeout on serial read
    header = headerTxt.split(DATA_DELIM)

    print(f"\nHeader:\n{headerTxt}\n")
    # Ask plot questions if the graph will appear at any point
    if (user_GC == GraphChoice.LIVE) or ((user_GC == GraphChoice.EXCEL_ONLY) and (saveAsXLSX)):
        timePrompt = "Enter the column index (start at 0) for the x-axis in the transmitted data: "
        dataPrompt = "Enter the column index (start at 0) for the y-axis in the transmitted data: "
        colHBnd = len(header) - 1
        timeColInd = getValidIntInput(timePrompt, 0, colHBnd)
        dataColInd = getValidIntInput(dataPrompt, 0, colHBnd)
    else:
        timeColInd = dataColInd = None
    choicePrompt = "Enter 0 to save as an Excel workbook, or enter 1 to save as a CSV file: "
    saveAsXLSX = (getValidIntInput(choicePrompt, 0, 1) == 0)

    fileName = getValidFileName(saveAsXLSX)
    
    # Get and write data
    ser = serial.Serial(port, buad, timeout=TO_time)
    ser.close()
    getAndWriteData(saveAsXLSX, fileName, headerTxt, DATA_DELIM, INTERVAL_PLOT, delayArd, timeColInd, dataColInd, ser, \
        DATA_START_AFTER, graphPause, user_GC)
    # Print confirmation
    print("Done.")


# Run main()
if (__name__ == "__main__"):
    main()
