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
# Python has a built-in os library
import os
# Python has a built-in datetime library
from datetime import datetime
# Python has a built-in time library
import time
# Python has a built-in traceback library
import traceback


# This function gets a valid integer input from the user
def getValidIntInput(prompt, lBnd=None, hBnd=None):
    valid = False;
    while (not valid):
        x = input(prompt).strip();
        try:
            x = int(x);
            if (x == round(x, 0)):
                if (lBnd == None): lBnd = x - 1;
                if (hBnd == None): hBnd = x + 1;
                valid = (lBnd <= x) and (x <= hBnd);
                if (not valid): print(f"Error: Integer out of range [{lBnd},{hBnd}]");
        except:
            print("Error: Numeric input not an integer");
    return x;


# This function prompts the user to either overwrite the specified file or enter another file name
# It will return the final file name (as a string)
def resolveDupFile(filepath, ext):
    fileOverwrite = False;
    while (os.path.exists(filepath)) and (not fileOverwrite):
        print("This file already exists:", filepath);
        fileOverwrite = input("Do you wish to overwrite it? ('y'/'n'): ").upper() == "Y";
        if not fileOverwrite:
            filepath = os.path.normpath(input("Enter another workbook/file name or path (without the '" + ext + "' at the end): ") + ext);
    # Make sure the directory exists
    fileDir = os.path.split(filepath)[0];
    if (fileDir != ""):
        # Using the following os methods on "" would throw errors
        if (not os.path.exists(fileDir)):
            os.makedirs(fileDir);
    return filepath;


# This function finds the header of the data and the Arduino delay time between data lines
# It also finds the time the program should pause for after updating the graph
def getHeaderAndDelay(ser, DATA_START_AFTER):
    dataStarted = readHeader = time0Found = False;
    runHeaderLoop = True;
    headerTxt = "";
    print("\nMeasuring delay between Arduino data packets...");
    ser.open();
    while runHeaderLoop:
        dataIn = (ser.readline()).decode().rstrip('\r\n');
        if (not dataStarted):
            if (dataIn.upper() == DATA_START_AFTER):
                dataStarted = True;
        elif (not readHeader):
            headerTxt = dataIn; # Save the header
            readHeader = True;
        elif (not time0Found):
            t0 = time.time();
            time0Found = True;
        else:
            t1 = time.time();
            runHeaderLoop = False;
    ser.close();
    delayArd = round(t1 - t0, 3);
    # Determine pause time for graph
    graphPause = 0.5*delayArd; # Account for data processing time (the 0.5 is arbitrary)
    if (delayArd == 0):
        # This is unlikely to happen, but I must account for it
        print("No notable Arduino delay between messages. There should be some sort of delay on the order of at least milliseconds.");
        print("If you want to see the live graph, there must be some pause for it to update.");
        print("If a delay is introduced, the graph and data-writing will lag behind, but there will be no gaps in the data stream.");
        addDelay = input("Add a 1 ms delay? ('y'/'n'): ").upper() == "Y";
        if (addDelay):
            delayArd = 0.001;
            graphPause = 0.001;
    else:
        print("Delay:", delayArd, "s.");
    return [headerTxt, delayArd, graphPause];


# This function determines what a row is (i.e., LABEL or DATA)
# None is returned if the row is blank
# If the first value is a number, then the row is assumed to be DATA
def getRowTypeAndNumCols(rowArr):
    rowType = "DATA"; # Default
    numCols = len(rowArr);
    if (numCols == 0):
        rowType == None;
    else:
        col1 = rowArr[0].strip();
        if (col1 == ""): rowType == None;
        elif (not col1[0].isnumeric()): rowType = col1.upper();
    return [rowType, numCols];


# This function writes text to a file
# This function will allow for more concise if-else statements
def writeToTextFile(file, text, append=True):
    mode = "at" if (append) else "wt";
    with open(file, mode) as fOut: fOut.write(text);


# This function will be used to overwrite the buffers with None
def fillTwoBufsWithNone(bufSize):
    buf1 = [None]*bufSize;
    buf2 = buf1.copy(); # This way, the arrays are independent of each other
    return [buf1, buf2];


# This function does the reading of serial data and writing of the output file
def getAndWriteData(saveChoice, fileName, headerTxt, DATA_DELIM, INTERVAL_PLOT, delayArd, timeColInd, dataColInd, ser, DATA_START_AFTER, TIMER_RESET, graphPause):
    if (saveChoice == 0):
        sheetName = "Data";
        workbook = xlsxwriter.Workbook(fileName);
        sheet = workbook.add_worksheet(sheetName);
        rowNum = 0;
        format_time = workbook.add_format({'num_format': 'hh:mm:ss.000'});
        sheet.set_column(1, 1, 15);
        sheet.set_column(3, 3, 15);
    else:
        writeToTextFile(fileName, "", False);

    # Prepare the list that will be filled with values before being plotted
    # Index 0 will be reserved for last value of the previous plot, so add 1 to the buffer size for that
    bufSize = int(INTERVAL_PLOT/delayArd) + 1 + 1; # The other +1 is to make sure at least 1 new value is plotted
    [bufXPlot, bufYPlot] = fillTwoBufsWithNone(bufSize);
    bufInd = 0; # Start the index count (afterwards, reserve 0 for the last plotted value)

    header = headerTxt.split(DATA_DELIM);
    numHeaderCols = len(header);
    xLabel = header[timeColInd];
    yLabel = header[dataColInd];

    fig, ax = plt.subplots(1,1); plt.ion();
    ax.set_xlabel(xLabel); ax.set_ylabel(yLabel);

    dataStarted = False;
    timerT0 = time.time();
    ser.open();
    try:
        print("\nThere are three ways to stop the program:");
        print("  Press any key while the graph window is selected.");
        print("  Press the Reset button on the Arduino.");
        print("  Press Ctrl+C (use as last resort).\n");
        while True:
            # The rows are iterated by the while loop, but columns will be iterated by the for loop
            currTime = currData = None; # Reset the time and data values
            dataIn = (ser.readline()).decode().strip();
            if (not dataStarted):
                dataStarted = (dataIn.upper() == DATA_START_AFTER);
            else:
                print(dataIn);
                row = dataIn.split(DATA_DELIM);
                [rowType, numCols] = getRowTypeAndNumCols(row);
                rowIsData = (rowType == "DATA");
                if ((numCols < numHeaderCols) and (rowIsData)) or (rowType == None):
                    print("\nSerial timed out.");
                    raise KeyboardInterrupt;
                bypassRow = False; # Used to avoid using "break;" in the for loop
                for col in range(numCols):
                    if (not bypassRow):
                        cellData = row[col];
                        isTime = False; # Used in later if-else statements since cellData will be overwritten
                        # Swap out key words with the values
                        if (cellData.upper() == "TIMER"): cellData = round(time.time() - timerT0, 3);
                        elif (cellData.upper() == "TIME"):
                            isTime = True;
                            cellData = datetime.now().time();
                            strData = str(cellData); # Datetime objects can't be plotted
                            if (col == dataColInd): currData = strData;
                            if (col == timeColInd): currTime = strData;
                        # Write to file accordingly
                        if (rowIsData):
                            if (saveChoice == 0):
                                try:
                                    if (not isTime):
                                        cellData = float(cellData);
                                        if (col == dataColInd): currData = cellData;
                                        if (col == timeColInd): currTime = cellData;
                                except:
                                    pass;
                                finally:
                                    if (isTime): sheet.write(rowNum, col, cellData, format_time);
                                    else: sheet.write(rowNum, col, cellData);
                            else:
                                try:
                                    if (not isTime):
                                        # Don't make cellData a float since it will be converted back to string for CSV
                                        if (col == dataColInd): currData = float(cellData);
                                        if (col == timeColInd): currTime = float(cellData);
                                except:
                                    pass;
                                finally:
                                    if (isTime): cellData = str(cellData);
                                    writeToTextFile(fileName, f"{cellData},");
                        elif (rowType == TIMER_RESET):
                            timerT0 = time.time();
                            bypassRow = True; # Move on to the next row
                        elif (rowType == DATA_START_AFTER):
                            if (saveChoice == 0):
                                # Overwrite all rows after header then reset rowNum
                                blankCol = [None]*(rowNum-2); # Current row is empty, and ignore header row
                                for col_i in range(numHeaderCols):
                                    sheet.write_column(1, col_i, blankCol, workbook.add_format());
                                rowNum = 1; # The row incrementer will be bypassed, so start at 1
                            else:
                                writeToTextFile(fileName, f"{headerTxt}\n", False);
                            # Reset graph
                            plt.cla(); ax.set_xlabel(xLabel); ax.set_ylabel(yLabel);
                            # Reset buffers and index
                            [bufXPlot, bufYPlot] = fillTwoBufsWithNone(bufSize);
                            bufInd = 0; # Restart the index count (afterwards, reserve 0 for the last plotted value)
                            bypassRow = True; # Move on to the next row
                        else:
                            if (saveChoice == 0): sheet.write(rowNum, col, cellData);
                            else: writeToTextFile(fileName, f"{cellData},");
                # Check to see if the row has been bypassed before writing and plotting data
                if (not bypassRow):
                    # Increment row/line if the current row is not bypassed
                    if (saveChoice == 0): rowNum += 1;
                    else: writeToTextFile(fileName, "\n");
                    # Add data to buffer and/or plot it
                    if (rowIsData):
                        bufXPlot[bufInd] = currTime; bufYPlot[bufInd] = currData;
                        bufInd += 1;
                        if (bufInd == bufSize):
                            ax.plot(bufXPlot, bufYPlot, '-b'); # Plot values
                            bufXPlot[0] = currTime; bufYPlot[0] = currData; # Put the last plotted values in (for graph continuity)
                            bufInd = 1; # Reset the index
                            if (plt.waitforbuttonpress(graphPause)): raise KeyboardInterrupt; # This will wait for your keypress
    except KeyboardInterrupt:
        print("\nExiting...");
    except:
        print(f"\nSomething went wrong:\n{traceback.format_exc()}\n");
    finally:
        ser.close(); plt.close(fig);
        
    if (saveChoice == 0):
        # Create a new chart object before closing the workbook
        capitalA_Int = ord("A");
        timeCol = chr(capitalA_Int + timeColInd);
        dataCol = chr(capitalA_Int + dataColInd);
        chartCol = chr(capitalA_Int + numHeaderCols + 1);
        chart = workbook.add_chart({'type': 'line'});
        finalRowStr = str(rowNum);
        chart.add_series({
            'categories': f'={sheetName}!${timeCol}$2:${timeCol}${finalRowStr}',
            'values':     f'={sheetName}!${dataCol}$2:${dataCol}${finalRowStr}',
        });
        chart.set_x_axis({'name': f'={sheetName}!${timeCol}$1'});
        chart.set_y_axis({'name': f'={sheetName}!${dataCol}$1'});
        chart.set_legend({'none': True});
        # Insert the chart into the worksheet
        sheet.insert_chart(chartCol + '2', chart);
        workbook.close();


# main()
def main():
    # Constants
    INTERVAL_PLOT = 0.5; # Minimum number of seconds before plot is updated (semi-arbitrary)
    DATA_START_AFTER = "CLEARDATA";
    TIMER_RESET = "RESETTIMER";
    DATA_DELIM = ",";
    
    # This first part will find the available serial ports. The Arduino should be a USB port.
    # To be sure of the Arduino's port, run this part before and after plugging in the Arduino, and compare the
    # output. To minimize confusion, make sure no other devices are also being plugged in between the two runs.
    portList = list(list_ports.comports());
    p_ind = 0;
    print("Ports:");
    for p in portList:
        print(f"{p_ind}: {p.device}");
        p_ind += 1;
    portPrompt = "Enter the index of the port you want to use, or -1 to exit.\nChoice: ";
    portChoice = getValidIntInput(portPrompt, -1, len(portList)-1);

    if (portChoice == -1):
        print("Exiting...");
    else:
        # This second part will actually read the serial data from the Arduino and write it to a file.
        # A live graph of the numerical data will also be generated.

        # Get port info from user
        port = portList[portChoice].device;
        buad = getValidIntInput("Enter the buad rate: ", 1);
        # See the rest of serial.Serial()'s parameters here:
        # https://pyserial.readthedocs.io/en/latest/pyserial_api.html#serial.Serial.__init__
        ser = serial.Serial(port, buad);
        ser.close();

        # Find the header and delay time between data (and for graph)
        [headerTxt, delayArd, graphPause] = getHeaderAndDelay(ser, DATA_START_AFTER);
        TO_time = 1.25*delayArd; # Amount of time before timeout on serial read
        header = headerTxt.split(DATA_DELIM);

        print("\nHeader:\n" + headerTxt + "\n");
        timePrompt = "Enter the column index (start at 0) for the x-axis in the transmitted data: ";
        dataPrompt = "Enter the column index (start at 0) for the y-axis in the transmitted data: ";
        choicePrompt = "Enter 0 to save as an Excel workbook, or enter 1 to save as a CSV file: ";
        colHBnd = len(header) - 1;
        timeColInd = getValidIntInput(timePrompt, 0, colHBnd);
        dataColInd = getValidIntInput(dataPrompt, 0, colHBnd);
        saveChoice = getValidIntInput(choicePrompt, 0, 1);

        fileName = os.path.normpath(input("Enter workbook/file name or path (without the file-specific extension): "));
        ext = ".xlsx" if (saveChoice == 0) else ".csv";
        fileName += ext;
        fileName = resolveDupFile(fileName, ext);
        
        # Get and write data
        ser = serial.Serial(port, buad, timeout=TO_time);
        ser.close();
        getAndWriteData(saveChoice, fileName, headerTxt, DATA_DELIM, INTERVAL_PLOT, delayArd, timeColInd, dataColInd, ser, DATA_START_AFTER, TIMER_RESET, graphPause);
        # Print confirmation
        print("Done.");


# Run main()
main();
