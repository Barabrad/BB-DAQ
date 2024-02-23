'''
Brad Barakat
Made for AME-341b

This script is meant to be a Mac workaround for PLX-DAQ since PLX-DAQ uses COM ports (Macs do not have these).
However, this has worked on a Windows computer, so this is not exclusive to Mac.
There are two parts to this script: the port finder, and the serial reader and publisher (to Excel or CSV file).
'''

# If serial is not installed, type "python3 -m pip install pyserial" into a Terminal window
# Note that if you have another serial library installed, it may interfere with this one
import serial
from serial.tools import list_ports
# If xlsxwriter is not installed, type "pip3 install xlsxwriter" into a Terminal window
import xlsxwriter
# If matplotlib is not installed, type "pip3 install matplotlib" into a Terminal window
from matplotlib import pyplot as plt
# Python has a built-in datetime library
from datetime import datetime
# Python has a built-in time library
import time
# Python has a built-in traceback library
import traceback


# This function returns a boolean based on if a given file already exists
def fileAlreadyExists(filepath):
    exists = True;
    try:
        with open(filepath, "rt") as f:
            pass;
    except:
        exists = False;
    return exists;


# This function prompts the user to either overwrite the specified file or enter another file name
# It will return the final file name (as a string)
def resolveDupFile(filepath, ext):
    fileOverwrite = False;
    while fileAlreadyExists(filepath) and not fileOverwrite:
        print("This file already exists:", filepath);
        fileOverwrite = input("Do you wish to overwrite it? ('y'/'n'): ").upper() == "Y";
        if not fileOverwrite:
            filepath = input("Enter another workbook/file name or path (without the '" + ext + "' at the end): ") + ext;
    return filepath;


# This main() function could be broken up more, but given that this is not an assignment, I can relax the standards a bit.
def main():
    # Constants
    INTERVAL_PLOT = 0.5; # Minimum number of seconds before plot is updated (semi-arbitrary)
    
    # This first part will find the available serial ports. The Arduino should be a USB port.
    # To be sure of the Arduino's port, run this part before and after plugging in the Arduino, and compare the
    # output. To minimize confusion, make sure no other devices are also being plugged in between the two runs.
    portList = list(list_ports.comports());
    p_ind = 0;
    print("Ports:");
    for p in portList:
        print(str(p_ind) + ": " + p.device);
        p_ind += 1;
    portChoice = int(input("Enter the index of the port you want to use, or -1 to exit.\nChoice: "));

    if portChoice == -1:
        print("Exiting...");
    elif portChoice not in range(len(portList)):
        print("Invalid index. Exiting...");
    else:
        # This second part will actually read the serial data from the Arduino and write it to a file.
        # A live graph of the numerical data will also be generated.
        dataStartAfter = "CLEARDATA";
        delim = ",";
        dataStarted = False;
        readHeader = False;
        time0Found = False;
        runHeaderLoop = True;

        # Get port info from user
        port = portList[portChoice].device;
        buad = int(input("Enter the buad rate: "));
        # See the rest of serial.Serial()'s parameters here:
        # https://pyserial.readthedocs.io/en/latest/pyserial_api.html#serial.Serial.__init__
        ser = serial.Serial(port, buad);
        ser.close();
        ser.open();

        # Find the header and Arduino delay time between data
        headerTxt = "";
        print("\nMeasuring delay between Arduino data packets...");
        while runHeaderLoop:
            dataIn = (ser.readline()).decode().rstrip('\r\n');
            if not dataStarted:
                if dataIn.upper() == dataStartAfter:
                    dataStarted = True;
            elif not readHeader:
                headerTxt = dataIn; # Save the header
                readHeader = True;
            elif not time0Found:
                t0 = time.time();
                time0Found = True;
            else:
                t1 = time.time();
                runHeaderLoop = False;
        ser.close();
        dataStarted = False; # Reset this for the actual reading

        delayArd = round(t1 - t0, 3);
        graphPause = 0.5*delayArd; # Account for data processing time
        if delayArd == 0:
            # This is unlikely to happen, but I must account for it
            print("No notable Arduino delay between messages. There should be some sort of delay on the order of at least milliseconds.");
            print("If you want to see the live graph, there must be some pause for it to update.");
            print("If a delay is introduced, the graph and data-writing will lag behind, but there will be no gaps in the data stream.");
            addDelay = input("Add a 1 ms delay? ('y'/'n'): ").upper() == "Y";
            if addDelay:
                delayArd = 0.001;
                graphPause = 0.001;
        else:
            print("Delay:", delayArd, "s.");
        TO_time = 1.25*delayArd; # Amount of time before timeout on serial read

        # Prepare the list that will be filled with values before being plotted
        # Index 0 will be reserved for last value, so add 1 to the buffer size for that
        bufSize = int(INTERVAL_PLOT/delayArd) + 1 + 1; # The other +1 is to make sure at least 1 new value is plotted
        bufXPlot = [None]*bufSize;
        bufYPlot = [None]*bufSize; # This way, the X and Y arrays are not linked
        bufInd = 1; # Start the index count (reserving 0 for the last plotted value)

        print("\nHeader:\n" + headerTxt + "\n");
        timeColInd = int(input("Enter the column index (start at 0) for the x-axis in the transmitted data: "));
        dataColInd = int(input("Enter the column index (start at 0) for the y-axis in the transmitted data: "));
        saveChoice = int(input("Enter 0 to save as an Excel workbook, or 1 to save as a CSV file.\nChoice: "));
        fileName = input("Enter workbook/file name or path (without the file-specific extension): ");
        fileOverwrite = False;

        if saveChoice == 0:
            ext = ".xlsx";
            fileName += ext;
            fileName = resolveDupFile(fileName, ext);
            sheetName = "Data";
            workbook = xlsxwriter.Workbook(fileName);
            sheet = workbook.add_worksheet(sheetName);
            rowNum = 0;
            format_time = workbook.add_format({'num_format': 'hh:mm:ss.000'});
            sheet.set_column(1, 1, 15);
            sheet.set_column(3, 3, 15);
        else:
            ext = ".csv";
            fileName += ext;
            resolveDupFile(fileName, ext);
            with open(fileName, "wt") as f:
                pass;

        fig, ax1 = plt.subplots(1,1);
        plt.ion();
        header = headerTxt.split(delim);
        numHeaderCols = len(header);
        ax1.set_xlabel(header[timeColInd]);
        ax1.set_ylabel(header[dataColInd]);
        currTime = None;
        currData = None;

        ser = serial.Serial(port, buad, timeout=TO_time);
        ser.close();
        ser.open();
        try:
            print("\nThere are three ways to stop the program:");
            print("  Press any key while the graph window is selected.");
            print("  Press the Reset button on the Arduino.");
            print("  Press Ctrl+C (use as last resort).\n");
            while True:
                dataIn = (ser.readline()).decode().rstrip('\r\n');
                if not(dataStarted):
                    if dataIn.upper() == dataStartAfter:
                        dataStarted = True;
                else:
                    print(dataIn);
                    row = dataIn.split(delim);
                    numCols = len(row);
                    if (numCols < numHeaderCols):
                        print("\nSerial timed out.")
                        raise KeyboardInterrupt;
                    for col in range(len(row)):
                        cellData = row[col];
                        if saveChoice == 0:
                            try:
                                cellData = float(cellData);
                                if col == dataColInd:
                                    currData = cellData;
                                if col == timeColInd:
                                    currTime = cellData;
                                sheet.write(rowNum, col, cellData);
                            except:
                                if cellData.upper() == "TIME":
                                    cellData = datetime.now().time();
                                    sheet.write(rowNum, col, cellData, format_time);
                                else:
                                    sheet.write(rowNum, col, cellData);
                        else:
                            try:
                                if col == dataColInd:
                                    currData = float(cellData);
                                if col == timeColInd:
                                    currTime = float(cellData);
                            except:
                                pass;
                            if cellData.upper() == "TIME":
                                cellData = str(datetime.now().time());
                            with open(fileName, "at") as f:
                                f.write(cellData);
                                f.write(",");
                    if saveChoice == 0:
                        rowNum += 1;
                    else:
                        with open(fileName, "at") as f:
                            f.write("\n");

                    # Add to buffer and/or plot
                    bufXPlot[bufInd] = currTime;
                    bufYPlot[bufInd] = currData;
                    bufInd += 1;
                    if (bufInd == bufSize):
                        ax1.plot(bufXPlot, bufYPlot, '-b');
                        # Put the last plotted values in (to make sure graph is continuous)
                        bufXPlot[0] = currTime;
                        bufYPlot[0] = currData;
                        # Reset the index
                        bufInd = 1;
                        # This will wait for your keypress
                        if plt.waitforbuttonpress(graphPause):
                            raise KeyboardInterrupt;
        except KeyboardInterrupt:
            print("\nExiting...");
        except:
            print("\nSomething went wrong:");
            print(traceback.format_exc(), "\n");
        finally:
            ser.close();
            plt.close();
            
        if saveChoice == 0:
            # Create a new chart object before closing the workbook
            capitalA_Int = ord("A");
            timeCol = chr(capitalA_Int + timeColInd);
            dataCol = chr(capitalA_Int + dataColInd);
            chartCol = chr(capitalA_Int + numHeaderCols + 1);
            chart = workbook.add_chart({'type': 'line'});
            finalRowStr = str(rowNum);
            chart.add_series({
                'categories': '=' + sheetName + '!$' + timeCol + '$2:$' + timeCol + '$' + finalRowStr,
                'values':     '=' + sheetName + '!$' + dataCol + '$2:$' + dataCol + '$' + finalRowStr,
            });
            chart.set_x_axis({'name': '=' + sheetName + '!$' + timeCol + '$1'});
            chart.set_y_axis({'name': '=' + sheetName + '!$' + dataCol + '$1'});
            chart.set_legend({'none': True});
            # Insert the chart into the worksheet
            sheet.insert_chart(chartCol + '2', chart);
            workbook.close();

        print("Done.");


# Run main()
main();
