# BB-DAQ/src
This folder contains `BB_DAQ.py` and `BB_BoardTester.py`. The former is the PLX-DAQ workaround, and the latter is a file to determine the properties of the board/microcontroller being used.
Both files have a couple of functions in common, but I did not make a third file with them because I wanted both files to be stand-alone. That way, there would be fewer dependencies to worry about when incorporating either of these files into a project.

## Documentation

### Introduction
This script is meant to be a (limited) Mac workaround for PLX-DAQ, an Excel file with a macro that uses COM ports (Macs do not have these). (See [**Warning**](#warning) for specifics on the limitations of BB-DAQ.) However, this script has worked on Windows, so it is not exclusive to Mac. **BB-DAQ does not interface with PLX-DAQ, so there is no need to download the latter.** Although there are comments in the code, I figured a document with a tutorial and warnings would be better. In this document, "terminal window" (for Mac) will mean "command prompt" for Windows.

I found out in Spring 2024 that different boards behave differently when the serial connection is closed. The Arduino Uno R3 (the board used in 2023) effectively resets, which my code takes for granted, but the Arduino Uno R4 Minima (the board used in 2024) does not. This difference will cause `BB_DAQ.py` to get stuck waiting for the "CLEARDATA" that marks the beginning of the serial stream when the Uno R4 Minima is used. I found a quick way to fix this on the user end, and I made a script (`BB_BoardTester.py`) to determine if any boards used in the future are similar to the Uno R3 or the Uno R4 Minima (theoretically, the board being tested might not even be an Arduino). In the tutorial below, "[**If R4** ...]" will contain instructions necessary for boards in the latter category. See [**Appendix A**](#appendix-a-bb-boardtester-tutorial) for the BB_BoardTester tutorial.

### Libraries
The libraries this script uses are listed below, as well as the download instructions. **My assumption is that you already have Python 3 installed on your computer.** To check, open a terminal window and type `python3 -V`. If the output does not display a version number, try `python -V`. If the latter command works, use `python` and `pip` instead of `python3` and `pip3`, respectively. If neither command shows a version number, install Python 3 first, and then return here. To see which non-built-in libraries are already installed, open a terminal window and type `pip3 list`. Depending on your system, you may need to make a [virtual environment](https://docs.python.org/3/library/venv.html) to install these libraries. The `requirements.txt` file contains the non-built-in libraries, so you could type `pip3 install requirements.txt`.
1. pyserial
    * This library is imported in the code as "serial" instead of "pyserial," so make sure you do not have another library installed named "serial." If you do, open a terminal, uninstall it while running the files in this repository, and reinstall it afterwards (or if you're comfortable enough with Python, create a [virtual environment](https://docs.python.org/3/library/venv.html) for running the files in this repository).
    * This library is not built-in, so you need to open a terminal window and enter `python3 -m pip install pyserial` if you do not have the library.
2. xlsxwriter
    * This library is not built-in, so you need to open a terminal window and enter `pip3 install xlsxwriter` if you do not have the library.
3. matplotlib
    * This library is not built-in, so you need to open a terminal window and enter `pip3 install matplotlib` if you do not have the library.
4. enum
    * This library is built-in, so you should not need to install anything.
5. os
    * This library is built-in, so you should not need to install anything.
6. datetime
    * This library is built-in, so you should not need to install anything.
7. time
    * This library is built-in, so you should not need to install anything.
8. traceback
    * This library is built-in, so you should not need to install anything.

### Warning
**This script does not replicate all of the features of PLX-DAQ!** This script was originally made to read data serially from an Arduino (see [**Appendix B**](#appendix-b-arduino-code) for the specific Arduino file), plot the data, and write to Excel. Replications for commands like "RESETTIMER" and "CLEARDATA" were added over a year later as an afterthought. See [**Current Key Words**](#current-key-words) for the current list of PLX-DAQ directives and special data strings this code can replicate.

**You cannot append to an existing Excel file with BB-DAQ!** Using an existing file will overwrite it (the code will warn you). However, you will be given the option at the end of a run to collect more data on a new sheet (also referred to as "worksheet") within the Excel file (also referred to as "workbook") that is being created.

Also, I am using a Mac, so the path slashes in the tutorial are different from those for Windows: "/" versus "\\" (the script accounts for this difference, but the tutorial does not).
* Note that Python's `os.path.normpath()` will correct "/" to "\\" on Windows, but will not correct "\\" to "/" on Mac.

### Current Key Words
Each line of data received serially will be split at each comma to form a list (or a row), and the **first value** will determine the row type. **If no row type is provided, or the type is not supported, the DATA type will be assumed and added to the beginning of the row.** The table below shows the current row types and behaviors:

Row Type | Properties
--- | ---
DATA | Each value in the row will be checked for any of the key words in the "Key Words" table
LABEL | The row will be written to the output file as-is
MSG | The row will not be written to the output file, but it will still be printed in the command line where the script was run

If any of the below words is the **first value** of a row, it is a directive that will perform the action specified below:

Directive | Action
--- | ---
CLEARDATA | Signals the start of data collection (if at the beginning) or erases the sheet except for the header row (if at any time after the beginning)
RESETTIMER | Sets the reference time of the timer to the current time (see "TIMER" in the "Key Words" table)

If any of the below words are in a row of **data**, it will get replaced by a value, as shown below:

Key Word | String Replacement | Format
--- | --- | ---
TIME | Computer time | hh:mm:ss.000
TIMER | Number of seconds since the serial connection opened (or last timer reset) | 0.00
DATE | Computer date | mm-dd-yyyy

### Tutorial
If all of the libraries are installed, and the thermocouple code from E13.5 is on your Arduino (see [**Appendix B**](#appendix-b-arduino-code)), you are ready for the tutorial.

1. Before plugging in the Arduino to your computer, run `BB_DAQ.py`. **You can do this either from your IDE or a terminal window.** For this tutorial, I will use the terminal window (I used the `cd` command to get to the directory with the code).
```
brad@Brads-MBP ~ % cd "/Users/brad/Desktop/Courses/AME 341b/HW/Assignment Submissions/E13p5/Mac Workaround"
brad@Brads-MBP Mac Workaround % python3 BB_DAQ.py
Ports:
0: /dev/cu.Bluetooth-Incoming-Port
Enter the index of the port you want to use, or -1 to exit: 
```

2. At this point, take note of the ports available before plugging the Arduino in. After doing so, enter `-1` to exit, and then plug your Arduino into your computer. [**If R4**, press the Reset button on the Arduino.] Run the script again, and choose the new port (assuming you did not add or remove any other serial ports). Beyond this step, the process is the same whether you use an IDE or terminal window.
```
Choice: -1
Exiting...
brad@Brads-MBP Mac Workaround % python3 BB_DAQ.py
Ports:
0: /dev/cu.Bluetooth-Incoming-Port
1: /dev/cu.usbmodem11401
Enter the index of the port you want to use, or -1 to exit: 1
```

3. Afterwards, enter the buad rate (which should be 9600, but check the parameter in the `Serial.begin()` line in your Arduino code).
```
Enter the buad rate: 9600
```

4. Choose whether you want to save the data in an Excel file or a CSV file. (The assignment this tutorial was made for requires an Excel file, so `0` will be entered.)
```
Enter 0 to save as an Excel workbook, or enter 1 to save as a CSV file: 0
```

5. Enter the file name **without** the extension, but note that you will be asked to confirm the choice if there is already another file of the chosen type with the same name.
```
Enter workbook/file name or path (without the file-specific extension): Tutorial
```

6. The script will then measure the delay between consecutive packets of data, which should be close to the value in the Arduino code (the set delay should be 200 ms, but check the parameter in the `delay()` line in your Arduino code). The header for the data (the first line after the starting cue) will also appear.
    * Note that the delay should not be 0 ms. If it is, make sure there is a delay programmed in your Arduino code. If there must be no delay at all for your use, you can continue, but the live graph will not appear (the only ways to stop the code will be to press the Reset button on the Arduino or use Ctrl+C, but the latter may not end well despite my try-except statement).
    * Note that the header will be parsed as a string, meaning that there should be **no data** in the header.
```
Measuring delay between Arduino data packets...
Delay: 0.205 s.

Header:
LABEL,Computer Time,SNo,Time (Milli Sec.),Temp C
```

7. [**If R4**, press the Reset button on the Arduino.] You will have the option to choose when to see the graph of the data. For this tutorial, the live graph will be selected (`0`).
    * Note that if no graph is selected (`2`), you will not see some of the lines in the next steps that are needed for the graph.
```
Enter 0 to see the live graph, 1 to see the graph only in the Excel output, or 2 to not see the graph at all: 0
```

8. If a graph will be displayed, whether live or in the Excel sheet, you will be asked for the column indices for the x and y axes. For this assignment, enter `3` for the x-axis, then enter `4` for the y-axis.
```
Enter the column index (start at 0) for the x-axis in the data: 3
Enter the column index (start at 0) for the y-axis in the data: 4
```

9. If you chose to save the data as an Excel file, you will be asked to name the sheet.
```
Refer to the following website for sheet-naming rules:
https://support.microsoft.com/en-us/office/rename-a-worksheet-3f1f7148-ee83-404d-8ef0-9ff99fbad1f9

Enter valid sheet name: Sheet1
```

10. After either case, the following output will be displayed before the data from the Arduino is read and processed:
```
There are three ways to stop the program:
  Press any key while the graph window is selected.
  Press the Reset button on the Arduino.
  Press Ctrl+C (use as last resort).
```

11. At this point, the live graph should appear, and the data will be displayed in the output as it is being read and processed into an Excel file or CSV file. When you are finished, press any key while the graph window is selected to stop the code.
    * Note that if you use the Reset button on the Arduino to stop the program, you may get an error message besides the graceful exit depending on what line in the code is running.

12. After the serial stream stops, you will be asked if you want to run the program again with the same settings, but with the output in a new file or worksheet. This tutorial will exit here (`0`). There is no need to continue the tutorial down the branch where the program is run again since the questions are straightforward. [**If R4**, press the Reset button on the Arduino if you choose to run the program again (`1`).]
```
Exiting...

Would you like to run BB-DAQ again with the same settings, but with the output in a new file/worksheet?
Enter 0 to exit, or enter 1 to run again: 0
Done.
```

13. After completely exiting `BB_DAQ.py`, the output files should be available.

## Appendix A: BB_BoardTester Tutorial
If the pyserial library is installed, and the thermocouple code from E13.5 is on your Arduino (or any code with a value that increments with each loop), you are ready for the tutorial.
1. Before plugging in the Arduino to your computer, run the script. **You can do this either from your IDE or a terminal window.** For this tutorial, I will use the terminal window (I used the `cd` command to get to the directory with the code).
```
brad@Brads-MBP ~ % cd "/Users/brad/Desktop/Courses/AME 341b/HW/Assignment Submissions/E13p5/Mac Workaround"
brad@Brads-MBP Mac Workaround % python3 BB_BoardTester.py
Ports:
0: /dev/cu.Bluetooth-Incoming-Port
Enter the index of the port you want to use, or -1 to exit: 
```

2. At this point, take note of the ports available before plugging the Arduino in. After doing so, enter `-1` to exit, and then plug your Arduino into your computer. Run the script again, and choose the new port (assuming you did not add or remove any other serial ports). Beyond this step, the process is the same whether you use an IDE or terminal window.
```
Choice: -1
Exiting...
brad@Brads-MBP Mac Workaround % python3 BB_BoardTester.py
Ports:
0: /dev/cu.Bluetooth-Incoming-Port
1: /dev/cu.usbmodem11401
Enter the index of the port you want to use, or -1 to exit: 1
```

3. Afterwards, enter the buad rate (which should be 9600, but check the parameter in the `Serial.begin()` line in your Arduino code). The script will then open the serial port, take in 10 lines (an arbitrary hard-coded number), close the serial port, and do those three steps a second time.
    * If the incremented value resets to the original value in the beginning of the second set of lines, the board is similar to the Uno R3. If not, the board is similar to the Uno R4 Minima. The example below shows the output from an Uno R3 with the thermocouple code uploaded to it (with the thermocouple unplugged):
```
Enter the buad rate: 9600

Serial Connection Opening...

CLEARDATA
LABEL,Computer Time,SNo,Time (Milli Sec.),Temp C
DATA,TIME,1,0,0.00
DATA,TIME,2,233,0.00
DATA,TIME,3,467,0.00
DATA,TIME,4,702,0.00
DATA,TIME,5,935,0.00
DATA,TIME,6,1170,0.00
DATA,TIME,7,1403,0.00
DATA,TIME,8,1638,0.00

Serial Connection Closed...


Serial Connection Opening...

CLEARDATA
LABEL,Computer Time,SNo,Time (Milli Sec.),Temp C
DATA,TIME,1,0,0.00
DATA,TIME,2,233,0.00
DATA,TIME,3,467,0.00
DATA,TIME,4,702,0.00
DATA,TIME,5,935,0.00
DATA,TIME,6,1170,0.00
DATA,TIME,7,1403,0.00
DATA,TIME,8,1638,0.00

Serial Connection Closed...

Done.
```

## Appendix B: Arduino Code
```
// Thermocouple Arduino Code (AME-341b, Sp2023)
// Comments were modified to prevent spilling over to the next line

#include <Thermocouple.h>
#include <MAX6675_Thermocouple.h>

// Change the PIN locations as setup
#define SCK_PIN 10
#define CS_PIN 9
#define SO_PIN 8

unsigned long int milli_time; // Define milliseconds time variable

int i = 1; // Define serial number increment variable

// Assign Null as no value reading for thermocouple
Thermocouple* thermocouple = NULL;

// The setup function runs once when you press reset or power the board
void setup() {
  Serial.begin(9600); // Baud rate.
  Serial.println("CLEARDATA"); // Clear data each cycle in PLX-DAQ
  // Assign labels for excel columns in PLX-DAQ
  Serial.println("LABEL,Computer Time,SNo,Time (Milli Sec.),Temp C");
  // Assign data to thermocouple variable
  thermocouple = new MAX6675_Thermocouple(SCK_PIN, CS_PIN, SO_PIN);
}

// The loop function runs over and over again forever
void loop() {
  // Assign milliseconds time variable
  milli_time = millis();
  // Read temperature with built in headers to convert voltage to ºC
  const double celsius = thermocouple->readCelsius();
  // Write the data to serial
  Serial.print("DATA,TIME,");   // Have PLX-DAQ record computer time
  Serial.print(i);              // Serial number for the data
  Serial.print(",");            // Insert comma
  Serial.print(milli_time);     // Display milliseconds time
  Serial.print(",");            // Insert comma
  Serial.print(celsius);        // Display temperature
  Serial.println();             // Next line
  
  delay(200); // Delay the output of information (200 ms was suggested)
  i++; // Increment serial number
}
```
