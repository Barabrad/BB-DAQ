'''
Brad Barakat
Made for AME-341b

This script is meant to test the board used in to collect data via BB-DAQ.
Boards like the Arduino Uno R3 reset when the serial connection is closed in BB-DAQ. Boards like the Arduino Uno R4 Minima don't.
The board will need a program that would increment a value and send it serially.
'''

# If serial is not installed, type "python3 -m pip install pyserial" into a Terminal window
# Note that if you have another serial library installed, it may interfere with this one
import serial
from serial.tools import list_ports


# This function gets a valid integer input from the user
def getValidIntInput(prompt, lBnd=None, hBnd=None):
    valid = False;
    while (not valid):
        x = input(prompt).strip();
        try:
            x = int(x);
            if (lBnd == None): lBnd = x;
            if (hBnd == None): hBnd = x;
            valid = (lBnd <= x) and (x <= hBnd);
            if (not valid): print(f"Error: Integer out of range [{lBnd},{hBnd}]");
        except:
            print("Error: Numeric input not an integer");
    return x;


# This function opens the serial connection, gets a batch of data, then closes the connection
def getSerialBatch(ser, numLines):
    print("\nSerial Connection Opening...\n");
    ser.open();
    for _ in range(numLines):
        dataIn = (ser.readline()).decode().rstrip('\r\n');
        print(dataIn);
    ser.close();
    print("\nSerial Connection Closed...\n");


# main()
def main():
    # This first part will find the available serial ports. The Arduino should be a USB port.
    # To be sure of the Arduino's port, run this part before and after plugging in the Arduino, and compare the
    # output. To minimize confusion, make sure no other devices are also being plugged in between the two runs.
    portList = list(list_ports.comports());
    p_ind = 0;
    print("Ports:");
    for p in portList:
        print(str(p_ind) + ": " + p.device);
        p_ind += 1;
    portPrompt = "Enter the index of the port you want to use, or -1 to exit: ";
    portChoice = getValidIntInput(portPrompt, -1, len(portList)-1);

    if (portChoice == -1):
        print("Exiting...");
    else:
        # This second part will show the user if the board resets upon closing the serial connection.
        numLines = 10; # Arbitrary, but it must be enough to show whether or not the board reset

        # Get port info from user
        port = portList[portChoice].device;
        buad = getValidIntInput("Enter the buad rate: ", 1);
        # See the rest of serial.Serial()'s parameters here:
        # https://pyserial.readthedocs.io/en/latest/pyserial_api.html#serial.Serial.__init__
        ser = serial.Serial(port, buad);
        ser.close();

        # Print two batches of serial data
        for _ in range(2): getSerialBatch(ser, numLines);

        print("Done.");


# Run main()
main();
