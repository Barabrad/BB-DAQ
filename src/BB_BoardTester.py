'''
Brad Barakat
Made for AME-341b

This script is meant to test the board used in to collect data via BB-DAQ.
Boards like the Arduino Uno R3 reset when the serial connection is closed in BB-DAQ. Boards like
the Arduino Uno R4 Minima don't.
The board will need a program that would increment a value and send it serially.
'''

# If serial is not installed, type "python3 -m pip install pyserial" into a Terminal window
# Note that if you have another serial library installed, it may interfere with this one
import serial
from serial.tools import list_ports


def is_num_str(x_str:str, num_type:type=float) -> bool:
    """
    This function checks if a string represents a specified numeric type (default: float)
    @param x_str: the string supposedly representing a number
    @param num_type: the numerical type to check for
    @return: a boolean that is true if the string represents the specified numeric type
    """
    valid = True
    try:
        x_flt = float(x_str)
        if num_type == int:
            valid = int(x_flt) == x_flt
    except (ValueError, TypeError):
        valid = False
    return valid


def get_int_input(prompt:str, l_bnd:int=None, u_bnd:int=None) -> int:
    """
    This function gets a valid integer input from the user
    @param prompt: the string asking the user for an integer
    @param l_bnd: lower bound (optional)
    @param u_bnd: upper bound (optional)
    @return: a valid integer
    """
    valid = False
    while (not valid):
        x = input(prompt).strip()
        if not is_num_str(x, int):
            print("Error: Numeric input not an integer")
            continue
        x = int(x)
        if l_bnd is None:
            l_bnd = x
        if u_bnd is None:
            u_bnd = x
        valid = (l_bnd <= x <= u_bnd)
        if not valid:
            print(f"Error: Integer out of range [{l_bnd},{u_bnd}]")
    return x


def get_serial_batch(ser:serial.Serial, num_lines:int) -> None:
    """
    This function opens the serial connection, prints a batch of data, then closes the connection
    @param ser: the Serial object that is connected to the device
    @param num_lines: the number of lines to read then print
    @return: a valid integer
    """
    print("\nSerial Connection Opening...\n")
    ser.open()
    for _ in range(num_lines):
        data_in = ser.readline().decode().strip()
        print(data_in)
    ser.close()
    print("\nSerial Connection Closed...\n")


def main() -> None:
    """
    This is the main function
    @return: None
    """
    # This first part will find the available serial ports. The Arduino should be a USB port.
    # To be sure of the Arduino's port, run this part before and after plugging in the Arduino, and
    # compare the output. To minimize confusion, make sure no other devices are also being plugged
    # in between the two runs.
    port_list = list_ports.comports() # Outputs a list
    p_ind = 0
    print("Ports:")
    for p in port_list:
        print(f"{p_ind}: {p.device}")
        p_ind += 1
    port_prompt = "Enter the index of the port you want to use, or -1 to exit: "
    port_choice = get_int_input(port_prompt, -1, p_ind-1) # At this point, p_ind = len(port_list)

    if (port_choice == -1):
        print("Exiting...")
        return

    # This second part will show the user if the board resets upon closing the serial connection.
    num_lines = 10 # Arbitrary, but it must be enough to show whether the board reset

    # Get port info from user
    port = port_list[port_choice].device
    buad = get_int_input("Enter the buad rate: ", 1)
    # See the rest of serial.Serial()'s parameters here:
    # https://pyserial.readthedocs.io/en/latest/pyserial_api.html#serial.Serial.__init__
    ser = serial.Serial(port, buad)
    # Close the port in case it is already open (this can happen when a serial connection isn't
    # closed gracefully)
    ser.close()
    # Print two batches of serial data
    for _ in range(2):
        get_serial_batch(ser, num_lines)

    print("Done.")


# Run main()
if (__name__ == "__main__"):
    main()
