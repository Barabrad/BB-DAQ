'''
Brad Barakat
Made for AME-341b

This script is meant to be a Mac workaround for PLX-DAQ since PLX-DAQ uses COM ports (Macs do not
have these).
However, this has worked on a Windows computer, so this is not exclusive to Mac.
There are two parts to this script: the port finder, and the serial reader and publisher (to Excel
or CSV file).

NOTE: This script does not replicate all of the features of PLX-DAQ! This script was originally
  made to read data serially from an Arduino (see Appendix B in the README for the specific Arduino
  file), plot the data, and write to Excel. Replications for commands like "RESETTIMER" and
  "CLEARDATA" were added over a year later as an afterthought.

DEV NOTE: the following structures are mutable, so changes made in functions will be carried
  outside the functions, so there is no need to return them unless you're deleting and recreating
  them, or initially creating them.
  - matplotlib Axes
  - xlsxwriter Workbook
  - xlsxwriter Worksheet
  - Python list
'''

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
# If serial is not installed, type "python3 -m pip install pyserial" into a Terminal window
# Note that if you have another serial library installed, it may interfere with this one
import serial
from serial.tools import list_ports
# If xlsxwriter is not installed, type "pip3 install xlsxwriter" into a Terminal window
import xlsxwriter
import xlsxwriter.worksheet
# If matplotlib is not installed, type "pip3 install matplotlib" into a Terminal window
from matplotlib import axes as pltaxes, pyplot as plt


# Make aliases for long class names for type-hinting
PySerial = serial.Serial
XlsxWkbk = xlsxwriter.Workbook
XlsxSheet = xlsxwriter.worksheet.Worksheet
XlsxFormat = xlsxwriter.workbook.Format
PltAxes = pltaxes.Axes


# Constants
INTERVAL_PLOT: float = 0.5 # Minimum number of seconds before plot is updated (semi-arbitrary)
DATA_START_AFTER: str = "CLEARDATA"
DATA_DELIM: str = ","
# Spreadsheet name bad characters
BAD_STARTS: set[str] = {"'"}
BAD_ENDS: set[str] = {"'"}
BAD_CHARS: set[str] = {"/", "\\", "?", "*", ":", "[", "]"}
# Supported directives
RESET_TIMER: str = "RESETTIMER"
CLEAR_DATA: str = "CLEARDATA"
# Supported row types
DATA_ROW: str = "DATA"
LABEL_ROW: str = "LABEL"
MSG_ROW: str = "MSG"
# All supported key words
KEY_WORDS: set[str] = {RESET_TIMER, CLEAR_DATA, DATA_ROW, LABEL_ROW, MSG_ROW}
# Special data words
TIME_WORD: str = "TIME"
TIMER_WORD: str = "TIMER"
DATE_WORD: str = "DATE"


# Enums
class GraphChoice(Enum):
    """
    Enum representing the choices for viewing the data graph
    """
    LIVE = 0
    EXCEL_ONLY = 1
    NONE = 2


# Classes
class GraphData():
    """
    Class containing graph-related data
    """

    def __init__(self, user_gc:GraphChoice, time_col_ind:int, data_col_ind:int, \
                 graph_pause:float, buf_size:int) -> None:
        """
        This method is the constructor
        @param self: Not needed in calls
        @param user_gc: the GraphChoice enum representing the user's choice
        @param time_col_ind: the index (0-based) of the time column ("x"), used for graphing
        @param data_col_ind: the index (0-based) of the data column ("y"), used for graphing
        @param graph_pause: the amount of time to pause the live graph for
        @param buf_size: the length of the buffers (one of them, not combined)
        @return: None
        """
        # Define parameters based on user choice
        self.user_gc = user_gc
        self.is_live = (user_gc == GraphChoice.LIVE)
        self.is_graphed = (user_gc != GraphChoice.NONE)
        self.time_col_ind = time_col_ind
        self.data_col_ind = data_col_ind
        if self.is_live:
            self.graph_pause = graph_pause
            self.buf_size = buf_size
            self.buf_ind = 0
            self.buf_x_plot:list[float] = [None]*buf_size
            self.buf_y_plot:list[float] = [None]*buf_size
            self.num_plot_bufs = 0 # Count number of buffers on current plot
            self.fig, self.ax = plt.subplots(1,1)
            plt.ion()

    def disable_graph(self) -> None:
        """
        This method will set the graph choice to NONE
        (Use case: user chooses CSV but with graph choice as EXCEL_ONLY)
        """
        self.user_gc = GraphChoice.NONE
        self.is_graphed = False
        self.time_col_ind = self.data_col_ind = -1

    def set_ax_labels(self, x:str, y:str) -> None:
        """
        This method sets the label of the axes IFF there is a live graph
        @param self: Not needed in calls
        @param x: x-axis label
        @param y: y-axis label
        @return: None
        """
        if not self.is_live:
            return
        self.ax.set_xlabel(x)
        self.ax.set_ylabel(y)

    def add_to_buffers(self, x:float, y:float) -> None:
        """
        This method adds data to the x and y buffers IFF there is a live graph, then plots the
        data when the buffers are full
        @param self: Not needed in calls
        @param x: x value
        @param y: y value
        @return: None
        """
        if not self.is_live:
            return
        self.buf_x_plot[self.buf_ind] = x
        self.buf_y_plot[self.buf_ind] = y
        self.buf_ind += 1
        if self.buf_ind == self.buf_size:
            self.plot_buffer_data()

    def plot_buffer_data(self) -> None:
        """
        This method plots the data in the buffers IFF there is a live graph, then resets the
        buffer index
        @param self: Not needed in calls
        @return: None
        """
        if not self.is_live:
            return
        self.ax.plot(self.buf_x_plot, self.buf_y_plot, "-b")
        if plt.waitforbuttonpress(self.graph_pause):
            raise KeyboardInterrupt # This will wait for keypress
        self.num_plot_bufs += 1
        # Make sure the lines connect by saving the most recent values at index 0
        self.buf_x_plot[0] = self.buf_x_plot[-1]
        self.buf_y_plot[0] = self.buf_y_plot[-1]
        self.buf_ind = 1

    def overwrite_buffers(self) -> None:
        """
        This method overwrites the buffers with None IFF there is a live graph
        @param self: Not needed in calls
        @return: None
        """
        if not self.is_live:
            return
        for i in range(self.buf_size):
            self.buf_x_plot[i] = None
            self.buf_y_plot[i] = None
        self.buf_ind = 0

    def reset_axes(self) -> None:
        """
        This method resets the axes (while preserving labels) IFF there is a live graph
        @param self: Not needed in calls
        @return: None
        """
        if not self.is_live:
            return
        x_label = self.ax.get_xlabel()
        y_label = self.ax.get_ylabel()
        self.ax.cla()
        self.set_ax_labels(x_label, y_label)
        self.num_plot_bufs = 0

    def close_fig(self) -> None:
        """
        This method closes the figure IFF there is a live graph
        @param self: Not needed in calls
        @return: None
        """
        if not self.is_live:
            return
        plt.ioff()
        plt.delaxes(self.ax)
        plt.pause(0.01)
        plt.close(self.fig)


class FileData():
    """
    Class containing file-related data
    """

    def __init__(self, save_as_xlsx:bool, file_name:str, header_txt:str) -> None:
        """
        This method is the constructor
        @param self: Not needed in calls
        @param save_as_xlsx: a boolean for the file extension (true:".xlsx", false:".csv")
        @param file_name: the name of the file that the data will be written to
        @param header_txt: the joined delimeter-separated values that make up the header
        @return: None
        """
        # Define parameters based on user choice
        self.file_name = file_name
        self.is_xlsx = save_as_xlsx
        self.header_txt = header_txt
        self.row_num = 0
        if self.is_xlsx:
            self.workbook:XlsxWkbk = None
            self.curr_sheet:XlsxSheet = None
            self.create_workbook(file_name) # Populates self.workbook
            # Formatters
            self.format_time:XlsxFormat = None
            self.format_timer:XlsxFormat = None
            self.format_date:XlsxFormat = None
            self.add_workbook_formats()
        else:
            # Clear the text file
            self.write_to_file("", append=False)

    def add_workbook_formats(self):
        """
        This method adds the Format objects for specific cells
        """
        self.format_time = self.workbook.add_format({'num_format': 'hh:mm:ss.000'})
        self.format_timer = self.workbook.add_format({'num_format': '0.00'})
        self.format_date = self.workbook.add_format({'num_format': 'mm-dd-yyyy'})

    def write_to_file(self, text:str|float|list[str], col:int=0, cell_fmt:XlsxFormat=None, \
                      inc_row_num:bool=False, append:bool=True) -> None:
        """
        This method writes text (or a row of text) to the file/sheet
        @param self: Not needed in calls
        @param text: the single string or list of text
        @param col: the column in the sheet to start writing at (only for Excel)
        @param cell_fmt: the format for the single cell (only for Excel)
        @param inc_row_num: a boolean for incrementing the current row number (only for Excel)
        @param append: a boolean for the CSV file writing mode (only for CSV)
        @return: None
        """
        is_list = isinstance(text, list)
        if self.is_xlsx:
            if is_list:
                self.curr_sheet.write_row(self.row_num, col, text)
            else:
                self.curr_sheet.write(self.row_num, col, text, cell_fmt)
        else:
            if is_list:
                text = ",".join(text) + "\n" # DATA_DELIM may not always be a comma
            opt = "at" if append else "wt"
            with open(self.file_name, mode=opt, encoding='utf-8') as f_out:
                f_out.write(text)
        self.row_num += 1 if inc_row_num else 0

    def add_formatted_sheet(self, valid_sheet_name:str=None) -> str:
        """
        This method adds a formatted sheet to the workbook IFF the file is a workbook
        @param self: Not needed in calls
        @param valid_sheet_name: a valid sheet name (this argument is used internally)
        @return: the sheet name
        """
        if not self.is_xlsx:
            return None
        if valid_sheet_name is None:
            valid_sheet_name = self.get_valid_sheet_name()
        self.curr_sheet = self.workbook.add_worksheet(valid_sheet_name)
        # Make sure the row number is 0 (especially if switching sheets)
        self.row_num = 0
        # Make columns 1 and 3 (0-indexed) wider
        # (This formatting is for BB-DAQ's original purpose, so feel free to change it)
        self.curr_sheet.set_column(1, 1, 15)
        self.curr_sheet.set_column(3, 3, 15)
        return valid_sheet_name

    def get_valid_sheet_name(self) -> str:
        """
        This method gets a valid sheet name from the user IFF the file is a workbook
        @param self: Not needed in calls
        @return: the sheet name
        """
        if not self.is_xlsx:
            return None
        print("\nRefer to the following website for sheet-naming rules:")
        print("https://support.microsoft.com/en-us/office/rename-a-worksheet-3f1f7148-" \
            "ee83-404d-8ef0-9ff99fbad1f9\n")
        sheet_prompt = "Enter valid sheet name: "
        is_good_name = False
        # Make sure no duplicates
        all_sheet_names:set[str] = self.workbook.sheetnames.keys()
        all_sheet_names = set(name.upper() for name in all_sheet_names)
        while not is_good_name:
            sheet_name = input(sheet_prompt).strip()
            len_name = len(sheet_name)
            # Run checks
            if (len_name > 32) or (len_name == 0):
                continue
            if sheet_name.lower() == "history":
                continue
            if sheet_name[0] in BAD_STARTS:
                continue
            if sheet_name[-1] in BAD_ENDS:
                continue
            if sheet_name.upper() in all_sheet_names:
                print("There is already a sheet by that name (case-insensitive)")
                continue
            char_set = set(sheet_name)
            is_good_name = char_set.isdisjoint(BAD_CHARS)
        return sheet_name

    def add_chart_to_sheet(self, time_col_ind:int, data_col_ind:int) -> None:
        """
        This method adds a chart to the current sheet IFF the file is a workbook
        @param self: Not needed in calls
        @param time_col_ind: the index (0-based) of the time column ("x"), used for graphing
        @param data_col_ind: the index (0-based) of the data column ("y"), used for graphing
        @return: None
        """
        if not self.is_xlsx:
            return
        capital_a_int = ord("A")
        time_col = chr(capital_a_int + time_col_ind)
        data_col = chr(capital_a_int + data_col_ind)
        num_header_cols = len(self.header_txt.split(DATA_DELIM))
        chart_col = chr(capital_a_int + num_header_cols + 1)
        chart = self.workbook.add_chart({'type': 'line'})
        final_row_str = str(self.row_num)
        sheet_name = self.curr_sheet.name # To save steps since used a lot
        chart.add_series({
            'categories': f'={sheet_name}!${time_col}$2:${time_col}${final_row_str}',
            'values':     f'={sheet_name}!${data_col}$2:${data_col}${final_row_str}',
        })
        chart.set_x_axis({'name': f'={sheet_name}!${time_col}$1'})
        chart.set_y_axis({'name': f'={sheet_name}!${data_col}$1'})
        chart.set_legend({'none': True})
        # Insert the chart into the worksheet
        self.curr_sheet.insert_chart(chart_col + '2', chart)

    def reset_current_page(self) -> None:
        """
        This method resets the CSV file or Excel sheet (not the whole workbook), and prints the
        header again to replicate PLX-DAQ's "CLEARDATA"
        @param self: Not needed in calls
        @return: None
        """
        if self.is_xlsx:
            # Re-create sheet by deleting and adding it
            sheet_name = self.curr_sheet.name
            self.workbook.worksheets().remove(self.curr_sheet)
            self.curr_sheet = None
            self.add_formatted_sheet(sheet_name)
            # Write header
            self.write_to_file(self.header_txt.split(DATA_DELIM))
        else:
            self.write_to_file(f"{self.header_txt}\n", append=False)
        self.row_num = 1

    def create_workbook(self, file_name:str) -> None:
        """
        This method creates a workbook IFF the file is a workbook
        @param self: Not needed in calls
        @param file_name: The file path of the workbook to be made
        @return: None
        """
        if not self.is_xlsx:
            return
        self.workbook = xlsxwriter.Workbook(file_name, {'constant_memory': True})

    def close_workbook(self) -> None:
        """
        This method closes the workbook IFF the file is a workbook
        @param self: Not needed in calls
        @return: None
        """
        if not self.is_xlsx:
            return
        self.workbook.close()

    def switch_to_new_file(self, new_file_name:str) -> None:
        """
        This method switches the file/workbook while preserving the other attributes
        @param self: Not needed in calls
        @param new_file_name: The file path of the file to switch to
        @return: None
        """
        self.file_name = new_file_name
        self.row_num = 0
        if self.is_xlsx:
            self.close_workbook()
            self.create_workbook(new_file_name)
            self.add_workbook_formats()


# Functions
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
        print(x)
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


def resolve_dup_file(filepath:str, ext:str) -> str:
    """
    This function prompts the user to either overwrite the specified file or enter another file
    name
    @param filepath: the path to the desired file, including the extension
    @param ext: the file extension
    @return: the final file name, including the extension
    """
    file_overwrite = False
    retry_prompt = f"Enter another workbook/file name or path (without the '{ext}' at the end): "
    while os.path.exists(filepath) and (not file_overwrite):
        print("This file already exists:", filepath)
        file_overwrite = input("Do you wish to overwrite it? ('y'/'n'): ").upper() == "Y"
        if not file_overwrite:
            raw_file = input(retry_prompt)
            filepath = os.path.normpath(raw_file + ext)
    # Make sure the directory exists
    file_dir = os.path.split(filepath)[0]
    if file_dir != "":
        # Using the following os methods on "" would throw errors
        if (not os.path.exists(file_dir)):
            os.makedirs(file_dir)
    return filepath


def get_file_name(save_as_xlsx:bool) -> str:
    """
    This function gets a valid file name from the user
    @param save_as_xlsx: a boolean that determines the file extension (true:".xlsx", false:".csv")
    @return: the final file name, including the extension
    """
    raw_file = input("Enter workbook/file name or path (without the file-specific extension): ")
    file_name = os.path.normpath(raw_file)
    ext = ".xlsx" if save_as_xlsx else ".csv"
    file_name += ext
    file_name = resolve_dup_file(file_name, ext)
    return file_name


def get_header_and_delay(ser:PySerial) -> tuple[str, float, float]:
    """
    This function finds the header of the data and the Arduino delay time between data lines
    It also finds the time the program should pause for after updating the graph
    @param ser: the Serial object that is connected to the device
    @return: a tuple containing the header, the Arduino delay, and the graph pause time
    """
    t0 = t1 = 0
    data_started = read_header = t0_found = False
    run_header_loop = True
    header_txt = ""
    print("\nMeasuring delay between Arduino data packets...")
    ser.open()
    while run_header_loop:
        data_in = ser.readline().decode().strip()
        if not data_started:
            if data_in.upper() == DATA_START_AFTER:
                data_started = True
        elif not read_header:
            header_txt = data_in # Save the header
            read_header = True
        elif not t0_found:
            t0 = time.time()
            t0_found = True
        else:
            t1 = time.time()
            run_header_loop = False
    ser.close()
    delay_ard = round(t1 - t0, 3)
    # Determine pause time for graph
    graph_pause = 0.5*delay_ard # Account for data processing time (the 0.5 is arbitrary)
    if (delay_ard == 0):
        # This is unlikely to happen, but I must account for it
        print("No notable Arduino delay between messages. There should be some sort of delay" \
              " on the order of at least milliseconds.")
        print("If you want to see the live graph, there must be some pause for it to update.")
        print("If a delay is introduced, the graph and data-writing will lag behind, but there" \
              " will be no gaps in the data stream.")
        add_delay = input("Add a delay of 1 ms? ('y'/'n'): ").upper() == "Y"
        if add_delay:
            delay_ard = 0.001
            graph_pause = 0.001
    else:
        print("Delay:", delay_ard, "s.")
    return (header_txt, delay_ard, graph_pause)


def get_row_type_and_num_cols(row_arr:list[str], default_row:str) -> tuple[str, int, bool]:
    """
    This function determines what a row is (i.e., LABEL or DATA)
    If the first value is not a row type or directive, then the row is assumed to be the given
    default
    If the list is empty, the returned type will be None
    @param row_arr: a list of each delimiter-separated value in the row
    @param default_row: the default row type
    @return: a tuple containing the row type, the number of columns, and a flag for a missing label
    """
    num_cols = len(row_arr)
    missing_label = False
    if num_cols == 0:
        row_type = None
    else:
        col1 = row_arr[0].strip().upper()
        if col1 == "":
            row_type = None
        elif col1 in KEY_WORDS:
            row_type = col1
        else:
            missing_label = True
            row_type = default_row # Set to default
    return (row_type, num_cols, missing_label)


def process_data_row(row:list[str], num_cols:int, timer_t0:float, file_struct:FileData, \
                     graph_struct:GraphData):
    """
    This function processes a data row, which entails checking for key words, writing to file, and
    graphing
    @param row: a list of each delimiter-separated value in the row
    @param num_cols: the number of delimeter-separated values in the row
    @param timer_t0: the reference second count for the timer
    @param file_struct: the FileData object containing the file-related information
    @param graph_struct: the GraphData object containing the graph-related information
    @return: None
    """
    # Reset the time and data values
    curr_time = curr_data = None
    # Pull often-used class variables
    time_col_ind = graph_struct.time_col_ind
    data_col_ind = graph_struct.data_col_ind
    is_graphed = graph_struct.is_graphed
    save_as_xlsx = file_struct.is_xlsx
    format_time = file_struct.format_time if save_as_xlsx else None
    format_timer = file_struct.format_timer if save_as_xlsx else None
    format_date = file_struct.format_date if save_as_xlsx else None
    # Begin data processing
    for col in range(num_cols):
        cell_data = row[col]
        # These booleans are used in later if-else statements since cell_data will be overwritten
        is_time = is_timer = is_date = False
        # Swap out key words with the values
        cell_data_upper = cell_data.upper()
        if cell_data_upper == TIME_WORD:
            is_time = True
            cell_data = datetime.now().time()
            cell_format = format_time
        elif cell_data_upper == TIMER_WORD:
            is_timer = True
            cell_data = round(time.time() - timer_t0, 3)
            cell_format = format_timer
        elif cell_data_upper == DATE_WORD:
            is_date = True
            cell_data = datetime.now().date()
            cell_format = format_date
        else:
            cell_format = None
        # Check if the data is a graphed value
        is_x_axis = (col == time_col_ind)
        is_y_axis = (col == data_col_ind)
        is_datetime = is_time or is_date
        is_numeric = (not is_datetime) and is_num_str(cell_data)
        if is_graphed and (is_x_axis or is_y_axis):
            if is_datetime:
                plot_data = str(cell_data) # Datetime objects can't be plotted
            elif is_numeric:
                # Don't cast cell_data to a float since it will be converted back to string for CSV
                plot_data = float(cell_data) # Datetime objects can't be plotted
            else:
                plot_data = cell_data
            if is_y_axis:
                curr_data = plot_data
            else: # is_x_axis
                curr_time = plot_data
        # Write to file accordingly
        if save_as_xlsx:
            if is_numeric:
                cell_data = float(cell_data)
            file_struct.write_to_file(cell_data, col, cell_format)
        else:
            # The row array is unused after the column iteration, so it can be reused for holding
            # CSV values
            if is_datetime or is_timer:
                row[col] = str(cell_data) # All values in CSV are strings
    # Deal with plot (class has live-checking logic)
    graph_struct.add_to_buffers(curr_time, curr_data)
    # Write row array to CSV file
    if not save_as_xlsx:
        file_struct.write_to_file(row)
    # Increment row count
    file_struct.row_num += 1


def process_label_row(row:list[str], file_struct:FileData) -> None:
    """
    This function processes a label row
    @param file_struct: the FileData object containing the file-related information
    @return: None
    """
    file_struct.write_to_file(row, inc_row_num=True)


def process_msg_row() -> None:
    """
    This function processes a message row
    (Nothing is to be done in this case yet)
    @return: None
    """


def process_reset_timer() -> float:
    """
    This function processes the reset timer directive
    @return: new reference second count for the timer
    """
    return time.time() # New timer_t0


# This function processes the clear data directive
def process_clear_data(file_struct:FileData, graph_struct:GraphData) -> None:
    """
    This function processes the clear data directive
    @param file_struct: the FileData object containing the file-related information
    @param graph_struct: the GraphData object containing the graph-related information
    @return: None
    """
    file_struct.reset_current_page()
    # GraphData has live-checking logic
    graph_struct.reset_axes()
    graph_struct.overwrite_buffers()


def get_and_write_data(ser:PySerial, file_struct:FileData, graph_struct:GraphData) -> None:
    """
    This function does the reading of serial data and writing of the output file
    (The optional parameters are populated internally if the user wants to run it again)
    @param ser: the Serial object that is connected to the device
    @param file_struct: the FileData object containing the file-related information
    @param graph_struct: the GraphData object containing the graph-related information
    @return: None
    """
    # Find how many columns the header has
    header = file_struct.header_txt.split(DATA_DELIM)

    # FileData has the logic to check if a spreadsheet is used
    file_struct.add_formatted_sheet()

    # GraphData has the logic to check if the graph is live
    x_label = header[graph_struct.time_col_ind]
    y_label = header[graph_struct.data_col_ind]
    graph_struct.set_ax_labels(x_label, y_label)

    # Make local aliases for often-used class variables, especially in time-sensitive loops
    save_as_xlsx = file_struct.is_xlsx

    # Make sure that the graphing choice makes sense
    if (not save_as_xlsx) and (graph_struct.user_gc == GraphChoice.EXCEL_ONLY):
        graph_struct.disable_graph()

    data_started = False
    timer_t0 = time.time()
    ser.open()
    try:
        print("\nThere are three ways to stop the program:")
        print("  Press any key while the graph window is selected.")
        print("  Press the Reset button on the Arduino.")
        print("  Press Ctrl+C (use as last resort).\n")
        # Loop until we hit DATA_START_AFTER
        while not data_started:
            # Read in a line of data and parse it
            data_in = ser.readline().decode().strip()
            data_started = (data_in.upper() == DATA_START_AFTER)
        # Now we're onto the header
        _ = ser.readline() # Discard the header since we already have it
        file_struct.write_to_file(file_struct.header_txt.split(DATA_DELIM), inc_row_num=True)
        # Now we're onto the data
        while True:
            # The rows are iterated by the while loop, but columns will be iterated by the for loop
            # Read in a line of data and parse it
            data_in = ser.readline().decode().strip()
            print(data_in)
            row = data_in.split(DATA_DELIM)
            (row_type, num_cols, missing_label) = get_row_type_and_num_cols(row, DATA_ROW)
            row_is_data = (row_type == DATA_ROW)
            row_is_msg = (row_type == MSG_ROW)
            # Check if the data stopped coming in
            if row_type is None:
                print("\nNo data received. Serial must've timed out.")
                raise KeyboardInterrupt
            # If the label is missing, add it
            if missing_label and (not row_is_msg):
                row = [row_type] + row
                num_cols += 1
            # Perform actions depending on the row type
            if row_is_data:
                process_data_row(row, num_cols, timer_t0, file_struct, graph_struct)
            elif row_type == RESET_TIMER:
                timer_t0 = process_reset_timer()
            elif row_type == CLEAR_DATA:
                process_clear_data(file_struct, graph_struct)
            elif row_is_msg:
                process_msg_row()
            elif row_type == LABEL_ROW:
                process_label_row(row, file_struct)
            else:
                # This line should not be reached, so it's good for troubleshooting
                print(f"Unexpected row type: {row_type}")
    except KeyboardInterrupt:
        print("\nExiting...")
    except:
        print(f"\nSomething went wrong:\n{traceback.format_exc()}\n")
    finally:
        ser.close()
        # GraphData has the logic to check if the graph is live
        graph_struct.close_fig()

    # Give the user the option to run BB-DAQ again with the same settings
    # (but in a new file/worksheet)
    print("\nWould you like to run BB-DAQ again with the same settings,"\
          " but with the output in a new file/worksheet?")
    rerun_prompt = "Enter 0 to exit, or enter 1 to run again: "
    run_again = (get_int_input(rerun_prompt, 0, 1) == 1)
    new_file = True # Default for CSV
    if run_again and save_as_xlsx:
        rerun_prompt_xlsx = "Enter 0 to make a new worksheet in the same workbook,"\
            " or enter 1 to make a new workbook: "
        new_file = (get_int_input(rerun_prompt_xlsx, 0, 1) == 1)

    # Add the chart before moving on from the worksheet
    if graph_struct.is_graphed:
        file_struct.add_chart_to_sheet(graph_struct.time_col_ind, graph_struct.data_col_ind)

    # Run again (generalized for both cases)
    if run_again:
        if new_file:
            file_name = get_file_name(save_as_xlsx)
            file_struct.switch_to_new_file(file_name)
        # file_struct.sheet will be overwritten in this function
        get_and_write_data(ser, file_struct, graph_struct)
    else:
        file_struct.close_workbook()


def get_port_info() -> tuple[str, int]:
    """
    This function gets the port info from the user
    @return: a tuple with the port name and the buad rate
    """
    # First get the choice from all available ports
    port_list = list_ports.comports() # Outputs a list
    p_ind = 0
    print("Ports:")
    for p in port_list:
        print(f"{p_ind}: {p.device}")
        p_ind += 1
    port_prompt = "Enter the index of the port you want to use, or -1 to exit: "
    port_choice = get_int_input(port_prompt, -1, p_ind-1) # At this point, p_ind = len(portList)
    # Now get the other info if the user selects a port
    if port_choice == -1:
        port = buad = None
    else:
        port = port_list[port_choice].device
        buad = get_int_input("Enter the buad rate: ", 1)
    return (port, buad)


def get_graph_info(save_as_xlsx:bool, header_txt:str) -> tuple[GraphChoice, int, int]:
    """
    This function gets the graph preferences and info from the user
    @param save_as_xlsx: a boolean that determines the file extension (true:".xlsx", false:".csv")
    @param header_txt: the joined delimeter-separated values that make up the header
    @return: a tuple with the user's GraphChoice enum, time column index, and data column index
    """
    graph_prompt = "Enter 0 to see the live graph, 1 to see the graph only in the Excel output, " \
        "or 2 to not see the graph at all: "
    user_gc = GraphChoice(get_int_input(graph_prompt, 0, 2))
    # Ask plot questions if the graph will appear at any point
    if (user_gc == GraphChoice.LIVE) or ((user_gc == GraphChoice.EXCEL_ONLY) and save_as_xlsx):
        time_prompt = "Enter the column index (start at 0) for the x-axis in the data: "
        data_prompt = "Enter the column index (start at 0) for the y-axis in the data: "
        col_upper_bnd = len(header_txt.split(DATA_DELIM)) - 1
        time_col_ind = get_int_input(time_prompt, 0, col_upper_bnd)
        data_col_ind = get_int_input(data_prompt, 0, col_upper_bnd)
    else:
        time_col_ind = data_col_ind = -1
    return (user_gc, time_col_ind, data_col_ind)


def main() -> None:
    """
    This is the main function
    @return: None
    """
    # This first part will find the available serial ports. The Arduino should be a USB port.
    # To be sure of the Arduino's port, run this part before and after plugging in the Arduino,
    # and compare the output. To minimize confusion, make sure no other devices are also being
    # plugged in between the two runs.
    (port, buad) = get_port_info()

    if port is None:
        print("Exiting...")
        return

    # This second part will actually read the serial data from the Arduino and write it to a file.
    # A live graph of the numerical data will also be generated (if desired by the user).

    # Get file name
    print("")
    choice_prompt = "Enter 0 to save as an Excel workbook, or enter 1 to save as a CSV file: "
    save_as_xlsx = (get_int_input(choice_prompt, 0, 1) == 0)
    file_name = get_file_name(save_as_xlsx)

    # See the rest of serial.Serial()'s parameters here:
    # https://pyserial.readthedocs.io/en/latest/pyserial_api.html#serial.Serial.__init__
    ser = serial.Serial(port, buad)
    # Close the port in case it is already open
    # (this can happen when a serial connection isn't closed gracefully)
    ser.close()

    # Find the header and delay time between data (and for graph)
    (header_txt, delay_ard, graph_pause) = get_header_and_delay(ser)
    print(f"\nHeader:\n{header_txt}\n")

    # Check to see if the user wants the graph, and get the column indices if so
    (user_gc, time_col_ind, data_col_ind) = get_graph_info(save_as_xlsx, header_txt)
    # Determine graph buffer size
    if delay_ard == 0:
        buf_size = 0
        # Make sure the graph choice makes sense
        if user_gc == GraphChoice.LIVE:
            user_gc = GraphChoice.EXCEL_ONLY
    else:
        # Index 0 will be reserved for last value of the previous plot, so add 1 to the buffer size
        # The other +1 is to make sure at least 1 new value is plotted
        # (in the case that delay_ard > INTERVAL_PLOT)
        buf_size = int(INTERVAL_PLOT/delay_ard) + 1 + 1

    # Prepare structures for data
    graph_struct:GraphData = GraphData(user_gc, time_col_ind, data_col_ind, graph_pause, buf_size)
    file_struct:FileData = FileData(save_as_xlsx, file_name, header_txt)

    # Get and write data
    ser = serial.Serial(port, buad, timeout=(1.25*delay_ard))
    ser.close()
    get_and_write_data(ser, file_struct, graph_struct)
    # Print confirmation
    print("Done.")


# Run main()
if __name__ == "__main__":
    main()
