'''
Brad Barakat
Made for testing BB_DAQ.py

The goal here is to automate as much testing as possible without needing to set up the Arduino.
A user would not need to see or even use this file.
'''

# Import standard libraries
from os import listdir, remove as os_rmv, makedirs
from os.path import normpath, join as os_join, split as os_split, isdir
from unittest.mock import patch
from time import sleep
# Import 3rd party libraries
import pytest
# Import BB_DAQ from src directory
from src import BB_DAQ


# Constants
DATA_HEADER = "Type,Date,Timer,Time,No.,Value"
DATA_ROW_START = f"{BB_DAQ.DATA_ROW},{BB_DAQ.DATE_WORD},{BB_DAQ.TIMER_WORD},{BB_DAQ.TIME_WORD}"
TEST_OUT_DIR = normpath(os_split(__file__)[0] + "/out/")


class SerialMock:
    """
    This class provides a simple mock for a pyserial serial object
    """

    def __init__(self, lines:list[str], delay_s:float) -> None:
        # Reverse the list of lines so we can just pop in O(1)
        if lines[-1].strip() != "":
            lines.append("") # So the last message signals a stop to BB-DAQ loops
        lines.reverse()
        self.line_stack = lines
        self.num_lines = len(lines)
        self.delay_s = delay_s

    def readline(self) -> bytes:
        """
        Returns an encoded string from the list of lines, and pops it from the list,
        while using time.sleep() to mimic the Arduino delay time
        (If the list is empty, "" will be returned)
        """
        sleep(self.delay_s)
        if self.num_lines == 0:
            raise StopIteration()
        line_i = self.line_stack[-1]
        self.line_stack.pop()
        self.num_lines -= 1
        return line_i.encode()

    def add_to_buffer(self, new_lines:list[str]) -> None:
        """
        A method that allows me to refresh the buffer without making another object
        """
        # Remember that the buffer is reversed to make pop easier
        self.line_stack.reverse()
        if self.num_lines > 0:
            self.line_stack.pop() # Get rid of ""
        for nl in new_lines:
            self.line_stack.append(nl)
        self.line_stack.append("") # Put "" back
        self.line_stack.reverse()
        self.num_lines += len(new_lines)

    def open(self) -> None:
        """
        A method that does nothing since this buffer is already "open"
        """

    def close(self) -> None:
        """
        A method that does nothing since this buffer isn't meant to be closed
        """


class TestClass:
    """
    The class containing the tests for BB_DAQ.py
    """

    def test_clear_files(self):
        """
        Use this method to clear the output directory to prevent overwrites
        This method must be run first!
        """
        if not isdir(TEST_OUT_DIR):
            makedirs(TEST_OUT_DIR)
        else:
            file_list = listdir(TEST_OUT_DIR)
            for filename in file_list:
                # Make sure the file is an Excel or CSV file
                if filename.endswith((".xlsx", ".csv")):
                    os_rmv(os_join(TEST_OUT_DIR,filename)) # Doesn't alter file_list
        assert True

    @pytest.mark.parametrize("test_row_type,expected_tuple", \
                             [(BB_DAQ.DATA_ROW, (BB_DAQ.DATA_ROW,False)), \
                              (BB_DAQ.LABEL_ROW, (BB_DAQ.LABEL_ROW,False)), \
                              (BB_DAQ.MSG_ROW, (BB_DAQ.MSG_ROW,False)), \
                              ("BadRowType", ("DEFAULT",True))])
    def test_get_row_type_and_num_cols(self, test_row_type, expected_tuple):
        """
        This method tests BB_DAQ.get_row_type_and_num_cols()
        """
        def_row = "DEFAULT"
        row = [test_row_type, "rest_of_data"]
        row_len = len(row)
        (row_type, num_cols, missing_label) = BB_DAQ.get_row_type_and_num_cols(row, def_row)
        assert row_type == expected_tuple[0]
        assert num_cols == row_len
        assert missing_label == expected_tuple[1]

    @pytest.mark.parametrize("test_tuple,expected", \
                             [(("7", float), True), \
                              (("-7.2", float), True), \
                              (("7", int), True), \
                              (("-7", int), True), \
                              (("-7.0", int), True), \
                              (("-7.1", int), False), \
                              (("ruh roh rhaggy", float), False)])
    def test_is_num_str(self, test_tuple, expected):
        """
        This method tests BB_DAQ.is_num_str()
        """
        res = BB_DAQ.is_num_str(test_tuple[0], test_tuple[1])
        assert res == expected

    @patch("builtins.input", side_effect=['ruh roh rhaggy', '3.2', '-4'])
    def test_get_int_input(self, _):
        """
        This method tests BB_DAQ.get_int_input() mainly to test input-patching
        Patching requires another argument, but it's unused, so I put _
        """
        res = BB_DAQ.get_int_input("pls send ints")
        assert res == -4

    def test_get_header_and_delay(self):
        """
        This method tests BB_DAQ.get_header_and_delay()
        """
        msg_list = [BB_DAQ.DATA_START_AFTER, DATA_HEADER, "random text"]
        sleep_time = 0.25 # s
        ser = SerialMock(msg_list, sleep_time)
        (header_txt, delay_ard, _) = BB_DAQ.get_header_and_delay(ser)
        assert header_txt == DATA_HEADER
        # We won't get delay_ard == sleep_time exactly, so we need a tolerance
        tol = 0.025*sleep_time
        assert -tol <= delay_ard - sleep_time <= tol

    @patch("builtins.input", side_effect=['Timer_Reset', '1', '0', 'history', 'Clear_Data', '1', \
           '1', normpath(f"{TEST_OUT_DIR}/test_gc_none_w2"), 'New_Wkbk', '0'])
    def test_get_and_write_data_excel_gc_none(self, _):
        """
        This method tests BB_DAQ.get_and_write_data() for an Excel output with no graph:
        - A sheet testing the timer reset, and another testing the data clear (and sheet naming)
        - A new workbook with row types to also test file-switching
        Patching requires another argument, but it's unused, so I put _
        """
        timer_reset_list = [BB_DAQ.DATA_START_AFTER, \
                DATA_HEADER, \
                f"{DATA_ROW_START},1,0", \
                f"{DATA_ROW_START},2,1", \
                BB_DAQ.RESET_TIMER, \
                f"{BB_DAQ.LABEL_ROW},{BB_DAQ.DATE_WORD},Label,Timer Reset", \
                f"{DATA_ROW_START},5,2", ""]
        clear_data_list = [BB_DAQ.DATA_START_AFTER, \
                DATA_HEADER, \
                f"{DATA_ROW_START},1,0", \
                f"{DATA_ROW_START},2,1", \
                BB_DAQ.CLEAR_DATA, \
                f"{BB_DAQ.LABEL_ROW},{BB_DAQ.DATE_WORD},Label,Data Cleared", \
                f"{DATA_ROW_START},5,2", ""]
        new_wkbk_rt_list = [BB_DAQ.DATA_START_AFTER, \
                DATA_HEADER, \
                f"{DATA_ROW_START},1,0", \
                f"{BB_DAQ.MSG_ROW},Message", \
                f"{BB_DAQ.LABEL_ROW},{BB_DAQ.DATE_WORD},Label,Message Sent", \
                # Intentionally left out a row type below
                f"{BB_DAQ.DATE_WORD},{BB_DAQ.TIMER_WORD},{BB_DAQ.TIME_WORD},4,1",
                f"{DATA_ROW_START},5,2", ""]
        msg_list = timer_reset_list + clear_data_list + new_wkbk_rt_list
        ser = SerialMock(msg_list, 0.25)
        fpath = normpath(f"{TEST_OUT_DIR}/test_gc_none_w1.xlsx")
        file_struct = BB_DAQ.FileData(True, fpath, DATA_HEADER)
        graph_struct = BB_DAQ.GraphData(BB_DAQ.GraphChoice.NONE,-1,-1,0,0)
        BB_DAQ.get_and_write_data(ser, file_struct, graph_struct)
        # DATA_START_AFTER, MSG_ROW, RESET_TIMER, and "" won't be shown in the file
        num_disp_rows = len(new_wkbk_rt_list) - 1 - 1 - 1 # This is 1-indexed
        # row_num is 0-indexed and incremented after each row is done
        assert file_struct.row_num == num_disp_rows

    @patch("builtins.input", side_effect=['Sheet_1', '0'])
    def test_get_and_write_data_excel_gc_sheet(self, _):
        """
        This method tests BB_DAQ.get_and_write_data() for an Excel output with the graph only in
        the sheet (not live)
        Patching requires another argument, but it's unused, so I put _
        """
        num_data_lines = 10
        msg_list = [BB_DAQ.DATA_START_AFTER, DATA_HEADER]
        for i in range(num_data_lines):
            msg_list.append(f"{DATA_ROW_START},{i},{(i-1)**2}")
        ser = SerialMock(msg_list, 0.25)
        fpath = normpath(f"{TEST_OUT_DIR}/test_gc_sheet.xlsx")
        file_struct = BB_DAQ.FileData(True, fpath, DATA_HEADER)
        graph_struct = BB_DAQ.GraphData(BB_DAQ.GraphChoice.EXCEL_ONLY,4,5,0,0)
        BB_DAQ.get_and_write_data(ser, file_struct, graph_struct)
        # DATA_START_AFTER, MSG_ROW, RESET_TIMER, and "" won't be shown in the file
        num_disp_rows = num_data_lines + 1 # This is 1-indexed
        # row_num is 0-indexed and incremented after each row is done
        assert file_struct.row_num == num_disp_rows
        assert len(file_struct.curr_sheet.charts) == 1

    @patch("builtins.input", side_effect=['Sheet_1', '0'])
    def test_get_and_write_data_excel_gc_live(self, _):
        """
        This method tests BB_DAQ.get_and_write_data() for an Excel output with the live graph
        Patching requires another argument, but it's unused, so I put _
        """
        num_data_lines = 10
        msg_list = [BB_DAQ.DATA_START_AFTER, DATA_HEADER]
        for i in range(num_data_lines):
            msg_list.append(f"{DATA_ROW_START},{i},{(i-1)**2}")
        ser = SerialMock(msg_list, 0.25)
        fpath = normpath(f"{TEST_OUT_DIR}/test_gc_live.xlsx")
        file_struct = BB_DAQ.FileData(True, fpath, DATA_HEADER)
        graph_struct = BB_DAQ.GraphData(BB_DAQ.GraphChoice.LIVE,4,5,0.1,5)
        BB_DAQ.get_and_write_data(ser, file_struct, graph_struct)
        # DATA_START_AFTER, MSG_ROW, RESET_TIMER, and "" won't be shown in the file
        num_disp_rows = num_data_lines + 1 # This is 1-indexed
        # row_num is 0-indexed and incremented after each row is done
        assert file_struct.row_num == num_disp_rows
        assert len(file_struct.curr_sheet.charts) == 1
        assert graph_struct.num_plot_bufs == 2

    @patch("builtins.input", side_effect=['1', normpath(f"{TEST_OUT_DIR}/test_gc_none_f2"), '0'])
    def test_get_and_write_data_csv_gc_none(self, _):
        """
        This method tests BB_DAQ.get_and_write_data() for a CSV output with no graph:
        - A file testing the data clear and timer reset
        - A new file with row types to also test file-switching
        Patching requires another argument, but it's unused, so I put _
        """
        clear_timer_list = [BB_DAQ.DATA_START_AFTER, \
                DATA_HEADER, \
                f"{DATA_ROW_START},1,0", \
                BB_DAQ.CLEAR_DATA, \
                f"{DATA_ROW_START},3,1", \
                BB_DAQ.RESET_TIMER, \
                f"{BB_DAQ.LABEL_ROW},{BB_DAQ.DATE_WORD},Label,Timer Reset", \
                f"{DATA_ROW_START},6,2", ""]
        new_file_rt_list = [BB_DAQ.DATA_START_AFTER, \
                DATA_HEADER, \
                f"{DATA_ROW_START},1,0", \
                f"{BB_DAQ.MSG_ROW},Message", \
                f"{BB_DAQ.LABEL_ROW},{BB_DAQ.DATE_WORD},Label,Message Sent", \
                # Intentionally left out a row type below
                f"{BB_DAQ.DATE_WORD},{BB_DAQ.TIMER_WORD},{BB_DAQ.TIME_WORD},4,1",
                f"{DATA_ROW_START},5,2", ""]
        msg_list = clear_timer_list + new_file_rt_list
        ser = SerialMock(msg_list, 0.25)
        fpath = normpath(f"{TEST_OUT_DIR}/test_gc_none_f1.csv")
        file_struct = BB_DAQ.FileData(False, fpath, DATA_HEADER)
        graph_struct = BB_DAQ.GraphData(BB_DAQ.GraphChoice.NONE,-1,-1,0,0)
        BB_DAQ.get_and_write_data(ser, file_struct, graph_struct)
        # DATA_START_AFTER, MSG_ROW, RESET_TIMER, and "" won't be shown in the file
        num_disp_rows = len(new_file_rt_list) - 1 - 1 - 1 # This is 1-indexed
        # row_num is 0-indexed and incremented after each row is done
        assert file_struct.row_num == num_disp_rows

    @patch("builtins.input", side_effect='0')
    def test_get_and_write_data_csv_gc_sheet(self, _):
        """
        This method tests BB_DAQ.get_and_write_data() for a CSV output with the graph only in the
        sheet (not live), which should be no different than the no graph case due to file type
        Patching requires another argument, but it's unused, so I put _
        """
        num_data_lines = 10
        msg_list = [BB_DAQ.DATA_START_AFTER, DATA_HEADER]
        for i in range(num_data_lines):
            msg_list.append(f"{DATA_ROW_START},{i},{(i-1)**2}")
        ser = SerialMock(msg_list, 0.25)
        fpath = normpath(f"{TEST_OUT_DIR}/test_gc_sheet.csv")
        file_struct = BB_DAQ.FileData(False, fpath, DATA_HEADER)
        graph_struct = BB_DAQ.GraphData(BB_DAQ.GraphChoice.EXCEL_ONLY,4,5,0,0)
        BB_DAQ.get_and_write_data(ser, file_struct, graph_struct)
        # DATA_START_AFTER, MSG_ROW, RESET_TIMER, and "" won't be shown in the file
        num_disp_rows = num_data_lines + 1 # This is 1-indexed
        # row_num is 0-indexed and incremented after each row is done
        assert file_struct.row_num == num_disp_rows
        with pytest.raises(AttributeError):
            _ = len(file_struct.curr_sheet.charts)

    @patch("builtins.input", side_effect='0')
    def test_get_and_write_data_csv_gc_live(self, _):
        """
        This method tests BB_DAQ.get_and_write_data() for a CSV output with the live graph
        Patching requires another argument, but it's unused, so I put _
        """
        num_data_lines = 10
        msg_list = [BB_DAQ.DATA_START_AFTER, DATA_HEADER]
        for i in range(num_data_lines):
            msg_list.append(f"{DATA_ROW_START},{i},{(i-1)**2}")
        ser = SerialMock(msg_list, 0.25)
        fpath = normpath(f"{TEST_OUT_DIR}/test_gc_live.csv")
        file_struct = BB_DAQ.FileData(False, fpath, DATA_HEADER)
        graph_struct = BB_DAQ.GraphData(BB_DAQ.GraphChoice.LIVE,4,5,0.1,5)
        BB_DAQ.get_and_write_data(ser, file_struct, graph_struct)
        # DATA_START_AFTER, MSG_ROW, RESET_TIMER, and "" won't be shown in the file
        num_disp_rows = num_data_lines + 1 # This is 1-indexed
        # row_num is 0-indexed and incremented after each row is done
        assert file_struct.row_num == num_disp_rows
        with pytest.raises(AttributeError):
            _ = len(file_struct.curr_sheet.charts)
        assert graph_struct.num_plot_bufs == 2
