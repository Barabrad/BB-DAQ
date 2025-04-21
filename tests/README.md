# BB-DAQ/tests
This folder contains `test_BB_DAQ.py`, which will run automated tests on `BB_DAQ.py`.

## Documentation

### Introduction
Previously, I had to set up my Arduino and manually run through various tests. Now, this script will do that and output files for me to look at if the rest of the checks pass.

### Libraries
The libraries this script and `BB_DAQ.py` use are listed below, as well as the download instructions. **My assumption is that you already have Python 3 installed on your computer.** To check, open a terminal window and type `python3 -V`. If the output does not display a version number, try `python -V`. If the latter command works, use `python` and `pip` instead of `python3` and `pip3`, respectively. If neither command shows a version number, install Python 3 first, and then return here. To see which non-built-in libraries are already installed, open a terminal window and type `pip3 list`. Depending on your system, you may need to make a [virtual environment](https://docs.python.org/3/library/venv.html) to install these libraries. The `requirements.txt` file contains the non-built-in libraries, so you could type `pip3 install requirements.txt`.
1. pyserial
    * This library is imported in the code as "serial" instead of "pyserial," so make sure you do not have another library installed named "serial." If you do, open a terminal, uninstall it while running the files in this repository, and reinstall it afterwards (or if you're comfortable enough with Python, create a [virtual environment](https://docs.python.org/3/library/venv.html) for running the files in this repository).
    * This library is not built-in, so you need to open a terminal window and enter `python3 -m pip install pyserial` if you do not have the library.
2. xlsxwriter
    * This library is not built-in, so you need to open a terminal window and enter `pip3 install xlsxwriter` if you do not have the library.
3. matplotlib
    * This library is not built-in, so you need to open a terminal window and enter `pip3 install matplotlib` if you do not have the library.
4. pytest
    * This library is not built-in, so you need to open a terminal window and enter `pip3 install pytest` if you do not have the library.
5. enum
    * This library is built-in, so you should not need to install anything.
6. os
    * This library is built-in, so you should not need to install anything.
7. datetime
    * This library is built-in, so you should not need to install anything.
8. time
    * This library is built-in, so you should not need to install anything.
9. traceback
    * This library is built-in, so you should not need to install anything.
10. unittest
    * This library is built-in, so you should not need to install anything.

### Tutorial
If all of the libraries are installed, you are ready for the tutorial.

1. Upon navigating to the `tests` directory (or even its parent directory) in your terminal window, run the following command: `pytest -v`.

2. At this point, the tests will run and show if they passed or failed.
    * Note that a couple tests will generate graphs.

3. Ideally, all tests will pass. You can then double-check the output files in the `tests/out/` subdirectory.
