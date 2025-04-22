# BB-DAQ
This repository contains `BB_DAQ.py`, a (limited) workaround for PLX-DAQ meant for Mac that also works on Windows.

## Repository Structure

This repository contains the following directories and files:
- `src/`
- `tests/`
- `requirements.txt`
- `.pylintrc`

### src
The `src` directory contains the following files:
- `__init__.py`
  - This file is blank. It was added to make imports easier during testing.
- `BB_BoardTester.py`
  - This file is used to determine the properties of the board/microcontroller being used.
- `BB_DAQ.py`
  - This file is the PLX-DAQ workaround.
- `requirements.txt`
  - This file contains the Python libraries to import.
- `README.md`
  - This file gives a thorough description of the other files in the directory (besides `__init__.py` and `requirements.txt`), along with tutorials.

### tests
The `tests` directory contains the following files:
- `__init__.py`
  - This file is blank. It was added to make imports easier during testing.
- `test_BB_DAQ.py`
  - This file runs automated tests on `BB_DAQ.py` using [pytest](https://docs.pytest.org/en/stable/).
- `requirements.txt`
  - This file contains the Python libraries to import.
- `README.md`
  - This file gives a thorough description of the other files in the directory (besides `__init__.py` and `requirements.txt`).

### requirements
The `requirements.txt` file lists all of the libraries needed to run any file in the repository.

### pylintrc
The `.pylintrc` file is a configuration file for [pylint](https://www.pylint.org), which will be used by GitHub in the pipeline I am setting up. I had to alter a few settings to make sure that the current functional code gets a 10/10.
