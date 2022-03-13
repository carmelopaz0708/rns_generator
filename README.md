# RNS Generator
A command-line script for generating random number sequences from an external reference table.

## Description
Python script that generates a random sequence of integers. The script infers its sequence values from an external reference file stored in `rns_generator\tables\`. All other input parameters should be provided by the user during runtime. This script can be used for probability sampling and can generate sequences for random, systematic and monetary sampling routines. An external reference table is **required** when running this script and should be provided by the user.

## Dependencies
The latest version of [Python](https://www.python.org/downloads/) must be installed and configured in your operating system in order to run this script. Next, you will need to install the latest version of `pip` (Python's package installer). To install `pip`, please refer to the manual [here](https://pip.pypa.io/en/stable/installation/). If you already have `pip`, skip this step. 

Lastly are the script's package dependencies, which are as follows:
1. [openpyxl](https://pypi.org/project/openpyxl/)
2. [et-xmlfile](#)

Note that `et-xmlfile` ships with `openpyxl` when installing the latter as a package dependency. That means you will only ever need to `openpyxl` to complete this list. Run the command shown below to your terminal of choince to install the package dependencies.

```python
pip install openpyxl
```

## Debugging
The best way to read and/or debug this program is through Python's virtual environment. The name of the virtual environment is irrelevant (Set it to any name you want. Personally, I use `env`).

### Enabling Execution Policy (Windows 10 and above)
Windows 10 and above restricts running custom scripts by default. This is an added safety feature that prevents unauthorized users from executing malicious code into your system. In order to run this program, we can either do one of two things. One, set Window's execution policy to `Bypass` on `Process` scope. Or two, set the execution policy to `RemoteSigned` on `CurrentUser` scope. The first option allows all scripts to run on only the current terminal session, and will return to its default state (restricted) after the terminal is closed. The second option ensures you can run the script even after ending your terminal session, since the execution policy is saved in your registry subkey.

Basically, choose option one when you read/debug this code infrequently or if you do not own the computer you're using to run this script. Otherwise, choose option two. Note that for option one, you'll have to re-enable the execution policy if you plan to run the script again after you've closed a terminal.

The command below sets the execution policy to `Bypass` on `Process` scope (option one).
```powershell
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process
```

This other command sets the execution policy to `RemoteSigned` on `CurrentUser` scope (option two). 
```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

You can read more of Window's execution policy [here](https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_execution_policies?view=powershell-7.1).

## Preparation
### Importing tables
As mentioned, external reference files, called *tables*, should be stored in the `rns_generator\tables\` directory. Tables should be in standard Excel format. Allowable file extensions for the reference file should be in `.xlsx`, `.xlsm`, `.xltx` or `.xltm` only. The script will not read any other file types not included in this list.

Likewise, all cells in the reference file comprising the *table* should contain integers only. **Do not include row nor column headers in the reference file**. Do not include letters, symbols and/or special characters in the *table*. Doing so will prevent the code from parsing the *table* and will undoubtedly crash the program. Do not use special formatting on the *table* as well; just enter the data in each cell until you complete the *table*. You can opt to format the spacing in each cell of the reference table for better readability; doing so has no effect on the program.

Lastly, while not explicitly required, is highly recommended to name all the available worksheets in your reference file. The script will parse worksheet names and help the user disambiguate from each table. 

### Running the script
Clone this repo and save it to any directory in your system. From the terminal, run a `cd` command to the repo. Make sure you are in `rns_generator\`. You can check your current directory by running a `dir` command on Windows or an `ls` command on Mac or Linux.

Once you're in `rns_generator`, enter the code below. Change *reference.xlsx* to your table's file name and provide the proper file extension.

```python
python source.py reference.xlsx
```

### Obtaining output files
Output files are named automatically as *output.txt*. They are stored in the `rns_generator\output` directory. Note that new output files will replace old ones in the directory every time the script is executed to completion (when the final prompt says "exit with success"). If you want to retain these files, you'll have to manually recover them from the directory before running a new instance of the script.

## Copyright
Copyright (C) 2021  Carmelo Julius Paz
This program comes with ABSOLUTELY NO WARRANTY. This is free software, and you are welcome to redistribute it under certain conditions.

This program is free software: you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation, either version 3 of the License, or (at your option) any later version.

This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.

You should have received a copy of the GNU General Public License along with this program.  If not, see <https://www.gnu.org/licenses/>.
