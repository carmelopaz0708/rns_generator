# RNS Generator
A command-line script for generating random number sequences from an external reference table.

## Description
Python script that generates a random sequence of integers. The script infers its sequence values from an external reference file stored in `rns_generator\tables\`. All other parameters should be provided by the user at runtime. This script can be used for probability sampling and can generate sequences for random, systematic and monetary sampling routines. An external reference table is **required** when running this script should be provided by the user.

## Dependencies
The latest version of [Python](https://www.python.org/downloads/) is required to run this script. Additionally, you will need to install the latest version of the [openpyxl](https://pypi.org/project/openpyxl/) library. This will allow the script to load tables. In order to get `openpyxl`, run the code below into your terminal.

```python
pip install openpyxl
```

Note that [pip](https://pypi.org/project/pip/) is required to install external Python packages. To install `pip`, please refer to [pip installation](https://pip.pypa.io/en/stable/installation/).

## Preparation
### Enable execution policy
Windows, by default, restricts running custom scripts. This is an added safety feature by Windows to prevent malicious scripts from booting into your terminal and destroying precious files and configurations. In order to run this program, we'll have to set Window's execution policy to `Bypass` every time we start up. You'll only ever need to do this once, and the execution policy will remain active until the terminal session is eventually closed. To start, run the code below into your terminal.

```powershell
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process
```

This is only a temporary solution. A future update will patch this startup workaround. You can also choose to enable bypass scope from `Process` to `MachinePolicy` or `UserPolicy`, however, this is highly discouraged as it will leave your system extremely vulnereable to running dangerous or malicious scripts. 

### Preparing tables
As mentioned, external reference files, called *tables*, should be stored in the `rns_generator\tables\` directory. Tables should be in standard Excel format. Allowable file extensions for this type are `.xlsx`, `.xlsm`, `.xltx` and `.xltm`. The script will not read any other file types that weren't mentioned previously.

Likewise, all cells in the reference file comprising the *table* should contain integers only. **Do not include row nor column headers in the reference file**. Do not include letters, symbols nor special characters in the *table*. Doing so will prevent the code from parsing the *table* and will undoubtedly crash the script. No formatting is required for the reference file either; just enter the data until it forms the *table*. You can opt to format the cells in the reference table for better readability as doing so will have no effect on the script parsing data.

### Running the script
Clone this repo and save it to any directory in your system. From the terminal, run a `cd` command to the repo. Make sure you are in `rns_generator\`. You can check your current directory by running a `dir` command. 

Once you're in `rns_generator`, enter the code below. Change *reference.xlsx* to your table's file name.

```python
python source.py reference.xlsx
```

### Obtaining output files
Output files are named automatically as *output.txt*. They are stored in the `rns_generator\output` directory. Note that new output files will replace old ones in the directory every time the script is executed to completion (when the final prompt says "exit with success"). If you want to retain these files, you'll have to manually recover them from the directory before running a new instance of the script.

## Copyright
(C) 2021 Carmelo Julius Paz
(email) carmelopaz0708@gmail.com
(github) [Carmelo Paz](https://github.com/carmelopaz0708)

This program is free software; you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.

This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.

You should have received a copy of the GNU General Public License along with this program; if not, write to the Free Software Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA 02110-1301 USA.

Released under the GNU General Public License (GPLv2)
