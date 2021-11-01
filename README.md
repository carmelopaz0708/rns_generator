# RNS Generator
A command-line script for generating random number sequences from an external reference table.

## Description
Python script that generates a random sequence of integers. The script infers its sequence values from an external reference file stored in `rns_generator\tables\`. All other parameters should be provided by the user at runtime. This script can be used for probability sampling and can generate sequences for random, systematic and monetary sampling routines. An external reference table is **required** when running this script should be provided by the user.

## Preparation
**Preparing reference files**
As mentioned, external reference files, called *tables*, should be stored in `rns_generator\tables\` directory. Reference files should be in standard Excel 2010 format. Allowable file types for this 


# Preparation

Before using the program, ensure ```Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process```. Then, activate virtual environment with ```env/Scripts/activate```.

# Troubleshooting
1. Unresolved import issue.
Set the interpreter to use Python virtual environment

2. Cannot run script.
This problem is inherent with Windows users. By default, Windows prohibits running unauthorized scripts. To get around this problem, make sure to set your execution policy to `Bypass` on `Process` scope. Note that this is only a temporary solution, and closing the command line will end the bypass (meaning you have to set execution policy again). A future update will patch this issue (either set this script to allsigned or have scope execution policy to unrestricted, the latter being prone to a security breach).