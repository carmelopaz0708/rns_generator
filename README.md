# RNS Generator

A program that creates a random sequence of number from a set algorithm

# Preparation

Before using the program, ensure ```Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process```. Then, activate virtual environment with ```env/Scripts/activate```.

# Troubleshooting
1. Unresolved import issue.
Set the interpreter to use Python virtual environment

2. Cannot run script.
This problem is inherent with Windows users. By default, Windows prohibits running unauthorized scripts. To get around this problem, make sure to set your execution policy to `Bypass` on `Process` scope. Note that this is only a temporary solution, and closing the command line will end the bypass (meaning you have to set execution policy again). A future update will patch this issue (either set this script to allsigned or have scope execution policy to unrestricted, the latter being prone to a security breach).