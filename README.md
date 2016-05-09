# Binary_code_format_Excel
Excel Macro that converts binary code from single cell input to multiple cells.

## Problem Statement
Binary content from orthogonal arrays can be used to run calculations in Excel. However, often importing these arrays into Excel can create problems, resulting in array content being imported into a single Excel cell per array row.

## Solution
The Excel formula `MID` can be used to solve this problem. The included Excel Macro automates the process and can be used for importing arrays of dimensions up to **144x144**.

## Usage Instruction
- Import text file with the array content into Excel, formated as text (or leading zeroes might get lost). Direct typing into each row of column D is also ok.
- Open the Excel Macro File and allow macros to be enabled.
- Copy-Paste all cells into **column D** of the Excel Sheet **Input**.
- Click on `Format` to get the result on the **Output** Sheet.

## Limitations
- Make sure that macros are allowed in Excel before using this program.
- Maximum Dimensions: **144x144**
- This macro has been written and tested with MS Excel, 2016. It might not work with other versions of Excel.

## Warranty & Disclaimer
This macro program may be freely distributed, provided that no charge are levied. The program is provided as is without any guarantees or warranty. Although the author has attempted to find and correct any bugs in this macro program, the author is not responsible for any damage or losses of any kind caused by the use or misuse of the programs. The author is under no obligation to provide support, service, corrections, or upgrades to this macro program.
