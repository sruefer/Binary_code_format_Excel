# Binary_code_format_Excel
Excel Macro that converts binary code from single cell input to multiple cells.

## Problem Statement
Binary content from orthogonal arrays can be used to run calculations in Excel. However, often importing these arrays 
into Excel can create problems, resulting in array content being imported into a single Excel cell per array row.

## Solution
The Excel formula `MID` can be used to solve this problem. The included Excel Macro automates the process and can be 
used for importing arrays of dimensions up to 128x128.

## Usage Instruction
- Import text file with the array content into Excel, formated as text (or leading zeroes might get lost).
- Open the Excel Macro File and allow macros to be enabled.
- Copy-Paste all cells into **column D** of the Excel Sheet **Input**.
- Click on `Format` to get the result on the **Output** Sheet.
