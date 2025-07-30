# Excel Sudoku Solver

A formula-based Sudoku puzzle solver built entirely in Microsoft Excel without any VBA or macros. Once I learned about Excel's iterative calculation setting, I had to build a little project to understand how it can be used.

https://github.com/user-attachments/assets/d636a186-fcb4-4cf6-a403-fd83e8a53a27

## Overview

This Excel workbook can solve easy Sudoku puzzles using only Excel formulas. It applies logical constraints to progressively eliminate impossible values until only one possibility remains for each cell.

## Features
- Pure formula-based sudoku solution (no VBA/macros required!)

## Limitations

Unfortunately, this solver only really works with easy to medium difficulty puzzles. This is partly due to the limit that you manually impose on the iterative calculation setting. Obviously, we don't want to set this to a number too large, or else we will be stuck in a brute-force loop that might not terminate. Excel does not crash with as much grace as some other programs.

Difficult puzzles requiring advanced techniques like "X-Wing" or "Swordfish" probably will not be solved with this formula-based approach, though I imagine difficult solves are possible.

## How It Works

The solver consists of two worksheets:

1. **Interface Sheet**: Where you enter your Sudoku puzzle and view the solution
2. **Calculation Engine Sheet**: Where all the logic happens

The Calculation Engine maps each cell of the Sudoku grid to a row and applies three types of constraints:
- Row constraints (no repeated values in a row)
- Column constraints (no repeated values in a column)
- Box constraints (no repeated values in a 3x3 box)

If X-Wing and Swordfish logic were to be imposed, it would be here.

## Usage Instructions

1. Enter your Sudoku puzzle in cells B2:J10
2. The solution will appear in cells L2:T10, typically in less than one second
3. The status indicator in cell C12 will inform the user if the puzzle is solved
4. To enter a new puzzle, simply clear cells B2:J10 and enter new values

## Requirements

- Microsoft Excel (2016 or newer recommended)
- Iterative calculations enabled (File > Options > Formulas > Enable iterative calculation)
