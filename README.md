# Excel Sudoku Solver

A formula-based Sudoku puzzle solver built entirely in Microsoft Excel without any VBA or macros.

## Overview

This Excel workbook can solve most Sudoku puzzles using only Excel formulas. It applies logical constraints to progressively eliminate impossible values until only one possibility remains for each cell.

## Features

- Pure formula-based solution (no VBA/macros required)
- Works for easy to medium difficulty puzzles
- Automatic solving with iterative calculations
- Clean interface with status indicators
- Easy to use with any valid Sudoku puzzle

## How It Works

The solver consists of two worksheets:

1. **Interface Sheet**: Where you enter your Sudoku puzzle and view the solution
2. **Calculation Engine Sheet**: Where all the logic happens

The Calculation Engine maps each cell of the Sudoku grid to a row and applies three types of constraints:
- Row constraints (no repeated values in a row)
- Column constraints (no repeated values in a column)
- Box constraints (no repeated values in a 3x3 box)

It uses Excel's iterative calculation capability to repeatedly apply these constraints until the puzzle is solved.

## Usage Instructions

1. Enter your Sudoku puzzle in cells B2:J10
2. The solution will appear in cells L2:T10
3. Check the status indicator in cell C12 to see if the puzzle is solved
4. To enter a new puzzle, simply clear cells B2:J10 and enter new values

## Requirements

- Microsoft Excel (2016 or newer recommended)
- Iterative calculations enabled (File > Options > Formulas > Enable iterative calculation)

## Limitations

This solver works best with easy to medium difficulty puzzles. Very difficult puzzles requiring advanced techniques like "X-Wing" or "Swordfish" may not be solved completely with this formula-based approach.

## Creator

Created by Anthony Gallante using Excel's built-in formula capabilities.

## License

Feel free to use, modify, and share this workbook.
