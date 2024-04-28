# VBA_challenge - Stock Analysis 

## Description
This VBA script loops through all the stocks for each quarter within an Excel workbook and outputs the following information:
- The ticker symbol
- Quarterly change from the opening price at the beginning of a given quarter to the closing price at the end of that quarter
- The percentage change from the opening price at the beginning of a given quarter to the closing price at the end of that quarter
- The total stock volume of the stock

Additionally, functionality has been added to the script to return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume". The script runs on every worksheet (quarter) in the workbook at once.

## Files
- `StockAnalysis.vbs`: VBA script file containing the stock analysis code
- `README.md`: This file

## Usage
1. Open the Excel workbook containing the stock data.
2. Press `Alt + F11` to open the VBA editor.
3. Insert a new module and paste the contents of `StockAnalysis.vbs` into the module.
4. Save the workbook with macros enabled.
5. Run the `StockAnalysis` subroutine to execute the script.

## Results
Screenshots of the results can be found in the gitlab submission.



