# vbachallenge
J. Savage
In this challenge, I wrote a script which will make the following things happen:
    1.) The script will create columns called "Ticker Name", "Yearly Change", "Percent Change", and "Total Stock Volume" as well as a column for values being calculated ("Greatest % Increase", "Greatest % Decrease", and "Greatest Total Volume" and the corresponding "Ticker" and "Value".
    2.) The script will loop through all the names of stocks and total up the change in value for that stock over the year. In addition, it will calculate the percent change and total stock volume within the year. It will then place those values into the corresponding columns of "Ticker Name", "Yearly Change", "Percent Change", and "Total Stock Volume".
    3.) The script will change the conditional formatting of the "Yearly Change" Column (K) based on the the value of the cell to either green for positive change or red for negative change. Zero change will remain unfilled.
    4.) The script will look through all the rows in the new columns from number 2 and calculate the ticker with the greatest % increase, the ticker with the greatest % decrease, and the ticker with the greatest total volume. It will paste the corresponding values into the "Ticker" and "Value" columns.
    5.) The script will format each of the columns on the sheet to fit the size of the values.
    6.) The script will repeat numbers 1 through 5 on each sheet in the file.
    
Sources:
https://support.microsoft.com/en-us/help/142126/macro-to-loop-through-all-worksheets-in-a-workbook

https://stackoverflow.com/questions/45510730/vba-how-to-convert-a-column-to-percentages
