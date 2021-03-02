# vba-challenge

BACKGROUND:
- The purpose of this homework was to showcase a variety of skills learned during the VBA portion of Northwestern's Data Science Bootcamp. The code created was practiced on a smaller dataset, and then applied to a larger dataset that used stock market data.


INSTRUCTIONS:
Create a script that will loop through all the stocks for one year and output the following information.
- The ticker symbol.
- Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
- The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
- The total stock volume of the stock.
- You should also have conditional formatting that will highlight positive change in green and negative change in red.

PROCESS:
- Used the alphabetical_testing.xlsx worksheet to test drive the code created.
- Started by creating individual subroutines for each bullet point in the instructions to ensure that each piece of code worked individually. I was also able to isolate issues with my code more easily.
- Once each piece of code worked within it's own subroutine, I was able to merge the code into one large subroutine.

CHALLENGES:
- The code did not populate results into each row. Instead, the results appeared  in one row. To address this, Erin suggested I add k = k + 1 to some output fields, which worked.


RESOURCES:
- The following classroom exercises were used to help develop this code: Wells Fargo, Formatter, and Credit Card Checker.
- I also Googled methods to apply VBA code to multiple worksheets.
