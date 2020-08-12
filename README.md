# python_data_fix
Fixes for oddly formatted data for a client
Anonymized data and code posted with permission of the client

This is a fix I made for a client. The data they could download was weirdly and unhelpfully formatted.
The base_fix file is the base code, the actual fix file runs the command to fix the data, the monthly_ytd file has the code that runs and makes a summary from some unfixed data (it also runs the base fix), and I have included a sample file with what the data looked like before, what the data looks like after the fix, and what the summary looks like when you run it on some data. 

The csv_fix fixes the following issues according to the employee's wishes:
-The date is fixed into three separate columns for month, date, and year from a single date column 
-It fixes the date so python recognizes each column's values as part of a date that can be used elsewhere (particularly when doing the aggregated summary data program). It makes it easier to format. 

-The item description was a bunch of combined data separated by an asterisk. The problem was that some rows had different amounts of information (some had OD, ID, and Length, while others had only Width and Length), which made it difficult to separate based on the asterisk via Excel. The csv_fix takes the item description and inserts delimiters as needed for each row to make sure each item description would have the same number of columns and in the correct places when separating into different columns by the delimiter. It also changes the delimiter to a comma so it fit back into the overall file as already separated columns.

-The data comes originally where there are three rows for each order, with everything the same except the last few rows where one row shows the quantity shipped, another the pounds shipped, and another showing the sale amount. The program fixes the data where all of the information for each order is in one row instead of three nearly identical rows. 

-It renames columns into a more readable format according to the client's wishes.
-Uneeded columns are deleted
-The company names have the either two or three letter prefix with a dash removed (depending on which prefix they have), then only keeps the first 20 characters after that (a specific request of the client)
-The columns are ordered automatically according to the client's wishes. 

The monthly_ytd is a program that aggregates the data according to the client's wishes (by company, month and year, and salesperson). It shows the total pounds and total sales for each group and includes a column for the total pounds and total sales per company (divided by salesperson if more than one salesperson sold shipments for that company) across the entire period the data covers. It also already includes the data fix so if the client wants the summary they only have to run the one program.

