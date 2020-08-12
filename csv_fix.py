from csv_fix_base import input_csv_fix
from time import process_time
#runs csv_fix from csv_fix_base
start = process_time()
input_csv_fix()
end = process_time()
total_time_csv_fix = "Total time for csv fix: " + str(end - start)
#prints total time for the csv_fix to run. Commented out unless needed.
#print(total_time_csv_fix)


#1. Download Data file (not crosstab) from your tableau
#2. Save to One Drive (the charter one, not your personal one) or email to yourself
#3. On the computer with Python on it, open your OneDrive and download the file to Downloads (or the attachment on the email)
#4. Open Visual Studio Code
#5. If not on Dad Workspace, click 'File', 'Open Workspace' and 'Dad Workspace', else continue
#6. click on the csv_fix.py file under dad-master
#7. Click the Play button at the top right
#8. Copy the path of the file you downloaded and paste it where it tells you to (don't delete the r before the quotation!)
#9. Enter what you want as the name of the fixed file when and where it tells you to.
#10. The fixed file will be in your One Drive under Fixed Data