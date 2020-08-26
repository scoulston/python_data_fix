from csv_fix_base import input_csv_fix
from time import process_time
#runs csv_fix from csv_fix_base
start = process_time()
input_csv_fix()
end = process_time()
total_time_csv_fix = "Total time for csv fix: " + str(end - start)
#prints total time for the csv_fix to run. Commented out unless needed.
#print(total_time_csv_fix)


