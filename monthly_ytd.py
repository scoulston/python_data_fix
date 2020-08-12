from csv_fix_base import input_csv_fix, default_base_file_path
import pandas as pd
from datetime import date, datetime
from time import process_time


#function that makes a csv file with summary of data (total sales and lbs shipped per month for each company, including a column for the salesman)
def input_monthly_ytd_summary(input_csv_fix=input_csv_fix):
    #runs input_csv_fix to fix a csv, then takes the fixed file name to open it and work with the fixed data to start summarizing
    fixed_file_name = input_csv_fix()
    #uses the fixed file name to open the fixed file
    summary_data = pd.read_csv(default_base_file_path + fixed_file_name + ".csv")
    start = process_time()
    #takes the month and year from the data and strings them together into "YR Month" format
    summary_data['Month_YR'] = summary_data['YR'].astype(str) + " " + summary_data['Month'].astype(str)
    #makes a list of the month_YR column's unique values to use later in renaming meshed columns
    months = summary_data['Month_YR'].unique().tolist()
    #changes the list months to datetime so we can get the wanted formatting later when renaming meshed columns
    for i in range(len(months)):
        months[i] = datetime.strptime(months[i], '%Y %m')
    end = process_time()
    fix_date_as_datetime_time = "Time to fix date column and store as datetime: " + str(end - start)

    start = process_time()
    #groups the data by company, salesperson, and month_yr
    data_by_month_and_comp = summary_data.groupby(['Company', 'Salesperson', 'Month_YR']).agg(
    {
         'Sales':sum, # Sum Sales per group
         'Pounds Shipped': sum,  # Sum Pounds Shipped per group
            }

).reset_index()
    end = process_time()
    group_by_table_time = "Time to group by the table: " + str(end - start)


    #uncomment the below line to get a file (group_by_test.csv) to test the above grouped table
    #data_by_month_and_comp.to_csv(default_base_file_path + "group_by_test.csv")

    #found this function online that supports pivoting multiindex, multilevel tables (where duplicated indices are not a problem)
    def multiindex_pivot(df, index=None, columns=None, values=None):
        if index is None:
            names = list(df.index.names)
            df = df.reset_index()
        else:
            names = index
        list_index = df[names].values
        tuples_index = [tuple(i) for i in list_index] # hashable
        df = df.assign(tuples_index=tuples_index)
        df = df.pivot(index="tuples_index", columns=columns, values=values)
        tuples_index = df.index  # reduced
        index = pd.MultiIndex.from_tuples(tuples_index, names=names)
        df.index = index
        return df
    
    start = process_time()
    #uses the above function to make a multiindex, multilevel table with the company and salesperson as indexes, the month-yr as one level of columns, the sales and pounds shipped as another level of columns
    data_by_month_and_comp_pivot = multiindex_pivot(data_by_month_and_comp, index = ['Company', 'Salesperson'], columns= 'Month_YR', values = ['Sales', 'Pounds Shipped']).reset_index() 
    end = process_time()
    multiindex_table_time = "Time to make the multiindex table: " + str(end - start)

    start = process_time()
    #fixes number months to abbreviated months in column labels using the sorted months list and converting it to the datetime format we want while still keeping the columns in the right order
    data_by_month_and_comp_pivot.columns.set_levels([x.strftime('%b %Y') for x in sorted(months)], 1, inplace=True, verify_integrity = False)
    end = process_time()
    fix_months_column_headers_time = "Time to fix the months in column headers: " + str(end - start)

    start = process_time()
    #adds columns for the total sales and total lbs for that company/salesperson during the timeperiod the data runs across 
    data_by_month_and_comp_pivot.insert(2, 'Total Sales', data_by_month_and_comp_pivot['Sales'].sum(axis=1))
    data_by_month_and_comp_pivot.insert(3, 'Total Lbs', data_by_month_and_comp_pivot["Pounds Shipped"].sum(axis=1))
    end = process_time()
    make_total_sales_and_lbs_time = "Time to make the total sales and lbs columns: " + str(end - start)

    start = process_time()
    #fixes the stacked columns into one column name per column (meshes the column level names)
    data_by_month_and_comp_pivot.columns = list(map(" ".join, data_by_month_and_comp_pivot.columns))
    data_by_month_and_comp_pivot.reset_index()
    end = process_time()
    merge_levels_time = "Time to merge levels on columns: " + str(end - start)

    start = process_time()
    #saves the data to the fixed csv filename 
    data_by_month_and_comp_pivot.to_csv(default_base_file_path + fixed_file_name + ".csv", index= False)
    end = process_time()
    ot_export_fixed_time = "Time ot export as fixed file: " + str(end - start)

    #prints all the times to do certain sections of code. Commented out unless needed.
    #print(fix_date_as_datetime_time)
    #print(group_by_table_time)
    #print(multiindex_table_time)
    #print(fix_months_column_headers_time)
    #print(make_total_sales_and_lbs_time)
    #print(merge_levels_time)
    #print(ot_export_fixed_time)

start2 = process_time()   
input_monthly_ytd_summary()
end2 = process_time()
#prints total time to run input_monthly_ytd_summary(). Commented out unless needed.
#print("Total time: " + str(end2 - start2))



