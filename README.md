# Read_PR_source_data
A Jupyter Notebook to iterate through a number of Excel files stored in the same folder, combine them into a pandas dataframe and pivot and resample the data into daily and monthly frequencies.

## 1. Description
The O&M operator of this solar PV project provides monthly reports composed of 5-minute generation and irradiation data to calculate the Performance Ratio (PR%) of the facility. Each month a new Excel file is added to a repository that stores previous reports. The challenge comes when the historical reports have to be aggregated to produce lower frequency reports (e.g. yearly).

This script iterates through all Excel files included in the repository, combining them into a single pandas dataframe with +100,000 rows, which is practically unmanageable in Excel. A pandas dataframe can however operate with such large datasets in a matter of seconds.

Once the dataframe is built, then the 5-minute data is pivoted and aggregated into hourly and daily dataframes, and then exported to Excel files that can be subsequently worked with much more easily. For instance, the calculation of hourly and daily PR% is possible directly in Excel.

## 2. How to use the script
The script is written in Python in a Jupyter notebook. The script will look for the monthly Excel reports in a folder called "1_Data". After combining, pivoting and aggregating the data, the consolidated Excel file will be produced in the same folder as the script.

## 3. Output
The script output is a .xlsx file with four tabs summarizing data in daily and hourly frequencies.
