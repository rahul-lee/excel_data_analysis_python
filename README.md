
Overview
This script is a Flask-based application that performs data analysis on an Excel file containing sales data. It dynamically identifies relevant columns, cleans the data, and performs various analyses, saving the results in an Excel file with multiple sheets.
________________________________________
Prerequisites
1.	Python 3.x installed.
2.	Required Python libraries:
o	Flask
o	pandas
o	openpyxl
o	xlsxwriter
3.	An Excel file (sales_data.xlsx) located at C:/Users/raghu/OneDrive/Desktop/data_analysis/.
4.	data_cleaning.py module containing the functions clean_data and clean_dates.
________________________________________
Installation and Setup
1.	Install dependencies using pip:
pip install flask pandas openpyxl xlsxwriter
2.	Ensure the data_cleaning.py module is in the same directory or accessible in the Python path.
3.	Save your sales data file in the specified directory.
________________________________________
Key Functionalities
1. Flask Route
•	Endpoint: /
•	Methods: GET, POST
•	Purpose: Reads the input Excel file, cleans the data, performs analyses, and generates an output Excel file containing the results.
2. Data Cleaning
Uses the clean_data and clean_dates functions from data_cleaning.py to:
•	Remove duplicates.
•	Normalize and standardize date columns.
3. Dynamic Column Detection
Automatically identifies key columns based on partial matches (e.g., city, total, invoice ID). Raises an error if any required column is missing.
4. Analytical Functions
•	Sales Analysis by Branch: Groups data by branch to calculate total sales, transaction count, and average rating.
•	Sales Analysis by Product Line: Groups data by product line to calculate total sales, transaction count, gross income, and average rating.
•	Top 3 Product Lines: Identifies the top 3 product lines by total sales.
•	Date-Wise Sales: Aggregates total sales and transactions for each date.
•	City-Wise Sales: Analyzes total sales, transaction count, and average rating by city.
•	Customer Type Analysis: Groups data by customer type to calculate total sales and average rating.
•	Top Performing Products: Analyzes product performance by quantity sold and total sales.
•	Customer Insights: Provides insights based on gender and customer type, including total sales and average rating.
•	Sales in Date Range: Filters sales data within a specified date range.
•	Sales on Specific Day: Calculates total sales for a specific date.
•	Branch Sales on Specific Day: Analyzes branch-level sales for a given date.
5. Output Generation
•	Results are saved in an Excel file (sales_analysis_output.xlsx) with multiple sheets for each analysis.
•	If a file with the same name exists, a versioned file name is created (e.g., sales_analysis_output(1).xlsx).
________________________________________
Execution
Run the script using the following command:
python script_name.py
•	Open your browser and navigate to http://127.0.0.1:5000/.
•	The results are returned as a JSON response and saved in the output Excel file.
________________________________________
Error Handling
The script gracefully handles errors and returns a JSON response with the error message if any issue occurs during execution.
________________________________________
Key Functions and Their Purpose
Function Name	Description
analyze_sales_by_branch	Analyzes sales by branch.
analyze_sales_by_product_line	Analyzes sales by product line.
get_top_3_product_lines	Identifies top 3 product lines based on sales.
analyze_date_wise_sales	Aggregates sales data by date.
analyze_city_sales	Analyzes sales data by city.
analyze_customer_type_sales	Analyzes sales data by customer type.
analyze_top_performing_products	Analyzes top-performing products based on quantity and sales.
analyze_customer_insights	Provides customer insights by gender and type.
Filtered_data_by_date_range	Filters sales data for a specific date range.
specific_day_sale	Calculates total sales for a specific date.
Filter_branch_sales_by_specific_date	 Analyzes branch-level sales on a specific date.
________________________________________
Output Example
The output file will contain the following sheets:
1.	Overall Sales by Branches
2.	Product Line Sales
3.	Top 3 Product Lines
4.	Date-Wise Sales
5.	City-Wise Sales
6.	Customer Type
7.	Top Selling Products
8.	Customer Purchase Behavior
9.	Date Range Filtered Data
10.	Specific Day Overall Sales
11.	Branch Sales on Specific Day
________________________________________
Contact Information
For any queries or issues, contact the developer at raghul.officialmail@gmail.com

