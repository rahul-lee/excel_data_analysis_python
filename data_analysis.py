from flask import Flask, jsonify
import pandas as pd
import os
from data_cleaning import clean_data, clean_dates
from datetime import datetime

app = Flask(__name__)

@app.route('/', methods=['GET', 'POST'])
def analyze_data():
    file_path = 'C:/Users/raghu/OneDrive/Desktop/data_analysis/sales_data.xlsx'

    try:
        data = pd.read_excel(file_path)
        data.drop_duplicates(inplace = True)

        # Dynamically detect column names
        city = next((col for col in data.columns if 'city' in col.lower()), None)
        total = next((col for col in data.columns if 'total' in col.lower()), None)
        invoice_id = next((col for col in data.columns if 'invoice' in col.lower()), None)
        rating = next((col for col in data.columns if 'rating' in col.lower()), None)
        customer_type = next((col for col in data.columns if 'customer type' in col.lower()), None)  
        product_line = next((col for col in data.columns if 'product line' in col.lower()), None)
        gross_income = next((col for col in data.columns if 'gross income' in col.lower()), None)
        date_col = next((col for col in data.columns if 'date' in col.lower()), None)
        invoice_id = next((col for col in data.columns if 'invoice id' in col.lower()), None)
        branch = next((col for col in data.columns if 'branch' in col.lower()), None)
        quantity_column = next((col for col in data.columns if 'quantity' in col.lower()), None)
        total_sales_column = next((col for col in data.columns if 'total' in col.lower()), None)
        gender = next((col for col in data.columns if 'gender' in col.lower()), None)

        # Check if all required columns are found
        if not all([branch, city, total, invoice_id, rating, customer_type, product_line, gross_income, date_col]):
            raise ValueError("Missing required columns in the input data.")

        data = clean_data(data, city)
        data = clean_dates(data, date_col)

        start_date = '2019-10-31'
        specific_date = '2019-11-03'

        # Perform Analysis using separate methods
        sales_by_branch = analyze_sales_by_branch(data, branch, total, invoice_id, rating)
        product_line_sales = analyze_sales_by_product_line(data, product_line, total, invoice_id, rating, gross_income)
        top_3_product_lines = get_top_3_product_lines(data, product_line, total)
        date_wise_sales = analyze_date_wise_sales(data, date_col, total, invoice_id)
        city_sales = analyze_city_sales(data, city, total, invoice_id, rating)
        customer_type_sales = analyze_customer_type_sales(data, customer_type, total, invoice_id, rating)
        top_products = analyze_top_performing_products(data, product_line, quantity_column, total_sales_column)
        customer_insights_result = analyze_customer_insights(data, gender, customer_type, total_sales_column, rating)
        filtered_data_by_date_range = Filtered_data_by_date_range(data, start_date)
        sale_on_specific_day = specific_day_sale(data, specific_date, date_col, total_sales_column)
        branch_sale_specific_day = Filter_branch_sales_by_specific_date(data, specific_date,  date_col, branch, total, rating)

        output_file = 'sales_analysis_output.xlsx'

        if os.path.exists(output_file):
            base_name, ext = os.path.splitext(output_file)
            counter = 1
            new_output_file = f"{base_name}({counter}){ext}"
            while os.path.exists(new_output_file):  # Increment counter if file exists
                counter += 1
                new_output_file = f"{base_name}({counter}){ext}"
            output_file = new_output_file

        # Save results to Excel with multiple tabs
        with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
            sales_by_branch.to_excel(writer, sheet_name='Overall Sales by Branches', index=False)
            product_line_sales.to_excel(writer, sheet_name='Product Line Sales', index=False)
            top_3_product_lines.to_excel(writer, sheet_name='Top 3 Product Lines', index=False)
            if date_wise_sales is not None:
                date_wise_sales.to_excel(writer, sheet_name='Date-Wise Sales', index=False)
            city_sales.to_excel(writer, sheet_name='City-Wise Sales', index=False)
            customer_type_sales.to_excel(writer, sheet_name='Customer Type', index=False)
            top_products.to_excel(writer, sheet_name='Top Selling Products', index=False)
            customer_insights_result.to_excel(writer, sheet_name='Customer Purchase Behavior', index=False)
            filtered_data_by_date_range.to_excel(writer, sheet_name='Date range Filtered data', index=False)
            sale_on_specific_day.to_excel(writer, sheet_name='Specific day overall sales', index=False)
            branch_sale_specific_day.to_excel(writer, sheet_name='Branch sales on specific day', index=False)

        return jsonify({
            "status": "success",
            "message": f"Analysis saved to {output_file}"
        })

    except Exception as e:
        return jsonify({
            "status": "error",
            "message": str(e)
        })


def analyze_sales_by_branch(data, branch, total, invoice_id, rating):
    """Analyze sales by branch."""
    try:
        return data.groupby(branch).agg(
            Total_Sales=(total, 'sum'),
            Transactions=(invoice_id, 'count'),
            Avg_Rating=(rating, 'mean')
        ).reset_index()
    except Exception as e:
        print(f"An error occurred: {e}")
        return None



def analyze_sales_by_product_line(data, product_line, total, invoice_id, rating, gross_income):
    """Analyze sales by product line."""
    try:
        return data.groupby(product_line).agg(
            Total_Sales=(total, 'sum'),
            Transactions=(invoice_id, 'count'),
            Avg_Rating=(rating, 'mean'),
            Gross_Income=(gross_income, 'sum')
        ).reset_index()
    except Exception as e:
        print(f"An error occurred: {e}")
        return None



def get_top_3_product_lines(data, product_line, total):
    try:
        # Group data by Product Line and calculate Total Sales
        product_line_sales = data.groupby(product_line).agg(
            Total_Sales=(total, 'sum')
        ).reset_index()

        top_3_product_lines = product_line_sales.sort_values(by='Total_Sales', ascending=False).head(3)
        return top_3_product_lines
    
    except Exception as e:
        print(f"Error in get_top_3_product_lines: {e}")
        return None


def analyze_date_wise_sales(data, date_column, total_column, invoice_id_column):
    try:
        if not date_column:
            raise ValueError("Date column not found in the data.")

        # proper date format and remove the timestamp
        data[date_column] = pd.to_datetime(data[date_column], errors='coerce').dt.date

        date_wise_sales = data.groupby(date_column).agg(
            Total_Sales=(total_column, 'sum'),
            Transactions=(invoice_id_column, 'count')
        ).reset_index()

        # Sort by date
        date_wise_sales = date_wise_sales.sort_values(by=date_column)
        return date_wise_sales

    except Exception as e:
        print(f"Error in analyze_date_wise_sales: {e}")
        return None


def analyze_city_sales(data, city, total, invoice_id, rating):
    """Analyze sales by city."""
    try:
        return data.groupby(city).agg(
            Total_Sales=(total, 'sum'),
            Transactions=(invoice_id, 'count'),
            Avg_Rating=(rating, 'mean')
        ).reset_index()
    except Exception as e:
        print(f"An error occurred: {e}")
        return None



def analyze_customer_type_sales(data, customer_type, total, invoice_id, rating):
    """Analyze sales based on customer type."""
    try:
        return data.groupby(customer_type).agg(
            Total_Sales=(total, 'sum'),
            Transactions=(invoice_id, 'count'),
            Avg_Rating=(rating, 'mean')
        ).reset_index()
    except Exception as e:
        print(f"An error occurred: {e}")
        return None



def analyze_top_performing_products(data, product_line, quantity_column, total_sales_column):
    try:
        if not product_line or not quantity_column or not total_sales_column:
            raise ValueError("Required columns are missing in the data.")

        # Convert sales and quantity columns to numeric to avoid any data issues
        data[quantity_column] = pd.to_numeric(data[quantity_column], errors='coerce')
        data[total_sales_column] = pd.to_numeric(data[total_sales_column], errors='coerce')

        top_performing_products = data.groupby(product_line).agg(
            Total_Quantity_Sold=(quantity_column, 'sum'),
            Total_Sales=(total_sales_column, 'sum')
        ).reset_index()

        top_performing_products_sorted = top_performing_products.sort_values(
            by='Total_Quantity_Sold', ascending=False  # Sorting by Total Sales, change to 'Total_Quantity_Sold' to sort by quantity
        )
        return top_performing_products_sorted

    except Exception as e:
        print(f"Error in analyze_top_performing_products: {e}")
        return None


def analyze_customer_insights(data, gender, customer_type, total_sales, rating):
    try:
        if not gender or not customer_type or not total_sales or not rating:
            raise ValueError("Required columns are missing in the data.")
        
        data[total_sales] = pd.to_numeric(data[total_sales], errors='coerce')
        data[rating] = pd.to_numeric(data[rating], errors='coerce')

        # Group by gender and customer type to calculate total sales and average rating
        customer_insights = data.groupby([gender, customer_type]).agg(
            Total_Sales=(total_sales, 'sum'),
            Average_Rating=(rating, 'mean')
        ).reset_index()

        customer_insights_sorted = customer_insights.sort_values(by='Total_Sales', ascending=False)
        return customer_insights_sorted

    except Exception as e:
        print(f"Error in analyze_customer_insights: {e}")
        return None


def Filtered_data_by_date_range(data, start_date, end_date=None):
    try:
        date_column = next((col for col in data.columns if 'date' in col.lower()), None)
        
        if not date_column:
            raise ValueError("No date column found in the data.")

        start_date = datetime.strptime(start_date, '%Y-%m-%d').date()

        if end_date is None:
            end_date = datetime.today().date()
        else:
            end_date = datetime.strptime(end_date, '%Y-%m-%d').date()

        filtered_data = data[(data[date_column] >= start_date) & (data[date_column] <= end_date)]
        return filtered_data
    
    except Exception as e:
        print(f"Error in Filtered_data_by_date_range: {e}")
        return None


def specific_day_sale(data, start_date, date_col, total_sales_column):
    try:
        if not date_col or not total_sales_column:
            raise ValueError("Required columns (date and total) are missing in the data.")
            
        # Ensure the sales column is numeric
        data[total_sales_column] = pd.to_numeric(data[total_sales_column], errors='coerce')
        start_date = datetime.strptime(start_date, '%Y-%m-%d').date()

        filtered_data = data[data[date_col] == start_date]
        total_sales = filtered_data[total_sales_column].sum()

        result_df = pd.DataFrame({"Date": [start_date], "Total Sales": [total_sales]})
        return result_df

    except Exception as e:
        print(f"Error in Specific day sale method: {e}")
        return None


def Filter_branch_sales_by_specific_date(data, start_date,  date_col, branch, total, rating):
    try:
        
        if not date_col:
            raise ValueError("No date column found in the data.")

        start_date = datetime.strptime(start_date, '%Y-%m-%d').date()

        filtered_data = data[(data[date_col] == start_date)]
        data = filtered_data
        
        branch_sales = []

        # Iterate over unique branch names
        for branch_name in data[branch].unique():
            branch_data = data[data[branch] == branch_name]
            total_sales = branch_data[total].sum()
            avg_rating = branch_data[rating].mean()
    
            branch_sales.append({
                'Branch': branch_name,
                'Total_Sales': total_sales,
                'Current Rating': avg_rating
            })
    
        branch_sales_data = pd.DataFrame(branch_sales)
        return branch_sales_data
    
    except Exception as e:
        print(f"Error in Filtered_branch_sales_data_by_specific_date: {e}")
        return None


if __name__ == '__main__':
    app.run(debug=True)
