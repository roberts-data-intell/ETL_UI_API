# #set up and activate environment on terminal 
# -------------------------------------------------------- if data base already built pass this to next block [ ]


# on MySQL SHell

# #1 Switch to SQL Mode:
# Type 
# -----------------
# \sql 
# -----------

# #2 Connect to the MySQL Server:
# -----------------------
# \connect root@localhost:3306
# --------------------------

# #3 Create a Database called Sales_Project
# ---------------------------
# CREATE DATABASE Sales_Project;
# -----------------------------

# #4 quit
# -----------
# \quit or \q
# ---------------


# ON VS CODE
# Download MYSQL Extension by yun han
# connect to it with credentials above

# ----------------------------------------------------------------- build environment to build api block [ ]
# #
# on terminal create environment anywhere you want to begin

# 1 create env dataworld_mysql or change the name dataworld_mysql if you want (no #) 
# --------------------------
# python -m venv dataworld_mysql
# --------------------------------------


#if not working -
# Press Windows + X to open the menu 
# then select terminal (admin)
# the run this code below
# ------------------------------------
# Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Force
# ---------------------------------------

# 2 activate environment 
# ----------------------------
# .\dataworld_mysql\Scripts\activate
# -------------------------------
# correct result goes from 
# PS C:\Users\c\projects\projects_hero> 
# to  
# (dataworld_mysql) PS C:\Users\c\projects\projects_hero> now it is activated 


# 3 changing directories or folders from command line to run the code
# navigate to datawolrd mysql folder
# -------------------------------------
# cd dataworld_mysql
# --------------------------------------
# the go to the folder Scripts which is inside the environment 
# ----------------------------------
#  cd Scripts
# -------------------------------------
# this changes the directory to the Scripts folder where you want to be to run your bash shell for the api creation
    # other similar commandsto go back and forth tips inside (   )
    # go up one level (cd ..)
    # home directory (cd ~)
    # go to root directory (cd /)
    # list direcories ( ls )
    
# install
# pip install mysql-connector-python
# ==========================================================================================================================


# installs ---------------------------------------------------------------------------install block [ ]
# 1 need to download this package (wkhtmltopdf) from this site below install on pc 
# this one ( Windows    Installer (Vista or later)  64-bit  32-bit) 
# *// download https://wkhtmltopdf.org/downloads.html

# 2 pip install on command line in terminal
# pip install flask
# pip install sqalchemy
# pip install pdfkit
# pip install pandas
# pip install dicttoxml
# pip install fpdf
# pip install mysql-connector-python



# ================================================================================== code for the api block [ ]

# make sure this is saved in the Scripts folder  as name_of_api.py


import subprocess

# List of required packages in requirements.txt
required_packages = [
    "pandas",
    "dicttoxml",
    "matplotlib",
    "fpdf",
    "sqlalchemy",
    "csv",
    "pymysql"

]

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from SQLAlchemy import create_engine, text
import os
import pandas as pd
from dicttoxml import dicttoxml
from fpdf import FPDF 
import logging


# Initialize logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s [%(levelname)s]: %(message)s')

# Connect to MySQL Database with connection pooling
DATABASE_URL = "mysql+mysqlconnector://root:password@localhost/Sales_Project"
engine = create_engine(DATABASE_URL, pool_size=10, max_overflow=20)





def get_all_queries():
    # SQL Queries dictionary
    queries = {
        # Query 1 top ten products and their marketing budget and the profit
        1:  """
            SELECT `product_type`, SUM(`marketing`) as `totalmarketing`, SUM(`profit`) as `totalprofit`
                           FROM `business_data`
                            GROUP BY `product_type`
                             ORDER BY `totalmarketing` DESC
                              LIMIT 10;
                            """,

        # Note Query 2 total profit and marketing expenses
        
        2:  """
            SELECT market, SUM(profit) as totalprofit, SUM(marketing) as totalmarketing
                           FROM business_data
                            GROUP BY market;
            """,

        # Note Query 3 Correlation Between Total Expenses and Profit in Each Market Size
        
        3:  """
            SELECT `market_size`,(COUNT(*) * SUM(`profit` * `total_expenses`) - SUM(`profit`) * SUM(`total_expenses`)) / 
                           (SQRT(COUNT(*) * SUM(`profit` * `profit`) - SUM(`profit`) * SUM(`profit`)) * SQRT(COUNT(*) * SUM(`total_expenses` * `total_expenses`) - SUM(`total_expenses`) * SUM(`total_expenses`)))
                            AS Correlation
                             FROM `business_data`
                              GROUP BY `market_size` HAVING COUNT(*) * SUM(`profit` * `profit`) - SUM(`profit`) * SUM(`profit`) > 0 AND 
                               COUNT(*) * SUM(`total_expenses` * `total_expenses`) - SUM(`total_expenses`) * SUM(`total_expenses`) > 0;""",

        # Note Query 4 Product Types That Have Seen Sales Growth Over Time and how much
        4:  """
            SELECT `product_type`, (MAX(sales) - MIN(sales)) AS SalesGrowth
                            FROM (SELECT `product_type`, Date, SUM(sales) as sales
                             FROM business_data
                             GROUP BY `product_type`, Date)
                              as SubQuery7
                               GROUP BY `product_type` HAVING COUNT(sales) > 1 AND MAX(sales) > MIN(sales);
                                
                            """,

        # Note Query 5 States Where Actual COGS Are Different From 
        # Budget COGS and show by how much is the difference
        
        5:  """
            SELECT state, ABS(actual_cogs - budget_cogs) AS difference
                           FROM (SELECT state, SUM(cogs) as actual_cogs, SUM(budget_cogs) as budget_cogs
                            FROM business_data
                             GROUP BY state) 
                              as subquery6
                               WHERE actual_cogs != budget_cogs;
            """,

        # Note Query 6 Average Inventory for Each Product Type in Each State
        6:  """
            SELECT state, product_type, ROUND(AVG(inventory), 2) as avg_inventory
                           FROM business_data
                            GROUP BY state, product_type;
            """,

        # Note Query 7 Top 5 States with Highest Sales
        7:  """
            SELECT state, MAX(sales) as MaxSales
                           FROM (SELECT state, market, SUM(sales) as sales
                            FROM business_data
                             GROUP BY state, market) as sub_query_3
                              GROUP BY state
                               ORDER BY MaxSales DESC
                                LIMIT 5;
            """,

        # Note Query 8 Highest selling product in each market and what is the profit and Total Expenses
        8:  """
            SELECT market, MAX(product) as product,  
                           MAX(sales) as HighestSales,
                            SUM(profit) as total_profit,  
                             SUM(total_expenses) as total_expenses_sum  
                              FROM ( SELECT market, product,
                                SUM(sales) as sales,
                                 SUM(profit) as profit,
                                  SUM(total_expenses) as total_expenses
                                   FROM business_data
                                    GROUP BY market, product) AS sub_query
                                     GROUP BY market;
            """,

        # Note Query 9 Most Profitable Markets in Each State
        9:  """
            SELECT state, MAX(profit) as MaxProfit
                           FROM ( SELECT state, market, 
                            SUM(profit) as profit
                             FROM business_data
                              GROUP BY state, market) AS sub_query1
                               GROUP BY state;
            """,

        # Query 10 Least Profitable Markets in Each State
        10: """
            SELECT `state`, MIN(`profit`) as `MinProfit`
                        FROM (
                         SELECT `state`, `market`, SUM(`profit`) as `profit`
                          FROM `business_data`
                           GROUP BY `state`, `market`) AS sub_query1
                            GROUP BY `state`
                             ORDER BY `MinProfit` ASC;

                    """
    }

    return queries

# Descriptive names for the queries
query_descriptions = {
    1: "# 1- Most Profitable Markets in Each State",
    2: "# 2- Least Profitable Markets in Each State",
    3: "# 3- Top 5 States with Highest Sales",
    4: "# 4- Markets Exceeding Budget Profits",
    5: "# 5- Avg Inventory by Product Type and State",
    6: "# 6- States with Different Actual vs Budget COGS",
    7: "# 7- Product Types with Sales Growth Over Time",
    8: "# 8- Correlation of Expenses and Profit by Market Size",
    9: "# 9 - Markets with Low Profit but High Marketing Expenses",
    10: "# 10 - Top 10 Product Types by Marketing Expense"
}

#______________________________________________________________________________________ end of Query defintions

# ---------------------------------------------------------------------Start Function Definitions


def execute_query(query):
    with engine.connect() as connection:
        return connection.execute(text(query)).fetchall()

def get_data_as_dataframe(query_id):
    queries = get_all_queries()
    data = execute_query(queries[query_id])
    return pd.DataFrame(data)

def show_results():
    selected_description = combo_query.get()
    query_id = [key for key, value in query_descriptions.items() if value == selected_description][0]
    
    df = get_data_as_dataframe(query_id)
    for i in tree.get_children():
        tree.delete(i)
    tree["column"] = list(df.columns)
    tree["show"] = "headings"
    for column in tree["columns"]:
        tree.heading(column, text=column)
    for index, row in df.iterrows():
        tree.insert("", "end", values=list(row))


def show_full_description(event):
    widget = event.widget
    value = widget.get()
    if value in truncated_descriptions.values():
        key = [k for k, v in query_descriptions.items() if v.startswith(value)][0]
        tooltip = query_descriptions[key]
        widget["values"] = list(truncated_descriptions.values())
        widget.set(tooltip)
        app.after(2000, lambda: widget.set(value))   

# ---------------------------------------------------------FILE TYPE FUNCTIONS
# --------------------------------------------------- PDF FUNCTIONS
# --------------------------------------------------- PDF FUNCTIONS
# Function to execute query and return a DataFrame
def execute_query_pdf(query):
    with engine.connect() as conn:
        result = pd.read_sql(query, conn)
    return result

# Function to export DataFrame to PDF
def dataframe_to_pdf(df, pdf):
    # Adding column headers
    for header in df.columns:
        pdf.cell(40, 10, header, 1)
    pdf.ln()

    for i in range(0, df.shape[0]):
        for j in range(0, df.shape[1]):
            pdf.cell(40, 10, str(df.values[i,j]), 1)
        pdf.ln()

# Function to handle the export button click
def export_to_pdf():
    queries = get_all_queries()
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=12)

    # Adding header to PDF
    pdf.cell(200, 10, "RDI Data Utility", ln=True, align="C")

    for query_key, query_value in queries.items():
        df = execute_query_pdf(query_value)
        # Adding query title as a separator between queries
        pdf.cell(200, 10, query_descriptions[query_key], ln=True, align="C")
        dataframe_to_pdf(df, pdf)

    pdf_output = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
    if pdf_output:
        pdf.output(pdf_output)
        messagebox.showinfo("Success", "PDF file has been exported successfully.")



# ----------------------------------------------------------------- CSV
# Function to execute query and return a DataFrame
def execute_query_csv(query):
    with engine.connect() as conn:
        result = pd.read_sql(query, conn)
    return result

# Function to export DataFrame to CSV
def dataframe_to_csv(df, csv_path):
    df.to_csv(csv_path, index=False)

# Function to handle the export button click
def export_to_csv():
    queries = get_all_queries()
    csv_output = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")])

    if csv_output:
        with open(csv_output, 'w') as f:
            # Adding header to CSV
            f.write("RDI Data Utility\n")
            for query_key, query_value in queries.items():
                df = execute_query_csv(query_value)
                # Adding query title as a separator between queries
                f.write(f"\n{query_descriptions[query_key]}\n")
                # Including header=True to include column names
                df.to_csv(f, index=False, mode='a', header=True)
            f.close()

        messagebox.showinfo("Success", "CSV file has been exported successfully.")


        
# -------------------------------------------------------------  HTML

# Function to execute query and return a DataFrame
def execute_query_html(query):
    with engine.connect() as conn:
        result = pd.read_sql(query, conn)
    return result

# Function to export DataFrame to HTML
def dataframe_to_html(df, html_path):
    df.to_html(html_path, index=False)

# Function to handle the export button click
def export_to_html():
    queries = get_all_queries()
    html_output = filedialog.asksaveasfilename(defaultextension=".html", filetypes=[("HTML files", "*.html")])

    if html_output:
        with open(html_output, 'w') as f:
            f.write("<html>\n")
            f.write("<head><title>RDI Data Utility</title></head>\n")
            f.write("<body>\n")
            f.write("<h1>RDI Data Utility</h1>\n")

            for query_key, query_value in queries.items():
                df = execute_query_html(query_value)
                f.write(f"<h2>{query_descriptions[query_key]}</h2>\n")
                f.write(df.to_html(index=False))
                f.write("<br/>\n")

            f.write("</body>\n")
            f.write("</html>\n")

        messagebox.showinfo("Success", "HTML file has been exported successfully.")
#______________________________________________________________________________________ end of functions


# --------------------------------------------------------- function to export all formats
def export_selected_format():
    selected_format = combo_export_format.get()
    if selected_format == "PDF":
        export_to_pdf()
    elif selected_format == "CSV":
        export_to_csv()
    elif selected_format == "HTML":
        export_to_html()
    else:
        print("Invalid format selected")
           
# ----------------------------------------------------------Tkinter GUI
app = tk.Tk()
app.title("RDI Database Query Utility")



frame = ttk.Frame(app, padding="12")
frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

# Dropdown for selecting export formats
export_formats = ["PDF", "CSV", "HTML"]
combo_export_format = ttk.Combobox(frame, values=export_formats, state="readonly")
combo_export_format.grid(row=2, column=0, padx=5, pady=5)
combo_export_format.set("Select Export Format")

# Button to execute the selected export format
btn_export = ttk.Button(frame, text="Export", command=export_selected_format)
btn_export.grid(row=2, column=1, padx=5, pady=5)

# Dropdown for selecting queries
max_width = max([len(value) for value in query_descriptions.values()])
combo_query = ttk.Combobox(frame, values=list(query_descriptions.values()), state="readonly", width=max_width)

combo_query.grid(row=0, column=0, padx=5, pady=5)
combo_query.set("Select a Query")

btn_show = ttk.Button(frame, text="Show Results", command=show_results)
btn_show.grid(row=0, column=1, padx=5, pady=5)

# Treeview for displaying results
tree = ttk.Treeview(frame)
tree.grid(row=1, column=0, columnspan=2, padx=5, pady=5, sticky=(tk.W, tk.E))

# ------------ Export to PDF
export_button = ttk.Button(app, text="Export to PDF", command=export_to_pdf)
export_button.grid(row=2, column=0, padx=5, pady=50)




app.mainloop()


# ====================================================================== after words instructions
#  after code above is set up and saved in the Scripts folder name_of_api.py
# on command line run to run the code

# -----------------------------
#   python name_of_api.py 
# ------------------------------