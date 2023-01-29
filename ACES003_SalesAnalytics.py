# -*- coding: UTF-8 -*-
"""
This script is to analyze sales data
Created on Wed Dec 14 13:50:00 2022
@author ACES ANALYTICS TEAM

"""

"""
Index
## 1. Create start time of running script
## 2. Import modules, packages
## 3. Define path and input file name
## 4. Import input file - sales data
## 5. Create string of date 
## 6. Define export path and output file name
## 7. Manipulate raw data
## 8. Analyze data
## 9. Plots
## 10. Format worksheets
   10.1 Format sales record worksheet
   10.2 Format analysis reports
   10.3 Add pictures into Visualization worksheet
## 11. Export results to Excel file
"""

"""
## 1. Create start time of running script
"""
from timeit import default_timer as timer
start = timer()

"""
## 2. Import modules, packages
"""
import numpy as np
import pandas as pd
import random, os, string
import matplotlib
matplotlib.use("Qt5Agg")  # Do this before importing pyplot.
import matplotlib.pyplot as plt

"""
## 3. Define import path and input file name
"""

path_from = r'D:\Python\ACES_Analytics\ACES003\Input'
input_file = path_from + "\\" + "ACES001_Sales Records.xlsx"

"""
## 4. Import input file - sales data
"""

raw_data = pd.read_excel(input_file)

"""
## 5. Create string of date
"""
from datetime import datetime
raw_date = raw_data.iat[0,1]
str_mth_year1 = datetime.strptime(raw_date,"%Y-%m-%d")
str_mth_year2 = datetime.strftime(str_mth_year1,"%d-%b-%y")
str_mth_year3 = datetime.strftime(str_mth_year1,"%b-%y")

"""
## 6. Define export path and output file name
"""
path_to = r'D:\Python\ACES_Analytics\ACES003\Output'
output_file = path_to + "\\" + "ACES004_Monthly Sales Report - " + str_mth_year3 + ".xlsx"

"""
## 7. Manipulate raw data
"""
# Calculate profit
raw_data['Profit'] = raw_data['Sales Revenue'] - raw_data['Cost of Sales']

#Calculate total volume, revenue, cost profit
ttl_volume = raw_data['Sales Volume'].sum()
ttl_revenue = raw_data['Sales Revenue'].sum()
ttl_cost = raw_data['Cost of Sales'].sum()
ttl_profit = raw_data['Profit'].sum()
ave_profit = round((ttl_profit / ttl_revenue) * 100, 4)

# Generate one series to contain total data
ttl = pd.DataFrame({"Cost of Sales": ttl_cost,
                    "Profit": ttl_profit,
                    "Sales Revenue": ttl_revenue,
                    "Sales Volume": ttl_volume,
                    "Profit%": ave_profit,
                    "Mtl Text":"Total"},index = ['Total'])

ttl2 = pd.DataFrame({"Cost of Sales": ttl_cost,
                    "Profit": ttl_profit,
                    "Sales Revenue": ttl_revenue,
                    "Sales Volume": ttl_volume,
                     "Mtl Text": "Total"
                    },index = ['Total'])

# Add ttl to raw_data
sales_records = pd.concat([raw_data,ttl2])
sales_records = sales_records.fillna(" ")

"""
## 8. Analyze data
"""

# 8.1 By sales team
sales_sltm = pd.pivot_table(raw_data,values = ['Sales Volume', 'Sales Revenue', 'Cost of Sales', 'Profit'],
                            index = ['Sales Team'], aggfunc = np.sum)

sales_sltm['Profit%'] = round(sales_sltm['Profit'] / sales_sltm['Sales Revenue'] * 100, 4)
sales_sltm2 = sales_sltm.sort_values(by="Sales Revenue",ascending= False)
sales_sltm_f = pd.concat([sales_sltm2,ttl])

sales_sltm_f['Revenue Share%'] = round(sales_sltm_f['Sales Revenue'] / ttl_revenue * 100, 2)
sales_sltm_f['Volume Share%'] = round(sales_sltm_f['Sales Volume'] / ttl_volume * 100, 2)

sales_sltm_f['Ave Price'] = round(sales_sltm_f['Sales Revenue'] / sales_sltm_f['Sales Volume'],4)

analysis_sltm = sales_sltm_f.loc[:,["Sales Volume","Sales Revenue", "Cost of Sales", "Profit",
                                    "Profit%", "Revenue Share%", "Volume Share%","Ave Price"]]

# 8.2 By salesman
sales_slmn = pd.pivot_table(raw_data,values = ['Sales Volume', 'Sales Revenue', 'Cost of Sales', 'Profit'],
                            index = ['Salesman'], aggfunc = np.sum)

sales_slmn['Profit%'] = round(sales_slmn['Profit'] / sales_slmn['Sales Revenue'] * 100, 2)
sales_slmn2 = sales_slmn.sort_values(by="Sales Revenue",ascending= False)

sales_slmn_f = pd.concat([sales_slmn2,ttl])

sales_slmn_f['Revenue Share%'] = round(sales_slmn_f['Sales Revenue'] / ttl_revenue * 100, 2)
sales_slmn_f['Volume Share%'] = round(sales_slmn_f['Sales Volume'] / ttl_volume * 100, 2)

sales_slmn_f['Ave Price'] = round(sales_slmn_f['Sales Revenue'] / sales_slmn_f['Sales Volume'],4)

analysis_slmn = sales_slmn_f.loc[:,["Sales Volume","Sales Revenue", "Cost of Sales", "Profit",
                                    "Profit%", "Revenue Share%", "Volume Share%","Ave Price"]]

# 8.3 By profit center
sales_prf = pd.pivot_table(raw_data,values = ['Sales Volume', 'Sales Revenue', 'Cost of Sales', 'Profit'],
                            index = ['Profit Center'], aggfunc = np.sum)

sales_prf['Profit%'] = round(sales_prf['Profit'] / sales_prf['Sales Revenue'] * 100, 2)
sales_prf2 = sales_prf.sort_values(by="Sales Revenue",ascending= False)

sales_prf_f = pd.concat([sales_prf2,ttl])

sales_prf_f['Revenue Share%'] = round(sales_prf_f['Sales Revenue'] / ttl_revenue * 100, 2)
sales_prf_f['Volume Share%'] = round(sales_prf_f['Sales Volume'] / ttl_volume * 100, 2)

sales_prf_f['Ave Price'] = round(sales_prf_f['Sales Revenue'] / sales_prf_f['Sales Volume'],4)
analysis_prf = sales_prf_f.loc[:,["Sales Volume","Sales Revenue", "Cost of Sales", "Profit",
                                    "Profit%", "Revenue Share%", "Volume Share%","Ave Price"]]

# 8.4 By Country
sales_cnty = pd.pivot_table(raw_data,values = ['Sales Volume', 'Sales Revenue', 'Cost of Sales', 'Profit'],
                            index = ['Country'], aggfunc = np.sum)

sales_cnty['Profit%'] = round(sales_cnty['Profit'] / sales_cnty['Sales Revenue'] * 100, 2)
sales_cnty2 = sales_cnty.sort_values(by="Sales Revenue",ascending= False)

sales_cnty_f = pd.concat([sales_cnty2,ttl])

sales_cnty_f['Revenue Share%'] = round(sales_cnty_f['Sales Revenue'] / ttl_revenue * 100, 2)
sales_cnty_f['Volume Share%'] = round(sales_cnty_f['Sales Volume'] / ttl_volume * 100, 2)

sales_cnty_f['Ave Price'] = round(sales_cnty_f['Sales Revenue'] / sales_cnty_f['Sales Volume'],4)
analysis_cnty = sales_cnty_f.loc[:,["Sales Volume","Sales Revenue", "Cost of Sales", "Profit",
                                    "Profit%", "Revenue Share%", "Volume Share%","Ave Price"]]

"""
## 9. Plots
"""

# 9.1 Pie plot for sales by sales team
# Create dataset
sltm_rev_plt = sales_sltm.loc[:,'Sales Revenue'].tolist()
sltm_plt = sales_sltm.index.tolist()

# Create explode data
cat_len = len(sales_sltm)
exp_list = [round(random.uniform(0.11,0.21),1) for i in range(cat_len)]

#exp_list
explode = (exp_list)

# Creat color parameters
colors = ("#e4dfae","#b3d966","#ff6700","#ffb300","#e4dfae", "#b3d966","#00b3a6","#807fff")

# Wedge properties
wp = { 'linewidth' : 1, 'edgecolor': "#ffb300"}

# Creat autopct arguments
def func(pct1, values1):
    num1 = int(pct1 / 100 * np.sum(values1))
    return "{:.1f}%\n {:,.0f} $".format(pct1,num1)

# Create plot
fig1, ax1 = plt.subplots(figsize = (4,4))
wedges, texts, autotexts = ax1.pie(sltm_rev_plt,
                                  autopct = lambda pct1: func(pct1, sltm_rev_plt),
                                  explode = explode,
                                  labels = sltm_plt,
                                  shadow = True,
                                  colors = colors,
                                  startangle = 90,
                                  wedgeprops =wp,
                                  textprops = dict(color = "#000000", size = 10, weight = "bold"))
# Set autotext
plt.setp(autotexts, size = 10, weight = "bold")

# Add title
ax1.set_title("By Sales Team ", size = 12, weight ="bold")

# Add text watermark
fig1.text(0.5, 0.5, 'ACES Analytics Team', fontsize = 16, color = '#2ea7db', ha = 'center', va = 'top',
          alpha = 0.5)

# Set background as transparent
fig1.patch.set_facecolor('none')
# show plot
plt.show()

# 9.2 Pie plot for sales by profit center
# Create dataset
prf_rev_plt = sales_prf.loc[:,'Sales Revenue'].tolist()
prf_plt = sales_prf.index.tolist()

# Create explode data
cat_len2 = len(sales_prf)
exp_list2 = [round(random.uniform(0.11,0.21),1) for i in range(cat_len2)]

#exp_list
explode2 = (exp_list2)

# Creat color parameters
#colors2 = ( "cyan",  "orange", "beige",  "#33d4f6",  "#ffa500", "#3ce7bc", "#E4E29A","#80c342")
colors2 = ("#e4dfae","#b3d966","#ff6700","#ffb300","#e4dfae", "#b3d966","#00b3a6","#807fff")
# Wedge properties
wp2 = { 'linewidth' : 1, 'edgecolor': "#ffb300"}

# Creat autopct arguments
def func(pct2, values2):
    num2 = int(pct2 / 100 * np.sum(values2))
    return "{:.1f}%\n {:,.0f} $".format(pct2,num2)

# Create plot
fig2, ax2 = plt.subplots(figsize = (4,4))
wedges2, texts2, autotexts2 = ax2.pie(prf_rev_plt,
                                  autopct = lambda pct2: func(pct2, prf_rev_plt),
                                  explode = explode2,
                                  labels = prf_plt,
                                  shadow = True,
                                  colors = colors2,
                                  startangle = 90,
                                  wedgeprops =wp2,
                                  textprops = dict(color = "#000000", size = 10, weight = "bold"))


# Set autotext
plt.setp(autotexts2, size = 10, weight = "bold")

# Add title
ax2.set_title("By Profit Center", size = 12, weight ="bold")

# Add text watermark
fig2.text(0.5, 0.5, 'ACES Analytics Team', fontsize = 16, color = '#2ea7db', ha = 'center', va = 'top',alpha = 0.5)
fig2.patch.set_facecolor('none')
# show plot
plt.show()

#9.3 Bar plot
#Bar plot of sales revenue by salesman
slmn_rev_plt = sales_slmn.loc[:,'Sales Revenue'].tolist()
slmn_name_plt = sales_slmn.index.tolist()

# Figure Size
fig3, ax3 = plt.subplots(figsize = (4,4))

plt.tick_params(left = True, right = False , labelleft = True ,
                labelbottom = False, bottom = True)

# Horizontal Bar Plot
ax3.barh(slmn_name_plt,slmn_rev_plt,color = '#fed1f5') ##fed1f5 - pink, #40e3bd - blue
#ax3.set_yticklabels(slmn_name_plt, fontsize=8)
plt.yticks(fontsize=8)

# Remove axes splines
for s in ['top','bottom','left','right']:
    ax3.spines[s].set_visible(False)

# Remove x,y Ticks
ax3.xaxis.set_ticks_position('none')
ax3.yaxis.set_ticks_position('none')

# Add padding between axes and labels
ax3.xaxis.set_tick_params(pad=2.0)
ax3.yaxis.set_tick_params(pad=2.0)

# Add x, y gridlines
#ax3.grid(b = 'True', color ='grey',
        #linestyle ='-', linewidth = 0.5,
        #alpha = 0.2)

# Show top values
ax3.invert_yaxis()

# Add bar label to bars
container = ax3.containers[0]
ax3.bar_label(container,labels = [f'{x:,.0f}' for x in container.datavalues], color = 'grey', fontsize = 6)

# Add Plot Title
ax3.set_title('By Sales Representative', fontsize = 12, loc = 'left',weight ="bold")

# Add Axis Labels
plt.ylabel('Name of Sale Representative', fontsize = 12, loc = 'center' )

# Add text watermark
fig3.text(0.5, 0.5, 'ACES Analytics Team', fontsize = 16, color = '#2ea7db', ha = 'center', va = 'top',alpha = 0.5)

#plt.margins(0.2)
plt.subplots_adjust(bottom=0.067)
plt.subplots_adjust(top=0.903)
plt.subplots_adjust(left=0.172)
plt.subplots_adjust(right=0.978)
plt.subplots_adjust(hspace=0.2)
plt.subplots_adjust(wspace=0.2)

#Delete 1e in x label
#current_values =plt.gca().get_xticks()
#plt.gca().set_xticklabels(['{:,.0f}'.format(x) for x in current_values])

# Set background as transparent
plt.rcParams['axes.facecolor']='none'
plt.rcParams['savefig.facecolor']='none'

plt.show()

# 9.4 Bar plot for sales by country
#Bar plot of sales revenue by country sold to
rev_plt4 = sales_cnty.loc[:,'Sales Revenue'].tolist()
cnty_plt4 = sales_cnty.index.tolist()

# Figure Size
fig4, ax4 = plt.subplots(figsize = (4,4))

plt.tick_params(left = True, right = False , labelleft = True ,
                labelbottom = False, bottom = True)

# Horizontal Bar Plot
ax4.barh(cnty_plt4,rev_plt4,color = '#70d1d4') ##fed1f5 - pink, #40e3bd - blue #50a3e4
#ax4.set_yticklabels(cnty_plt4, fontsize=8)
plt.yticks(fontsize=8)

# Remove axes splines
for s in ['top','bottom','left','right']:
    ax4.spines[s].set_visible(False)

# Remove x,y Ticks
ax4.xaxis.set_ticks_position('none')
ax4.yaxis.set_ticks_position('none')

# Add padding between axes and labels
ax4.xaxis.set_tick_params(pad=2.0)
ax4.yaxis.set_tick_params(pad=2.0)

# Add x, y gridlines
#ax4.grid(b = 'True', color ='grey',
        #linestyle ='-', linewidth = 0.5,
        #alpha = 0.2)

# Show top values
ax4.invert_yaxis()

# Add bar label to bars
container = ax4.containers[0]
ax4.bar_label(container,labels = [f'{x:,.0f}' for x in container.datavalues], color = 'grey', fontsize = 6)

# Add Plot Title
ax4.set_title('By Country Sold to', fontsize = 12, loc = 'left', weight ="bold")

# Add Axis Labels
plt.ylabel('Name of Country', fontsize = 12, loc = 'center' )

# Add text watermark
fig4.text(0.5, 0.5, 'ACES Analytics Team', fontsize = 16, color = '#fed1f5', ha = 'center', va = 'top',alpha = 0.9)

#plt.margins(0.2)
plt.subplots_adjust(bottom=0.067)
plt.subplots_adjust(top=0.903)
plt.subplots_adjust(left=0.172)
plt.subplots_adjust(right=0.978)
plt.subplots_adjust(hspace=0.2)
plt.subplots_adjust(wspace=0.2)

#Delete 1e in x label
#current_values =plt.gca().get_xticks()
#plt.gca().set_xticklabels(['{:,.0f}'.format(x) for x in current_values])

# Set background as transparent
plt.rcParams['axes.facecolor']='none'
plt.rcParams['savefig.facecolor']='none'

plt.show()

"""
## 10. Format worksheets
"""
# 10.1  Format sales record worksheet
# Open blank work book
import xlwings as xw
wb = xw.Book()

# Define name of worksheet
sheet = wb.sheets["Sheet1"]
sheet.name = "Sales Records"

# Hide gridlines for sheet
app = xw.apps.active
app.api.ActiveWindow.DisplayGridlines = False

# Assign values of Payable Records to worksheet
sheet.range("A1" ).options(index=False).value = sales_records

# define range of whole worksheet
data_rng = sheet.range("A1" ).expand('table')

# define height and width of each cell
data_rng.row_height = 23
data_rng.column_width = 13

## Format last row
# Calculate number of last row
len_row = str(len(sales_records) + 1)
last_row = "A"+ len_row

len_row_minus1 = str(len(sales_records))

# Get lenght of columns
len_col = len(sales_records.columns)

# Convert column number to column name
def excel_column_name(n):
    """Number to Excel-style column name, e.g., 1 = A, 26 = Z, 27 = AA, 703 = AAA."""
    name = ''
    while n > 0:
        n, r = divmod (n - 1, 26)
        name = chr(r + ord('A')) + name
    return name

name_last_col = excel_column_name(len_col)

last_row_e = name_last_col + len_row
last_rowminus1_e = name_last_col + len_row_minus1

last_row_rng = sheet.range(last_row, last_row_e)
for bi in range(7,11):
    last_row_rng.api.Borders(bi).Weight = 2
    last_row_rng.api.Borders(bi).Color = 0x70ad47
last_row_rng.api.Font.Name = 'Verdana'
last_row_rng.api.Font.Size = 9
last_row_rng.color = ('#c6e0b4') #6c8d84 #2da8dc #c6e0b4 green  #d8eaf1 light blue
last_row_rng.api.Font.Color = 0x000000 #000000 #0xffffff
last_row_rng.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignRight

## Format border of cells
border_rng = sheet.range("A1",last_rowminus1_e )
for bi in range(7,13):
    border_rng.api.Borders(bi).Weight = 2
    border_rng.api.Borders(bi).Color = 0x70ad47
data_rng.api.Font.Name = 'Verdana'
data_rng.api.Font.Size = 9

# Format all range of the worksheet
data_rng.api.Font.Name = 'Verdana'
data_rng.api.Font.Size = 8
data_rng.api.VerticalAlignment = xw.constants.HAlign.xlHAlignCenter
data_rng.api.WrapText = True

# Format header
header_rng = sheet.range("A1" ).expand('right')
header_rng.color = ('#70ad47') #cca989 #2da8dc
header_rng.api.Font.Color = 0xffffff
header_rng.api.Font.Bold = True
header_rng.api.Font.Size = 9
header_rng.api.VerticalAlignment = xw.constants.HAlign.xlHAlignCenter
header_rng.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignCenter

# Format first column
id_column_rng = sheet.range("A2").expand('down')
id_column_rng.color = ('#c6e0b4') #6c8d84 #2da8dc #c6e0b4 green  #d8eaf1 light blue
id_column_rng.api.Font.Color = 0x000000 #000000 #0xffffff
id_column_rng.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignLeft
id_column_rng.column_width = 9

# Format B column - Date
B_account_rng = sheet.range("B2").expand('down')
B_account_rng.column_width = 10
B_account_rng.api.VerticalAlignment = xw.constants.HAlign.xlHAlignCenter
B_account_rng.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignLeft

# Format C column - Date
C_account_rng = sheet.range("C2").expand('down')
C_account_rng.column_width = 4

# Format D column - Customer
D_account_rng = sheet.range("D2").expand('down')
D_account_rng.column_width = 20

# Format G column - Sales Team
G_account_rng = sheet.range("G2").expand('down')
G_account_rng.column_width = 10

# Format H column - Profit Center
H_account_rng = sheet.range("H2").expand('down')
H_account_rng.column_width = 10

# Format I column - Material Group
I_account_rng = sheet.range("I2").expand('down')
I_account_rng.column_width = 8

# Format J column - Material Group
J_account_rng = sheet.range("J2").expand('down')
J_account_rng.column_width = 8
J_account_rng.api.VerticalAlignment = xw.constants.HAlign.xlHAlignCenter
J_account_rng.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignLeft

# Format K column - Material Text
K_account_rng = sheet.range("K2").expand('down')
K_account_rng.column_width = 15

# Format L column - Sales Volume
l_rng = sheet.range("l2").expand('down')
l_rng.column_width = 10
l_rng.number_format = "#,###"

# Format M column - Unit Price
M_rng = sheet.range("M2").expand('down')
M_rng.column_width = 8
M_rng.number_format = "#,###.0000"

# Format N column - Unit Price
N_rng = sheet.range("N2").expand('down')
N_rng.column_width = 8
N_rng.number_format = "#,###.0000"

# Format OPQ columns - Sales Revenue, Cost of Sales, Profit
OPQ_rng = sheet.range("O2").expand('table')
OPQ_rng.column_width = 13
OPQ_rng.number_format = "#,###.00"

# Format first row
firstrow_rng = sheet.range("A1" ).expand('right')
firstrow_rng.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignCenter



# Add title of workseet of sales records
# Get rows of sales records
rowl = "{:,d}".format(len(sales_records))

# Insert three rows above row 1 in worksheet of Sales Records
sheet.range("1:3").insert('down')
sheet.range("A1" ).value = "Sales Records"
sheet.range("A2").value = "Length of Rows: " + rowl
sheet.range("A3").value = "Unit: USD " + " " + "UoM:kg"

# 10.2  Format analysis reports
# 10.2.1 Add new sheet
sheet2 = wb.sheets.add("Analysis")

# Hide gridlines for sheet2
app = xw.apps.active
app.api.ActiveWindow.DisplayGridlines = False

# 10.2.2 sales analysis by sales team
#Locate rows
lst_row2 =  str(len(analysis_sltm) + 1)
fst_row2_s = "A1"
fst_row2_s_plus1 = "A2"

# Assign values 

sheet2.range("A1" ).options(index=True).value = analysis_sltm
sheet2.range("A1" ).value = "Sales Team"

# Calculate number of last row
last_row2 = "A"+ lst_row2
last_row_e2 = "I"+ lst_row2

# define range of whole worksheet
data_rng2 = sheet2.range("A1").expand('table')

# define height and width of each cell
data_rng2.row_height = 23
data_rng2.column_width = 9
data_rng2.api.Font.Name = 'Verdana'
data_rng2.api.Font.Size = 9

# Format border of cells
border_rng2 = sheet2.range("A1" ).expand('table')
for bi in range(7,13):
    border_rng2.api.Borders(bi).Weight = 2
    border_rng2.api.Borders(bi).Color = 0x70ad47

# Format first row
border_rng2 = sheet2.range("A1" ).expand('right')
for bi in range(7,13):
    border_rng2.api.Borders(bi).Weight = 2
    border_rng2.api.Borders(bi).Color = 0x70ad47

# Format all range of the worksheet
data_rng2.api.Font.Name = 'Verdana'
data_rng2.api.Font.Size = 8
data_rng2.api.VerticalAlignment = xw.constants.HAlign.xlHAlignCenter
data_rng2.api.WrapText = True

# Format header
header_rng2 = sheet2.range("A1" ).expand('right')
header_rng2.color = ('#70ad47') #cca989 #2da8dc
header_rng2.api.Font.Color = 0xffffff
header_rng2.api.Font.Bold = True
header_rng2.api.Font.Size = 9
header_rng2.api.VerticalAlignment = xw.constants.HAlign.xlHAlignCenter
header_rng2.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignCenter
header_rng2.api.WrapText = True

# Format first column
id_column_rng2 = sheet2.range("A2").expand('down')
id_column_rng2.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignLeft
#id_column_rng2.color = ('#c6e0b4') #6c8d84 #2da8dc #c6e0b4 green  #d8eaf1 light blue
#id_column_rng2.api.Font.Color = 0x000000 #000000 #0xffffff

#id_column_rng2.column_width = 13
#id_column_rng2.api.WrapText = True
#id_column_rng2.api.VerticalAlignment = xw.constants.HAlign.xlHAlignCenter

# Format last row
bottom_rng2 = sheet2.range(last_row2, last_row_e2)
for bi in range(7,13):
    bottom_rng2.api.Borders(bi).Weight = 2
    bottom_rng2.api.Borders(bi).Color = 0x70ad47
bottom_rng2.api.Font.Name = 'Verdana'
bottom_rng2.api.Font.Size = 9
bottom_rng2.color = ('#c6e0b4') #6c8d84 #2da8dc #c6e0b4 green  #d8eaf1 light blue
bottom_rng2.api.Font.Color = 0x000000 #000000 #0xffffff
bottom_rng2.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignRight

# Format BCD columns - Sales Volume, Sales Revenue, Cost of Sales, Profit
B_rng2 = sheet2.range("B2").expand('down')
B_rng2.number_format = "#,###"
B_rng2.column_width = 12

# Format BCD columns - Sales Volume, Sales Revenue, Cost of Sales, Profit
C_rng2 = sheet2.range("C2").expand('down')
C_rng2.number_format = "#,###"
C_rng2.column_width = 12

# Format BCD columns - Sales Volume, Sales Revenue, Cost of Sales, Profit
D_rng2 = sheet2.range("D2").expand('down')
D_rng2.number_format = "#,###"
D_rng2.column_width = 12

# Format E columns - Profit
E_rng2 = sheet2.range("E2").expand('down')
E_rng2.number_format = "#,###"
E_rng2.column_width = 12

# Format F columns - Profit%
F_rng2 = sheet2.range("F2").expand('down')
F_rng2.number_format = "#,###.0"
F_rng2.column_width = 8

# Format G Revenue Share
G_rng2 = sheet2.range("G2").expand('down')
G_rng2.number_format = "#,###.0"
G_rng2.column_width = 8

# Format H Revenue Share
H_rng2 = sheet2.range("H2").expand('down')
H_rng2.number_format = "#,###.0"
H_rng2.column_width = 8

# Format I Revenue Share
I_rng2 = sheet2.range("I2").expand('down')
I_rng2.number_format = "#,###.0000"
I_rng2.column_width = 8

# 10.2.3 sales analysis by profit center
# Locate rows
fst_row3 = str(len(analysis_sltm) + 3)
lst_row3 =  str(len(analysis_sltm) + len(analysis_prf) + 3)
fst_row3_s = "A" + fst_row3
fst_row3_s_plus1 = "A" + str(str(len(analysis_sltm) + 4))

# Assign values of sales analysis by profit center to worksheet
#sheet2.range(fst_row3_s ).value = "By Profit Center"
sheet2.range(fst_row3_s ).options(index=True).value = analysis_prf
sheet2.range(fst_row3_s ).value = "Profit Center"

# define range of whole worksheet
data_rng3 = sheet2.range(fst_row3_s ).expand('table')

# define height and width of each cell
data_rng3.row_height = 23
data_rng3.column_width = 9
data_rng3.api.Font.Name = 'Verdana'
data_rng3.api.Font.Size = 9

# Format border of cells
border_rng3 = sheet2.range(fst_row3_s ).expand('table')
for bi in range(7,13):
    border_rng3.api.Borders(bi).Weight = 2
    border_rng3.api.Borders(bi).Color = 0x70ad47

# Format first row
border_rng3 = sheet2.range(fst_row3_s ).expand('right')
for bi in range(7,13):
    border_rng2.api.Borders(bi).Weight = 2
    border_rng2.api.Borders(bi).Color = 0x70ad47

# Format all range of the worksheet
data_rng3.api.Font.Name = 'Verdana'
data_rng3.api.Font.Size = 8
data_rng3.api.VerticalAlignment = xw.constants.HAlign.xlHAlignCenter
data_rng3.api.WrapText = True

# Format header
header_rng3 = sheet2.range(fst_row3_s ).expand('right')
header_rng3.color = ('#70ad47') #cca989 #2da8dc
header_rng3.api.Font.Color = 0xffffff
header_rng3.api.Font.Bold = True
header_rng3.api.Font.Size = 9
header_rng3.api.VerticalAlignment = xw.constants.HAlign.xlHAlignCenter
header_rng3.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignCenter
header_rng3.api.WrapText = True

# Format first column
id_column_rng3 = sheet2.range(fst_row3_s_plus1).expand('down')
id_column_rng3.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignLeft
#id_column_rng2.color = ('#c6e0b4') #6c8d84 #2da8dc #c6e0b4 green  #d8eaf1 light blue
#id_column_rng2.api.Font.Color = 0x000000 #000000 #0xffffff

#id_column_rng2.column_width = 13
#id_column_rng2.api.WrapText = True
#id_column_rng2.api.VerticalAlignment = xw.constants.HAlign.xlHAlignCenter

# Calculate number of last row
last_row3 = "A"+ lst_row3
last_row_e3 = "I"+ lst_row3

# Format last row
bottom_rng3 = sheet2.range(last_row3, last_row_e3)
for bi in range(7,13):
    bottom_rng3.api.Borders(bi).Weight = 2
    bottom_rng3.api.Borders(bi).Color = 0x70ad47
bottom_rng3.api.Font.Name = 'Verdana'
bottom_rng3.api.Font.Size = 9
bottom_rng3.color = ('#c6e0b4') #6c8d84 #2da8dc #c6e0b4 green  #d8eaf1 light blue
bottom_rng3.api.Font.Color = 0x000000 #000000 #0xffffff
bottom_rng3.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignRight

# Format BCD columns - Sales Volume, Sales Revenue, Cost of Sales, Profit
B_clm_3 = "B" + str((len(analysis_sltm) + 4))
B_rng3 = sheet2.range(B_clm_3).expand('down')
B_rng3.number_format = "#,###"
B_rng3.column_width = 12

# Format BCD columns - Sales Volume, Sales Revenue, Cost of Sales, Profit
C_clm_3 = "C" + str((len(analysis_sltm) + 4))
C_rng3 = sheet2.range(C_clm_3).expand('down')
C_rng3.number_format = "#,###"
C_rng3.column_width = 12

# Format BCD columns - Sales Volume, Sales Revenue, Cost of Sales, Profit
D_clm_3 = "D" + str((len(analysis_sltm) + 4))
D_rng3 = sheet2.range(D_clm_3).expand('down')
D_rng3.number_format = "#,###"
D_rng3.column_width = 12

# Format E columns - Profit
E_clm_3 = "E" + str((len(analysis_sltm) + 4))
E_rng3 = sheet2.range(E_clm_3).expand('down')
E_rng3.number_format = "#,###"
E_rng3.column_width = 12

# Format F columns - Profit%
F_clm_3 = "F" + str((len(analysis_sltm) + 4))
F_rng3 = sheet2.range(F_clm_3 ).expand('down')
F_rng3.number_format = "#,###.0"
F_rng3.column_width = 8

# Format G Revenue Share
G_clm_3 = "G" + str((len(analysis_sltm) + 4))
G_rng3 = sheet2.range(G_clm_3).expand('down')
G_rng3.number_format = "#,###.0"
G_rng3.column_width = 8

# Format H Revenue Share
H_clm_3 = "H" + str((len(analysis_sltm) + 4))
H_rng3 = sheet2.range(H_clm_3).expand('down')
H_rng3.number_format = "#,###.0"
H_rng3.column_width = 8

# Format I Revenue Share
I_clm_3 = "I" + str((len(analysis_sltm) + 4))
I_rng3 = sheet2.range(I_clm_3).expand('down')
I_rng3.number_format = "#,###.0000"
I_rng3.column_width = 8

# Calculate number of last row
last_row3 = "A"+ lst_row3
last_row_e3 = "I"+ lst_row3

# Format last row
bottom_rng3 = sheet2.range(last_row3, last_row_e3)
for bi in range(7,11):
    bottom_rng3.api.Borders(bi).Weight = 2
    bottom_rng3.api.Borders(bi).Color = 0x70ad47
bottom_rng3.api.Font.Name = 'Verdana'
bottom_rng3.api.Font.Size = 9
bottom_rng3.color = ('#c6e0b4') #6c8d84 #2da8dc #c6e0b4 green  #d8eaf1 light blue
bottom_rng3.api.Font.Color = 0x000000 #000000 #0xffffff
bottom_rng3.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignRight

#10.2.4 Assign values of sales analysis by salesman to worksheet
# Locate rows
fst_row4 = str(len(analysis_sltm) + len(analysis_prf) + 5)
lst_row4 =  str(len(analysis_sltm) + len(analysis_prf) + len (analysis_slmn) + 5)
fst_row4_s = "A" + fst_row4
fst_row4_s_plus1 = "A" + str(str(len(analysis_sltm) + len(analysis_prf) + 6))

# Assign values of sales analysis by profit center to worksheet
sheet2.range(fst_row4_s).options(index=True).value = analysis_slmn
sheet2.range(fst_row4_s).value = "Sale Representative"

# Calculate number of last row
last_row4 = "A"+ lst_row4
last_row_e4 = "I"+ lst_row4

# define range of whole worksheet
data_rng4 = sheet2.range(fst_row4_s).expand('table')

# define height and width of each cell
data_rng4.row_height = 23
data_rng4.column_width = 9
data_rng4.api.Font.Name = 'Verdana'
data_rng4.api.Font.Size = 9

# Format border of cells
border_rng4 = sheet2.range(fst_row4_s).expand('table')
for bi in range(7,13):
    border_rng4.api.Borders(bi).Weight = 2
    border_rng4.api.Borders(bi).Color = 0x70ad47

# Format first row
border_rng4 = sheet2.range(fst_row4_s).expand('right')
for bi in range(7,13):
    border_rng4.api.Borders(bi).Weight = 2
    border_rng4.api.Borders(bi).Color = 0x70ad47

# Format all range of the worksheet
data_rng4.api.Font.Name = 'Verdana'
data_rng4.api.Font.Size = 8
data_rng4.api.VerticalAlignment = xw.constants.HAlign.xlHAlignCenter
data_rng3.api.WrapText = True

# Format header
header_rng4 = sheet2.range(fst_row4_s).expand('right')
header_rng4.color = ('#70ad47') #cca989 #2da8dc
header_rng4.api.Font.Color = 0xffffff
header_rng4.api.Font.Bold = True
header_rng4.api.Font.Size = 9
header_rng4.api.VerticalAlignment = xw.constants.HAlign.xlHAlignCenter
header_rng4.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignCenter
header_rng4.api.WrapText = True

# Format first column
id_column_rng4 = sheet2.range(fst_row4_s_plus1).expand('down')
id_column_rng4.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignLeft
#id_column_rng2.color = ('#c6e0b4') #6c8d84 #2da8dc #c6e0b4 green  #d8eaf1 light blue
#id_column_rng2.api.Font.Color = 0x000000 #000000 #0xffffff

id_column_rng2.column_width = 18
#id_column_rng2.api.WrapText = True
#id_column_rng2.api.VerticalAlignment = xw.constants.HAlign.xlHAlignCenter

# Format last row
bottom_rng4 = sheet2.range(last_row4, last_row_e4)
for bi in range(7,13):
    bottom_rng4.api.Borders(bi).Weight = 2
    bottom_rng4.api.Borders(bi).Color = 0x70ad47
bottom_rng4.api.Font.Name = 'Verdana'
bottom_rng4.api.Font.Size = 9
bottom_rng4.color = ('#c6e0b4') #6c8d84 #2da8dc #c6e0b4 green  #d8eaf1 light blue
bottom_rng4.api.Font.Color = 0x000000 #000000 #0xffffff
bottom_rng4.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignRight

# Format BCD columns - Sales Volume, Sales Revenue, Cost of Sales, Profit
B_clm_4 = "B" + str((len(analysis_sltm) + len(analysis_prf) + 6))
B_rng4 = sheet2.range(B_clm_4).expand('down')
B_rng4.number_format = "#,###"
B_rng4.column_width = 12

# Format BCD columns - Sales Volume, Sales Revenue, Cost of Sales, Profit
C_clm_4 = "C" + str((len(analysis_sltm) + len(analysis_prf) + 6))
C_rng4 = sheet2.range(C_clm_4).expand('down')
C_rng4.number_format = "#,###"
C_rng4.column_width = 12

# Format BCD columns - Sales Volume, Sales Revenue, Cost of Sales, Profit
D_clm_4 = "D" + str((len(analysis_sltm) + len(analysis_prf) + 6))
D_rng4 = sheet2.range(D_clm_4).expand('down')
D_rng4.number_format = "#,###"
D_rng4.column_width = 12

# Format E columns - Profit
E_clm_4 = "E" + str((len(analysis_sltm) + len(analysis_prf) + 6))
E_rng4 = sheet2.range(E_clm_4).expand('down')
E_rng4.number_format = "#,###"
E_rng4.column_width = 12

# Format F columns - Profit%
F_clm_4 = "F" + str((len(analysis_sltm) + len(analysis_prf) + 6))
F_rng4 = sheet2.range(F_clm_4).expand('down')
F_rng4.number_format = "#,###.0"
F_rng4.column_width = 8

# Format G Revenue Share
G_clm_4 = "G" + str((len(analysis_sltm) + len(analysis_prf) + 6))
G_rng4 = sheet2.range(G_clm_4).expand('down')
G_rng4.number_format = "#,###.0"
G_rng4.column_width = 8

# Format H Revenue Share
H_clm_4 = "H" + str((len(analysis_sltm) + len(analysis_prf) + 6))
H_rng4 = sheet2.range("H15").expand('down')
H_rng4.number_format = "#,###.0"
H_rng4.column_width = 8

# Format I Revenue Share
I_clm_4 = "I" + str((len(analysis_sltm) + len(analysis_prf) + 6))
I_rng4 = sheet2.range(I_clm_4).expand('down')
I_rng4.number_format = "#,###.0000"
I_rng4.column_width = 8

# Format last row
bottom_rng4 = sheet2.range(last_row4, last_row_e4)
for bi in range(7,11):
    bottom_rng4.api.Borders(bi).Weight = 2
    bottom_rng4.api.Borders(bi).Color = 0x70ad47
bottom_rng4.api.Font.Name = 'Verdana'
bottom_rng4.api.Font.Size = 9
bottom_rng4.color = ('#c6e0b4') #6c8d84 #2da8dc #c6e0b4 green  #d8eaf1 light blue
bottom_rng4.api.Font.Color = 0x000000 #000000 #0xffffff
bottom_rng4.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignRight

#10.2.5 Assign values of sales analysis by country to worksheet

# Locate first row and last row of analysis_cnty
fst_row5 = str(len(analysis_sltm) + len(analysis_prf) + len (analysis_slmn) + 7)
lst_row5 =  str(len(analysis_cnty) + len(analysis_sltm) + len(analysis_prf) + len (analysis_slmn) + 7)

fst_row5_s = "A" + fst_row5
fst_row5_s_plus1 = "A" + str(str(len(analysis_sltm) + len(analysis_prf) + len (analysis_slmn) + 8))

# Assign values of sales analysis by profit center to worksheet
sheet2.range(fst_row5_s).options(index=True).value = analysis_cnty
sheet2.range(fst_row5_s).value = "Country Sold to"

# Calculate number of last row
last_row5 = "A" + lst_row5
last_row_e5 = "I"+ lst_row5

# define range of whole worksheet
data_rng5 = sheet2.range(fst_row5_s).expand('table')

# define height and width of each cell
data_rng5.row_height = 23
data_rng5.column_width = 9
data_rng5.api.Font.Name = 'Verdana'
data_rng5.api.Font.Size = 9

# Format border of cells
border_rng5 = sheet2.range(fst_row5_s).expand('table')
for bi in range(7,13):
    border_rng5.api.Borders(bi).Weight = 2
    border_rng5.api.Borders(bi).Color = 0x70ad47

# Format first row
border_rng5 = sheet2.range(fst_row5_s).expand('right')
for bi in range(7,13):
    border_rng5.api.Borders(bi).Weight = 2
    border_rng5.api.Borders(bi).Color = 0x70ad47

# Format all range of the worksheet
data_rng5.api.Font.Name = 'Verdana'
data_rng5.api.Font.Size = 8
data_rng5.api.VerticalAlignment = xw.constants.HAlign.xlHAlignCenter
data_rng5.api.WrapText = True

# Format header
header_rng5 = sheet2.range(fst_row5_s).expand('right')
header_rng5.color = ('#70ad47') #cca989 #2da8dc
header_rng5.api.Font.Color = 0xffffff
header_rng5.api.Font.Bold = True
header_rng5.api.Font.Size = 9
header_rng5.api.VerticalAlignment = xw.constants.HAlign.xlHAlignCenter
header_rng5.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignCenter
header_rng5.api.WrapText = True

# Format first column
id_column_rng5 = sheet2.range(fst_row5_s_plus1).expand('down')
id_column_rng5.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignLeft
#id_column_rng2.color = ('#c6e0b4') #6c8d84 #2da8dc #c6e0b4 green  #d8eaf1 light blue
#id_column_rng2.api.Font.Color = 0x000000 #000000 #0xffffff

id_column_rng5.column_width = 18
#id_column_rng2.api.WrapText = True
#id_column_rng2.api.VerticalAlignment = xw.constants.HAlign.xlHAlignCenter

# Format last row
bottom_rng5 = sheet2.range(last_row5, last_row_e5)
for bi in range(7,13):
    bottom_rng5.api.Borders(bi).Weight = 2
    bottom_rng5.api.Borders(bi).Color = 0x70ad47
bottom_rng5.api.Font.Name = 'Verdana'
bottom_rng5.api.Font.Size = 9
bottom_rng5.color = ('#c6e0b4') #6c8d84 #2da8dc #c6e0b4 green  #d8eaf1 light blue
bottom_rng5.api.Font.Color = 0x000000 #000000 #0xffffff
bottom_rng5.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignRight

# Format BCD columns - Sales Volume, Sales Revenue, Cost of Sales, Profit
B_clm_5 = "B" + str(str(len(analysis_sltm) + len(analysis_prf) + len (analysis_slmn) + 8))
B_rng5 = sheet2.range(B_clm_5).expand('down')
B_rng5.number_format = "#,###"
B_rng5.column_width = 12

# Format BCD columns - Sales Volume, Sales Revenue, Cost of Sales, Profit
C_clm_5 = "C" + str(str(len(analysis_sltm) + len(analysis_prf) + len (analysis_slmn) + 8))
C_rng5 = sheet2.range(C_clm_5).expand('down')
C_rng5.number_format = "#,###"
C_rng5.column_width = 12

# Format BCD columns - Sales Volume, Sales Revenue, Cost of Sales, Profit
D_clm_5 = "D" + str(str(len(analysis_sltm) + len(analysis_prf) + len (analysis_slmn) + 8))
D_rng5 = sheet2.range(D_clm_5).expand('down')
D_rng5.number_format = "#,###"
D_rng5.column_width = 12

# Format E columns - Profit
E_clm_5 = "E" + str(str(len(analysis_sltm) + len(analysis_prf) + len (analysis_slmn) + 8))
E_rng5 = sheet2.range(E_clm_5).expand('down')
E_rng5.number_format = "#,###"
E_rng5.column_width = 12

# Format F columns - Profit%
F_clm_5 = "F" + str(str(len(analysis_sltm) + len(analysis_prf) + len (analysis_slmn) + 8))
F_rng5 = sheet2.range(F_clm_5).expand('down')
F_rng5.number_format = "#,###.0"
F_rng5.column_width = 8

# Format G Revenue Share
G_clm_5 = "G" + str(str(len(analysis_sltm) + len(analysis_prf) + len (analysis_slmn) + 8))
G_rng5 = sheet2.range(G_clm_5).expand('down')
G_rng5.number_format = "#,###.0"
G_rng5.column_width = 8

# Format H Revenue Share
H_clm_5 = "H" + str(str(len(analysis_sltm) + len(analysis_prf) + len (analysis_slmn) + 8))
H_rng5 = sheet2.range(H_clm_5).expand('down')
H_rng5.number_format = "#,###.0"
H_rng5.column_width = 8

# Format I Revenue Share
I_clm_5 = "I" + str(str(len(analysis_sltm) + len(analysis_prf) + len (analysis_slmn) + 8))
I_rng5 = sheet2.range(I_clm_5 ).expand('down')
I_rng5.number_format = "#,###.0000"
I_rng5.column_width = 8

# Format last row
bottom_rng5 = sheet2.range(last_row5, last_row_e5)
for bi in range(7,11):
    bottom_rng5.api.Borders(bi).Weight = 2
    bottom_rng5.api.Borders(bi).Color = 0x70ad47
bottom_rng5.api.Font.Name = 'Verdana'
bottom_rng5.api.Font.Size = 9
bottom_rng5.color = ('#c6e0b4') #6c8d84 #2da8dc #c6e0b4 green  #d8eaf1 light blue
bottom_rng5.api.Font.Color = 0x000000 #000000 #0xffffff
bottom_rng5.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignRight

# Add Title
# Insert three rows above row 1 in worksheet of Sales Records
sheet2.range("1:2").insert('down')
sheet2.range("A1" ).value = "Monthly Sales Analysis" + "  " + "in" + "  " + str_mth_year3
sheet2.range("A2").value = "Unit: USD " + " " + "UoM:kg"

#10.3 Visualization
# 10.3.1 Add new sheet
sheet3 = wb.sheets.add("Visualization")
# Hide gridlines for sheet2
app = xw.apps.active
app.api.ActiveWindow.DisplayGridlines = False

sheet3.pictures.add(fig1, name='ACES_Plot1', update=True,
                     left=sheet2.range('A1').left, top=sheet2.range('A1').top)

sheet3.pictures.add(fig4, name='ACES_Plot2', update=True,
                     left=sheet2.range('G1').left, top=sheet2.range('G1').top)

sheet3.pictures.add(fig3, name='ACES_Plot3', update=True,
                     left=sheet2.range('A15').left, top=sheet2.range('A15').top)

sheet3.pictures.add(fig2, name='ACES_Plot4', update=True,
                     left=sheet2.range('H15').left, top=sheet2.range('H15').top)

# Get time
from time import strftime, localtime
time_local = strftime ("%A, %d %b %Y, %H:%M")

# Insert four rows above row 1 of worksheet of visualization
sheet3.range("1:4").insert('down')
sheet3.range("A1" ).value = "Monthly Sales Report" + "  " + "in" + "  " + str_mth_year3
sheet3.range("A2").value = "The last update time :  " + time_local  +"."

# Calculate running time to be shown in output file
end = timer()
running_time = "{:,.2f}".format(end - start)
sheet3.range("A3").value = "Running time:  " +  running_time + "  s"

"""
## 11. Export results to Excel file
"""
wb.save(output_file)
wb.close()
app.quit()

# Print the end for this script
print("The run of script is completed successfully.")
time_local_end = strftime ("%A, %d %b %Y, %H:%M")

#Print running time
running_time2 = "{:,.2f}".format(end - start)
print (running_time2 )
