import os
import pandas as pd

from app import app
from flask import render_template, request, redirect, url_for, send_file, send_from_directory
from flask_wtf import FlaskForm
from wtforms import RadioField
from flask_wtf.file import FileField
from werkzeug.utils import secure_filename

class UploadForm(FlaskForm):
    file = FileField()

# Upload folder where xls files will be saved and retrieved by auto-downloader  	
UPLOAD_FOLDER = '/home/sysadmin/ccwparser/files/'
ALLOWED_EXTENSIONS = set(['txt', 'xlsx'])
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER


### Importing CCW-R quote for parsing to tables
def get_tables(file):
    
    # Read csv file
    # Returns tables found in the CCW-R quote
    df = pd.read_excel(file, header=None)
    table_names = ["Invoice Summary", "Invoice Details", "Quote Details"]
    
    # Parse dataframe to split tables
    groups = df[0].isin(table_names).cumsum()
    tables = {g.iloc[0,0]:g.iloc[1:] for k,g in df.groupby(groups)}
    return tables

### Preparing tables function
def prep_tables(table_data, table_name):
    # Fix Invoice Summary table formatting. Remove NaN rows.
    df = table_data[table_name].dropna(axis=1, how='all').dropna(axis=0, how='all')

    # Remove leading and trailing space in row0
    cols = list(map(str.strip, list(df.iloc[0])))
    
    # Replace space with _ in row0
    cols = map(lambda x: x.replace(' ', '_'), cols)

    # Assigning row0 as column names
    df.columns = list(cols)

    # Removing row0
    df = df.drop(df.index[0])
    # Format '*Quantity' to integer
    df['Quantity'] = df.filter(regex='Quantity').apply(pd.to_numeric)
    return df

### Fix Invoice Summary Formatting  
def clean_inv(df, l_mod, c_mod):
    # Setup summary invoice table column names

    invoice_summary_c_names = ["Service_Ordering_SKU", "Quantity"]
    df = df[invoice_summary_c_names]    
    
    # Extracting Unit_Price from Service_Ordering_SKU  
    df['Unit_Price'] = df['Service_Ordering_SKU'].str.replace(r'\D+', '').astype('int')

    # add following columns: L modifier, C modifier, Extended Price, Extended Cost
    df['Client_Discount'] = "L"+str(l_mod)
    df['Tekscape_Discount'] = "C"+str(c_mod)
    df['Extended_Price'] = (df.Quantity * df.Unit_Price) * (1.0 - (l_mod/100))
    df['Extended_Cost'] = (df.Quantity * df.Unit_Price) * (1.0 - (c_mod/100))
    
    return df

### Fix Quote Details formatting
def clean_quo(df, l_mod, c_mod):

    # Fix Quantity/Price issue
    df['Unit_Price'] = df.Extended_Net_Price/df.Quantity

    # add following columns: L modifier, C modifier, Extended Price, Extended Cost
    df['Client_Discount'] = "L"+str(l_mod)
    df['Tekscape_Discount'] = "C"+str(c_mod)
    df['Extended_Price'] = (df.Quantity * df.Unit_Price) * (1.0 - (l_mod/100))
    df['Extended_Cost'] = (df.Quantity * df.Unit_Price) * (1.0 - (c_mod/100))
    
    # Setup quote details table column names
    quote_details_c_names = ["Instance_Number", "PAK/Serial_Number", "SKU", "Product_Number", "Quantity", \
                             "Target_Contract_Number", "Start_Date", "End_Date", "Unit_Price", "Client_Discount", \
                            "Tekscape_Discount", "Extended_Price", "Extended_Cost"]

    df = df[quote_details_c_names]
    
    return df

	
#### Function to check whether the quote is a Takeover or Incumbent quote and return correct l_mod and c_mod values.
def check_takeover(df):
    if df['Takeover_Line'].iloc[0] == 'No':
    #Incumbent values l_mod = 15, c_mod = 23
        l_mod = 15
        c_mod = 23
    else:
    #Takeover values l_mod = 10, c_mod = 15
        l_mod = 10
        c_mod = 15
    return l_mod, c_mod	

def create_quosal(file):
    ccw = get_tables(file)
    ccw_inv = prep_tables(ccw, "Invoice Summary")
    ccw_quo = prep_tables(ccw, "Quote Details")
    l_mod, c_mod = check_takeover(ccw_quo)
    inv = clean_inv(ccw_inv, l_mod, c_mod)
    quo = clean_quo(ccw_quo, l_mod, c_mod)
    
#   There will may be a discrepancy in the Invoice total amount and Quote total amount. This is because of the way parent items are priced in the quote line item details list.
#   The invoice amount total is to be followed.  

    # Writing to Excel
#    new_file = 'quosal_' + file
#   Trying to save to files folder
    new_file = app.config['UPLOAD_FOLDER'] + 'quosal_' + file  
    writer = pd.ExcelWriter(new_file, engine='xlsxwriter')
    inv.to_excel(writer, sheet_name='Invoice Line Item Detail', columns = ['Service_Ordering_SKU', 'Quantity', 'Unit_Price', 'Client_Discount', 'Tekscape_Discount'], index=False)
    quo.to_excel(writer, sheet_name='Quote Line Item Detail', columns = ['PAK/Serial_Number', 'SKU', 'Product_Number', 'Quantity', 'Target_Contract_Number', 'Start_Date', 'End_Date', 'Unit_Price', 'Client_Discount', 'Tekscape_Discount'], index=False)
    ccw_quo.to_excel(writer, sheet_name=file, index=False)
    writer.save()
    return 

@app.route('/', methods=['GET', 'POST'])
@app.route('/index')
def upload_file():
    form = UploadForm(csrf_enabled=False)
    if form.validate_on_submit():
	    filename = secure_filename(form.file.data.filename)
	    form.file.data.save(filename)
	    create_quosal(filename)
	    return send_from_directory(app.config['UPLOAD_FOLDER'], 'quosal_' + filename, as_attachment=True)
	    return redirect(url_for('upload_file', filename=filename))
    return render_template('index.html', title='CCW-R Parser', form=form)
