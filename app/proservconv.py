# -*- coding: utf-8 -*-
"""
Created on Thu Dec 27 10:36:00 2018

@author: ogianan
"""

import os
import pandas as pd  
from string import punctuation  

from app import app
from flask import render_template, request, redirect, url_for, send_file, send_from_directory
from flask_wtf import FlaskForm
from wtforms import RadioField
from flask_wtf.file import FileField
from werkzeug.utils import secure_filename

class UploadForm(FlaskForm):
    file = FileField()
    
# Upload folder where xls files will be saved and retrieved by auto-downloader  	
UPLOAD_FOLDER = '/home/sysadmin/ProServConvert/files/'
ALLOWED_EXTENSIONS = set(['txt', 'csv'])
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER


def load_skus(file, subitems, subcons):
    # Create columns order list
    cols = ['Product ID', 'Quantity', 'Unit Price', 'Unit Cost']
    # Read csv file  
    soitems = pd.read_csv(file)
    soitems['Product ID'] = soitems['Product ID'][:].str.strip(punctuation).str.lstrip()
    soitems_sub = soitems[soitems['Product ID'].isin(subitems)]
    soitems_sub = soitems_sub.groupby(['Product ID', 'Unit Price', 'Unit Cost'])['Quantity'].sum().reset_index(name = 'Quantity')
    return soitems, soitems_sub[cols]





def create_proserv(file):  
    # Create tables template  
    # Create sub-items series  
    subitems = pd.Series(['PRO-UC-L1', 'PRO-NET-L1', 'PRO-SVR-L1', 'PRO-PMO-L1', 'Risk', 'PRO-SMARTHANDS-L1','PRO-SMARTHANDS'])  
    # Create SUBCON variations series  
    subcons = pd.Series(['PRO-SMARTHANDS-L1', 'PRO-SMARTHANDS'])  
    # Create output template  
    template = { 'Product ID' : ['PKG-PRO-PROFSVCS', 'PKG-PRO-PROFSVCS', 'PKG-PRO-PROFSVCS', 'PKG-PRO-PROFSVCS'] , 'Quantity' : [1, 1, 1, 1] , 'Price' : [0, 0, 0, 0] , 'Cost' : [0, 0, 0, 0] , 'Customer Description' : ['Professional Services: 50%', 'Professional Services: 20%', 'Professional Services: 20%', 'Professional Services: 10%'] }
    out_table = pd.DataFrame(template, columns=['Product ID', 'Quantity', 'Unit Price', 'Unit Cost', 'Customer Description'])
    # Create a Pandas Excel writer using XlsxWriter as the engine.
    #new_file = app.config['UPLOAD_FOLDER'] + 'proserv_' + file

    new_file = app.config['UPLOAD_FOLDER'] + 'proserv_converted.xlsx'
    
    writer = pd.ExcelWriter(new_file, engine='xlsxwriter')

    # Load soitems file  
    df, agg_df = load_skus(file, subitems, subcons)
    
    # Get total proserv price  
    t_proserv_price = df[df['Product Class'].str.contains('Bundle')]['Unit Price'].sum()  
    # Get total proserv cost
    t_proserv_cost = df[df['Product Class'].str.contains('Bundle')]['Unit Cost'].sum()  
    
    # Split smarthands from sub_agg  
    subcons_df = agg_df.loc[agg_df['Product ID'].isin(subcons)]
    subitems_df = agg_df.loc[~agg_df['Product ID'].isin(subcons)]
    
    # Get total subcons cost  
    t_subcons_cost = (subcons_df['Quantity'] * subcons_df['Unit Cost']).sum()
    
    # Get adjusted proserv cost 
    t_proserv_cost = t_proserv_cost - t_subcons_cost
    
    # Assign values to out_table  
    out_table.at[0, 'Unit Price'] = .5 * t_proserv_price
    out_table.at[1, 'Unit Price'] = .2 * t_proserv_price
    out_table.at[2, 'Unit Price'] = .2 * t_proserv_price
    out_table.at[3, 'Unit Price'] = .1 * t_proserv_price
    out_table.at[0, 'Unit Cost'] =  .5 * t_proserv_cost
    out_table.at[1, 'Unit Cost'] = .2 * t_proserv_cost
    out_table.at[2, 'Unit Cost'] = .2 * t_proserv_cost
    out_table.at[3, 'Unit Cost'] = .1 * t_proserv_cost
    
    # Zero out subcon pricing  
    #subcons_df = subcons_df.replace([subcons_df['Unit Price']], 0)
    # the above code quits when df has only one row  
    # Disable warnings
    pd.options.mode.chained_assignment = None
    subcons_df['Unit Price'] = 0
    # the above code outputs a warning  
    # rename PRO-SMARTHANDS* to SUBCON  
    subcons_df['Product ID'] = "SUBCON"  
    # the above code outputs a warning  

    # Output final table
    final = pd.concat([out_table.loc[:0], subitems_df, out_table.loc[1:], subcons_df]).reset_index(drop=True)
    final_cols = ['Product ID', 'Quantity', 'Unit Price', 'Unit Cost', 'Customer Description']
    final = final[final_cols]
    
    # Number of rows of original table
    df_rows = df.shape[0]
    final_start_row = df_rows + 3
    
    # Writing the dataframes into the worksheet  
    df.to_excel(writer, sheet_name = 'Sheet1', index = False)
    final.to_excel(writer, sheet_name = 'Sheet1', startrow=final_start_row, index = False)
    
    # Conditional formatting  
    # Get the xlsxwriter workbook and worksheet objects.
    workbook  = writer.book
    worksheet = writer.sheets['Sheet1']
    # Add a format. Light red fill with dark red text.
    red_format = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})
    
    # Apply format to the cell range.
    marker = df['Product ID'].isin(subitems) | df['Product Class'].str.contains('Bundle')
    # Get index of first TRUE value
    f_in = 1 + marker[marker==True].index.tolist()[0]
    
    # Get index of last TRUE value  
    l_in = 1 +  marker[marker==True].index.tolist()[-1]
    #str('L%d:L%d' % (f_in, l_in))
    
    for row in range (f_in, l_in+1):
    #    print(row)
        worksheet.set_row(row, 15, red_format)
    
    #print(marker)
    writer.save()
    return


@app.route('/', methods=['GET', 'POST'])
@app.route('/index')
def upload_file():
    form = UploadForm(csrf_enabled=False)
    if form.validate_on_submit():
	    filename = secure_filename(form.file.data.filename)
	    form.file.data.save(filename)
	    create_proserv(filename)
	    return send_from_directory(app.config['UPLOAD_FOLDER'], 'proserv_converted.xlsx', as_attachment=True)
	    return redirect(url_for('upload_file', filename=filename))
    return render_template('index.html', title='ProServ Conversion Tool', form=form)
