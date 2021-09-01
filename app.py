from flask import Flask, render_template, redirect
import requests
import json
import os
import pandas as pd
import xlrd
from flask import request
import zipfile
import shutil
from flask import send_file
import numpy as np



app = Flask(__name__)
cache= {}


@app.route("/")
@app.route("/home")
def home():
    return render_template("home.html")

@app.route("/generate", methods = ["GET", "POST"])
def generate():
    if request.method == "POST":
        sheet = request.form.getlist('sheet')
        
        
        inputFile = cache['file']

#getting sheet names
        xls = xlrd.open_workbook(inputFile, on_demand=True)
        sheet_names = xls.sheet_names()

        path = "static/"
        zipf = zipfile.ZipFile('Name.zip','w', zipfile.ZIP_DEFLATED)

        #create a new excel file for every sheet
        for name in sheet_names:
            
            if name in sheet :
                
    
                parsing = pd.ExcelFile(inputFile).parse(sheetname = name)

                #writing data to the new excel file
                df = pd.read_excel(inputFile, sheet_name=name)
                df1 = pd.read_excel(inputFile, sheet_name=name)
                col_list=list(df1.columns)
               
                for x in col_list:
                    if (df1.dtypes[x]== np.float64):
                        
                        df1[x] = df1[x].astype('object')
                        
 
                df1=df1.to_numpy() 
                row_list=df1.tolist()
                
                
                with open(path+str(name)+".txt", 'w') as outfile:
             ## From here code will change it will be specific to a sheet
             ## start if sheet==inventory:
                    c=0
                    for j in row_list:
                        
                        
                        d=0
                        for i in col_list:
                            
                            if d==0:
                                
                                outfile.write("%s" % i)
                                outfile.write(" ....")
                                outfile.write(' ')
                    
                        
                            elif d>1:
                                outfile.write("%s" % i)
                                outfile.write(' ')
                                outfile.write(str(row_list[c][d]))
                                outfile.write(' ')
                            d=d+1    
                        c=c+1        
                        outfile.write('\n')
               ## Here code will end for a speific sheet type
               ## start for new sheet type
               ## if sheet==breakout: Then follow same procedure for all sheets
                                
                            
                        
                    
                
                zipf.write(path+str(name)+".txt")
                
                print(request)
    zipf.close()
    return send_file('Name.zip',
            mimetype = 'zip',
            attachment_filename= 'Name.zip',
            as_attachment = True)


    

@app.route('/uploader', methods = ['GET', 'POST'])

def upload_file():
   if request.method == 'POST':
      f = request.files['file']
      if 'file' in cache.keys():
          
          os.remove(cache['file'])
          shutil.rmtree("static")
          os.mkdir("static")
      cache['file']=f.filename
      f.save(f.filename)
      return render_template("success.html", name = f.filename)
      

     



if __name__ == '__main__':
    os.environ['FLASK_ENV'] = 'development'
    app.run(debug=True)
    