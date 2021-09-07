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
from flask import * 
from werkzeug.utils import secure_filename
import openpyxl



app = Flask(__name__)
app.config['UPLOAD_EXTENSIONS'] = ['.xls', '.xlsx']
cache= {}

    
def createcommand(command, key, value):
    
    
        
    if (value!='dummy'):
        
        
        if isinstance(value,list):
            value_tmp=''
            for i in value:
                
                value_tmp=str(value_tmp + " --" + key + ' ' + str(i))
            return str(command + value_tmp)
        elif isinstance(value,float):
            
            value = str(int(value))
            return str(command + " --" + key + ' ' + value )
        else:
            value=str(value)
            return str(command + " --" + key + ' ' + value )
    else:
        
        return command


@app.route("/")
@app.route("/home")
def home():
    return render_template("home.html")

@app.route("/generate", methods = ["GET", "POST"])
def generate():
    
    if cache == {}:
        
        
        return render_template("upload_first.html") 
    if request.method == "POST":
        sheet = request.form.getlist('sheet')
        
        
        inputFile = cache['file']

#getting sheet names
        xls = xlrd.open_workbook(inputFile, on_demand=True)
        sheet_names = xls.sheet_names()

        path = "static/"
        zipf = zipfile.ZipFile('Name.zip','w', zipfile.ZIP_DEFLATED)

        #create a new excel file for every sheet
        for name1 in sheet_names:
            
            
            if name1 in sheet :
                
    
                

                #writing data to the new excel file
                ## From here code will change it will be specific to a sheet
                if name1 == "Tenant":#########################################
                    
                    
                    df = pd.read_excel(inputFile, sheet_name=name1, engine='openpyxl')
                    df1 = pd.read_excel(inputFile, sheet_name=name1, engine='openpyxl')
                    col_list=list(df1.columns)
                    df = df.replace(np.nan, 'dummy')
                    
                    with open(path+str(name1)+".txt", 'w') as outfile:
                        
                        tenant = []
                        for k in range(0,len(df['efa tenant create'])):
                            
                            
    
                            tenant_temp = {}
                            tenant_temp['name']=df['name'][k]
                            tenant_temp['description']=df['description'][k]
                            tenant_temp['type']=df['type'][k]
                            tenant_temp['l2-vni-range']=df['l2-vni-range'][k]
                            tenant_temp['l3-vni-range']=df['l3-vni-range'][k]
                            tenant_temp['vlan-range']=df['vlan-range'][k]
                            tenant_temp['vrf-count']=df['vrf-count'][k]
                            tenant_temp['enable-bd']=str(df['enable-bd'][k]).lower
                            tenant_temp['port_sw']=df['sw-ip'][k]+"["+df['sw-port'][k]+"]"
                            tenant.append(tenant_temp)
	                       
	                       
                        
               
                        
               
			         
                        lst1 = []
                        for i in tenant:
                            
                            if (i['name'] !="dummy"):
                                
                                command = "efa tenant create"
                                name = createcommand(command,'name',i['name'])
                                description = createcommand(name,'description',i['description'])
                                type = createcommand(description,'type',i['type'])
                                l2_vni_range = createcommand(type,'l2-vni-range',i['l2-vni-range'])
                                l3_vni_range = createcommand(l2_vni_range,'l3-vni-range',i['l3-vni-range'])
                                vlan_range = createcommand(l3_vni_range,'vlan-range',i['vlan-range'])
                                vrf_count = createcommand(vlan_range,'vrf-count',i['vrf-count'])
                                enable_bd = createcommand(vrf_count,'enable-bd',i['enable-bd'])
                                port = createcommand(enable_bd,'port',i['port_sw'])
                                lst1.append(port)
        
                            else:
                                
                                port = "," + i['port_sw']
                                last_element = lst1[-1]
                                lst1.pop()
                                tmp = last_element + port
                                lst1.append(tmp)
                                
                        for i in lst1:
                            outfile.write("%s\n\n" % i)

                            
                            
				             
                            
                    zipf.write(path+str(name1)+".txt")
                    print(request)
                    ##Here the code will end for a specific sheet type
                    
                ## start for new sheet type from her follow the same identation    
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
      if f.filename == '':
          
            
          return render_template("failure.html") 
      filename1 = secure_filename(f.filename)
      if filename1 != '':
          
          file_ext = os.path.splitext(filename1)[1]
          if file_ext not in app.config['UPLOAD_EXTENSIONS']:
              
              return render_template("file_type.html")
          
      cache['file']=f.filename
      f.save(f.filename)  
      #return redirect("/")
      return render_template("success.html", name = f.filename)
      

     



if __name__ == '__main__':
    os.environ['FLASK_ENV'] = 'development'
    app.run(debug=True)
    
