import azure.functions as func
import logging

#!/usr/bin/env python
# coding: utf-8

app = func.FunctionApp(http_auth_level=func.AuthLevel.ANONYMOUS)

import pandas as pd
import numpy as np
import warnings
import json
import os
import sys
import pyodbc
import pytz
from azure.storage.blob import BlobServiceClient, BlobClient, ContainerClient
from io import BytesIO
from zipfile import ZipFile
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders


# try:
#     import pyodbc
#     from azure.storage.blob import BlobServiceClient, BlobClient, ContainerClient
#     from io import BytesIO
#     from zipfile import ZipFile
#     import smtplib
#     from email.mime.multipart import MIMEMultipart
#     from email.mime.text import MIMEText
#     from email.mime.base import MIMEBase
#     from email import encoders
#     import pytz
# except Exception as e:
#     log_files(e,'Certain python libraries is not working, Kindly reach out to Tech Team:')


# custom functions here

exception_list=[]

warnings.filterwarnings('ignore')

connection_string = ""
container_name = ""
blob_name = f""

def log_files(exec,name):
    exception_list.append(name+str(exec))

try:
    blob_service_client = BlobServiceClient.from_connection_string(connection_string)
    container_client = blob_service_client.get_container_client(container_name)
    blob_list = list(container_client.list_blobs())
except Exception as e:
    log_files(e, 'Blob connection string is invalid or empty')


try:                  ## This is to signify if it contains any file in the blob then it will work else returns an error
    first_blob = blob_list[0]
    #blob_client = container_client.get_blob_client(first_blob)
except Exception as e:
    log_files(e,'The files couldnt be fetched: ')


def athena_email():
    sendermail=''
    senderpass=''
    return sendermail,senderpass 


def timestamp_blob(connec_string,blob):
    """
    Input: Connection String and Name of the first file in the blob
    **Note: The function fetches the file in Alphabetical order (if files present in the blob)
    Output: Returns the timestamp (IST) when the file has been uploaded
    """
    try:
        blob_client = container_client.get_blob_client(first_blob.name)
        properties = blob_client.get_blob_properties()
        upload_time_utc = properties['creation_time']
        utc_zone = pytz.utc
        ist_zone = pytz.timezone('Asia/Kolkata')
        upload_time_utc = upload_time_utc.replace(tzinfo=utc_zone)
        upload_time_ist = upload_time_utc.astimezone(ist_zone)
        return upload_time_ist
    except Exception as e:
        #str = 'The timestamp is not be made available'
        log_files(e,'The timestamp couldnt be made available:')
        #print(f'An error occured while timestamp:{e}')

def delete_blob(conn_string,blob):
    container_client = ContainerClient.from_connection_string(conn_string, container_name)
    blob_service_client = BlobServiceClient.from_connection_string(conn_string)
    container_client = blob_service_client.get_container_client(container_name)  
    blob_client = container_client.get_blob_client(blob)
    try:
        blob_client.delete_blob()  
        print(f"Blob '{blob.name}' deleted successfully.") 
    except Exception as e:
        log_files(e,'The blob couldnt be deleted:')

## The below function will fetch the json file contanining all the email of different markets
def json_fetch():    
    try:
    ## This function is hard-coded since it has to be only once for a single purpose
        connection_string = ""
        container_name = ""
        blob_name = "" 
        json_file_name = "" 
        blob_service_client = BlobServiceClient.from_connection_string(connection_string)
        container_client = blob_service_client.get_container_client(container_name)
        blob_client = container_client.get_blob_client(blob_name)
        download_stream = blob_client.download_blob()
        txt_content = download_stream.readall().decode('utf-8')
        data = json.loads(txt_content)
        return data
    except Exception as e:
        log_files(e,'The list of email recepient couldnt be fetched, the error for the same is: ')

def file_name(conn_string,blob):
    
    container_client = ContainerClient.from_connection_string(conn_string, container_name)
    blob_service_client = BlobServiceClient.from_connection_string(conn_string)
    container_client = blob_service_client.get_container_client(container_name)  
    try:
        blob_client = container_client.get_blob_client(blob)
        downloaded_blob = container_client.download_blob(blob)
        mci = pd.read_excel(BytesIO(downloaded_blob.content_as_bytes()), sheet_name='Market Channel info') ## First sheet in the excel workbook
        if(mci['BE'].values[0]==1 and mci.iloc[:,1:-1].values.sum()==0):
            file_name = mci['Market'].values[0]+'_BE.xlsx'
        elif(mci.iloc[:,1:].values.sum()>1):
            file_name = mci['Market'].values[0]+'_ACC.xlsx'
        return file_name
    except Exception as e:
        log_files(e,'  ')
def failure_email(connection_string, first_blob,exception_list):
    container_name = "" 
    container_client = ContainerClient.from_connection_string(connection_string, container_name)
    blob_service_client = BlobServiceClient.from_connection_string(connection_string)
    container_client = blob_service_client.get_container_client(container_name)  

    data = json_fetch()
    for i in data['Failure']:
        sender_email = ""
        receiver_email = i
        subject = f"Interim Failure Update! Modified files not qualified for final inputs for BO {first_blob.name}"
        time = timestamp_blob(connection_string,first_blob)
        body = f'''<html>
                    Dear Uploader,
                    <br>The file uploaded at {time} does not qualifies the quality checks.</br>
                    <br>The reason of rejection of the file is/are as follows:</br>
                    </html>
                ''' 
        
        for k in range(len(exception_list)):
            body+=f'''
                  <html>
                  <br>
                  {k+1}. {exception_list[k]} </br>
                  </html>
                   '''
        body += '''<html>
        <p>   <br>Kindly upload the file again after rectification.</br>
            <br><br>
            Thank you<br>
            Regards,<br>
            BO Quality Engine
            </p>
            </html>'''
    
        # Create the email object
        message = MIMEMultipart()
        message["From"] = sender_email
        message["To"] = receiver_email
        message["Subject"] = subject
        
        # Attach the body of the email
        message.attach(MIMEText(body, "html"))
 
        # Send the email
        smtp_server = "smtp.office365.com"
        smtp_port = 587
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(athena_email()[0], athena_email()[1])
            server.send_message(message)
        exception_list=[]
        #attachment.close()
        #server.quit()
    print(f'Failure Email for modified files sent to {i}')
    exception_list=[]
    exit()
    delete_blob(connection_string, first_blob)

def success_email(connection_string,first_blob):
    container_name = "" 
    container_client = ContainerClient.from_connection_string(connection_string, container_name)
    blob_service_client = BlobServiceClient.from_connection_string(connection_string)
    container_client = blob_service_client.get_container_client(container_name)  
    mci,_,_,_=modify_blob_fetch(connection_string,first_blob)
    market = mci['Market'].values[0]
    if(mci.iloc[:,1:].values.sum()>1):
        channel='ACC'
    elif(mci['BE'].values[0]==1 and mci.iloc[:,1:-1].values.sum()==0):
        channel='BE'
        
    data = json_fetch()
    for i in data[market][channel]:
        sender_email = athena_email()[0]
        receiver_email = i
        subject = f"Interim Success Update! Modified files qualified for final inputs for BO  {first_blob.name}"
        time = timestamp_blob(connection_string,first_blob)
        body = f'''<html> Dear Uploader,
                     <br>The file uploaded at {time} has passed all the relevant quality checks.</br>
                     <br>The file is further getting processed to final input creation, updates will be informed sooner.</br>
                     <p>
                     <br><br>
                     Thank you <br>
                     BO Quality Engine
                     </p>
                   </html>
                '''
        # Create the email object
        message = MIMEMultipart()
        message["From"] = athena_email()[0]
        message["To"] = receiver_email
        message["Subject"] = subject
        
        # Attach the body of the email
        message.attach(MIMEText(body, "html"))
 
        smtp_server = "smtp.office365.com"
        smtp_port = 587
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(athena_email()[0], athena_email()[1])
            server.send_message(message)
        print(f'Success Email for modified files sent to {i}')

def modify_blob_fetch(connection_string,blob):
    try:
        container_name = ""
        container_client = ContainerClient.from_connection_string(connection_string, container_name)
        downloaded_blob = container_client.download_blob(blob)
        mci = pd.read_excel(BytesIO(downloaded_blob.content_as_bytes()), sheet_name='Market Channel info') 
        container_client = ContainerClient.from_connection_string(connection_string, container_name)
        downloaded_blob = container_client.download_blob(blob)
        base_loc = pd.read_excel(BytesIO(downloaded_blob.content_as_bytes()), sheet_name='Base Location')
        container_client = ContainerClient.from_connection_string(connection_string, container_name)
        downloaded_blob = container_client.download_blob(blob)
        allocation_up = pd.read_excel(BytesIO(downloaded_blob.content_as_bytes()), sheet_name='Allocation')
        container_client = ContainerClient.from_connection_string(connection_string, container_name)
        downloaded_blob = container_client.download_blob(blob)
        dates = pd.read_excel(BytesIO(downloaded_blob.content_as_bytes()), sheet_name='Considered dates') 
        
        return mci,base_loc,allocation_up,dates
    except Exception as e:
        log_files(e,'The modified files uploaded couldnt be fetched, the possible reasons could be:')
        
        #exit()

def static_blob():
    try:
        conn_string = ""
        container_name = ""
        blob_name = ""
        container_client = ContainerClient.from_connection_string(conn_string, container_name)
        channel_mapping = {'On': 'ON', 'Off': 'OFF', 'BR': 'BR', 'Be': 'BE'}
        container_client = ContainerClient.from_connection_string(conn_string, container_name)
        downloaded_blob = container_client.download_blob(blob_name)
        mci = modify_blob_fetch(connection_string, first_blob)[0]
        
        field_visit = pd.read_excel(BytesIO(downloaded_blob.content_as_bytes()), sheet_name='Field Visit Detail')
        field_visit=field_visit[field_visit['Market']==mci['Market'].values[0]]
        field_visit = field_visit.reset_index().drop(columns=['index','Market'])
        filtered_dfs = []
        for short_name, long_name in channel_mapping.items():
            if mci[short_name].values[0] == 1:
                filtered_dfs.append(field_visit[field_visit['Channel'] == long_name])
        field_v = pd.concat(filtered_dfs, ignore_index=True) if filtered_dfs else pd.DataFrame()
        
        container_client = ContainerClient.from_connection_string(conn_string, container_name)
        downloaded_blob = container_client.download_blob(blob_name)
        visit_freq = pd.read_excel(BytesIO(downloaded_blob.content_as_bytes()), sheet_name='Visit Frequency Config')
        visit_freq=visit_freq[visit_freq['Market']==mci['Market'].values[0]]
        visit_freq = visit_freq.reset_index().drop(columns=['index'])
        
        container_client = ContainerClient.from_connection_string(conn_string, container_name)
        downloaded_blob = container_client.download_blob(blob_name)
        outlet_time_spent = pd.read_excel(BytesIO(downloaded_blob.content_as_bytes()), sheet_name='Time spent at outlet') 
        filtered_dfs = []
        for short_name, long_name in channel_mapping.items():
            if mci[short_name].values[0] == 1:
                filtered_dfs.append(outlet_time_spent[outlet_time_spent['Channel'] == long_name])
        outlet_time_spents = pd.concat(filtered_dfs, ignore_index=True) if filtered_dfs else pd.DataFrame()
        return field_v,visit_freq,outlet_time_spents
    except Exception as e:
        log_files(e,'The static files from the blob couldnt be fetched: ')
        
        #exit()

def qc_check_modify(connection_string,blob):
    try:
        visit_freq = static_blob()[1]
        mci,modify_base_loc,modify_allocation_up,_ = modify_blob_fetch(connection_string,blob)
        user = sql_data(connection_string,blob)[3]
#    except Exception as e:
#        log_files(e,'There is some error in the files which needs the revision, error is: ')
    
        #case 1-> 
        if (mci['Market'].values[0] not in visit_freq['Market'].unique()):
            log_files(' ','The market is not  valid, kindly change it and re-upload the file')
        else:
            pass
        ## case 2-> There should be 5 columns in the Market Channel info
        if(len(mci.columns)==5):
            pass
        else:
            log_files(' ','Certain columns in Market Channel info are missing')
        ## case 3-> The first column of Market should be of string type and other to be int or float type
        if(type(mci.columns[0])==str and pd.api.types.is_integer_dtype(mci[col]) for col in mci.iloc[:,1:]):
            pass
        else:
            log_files(' ','The file entered is not in the required format.')
        ## case 4-> There should be atleast one channel active as 1
        if(mci.iloc[:,1:].values.sum()>=1):
            pass
        else:
            log_files(' ','In Market Channel Info, atleast one of the channel should be one.')
        ## Cases to be considered for base location
        if(modify_base_loc.shape[0]>0):
            for i in range(0,modify_base_loc.shape[0]):
                emp = modify_base_loc['Sales_Rep'][i] in user['Name'].values
                sale_rep = modify_base_loc['Sales_Rep'].values[i]
                repl = str(modify_base_loc['Replacement Employee ID'][i]).split('.')[0] in user['EmployeeNumber'].values
                repl_rep=str(modify_base_loc['Replacement Employee ID'][i]).split('.')[0]
                channel = ['ON','OFF','BR','BE']
                for j in channel:
                    if(pd.isna(modify_base_loc[j].values[i])==True):
                        log_files(" ",f"{j} Channel entry is required for the sales representative {emp} in Base Location Sheet.")
                    else:
                        zero = modify_base_loc[j].values[i]==0
                        one = modify_base_loc[j].values[i]==1
                        if(zero or one):
                            pass
                        elif(not zero or not one):
                            log_files(" ",f"{j} Channel entry for sales representative {sale_rep} is not valid, it should be either 1 or 0 in Base Location Sheet.")
                if(modify_base_loc['Leaving?'].values[i]=='Y' or modify_base_loc['Leaving?'].values[i]=='N'):
                    if(modify_base_loc['Leaving?'].values[i]=='Y'):
                        if(pd.isna(modify_base_loc['Replacement Employee ID'].values[i])==True):
                            log_files(' ',f'For row {i+1}, the employee ID is a required field and not present in the Base Location sheet.')
                        elif(not repl):
                            log_files(" ",f'Sales representative ID: {repl_rep} in Base Location Sheet is not present in the SFDC records. ')
                        #else:
                        #    log_files(" ",'There is some technical error intercepted, kindly reach out to Technical team.')
                    elif(modify_base_loc['Leaving?'].values[i]=='N'):
                        
                        if(pd.isna(modify_base_loc['Rep Type'].values[i])==True):
                            log_files(" ",f"The sales rep at {i+1} entry in Base Location sheet doesn't contain Rep Type.")
                        else:
                            pass
                        if(modify_base_loc.iloc[i:i+1,5:].values.sum()>0):
                            pass
                        else:
                            log_files(" ",f"There should be atleast one channel to be allocated to the sales rep {sale_rep}")
                else:
                    log_files(' ',f'Wrong Input submitted in Leaving column of Base Location Sheet, it has to be in Y/N form.')
        if(modify_allocation_up.shape[0]>0):
            for i in range(0,modify_allocation_up.shape[0]):
                mm = modify_allocation_up['Sales_Rep'].values[i]
               
                if(pd.isna(modify_allocation_up['ID'].values[i])==True and pd.isna(modify_allocation_up['Sales_Rep'].values[i])==False):
                    log_files(" ",f"The ID has not been allocated to {mm} in the Allocation Sheet.")
                elif(pd.isna(modify_allocation_up['ID'].values[i])==True and pd.isna(modify_allocation_up['Sales_Rep'].values[i])==True):
                    log_files(" ",f"The ID and the sales rep hasnt been allocated for the outlet.")
                else:
                    pass
            return exception_list
    except Exception as e:
        log_files(e,"The operations of Quality checks couldnt be performed, the reason is:   ")
        
def modify_main(connection_string,first_blob,exception_list): 
    qc_check_modify(connection_string,first_blob)

    if(exception_list==[]):
        mci,base_loc,allocation_up,_ = modify_blob_fetch(connection_string,first_blob)
        print(mci)
        print()
        print(base_loc)
        print()
        print(allocation_up)
        success_email(connection_string,first_blob)
    elif(exception_list!=[]):
        try:
            mci,base_loc,allocation_up,_ = modify_blob_fetch(connection_string,first_blob)
            print(mci)
            print()
            print(base_loc)
            print()
            print(allocation_up)
        except Exception as e:
            print('Failure email for Modified files, atleast one of the sheet in modified files not recognized')
        failure_email(connection_string, first_blob,exception_list)
    else:
        pass
def sql_data(connection_string,blob):
    try:
        mci = modify_blob_fetch(connection_string,blob)[0]
        connection_string = ("")
        connection = pyodbc.connect(connection_string)
        cursor  = connection.cursor()
        query1 = ""
        query2 = ""
        query3 = ""
        query4 = ""
        query5 = ""
        sfdc={"outlet_data":query2,"allocation":query3,"base_location":query4,"user":query5}
        
        try:
            area_name = cursor.execute(query1%mci.values[0])
            area_name = cursor.fetchall()
            area_name = pd.read_sql_query(query1, connection)
            #arr = 
            area = area_name[area_name['Name']==mci['Market'].values[0]]['Id'].values
            arr = area[0]
        except Exception as e:
            log_files(e,'The Area is not valid in nature, hence need revision: ')
            final_input_check(connection_string,first_blob)
            #exit()
        
        qq = query2%arr
        outlet_data = cursor.execute(query2%arr)
        outlet_data = cursor.fetchall()
        channel_mapping = {'On': 'ON', 'Off': 'OFF', 'BR': 'BR', 'BE': 'BE'}
        
        outlet_data = pd.read_sql_query(query2%arr, connection)
        
        df = pd.DataFrame()
        dt = {'ON':'On', 'OFF':'Off', 'BE':'BE', 'BR':'BR'}
        for i, (key, j) in enumerate(dt.items()):
            if mci[j].values[0] == 0:
                continue
            else:
                filtered_data = outlet_data[outlet_data[key] == mci[j].values[0]]
                filtered_data['Channel'] = key
                df = pd.concat([df, filtered_data],ignore_index=True)   
         
        outlet_data  = df.drop(columns=['ON','OFF','BE','BR','AREA_ID'])   
        filtered_dfs = []
        for short_name, long_name in channel_mapping.items():
            if mci[short_name].values[0] == 1:
                filtered_dfs.append(outlet_data[outlet_data['Channel'] == long_name])
        outlet_datas = pd.concat(filtered_dfs, ignore_index=True) if filtered_dfs else pd.DataFrame()
        outlet_datas['New (Y)/Current(N) outlet']=np.where(outlet_datas['L3M sales ']>0.001,'N','Y')
    
    
        qqa = query3%area
        allocation = cursor.execute(query3%arr)
        allocation = cursor.fetchall()
        allocation = pd.read_sql_query(query3%arr, connection)
        #allocation = allocation[allocation['Market_Name__c']==mci['Market'].values[0]] 
        allocation = allocation.drop(columns=['Area_Lookup__c'])            
        ## This needs to be aligned with the mci
        filtered_dfs = []
        for short_name, long_name in channel_mapping.items():
            if mci[short_name].values[0] == 1:
                filtered_dfs.append(allocation[allocation['Previous Month Channel'] == long_name])
        allocations = pd.concat(filtered_dfs, ignore_index=True) if filtered_dfs else pd.DataFrame()
    
        qqb = query4%arr
        base_location = cursor.execute(query4%arr)
        base_location = cursor.fetchall()
        base_location = pd.read_sql_query(query4%arr, connection)
        base_location=base_location.groupby(['OwnerId']).max().reset_index().drop(columns=['OwnerId','Area_Lookup__c'])
        ## This needs to be aligned with the mci (channel)
        base_location['Cluster']=np.nan
        
        user = cursor.execute(query5)
        user = cursor.fetchall()
        user = pd.read_sql_query(query5, connection)
        return outlet_datas,allocations,base_location,user
    except Exception as e:
        log_files(e,"The files couldnt be fetched from SFDC: ")

def transformations(connection_string, first_blob):
    try:
        outlet_data,allocation,base_location,user = sql_data(connection_string,first_blob)
        mci,base_loc,allocation_up,dates=modify_blob_fetch(connection_string,first_blob)
        field_v,visit_freq,outlet_time_spent=static_blob()
        visit_freq=visit_freq.drop(columns=['Market'])
        if base_loc.shape[0]>0:
            for i in range(base_loc.shape[0]):
                emp = base_loc['Sales_Rep'][i] in user['Name'].values
                sale_rep = base_loc['Sales_Rep'].values[i]
                repl_rep=str(base_loc['Replacement Employee ID'][i]).split('.')[0]
                repl = str(base_loc['Replacement Employee ID'][i]).split('.')[0] in user['EmployeeNumber'].values
                updated_values = base_loc[['Rep Type', 'Sales_Rep','Replacement','ON', 'OFF', 'BE', 'BR']].iloc[i]
                name = user[user['EmployeeNumber']==str(base_loc['Replacement Employee ID'].values[i]).split('.')[0]]
#                if(not emp):
#                    log_files(" ",f'Sales representative: {sale_rep}, is not present in the SFDC records.  ')
                if(emp):
                    if(base_loc['Leaving?'].values[i]=='N'):
                        base_location.loc[base_location['Sales_Rep'] == base_loc['Sales_Rep'].values[i], 'Rep Type'] = updated_values['Rep Type']
                        base_location.loc[base_location['Sales_Rep'] == base_loc['Sales_Rep'].values[i], 'ON'] = updated_values['ON']
                        base_location.loc[base_location['Sales_Rep'] == base_loc['Sales_Rep'].values[i], 'OFF'] = updated_values['OFF']
                        base_location.loc[base_location['Sales_Rep'] == base_loc['Sales_Rep'].values[i], 'BE'] = updated_values['BE']  # Note the name difference
                        base_location.loc[base_location['Sales_Rep'] == base_loc['Sales_Rep'].values[i], 'BR'] = updated_values['BR']
                    
                    elif(base_loc['Leaving?'].values[i]=='Y'):
                        if(repl):
                            base_location.loc[base_location['Sales_Rep'] == base_loc['Sales_Rep'].values[i], 'Sales_Rep'] = updated_values['Replacement']
                            base_location.loc[base_location['Sales_Rep'] == name['Name'].values[i], 'ON'] = updated_values['ON']
                            
                            base_location.loc[base_location['Sales_Rep'] == name['Name'].values[i], 'OFF'] = updated_values['OFF']
                            
                            base_location.loc[base_location['Sales_Rep'] == name['Name'].values[i], 'BE'] = updated_values['BE']
                            
                            base_location.loc[base_location['Sales_Rep'] == name['Name'].values[i], 'BR'] = updated_values['BR']
                            
                            if(pd.isna(name['Base_Location__Latitude__s'].values[0])==True or pd.isna(name['Base_Location__Longitude__s'].values[0])==True):
                                va = name['Name'].values[0]
                                log_files(" ",f'The Geo-Coordinates (Lat/Long) of Sales representative {va} is not present in the SFDC records.')
                            elif(pd.isna(name['Base_Location__Latitude__s'].values[0])==False and pd.isna(name['Base_Location__Longitude__s'].values[0])==False):
                                print('Going into the geo-coordinates part')
                                base_location.loc[base_location['Sales_Rep']==name['Name'].values[0],'Lat']=name['Base_Location__Latitude__s'].values[0]
                                base_location.loc[base_location['Sales_Rep']==name['Name'].values[0],'Long']=name['Base_Location__Longitude__s'].values[0]
                            if(pd.isna(name['Rep Type'].values[0])==True):
                                bb = name['Name'].values[0]
                                log_files(" ",f'Sales Role is not defined for Sales Representative {bb} in the SFDC records.')
                            elif(pd.isna(name['Rep Type'].values[0])==False):
                                print('Going into the Rep Type part')
                                base_location.loc[base_location['Sales_Rep']==name['Name'].values[0],'Rep Type']=name['Rep Type'].values[0]    
                            if(base_loc['Sales_Rep'].values[i] in allocation['Sales_Rep'].values):
                                print('Going into the name change in allocation sheet')
                                allocation.loc[allocation['Sales_Rep']==base_loc['Sales_Rep'].values[i],'Sales_Rep']=updated_values['Replacement']
                            else:
                                rr = base_loc['Sales_Rep'].values[i]
                                log_files(" ",f'The leaving sales representative: {rr} didnt had any outlets allocated.')
                        else:
                            log_files(" ",f'Sales representative ID: {repl_rep} in Base Location Sheet is not present in the SFDC records.  ')
                else:
                    log_files(" ",f'Sales representative ID: {repl_rep} in Base Location Sheet is not present in the SFDC records.  ')
                                
        if(allocation_up.shape[0]>0):    
            for i in range(allocation_up.shape[0]):
                allocation.loc[allocation['ID']==allocation_up['ID'][i],'Sales_Rep']=allocation_up['Sales_Rep'].iloc[i]        
        past_pjp = pd.DataFrame()
        past_pjp['Account: Account Name']=np.nan
        past_pjp['Account: Channel']=np.nan
        past_pjp['Account: ID']=allocation['ID']
        past_pjp['Event: Owner Name']=allocation['Sales_Rep']
        past_pjp['Account: Account Owner']=np.nan
        past_pjp['Start date and time']=np.nan
        past_pjp['End date and time']=np.nan
        visit_freq = visit_freq.rename(columns = {'Segment':'segment'})
        visit_freq['Channel'] = visit_freq['Channel'].str.upper()
        outlet_data['Channel'] = outlet_data['Channel'].str.upper()
        outlet_data = outlet_data.merge(visit_freq, on = ['Channel','segment'], how = 'left')  
        outlet_data['ID_A']=outlet_data['ID']
        outlet_data = outlet_data.rename(columns = {'SO':'SO_Frequency', 'SM':'SM_Frequency'})
        outlet_data = outlet_data.drop('Key',axis = 1)
        outlet_data['New (Y)/Current(N) outlet'] = np.where(outlet_data['L3M sales ']>0.001, "N","Y")
        visit_freq = visit_freq.rename(columns = {'segment':'Segment'})
        return outlet_data,visit_freq,base_location,outlet_time_spent,allocation,past_pjp,field_v,dates    
    except Exception as e:  
        log_files(e,'The files couldnt be overwritten   ')
try:
    def data_modelling_allocation(allocation):
        allocation['ID_A']=allocation['ID']
        return allocation
    
    
    def header_check(data, var_cols,data_name):
        error  = []
        data_columns = data.columns
        column_missing = list(set(var_cols) - set(data_columns))
        columns_extra = list(set(data_columns) - set(var_cols))
        if len(column_missing) > 0:
            error.append(f"{(column_missing)} columns are missing in {data_name}")
        if len(columns_extra)>0:
            error.append(f"{(columns_extra)} columns are not expected in {data_name}")
        return error
    
    def Triangulation_Check(outlet_data, allocation, base_location):
        error = []
        new_rows = []
    
        for index, row in base_location.iterrows():
            for col in ['OFF', 'ON', 'BE', 'BR']:
                if row[col] == 1:
                    new_row = row.copy()
                    new_row['Channel Type'] = col  
                    new_rows.append(new_row)
    
        base_location = pd.DataFrame(new_rows)
    
        outlet_data['Channel'] = outlet_data['Channel'].str.lower()
        #allocation['ID_A']=allocation['ID']
        alloc = allocation[['ID_A', 'Sales_Rep', 'Previous Month Channel']]
        alloc['Previous Month Channel'] = alloc['Previous Month Channel'].str.lower()
        
        base_loc = base_location[['Sales_Rep', 'Rep Type', 'Channel Type']]
        base_loc = base_loc.rename(columns = {'Channel Type':'Sales Rep Channel'})
        base_loc['Sales Rep Channel'] = base_loc['Sales Rep Channel'].str.lower()
        
        outlet_data['Status'] = 'Hold'
        outlet_alloc = outlet_data.merge(alloc, on = 'ID_A', how = 'left')
        outlet_alloc['Status'] = np.where(outlet_alloc['Sales_Rep'].isnull(), 'Treat New', outlet_alloc['Status'])
        
        outlet_hold = outlet_alloc[outlet_alloc['Status'] == 'Hold']
        outlet_hold = outlet_hold.merge(base_loc, left_on = ['Sales_Rep', 'Channel'], right_on = ['Sales_Rep', 'Sales Rep Channel'], how = 'left')  
        ## Change to be verified. Join on Sales Rep only was there
        outlet_hold['Channel_change_flag'] = np.where(outlet_hold['Channel']==outlet_hold['Previous Month Channel'], 0, 1)
        
        outlet_hold_true = outlet_hold[outlet_hold['Channel_change_flag'] == 0]
        SM_VF_for_SM_alloc = outlet_hold_true[outlet_hold_true['Rep Type'] == 'SM']['SM_Frequency'].unique()
    
        outlet_hold_true['Rep_channel_flag'] = np.where(outlet_hold_true['Channel']==outlet_hold_true['Sales Rep Channel'], 0, 1)
        rep_of_wrong_channel = outlet_hold_true[outlet_hold_true['Rep_channel_flag'] == 1]
        
        if 0 in SM_VF_for_SM_alloc:
            error.append(f"Outlets allocated to SM but SM_VF = 0 for outlets {list(outlet_hold_true[(outlet_hold_true['Rep Type'] == 'SM') & (outlet_hold['SM_Frequency'] == 0)]['ID'].unique())} ")
        elif len(rep_of_wrong_channel)!=0:
            error.append(f"Sales Rep of Wrong Channel is Allocated for outlets {rep_of_wrong_channel['ID'].unique()}")
        else:
            None
            
        alloc_baseloc_inner = alloc.merge(base_loc, left_on = ['Sales_Rep', 'Previous Month Channel'], right_on = ['Sales_Rep', 'Sales Rep Channel'], how = 'inner')
        alloc_baseloc_outer = alloc.merge(base_loc, left_on = ['Sales_Rep', 'Previous Month Channel'], right_on = ['Sales_Rep', 'Sales Rep Channel'], how = 'outer')
    
        if alloc_baseloc_inner.shape[0] != alloc_baseloc_outer.shape[0]:
            alloc['chan_rep'] = alloc['Sales_Rep']+"_"+alloc['Previous Month Channel'].str.upper()
            base_loc['chan_rep'] = base_loc['Sales_Rep'] + "_" + base_loc['Sales Rep Channel'].str.upper()
            chanrep_in_alloc_notin_baseloc = set(alloc['chan_rep'].unique()) - set(base_loc['chan_rep'].unique())
            chanrep_in_baseloc_notin_alloc = set(base_loc['chan_rep'].unique()) - set(alloc['chan_rep'].unique())
            if len(chanrep_in_alloc_notin_baseloc) > 0:
                error.append(f"channel*rep not matching in allocation and base location. channel*rep present in allocation but not in base locations are {chanrep_in_alloc_notin_baseloc}.Please ignore if the Channel of these outlets have changed.")
            #elif len(chanrep_in_baseloc_notin_alloc) > 0:
            #    error.append(f"channel*rep not matching in allocation and base location. channel*rep present in base location but not in allocations are {chanrep_in_baseloc_notin_alloc}.Please ignore if there is any new sales representative.")
        return error
    
    def Null_Check(data, data_name):
        error = []
        df_null = data.isnull().sum().to_frame().reset_index()
        df_null.columns = ['col_name', 'null_count']
        null_columns = df_null[df_null['null_count'] != 0]['col_name'].to_list()
        if len(null_columns) !=0:
            error.append(f"{data_name} has Columns with Null Values: {null_columns}")
        else:
            pass
        return error
    
    def Unique_Check(data, column, data_name):
        error = []
        tot_cnt = data.shape[0]
        #data[column] = data[column].apply(lambda x: x.strip())
        uniq_cnt_col = len(data[column].unique())
        if uniq_cnt_col != tot_cnt:
            error.append(f"{column} has dupID_Aate values in {data_name}")
    
        return error
    
    def Missing_Geocode_Allocation(data):
        error = []
        missing_lat_cnt = data['Previous Month Lat'].isnull().sum()
        missing_long_cnt = data['Previous Month Long'].isnull().sum()
        if (missing_lat_cnt != 0) or (missing_long_cnt != 0):
            error.append(f"Lat or Long missing in Allocation sheet for outlets {list(data[(data['Previous Month Lat'].isnull())|(data['Previous Month Long'].isnull())]['ID'].unique())}")
        return data.dropna(subset=['Previous Month Lat', 'Previous Month Long'], how='any'), error
        
    def Sales_Rep_Dropped(data):
        error = []
        uniq_val = data['Sales_Rep'].str.lower().unique()
        for val in uniq_val:
            if val in ['dropped', 'unassigned',np.nan]:
                error.append(f"unexpected values present in Sales_Rep of allocation sheet for outlets: {list(data[data['Sales_Rep'].str.lower().isin(['dropped', 'unassigned',np.nan])]['ID'].unique())}")
        data_True = data[~data['Sales_Rep'].isin(['dropped', 'unassigned',np.nan])]
        return data_True,error
    
    def Zero_Sales_Check(data):
        error = []
        tot_cnt = data.shape[0]
        zero_sales_cnt = data[data['L3M sales '] == 0.001].shape[0]
        percent_zero_sales = zero_sales_cnt/tot_cnt*100
        if percent_zero_sales > 75:
            error.append("In outlet Level Data, change all 0 values of L3M sales  to 0.001")
        return error
    
    def base_location_channel(base_location):
        error = []
        base_location['Channel_sum'] = base_location['OFF'] + base_location['ON'] + base_location['BE'] + base_location['BR']
        non_channel_reps = base_location[base_location['Channel_sum'] == 0]['Sales_Rep'].unique()
        if len(non_channel_reps) > 0:
            error.append(f"Sales Rep with no channel marked in base location are : {non_channel_reps}")
        return error
    
    def Current_Allocation():
        all_error = []
        header_error = []
        null_error = []
        uniq_error = []
        outlet_data,vf_config,base_location,time_spent,allocation,past_pjp,field_visit,_ = transformations(connection_string, first_blob)
        #_,outlet_data, vf_config, base_location, time_spent, allocation, past_pjp, field_visit,_=transformation()
        allocation =  data_modelling_allocation(allocation)
        #outlet_data =  transformation()[1]  
    
    
        outlet_cols = ['ID_A', 'ID', 'BUYER NAME', 'SHOP ADDRESS', 'Revised Lat','Revised Long', 'L3M sales ', 'segment', 'SO_Frequency','SM_Frequency', 'Channel', 'New (Y)/Current(N) outlet']
        vf_config_cols = ['Channel', 'Segment', 'SO', 'SM', 'Key']
        base_loc_cols = ['Rep Type', 'Cluster', 'Sales_Rep', 'Lat', 'Long', 'OFF', 'ON', 'BE', 'BR']
        time_spent_cols = ['Channel', 'Segment', 'Time(mins)']
        allocation_cols = ['Sales_Rep', 'ID_A', 'ID', 'Outlet Name', 'Previous Month Lat', 'Previous Month Long', 'Previous Month Channel']
        past_pjp_cols = ['Account: Account Name', 'Account: Channel', 'Account: ID','Event: Owner Name', 'Account: Account Owner', 'Start date and time','End date and time']
        field_visit_cols = ['Channel', 'Rep Type', '# Days of field visit', 'Max visits per day']
    
        datas_header = [outlet_data, vf_config, base_location, time_spent, allocation, past_pjp, field_visit]
        data_names_header = ['Outlet Level Data','Visit Frequency Config','Base Location','Time spent at outlet','Allocation','Past PJP for overlap matrix','Field Visit Detail']
        vars = [outlet_cols, vf_config_cols, base_loc_cols, time_spent_cols, allocation_cols, past_pjp_cols, field_visit_cols]
    
        for file, var_col, file_name in zip(datas_header, vars,data_names_header):
            if len(header_check(file, var_col,file_name))>0:
                header_error.append(header_check(file, var_col,file_name)[0])
    
    
        base_loc_error = base_location_channel(base_location)
    
        datas_null = [outlet_data, vf_config, time_spent, allocation, field_visit]
        data_names_null = ['Outlet Level Data','Visit Frequency Config','Time spent at outlet','Allocation','Field Visit Detail']
        for file,file_name in zip(datas_null,data_names_null):
            if len(Null_Check(file, file_name)) > 0:
                null_error.append(Null_Check(file, file_name)[0])
    
        datas_uniq = [outlet_data, allocation]
        data_names_null = ['Outlet Level Data','Allocation']
        column_names = ['ID','ID']
        for file, col, file_name in zip(datas_uniq,column_names, data_names_null):
            if len(Unique_Check(file, col, file_name))>0:
                uniq_error.append(Unique_Check(file, col, file_name)[0])
    
        allocation_corrected1, geocode_error = Missing_Geocode_Allocation(allocation)
        allocation_corrected2, salesrep_error = Sales_Rep_Dropped(allocation_corrected1)
        triangulation_error = Triangulation_Check(outlet_data, allocation_corrected2, base_location)
        zerosales_error = Zero_Sales_Check(outlet_data)
    
        for i in header_error:
            all_error.append(i)
    
        for error in [null_error,uniq_error,base_loc_error,geocode_error,salesrep_error,triangulation_error,zerosales_error]:
            if len(error) > 0:
                all_error = all_error + error
    
        return all_error
    
    def Fresh_Allocation():
        all_error = []
        header_error = []
        null_error = []
        uniq_error = []
        #_,outlet_data, vf_config, base_location, time_spent, allocation, past_pjp, field_visit,_=transformation()
        outlet_data,vf_config,base_location,time_spent,allocation,past_pjp,field_visit,_ = transformations(connection_string, first_blob)
        allocation = data_modelling_allocation(allocation)
        #outlet_data = transformation()[1]
    
    
        outlet_cols = ['ID_A', 'ID', 'BUYER NAME', 'SHOP ADDRESS', 'Revised Lat','Revised Long', 'L3M sales ', 'segment', 'SO_Frequency','SM_Frequency', 'Channel', 'New (Y)/Current(N) outlet']
        vf_config_cols = ['Channel', 'Segment', 'SO', 'SM', 'Key']
        base_loc_cols = ['Rep Type', 'Cluster', 'Sales_Rep', 'Lat', 'Long', 'OFF', 'ON', 'BE', 'BR']
        time_spent_cols = ['Channel', 'Segment', 'Time(mins)']
        allocation_cols = ['Sales_Rep', 'ID_A', 'ID', 'Outlet Name', 'Previous Month Lat', 'Previous Month Long', 'Previous Month Channel']
        past_pjp_cols = ['Account: Account Name', 'Account: Channel', 'Account: ID','Event: Owner Name', 'Account: Account Owner', 'Start date and time','End date and time']
        field_visit_cols = ['Channel', 'Rep Type', '# Days of field visit', 'Max visits per day']
    
        datas = [outlet_data, vf_config, base_location, time_spent, allocation, past_pjp, field_visit]
        data_names = ['Outlet Level Data','Visit Frequency Config','Base Location','Time spent at outlet','Allocation','Past PJP for overlap matrix','Field Visit Detail']
        vars = [outlet_cols, vf_config_cols, base_loc_cols, time_spent_cols, allocation_cols, past_pjp_cols, field_visit_cols]
        
        for file, var_col, file_name in zip(datas, vars,data_names):
            if len(header_check(file, var_col,file_name))>0:
                header_error.append(header_check(file, var_col,file_name)[0])
    
        base_loc_error = base_location_channel(base_location)
    
        datas_null = [outlet_data, vf_config, time_spent, field_visit]
        data_names_null = ['Outlet Level Data','Visit Frequency Config','Time spent at outlet','Field Visit Detail']
        for file,file_name in zip(datas_null,data_names_null):
            if len(Null_Check(file, file_name)) > 0:
                null_error.append(Null_Check(file, file_name)[0])
    
        datas_uniq = [outlet_data, past_pjp]
        data_names_null = ['Outlet Level Data','Past PJP for overlap matrix']
        column_names = ['ID','Account: ID']
        for file, col, file_name in zip(datas_uniq,column_names, data_names_null):
            if len(Unique_Check(file, col, file_name))>0:
                uniq_error.append(Unique_Check(file, col, file_name)[0])
                
    
        zerosales_error = Zero_Sales_Check(outlet_data)
    
        for i in header_error:
            all_error.append(i)
    
        for error in [null_error,uniq_error,base_loc_error,zerosales_error]:
            if len(error) > 0:
                all_error = all_error + error
    
        return all_error
    
    def main():
        outlet_data,vf_config,base_location,time_spent,allocation,past_pjp,field_visit,_ = transformations(connection_string, first_blob)
        outlet_data['ID_A']=outlet_data['ID']
        allocation['ID_A']=allocation['ID']
        allocation = allocation.drop_dupID_Aates().dropna()
        if len(allocation)>0:
            #print('Current Allocation')
            error = Current_Allocation()
            for i in error:
                log_files(i," The error in in Current Allocation ")    
            
        else:
            #print('Fresh Allocation')
            error = Fresh_Allocation()
            for i in error:
                log_files(i," The error is in Fresh Allocation")
except Exception as e:
    log_files(e,'There are certain errors in the final QC checks: ')
    
def final_input_check(connection_string,first_blob,exception_list):
    try:
        data = json_fetch()    
        
        outlet_data,visit_freq,base_location,outlet_time_spent,allocation,past_pjp,field_v,dates=transformations(connection_string,first_blob)
        excel_file_path = file_name(connection_string,first_blob)
        time = timestamp_blob(connection_string,first_blob.name)
        mci,base_loc,allocation_up,dates=modify_blob_fetch(connection_string,first_blob)
        
        #mci,base_loc,allocation_up,dates = modified_files(connection_string,first_blob)    
        
        mci = modify_blob_fetch(connection_string,first_blob)[0]
        market = mci['Market'].values[0]
        if(mci.iloc[:,1:].values.sum()>1):
            channel = 'ACC'
        elif(mci['BE'].values[0]==1 and mci.iloc[:,1:-1].values.sum()==0):
            channel = 'BE'
        
        modified_sfdc = file_name(connection_string,first_blob)
        if (exception_list==[]):
            excel_file_path = file_name(connection_string,first_blob)    
            with pd.ExcelWriter(excel_file_path, engine='xlsxwriter') as writer:
                outlet_data.to_excel(writer, sheet_name='Outlet Level Data', index=False)
                visit_freq.to_excel(writer, sheet_name='Visit Frequency', index=False)
                base_location.to_excel(writer, sheet_name='Base Location', index=False)
                outlet_time_spent.to_excel(writer, sheet_name='Outlet Time Spent', index=False)
                allocation.to_excel(writer, sheet_name='Allocation', index=False)
                past_pjp.to_excel(writer, sheet_name='Past PJP', index=False)
                field_v.to_excel(writer, sheet_name='Field Visit', index=False)
                dates.to_excel(writer, sheet_name='Considered Dates', index=False)
            for i in data[market][channel]+data['Success']:
                sender_email = athena_email()[0]
                receiver_email = i
                subject = f'Final Success Update! Input file is ready for {excel_file_path} '
                body = f'''<html>
                            <p>
                                Dear uploader,<br>
                                The input file (PFA) created at {time} for BO has passed all the quality checks and is ready for next 
                                steps. Please take it forward from here.
                            </p>
                        </html>'''
                body += '''<html>
                            <p>
                                <br><br>
                                Thank you<br>
                                Regards,<br>
                                BO Quality Engine
                            </p>
                        </html>'''
    
               
                msg = MIMEMultipart()
                msg['From'] = athena_email()[0]
                msg['To'] = i
                msg['Subject'] = subject
                msg.attach(MIMEText(body, 'html'))
                smtp_server = 'smtp.office365.com'
                smtp_port = 587
                
                with open(excel_file_path, 'rb') as attachment:
                    part = MIMEBase('appID_Aation', 'vnd.ms-excel')
                    part.set_payload(attachment.read())
                    encoders.encode_base64(part)
                    part.add_header(
                        'Content-Disposition',
                        f'attachment; filename={excel_file_path}')
                    msg.attach(part)
                
                # Attach the body of the email
                msg.attach(MIMEText(body, 'html'))
                server = smtplib.SMTP(smtp_server, smtp_port)
                server.starttls()
                server.login(athena_email()[0], athena_email()[1])
                server.sendmail(athena_email()[0], i, msg.as_string())
                print('The success email sent to: ',i)
                attachment.close()
                server.quit()
            for i in data['Modifications']:
                if(base_loc.shape[0]>0 or allocation_up.shape[0]>0):
                        with pd.ExcelWriter(modified_sfdc, engine='xlsxwriter') as writer:
                            base_loc.to_excel(writer, sheet_name='Base Location', index=False)
                            allocation_up.to_excel(writer, sheet_name='Allocation', index=False)
                        #for i in 'praveen.shahani@bira91.com':##data['Modifications']:
                        subject = f'Final File Failure Update! Input file requires correction {modified_sfdc}'
                        body = f'''<html>
                                    <p>
                                        Dear SFDC Team,<br>
                                        PFA the modifications received for User Master and User-Account Mapping at {time}. 
                                        
                                        Please complete 
                                        your due diligence with the regional teams and make relevant changes in SFDC.
                                    </p>
                                </html>'''
                        body += '''<html>
                                    <p>
                                        <br><br>
                                        Thank you<br>
                                        Regards,<br>
                                        BO Quality Engine
                                    </p>
                                </html>'''
                        smtp_server = 'smtp.office365.com'
                        smtp_port = 587
                        
                        # Create a multipart message
                        mail = json_fetch()
                        
                        
                        msg = MIMEMultipart()
                        msg['From'] = athena_email()[0]
                        msg['To'] = i
                        msg['Subject'] = 'Modified files successfully triggered'
                        msg.attach(MIMEText(body, 'html'))
                        mci,base_loc,allocation_up
            
                        with open(modified_sfdc, 'rb') as attachment:
                            part = MIMEBase('appID_Aation', 'vnd.ms-excel')
                            part.set_payload(attachment.read())
                            encoders.encode_base64(part)
                            part.add_header(
                                'Content-Disposition',
                                f'attachment; filename={modified_sfdc}',
                            )
                        msg.attach(part)
                        msg.attach(MIMEText(body, 'html'))
                        server = smtplib.SMTP(smtp_server, smtp_port)
                        server.starttls()
                        server.login(athena_email()[0], athena_email()[1])
                        server.sendmail(athena_email()[0],i, msg.as_string())
                        print('The SFDC email sent to: ',i)
                        attachment.close()
                        server.quit()
                else:
                    pass
        
        else:
            with pd.ExcelWriter(excel_file_path, engine='xlsxwriter') as writer:
                outlet_data.to_excel(writer, sheet_name='Outlet Level Data', index=False)
                visit_freq.to_excel(writer, sheet_name='Visit Frequency', index=False)
                base_location.to_excel(writer, sheet_name='Base Location', index=False)
                outlet_time_spent.to_excel(writer, sheet_name='Outlet Time Spent', index=False)
                allocation.to_excel(writer, sheet_name='Allocation', index=False)
                past_pjp.to_excel(writer, sheet_name='Past PJP', index=False)
                field_v.to_excel(writer, sheet_name='Field Visit', index=False)
                dates.to_excel(writer, sheet_name='Considered Dates', index=False)
            for i in data[market][channel]:
                #i = "\n".join(exception_list)
                subject = f'Modified files to be re-uploaded: Process Stopped for {file_name(connection_string,first_blob)} '
                body = f'''<html>
                            <p>
                                Dear uploader,<br>
                                The input file (PFA) created at {time} for BO has not passed all the 
                                quality checks.There are certain errors which are as follows.
                            </p>
                        </html>'''
                for k in range(len(exception_list)):
                    body+=f'''
                          <html>
                          <br>
                          {k}. {exception_list[k]}
                          </html>
                           '''
                body += '''<html>
                            <p>
                                Please rectify the aforementioned errors by updating the back-end data and kindly re-upload the file in the blob.
                                <br><br>
                                Thank you<br>
                                Regards,<br>
                                BO Quality Engine
                            </p>
                        </html>'''
                smtp_server = i
                smtp_port = 587
        
                # Create a multipart message
                msg = MIMEMultipart()
                msg['From'] = athena_email()[0]
                msg['To'] = i
                msg['Subject'] = subject
                
            
            # Attach the body of the email
                msg.attach(MIMEText(body, 'html'))
                with open(excel_file_path, 'rb') as attachment:
                    part = MIMEBase('appID_Aation', 'vnd.ms-excel')
                    part.set_payload(attachment.read())
                    encoders.encode_base64(part)
                    part.add_header(
                        'Content-Disposition',
                        f'attachment; filename={excel_file_path}',
                    )
                    msg.attach(part)
                server = smtplib.SMTP(smtp_server, smtp_port)
                server.starttls()
                server.login(athena_email()[0], athena_email()[1])
                server.sendmail(athena_email()[0], i, msg.as_string())
                print('Failure Mail shared to',i)
                attachment.close()
                server.quit()
                exception_list=[]
        #delete_blob(connection_string, first_blob)        
    except Exception as e:
        exit()
        # delete_blob(connection_string, first_blob)

def bo_automation(connection_string,first_blob, exception_list):
    modify_main(connection_string,first_blob,exception_list)
    final_input_check(connection_string,first_blob,exception_list)
    delete_blob(connection_string,first_blob)


# Invoked Custom Function

def testFunctionCreation():
    a = modify_main(connection_string,first_blob,exception_list)
    b = final_input_check(connection_string,first_blob,exception_list)
    return a, b

# ========================================

@app.route(route="testFunction")
def testFunction(req: func.HttpRequest) -> func.HttpResponse:
    logging.info('Python HTTP trigger function processed a request.')

    name = req.params.get('name')
    if not name:
        try:
            result = testFunctionCreation()
            req_body = req.get_json()
        except ValueError:
            pass
        else:
            name = req_body.get('name')

    if name:
        return func.HttpResponse(f"Hello, {name}. This HTTP triggered function executed successfully.")
    else:
        return func.HttpResponse(
             "This HTTP triggered function executed successfully. Pass a name in the query string or in the request body for a personalized response.",
             status_code=200
        )