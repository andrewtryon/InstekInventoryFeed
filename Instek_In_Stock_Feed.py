from dotenv import load_dotenv
load_dotenv()
import os
from datetime import date
import subprocess
import time
import pandas as pd
import pyodbc
import requests

#This is the basic info for making wrike tasks.
def makeWrikeTask (title = "New Pricing Task", description = "No Description Provided", status = "Active", assignees = "KUALCDZR", folderid = "IEAAJKV3I4JBAOZD"):
    url = "https://www.wrike.com/api/v4/folders/" + folderid + "/tasks"
    querystring = {
        'title':title,
        'description':description,
        'status':status,
        'responsibles':assignees
        } 
    headers = {
        'Authorization': 'bearer TOKEN'.replace('TOKEN',os.environ.get(r"WRIKE_TOKEN"))
        }        
    response = requests.request("POST", url, headers=headers, params=querystring)
    print(response)
    return response

def markWrikeTaskComplete (taskid):
    url = "https://www.wrike.com/api/v4/tasks/" + taskid + "/"
    querystring = {
        'status':'Completed'
        }     
    headers = {
        'Authorization': 'bearer TOKEN'.replace('TOKEN',os.environ.get(r"WRIKE_TOKEN"))
    }

    response = requests.request("PUT", url, headers=headers, params=querystring)
    return response            

def attachWrikeTask (attachmentpath, taskid):
    url = "https://www.wrike.com/api/v4/tasks/" + taskid + "/attachments"
    headers = {
        'Authorization': 'bearer TOKEN'.replace('TOKEN',os.environ.get(r"WRIKE_TOKEN"))
    }

    files = {
        'X-File-Name': (attachmentpath, open(attachmentpath, 'rb')),
    }

    response = requests.post(url, headers=headers, files=files)
    return response     


if __name__ == '__main__':

    #Find most recently modified Inventory File
    workingdir = "C:\\Users\\andrew.tryon\\Dropbox\\Instek Inventory Files\\"
    directory = os.fsencode(workingdir)
    for file in os.listdir(directory):
        filename = os.fsdecode(file)
        print(filename)
        if filename.endswith(".xlsx"):  
            print('good file')  
            if 'modifieddate' in locals():
                if os.path.getmtime(workingdir + filename) > modifieddate:
                    print('better modified date found')
                    modifieddate = os.path.getmtime(workingdir + filename)
                    choosenFilename = workingdir + filename                         
            else:
                modifieddate = os.path.getmtime(workingdir + filename)
                choosenFilename = workingdir + filename                

    instekInv_df = pd.read_excel(choosenFilename)
    print(instekInv_df)
    instekInv_df.rename(columns={ instekInv_df.columns[0]: "Model" , instekInv_df.columns[1]: "Inventory" }, inplace=True)

    #Cleanup Model Numbers
    instekInv_df['Model'] = instekInv_df['Model'].astype(str)
    instekInv_df['Model'] = instekInv_df['Model'].str.strip()
    instekInv_df['Model'] = instekInv_df['Model'].str.encode('ascii', 'ignore').str.decode('ascii')
    instekInv_df['Model'] = instekInv_df['Model'].str.upper()
        
    print(instekInv_df)   
      
    #This is the connection string to Sage.
    sage_conn_str = os.environ.get(r"sage_conn_str").replace("UID=;","UID=" + os.environ.get(r"sage_login") + ";").replace("PWD=;","PWD=" + os.environ.get(r"sage_pw") + ";") 
            
    #This makes the connection to Sage based on the string above.
    cnxn = pyodbc.connect(sage_conn_str, autocommit=True)

    print('Grabbing Sage Data')

    #This is responsible for selecting what data to pull from Sage.
    sql = """SELECT 
                CI_Item.ItemCode, 
                CI_Item.InactiveItem,    
                IM_ItemVendor.VendorAliasItemNo,
                CI_Item.UDF_WEB_DISPLAY_MODEL_NUMBER
            FROM 
                CI_Item CI_Item, IM_ItemVendor IM_ItemVendor
            WHERE
                (CI_Item.PrimaryVendorNo = 'INST001') AND
                (CI_Item.ItemCode = IM_ItemVendor.ItemCode)
    """

    sagedf = pd.read_sql(sql,cnxn)
    sagedf['UDF_VENDOR_STOCK_LEVEL'] = 0
    sagedf['UDF_VENDOR_STOCK_LEVEL_DATE'] = date.today().strftime('%m/%d/%y')    

    print(sagedf)

    master_df = instekInv_df.merge(sagedf, left_on='Model', right_on='VendorAliasItemNo').append(instekInv_df.merge(sagedf, left_on='Model', right_on='UDF_WEB_DISPLAY_MODEL_NUMBER'))

    print(master_df)

    master_df = master_df.drop_duplicates(subset='ItemCode').set_index('ItemCode')
    master_df.loc[master_df['Inventory'] > 0,'UDF_VENDOR_STOCK_LEVEL'] = master_df['Inventory']
    master_df.loc[master_df['UDF_VENDOR_STOCK_LEVEL'] > 0,'InactiveItem'] = 'N'

    print(master_df)

    #exit()
    if master_df.shape[0] > 0:
        print('Saving Vendor Stock Levels upload file')
        filename = r"\\FOT00WEB\Alt Team\Kris\GitHubRepos\Vendor Stock Feeds\Instek\instek_inv_feed.csv"
        master_df.to_csv(filename, columns=['UDF_VENDOR_STOCK_LEVEL','UDF_VENDOR_STOCK_LEVEL_DATE','InactiveItem'], index=True)    

        print('Sleeping for 30 secs minute')
        time.sleep(30)
        print('Attempting VI to Sage the Instek Inventory Feed')

        #Auto_VENDOR_STOCK_VIWI7H
        p = subprocess.Popen('Auto_VENDOR_STOCK_VIWI7H.bat', cwd= r'Y:\Qarl\Automatic VI Jobs\Maintenance', shell = True)
        stdout, stderr = p.communicate()   
        p.wait()

        print('Done')