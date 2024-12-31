import urllib, urllib.request, winreg as reg, sys, win32com.client, win32gui, os, pythoncom, json, subprocess, time, threading
import datetime, pandas as pd, shutil
from pywinauto.application import Application
from win10toast import ToastNotifier
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from calendar import monthrange
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.select import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException, NoAlertPresentException, StaleElementReferenceException
from msedge.selenium_tools import Edge, EdgeOptions  ####remove if not working
from webdriver_manager.chrome import ChromeDriverManager
from webdriver_manager.firefox import GeckoDriverManager
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException, NoAlertPresentException, StaleElementReferenceException
import socket, logging, traceback, configparser, base64, hashlib
from cryptography.hazmat.primitives.ciphers import Cipher, algorithms, modes
from cryptography.hazmat.backends import default_backend
from cryptography.hazmat.primitives import padding
import mysql.connector
sys.path.append(".\\packages")


class frameworktemplate:

    def __init__(self, Bot_ID, Bot_Name):
        self.Bot_ID = Bot_ID
        self.Bot_Name = Bot_Name
        self.Bot_Data = self.Bot_ID + "_" + self.Bot_Name
        self.OnedrivePath = (str(os.environ['USERPROFILE']) + "\\OneDrive - ZF Friedrichshafen AG\\RPA\\" + self.Bot_Data + "\\")
        self.pass_count = 0
        self.fail_count = 0
        self.unprocessed_count = 0
        self.total_count = 0
        
    def Get_UserRootpath(self, main_folder="ImmerAI", sub_folder=''):
        user_path = str(os.environ['USERPROFILE'])+"\\"
        ImmerAI_path = user_path+main_folder+"\\"
        Bot_path = ImmerAI_path+sub_folder+"\\"
        
        if not os.path.exists(ImmerAI_path):
            os.makedirs(ImmerAI_path)
        if os.path.exists(Bot_path):
            shutil.rmtree(Bot_path)
        os.makedirs(Bot_path)
        
        return Bot_path
    
    def Reset_Folders(self, folderlist):
        for folder in folderlist:
            if os.path.exists(folder):
                shutil.rmtree(folder)
            os.makedirs(folder)
            
    def Delete_Directories(self, folderlist):
        for folder in folderlist:
            if os.path.exists(folder):
                shutil.rmtree(folder)
                
    def Get_OnedrivePath(self):
        return (str(os.environ['USERPROFILE'])+"\\OneDrive - ZF Friedrichshafen AG\\")
    
    def Get_UserProfile(self):
        return str(os.environ['USERPROFILE'])
    
    def Check_Onedrive_SyncStatus(self, folder, seconds):
        (pd.DataFrame()).to_excel(folder + "dummy.xlsx")
        time.sleep(seconds)
        os.remove(folder + "dummy.xlsx")
        
    def RemoveFile(self, files):
        for file in files:
            if os.path.exists(file):
                os.remove(file)
    
    def delete_genpy_folder(self):
        gen_py_path = str(os.environ['USERPROFILE'])+"\\AppData\\Local\\Temp\\gen_py"
        if os.path.exists(gen_py_path):
            shutil.rmtree(gen_py_path)
    
    def Getting_FrameworkModule(self, framework_name: str, application_namelist):
        module = __import__(str(sys.argv[5])) if len(sys.argv) >= 4 else __import__(str(framework_name))
        List = []
        for app in application_namelist:
            framework_class = getattr(module, str(app))
            List.append(framework_class())
        return List

    def Get_UserEmail_Data(self, framework):
        user_mail = framework.get_user_mail()
        dev_mail = framework.get_developer_mail()
        return [user_mail, dev_mail]

    def Get_SMTP_Data(self, framework):
        Auth_mail_password = framework.MailData('password')
        smtp_authentication = framework.MailData('smtp_authentication')
        return [smtp_authentication, Auth_mail_password]

    def DecryptionStatus(self, framework):
        DECRYPTED_BOOL = framework.decrypt()
        if not DECRYPTED_BOOL:
            return "Anonymous Act found! Access denied"
        else:
            return "Proceed"

    def get_pyd_path(self):
        # Specify the registry key you want to access
        key_path = r"SYSTEM\CurrentControlSet\Control"
        key_name = "Immer"
        error = False
        error_message = ""
        file_path = ""
        try:
            # Open the registry key for reading
            with reg.OpenKey(reg.HKEY_CURRENT_USER, key_path) as key:
                # Read the value of the specified registry entry
                file_path, _ = reg.QueryValueEx(key, key_name)
        except Exception as err:
            error = True
            error_message = str(err)
        return {"error": error, "file_path": file_path, "error_message": error_message}

    def main(self):
        pyd_path = self.get_pyd_path()
        if not pyd_path['error']:
            sys.path.append(pyd_path['file_path'])
        if len(sys.argv) >= 4:
            return True
        else:
            return False

    def GetPath(self, RelativePath):
        return (self.Root + "\\" + RelativePath)

    def DataFrame_Excelwriter(self, dataframe, excelpath, sheetname, sheet_to_create=False):
        if not sheet_to_create:
            df = pd.concat(dataframe)
            writer = pd.ExcelWriter(path=excelpath, engine='openpyxl')
            df.to_excel(writer, sheet_name=sheetname, index=False, header=True)
            writer.save()
        else:
            df = pd.concat(dataframe)
            with pd.ExcelWriter(excelpath, engine='openpyxl', mode='a') as writer:
                workBook = writer.book
                try:
                    workBook.remove(workBook[sheetname])
                    print("Sheets removed")
                except:
                    print("There is no such sheet in this file")
                finally:
                    df.to_excel(writer, sheet_name=sheetname, index=False)
                    writer.save()

    def Exception_Subject_and_Body(self, framework, err, content):
        err = str(err) + "\n" + str(traceback.format_exc())
        subject = self.Bot_Data + " - Failed"
        body = "Hi Team, \n\nThis Email is to formally notify you that the " + str(self.Bot_Data) + " process has not been completed successfully. \n\n" + str(content) + "\n\n" +str(err) + "\n\n\n\nPlease reach out RPA Team for more information.\n* This is an automated report from the RPA system.\n\n\nBy,\nRPA Team\n\n\n"
        return [subject, body]

    def Get_EmailSubject(self):
        return (self.Bot_Data+" - Passed")

    def Get_EmailBody(self, content:str):
        return ("Hi Team, \n\nThis Email is to formally notify you that the "+str(self.Bot_Data)+" process has been completed successfully.\n\n"+content+"\n\n\n\nPlease reach out RPA Team (RPA.Support@zf.com) for more information.\n* This is an automated report from the RPA system.\n\n\nBy,\nRPA Team\n\n\n")

    def Complete_Success_Logging(self, framework):
        self.pass_count = 1
        framework.log_completed(self.pass_count, self.fail_count, self.unprocessed_count, self.total_count)

    def Complete_Failure_Logging(self, framework):
        self.fail_count = 1
        framework.close_function(self.pass_count, self.fail_count, self.unprocessed_count, self.total_count)

    def Getting_DB_Values(self, delaytime):
        database = mysql.connector.connect(host=host, user=user, passwd=passwd, db=db)
        cursor = database.cursor()
        sql = f"""SELECT 
                            tbpd.bot_details_id,
                            tbpd.`status`,
                            tbpd.executor_id,
                            tbpm.bot_id,
                            tbpm.bot_name,
                            tlm.user_name,
                            tlm.email,
                            tlm.user_position,
                            tbpd.scheduled_on_at,
                            tem.online_status,
                            tbpm.bot_type FROM tbl_bot_process_details tbpd
                            LEFT JOIN tbl_bot_process_master tbpm ON tbpd.bot_master_id = tbpm.bot_master_id
                            LEFT JOIN tbl_login_master tlm ON tlm.user_name = tbpd.runner_name
                            LEFT JOIN tbl_executor_master tem ON tem.executor_id = tbpd.executor_id WHERE tbpd.`status` = 4 AND 
                            tbpd.scheduled_on_at <= '{delaytime}' ORDER BY tbpd.bot_details_id desc"""
        cursor.execute(sql)
        result = cursor.fetchall()

        return result
    

if __name__ == "__main__":
    try:
        template = frameworktemplate(Bot_ID="Support_001", Bot_Name="FindMissedSchedules")
        Bot_Data = template.Bot_Data
        Root = template.Get_UserRootpath(main_folder="RPA_Support", sub_folder=Bot_Data)
        
        Production = template.main()
        FrameworkList = template.Getting_FrameworkModule(framework_name="framework_v4_3_1", application_namelist=["SAPFramework"])
        obj_frame = FrameworkList[0]
        obj_frame.immer_ai_Excelkill()
        onedrive_path = template.Get_OnedrivePath()+"RPA_Support\\"+Bot_Data+"\\"
        user_mail = dev_mail = smtp = smtp_auth = None
        
        if Production:
            if (template.DecryptionStatus(framework=obj_frame)) != "Proceed":
                raise Exception(template.DecryptionStatus(framework=obj_frame))
            
            system_asset = obj_frame.get_system_assets("Asset_Support_001")
            to = system_asset[0]['value']
            cc = ""
            host = system_asset[1]['value']
            db = system_asset[2]['db']
            user = system_asset[3]['value']
            passwd = system_asset[4]['passwd']
            duration = system_asset[5]['duration']
            
            smtp = obj_frame.MailData('smtp_authentication')
            smtp_auth = obj_frame.MailData('password')
        else:
            userdata = obj_frame.immer_ai_GetJsonValues(JsonPath=onedrive_path+"userinfo.json")
            to = userdata['To']
            cc = userdata['Cc']
            host = userdata['host']
            db = userdata['db']
            user = userdata['user']
            passwd = userdata['passwd']
            duration = userdata['duration']
            smtp = userdata['smtp']
            smtp_auth = userdata['smtp_auth']
            
        # Enter Code
        template.Reset_Folders([Root+"Report\\"])
        now = datetime.datetime.now().strftime("%d-%m-%Y_%H-%M")
        report_path = Root+"Report\\"
        report_file = report_path+"Report_"+str(now)+".xlsx"
        
        delaytime = (datetime.datetime.now(datetime.timezone.utc) - datetime.timedelta(minutes=duration)).strftime("%Y-%m-%d %H:%M:%S")
        db_result = template.Getting_DB_Values(delaytime)
        db_count = len(db_result)
        pending_no = "Date:"+str(now)+" - No of Pendings --> "+str(db_count)
        df=pd.DataFrame(columns=["Bot ID", "Bot Name", "Scheduled Time", "Machine Name", "Executor Status"])
        if db_count>0:
            for i in range(0, db_count):
                df.loc[i, "Bot ID"] = str((db_result[i])[3])
                df.loc[i, "Bot Name"] = str((db_result[i])[4])
                df.loc[i, "Scheduled Time"] = ((db_result[i])[8]).strftime("%d-%m-%Y_%H-%M_UTC")
                df.loc[i, "Machine Name"] = str((db_result[i])[2])
                if (db_result[i])[9]:
                    df.loc[i, "Executor Status"] = "Online"
                else:
                    df.loc[i, "Executor Status"] = "Offline"
            df.to_excel(report_file, index=False)
            df1 = pd.DataFrame()
            df2 = pd.DataFrame()
            df1 = df[df['Executor Status'].str.contains('Online')]
            df2 = df[df['Executor Status'].str.contains('Offline')]
            template.DataFrame_Excelwriter([df1], report_file, sheetname="Online", sheet_to_create=False)
            template.DataFrame_Excelwriter([df2], report_file, sheetname="Offline", sheet_to_create=True)
            subject = template.Get_EmailSubject()
            body = Body = "Hi Team, \n\nPlease check executor and bot running status in the production machine.\n\n**Note --> "+str(pending_no)+"\n\nBy,\nRPA Support Team."
            obj_frame.immer_ai_Email(to,cc,subject,body,smtp,smtp_auth,authenticated=True,cc_email='', attachments=[report_file], folder_path=None)
        
        template.Complete_Success_Logging(framework=obj_frame)
    except Exception as error:
        issue = str(error)+str(traceback.format_exc())
        subject = Bot_Data+" - Failed"
        body = Body = "Hi Team, \n\nPlease check "+Bot_Data+" has not been completed.\n\n"+str(issue)+"\n\nBy,\nRPA Support Team."
        obj_frame.immer_ai_Email(to,cc,subject,body,smtp,smtp_auth,authenticated=True,cc_email='', attachments=[], folder_path=None)
        template.Complete_Failure_Logging(framework=obj_frame)

    try:
        obj_frame.immer_ai_Excelkill()
        template.Delete_Directories(folderlist=[Root])
    except:
        pass
    