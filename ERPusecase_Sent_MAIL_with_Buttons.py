'''
Author   :  AnurajPilanku
Use Case :  ERP Quality Review

'''
import os
import datetime
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email.mime.application import MIMEApplication
from os.path import basename
import sys

filepath = ""
today = str(datetime.date.today())
colureddata="Please click the button to navigate to {chnge}!"

From = 'USSACPrd@mmm.com'
To = sys.argv[1]
cc =sys.argv[2]
bcc = sys.argv[3]
subject= "ERP Quality Review Macro Execution Output Mail {dte}".format(dte=str(today))
greetings = "Hi Team"
body= "ERP Incident Quality Review for {dte} has successfully Completed".format(dte=str(today))
signature = 'CAC Automation Centre'
fontStyle = 'Times New Roman'
def stle(dt):
    styledfilename = '''<span style="color:#6C6E7A">{filename}</span>'''.format(filename=dt)
    return styledfilename
mout=stle(colureddata.format(chnge="ERP Macro Output"))
sp=stle(colureddata.format(chnge="sharepoint"))


def sentmail():
    html_file = '''

    <!DOCTYPE html>
    <html>
    <head>
    </head>
    <body>
      <h1></h1>
      <body style="font-family:{fontStyle}">
      <br/><img src='cid:image1'<br/>
      <br>
      <br>
      <br /><font face='{fontStyle}'>''' + greetings + ''',</font><br/>

      <br /><font face='{fontStyle}'>''' + body + ''' </font><br/>

    <div style="overflow-x:auto;">

    </div>
    <br>
    <br /><font face='{fontStyle}'>''' + mout + ''' </font><br/>
    <br>
  
    <div style="font-family: Helvetica, sans-serif;font-size: 100%;margin: 0;padding: 0;margin-top: 0">  
    <v:rect style="height:40px;width:0;" fill="f" stroke="f" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:w="urn:schemas-microsoft-com:office:word" />   
    <v:roundrect style="width:490px;height:55px;position:relative;top:0;left:-4px;" arcsize="50%" stroke="f" fill="true" xmlns:v=&quot;urn:schemas-microsoft-com:vml&quot; xmlns:w=&quot;urn:schemas-microsoft-com:office:word&quot;>
        <v:fill type="gradientradial" color="transparent" color2="#112EED" focus="0" focusposition=".05,0.23" focussize=".9,0.25" />
    </v:roundrect>
        
    <v:roundrect href="\\\\acprd01\\E\\3M_CAC\\ERP_Quality_Review\\QualityCheck" style="width:220px;height:40px;position:relative;top:0;left:0;v-text-anchor:middle;" arcsize="50%" stroke="f" fillcolor="#1659DE" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:w="urn:schemas-microsoft-com:office:word">
        <w:anchorlock/>
        <v:textbox inset="0,0,0,0">
        <center>
    <a href="\\\\acprd01\\E\\3M_CAC\\ERP_Quality_Review\\QualityCheck" style="font-family: Helvetica, sans-serif;font-size: 14px;margin: 0;padding: 0 24px;color: #ffffff;margin-top: 0;font-weight: bold;line-height: 40px;letter-spacing: 0.1ch;text-transform: uppercase;background-color: #1659DE;border-radius: 5000px;box-shadow: 0 8px 14px 2px #f45c2324, 0 6px 20px 5px #f45c231f, 0 8px 10px -5px #f45c2333;display: inline-block;text-align: center;text-decoration: none;white-space: nowrap;-webkit-text-size-adjust: none">Macro Output</a>
        </center>
        </v:textbox>
    </v:roundrect>
    </div>

     
     <br>
    <br /><font face='{fontStyle}'>''' + sp + ''' </font><br/>
    <br>
  
    <div style="font-family: Helvetica, sans-serif;font-size: 100%;margin: 0;padding: 0;margin-top: 0">  
    <v:rect style="height:40px;width:0;" fill="f" stroke="f" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:w="urn:schemas-microsoft-com:office:word" />   
    <v:roundrect style="width:490px;height:55px;position:relative;top:0;left:-4px;" arcsize="50%" stroke="f" fill="true" xmlns:v=&quot;urn:schemas-microsoft-com:vml&quot; xmlns:w=&quot;urn:schemas-microsoft-com:office:word&quot;>
        <v:fill type="gradientradial" color="transparent" color2="#112EED" focus="0" focusposition=".05,0.23" focussize=".9,0.25" />
    </v:roundrect>
        
    <v:roundrect href="microsoft-edge:https://skydrive3m.sharepoint.com/:f:/r/teams/ProjectAlpineExecution/Shared%20Documents/Service%20Management%20Office/QualityTools/Quality%20Review%20-%20Files%20to%20Audit?csf=1&web=1&e=mjtKPr" style="width:220px;height:40px;position:relative;top:0;left:0;v-text-anchor:middle;" arcsize="50%" stroke="f" fillcolor="#964B00" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:w="urn:schemas-microsoft-com:office:word">
        <w:anchorlock/>
        <v:textbox inset="0,0,0,0">
        <center>
    <a href="microsoft-edge:https://skydrive3m.sharepoint.com/:f:/r/teams/ProjectAlpineExecution/Shared%20Documents/Service%20Management%20Office/QualityTools/Quality%20Review%20-%20Files%20to%20Audit?csf=1&web=1&e=mjtKPr" style="font-family: Helvetica, sans-serif;font-size: 14px;margin: 0;padding: 0 24px;color: #ffffff;margin-top: 0;font-weight: bold;line-height: 40px;letter-spacing: 0.1ch;text-transform: uppercase;background-color: #964B00;border-radius: 5000px;box-shadow: 0 8px 14px 2px #f45c2324, 0 6px 20px 5px #f45c231f, 0 8px 10px -5px #f45c2333;display: inline-block;text-align: center;text-decoration: none;white-space: nowrap;-webkit-text-size-adjust: none">Sharepoint</a>
        </center>
        </v:textbox>
    </v:roundrect>
    </div>
   
  
    <br /><font face='{fontStyle}'>Regards </font><br/>
    <br /><font face='{fontStyle}'>''' + signature + '''</font><br/>
    <br>
    <br>
    <br/><img src='cid:image3'<br/>
    </div>
    </body>
    </html>
    '''.format(fontStyle=fontStyle,excelpath=filepath)
    msgRoot = MIMEMultipart('related')
    msgRoot['Subject'] = subject
    msgRoot['From'] = From
    msgRoot['Cc'] = cc
    msgRoot['To'] = To
    msgRoot['Bcc'] = bcc
    msgRoot.preamble = '====================================================='
    msgAlternative = MIMEMultipart('alternative')
    msgRoot.attach(msgAlternative)
    msgText = MIMEText('Please find ')
    msgAlternative.attach(msgText)
    msgText = MIMEText(html_file, 'html')
    msgAlternative.attach(msgText)
    msgAlternative.attach(msgText)
    fp = open(r"\\acprd01\E\3M_CAC\Alphine\Mail_image\head.png", 'rb')
    # fp2 = open(sys.argv[7], 'rb')#"//acdev01/3M_CAC/IPM_FSM/Mail_elements/new.png"
    fp3 = open(r"\\acprd01\E\3M_CAC\Alphine\Mail_image\footer.png", 'rb')
    msgImage = MIMEImage(fp.read())
    # msgImage1 = MIMEImage(fp2.read())
    msgImage2 = MIMEImage(fp3.read())
    fp.close()
    fp3.close()
    msgImage.add_header('Content-ID', '<image1>')
    msgImage2.add_header('Content-ID', '<image3>')
    msgRoot.attach(msgImage)
    msgRoot.attach(msgImage2)
    smtp = smtplib.SMTP()
    smtp.connect('mailserv.mmm.com')
    # smtp.sendmail(From,To, msgRoot.as_string())
    smtp.send_message(msgRoot)
    smtp.quit()
    print("Email is sent successfully")

sentmail()

