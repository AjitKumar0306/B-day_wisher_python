# Turn on less secure app of your Account

import pandas as pd
import datetime 
import smtplib

# put your email it details
Gmail_ID = ''
Gmail_PSWD = ''

def sendEmail(to, sub, msg):
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()
    server.login(Gmail_ID, Gmail_PSWD)
    server.sendmail(Gmail_ID, to, f"Subject: {sub}\n\n{msg}")
    server.quit()




if __name__ == "__main__":
    df = pd.read_excel("B'dayData.xls")
    today = datetime.datetime.now().strftime("%d-%m")
    yearNow = datetime.datetime.now().strftime("%Y")
    writeInd = []
    for index, item in df.iterrows():
       bday = item['Birthday'].strftime("%d-%m")
       if (today == bday ) and (yearNow not in str(item['Year'])):
           sendEmail(item['Email'], "Happy Birthday", item["Dialogue"]) 
           writeInd.append(index) 
    
    
    if len(writeInd):
        for i in writeInd:
            year = df.loc[i, "Year"] 
            df.loc[i, "Year"] = str(year) + ', '+ str(yearNow)    
        df.to_excel("B'dayData.xls",index=False)   
        

    