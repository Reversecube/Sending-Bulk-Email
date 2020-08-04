####################################
# @author	Reversecube            #
# @author	Lance Ziyat            #
# @copyright	Copyright (c) 2020 #
# @license	Reversecube            #
# @link	http://reversecube.site    #
####################################
import smtplib as sm
import pandas as pd
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

class Sending:
    def __init__(self,email='',password='',myhost='smtp.gmail.com',myport=587):
        self.email=email
        self.password=password
        self.myhost=myhost
        self.myport=myport

    def Send(self,subject_,from_,to_addr,content):

        msg=MIMEMultipart()
        msg['Subject']=subject_
        msg['From']=from_
        msg['To']=to_addr
        #msg.set_payload(content, "UTF-8")
        msg.attach(MIMEText(content,"html","UTF-8"))
        try:
            #mail=sm.SMTP('10.18.93.128:25')
            #context = sm.ssl.create_default_context()
            mail=sm.SMTP(self.myhost,self.myport)
            print('Connect To Server :'+self.myhost+'/'+str(self.myport))
            #mail.ehlo()
            mail.starttls()
            #time.sleep(3)
            try:
                mail.login(self.email, self.password)
                print('Connect To Compte '+self.email+'/'+self.password)
            except:
                self.field_list(self.email+','+self.password+','+self.myhost+','+str(self.myport))      
            mail.sendmail(self.email, to_addr,msg.as_string())
            print('done')            
            mail.quit()
            return True
        except:
             print "-------------------------Msg is not Send-----------------"
             return False
    

class Excel:
    def __init__(self):
        pass

    def load_data(self):
        email_list=[]
        data = pd.read_csv("data.csv")
        email_list = data.to_numpy()
        # email_list=data.head()
        # email_list=data.tail()
        print("Data Was Loaded Successfuly !!")
        return email_list
    def load_account(self):        
        account = pd.read_excel("accounts.xlsx")
        account_list=account.to_numpy()
        # email_list=data.head()
        # email_list=data.tail()
        print("Account List Was Loaded Successfuly !!")
        return account_list
    def load_creative(self):
        creative = pd.read_excel("creatives.xlsx")
        creative_list=creative.to_numpy()
        # email_list=data.head()
        # email_list=data.tail()
        print("Creative List Was Loaded Successfuly !!")
        return creative_list


def main():
    excel = Excel()
    #!-- Get Data From Data.csv
    lsdata = excel.load_data()
    #!--Get Account From Excel File
    lsaccount = excel.load_account()
    #!--Get Creative From Excel File
    lscreative = excel.load_creative()
    #print(lscreative)
    s=Sending()
    print("------------------------------------Welcome To Sending-Bluk-Email---------------------------------")
    if(len(lscreative)>0):
        fn = int(input("Chose Creative By Number:"))
        opmax = int(input("SMTP Rotation (MAX 500 Per Day):"))

        subject_ = lscreative[fn-1][0]
        from_ = lscreative[fn-1][1]
        content_ = lscreative[fn-1][2]
        opmax = int(input("SMTP Rotation (MAX 500 Per Day):"))

        for index in xrange(0,len(lsaccount)):
            print("#################### --"+str(index+1)+"-- #############################")

            print('Using Email : '+lsaccount[index][0]+" Number "+str(index+1))
            s=Sending(lsaccount[index][0],lsaccount[index][1],'smtp.gmail.com',587)
            print("Port : 587")
            op=0
            mindex=0
            while (mindex<len(lsdata) and op<=opmax):
            
                print("Email Number "+str(mindex+1)+" withe ----->"+lsaccount[index][0]+" Sending to -+-+-+-+"+lsdata[mindex])
                if(s.Send(subject_,from_,lsdata[mindex],content_)==True):
                    #time.sleep(5)
                    print(lsdata[mindex]+" send it.")
                    mindex=mindex+1
                    op=op+1
                else:
                    break
    print("----------------------------------Done--------------------------------------------------")
    print("----------------------------------------------------------------------------------------")
if __name__ == '__main__':
    main()
