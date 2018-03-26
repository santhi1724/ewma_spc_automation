import gspread
from oauth2client.service_account import ServiceAccountCredentials
import sys
import sysconfig
import time
import pandas as pd
import csv
import numpy as np
import matplotlib.pyplot as plt
import plot
from matplotlib.pylab import rcParams
rcParams['figure.figsize'] = 100, 6
from io import StringIO
import matplotlib.pyplot as plt
import spc
import smtplib
from email.MIMEMultipart import MIMEMultipart
from email.MIMEText import MIMEText
from email.MIMEBase import MIMEBase
from email import Encoders
from decimal import Decimal



scope = ['https://spreadsheets.google.com/feeds']
creds = ServiceAccountCredentials.from_json_keyfile_name('client_secret.json', scope)
client = gspread.authorize(creds)

sheet = client.open("SPC_allherds.csv").sheet1

live_data = sheet.get_all_values()
live_data = pd.DataFrame(live_data)

live_data.to_csv('livedata.csv')


def calculate_spc(farm_name):
    def farm(farm_name):
        with open('livedata.csv') as csvfile:
            readCSV = csv.reader(csvfile, delimiter=',')
            Farm_Aborts = []
            Farm_Week = []
            Farm_PWM = []
            Farm_Alive = []
            for row in readCSV:
                if(row[1] == farm_name):
                    A1 = row[4]
                    D1 = row[2]
                    P1 = row[5]
                    L1 = row[11]
                    Farm_Aborts.append(A1)
                    Farm_Week.append(D1)
                    Farm_PWM.append(P1)
                    Farm_Alive.append(L1)
            df1 = pd.DataFrame({'date':Farm_Week})
            for row in df1:
                df1 = df1.apply(pd.to_datetime)

            k = np.array(Farm_Aborts).astype(np.int)
            l = [float(x) for x in Farm_PWM]
            m = [float(y) for y in Farm_Alive]
            #print l
            return df1,k,l,m

    def ewma(parameter_values):
        df2 = pd.DataFrame({'data':parameter_values})
        df3 = df2.ewm(com=None, span=None, halflife=None,
                      alpha=0.4, min_periods=3, freq=None,
                      adjust=True, ignore_na=False, axis=0).mean()
        return df3
       
    def spc_limits(paramter_values, lcl_value,ucl_value,center_value):
        lcl = []
        ucl = []
        center = []
        s = len(paramter_values)
        for i in xrange(s):    
            l = lcl_value
            lcl.append(l)
        for i in xrange(s):    
            u = ucl_value
            ucl.append(u)
        for i in xrange(s):    
            c = center_value
            center.append(c)

        return lcl, ucl, center

    def spcplot(ewmaofparameter_values, week_farm):
        plt.clf()
        colours = ['c', 'crimson', 'chartreuse']
        plt.plot(week_farm, ewmaofparameter_values)
        plt.xlabel('time')
        plt.ylabel('EWMA')
        plt.title('Plot of EWMA SPC for given parameter')
        plt.plot(week_farm, lcl, c = 'c')
        plt.plot(week_farm, ucl, c = 'crimson')
        plt.plot(week_farm, center, c = 'chartreuse')
        # Creates 3 Rectangles
        p1 = plt.Rectangle((0, 0), 0.1, 0.1, fc=colours[0])
        p2 = plt.Rectangle((0, 0), 0.1, 0.1, fc=colours[1])
        p3 = plt.Rectangle((0, 0), 0.1, 0.1, fc=colours[2])
        # Adds the legend into plot
        plt.legend((p1, p2, p3), ('lcl', 'ucl', 'center'), loc='best')
	fig = plt.figure(figsize=(18, 18))
        plt.savefig('ewma_spc.png')

        return plt

    def test_violating_runs(ewmaofparamter_values,week_period,lcl_value,ucl_value,center_value,farm,parameter):

        ewma = ewmaofparamter_values['data']
        week = week_period['date']
        l = len(ewmaofparamter_values)
        A1 = []
        B1 = []
        C1 = []
        D1 = []
        for i in xrange(l):
            if ewma[i] >= ucl_value:
                #print i
                #A1.append(i)
                B1.append(ewma[i])
                C1.append(week[i])
                D1.append(farm)
                A1.append(parameter)
            #if data[i] <= lcl:
                #print i
                #C1.append(i)
                #D1.appned(data[i])

        return B1,C1, D1, A1

    def violation_list(violations_values):
        B2 = violations_values[0]
        C2 = violations_values[1]
        D2 = violations_values[2]
        A2 = violations_values[3]
        df3 = pd.DataFrame({'violations_ucl': B2, 'date': C2, 'farm': D2, 'parameter':A2})
        return df3
    
    def automatic_alerts(from_email_address,to_email_address,violationsCsvfile,reciever_email_id1,reciever_email_id2):
        SUBJECT = "SPC of BR_Aborts violations_list"

        msg = MIMEMultipart()
        msg['Subject'] = SUBJECT 
        msg['From'] = from_email_address
        msg['To'] = to_email_address
        body = "Weekly status of EWMA-SPC violating points."
        msg.attach(MIMEText(body, 'plain'))

        part = MIMEBase('application', "octet-stream")
        part.set_payload(open(violationsCsvfile, "rb").read())
        Encoders.encode_base64(part)

        part.add_header('Content-Disposition', 'attachment; filename=violations_list.csv')
        #part.add_header('Content-Disposition', 'attachment; filename="test_pd_ewma_pd.png"')

        msg.attach(part)

        mail = smtplib.SMTP('smtp.gmail.com',587)
            #server = smtplib.SMTP(self.EMAIL_SERVER)
        #server.sendmail("bsanthi17@gmail.com", "bsanthi17@gmail.com", msg.as_string())

        mail.ehlo()

        mail.starttls()

        mail.login('bsanthi17@gmail.com', 'ABCabc123@')

        mail.sendmail(reciever_email_id1,reciever_email_id2,msg.as_string())

        mail.close()

    def automatic_alerts_graph(from_email_address,to_email_address,Violations_graph,reciever_email_id1,reciever_email_id2):
        SUBJECT = "SPC of BR_Aborts violations_list"

        msg = MIMEMultipart()
        msg['Subject'] = SUBJECT 
        msg['From'] = from_email_address
        msg['To'] = to_email_address
        body = "Weekly status of EWMA-SPC Graph."
        msg.attach(MIMEText(body, 'plain'))

        part = MIMEBase('application', "octet-stream")
        part.set_payload(open(Violations_graph, "rb").read())
        Encoders.encode_base64(part)

        #part.add_header('Content-Disposition', 'attachment; filename=violations_list.csv')
        part.add_header('Content-Disposition', 'attachment; filename="test_pd_ewma_pd.png"')

        msg.attach(part)

        mail = smtplib.SMTP('smtp.gmail.com',587)
        #server = smtplib.SMTP(self.EMAIL_SERVER)
        #server.sendmail("bsanthi17@gmail.com", "bsanthi17@gmail.com", msg.as_string())

        mail.ehlo()

        mail.starttls()

        mail.login('bsanthi17@gmail.com', 'ABCabc123@')

        mail.sendmail(reciever_email_id1,reciever_email_id2,msg.as_string())

        mail.close()

    with open('Control limits.csv') as csvfile:
        readCSV = csv.reader(csvfile, delimiter=',')
        for row in readCSV:
            if(row[0] == farm_name):
                Week, Aborts, farm_PWM, farm_Alive = farm(farm_name)
                ewma_Aborts = ewma(Aborts)
                lcl, ucl, center = spc_limits(Aborts, row[2], row[3], row[4])
                spcplot(ewma_Aborts,Week)
                if (row[1] == "Abortions"):
                    a = float(row[2])
                    b = float(row[3])
                    c = float(row[4])
                    #print a, b, c
                    ewma_violations = test_violating_runs(ewma_Aborts, Week, a, b, c, farm_name,row[1])
                    violations_ucl = violation_list(ewma_violations)
                    
                    violations_ucl.to_csv('violations_ucl_aborts.csv')
                    automatic_alerts("bsanthi17@gmail.com",row[6],"violations_ucl.csv",row[6],row[5])
                    automatic_alerts_graph("bsanthi17@gmail.com",row[6],"ewma_spc.png",row[5],row[6])
                    
                ewma_PWM = ewma(farm_PWM)
                #print ewma_PWM                
                lcl, ucl, center = spc_limits(farm_PWM, row[2], row[3], row[4])
                #print lcl               
                spcplot(ewma_PWM,Week)
                if (row[1] == "PWM"):
                    a = float(row[2])
                    b = float(row[3])
                    c = float(row[4])
                    #print a, b, c
                    ewma_violations = test_violating_runs(ewma_PWM,Week, a, b, c, farm_name,row[1])
                    #print ewma_violations
                    violations_ucl = violation_list(ewma_violations)
                    #print violations_ucl
                    violations_ucl.to_csv('violations_ucl_PWM.csv') 
                    automatic_alerts("bsanthi17@gmail.com",row[6],"violations_ucl.csv",row[6],row[5])
                    automatic_alerts_graph("bsanthi17@gmail.com",row[6],"ewma_spc.png",row[5],row[6])

                ewma_Alive = ewma(farm_Alive)
                #print ewma_Alive     
                lcl, ucl, center = spc_limits(farm_Alive, row[2], row[3], row[4])
                #print lcl               
                spcplot(ewma_Alive,Week)
                if (row[1] == "Prenatal losses"):
                    a = float(row[2])
                    b = float(row[3])
                    c = float(row[4])
                    #print a, b, c
                    ewma_violations = test_violating_runs(ewma_Alive,Week, a, b, c, farm_name,row[1])
                    #print ewma_violations
                    violations_ucl = violation_list(ewma_violations)
                    #print violations_ucl
                    violations_ucl.to_csv('violations_ucl_Pl.csv') 
                    automatic_alerts("bsanthi17@gmail.com",row[6],"violations_ucl.csv",row[5],row[6])
                    automatic_alerts_graph("bsanthi17@gmail.com",row[6],"ewma_spc.png",row[5],row[6])


with open('Control limits.csv') as csvfile:
    rows = csv.reader(csvfile, delimiter=',')
    newrows = []
    for row in rows:
        if row[0] not in newrows:
            newrows.append(row[0])
for i in range(len(newrows)):
    calculate_spc(newrows[i])

    
