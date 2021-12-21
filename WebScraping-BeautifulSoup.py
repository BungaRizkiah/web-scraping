from bs4 import BeautifulSoup
import requests
import pandas as pd
import datetime
from selenium import webdriver

page_link = 'http://lefthandditchcompany.com/SystemStatus.aspx'
page_response = requests.get(page_link, timeout=60, verify=False)
body = BeautifulSoup(page_response.content, 'lxml')
Creekflow = body.find("span", id="ctl00_MainContentPlaceHolder_CreekFlowAFLabel").get_text()
CFS = body.find("span", id="ctl00_MainContentPlaceHolder_CreekFlowCFSLabel").get_text()
Issues = body.find("span", id ="ctl00_MainContentPlaceHolder_CreekFlowIssueCFSPerShareLabel").get_text()
Current_Gold = body.find("span", id='ctl00_MainContentPlaceHolder_GoldAFLabel').get_text()
Current_Isabelle = body.find("span", id='ctl00_MainContentPlaceHolder_IsabelleAFLabel').get_text()
Current_LHP = body.find("span", id="ctl00_MainContentPlaceHolder_LHParkAFLabel").get_text()
Current_LHV = body.find("span", id="ctl00_MainContentPlaceHolder_LHValleyAFLabel").get_text()
Current_Allens = body.find("span", id='ctl00_MainContentPlaceHolder_AllensAFLabel').get_text()
Current_Total = body.find("span", id='ctl00_MainContentPlaceHolder_TotalAFLabel').get_text()
Empty_Gold = body.find("span", id="ctl00_MainContentPlaceHolder_GoldEmptyAFLabel").get_text()
Empty_Isabelle = body.find("span", id="ctl00_MainContentPlaceHolder_IsabelleEmptyAFLabel").get_text()
Empty_LHP = body.find("span",id="ctl00_MainContentPlaceHolder_LHParkEmptyAFLabel").get_text()
Empty_LHV = body.find("span", id="ctl00_MainContentPlaceHolder_LHValleyEmptyAFLabel").get_text()
Empty_Allens = body.find("span", id="ctl00_MainContentPlaceHolder_AllensEmptyAFLabel").get_text()
Full_Gold = body.find("span", id="ctl00_MainContentPlaceHolder_GoldFullAFLabel").get_text()
Full_Isabelle = body.find("span", id="ctl00_MainContentPlaceHolder_IsabelleFullAFLabel").get_text()
Full_LHP = body.find("span", id="ctl00_MainContentPlaceHolder_LHParkFullAFLabel").get_text()
Full_LHV = body.find("span", id='ctl00_MainContentPlaceHolder_LHValleyFullAFLabel').get_text()
Full_Allens = body.find("span", id="ctl00_MainContentPlaceHolder_AllensFullAFLabel").get_text()



dictionary = {'Creekflow': Creekflow, 'CFS': CFS, 'Issues': Issues,
              'CurrentGold': Current_Gold, 'CurrentIsabelle': Current_Isabelle, 'CurrentLHP': Current_LHP,
              'CurrentLHV': Current_LHV, 'CurrentAllens': Current_Allens, 'CurrentTotal': Current_Total,
              'EmptyGold': Empty_Gold, 'EmptyIsabelle': Empty_Isabelle, 'EmptyLHP': Empty_LHP,
              'EmptyLHV': Empty_LHV, 'EmptyAllens': Empty_Allens, 'FullGold': Full_Gold,
              'FullIsabelle': Full_Isabelle, 'FullLHP': Full_LHP, 'FullLHV': Full_LHV, 'FullAllens': Full_Allens}

df = pd.DataFrame(dictionary, index=[0])
df.CFS = df.CFS.str.replace('(','')
df.CFS = df.CFS.str.replace(')','')
df.CFS = df.CFS.str.replace('CFS','')
df.Issues = df.Issues.str.replace('(','')
df.Issues = df.Issues.str.replace(')','')
df.Issues = df.Issues.str.replace('Issue:','')
df.Issues = df.Issues.str.replace('CFS / Share','')
df.CurrentLHV = df.CurrentLHV.str.replace(',','')
df.CurrentTotal = df.CurrentTotal.str.replace(',','')
df.FullLHP = df.FullLHP.str.replace(',','')
df.FullLHV = df.FullLHV.str.replace(',','')
df.FullAllens = df.FullAllens.str.replace(',','')

df.Creekflow = pd.to_numeric(df.Creekflow)
df.CFS = pd.to_numeric(df.CFS)
df.Issues = pd.to_numeric(df.Issues)
df.CurrentGold = pd.to_numeric(df.CurrentGold)
df.CurrentIsabelle = pd.to_numeric(df.CurrentIsabelle)
df.CurrentLHP = pd.to_numeric(df.CurrentLHP)
df.CurrentLHV = pd.to_numeric(df.CurrentLHV)
df.CurrentAllens = pd.to_numeric(df.CurrentAllens)
df.EmptyGold = pd.to_numeric(df.EmptyGold)
df.EmptyIsabelle = pd.to_numeric(df.EmptyIsabelle)
df.EmptyLHP = pd.to_numeric(df.EmptyLHP)
df.EmptyLHV = pd.to_numeric(df.EmptyLHV)
df.EmptyAllens = pd.to_numeric(df.EmptyAllens)
df.FullGold = pd.to_numeric(df.FullGold)
df.FullIsabelle = pd.to_numeric(df.FullIsabelle)
df.FullLHP = pd.to_numeric(df.FullLHP)
df.FullLHV = pd.to_numeric(df.FullLHV)
df.FullAllens= pd.to_numeric(df.FullAllens)

#add datetime
today = datetime.datetime.today()
Position = 19
df.insert(loc=Position, column='Date Refreshed', value=today)

#merge data with the earlier data
df_origin= pd.read_excel(r'C:\Users\Anti.Rizkiah\Desktop\Farm101.xlsx', sheet_name='Farm101')
df_new = df_origin.append(df)

writer = pd.ExcelWriter(r'C:\Users\Anti.Rizkiah\Desktop\Farm101.xlsx', engine='xlsxwriter')
df_new.to_excel(writer, sheet_name='Farm101', index=False)
writer.save()