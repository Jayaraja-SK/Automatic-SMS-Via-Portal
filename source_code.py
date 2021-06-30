import requests
import openpyxl

wb=openpyxl.load_workbook('samplemarklist.xlsx')
sheet=wb['Sheet1']

a=sheet['B'] #Phone Number
L1=[]
for cell in a:
    n=cell.value
    L1.append(n)
print(L1)

b=sheet['A'] # Student Name
L2=[]
for cell in b:
    m=cell.value
    L2.append(m)
print(L2)

c=sheet['C'] #Physics Marks
L3=[]
for cell in c:
    p=cell.value
    L3.append(p)
print(L3)

d=sheet['D'] #Chemistry Marks
L4=[]
for cell in d:
    q=cell.value
    L4.append(q)
print(L4)

e=sheet['E'] #Maths Marks
L5=[]
for cell in e:
    r=cell.value
    L5.append(r)
print(L5)

f=sheet['F'] #E.G.
L6=[]
for cell in f:
    s=cell.value
    L6.append(s)
print(L6)

f=open('passcode.txt','r')#READS THE API KEY FROM A .TXT FILE
auth_code=f.read()

for i in range(1,len(L1)):
    url = "https://www.fast2sms.com/dev/bulk"
    payload = "sender_id=FSTSMS&message=\nUnit Test-2 \nName:"+L2[i]+",\nPhysics:"+str(L3[i])+",\nChemistry:"+str(L4[i])+",\nMaths:"+str(L5[i])+",\nE.G.:"+str(L6[i])+".&language=english&route=p&numbers="+str(L1[i])
    headers = {

    'authorization': auth_code,

    'Content-Type': "application/x-www-form-urlencoded",

    'Cache-Control': "no-cache",

     }
    response = requests.request("POST", url, data=payload, headers=headers)
    print(response.text)
    

    


