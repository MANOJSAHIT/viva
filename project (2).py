import xlrd as xl    #for accessing excel files
import datetime      #for date and time settings
import pandas as pd  #for forming dataframes
import webbrowser    #for opening files directly from editor
import matplotlib as mat   #for plotting graphs
from matplotlib import pyplot as plt
import openpyxl as opx
path='C:\\time_series_covid_19_deaths.xlsx'
path1='C:\\time_series_covid_19_recovered.xlsx'
path2='C:\\time_series_covid_19_confirmed.xlsx'
path3='C:\\AgeGroupDetails.xlsx'
wb=xl.open_workbook(path)
wb1=xl.open_workbook(path1)
wb2=xl.open_workbook(path2)
wb3=xl.open_workbook(path3)
sheet=wb.sheet_by_index(0)
sheet4=wb1.sheet_by_index(0)
sheet2=wb2.sheet_by_index(0)
sheet3=wb3.sheet_by_index(0)
country=[]
ndeaths=[]
nrecover=[]
ncases=[]
totald=[]
totalr=[]
for i in range(1,sheet.nrows):
    country.append(sheet.cell_value(i,1).lower())
for i1 in range(1,sheet.nrows):
    tempc=[]
    tempc1=[0]
    tempc2=[]
    for j1 in range(4,sheet.ncols):
        tempc.append(int(sheet.cell_value(i1,j1)))
        tempc1.append(int(sheet.cell_value(i1,j1)))
        tempc2.append(abs(tempc[j1-4]-tempc1[j1-4]))
    ndeaths.append(tempc2)
    totald.append(sum(tempc2))
dicco={country[j2]:totald[j2] for j2 in range(len(country))}  #dictionary for total deaths
print(dicco['south sudan'])
dicdata={}
dates=[]
for l in range(4,sheet2.ncols):
    value=sheet2.cell_value(0,l)
    value1=datetime.datetime.strptime(value,'%m/%d/%y').strftime('%d-%m-%y')    #for date and time settings
    dates.append(value1)
f=open('total.txt','w+')
rgt=[]
totalc=[]
for i3 in range(1,sheet2.nrows):
    tempcs=[]
    tempcs1=[0]
    tempcs2=[]
    for j3 in range(4,sheet2.ncols):
        tempcs.append(int(sheet2.cell_value(i3,j3)))
        tempcs1.append(int(sheet2.cell_value(i3,j3)))
        tempcs2.append(abs(tempcs[j3-4]-tempcs1[j3-4]))
    totalc.append(sum(tempcs2))
    ncases.append(tempcs2)
dicdata=pd.DataFrame({country[k]:[totald[k],totalc[k]] for k in range(len(country))}).transpose()  #add here used pandas for dataframe here
dicdata.columns=['DEATHS','CONFIRMED']                                                             #change here coleective data for deaths,recovered,confirmed 
dicdata.to_csv('total.txt',header=True,index=True,sep='\t',mode='a')
for j4 in range(len(totalc)):
    if totalc[j4]==0 or totalc[j4]==1:
        rgt.append(0)
    else:
        rgt.append((totald[j4]/totalc[j4])*100)
red=[]
rgtr=[]
for j5 in range(len(country)):
    if rgt[j5]>=3:
        red.append(country[j5]) 
        rgtr.append(rgt[j5])
    else:
        pass
x=rgtr
y=red
didt={red[l4]:rgtr[l4] for l4 in range(len(red))}
didts=dict(sorted(dicco.items(),key=lambda kv:(kv[1],kv[0])))
print('THE TOP 15 COUNTRIES THAT ARE DANGEROUS TO VIST ARE \n')
lis=list(didts.keys())
lis=lis[-1:-16:-1]
print(lis)
dicdrt={country[hj]:rgt[hj] for hj in range(len(country))}
listdr=[]
for de in lis:
    listdr.append(dicdrt[de])
plt.tick_params(axis="x", labelsize=10)
plt.tick_params(axis="y", labelsize=10)
plt.xticks(rotation=90)
mat.pyplot.title('DEATH RATE IN TOP 15 HIGH RISK COUNTRIES',loc='center')
mat.pyplot.xlabel('COUNTRIES')
mat.pyplot.ylabel('NUMBERS')
plt.bar(lis,listdr,color=['red','blue','green'])
plt.show()
y1=country[::3]
x1=rgt[::3]
plt.tick_params(axis='x',labelsize=3)
plt.tick_params(axis='y',labelsize=10)
plt.xticks(rotation=90)
mat.pyplot.title('DEATH RATE IN AL  COUNTRIES\n',loc='center')
mat.pyplot.xlabel('COUNTRIES')
mat.pyplot.ylabel('NUMBERS')
plt.bar(y1,x1,color=['red','blue','green'])
plt.show()
dicrgt={country[k1]:ncases[k1] for k1 in range(len(country))}  #dictionary for datawise data
dicrgtd={country[k2]:ndeaths[k2] for k2 in range(len(country))} 
print('THE FOLLOWING GRAPHS SHOW NUMBER OF CASES & DEATHS\n')
name=input('ENTER THE NAME OF COUNTRY FOR DATE WISE GRAPH\n').lower()
x2=dates[::2]
while (1):
    if name in dicrgt.keys():
        break
    elif name=='america':
        name='us'
        break
    else:
        name=input('ENTER A VALID COUNTRY')
print(name)
h2=dicrgt[name]
h9=dicrgtd[name]
y2=h2[::2]
y9=h9[::2]
plt.tick_params(axis='x',labelsize=3)
plt.tick_params(axis='y',labelsize=10)
plt.xticks(rotation=90)
mat.pyplot.title(name,loc='center')
mat.pyplot.xlabel('DATES')
mat.pyplot.ylabel('NUMBER OF CASES')
plt.plot(x2,y2,label='CASES')
plt.legend()
plt.show()
plt.tick_params(axis='x',labelsize=3)
mat.pyplot.title(name,loc='center')
plt.plot(x2,y9,label='DEATHS')
mat.pyplot.ylabel('NUMBER OF DEATHS')
plt.xticks(rotation=90)
plt.legend()
plt.show()
f1=open('redzo.txt','w+')
for b2 in range(len(red)):
    f1.write(red[b2]+'\n')
f1.close()
p6=int(input('FOR A TEXT FILE CONTAINING LIST OF DANGEROUS COUNTRIES FOR NEXT TWO YEARS TO VISIT ENTER 1 FOR A NORMAL LIST IN TERMINAL ENTER 2\n'))
while(1):
    if p6==1:
        webbrowser.open('redzo.txt')  #opens text file containing redzone countries
        break
    elif p6==2:
        print(red)  #prints list of redzone countries in terminal
        break
    else:
        p6=int(input('ENTER A VALID INPUT\n'))
namec=input('ENTER THE COUNTRY NAME TO VIEW THE GRAPHICAL REPRESNTATION OF AVERAGE WEEKLY STATISTICS\n').lower()
while(1):
    if namec in country:
        break
    elif namec=='america':
        namec='us'
        break
    else:
        namec=input('THE ENTERED NAME IS NOT CORRECT TRY ANOTHER COUNTRY OR TRY ANOTHER NAME OF SAME COUNTRY\n')
nrecover=[]
totalr=[]
country=[]
for i in range(1,sheet4.nrows):
    country.append(sheet4.cell_value(i,1).lower())
    tempr=[]
    tempr1=[0]
    tempr2=[]
    for j in range(4,sheet4.ncols):
        tempr.append(sheet4.cell_value(i,j))
        tempr1.append(sheet4.cell_value(i,j))
        tempr2.append(tempr[j-4]-tempr1[j-4])
    nrecover.append(tempr2)
    totalr.append(sum(tempr2))
di={country[j11]:nrecover[j11] for j11 in range(len(country))}
tem=dicrgt[namec]
temd=dicrgtd[namec]
temr=di[namec]
avcases=[]
avdeaths=[]
avrecover=[]
b=0
for we in range(sheet.ncols//7):
    jsum=0
    jsum1=0
    jsum2=0
    while(we*7<=b<=((we*7)+6)):
        jsum=jsum+tem[b]
        jsum1=jsum1+temd[b]
        jsum2=jsum2+temr[b]
        b+=1
    avcases.append(jsum/7)
    avdeaths.append(jsum1/7)
    avrecover.append(jsum2/7)
x3=[]
for gr in range(sheet.ncols//7):
    x3.append('WEEK-'+str(gr+1))
y3=avcases
y4=avdeaths
y5=avrecover
mat.pyplot.title(namec,loc='center')
mat.pyplot.xlabel('DATES')
mat.pyplot.ylabel('AVERAGE VALUES OF CASES')
plt.xticks(rotation=90)
plt.plot(x3,y3)
plt.show()
mat.pyplot.title(namec,loc='center')
plt.xticks(rotation=90)
plt.plot(x3,y4)
mat.pyplot.ylabel('AVERAGE VALUE OF DEATHS')
plt.show()
mat.pyplot.title(namec,loc='center')
plt.xticks(rotation=90)
plt.plot(x3,y5)
mat.pyplot.ylabel('AVERAGE VALUE OF RECOVERIES')
plt.show()
p7=int(input('FOR CHECKING A COUNTRY IS IN DANGER LIST OR NOT ENTER 1 ELSE ENTER 2'))
while(1):
    if p7==1:
        nam=input('ENTER THE NAME OF COUNTRY TO SEARCH\n').lower()
        if nam in red:
            print('THE ENTERED COUNTRY IS DANGER COUNTRY\n')
        else:
            print('THE ENTERED COUNTRY IS NOT A DANGER COUNTRY\n')
        break
    elif p7==2:
        break
    else:
        p7=int(input('ENTER A VALID ENTRY\n'))
        break
ages=[]
per=[]
for g in range(2,sheet3.nrows):
    if sheet3.cell_value(g,1)=='':
        pass
    else:
        ages.append(str(sheet3.cell_value(g,1)))
for g1 in range(2,sheet3.nrows):
    if sheet3.cell_value(g1,3)=='':
        pass
    else:
        per.append((sheet3.cell_value(g1,3)))
plt.title('AGES AND PERCENTAGE OF DEATHS IN THAT AGE')
plt.pie(per,labels=ages,shadow=True,startangle=180)
plt.axis('equal')
plt.show()
print('THE TOP 15 COUNTRIES DANGEROUS TO VISIT ARE \n')
print(lis)