import xlrd as xl
import openpyxl
import pandas as pd
import webbrowser
from matplotlib import pyplot as plt
def admin():
    years=[]
    statenamef=[]
    seasonf=[]
    cropf=[]
    areaf=[]
    productionf=[]
    i1=0
    n1=int(input('TO SEARCH BY YEAR ENTER 1 VIEW GRAPH OF CROPS PRODUCTION DETAILS ENTER 2 TO ADD DEATAILS ENTER 3\n'))
    while(1):
        if (n1==1):
            searyear=int(input('ENTER THE YEAR TO SEARCH IN DATA YEAR BETWEEN 1997 TO 2006 AND YOU CAN ALSO ENTER 2010\n'))
            while(i1<len(year)):
                if year[i1]==searyear:
                    years.append(year[i1])
                    statenamef.append(statename[i1])
                    seasonf.append(season[i1])
                    cropf.append(crop[i1].lower())
                    areaf.append(area[i1])
                    productionf.append(production[i1])
                    i1+=1
                else:
                    i1+=1
            d=pd.DataFrame(list(zip(years,statenamef,cropf,seasonf,areaf,productionf)),columns=['year','statename','crop','season','area','production'])
            d.to_excel('extd.xlsx',header=True)
            webbrowser.open('extd.xlsx')
            break
        elif (n1==2):
            print('VIEW THE LIST OF CROPS AND ENTER THE NAME OF THAT CROP\n')
            cropf=list(set(crop))
            print(cropf)
            prodf=[]
            yeaf=[]
            i2=0
            searcrop=input('ENTER THE NAME OF CROP FOR YEAR WISE PRODUCTION GRAPH').lower()
            if searcrop=='coconut':
                searcrop='coconut '
            while(i2<len(crop)):
                if crop[i2]==searcrop:
                    yeaf.append(year[i2])
                    prodf.append(production[i2])
                    i2+=1
                else:
                    i2+=1
            dicd={prodf[i3]:yeaf[i3] for i3 in range(len(yeaf))}
            dicd=dict(sorted(dicd.items(),key=lambda kv:(kv[1])))
            ndicd={}
            for key,value in dicd.items():
                if value in ndicd:
                    ndicd[value].append(key)
                else:
                    ndicd[value]=[key]
            for j in ndicd.keys():
                if len(ndicd[j])>=1:
                    ndicd[j]=float(sum(ndicd[j]))
            x=[]
            y=[]
            for i4 in ndicd.keys():
                x.append(int(i4))
                y.append(ndicd[i4])
            plt.plot(x,y,label='PRODUCTION')
            plt.locator_params(nbins=len(x)+2)
            plt.show()
            break
        elif(n1==3):
            name=input('enter the name of state\n')
            dname=input('enter the name of district\n')
            ye=int(input('enter the year\n'))
            sea=input('enter the season\n')
            cro=input('enter the name of crop\n')
            are=float(input('enter the area of crop\n'))
            rpr=float(input('enter the production of crop\n'))
            ln=[name,dname,ye,sea,cro,are,rpr]
            xfile=openpyxl.load_workbook(path)
            ws=xfile.active
            for j in range(1,8):
                ws.cell(row=sheet.nrows+1,column=j).value=ln[j-1]
            xfile.save('CropsDataFile1.xlsx')
            webbrowser.open('CropsDataFile1.xlsx')
            return ln
            break
        else:
            n1=int(input('select valid input\n'))
def user(f):
    n=int(input('enter 1 to see the data quit enter 2'))
    if n==1:
        if f!=None:
            statename.append(f[0])
            year.append(f[2])
            season.append(f[3])
            crop.append(f[4])
            area.append(f[5])
            production.append(f[6])
            d1=pd.DataFrame(list(zip(year,statename,crop,season,area,production)),columns=['year','statename','crop','season','area','production'])
            d1.to_csv('file.txt',sep='\t',header=True,index=False)
            webbrowser.open('file.txt')
        else:
            d1=pd.DataFrame(list(zip(year,statename,crop,season,area,production)),columns=['year','statename','crop','season','area','production'])
            d1.to_csv('file.txt',sep='\t',header=True,index=False)
            webbrowser.open('file.txt')
path='D:\\CropsDataFile.xlsx'
wb=xl.open_workbook(path)
sheet=wb.sheet_by_index(0)
year=[]
statename=[]
season=[]
crop=[]
area=[]
production=[]
for i in range(2,sheet.nrows):
    statename.append(sheet.cell_value(i,0))
    year.append(int(sheet.cell_value(i,2)))
    season.append(sheet.cell_value(i,3))
    crop.append(sheet.cell_value(i,4).lower())
    area.append(sheet.cell_value(i,5))
    production.append(sheet.cell_value(i,6))
n=int(input('IF YOU ARE AN ADMIN ENTER 1 USER ENTER 2 QUIT ENTER 3\n'))
if n==1:
    f=admin()
    en=int(input('if you want to see data as user enter 1 else enter 2\n'))
    if en==1:
        user(f)
elif n==2:
    user(f)
    en1=int(input('if you want to see data as admin enter 1 else enter 2\n'))
    if en1==1:
        admin()
#NEED TO EDIT THE ENTRY DETAILS

                
                
            