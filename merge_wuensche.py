#! /usr/bin/env python3
import re,os,subprocess,glob
import pandas as pd
import numpy as np
import xlsxwriter
from pandas import ExcelWriter
from pandas import ExcelFile
from tempfile import TemporaryFile


# files must have only one sheet
# max. rows excel = 1,048,576
# from 2009 to 2018 (9 years) we get 21729 rows (~1/48 from max) --> 432 (9*48) years more aprox we get the limit

xfolder = input("Enter the name of the folder where the excel files are:\n") or "123wunsche"

xtype = input("Enter the type of files [ex.: xls or xlsx] in folder "+xfolder+":\n") or "xlsx"

regfile = input("Enter name of Excel file with school distribution by region. It must contain the columns \"schulnr\" and \"bildungsregion_standort\" in first sheet:\n") or "SCHL_Schulen.xlsx"
cols_region = input("Enter name of 2 columns with \"school_ID\" and corresponding \"region\" separated by a space (ex. Schulnr. bildungsregion_standort):\n") or "Schulnr. BR"
column_id = cols_region.split(" ")[0]
column_reg = cols_region.split(" ")[1]


cwd = os.getcwd()
xfiles = sorted(glob.glob(cwd+"/"+xfolder+"/*."+xtype))

# getting : year, index end header and header
years = []
indices = []
headers = []
#totrows = 0
for xf in xfiles:
    indf = pd.read_excel(xf)
    name = xf.split("/")[-1].split(".")[0]+".temp"
    indf.to_csv(name,sep=';')
    m = 0
    
    with open(name,'r') as csv:
        #totrows += len(csv.readlines())
        #csv.seek(0)
        for l in csv:
            line = l.rstrip().split(";")
            if "Name" in line and "Strasse" in line:
                headers.append(line[1:])
                indices.append(int(line[0])+2)
                m = 1
            if m == 0:
                if "Parameterliste" in l:
                    for i,s in enumerate(line):
                        if "Schuljahr" in s:
                            # it must exist the line with the string "Parameterliste" containing the year, the format (example):
                            # Parameterliste: Schuljahr: 2018/2019; Schulaufsicht: F; Stufe: 5, 7, 11
                            year = line[i].split(":")[-1].split("/")[0]
                            years.append(year)

# working with data frames
writer = pd.ExcelWriter("alles_wuensche_raw.xlsx",engine='xlsxwriter') 
workbook=writer.book
# change here name worksheet
sheetname = "einzelwuensche_raw"
worksheet=workbook.add_worksheet(sheetname)
writer.sheets[sheetname] = worksheet
lastheaderpre = headers[-1]
lastheader = [x for x in lastheaderpre if x]
lastheader.insert(0,"Jahr")
lastheader.insert(len(lastheader),"Wunsche Summe")
row = 1
worksheet.set_column('A:Z', 12)
header_format = workbook.add_format({'bold': 1,'border': 1,'align': 'center','valign': 'vcenter','fg_color': '#FFFFFF'})


idxnam1 = [i for i, n in enumerate(lastheader) if n == "Name"][0]
lastheader.insert(idxnam1+1,"BR")
idxnam2 = [i for i, n in enumerate(lastheader) if n == "Name"][1]
lastheader.insert(idxnam2+1,"BR")

# writing header just at the beginning
for i,h in enumerate(lastheader):
    worksheet.write(0,i,h,header_format)

for j,xf in enumerate(xfiles):
    print("Working in file "+xf+" ...")
    h = headers[j]
    # adding region columns
    #searching indices
    idxnam1 = [i for i, n in enumerate(h) if n == "Name"][0]
    idxnam2 = [i for i, n in enumerate(h) if n == "Name"][1]
    idxid1 = [i for i, n in enumerate(h) if n == "Dst.Nr"][0]
    idxid2 = [i for i, n in enumerate(h) if n == "Dst.Nr"][1]
    # changing duplicated NAME in header
    h[idxnam1] = "Name1"
    h[idxnam2] = "Name2"
    h[idxid1] = "Dst.Nr_1"
    h[idxid2] = "Dst.Nr_2"
    
    df = pd.read_excel(xf,header=None)
    rowsdf = df.shape[0]
    listskip = [x for x in range(0,indices[j])]
    df = pd.read_excel(xf,names=h,header=None,index_col=None,skiprows=listskip)
    df = df.dropna(axis=1,how='all')
    rowsdf = df.shape[0]
    firstcol = [int(years[j]) for x in range(0,rowsdf)]
    df.insert(0,"Jahr",firstcol,True)
    #print(df.shape[0])
    #dfnew = df[indices[j]:]
    
    #regions file
    reg = pd.read_excel(regfile)
    reg = reg.dropna(axis=1,how='all')
    headerstr = list(reg.columns.values)
    idxschnr = headerstr.index(column_id)
    idxreg = headerstr.index(column_reg)

    idreg_pre = list(reg[headerstr[idxschnr]])
    regions_pre = list(reg[headerstr[idxreg]])
    
    idreg = []
    regions = []
    for i,re in enumerate(regions_pre):
        if pd.isna(re) != True:
            regions.append(re)
            idreg.append(idreg_pre[i])
    
    #extracting 2 "Name" columns
    colid1 = list(df["Dst.Nr_1"])
    colid2 = list(df["Dst.Nr_2"])
    
    newreg1 = []
    for i,id1 in enumerate(colid1):
        catch1 = 0
        for j,ir in enumerate(idreg):
            if id1 == ir:
                newreg1.append(regions[j])
                catch1 = 1
        if catch1 == 0:
            newreg1.append("")
    newreg2 = []
    for i,id2 in enumerate(colid2):
        catch2 = 0
        for j,ir in enumerate(idreg):
            if id2 == ir:
                newreg2.append(regions[j])
                catch2 = 1
        if catch2 == 0:
            newreg2.append("")
    
    newheader = list(df.columns.values)
    idxnam1 = [i for i, n in enumerate(newheader) if n == "Name1"][0]
    df.insert(idxnam1+1,"BR",newreg1,True)
    newheader = list(df.columns.values)
    idxnam2 = [i for i, n in enumerate(newheader) if n == "Name2"][0]
    df.insert(idxnam2+1,"BR",newreg2,True)
    newheader = list(df.columns.values)
    # add wünsches
    # remember that now lastheader has an extra column at the end ("Wünsche Summe")
    std_lenght = len(lastheader)-1
    if len(newheader) < std_lenght:
        rest = std_lenght - len(newheader)
        if rest == 2:
            rowsdf = df.shape[0]
            nullist = ["" for x in range(0,rowsdf)]
            df.insert(std_lenght-2,"Zweitwunsch",nullist,True)
            df.insert(std_lenght-1,"Drittwunsch",nullist,True)
        if rest == 1:
            rowsdf = df.shape[0]
            nullist = ["" for x in range(0,rowsdf)]
            df.insert(std_lenght-1,"Drittwunsch",nullist,True)
    # lastheader[-1] is "Wünsche Summe" column
    wsum = df[[lastheader[-4],lastheader[-3],lastheader[-2]]].sum(axis=1,skipna=True)
    df.insert(std_lenght,"Wunsche Summe",wsum,True)
    # where start counting rows. In this case start from 1 and not 0
    df.index = df.index + 1
    # print to excel
    df_obj = df.select_dtypes(['object'])
    df[df_obj.columns] = df_obj.apply(lambda x: x.str.strip())
    df.to_excel(writer,sheet_name=sheetname,startrow=row,startcol=0,index=False,header=False) 
    row = row + rowsdf 
writer.save()

pretempfiles = sorted(glob.glob("*.temp"))
for p in pretempfiles:
    os.remove(p)



