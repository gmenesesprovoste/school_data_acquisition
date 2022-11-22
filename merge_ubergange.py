#! /usr/bin/env python3
import re,os,subprocess,glob
import pandas as pd
import numpy as np
import xlsxwriter
from pandas import ExcelWriter
from pandas import ExcelFile
from tempfile import TemporaryFile


xfolder = input("Enter the name of the folder where the excel files are:\n") or "ubergange"

xtype = input("Enter the type of files [ex.: xls or xlsx] in folder "+xfolder+":\n") or "xlsx"

regfile = input("Enter name of Excel file with school distribution by region. It must contain the columns \"schulnr\" and \"bildungsregion_standort\" in first sheet:\n") or "SCHL_Schulen.xlsx"
cols_region = input("Enter name of 2 columns with \"school_ID\" and corresponding \"region\" separated by a space (ex. Schulnr. bildungsregion_standort):\n") or "Schulnr. BR"
column_id = cols_region.split(" ")[0]
column_reg = cols_region.split(" ")[1]


cwd = os.getcwd()
xfiles = sorted(glob.glob(cwd+"/"+xfolder+"/*."+xtype))

# getting : year, index end header and header
years = {} 
indices = {}
headers = {}
secondheaders = {}
sheets = {}
# very important that relevant sheets include or start with "5-F" and same number of header rows for a specific file (all sheets)
for xf in xfiles:
    nametemp = xf.split("/")[-1].split(".")[0]+".temp"
    namefile = xf.split("/")[-1]
    # this method is better for multiple sheet excel files
    indf = pd.ExcelFile(xf)
    # creating file-sheets dictionary
    all_sheets_pre = indf.sheet_names
    all_sheets = []
    for s in all_sheets_pre:
        if "5-" in s:
            all_sheets.append(s)
    sheets[namefile] = all_sheets
    
    headsheets = []
    for s in all_sheets:
        df = pd.read_excel(xf,sheet_name=s)
        hdf_find = df[df[list(df)[0]].str.contains("Jg.",na=False)]
        hdf_pre = hdf_find[list(hdf_find)[0:]]
        hdf = hdf_pre.values.tolist()[0]
        headsheets.append(hdf)
    headers[namefile] = headsheets
    
    df = pd.read_excel(xf,sheet_name=all_sheets[0])
    df.to_csv(nametemp,sep=';')
    
    m = 0
    with open(nametemp,'r') as csv:
        for l in csv:
            line = l.rstrip().split(";")
            cases = ["Name" in line,
                     "PLZ" in line]
            if all(cases):
                #headers[namefile] = line[1:]
                indices[namefile] = int(line[0])+4
                m = 1
            if m == 1 and not any(cases) :
                #secondheaders[namefile] = line[1:]
                m = 0
            if m == 0:
                if "Parameterliste" in l:
                    for i,s in enumerate(line):
                        if "Schuljahr" in s:
                            # it must exist the line with the string "Parameterliste" containing the year, the format (example):
                            # Parameterliste: Schuljahr: 2018/2019; Schulaufsicht: F; Stufe: 5, 7, 11
                            year = line[i].split(":")[-1].split("/")[0]
                            years[namefile] = year
        
    
# working with data frames
writer = pd.ExcelWriter("alles_ubergange_raw.xlsx",engine='xlsxwriter') 
workbook=writer.book
# change here name worksheet
sheetname = "ubergange_raw"
worksheet=workbook.add_worksheet(sheetname)
writer.sheets[sheetname] = worksheet


# finding definitive header
newheaders = []
for xf in xfiles:
    namefile = xf.split("/")[-1]
    for h in headers[namefile]:
        newh = [x for x in h if str(x) != 'nan' and str(x) != 'None']
        newheaders.append(newh)

lheader = newheaders[0]
inds = 0
for h in newheaders[1:]:
    for i,ele in enumerate(h):
        ele = str(ele).rstrip()
        if ele == "S":
            inds = i

        if ele not in lheader:
            lheader.append(ele)
for j in range(0,len(lheader)-inds):
    real = str(lheader[inds+(j*2)])
    lheader[inds+(j*2)] = str(lheader[inds+(j*2)])+"_abs"
    lheader.insert(inds+(j*2)+1,real+"_pro")


idxnam1 = [i for i, n in enumerate(lheader) if n == "Name"][0]
lheader.insert(idxnam1+1,"BR")
idxnam2 = [i for i, n in enumerate(lheader) if n == "Name"][1]
lheader.insert(idxnam2+1,"BR")
lheader[idxnam1] = "Name1"
lheader[idxnam2] = "Name2"
idxid1 = [i for i, n in enumerate(lheader) if n == "Nr."][0]
idxid2 = [i for i, n in enumerate(lheader) if n == "Nr."][1]
lheader[idxid1] = "Nr_1"
lheader[idxid2] = "Nr_2"

idxop1 = [i for i, n in enumerate(lheader) if n == "Ö/P"][0]
idxop2 = [i for i, n in enumerate(lheader) if n == "Ö/P"][1]
idxtyp1 = [i for i, n in enumerate(lheader) if n == "Typ"][0]
idxtyp2 = [i for i, n in enumerate(lheader) if n == "Typ"][1]
idxplz1 = [i for i, n in enumerate(lheader) if n == "PLZ"][0]
idxplz2 = [i for i, n in enumerate(lheader) if n == "PLZ"][1]
idxort1 = [i for i, n in enumerate(lheader) if n == "Ort"][0]
idxort2 = [i for i, n in enumerate(lheader) if n == "Ort"][1]
lheader[idxop1] = "Ö/P_1"
lheader[idxop2] = "Ö/P_2"
lheader[idxtyp1] = "Typ_1"
lheader[idxtyp2] = "Typ_2"
lheader[idxplz1] = "PLZ_1"
lheader[idxplz2] = "PLZ_2"
lheader[idxort1] = "Ort_1"
lheader[idxort2] = "Ort_2"

lheader_orig = []
for ele in lheader:
    if ele != "BR":
        lheader_orig.append(ele)

#print(lheader_orig)

lheader.insert(0,"Jahr")

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

worksheet.set_column('A:AD', 12)
header_format = workbook.add_format({'bold': 1,'border': 1,'align': 'center','valign': 'vcenter','fg_color': '#FFFFFF'})

# writing header just at the beginning
for i,h in enumerate(lheader):
    worksheet.write(0,i,h,header_format)


row = 1

# entering in the big loop ##########################################################################################################
for j,xf in enumerate(xfiles):

    namefile = xf.split("/")[-1]
    print("Working in file "+namefile+" ...")
    
    # loop by worksheet ##########################################
    list_sheets = sheets[namefile]
    for k,ws in enumerate(list_sheets):            
        h = headers[namefile][k]
        #new header with good labels
        hup = []
        for i,el in enumerate(h):
            if str(el) == 'nan' or str(el) == 'None':
                    hup.append("")
            else:
                    hup.append(el)
        #index from empty elements
        idxes = [i for i, n in enumerate(hup) if n == ""]
        for i in idxes:
            if str(h[i-1]) != 'nan' and str(h[i-1]) != 'None':
                hup[i] = str(hup[i-1]).rstrip()+"_pro"
                hup[i-1] = str(hup[i-1]).rstrip()+"_abs"
         
        
        # adding region columns
        
        #searching indices
        idxnam1 = [i for i, n in enumerate(hup) if n == "Name"][0]
        idxnam2 = [i for i, n in enumerate(hup) if n == "Name"][1]
        idxid1 = [i for i, n in enumerate(hup) if n == "Nr."][0]
        idxid2 = [i for i, n in enumerate(hup) if n == "Nr."][1]
        
        
        idxop1 = [i for i, n in enumerate(hup) if n == "Ö/P"][0]
        idxop2 = [i for i, n in enumerate(hup) if n == "Ö/P"][1]
        idxtyp1 = [i for i, n in enumerate(hup) if n == "Typ"][0]
        idxtyp2 = [i for i, n in enumerate(hup) if n == "Typ"][1]
        idxplz1 = [i for i, n in enumerate(hup) if n == "PLZ"][0]
        idxplz2 = [i for i, n in enumerate(hup) if n == "PLZ"][1]
        idxort1 = [i for i, n in enumerate(hup) if n == "Ort"][0]
        idxort2 = [i for i, n in enumerate(hup) if n == "Ort"][1]

        # changing duplicated NAME in header
        hup[idxnam1] = "Name1"
        hup[idxnam2] = "Name2"
        hup[idxid1] = "Nr_1"
        hup[idxid2] = "Nr_2"
        
        hup[idxop1] = "Ö/P_1"
        hup[idxop2] = "Ö/P_2"
        hup[idxtyp1] = "Typ_1"
        hup[idxtyp2] = "Typ_2"
        hup[idxplz1] = "PLZ_1"
        hup[idxplz2] = "PLZ_2"
        hup[idxort1] = "Ort_1"
        hup[idxort2] = "Ort_2"
        #print(hup) 
        # jumping header
        nr_skip = int(indices[namefile])
        listskip = [x for x in range(0,nr_skip)]
        #        
        newhup = []
        nulindex = []
        for i,el in enumerate(lheader_orig):
            if el in hup:
                newhup.append(el)
            else:
                nulindex.append(i)

        df = pd.read_excel(xf,sheet_name=ws,names=hup,header=None,index_col=None,skiprows=listskip)
        df = df.dropna(axis=1,how='all')
        df = df[newhup]
        rowsdf = df.shape[0]
        nulcol = ["" for x in range(0,rowsdf)]
        for ind in nulindex:
            df.insert(ind,"nul",nulcol,True)
        firstcol = [int(years[namefile]) for x in range(0,rowsdf)]
        df.insert(0,"Jahr",firstcol,True)
        
        #extracting 2 "Name" columns
        colid1 = list(df["Nr_1"])
        colid2 = list(df["Nr_2"])
        
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
        
        df_obj = df.select_dtypes(['object'])
        df[df_obj.columns] = df_obj.apply(lambda x: x.str.strip())
        df.to_excel(writer,sheet_name=sheetname,startrow=row,startcol=0,index=False,header=False,float_format = "%0.1f") 
        row = row + rowsdf  
writer.save()
#
pretempfiles = sorted(glob.glob("*.temp"))
for p in pretempfiles:
    os.remove(p)




