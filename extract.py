#! /usr/bin/env python3
import re,os,subprocess,glob
import pandas as pd
import numpy as np
import xlsxwriter
from pandas import ExcelWriter
from pandas import ExcelFile
from tempfile import TemporaryFile


# This program will work if schulforms from the input excel are contained in the following list:
#['Vorklasse', 'Grundschule', 'Hauptschule', 'Realschule', 'Eingangsstufe', 'Förderstufe', 'achtj. Gymnasium', 'Integr. Jahrg.stufe', 'Gymnasium', 'flex. Schulanfang', 'Seiteneinsteiger (in Intensivklassen)', 'Praxis und Schule allgemeinbildend']
# if there is one or more different ones, the program must be modified.



############################# interaction with the user ######################################################################################

# user have to create a work folder in which there is another folder with the .xlsx files. Place this script in the work folder
print("First, the user have to create a work folder into which there is another folder containing the excel files.\nPlace this script in the work folder.\n")

# entering type of files to read and the folder where they are
xfolder = input("Enter the name of the folder where the excel files are:\n") or "WIB"
xtype = input("Enter the type of files [ex.: xls or xlsx] in folder "+xfolder+":\n") or "xlsx"
keysheet= input("Enter the name of the worksheet where the relevant data is [same for all Excel files ]:\n") or "Tabellenblatt2"

# enter file with schools distribution by region
regfile = input("Enter name of Excel file with school distribution by region. It must contain the columns \"schulnr\" and \"bildungsregion_standort\" in first sheet:\n") or "SCHL_Schulen.xlsx"
cols_region = input("Enter name of 2 columns with \"school_ID\" and corresponding \"region\" separated by a space (ex. Schulnr. bildungsregion_standort):\n") or "Schulnr. BR"
column_id = cols_region.split(" ")[0]
column_reg = cols_region.split(" ")[1]

############################# list of the files (only these files in the folder) ################################################################

# in which folder am I
cwd = os.getcwd()
xfiles = sorted(glob.glob(cwd+"/"+xfolder+"/*."+xtype))
# os.chdir()


################################ input file format, getting information ###################################################################################
# schools type SOFS, LER and BS are exluded
included_schools = [' GH', ' G', ' H', ' GHR', ' HR', ' IGS', ' GYM', ' KGS', ' GOS', ' R', ' GYMM']

schulform_pre = []
years = []

for xfile in xfiles:
    start = 0
    # name temp file
    fin_pre = xfile.split(" ")[-1]
    fin = fin_pre.split('.')[0][1:-1]


# loading spreadsheet (recovering excel info of one file)
    indf = pd.read_excel(xfile,keysheet)
    namefilepre = fin+"_pre.temp"
    namefile = fin+".temp"

# conversion to csv format  
    indf.to_csv(namefilepre,sep=';')
# searching the line to start => skipping big header in a new csv temp file (namefile).
# this file will contain forbidden schools, but schulform_pre will not include schulforms of these schools
    with open(namefile,'w') as temp:
        with open(namefilepre,'r') as incsv:
            start = 100
            i = 0
            for nline in incsv:
                line =  nline.rstrip().split(";")
                # getting the year from inside the file
                if line:
                    if "Parameterliste" in nline:
                        for j,s in enumerate(line):
                            if "Schuljahr" in s:
                                # it must exist the line with the string "Parameterliste" containing the year, the format (example):
                                # Parameterliste: Schuljahr: 2018/2019; Schulaufsicht: F; Stufe: 5, 7, 11
                                year = line[j].split(":")[-1].split("/")[0]
                                years.append(year)
                    if "Schulaufsicht" and "Primarbereich" in line:
                        start = i
                    if i > start:
                        if "davon IB" not in line:
                            joinline = ";".join(line[1:])
                            print (str(i)+";"+joinline, file=temp)
                            #input_array_pre.append(line)
                            # filtering by admitted schulforms
                            if line[1]:
                                if "ÖFF" in line[1]:
                                    elemname = line[1].replace("\"","")
                                    type_sch = elemname.split(",")[2]
                            # making list all type schule contained in all input files        
                            if line[3] and "Dienststelle" not in line[1]:
                                rules = ["Klassen"  in line[4],
                                         "Insgesamt" in line[3]]
                                if not any(rules):
                                    if type_sch in included_schools:
                                        schulform_pre.append(line[3])
                    i = i + 1


# making the complete list of schulform unique
schulform = []
for s in schulform_pre:
    if s:
        if "Schulform" not in s and "Insgesamt" not in s:
            if s not in schulform:
                schulform.append(s)
##################################### list all files in csv format ##########################################################################################

yearfileslist = sorted(glob.glob('??-??.temp'))


###################### finding all schools (with repetitions) #########################################################################################################

# gran loop in each yearfile
schulen = []
IDs = []
allheaders = []
# name of the file with the form 
# it is important that the name of the year file is same than example
i = 0
for yearfile in yearfileslist:
    with open(yearfile,"r") as yf:
        year = int(years[i])
        i += 1
        for l in yf:
            linearr = l.split(";")
            elem1 = linearr[1]

            if "Dienststelle" in l:
                liarr = l.split(";")
                allheaders.append(liarr)
            if elem1:
                if "ÖFF" in elem1:
                    elemname = elem1.replace("\"","")
                    type_sch = elemname.split(",")[2]
                    # making list all type schule contained in all input files        
                    if type_sch in included_schools:
                        try:
                            firstelem1 = elem1.split(",")[0]
                            
                            #eliminating possible quotes
                            firstelem1 = firstelem1.replace("\"","")
                            
                            ID = int(firstelem1.split(" ")[0])
                            schule = " ".join(firstelem1.split(" ")[1:])
                            schulen.append(schule)
                            IDs.append(ID)
                        except ValueError:
                            pass
schulen.reverse()
IDs.reverse()

############################# finding total and definitive list of schools for all the data available ##################################################################
# making the first sheet with lists of school. Detecting as well schools that changed their name
change = open("change_name_school.dat","w")
print("ID ; last name school ; old name school",file=change)
schulenuniq = []
IDsuniq = []
rep = []
for k in range (0,len(IDs)):
    ID2 = IDs[k]
    schule2 = schulen[k]
    if ID2 not in IDsuniq:
        IDsuniq.append(ID2)
        schulenuniq.append(schule2)
    else:
        idxrep = IDsuniq.index(ID2)
        if schule2 != schulenuniq[idxrep]:
            rep.append(str(ID2)+" ; "+schulenuniq[idxrep]+" ; "+schule2)
repuniq = []
for el in rep:
    if el not in repuniq:
        repuniq.append(el)
        print(el,file=change)



change.close()


schulenuniq.reverse()
IDsuniq.reverse()


######## first sheet in output file
def firstsheet_schulen(firstworksheet,IDschulenlist,schulenlist):
    for i,elem in enumerate(IDschulenlist):
        firstworksheet.write('A'+str(i+2),i+1)
        firstworksheet.write('B'+str(i+2),elem)
        firstworksheet.write_url('C'+str(i+2),"internal:"+str(elem)+"!A1")
        firstworksheet.write('C'+str(i+2),schulenlist[i])

shiftlist = [1,1,1,1,7,3,11,11,11,3,9,18]
def sheet_by_schule(sheet, outputheaderlist, shiftlist, idschule, nameschule,rowheader):
        sheet.write(0,1,nameschule)
        sheet.write(0,0,idschule)
        ini = 2
    # creando header escuela sheet
        for i,oh in enumerate(outputheaderlist):
            shift = shiftlist[i]
            sheet.merge_range(rowheader,ini,rowheader,ini+shift,oh)
            ini = ini + shift + 1 
        sheet.write(rowheader+1,1,"Schuljahr")
        # CORREGIR 90 (generico) 
        for j in range(0,90,2):
            sheet.write(rowheader+1,j+2,"SuS")
            sheet.write(rowheader+1,j+3,"K")
        for k,y in enumerate(years):
            sheet.write(k+rowheader+2,1,int(y))


######################################### function that creates excel output file for a list od IDs schools ################################################################
def schoolinfo_by_region(IDslist,schulenlist):
    
    ######### loop by school 
    exc_schulen = []
    #for sh in schulenuniq[5:7]:
    for d in range(0,len(IDslist)):
        sh = schulenlist[d]
        idu = IDslist[d]
        print("********************************************",file=log)
        print(idu,file=log)
        print(sh,file=log)
        
        encounter1 = 0        
        m = 0
        #opening all year files for schule sh
        for yearfile in yearfileslist:
            # parameter to ignore insgesamt lines
            s = 0
            #parameter to ignore "Förderschule"
            t = 0
            # parameter to ignore "Förderschule" in lines after the first
            q = 0
            #parameter to include "Flex.Schulanfang"
            p = 0
    
            fromhere = 0
            #year = "20"+yearfile.split("-")[0]
            year = years[m]
            #print("Searching for "+sh+" school during the year "+year)
            print("**********************"+year+"*********************************************",file=log)
            headerpre = allheaders[m]
            header = [maybe_float(v) for v in headerpre]
            header[0] = ""
            try:
                diens_idx = header.index("Dienststelle")
                sform_idx = header.index("Schulform")
                kateg_idx = header.index("Kategorie")
                if not header[kateg_idx + 1]:
                    idxsign = kateg_idx + 1
                    
                idx0 = header.index(0)
                idx1 = header[1:].index(1)+1
                idx2 = header.index(2)
                idx3 = header.index(3)
                idx4 = header.index(4)
                idx5 = header.index(5)
                idx6 = header.index(6)
                idx7 = header.index(7)
                idx8 = header.index(8)
                idx9 = header.index(9)
                idx10 = header.index(10)
                idx11 = header.index(11)
                idx12 = header.index(12)
                idx13 = header.index(13)
                idx14 = header.index(14)
            except ValueError:
                print ("Header does not contain the right parameters. Might be that the program does not work properly.\n")
            
            inheaderindex = [[idx0],[idx1,idx2,idx3,idx4],[idx5,idx6,idx7,idx8,idx9,idx10],[idx5,idx6,idx7,idx8,idx9,idx10],[idx0],[idx5,idx6],[idx5,idx6,idx7,idx8,idx9],[idx5,idx6,idx7,idx8,idx9,idx10],[idx5,idx6,idx7,idx8,idx9,idx10,idx11,idx12,idx13,idx14],[idxsign,idx0,idx1,idx2,idx3,idx4,idx5,idx6,idx7,idx8,idx9,idx10,idx11,idx12,idx13,idx14],[idx1,idx2],[idxsign],[idx8,idx9]]
            
            # position in outfile following inheaderindex above
            outheaderindex = [[4],[10,12,14,16],[22,24,26,28,30,32],[34,36,38,40,42,44],[6],[18,20],[62,64,66,68,70],[46,48,50,52,54,56],[72,74,76,78,80,82,84,86,88,90],[92,94,96,98,100,102,104,106,108,110,112,114,116,118,120,122],[8],[2],[58,60]]
            
            
            
            addtogrund = 0
            addtogrundk = 0
            with open(yearfile,"r") as yf:
                for lin in yf:
                    linpre = lin.rstrip()
                    listlin = linpre.split(";")
                    nl = int(listlin[0])
################ detecta donde termina informacion de una escuela determinada
                    if fromhere != 0:
                        if listlin[diens_idx]:
                            break
################ linea que contiene nombre escuela             
                    if listlin[diens_idx]:
                        if str(idu) in listlin[diens_idx]:
                            # not to forget the "," for more clauses
                            rules1 = ["Förderschule" in listlin]
                            if any(rules1):  
                                t = 1
                                q = 1
                                fromhere = int(listlin[0])
                                exc_schulen.append(idu)
                            else:
                                if encounter1 == 0:
                                    worksheet = outfile.add_worksheet(str(idu))
                                    writer.sheets[str(idu)] = worksheet
                                    rowheader = 2
                                    # preparando sheet sh
                                    sheet_by_schule(worksheet, outheaderboth, shiftlist, idu, sh,rowheader)
                                    encounter1 += 1

                                t = 0
                                fromhere = int(listlin[0])
                                schulform = listlin[sform_idx]
                                print(schulform,file=log)
                                inde = inheaderboth.index(schulform)
                                inidx = inheaderindex[inde]
                                outidx = outheaderindex[inde]
                                print(listlin,file=log)
                                rules2 = ["flex. Schulanfang" in listlin]
                                          #"Eingangsstufe" in listlin]
                                if not any(rules2):
                                    p = 0
                                    for i in range(0,len(inidx)):
                                       if float(listlin[inidx[i]]) > 0:
                                           worksheet.write(m+rowheader+2,outidx[i],int(float(listlin[inidx[i]])))
                                       else:
                                           worksheet.write(m+rowheader+2,outidx[i],"")
                                if "flex. Schulanfang" in listlin:
                                    p = 1
                                    totflex = 0
                                    for i in range(0,len(inidx)):
                                        if float(listlin[inidx[i]]) > 0:
                                            totflex = totflex + int(float(listlin[inidx[i]]))
                                        else:
                                            totflex = totflex + 0
                                    if totflex == 0:
                                        worksheet.write(m+rowheader+2,outidx[0],"")
                                    else:    
                                        worksheet.write(m+rowheader+2,outidx[0],totflex)    
                                if "Eingangsstufe" in listlin:
                                    addtogrund = int(float(listlin[idx1]))
    
    
    
##################### lineas que vienen despues de la que contiene nombre escuela    
                    if fromhere != 0:
                        if not listlin[diens_idx]:
                            rules2 = ["Klassen" not in listlin,
                                      "Insgesamt" not in listlin]
                            #entro a linea despues de primera, sin klassen o insgesamt, with another schulform
                            if all(rules2):  
                                rules1 = ["Förderschule" in listlin]
                                # here forderschule as schulform, jump it          
                                if any(rules1):
                                    q = 1
                                    t = 1
                                    exc_schulen.append(idu)
                                # all the other options    
                                else:
                                    t = 0
                                    schulform = listlin[sform_idx]
                                    print(schulform,file=log)
                                    print(listlin,file=log)
                                    inde = inheaderboth.index(schulform)
                                    inidx = inheaderindex[inde]
                                    outidx = outheaderindex[inde]
    
                                    rules3 = ["flex. Schulanfang" in listlin,
                                              q == 1]
                                    # all the other options except the ones of rules3
                                    if not any(rules3):
                                        p = 0
                                        for i in range(0,len(inidx)):
                                           if float(listlin[inidx[i]]) > 0:
                                               if schulform == "Grundschule" and i == 0:
                                                   worksheet.write(m+rowheader+2,outidx[i],int(float(listlin[inidx[i]]))+addtogrund)
                                               else:
                                                   worksheet.write(m+rowheader+2,outidx[i],int(float(listlin[inidx[i]])))
                                           else:
                                               if schulform == "Grundschule" and i == 0 and addtogrund > 0:
                                                   worksheet.write(m+rowheader+2,outidx[i],addtogrund)
                                               else:    
                                                   worksheet.write(m+rowheader+2,outidx[i],"")
                                    # option flex. schulanfang (add 2 values)               
                                    if "flex. Schulanfang" in listlin and q == 0:
                                        p = 1
                                        totflex = 0
                                        for i in range(0,len(inidx)):
                                            if float(listlin[inidx[i]]) > 0:
                                                totflex = totflex + int(float(listlin[inidx[i]]))
                                            else:
                                                totflex = totflex + 0
                                        if totflex == 0:
                                            worksheet.write(m+rowheader+2,outidx[0],"")
                                        else:    
                                            worksheet.write(m+rowheader+2,outidx[0],totflex)    
                                        
                                        
                            # marking line of Insgesamt to ignore   
                            if "Insgesamt" in listlin:
                                s = 1
                            rulesk = ["Klassen" in listlin,
                                      s == 0,
                                      t == 0,
                                      q == 0]
                            # all the lines only with "klassen" (s to ignore "Insgesamt" line (s = 1), and t to ignore "forderschule" (t = 1))
                            if all(rulesk):    
                                print(listlin,file=log)
                                rules3 = [schulform == "flex. Schulanfang"]
                                          #schulform == "Eingangsstufe"]
                                if not any(rules3):
                                    for i in range(0,len(inidx)):
                                       if float(listlin[inidx[i]]) > 0:
                                           if schulform == "Grundschule" and i == 0:
                                               worksheet.write(m+rowheader+2,outidx[i]+1,int(float(listlin[inidx[i]]))+addtogrundk)
                                           else:
                                               worksheet.write(m+rowheader+2,outidx[i]+1,int(float(listlin[inidx[i]])))
                                       else:
                                           if schulform == "Grundschule" and i == 0 and addtogrund > 0:
                                               worksheet.write(m+rowheader+2,outidx[i]+1,addtogrundk)
                                           else:    
                                               worksheet.write(m+rowheader+2,outidx[i]+1,"")
                                if schulform == "Eingangsstufe":
                                    addtogrundk = int(float(listlin[idx1]))
                                if p > 0:
                                    totflexk = 0
                                    for i in range(0,len(inidx)):
                                       if float(listlin[inidx[i]]) > 0:
                                           totflexk = totflexk + int(float(listlin[inidx[i]]))
                                       else:
                                           totflexk = totflexk + 0
                                    if totflexk == 0:
                                        worksheet.write(m+rowheader+2,outidx[0]+1,"")
                                    else:    
                                        worksheet.write(m+rowheader+2,outidx[0]+1,totflexk)    
                                         
            m = m + 1        
    
    
    exc_sch_uniq = []
    for s in exc_schulen:
        if s not in exc_sch_uniq:
            exc_sch_uniq.append(s)
    
    
    #remove schools with "Förderschule" from the list of schools
    final_id = []
    final_sch = []
    for i,d in enumerate(idlist_region):
        if d not in exc_sch_uniq:
            final_id.append(d)
            final_sch.append(schulelist_region[i])
    
    print("Excluded schools with Förderschule:")
    print(exc_sch_uniq)

    firstsheet_schulen(firstworksheet,final_id,final_sch)


################################################## end function ######################################################################################



############################### input and output headers  ############################################################################################

inheaderboth = ['Vorklasse', 'Grundschule', 'Hauptschule', 'Realschule', 'Eingangsstufe', 'Förderstufe', 'achtj. Gymnasium', 'Integr. Jahrg.stufe', 'Gymnasium', 'Förderschule', 'flex. Schulanfang', 'Seiteneinsteiger (in Intensivklassen)', 'Praxis und Schule allgemeinbildend']
outheaderboth = ['Intensivklassen','Vorbereitungsklassen','Eingangsstufe','Flexiber Schulanfang (Jg. 1 und 2.)','Grundschule','Förderstufe','Hauptschule','Realschule','IGS','Praxis und Schule','Gymnasium (G8)','Gymnasium (G9)']



############################# regions file ############################################################################################################


reg = pd.read_excel(regfile)
reg = reg.dropna(axis=1,how='all')
headerstr = list(reg.columns.values)
idxschnr = headerstr.index(column_id)
idxreg = headerstr.index(column_reg)

idnr_pre = list(reg[headerstr[idxschnr]])
regions_pre = list(reg[headerstr[idxreg]])

#eliminating possible nan from both lists and creating a list of all regions available
idnr = []
regions = []
regionsuniq = []

for i,re in enumerate(regions_pre):
    if pd.isna(re) != True:
        regions.append(re)
        idnr.append(idnr_pre[i])
        if re not in regionsuniq:
            regionsuniq.append(re)



# function to get numeric values in the correct type
def maybe_float(s):
    try:
        return float(s)
    except (ValueError,TypeError):
        return s


#### making the loop by region, using the function schoolinfo_by_region() ##########################################################################

for region in regionsuniq:
    print("\n")
    print("Working on schools from region "+region)
    log = open("extract_output_"+region+".log","w")
    # creating output file 
    writer = pd.ExcelWriter('extract_output_'+region+"."+xtype,engine ='xlsxwriter')
    outfile=writer.book
    firstworksheet = outfile.add_worksheet("Schulen")
    writer.sheets["Schulen"] = firstworksheet
    secondworksheet = outfile.add_worksheet("1")
    writer.sheets["1"] = secondworksheet
    idlist_region = []
    schulelist_region = []
    for iduni in IDsuniq:
        for idreg in idnr:
            if int(iduni) == int(idreg):
                idx = int(idnr.index(idreg))
                idxsch = int(IDsuniq.index(iduni))
                re = regions[idx]
                if region == re:
                    schulelist_region.append(schulenuniq[idxsch])
                    idlist_region.append(iduni)
    schoolinfo_by_region(idlist_region,schulelist_region)                
   
    lastworksheet = outfile.add_worksheet("99999")
    writer.sheets["99999"] = lastworksheet
    
    outfile.close()
    log.close()

writer.save

# all schools #################################################

os.chdir(cwd)
pretempfiles = sorted(glob.glob("*.temp"))
for p in pretempfiles:
    os.remove(p)



















