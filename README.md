# school_data_acquisition

export.py

Description
This script extracts the information of n number of input files of the type:

Schüler- und Klassenstatistik für allgemeinbildende Schulen und Schulen für Erwachsene (??-??).xlsx 

, with “??-??” the period of time. It produces the following output files:
1. extract_output_region.xlsx
2. extract_output_region.log
3. change_name_school.dat

Output files from point 1 correspond to the recovered and re-arranged information from original files. There are as many files as regions available, with the regions contained in input file from point 1 (see below). Each file contains a list of schools (located in the region) in the first sheet and the information by school (for all available years) given in different sheets named with school ID. Only the following school forms are included (excluded “Foerderschule”) :

Intensivklassen, Vorbereitungsklassen, Eingangsstufe, Flexiber Schulanfang (Jg. 1 und 2.), Grundschule, Förderstufe, Hauptschule, Realschule, IGS, Praxis und Schule, Gymnasium (G8), Gymnasium (G9)

* if new input files contain additional school forms, the script must be modified.

Output file of point number 2 contains only technical details about how the program runs. This can be useful to detect future problems (changes in input files format might cause the malfunctioning of the program.)
Output file from point 3 gives a list of schools that have changed name. The last name is kept in the output excel files from point 1.

How it works
Before executing the program, the script (export.py) and the following input files must be in the same folder:
1. Excel file with the information about the distribution of the schools by region
2. A sub-folder containing input files with the information

File of point number 1 must contain its information in 2 columns (it’s not important the order, and can be that there are more columns with other information included in the file). The 2 columns correspond to ID of the school and its corresponding region. All this information must be in the first sheet of the  excel file. 
Name of files of point 2 are not relevant, although their format must be the same of the one used during the period 2009-2019.


Once the script is executed using the command:

python  export.py

, the user must answer the following questions:

1. Enter the name of the folder where the excel files are:
2. Enter the type of input files [ex.: xls or xlsx] in folder <here name given by user>:
3. Enter the name of the worksheet where the relevant data is [same for all Excel files ]:
4. Enter name of Excel file with school distribution by region:
5. Enter name of 2 columns with "school_ID" and corresponding "region" separated by a space (ex. Schulnr. BR):


