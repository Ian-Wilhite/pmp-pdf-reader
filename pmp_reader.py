"""
Order of the Arrow Performance Measurement Program PDF reader 1.1.4
Created by Ian Wilhite
1/22/2024

LIMITATIONS / USAGE:
this only views the ".pdf" files also located in the same folder as this script, and will create the 
csv in the same folder. This script is designed on the "PMP/JTE Detailed Worksheet" lodgemaster 
generated file (circa 2024). It also tends to begin to grab incorrect data on reports pre-2019.

INSTRUCTIONS:
download all reports and name files with the first word of the file name being the name of the lodge (case sensitive).
place all files into a single folder with the python script
run the script and find the generated csv file
paste the csv file into the excel sheet and verify data (the G3 and E3 are functionally identical) 

BACKGROUND:
as of writing this, I am the G3 PMP coordinator, and i got really tired of copying all the data 
from the lodgemaster-generated pdfs into an excel sheet that I could actually compare the different
lodges. I joked it might be faster for me to write a program to do it, so alas, bon apetit. If you 
are in a similar position in your lodge, section, or nationally, remember that the data only goes so 
far, and the real value in these positions is infomred decision making and goal setting. 

contact:
if you have any questions about PMP, this file, or anything else, shoot an email to iwilhite@netopalis.org
"""

import os 
#import PyPDF2 
from pandas import DataFrame
from tabula.io import read_pdf
#import tabulate


class LodgeData:
    
    years = []
    election = []
    induction = []
    activation = []
    m_retention = [] 
    m_growth = []
    e_participation = []
    bro_rate = []
    dol_mem = [] 
    ser_mem = []
    training = []
    y_ordeal = []
    y_bro = []
    y_vigil = []
    ya_ordeal = []
    ya_bro = []
    ya_vigil = []
    a_ordeal = []
    a_bro = []
    a_vigil = []
    y_all = []
    ya_all = []
    a_all = []
    ordeal_all = []
    bro_all = []
    vigil_all = []
    lodge_all = []
    
    def __init__(self) -> None:
        pass
    
    def __str__(self) -> str:
        to_str = [''.join(yr) for yr in self.years][0]
        return to_str
    
    def __get__(self, instance, owner):
        return list[str] #trust
    
    def __set__(self, instance, value): #I only ever want these to be list[str]
        self.years = list(self.years)
        self.election = list(self.election)
        self.induction = list(self.induction)
        self.activation = list(self.activation)
        self.m_retention = list(self.m_retention)
        self.m_growth = list(self.m_growth)
        self.e_participation = list(self.e_participation)
        self.bro_rate = list(self.bro_rate)
        self.dol_mem = list(self.dol_mem)
        self.ser_mem = list(self.ser_mem)
        self.training = list(self.training)
        self.y_ordeal = list(self.y_ordeal)
        self.y_bro = list(self.y_bro)
        self.y_vigil = list(self.y_vigil)
        self.ya_ordeal = list(self.ya_ordeal)
        self.ya_bro = list(self.ya_bro)
        self.ya_vigil = list(self.ya_vigil)
        self.a_ordeal = list(self.a_ordeal)
        self.a_bro = list(self.a_bro)
        self.a_vigil = list(self.a_vigil)
        self.y_all = list(self.y_all)
        self.ya_all = list(self.ya_all)
        self.a_all = list(self.a_all)
        self.ordeal_all = list(self.ordeal_all)
        self.bro_all = list(self.bro_all)
        self.vigil_all = list(self.vigil_all)
        self.lodge_all = list(self.lodge_all)
        
    
    def write_info(self, f_sum, folder):
        #convert to str & remove commas (numbers >1000 are formatted w commas)
        self.years = list(map(lambda x: str(x).replace(",",""), self.years))
        self.election = list(map(lambda x: str(x).replace(",",""), self.election))
        self.induction = list(map(lambda x: str(x).replace(",",""), self.induction))
        self.activation = list(map(lambda x: str(x).replace(",",""), self.activation))
        self.m_retention = list(map(lambda x: str(x).replace(",",""), self.m_retention))
        self.m_growth = list(map(lambda x: str(x).replace(",",""), self.m_growth))
        self.e_participation = list(map(lambda x: str(x).replace(",",""), self.e_participation))
        self.bro_rate = list(map(lambda x: str(x).replace(",",""), self.bro_rate))
        self.dol_mem = list(map(lambda x: str(x).replace(",",""), self.dol_mem))
        self.ser_mem = list(map(lambda x: str(x).replace(",",""), self.ser_mem))
        self.training = list(map(lambda x: str(x).replace(",",""), self.training))
        self.y_ordeal = list(map(lambda x: str(x).replace(",",""), self.y_ordeal))
        self.y_bro = list(map(lambda x: str(x).replace(",",""), self.y_bro))
        self.y_vigil = list(map(lambda x: str(x).replace(",",""), self.y_vigil))
        self.ya_ordeal = list(map(lambda x: str(x).replace(",",""), self.ya_ordeal))
        self.ya_bro = list(map(lambda x: str(x).replace(",",""), self.ya_bro))
        self.ya_vigil = list(map(lambda x: str(x).replace(",",""), self.ya_vigil))
        self.a_ordeal = list(map(lambda x: str(x).replace(",",""), self.a_ordeal))
        self.a_bro = list(map(lambda x: str(x).replace(",",""), self.a_bro))
        self.a_vigil = list(map(lambda x: str(x).replace(",",""), self.a_vigil))
        self.y_all = list(map(lambda x: str(x).replace(",",""), self.y_all))
        self.ya_all = list(map(lambda x: str(x).replace(",",""), self.ya_all))
        self.a_all = list(map(lambda x: str(x).replace(",",""), self.a_all))
        self.ordeal_all = list(map(lambda x: str(x).replace(",",""), self.ordeal_all))
        self.bro_all = list(map(lambda x: str(x).replace(",",""), self.bro_all))
        self.vigil_all = list(map(lambda x: str(x).replace(",",""), self.vigil_all))
        self.lodge_all = list(map(lambda x: str(x).replace(",",""), self.lodge_all))
        
        #write to file
        
        os.chdir(starting_folder)#to write to the file in pmp reader
        
        f_sum.write('years, ')# bc Im writing to a csv, a comma represents the next cell, and a \n represents the next row
        f_sum.write(', '.join(self.years) + "\n")
        f_sum.write('election rate, ')
        f_sum.write(', '.join(self.election) + "\n")
        f_sum.write('induction rate, ')
        f_sum.write(', '.join(self.induction) + "\n")
        f_sum.write('activation rate, ')
        f_sum.write(', '.join(self.activation) + "\n")
        f_sum.write('membership retention, ')
        f_sum.write(', '.join(self.m_retention) + "\n")
        f_sum.write('membership growth, ')
        f_sum.write(', '.join(self.m_growth) + "\n")
        f_sum.write('event participation, ')
        f_sum.write(', '.join(self.e_participation) + "\n")
        f_sum.write('brotherhood conversion, ')
        f_sum.write(', '.join(self.bro_rate) + "\n")
        f_sum.write('donations per memeber, ')
        f_sum.write(', '.join(self.dol_mem) + "\n")
        f_sum.write('servicer per member, ')
        f_sum.write(', '.join(self.ser_mem) + "\n")
        f_sum.write('training rate, ')
        f_sum.write(', '.join(self.training) + "\n")
        f_sum.write('youth ordeal members, ')
        f_sum.write(', '.join(self.y_ordeal) + "\n")
        f_sum.write('youth brotherhood members, ')
        f_sum.write(', '.join(self.y_bro) + "\n")
        f_sum.write('youth vigil members, ')
        f_sum.write(', '.join(self.y_vigil) + "\n")
        f_sum.write('young-adult ordeal members, ')
        f_sum.write(', '.join(self.ya_ordeal) + "\n")
        f_sum.write('young adult brotherhood members, ')
        f_sum.write(', '.join(self.ya_bro) + "\n")
        f_sum.write('young adult vigil members, ')
        f_sum.write(', '.join(self.ya_vigil) + "\n")
        f_sum.write('adult ordeal members, ')
        f_sum.write(', '.join(self.a_ordeal) + "\n")
        f_sum.write('adult brotherhood members, ')
        f_sum.write(', '.join(self.a_bro) + "\n")
        f_sum.write('adult vigil members, ')
        f_sum.write(', '.join(self.a_vigil) + "\n")
        f_sum.write('all youth members, ')
        f_sum.write(', '.join(self.y_all) + "\n")
        f_sum.write('all ya members, ')
        f_sum.write(', '.join(self.ya_all) + "\n")
        f_sum.write('all adult members, ')
        f_sum.write(', '.join(self.a_all) + "\n")
        f_sum.write('all ordeal members, ')
        f_sum.write(', '.join(self.ordeal_all) + "\n")
        f_sum.write('all brotherhood members, ')
        f_sum.write(', '.join(self.bro_all) + "\n")
        f_sum.write('all vigil members, ')
        f_sum.write(', '.join(self.vigil_all) + "\n")
        f_sum.write('all lodge members, ')
        f_sum.write(', '.join(self.lodge_all) + "\n")
        
        os.chdir(starting_folder+"/Reports")#to continue to read the files within Reports
        
        
    def wipe(self):
        self.years = []
        self.election = []
        self.induction = []
        self.activation = []
        self.m_retention = [] 
        self.m_growth = []
        self.e_participation = []
        self.bro_rate = []
        self.dol_mem = [] 
        self.ser_mem = []
        self.training = []
        self.y_ordeal = []
        self.y_bro = []
        self.y_vigil = []
        self.ya_ordeal = []
        self.ya_bro = []
        self.ya_vigil = []
        self.a_ordeal = []
        self.a_bro = []
        self.a_vigil = []
        self.y_all = []
        self.ya_all = []
        self.a_all = []
        self.ordeal_all = []
        self.bro_all = []
        self.vigil_all = []
        self.lodge_all = []
    
def str_search(regex: str, df, case=False):
    """Search all the text columns of `df`, return rows with any matches."""
    textlikes = df.select_dtypes(include=[object, "string"])
    return df[
        textlikes.apply(
            lambda column: column.str.contains(regex, regex=True, case=case, na=False)
        ).any(axis=1)
    ]
    

# PART 1 - - - - - file reading / data processing- - - - - 

#navigate into the reports folder
starting_folder = os.getcwd()


os.chdir(starting_folder+"/Reports")#file list needs to come from the reports folder
files = [f for f in os.listdir() if os.path.isfile(f)] #creates str array of all filenames
files.sort()#sorts alphabetically

pdfs = []
for i, f in enumerate(files): #new list pdfs of pdf filenames
    if f[-3:] == "pdf": #checks in case non-pdf files ended up in the Reports folder
        pdfs += [f]

# - - - - - - - -
print(pdfs) # prints filenames in current folder - for testing if the files are in the right place
# - - - - - - - -

#start gathering data
os.chdir(starting_folder)
f_sum = open('G3_summary.csv', "w") #file for summary
os.chdir(starting_folder+"/Reports")

mylodge = LodgeData()#creates lodge object for data management

current_lodge_name = pdfs[0].split()[0]

for i, file in enumerate(pdfs):   
    lodge_name = file.split()[0]#pointer for current lodge name
    print(lodge_name)
    #       writes data to file
    # if found new lodge or the last report, save the info before gathering new data
    if (current_lodge_name.lower() != lodge_name.lower()): #only writes to file upon finding a new lodge or the end of the list
        f_sum.write("\n" + current_lodge_name + "\n")#data gathered belongs to previous lodge, not  the one we just stumbled on
        mylodge.write_info(f_sum, starting_folder)
        mylodge.wipe() #resets the lodge data for the next round
        current_lodge_name = lodge_name#resets the lodge name for the next round
        
    #       gathers data
    dfs = read_pdf(file,pages="all") #returns list of dataframes representing pages from a file we know is a pdf
    
    """if lodge_name == "Nakona":
        try:
            print(dfs)
        except:
            pass"""
    
    #easy data to grab, independant try/except for issue isolation (probs overkill)
    try:
        mylodge.years += [dfs[0].iloc[18,0]]
    except: 
        mylodge.years += [" "]
    try:
        mylodge.election += [dfs[1].iloc[12,1]]
    except:
        mylodge.election += [" "]
    if mylodge.years[-1] == "2023":
        try:
            mylodge.induction += [dfs[1].iloc[25,1]]
        except:
            mylodge.induction += [" "]
        try:
            mylodge.activation += [dfs[1].iloc[39,1]]
        except:
            mylodge.activation += [" "]
    else:
        try:
            mylodge.induction += [dfs[1].iloc[26,1]]
        except:
            mylodge.induction += [" "]
        try:
            mylodge.activation += [dfs[1].iloc[40,1]]
        except:
            mylodge.activation += [" "]
    try:
        mylodge.m_retention += [dfs[2].iloc[14,1]]
    except:
        mylodge.m_retention += [" "]
    try:
        mylodge.m_growth += [str(int(str(dfs[2].iloc[21,1]).replace(",", "")) - int(str(dfs[2].iloc[21,6]).replace(",","")))]
    except:
        #print('typerror')
        mylodge.m_growth += [" "]
    try:
        mylodge.e_participation += [dfs[2].iloc[34,1]]
    except:
        mylodge.e_participation += [" "]
    try:
        mylodge.bro_rate += [dfs[3].iloc[12,1]]
    except:
        mylodge.bro_rate += [" "]
    try:
        mylodge.dol_mem += [dfs[3].iloc[24,1]]
    except:
        mylodge.dol_mem += [" "]
    try:
        mylodge.ser_mem += [dfs[3].iloc[39,1]]
    except:
        mylodge.ser_mem += [" "]
    try:
        mylodge.training += [dfs[4].iloc[12,1]]
    except:
        mylodge.training += [" "]
    try:
        y_ordeal_num = int(str(dfs[0].iloc[28,1]).replace(",", ""))
        mylodge.y_ordeal += [y_ordeal_num]
    except:
        mylodge.y_ordeal += [" "]
    try:
        y_bro_num = int(str(dfs[0].iloc[29,1]).replace(",", ""))
        mylodge.y_bro += [y_bro_num]
    except:
        mylodge.y_bro += [" "]
    try:
        y_vigil_num = int(str(dfs[0].iloc[30,1]).replace(",", ""))
        mylodge.y_vigil += [y_vigil_num]
    except:
        mylodge.y_vigil += [" "]
    try:
        ya_ordeal_num = int(str(dfs[0].iloc[28,2]).replace(",", ""))
        mylodge.ya_ordeal += [ya_ordeal_num]
    except:
        mylodge.ya_ordeal += [" "]
    try:
        ya_bro_num = int(str(dfs[0].iloc[29,2]).replace(",", ""))
        mylodge.ya_bro += [ya_bro_num]
    except:
        mylodge.ya_bro += [" "]
    try:
        ya_vigil_num = int(str(dfs[0].iloc[30,2]).replace(",", ""))
        mylodge.ya_vigil += [ya_vigil_num]
    except:
        mylodge.ya_vigil += [" "]
    try:
        a_ordeal_num = int(str(dfs[0].iloc[28,3]).replace(",", ""))
        mylodge.a_ordeal += [a_ordeal_num]
    except:
        mylodge.a_ordeal += [" "]
    try:
        a_bro_num = int(str(dfs[0].iloc[29,3]).replace(",", ""))
        mylodge.a_bro += [a_bro_num]
    except:
        mylodge.a_bro += [" "]
    try:
        a_vigil_num = int(str(dfs[0].iloc[30,3]).replace(",", ""))
        mylodge.a_vigil += [a_vigil_num]
    except:
        mylodge.a_vigil += [" "]
    try:
        mylodge.y_all += [str(y_ordeal_num + y_bro_num + y_vigil_num)]
    except:
        mylodge.y_all += [" "]
    try:
        mylodge.ya_all += [str(ya_ordeal_num + ya_bro_num + ya_vigil_num)]
    except:
        mylodge.ya_all += [" "]
    try:
        mylodge.a_all += [str(a_ordeal_num + a_bro_num + a_vigil_num)]
    except:
        mylodge.a_all += [" "]
    try:
        mylodge.ordeal_all += [str(a_ordeal_num + ya_ordeal_num + y_ordeal_num)]
    except:
        mylodge.ordeal_all += [" "]
    try:
        mylodge.bro_all += [str(a_bro_num + ya_bro_num + y_bro_num)]
    except:
        mylodge.bro_all += [" "]
    try:
        mylodge.vigil_all += [str(a_vigil_num + ya_vigil_num + y_vigil_num)]
    except:
        mylodge.vigil_all += [" "]
    try:
        mylodge.lodge_all += [str(sum([y_ordeal_num, ya_ordeal_num, a_ordeal_num, y_bro_num, ya_bro_num, a_bro_num, y_vigil_num, ya_vigil_num, a_vigil_num]))]
    except:
        mylodge.lodge_all += [" "]
    
    
    if (i == len(pdfs) - 1): #the end of the list
        f_sum.write("\n" + lodge_name + "\n")#data gathered belongs to current lodge
        mylodge.write_info(f_sum, starting_folder)
        
   
      
      
  
print("done.")
f_sum.close()


#a joke for those that made it this far
"""
print("def is_even(num:int) -> bool:")
for i in range(100):
    print(f'    if num == {i}:')
    print(f'        return {"True" if i%2 ==0 else "False"}')
"""

