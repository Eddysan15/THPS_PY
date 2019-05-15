#################################
#Author: Eddie Sanchez 
#Date: May 10, 2019
#Assignment: Evulatation of openpyxl 
#Final Project Python Script 2
#Demo2 

######################################

#documentation used for openpyxl: https://www.geeksforgeeks.org/python-arithmetic-operations-in-excel-file-using-openpyxl/
#documentation used for pands: https://pandas.pydata.org/pandas-docs/stable/reference/api/pandas.read_excel.html

##########################################
#Uses Anconda to process files 
#to run code first use Anaconda C:-> Program Files-> Anaconda3 -> python.exe
#########################################

#Libraries Used
import openpyxl
import pandas as pd 
import openpyxl as px

################################################################
#DelNorte2016 Excel File Manipulation 
##################################################################
#array methods used to maninpate excel spread sheet  
#An array is a list used to store temporarl used to define and excute in script 

######################################################
#Set up File Path 
Outputpath="C:\\Users\\ees407\\Desktop\\Python Final Project\\Outputs\\"

data_file=pd.read_excel(Outputpath+'DelNorte2016.xlsx') #used to read the file 

############################################
#set up Outputpath for Jim to make it easier to process

Outputpath="C:\\Users\\ees407\\Desktop\\Python Final Project\\Outputs\\"

##########################################
#define arces

acres=data_file['ACRES'] #Arces is defined to be used throughout the script 

######################################
# define the Spilt 

num= data_file['THP_NUM'] #spilt is used to know where the last digits of the identification number changes

#import the THP_NUM into the first list 
nums= list(data_file['THP_NUM'])

#########################################
#first array
my_list= []
count_list=[]
for x in nums: 
    try: 
        index=my_list.index(x)
    except:
        count_list.append(0)
        my_list.append(x)

#print(my_list)

index=0
##################################
#second array 
while (index<len(nums)): 
    a1=nums[index]
    
    #print(a1)
    b1= acres[index]
    #print(b1)
    
    index2=my_list.index(a1)
    count_list[index2]+=b1
   # print(count_list)
    index+=1
#########################
#loop that goes through both array 
for i in count_list: 
    try: 
        index=count_list.index(i)
    except:
        my_list.append(0)
        count_list.append(i)

       
#print(count_list)

index=0

#line of code that prints in columns 
LineCount=0
wb=openpyxl.load_workbook(Outputpath+"DelNorte2016.xlsx")
TheFile=(Outputpath+"DelNorte2016.xlsx","w")
sheet= wb.active
sheet=wb.get_sheet_by_name('DelNorte2016')
sheet["V1"] = ("Total Arces")
sheet["V2"] = (format(count_list))
print(count_list)
count=0
for areacount in count_list: 
    thecell=("V"+format(count+2))
    
    try:
        sheet[thecell]=count_list[int(count)]
    except:
        pass
    
    count+=1
    
#save work using xl
    
wb.save(Outputpath+"DelNorte2016.xlsx")

#########################################
#the results should be in the out put folder in the same excel sheet used to manipulate numbers 