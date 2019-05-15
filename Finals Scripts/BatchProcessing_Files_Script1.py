
##############################
#Author: Eddie Sanchez 
#Date: May 10, 2019
#Final Python Script 1 that makes the attribute table into an excel (.xlsx) file. 

##################################
###################################
#How to run script
#First Will need to import arcpy and os
#Second make sure that use ArcPro Python, Go to the C drive ->Programs Files->ArcGis-> Pro->bin ->Python->envs->arcgispro-py3->python.exe
#This code runs the output excel files to the output folder
######################################
####################################
#Libraries used
import arcpy 
import os 

arcpy.CheckOutExtension("Spatial") #this is neede to enable the extensions to do the processing 


# This allows us to run the script repeatedly without deleting the intermdiate files
arcpy.env.overwriteOutput=True 


# set the input (orignal files) and output paths (excel Files)

InputPath="C:\\Users\\ees407\\Desktop\\Python Final Project\\Inputs\\Orignals\\"
OutputPath="C:\\Users\\ees407\\Desktop\\Python Final Project\\Inputs\\Outputs\\"
################################
#Set up names to spefic shapefile, needed to make code run easier 
DelNorte2016="C:\\Users\\ees407\\Desktop\\Python Final Project\\Inputs\\Orignals\\DelNorte2016.shp"
DelNorte2017="C:\\Users\\ees407\\Desktop\\Python Final Project\\Inputs\\Orignals\\DelNorte2017.shp"
DelNorte2018="C:\\Users\\ees407\\Desktop\\Python Final Project\\Inputs\\Orignals\\DelNorte2018.shp"
Humboldt2016="C:\\Users\\ees407\\Desktop\\Python Final Project\\Inputs\\Orignals\\Humboldt2016.shp"
Humboldt2017="C:\\Users\\ees407\\Desktop\\Python Final Project\\Inputs\\Orignals\\Humboldt2017.shp"
Humboldt2018="C:\\Users\\ees407\\Desktop\\Python Final Project\\Inputs\\Orignals\\Humboldt2018.shp"
################################

TheList=os.listdir(InputPath) #This allows python to read the list  
################################

#ArcPY using ArcGis Pro For Del Norte 2016-2018

# Replace a layer/table view name with a path to a dataset (which can be a layer file) or create the layer/table view within the script
# The following inputs are layers or table views: "DelNorte2016"
arcpy.TableToExcel_conversion(Input_Table=InputPath+"DelNorte2016.shp", Output_Excel_File="C:\\Users\\ees407\\Desktop\\Python Final Project\\Outputs\\DelNorte2016.xlsx", Use_field_alias_as_column_header="NAME", Use_domain_and_subtype_description="CODE") 
arcpy.TableToExcel_conversion(Input_Table=InputPath+"DelNorte2017.shp", Output_Excel_File="C:\\Users\\ees407\\Desktop\\Python Final Project\\Outputs\\DelNorte2017.xlsx", Use_field_alias_as_column_header="NAME", Use_domain_and_subtype_description="CODE")   
arcpy.TableToExcel_conversion(Input_Table=InputPath+"DelNorte2018.shp", Output_Excel_File="C:\\Users\\ees407\\Desktop\\Python Final Project\\Outputs\\DelNorte2018.xlsx", Use_field_alias_as_column_header="NAME", Use_domain_and_subtype_description="CODE")   

##############################################
#ArcPY using ArcGis Pro For Humboldt 2016-2018
arcpy.TableToExcel_conversion(Input_Table=InputPath+"Humboldt2016.shp", Output_Excel_File="C:\\Users\\ees407\\Desktop\\Python Final Project\\Outputs\\Humboldt2016.xlsx", Use_field_alias_as_column_header="NAME", Use_domain_and_subtype_description="CODE") 
arcpy.TableToExcel_conversion(Input_Table=InputPath+"Humboldt2017.shp", Output_Excel_File="C:\\Users\\ees407\\Desktop\\Python Final Project\\Outputs\\Humboldt2017.xlsx", Use_field_alias_as_column_header="NAME", Use_domain_and_subtype_description="CODE") 
arcpy.TableToExcel_conversion(Input_Table=InputPath+"Humboldt2018.shp", Output_Excel_File="C:\\Users\\ees407\\Desktop\\Python Final Project\\Outputs\\Humboldt2018.xlsx", Use_field_alias_as_column_header="NAME", Use_domain_and_subtype_description="CODE") 
##############################################

#end of script Excel files should be in the output folder