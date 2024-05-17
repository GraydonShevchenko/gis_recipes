"""
Author: Amanda Lu
Updated: 2024-05-17
Description: This script grabs applications or tenures from TANTALIS 
and outputs a feature, kml and/or excel spreadsheet    

Variables: 
- Access to oracle database
- Area of Interest created in form of a shapefile. 

Clients:
Land Officers and anyone with an interest on the land. 
"""

#import libraries 
import arcpy
import getpass
import os
import shutil
import time
import pandas as pd


# Variables for database connection
input_database = r"bcgw_bcgov.sde" # Change to reflect your connection
instance = r"bcgw.bcgov/idwprod1.bcgov" # Database instance
srid = 3005 # EPSG code for projection. Almost always BC Albers/3005

# selection 1- things to change
# Name to open the query layer as
out_layer_name = r"tantalis_selection" 
# Uniqe Object ID field according to the sql query
oid_field = "FILE_NBR"
#location of your area of interest
aoi = r" " 
#location of results
resultant_path = r" "
#name of geodatabase
work_gdb = r" "


#creating some database variables
input_database_fullpath = resultant_path + "\\" + input_database
gdb_path = resultant_path+"\\"+work_gdb
out_path = gdb_path + "\\"+'tantalis_selection'


# Check geodatabse
#does the database exist, if yes delete
if arcpy.Exists(gdb_path):
    arcpy.Delete_management(gdb_path)
    arcpy.AddMessage(arcpy.GetMessages())
    print("Geodatabase Deleted")
#if there isn't existing geodatabse, create a new one
if not os.path.exists(os.path.dirname(gdb_path)):
    os.mkdir(os.path.dirname(gdb_path))
    print("new geodatabase created")

arcpy.CreateFileGDB_management(os.path.dirname(gdb_path),os.path.basename(gdb_path))
arcpy.AddMessage(arcpy.GetMessages())
arcpy.env.overwriteOutput=True
arcpy.env.workspace = work_gdb

#ask user to enter Oracle credential
user = str(input("Enter your username:\n"))
print("Connecting to {} as user {}".format(instance,user)) 
passwd = str(input("Enter your password:\n"))
print('password entered')

if os.path.exists(input_database_fullpath):
    delete_connection = input("The database connection {} already exists. Delete? y/n".format(input_database))
    if delete_connection.lower() == "y":
        os.remove(input_database_fullpath)

time_initial=time.time()

#creating connection to bcgw
bcgw = arcpy.management.CreateDatabaseConnection(
    resultant_path, 
    input_database, 
    "ORACLE", 
    instance, 
    "DATABASE_AUTH", 
    user,
    passwd, 
    "DO_NOT_SAVE_USERNAME")

# selection 2- things to change
#sql statement to configure with parameters 
query ="""
"""
print('sql statement successful')
#feeding the sql statement into the query 
geometryLayer = arcpy.management.MakeQueryLayer(input_database_fullpath, out_layer_name, query,oid_field ,'POLYGON', srid)
print('layer made')
selection = arcpy.management.SelectLayerByLocation(geometryLayer, 'INTERSECT',aoi)
g_g = arcpy.conversion.ExportFeatures(geometryLayer,out_path)

#making shapefile
output_shapefile = f"{resultant_path}/tantalis_selection.shp"
arcpy.CopyFeatures_management(g_g,output_shapefile)
print(f"'{output_shapefile}"" shapefile created")
#creating excel spreadsheet
output_spreadsheet= f"{resultant_path}/tantalis_selection_foiSC.xlsx"
arcpy.conversion.TableToExcel(g_g,output_spreadsheet)
print(f"'{output_spreadsheet}"" excel created")

#let's make some kmls 
#making file names
kml_filename =os.path.join(resultant_path, out_layer_name + "."+"kml")
arcpy.conversion.LayerToKML(geometryLayer, kml_filename)
print(f"'{kml_filename}' KML created")        
      

print('process completed')

os.remove(input_database_fullpath)
print("Successfully removed {}".format(input_database))

time_end=time.time()
time_delta=time_end-time_initial
time_min=int(time_delta//60)
print("Takes {} minutes to run\n".format(time_min))
