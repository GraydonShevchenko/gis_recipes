"""
Author: Amanda Lu
Updated: 2023-10-25
Description: This script grabs authorizations from an area of interest.     

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
input_database = r"" # Change to reflect your connection
instance = r"" # Database instance
srid = 3005 # EPSG code for projection. Almost always BC Albers/3005

#things to change
out_layer_name = r"" # Name to open the query layer as
oid_field = "" # Unique Object ID field according to the sql query
aoi = r"" #point python to your area of interest in shapefile
resultant_path = r"" #point python to a spot to output the results
work_gdb = r"query_aoi.gdb"#working geodatabase


# Parse inputs
input_database_fullpath = resultant_path + "\\" + input_database
arcpy.env.overwriteOutput=True

# Check file status
if not os.path.exists(resultant_path):
    arcpy.management.CreateFileGDB(resultant_path, work_gdb)

if os.path.exists(resultant_path + "\\" + work_gdb):
    delete_temp_gdb = input("Temporary geodatabase: {}\n\nThe temporary geodatabase already exists. Delete? y/n".format(work_gdb))
    if delete_temp_gdb.lower() == "y":
        shutil.rmtree(resultant_path + "\\" + work_gdb)
    else:
        print("Either rename the variables temp_out_path or temp_out_gdb or delete the existing temporary geodatabase")

user = ''
# str(input("Enter your username:\n"))
# print("Connecting to {} as user {}".format(instance,user)) 
passwd = ''
#str(input("Enter your password:\n"))

if os.path.exists(input_database_fullpath):
    delete_connection = input("The database connection {} already exists. Delete? y/n".format(input_database))
    if delete_connection.lower() == "y":
        os.remove(input_database_fullpath)

time_initial=time.time()

# Bake in a 5 minute cooldown
bcgw = arcpy.management.CreateDatabaseConnection(
    resultant_path, 
    input_database, 
    "ORACLE", 
    instance, 
    "DATABASE_AUTH", 
    user,
    passwd, 
    "DO_NOT_SAVE_USERNAME")

print(arcpy.Describe(input_database_fullpath).dataType)

query ="""

"""


geometryLayer = arcpy.management.MakeQueryLayer(input_database_fullpath, out_layer_name, query, oid_field, "POLYGON", srid)
gc= arcpy.management.SelectLayerByLocation(geometryLayer,"INTERSECT", aoi)


gc_kml = arcpy.management.MakeFeatureLayer(gc,'kml_tantalis_selection')
#creating a shapefile from the select by location 
gc_int_copy = arcpy.management.CopyFeatures(gc,resultant_path)
output_shapefile = f"{resultant_path}/tantalis_selection_foiSC.shp"
arcpy.CopyFeatures_management(gc_int_copy,output_shapefile)
output_spreadsheet= f"{resultant_path}/tantalis_selection_foiSC.xlsx"
arcpy.conversion.TableToExcel(gc_int_copy,output_spreadsheet)

#let's make some kmls 
#making file names
kml_filename =os.path.join(resultant_path, out_layer_name + "."+"kml")
arcpy.conversion.LayerToKML(gc_kml, kml_filename)
print(f"'{kml_filename}' KML created")        
      

print('process completed')

os.remove(input_database_fullpath)
print("Successfully removed {}".format(input_database))

time_end=time.time()
time_delta=time_end-time_initial
time_min=int(time_delta//60)
print("Takes {} minutes to run\n".format(time_min))




