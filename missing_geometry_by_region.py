import os 
from dotenv import load_dotenv, find_dotenv
import arcpy
import random 

#randomgeneratorfortest
rando =str(random.randint(1,100))

input_database = r"bcgw_bcgov.sde" # Change to reflect your connection
instance = r"bcgw.bcgov/idwprod1.bcgov" # Database instance
srid = 3005

out_layer_name = r"tantalis_selection"  # Name to open the query layer as
oid_field = "File_NBR" # Unique Object ID field according to the sql query
resultant_path = r"\\spatialfiles.bcgov\work\srm\nel\Local\Geomatics\Workarea\alu\GR_2024_1149\GR_2024_1149" #point python to a spot to output the results
arcpy.env.overwriteOutput=True

#loading credentials
load_dotenv(r"\\SFP.IDIR.BCGOV\U109\ALU$\secrets\bcgw.env")
user = os.environ.get("BCGW_USER")
password = os.environ.get("BCGW_PASS")
#connection to bcgw
bcgw = arcpy.management.CreateDatabaseConnection(
    resultant_path, 
    input_database, 
    "ORACLE", 
    instance, 
    "DATABASE_AUTH", 
    user,
    password, 
    "DO_NOT_SAVE_USERNAME")
print("new db connected")

work_gdb = r"GR_2024_1149.gdb"#working geodatabase
#location of the active tenures
active = r"\\sfp.idir.bcgov\s164\S63097\Share\GIS\2_scripts\sql-pantry\recipes\active.sql"


#creating some database variables
input_database_fullpath = resultant_path + "\\" + input_database
gdb_path = resultant_path+"\\"+work_gdb

#creating a function that select all active tantalis tenures and applications
def tantalis_selection():
    ten_name = "active" + rando
    file = open(active,'r')
    sql_query = file.read()
    print(sql_query)
    geometryLayer = arcpy.management.MakeQueryLayer(input_database_fullpath, out_layer_name,sql_query, oid_field, "POLYGON", srid)
    print("query made")
    out_path = gdb_path + "\\"+ten_name
    g_g = arcpy.management.CopyFeatures(geometryLayer,out_path)
    print("feature exported")

#let's call the function
tantalis_selection()


#region selection
regions = {"RCB","RKB","RNO","ROM","RSK","RSC","RTO","RWC"}


print('process completed')

os.remove(input_database_fullpath)
print("Successfully removed {}".format(input_database))
