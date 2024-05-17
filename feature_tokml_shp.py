import arcpy 
import os 

#add your geodatabase path here
arcpy.env.workspace=r"Z:\GIS\1_Request\Lands\PowellLakeFloatHomes\PowellLake_FloatHome\PowellLake_FloatHome.gdb"
arcpy.env.overwriteOutput=True

#add output folder path here
output_folder = r"\\sfp.idir.bcgov\s164\S63097\Share\GIS\1_Request\Lands\PowellLakeFloatHomes\3_export_data"
#let's check what's in the database
fc = arcpy.ListFeatureClasses() 
#for loop to loop through the database
for f in fc: 
    #layer to kml requires layers, so let's make some layers
    arcpy.management.MakeFeatureLayer(f,'out_feature')
    #making file names
    filename =os.path.join(output_folder, f + "."+"kml")
    #double checking file name names
    print(filename)
    #let's make the layers
    kml = arcpy.conversion.LayerToKML('out_feature',filename)
    #let's make the shapefiles too
    arcpy.conversion.FeatureClassToShapefile(f,output_folder)
print("completed")