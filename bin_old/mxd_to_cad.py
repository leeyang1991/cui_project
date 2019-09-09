# coding=gbk
import arcpy
import os

dir = 'D:\\project13\\output\\³ׯ401\\'.decode('gbk')

features = []
output_type = "DWG_R2010"
output_file = dir+"/cad.dwg"
flist = os.listdir(dir)
for i in flist:
    if i.endswith('shp'):
        features.append(dir+''+i[:-4])
for i in features:
    print i
feature = 'D:\\project13\\output\\³ׯ401\\bianyaqi.shp'
arcpy.ExportCAD_conversion(feature,output_type,output_file,"USE_FILENAMES_IN_TABLES", "OVERWRITE_EXISTING_FILES", "")

