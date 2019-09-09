# coding=gbk
import ogr, os
from osgeo import gdal

#chinese directory support
gdal.SetConfigOption('GDAL_FILENAME_IS_UTF8','NO')
#chinese attribute list support
gdal.SetConfigOption('SHAPE_ENCODING','')
#register
ogr.RegisterAll()


def point_to_shp(inputlist,outSHPfn):
    fieldType = ogr.OFTString
    # Create the output shapefile
    shpDriver = ogr.GetDriverByName("ESRI Shapefile")
    if os.path.exists(outSHPfn):
        shpDriver.DeleteDataSource(outSHPfn)
    outDataSource = shpDriver.CreateDataSource(outSHPfn)
    outLayer = outDataSource.CreateLayer(outSHPfn, geom_type=ogr.wkbPoint)


    # create a field
    idField1 = ogr.FieldDefn('val1', fieldType)
    idField2 = ogr.FieldDefn('val2', fieldType)
    idField3 = ogr.FieldDefn('val3', fieldType)
    idField4 = ogr.FieldDefn('val4', fieldType)
    idField5 = ogr.FieldDefn('val5', fieldType)

    outLayer.CreateField(idField1)
    outLayer.CreateField(idField2)
    outLayer.CreateField(idField3)
    outLayer.CreateField(idField4)
    outLayer.CreateField(idField5)

    # Create the feature and set values

    for i in range(len(inputlist)):
        point = ogr.Geometry(ogr.wkbPoint)
        point.AddPoint(inputlist[i][0],inputlist[i][1])

        featureDefn = outLayer.GetLayerDefn()
        outFeature = ogr.Feature(featureDefn)
        outFeature.SetGeometry(point)
        outFeature.SetField('val1', inputlist[i][2].encode('gbk'))
        outFeature.SetField('val2', inputlist[i][3].encode('gbk'))
        outFeature.SetField('val3', inputlist[i][4].encode('gbk'))
        outFeature.SetField('val4', inputlist[i][5].encode('gbk'))
        outFeature.SetField('val5', inputlist[i][6].encode('gbk'))

        outLayer.CreateFeature(outFeature)
        outFeature.Destroy()



    #create point geometry
    # point = ogr.Geometry(ogr.wkbPoint)
    # point.AddPoint(lon,lat)
    #
    # featureDefn = outLayer.GetLayerDefn()
    # outFeature = ogr.Feature(featureDefn)
    # outFeature.SetGeometry(point)
    # outFeature.SetField('val1', val1.encode('gb2312'))
    # outFeature.SetField('val2', val2.encode('gb2312'))
    # outFeature.SetField('val3', val3.encode('gb2312'))
    # outFeature.SetField('val4', val4.encode('gb2312'))
    # outFeature.SetField('val5', val5.encode('gb2312'))
    #
    # outLayer.CreateFeature(outFeature)
    outFeature = None


    # # create a field2
    # idField2 = ogr.FieldDefn('val2', fieldType)
    # outLayer.CreateField(idField2)
    # # Create the feature and set values
    # featureDefn = outLayer.GetLayerDefn()
    # outFeature = ogr.Feature(featureDefn)
    # # outFeature.SetGeometry(point)
    # outFeature.SetField('val2', '2')
    # outLayer.CreateFeature(outFeature)
    # outFeature = None
    #
    # # create a field3
    # idField3 = ogr.FieldDefn('val3', fieldType)
    # outLayer.CreateField(idField3)
    # # Create the feature and set values
    # featureDefn = outLayer.GetLayerDefn()
    # outFeature = ogr.Feature(featureDefn)
    # # outFeature.SetGeometry(point)
    # outFeature.SetField('val3', '3')
    # outLayer.CreateFeature(outFeature)
    # outFeature = None


def merge_point(outputMergefn,directory):
    # outputMergefn = 'merge.shp'
    # directory = "test_temp/"
    # fileStartsWith = 'test'
    fileEndsWith = '.shp'
    driverName = 'ESRI Shapefile'
    geometryType = ogr.wkbPoint

    out_driver = ogr.GetDriverByName(driverName)
    if os.path.exists(outputMergefn):
        out_driver.DeleteDataSource(outputMergefn)
    out_ds = out_driver.CreateDataSource(outputMergefn)
    out_layer = out_ds.CreateLayer(outputMergefn, geom_type=geometryType)

    fileList = os.listdir(directory)

    for file in fileList:
        if file.endswith(fileEndsWith):
            # print file
            ds = ogr.Open(directory+file,1)

            lyr = ds.GetLayer()
            for feat in lyr:
                # oLayer = ds.GetLayerByIndex(0)
                # oNewField = ogr.FieldDefn("val",ogr.OFTString)
                # oLayer.CreateField(oNewField,1)
                # feat.SetField(0,'val'.encode('gb2312'))
                # oLayer.SetFeature(feat)

                out_feat = ogr.Feature(out_layer.GetLayerDefn())
                out_feat.SetGeometry(feat.GetGeometryRef().Clone())
                out_layer.CreateFeature(out_feat)
                # feat.SetField()
                out_feat = None
                out_layer.SyncToDisk()


def write_shp_attrib(shp,arrtib_content):
    #attrib_content should be a list type
    ds = ogr.Open(shp,1)
    oLayer = ds.GetLayerByIndex(0)
    oNewField = ogr.FieldDefn("val",ogr.OFTString)
    oLayer.CreateField(oNewField,1)
    i=0
    for feature in oLayer:
        NumOfDefn = feature.GetFieldCount()
        feature.SetField(NumOfDefn-1,arrtib_content[i].encode('gb2312'))
        # print arrtib_content[i]
        oLayer.SetFeature(feature)
        i += 1


# point_to_shp(12,34,'..\\test\\aaa.shp')
# point_to_shp(32,22,'..\\test\\aab.shp')
# merge_point('../merge.shp','..\\test\\')
# if __name__ == '__main__':
#     input_list=[[1,2,'3','3','3','3','3'],[2,3,'3','3','3','3','3'],[2,4,'3','3','3','3','3'],[3,2,'3','3','3','3','3']]
#     point_to_shp(input_list,'../test.shp')