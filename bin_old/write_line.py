# coding = utf-8
import ogr, os
from osgeo import gdal
from math import radians, cos, sin, asin, sqrt
import math

#chinese directory support
gdal.SetConfigOption('GDAL_FILENAME_IS_UTF8','NO')
#chinese attribute list support
gdal.SetConfigOption('SHAPE_ENCODING','')
#register
ogr.RegisterAll()

def haversine(lon1, lat1, lon2, lat2):
    """
    Calculate the great circle distance between two points
    on the earth (specified in decimal degrees)
    """
    lon1, lat1, lon2, lat2 = map(radians, [lon1, lat1, lon2, lat2])
    dlon = lon2 - lon1
    dlat = lat2 - lat1
    a = sin(dlat/2)**2 + cos(lat1) * cos(lat2) * sin(dlon/2)**2
    c = 2 * asin(sqrt(a))
    r = 6371
    return c * r * 1000



def rad(d):
    return d*math.pi/180
def GetDistance(lng1,lat1,lng2,lat2):
    radLat1=rad(lat1)
    radLat2=rad(lat2)
    a=radLat1-radLat2
    b=rad(lng1)-rad(lng2)
    s=2 * math.asin(math.sqrt(math.pow(math.sin(a/2),2) +math.cos(radLat1)*math.cos(radLat2)*math.pow(math.sin(b/2),2)))
    s = s *6378.137*1000
    distance=round(s,4)
    return distance

    #### from https://kite.com/python/answers/how-to-find-the-distance-between-two-lat-long-coordinates-in-python
    # R = 6373.0
    # dlon = lng2 - lng1
    # dlat = lat2 - lat1
    # a = math.sin(dlat / 2) ** 2 + math.cos(lat1) * math.cos(lat2) * math.sin(dlon / 2) ** 2
    # c = 2 * math.atan2(math.sqrt(a), math.sqrt(1 - a))
    # distance = R * c
    # # print distance
    # # exit()
    # return distance
    pass


def line_to_shp(inputlist,outSHPfn):

    # start,end,outSHPfn,val1,val2,val3,val4,val5
    # _,_,_,_=start[1],start[0],end[0],end[1]

    shpDriver = ogr.GetDriverByName("ESRI Shapefile")
    if os.path.exists(outSHPfn):
        shpDriver.DeleteDataSource(outSHPfn)
    outDataSource = shpDriver.CreateDataSource(outSHPfn)
    outLayer = outDataSource.CreateLayer(outSHPfn, geom_type=ogr.wkbLineString)

    #create line geometry
    line = ogr.Geometry(ogr.wkbLineString)

    # create a field
    fieldType = ogr.OFTString
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

    for i in range(len(inputlist)):
        start = inputlist[i][0]
        end = inputlist[i][1]
        val1 = inputlist[i][2]
        val2 = inputlist[i][3]
        val3 = inputlist[i][4]
        val4 = inputlist[i][5]
        val5 = inputlist[i][6]

        line.AddPoint(start[0],start[1])
        line.AddPoint(end[0],end[1])

        featureDefn = outLayer.GetLayerDefn()
        outFeature = ogr.Feature(featureDefn)
        outFeature.SetGeometry(line)
        outFeature.SetField('val1', val1.encode('gb2312'))
        outFeature.SetField('val2', val2.encode('gb2312'))
        outFeature.SetField('val3', val3.encode('gb2312'))
        outFeature.SetField('val4', val4.encode('gb2312'))
        outFeature.SetField('val4', val5.encode('gb2312'))
        outLayer.CreateFeature(outFeature)
        outFeature.Destroy()
        line = ogr.Geometry(ogr.wkbLineString)
        outFeature = None


# def poly_line(points,outSHPfn):
#     shpDriver = ogr.GetDriverByName("ESRI Shapefile")
#     if os.path.exists(outSHPfn):
#         shpDriver.DeleteDataSource(outSHPfn)
#     outDataSource = shpDriver.CreateDataSource(outSHPfn)
#     outLayer = outDataSource.CreateLayer(outSHPfn, geom_type=ogr.wkbMultiLineString)
#
#     #create line geometry
#     line = ogr.Geometry(ogr.wkbMultiLineString)
#     fieldType = ogr.OFTString
#     idField = ogr.FieldDefn('val1', fieldType)
#     outLayer.CreateField(idField)
#     for i in points:
#         for j in i:
#             line.AddPoint(j)
#         outFeature.SetField('val1', dist)
#         featureDefn = outLayer.GetLayerDefn()
#         outFeature = ogr.Feature(featureDefn)
#         outFeature.SetGeometry(line)
#         outLayer.CreateFeature(outFeature)
#         outFeature.Destroy()
#         line = ogr.Geometry(ogr.wkbMultiLineString)

# def merge_line():
#     outputMergefn = 'merge.shp'
#     directory = "test_temp/"
#     # fileStartsWith = 'test'
#     fileEndsWith = '.shp'
#     driverName = 'ESRI Shapefile'
#     geometryType = ogr.wkbLineString
#
#     out_driver = ogr.GetDriverByName(driverName)
#     if os.path.exists(outputMergefn):
#         out_driver.DeleteDataSource(outputMergefn)
#     out_ds = out_driver.CreateDataSource(outputMergefn)
#     out_layer = out_ds.CreateLayer(outputMergefn, geom_type=geometryType)
#
#     fileList = os.listdir(directory)
#
#     for file in fileList:
#         if file.endswith(fileEndsWith):
#             print file
#             ds = ogr.Open(directory+file)
#             lyr = ds.GetLayer()
#             for feat in lyr:
#                 out_feat = ogr.Feature(out_layer.GetLayerDefn())
#                 out_feat.SetGeometry(feat.GetGeometryRef().Clone())
#                 out_layer.CreateFeature(out_feat)
#                 out_feat = None
#                 out_layer.SyncToDisk()
#
# def write_shp_attrib(shp,arrtib_content):
#     ds = ogr.Open(shp,1)
#     oLayer = ds.GetLayerByIndex(0)
#     oNewField = ogr.FieldDefn("s",ogr.OFSTFloat32)
#     oLayer.CreateField(oNewField,1)
#     i=0
#     for feature in oLayer:
#         NumOfDefn = feature.GetFieldCount()
#         feature.SetField(NumOfDefn-1,arrtib_content[i])
#         oLayer.SetFeature(feature)
#         i += 1

if __name__ == '__main__':
    # line_to_shp([1.,2.],[3.,4.],'../test/sss.shp','lon','lat','length')
    # points = [[23,43],[22,11],[11,43],[45,87],[89,12],[23,21],[32,11]],[[13,23],[12,23],[15,34],[15,37],[49,32],[63,81],[62,31]]
    # print points[0]
    # print points[1]
    # lines = []
    # for i in range(len(points[0])):
    #     lines.append([points[0][i],points[1][i],'','','','',''])
    #
    #
    # line_to_shp(lines,'../teststatat.shp')
    pass