# coding = utf-8
import arcpy

ptList =[[20.000,43.000,'l','llll'],[25.500, 45.085,'daa'],[26.574, 46.025,'dfd'], [28.131, 48.124,'adsf']]
pt = arcpy.Point()
ptGeoms = []
for p in ptList:
    pt.X = p[0]
    pt.Y = p[1]

    ptGeoms.append(arcpy.PointGeometry(pt))

arcpy.CopyFeatures_management(ptGeoms, r"../test/test.shp")