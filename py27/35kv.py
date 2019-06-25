# coding=utf-8

from osgeo import ogr
import os
this_root = os.getcwd()+'\\..\\'
import gdal
import xlrd
import time
import re
import numpy as np
import log_process



def mk_dir(dir):
    if not os.path.isdir(dir):
        os.makedirs(dir)


def line_to_shp(inputlist,outSHPfn):
    ############重要#################
    gdal.SetConfigOption("SHAPE_ENCODING", "GBK")
    ############重要#################
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
        outFeature.SetField('val1', val1)
        outFeature.SetField('val2', val2)
        outFeature.SetField('val3', val3)
        outFeature.SetField('val4', val4)
        outFeature.SetField('val5', val5)
        outLayer.CreateFeature(outFeature)
        outFeature.Destroy()
        line = ogr.Geometry(ogr.wkbLineString)
        outFeature = None


def point_to_shp(inputlist,outSHPfn):
    # gdal.SetConfigOption("GDAL_FILENAME_IS_UTF8", "YES")
    ############重要#################
    gdal.SetConfigOption("SHAPE_ENCODING", "GBK")
    ############重要#################
    fieldType = ogr.OFTString
    # fieldType.set
    # Create the output shapefile
    shpDriver = ogr.GetDriverByName("ESRI Shapefile")
    if os.path.exists(outSHPfn):
        shpDriver.DeleteDataSource(outSHPfn)
    outDataSource = shpDriver.CreateDataSource(outSHPfn)
    outLayer = outDataSource.CreateLayer(outSHPfn, geom_type=ogr.wkbPoint)


    # create a field
    idField1 = ogr.FieldDefn('val1',fieldType)
    idField2 = ogr.FieldDefn('val2', fieldType)
    idField3 = ogr.FieldDefn('val3', fieldType)
    # idField4 = ogr.FieldDefn('val4', fieldType)
    # idField5 = ogr.FieldDefn('val5', fieldType)

    outLayer.CreateField(idField1)
    outLayer.CreateField(idField2)
    outLayer.CreateField(idField3)
    # outLayer.CreateField(idField4)
    # outLayer.CreateField(idField5)

    # Create the feature and set values

    for i in range(len(inputlist)):
        point = ogr.Geometry(ogr.wkbPoint)
        point.AddPoint(inputlist[i][0],inputlist[i][1])

        featureDefn = outLayer.GetLayerDefn()
        outFeature = ogr.Feature(featureDefn)
        outFeature.SetGeometry(point)
        outFeature.SetField('val1', inputlist[i][2])
        outFeature.SetField('val2', inputlist[i][3])
        outFeature.SetField('val3', inputlist[i][4])
        # outFeature.SetField('val4', inputlist[i][5].encode('gbk'))
        # outFeature.SetField('val5', inputlist[i][6].encode('gbk'))

        outLayer.CreateFeature(outFeature)
        outFeature.Destroy()
    outFeature = None


def _35_kv_ganta_excel():
    f_excel = this_root + '190509\\35kV线路\\民权35千伏杆塔序号.xls'
    bk = xlrd.open_workbook(f_excel)
    sh = bk.sheet_by_index(0)
    nrows = sh.nrows
    ganta = []
    ganta_num = []
    xx = u"([0-9]+)"
    pattern = re.compile(xx)

    for i in range(nrows):
        ganta_attrib = sh.cell_value(i, 3)
        ganta_name = sh.cell(i, 0)
        find = pattern.findall(ganta_name.value)
        if len(find) > 0:
            ganta_num_i = find[1]

        else:
            ganta_num_i = None

        if ganta_attrib == '耐张'.decode('gbk') and ganta_num_i:
            ganta.append(ganta_name.value)
            ganta_num.append(ganta_num_i)

    ganta_dic = {}
    for i in range(len(ganta)):
        ganta_dic[ganta[i]] = ganta_num[i]

    return ganta_dic



# def gen_zhushang_bianyaqi_shp():
#     ganta = gen_35_kv_ganta_excel()
#     inlist = []
#     for i in ganta:
#         print(i)
#     pass



def gen_35kv_line_shp():
    '''
    composite shp
    将xian_dic_sort生成的shp合成为1个shp
    作为annotation
    :return:
    '''
    gdal.SetConfigOption("SHAPE_ENCODING", "GBK")
    fdir = this_root + '\\190509\\35kV线路\\dwg_to_shp\\'
    flist = os.listdir(fdir)
    time_init = time.time()
    time_i = 0
    inlist = []
    for folder in flist:
        # print(folder.decode('gbk'))
        # exit()
        time_start = time.time()
        # print(folder.decode('gbk'))
        shp_dir = fdir+folder+'\\'

        shp_list = os.listdir(shp_dir)

        for shp in shp_list:
            # print(shp.decode('gbk'))
            # exit()
            if shp.endswith('dwg_Polyline.shp'):
                # print(shp.decode('gbk'))
                # exit()
                daShapefile = shp_dir + shp
                # print(daShapefile.decode('gbk'))
                driver = ogr.GetDriverByName("ESRI Shapefile")
                dataSource = driver.Open(daShapefile.decode('gbk'), 0)
                layer = dataSource.GetLayer()
                # inlist_i = []
                for feature in layer:
                    geom = feature.GetGeometryRef()
                    Points = geom.GetPoints()
                    if Points:
                        for i in range(len(Points)):
                            if i == len(Points)-1:
                                break
                            inlist.append([Points[i],Points[i+1],'','','','',''])
                        # print(Points)
                        # print(inlist_i)
                    # exit()
                    # exit()
                    # val1 = feature.GetField("val1")
                    # val2 = feature.GetField("val2")
                    # val3 = feature.GetField("val3")
                    # val4 = feature.GetField("val4")
                    # inlist.append([Points[0],Points[1],val1,val2,val3,val4,''])

        time_end = time.time()
        time_i += 1
        log_process.process_bar(time_i,len(flist),time_init,time_start,time_end)
    output_fn = this_root + '190509\\35kV线路\\35_kv_line.shp'.decode('gbk').encode('utf-8')
    print('exporting line shp...')
    line_to_shp(inlist, output_fn)

    pass





def gen_shp_dic():

    fdir = this_root+u'190509\\35kV线路\\dwg_to_shp\\'
    flist = os.listdir(fdir)
    shp_dic = {}

    time_init = time.time()
    flag = 0
    for folder in flist:
        time_start = time.time()
        print(folder)
        shp_list = os.listdir(fdir+folder)
        for shp in shp_list:
            if shp.endswith('Annotation.shp'):
                daShapefile = fdir+folder+'\\'+shp
                driver = ogr.GetDriverByName("ESRI Shapefile")
                dataSource = driver.Open(daShapefile, 0)
                layer = dataSource.GetLayer()
                for feature in layer:
                    geom = feature.GetGeometryRef()
                    x = geom.GetX()
                    y = geom.GetY()
                    name = feature.GetField("RefName")
                    name_gbk = name.decode('utf-8')
                    shp_dic[name_gbk] = [x,y]
                    # print x,y,name_gbk
                    # time.sleep(0.5)
        time_end = time.time()
        log_process.process_bar(flag,len(flist),time_init,time_start,time_end)
        flag+=1
    print('saving shp dic...')
    np.save(this_root+'190509\\35kV线路\\shp_dic',shp_dic)
        # exit()




def foo():
    '''
    看shp_dic字典里是啥
    :return:
    '''
    npy = this_root+'190509\\35kV线路\\shp_dic.npy'
    shp_dic = np.load(npy).item()
    shp_dic = dict(shp_dic)
    # print(len(shp_dic))
    for i in shp_dic:
        print i,shp_dic[i]
        # time.sleep(0.5)


def gen_35kv_ganta_shp():
    naizhang = _35_kv_ganta_excel()
    npy = this_root + '190509\\35kV线路\\shp_dic.npy'
    shp_dic = np.load(npy).item()
    shp_dic = dict(shp_dic)
    inlist = []
    for ganta in naizhang:
        try:
            x = shp_dic[ganta][0]
            y = shp_dic[ganta][1]
            val1 = ganta
            val2 = naizhang[ganta]
            inlist.append([x,y,val1,val2,''])
            # print(x)
        except Exception as e:
            print ganta,'error'
            # print(e)
    for i in inlist:
        print(i)
    point_to_shp(inlist,this_root + '190509\\35kV线路\\naizhang_ganta.shp'.decode('gbk').encode('utf-8'))


def biandianzhan_shp():
    f_excel = this_root + '190509\\35kV线路\\变电站.xls'
    bk = xlrd.open_workbook(f_excel)
    sh = bk.sheet_by_index(0)
    nrows = sh.nrows
    inlist = []
    for i in range(nrows):
        biandianzhan_name = sh.cell_value(i, 0)
        x = sh.cell_value(i, 1)
        y = sh.cell_value(i, 2)
        try:
            inlist.append([float(x),float(y),biandianzhan_name,'',''])
        except:
            pass
    point_to_shp(inlist,this_root+'190509\\35kV线路\\biandianzhan.shp'.decode('gbk').encode('utf-8'))
    pass



def main():
    biandianzhan_shp()
    # pass
    # a='12312'
    # b='2'
    # print(a-b)

if __name__ == '__main__':
    main()