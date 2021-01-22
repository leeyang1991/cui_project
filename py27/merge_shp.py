# coding=gbk

from osgeo import ogr
import os
# this_root = os.getcwd()+'\\..\\'
# this_root = 'E:\\cui\\'
import gdal
import xlrd
import time
import re
import numpy as np
import log_process
# from matplotlib import pyplot as plt
import simple_tkinter as sg
import codecs
import coordinate_transformation as cs
from tqdm import tqdm
import math
import collections
import multiprocessing
import copy_reg
import types
from multiprocessing.pool import ThreadPool as TPool


def line_to_shp(inputlist, outSHPfn):
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

    # create line geometry
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

        line.AddPoint(start[0], start[1])
        line.AddPoint(end[0], end[1])

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

def point_to_shp(inputlist, outSHPfn):
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
        point.AddPoint(inputlist[i][0], inputlist[i][1])

        featureDefn = outLayer.GetLayerDefn()
        outFeature = ogr.Feature(featureDefn)
        outFeature.SetGeometry(point)
        outFeature.SetField('val1', inputlist[i][2])
        outFeature.SetField('val2', inputlist[i][3])
        outFeature.SetField('val3', inputlist[i][4])
        outFeature.SetField('val4', inputlist[i][5])
        outFeature.SetField('val5', inputlist[i][6])

        outLayer.CreateFeature(outFeature)
        outFeature.Destroy()
    outFeature = None


class Merge:
    def __init__(self):
        pass


    def run(self):

        # indir = r'E:\cui\20200516\单条线shp\\'
        # outdir = ur'E:\cui\20200516\data_merge\\'
        # self.merge(indir,outdir)

        # father = r'E:\cui\20200613\data\\'
        father = r'E:\cui\20201010\data\\'
        father_out_dir = ur'E:\cui\20201010\shp_merge\\'
        for folder in os.listdir(father):
            print folder.decode('gbk')
            for indir in os.listdir(father+folder):
                indir_ = father.encode('gbk')+folder+'\\'+indir.decode('gbk').encode('gbk')+'\\'
                outdir =father_out_dir + folder.decode('gbk') + '\\'+indir.decode('gbk')+'\\'
                self.merge(indir_,outdir)



        pass

    def run_gui(self,in_dir,out_dir):
        self.merge(in_dir,out_dir)

    def filter_val(self,val):

        if val == None:
            val = ''
        return val


    def is_line(self,shp):
        line = [
            'line_annotation1.shp',
            'daoxian.shp',
            'dianlan.shp',
        ]
        is_in = False
        for l in line:
            if l in shp:
                is_in = True
        return is_in
        pass

    def merge(self, indir, outdir):
        if not os.path.isdir(outdir):
            os.makedirs(outdir)
        shp_dic = {}
        for folder in os.listdir(indir):
            for shp in os.listdir(os.path.join(indir,folder)):
                if shp.endswith('.shp'):
                    # print folder.decode('gbk'),shp
                    if 'dwg' in shp:
                        continue
                    shp_dic[shp] = []
        for folder in os.listdir(indir):
            for shp in os.listdir(os.path.join(indir, folder)):
                if shp.endswith('.shp'):
                    # print folder.decode('gbk'),shp
                    if 'dwg' in shp:
                        continue
                    shp_dic[shp].append(os.path.join(indir, folder, shp))

        for shp in shp_dic:
            if len(shp_dic[shp]) > 1:
                print shp
                # is_line = self.is_line(shp)
                shp_paths = shp_dic[shp]
                point_inlist = []
                line_inlist = []
                for shp_path_i in shp_paths:
                    daShapefile = shp_path_i
                    # daShapefile = daShapefile.encode('gbk')

                    # print(daShapefile)
                    # print(daShapefile.decode('gbk'))
                    driver = ogr.GetDriverByName("ESRI Shapefile")
                    dataSource = driver.Open(daShapefile.decode('gbk'), 0)
                    layer = dataSource.GetLayer()
                    # inlist_i = []
                    for feature in layer:
                        geom = feature.GetGeometryRef()
                        # geom.
                        Points = geom.GetPoints()
                        val1 = feature.GetField("val1")
                        val2 = feature.GetField("val2")
                        val3 = feature.GetField("val3")
                        # val4 = feature.GetField("val4")

                        val1 = self.filter_val(val1)
                        val2 = self.filter_val(val2)
                        val3 = self.filter_val(val3)
                        # val4 = self.filter_val(val4)
                        # if np.shape(Points) == (1,2):
                        #     print Points
                        if np.shape(Points) == (1,2):
                            # if Points:
                            for i in range(len(Points)):
                                point_inlist.append((Points[i][0], Points[i][1], val1, val2, val3, '', ''))
                        if np.shape(Points) == (2,2):
                            # if Points:
                            for i in range(len(Points)):
                                if i == len(Points) - 1:
                                    break
                                line_inlist.append((Points[i], Points[i + 1], val1, val2, val3, '', ''))

                outfn = os.path.join(outdir,shp)
                outfn = outfn.decode('gbk').encode('utf-8')
                if len(line_inlist) > 0:
                    # outfn = outdir + shp
                    # outfn = outfn.encode('utf-8')
                    line_inlist = list(set(line_inlist))
                    print line_inlist
                    line_to_shp(line_inlist,outfn)
                if len(point_inlist) > 0:
                    # print outdir+shp
                    # outfn = outdir+shp
                    # outfn = outfn.encode('utf-8')
                    point_inlist = list(set(point_inlist))
                    # print point_inlist
                    point_to_shp(point_inlist,outfn)
            # exit()
        pass

    def merge_point_annotation_shp(self, indir, outdir):
        '''
        composite shp
        将xian_dic_sort生成的shp合成为1个shp
        作为annotation
        :return:
        '''
        gdal.SetConfigOption("SHAPE_ENCODING", "GBK")
        # fdir = this_root + '\\190725\\dwg_to_shp\\'
        fdir = indir + '/'
        flist = os.listdir(fdir)
        time_init = time.time()
        time_i = 0
        inlist = []
        for folder in flist:
            # print(folder.decode('gbk'))
            # exit()
            time_start = time.time()
            # print(folder.decode('gbk'))
            shp_dir = fdir + folder + '\\'

            shp_list = os.listdir(shp_dir)

            for shp in shp_list:
                # print(shp.decode('gbk'))
                # print(shp_type+'.shp')
                # exit()
                if shp.endswith('.shp'):
                    # print(shp.decode('gbk'))
                    # print(1)
                    # exit()
                    daShapefile = shp_dir + shp
                    # daShapefile = daShapefile.encode('gbk')
                    print(daShapefile)
                    # print(daShapefile.decode('gbk'))
                    driver = ogr.GetDriverByName("ESRI Shapefile")
                    dataSource = driver.Open(daShapefile.decode('gbk'), 0)
                    layer = dataSource.GetLayer()
                    # inlist_i = []
                    for feature in layer:
                        geom = feature.GetGeometryRef()
                        # geom.
                        Points = geom.GetPoints()
                        val1 = feature.GetField("val1")
                        # val2 = feature.GetField("val2")
                        # val3 = feature.GetField("val3")
                        # print(Points)
                        # exit()
                        if Points:
                            # print(Points)
                            # exit()
                            for i in range(len(Points)):
                                # if i == len(Points)-1:
                                #     break
                                inlist.append([Points[i][0], Points[i][1], val1, '', '', '', ''])
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
            log_process.process_bar(time_i, len(flist), time_init, time_start, time_end)
        # output_fn = this_root + '190725\\shp\\'+shp_type+'_merge.shp'.decode('gbk').encode('utf-8')
        output_fn = outdir + '\\merge_dwg_Annotation.shp'
        # output_fn = output_fn.encode('gbk')
        print('exporting line shp...')
        # print(output_fn)
        # exit()
        point_to_shp1(inlist, output_fn)

        pass

    def merge_line_annotation_shp(self):
        '''
        composite shp
        将xian_dic_sort生成的shp合成为1个shp
        作为annotation
        :return:
        '''
        gdal.SetConfigOption("SHAPE_ENCODING", "GBK")
        # fdir = this_root + '\\190714\\dwg_to_shp\\'
        fdir = r'E:\cui\190905\dwg_to_shp\shao\\'
        output_fn = r'E:\cui\190905\dwg_to_shp\shao_merge\line_annotation.shp'
        flist = os.listdir(fdir)
        time_init = time.time()
        time_i = 0
        inlist = []
        for folder in flist:
            # print(folder.decode('gbk'))
            # exit()
            time_start = time.time()
            # print(folder.decode('gbk'))
            shp_dir = fdir + folder + '\\annotation\\'

            try:
                shp_list = os.listdir(shp_dir)
            except:
                continue

            for shp in shp_list:
                # print(shp.decode('gbk'))
                # exit()
                if shp.endswith('annotation.shp'):

                    # print(shp.decode('gbk'))
                    # exit()
                    daShapefile = shp_dir + shp
                    print(daShapefile.decode('gbk'))
                    # print(daShapefile.decode('gbk'))
                    driver = ogr.GetDriverByName("ESRI Shapefile")
                    dataSource = driver.Open(daShapefile.decode('gbk'), 0)
                    layer = dataSource.GetLayer()
                    # inlist_i = []
                    for feature in layer:
                        geom = feature.GetGeometryRef()
                        val1 = feature.GetField("val1")
                        val2 = feature.GetField("val2")
                        val3 = feature.GetField("val3")
                        val4 = feature.GetField("val4")
                        Points = geom.GetPoints()
                        if Points:
                            for i in range(len(Points)):
                                if i == len(Points) - 1:
                                    break
                                inlist.append([Points[i], Points[i + 1], val1, val2, val3, val4, ''])
                            # print(Points)
                            # print(inlist_i)
                        # exit()
                        # exit()

                        # inlist.append([Points[0],Points[1],val1,val2,val3,val4,''])

            time_end = time.time()
            time_i += 1
            log_process.process_bar(time_i, len(flist), time_init, time_start, time_end)

        print('exporting line shp...')
        line_to_shp(inlist, output_fn)

        pass

    def merge_point_layer_shp(self, shp_type):
        '''

        :return:
        '''
        gdal.SetConfigOption("SHAPE_ENCODING", "GBK")
        fdir = r'E:\cui\190924\dwg_to_shp\\'
        out_dir = r'E:\cui\190924\merge\\'
        flist = os.listdir(fdir)
        time_init = time.time()
        time_i = 0
        inlist = []
        for folder in flist:
            # print(folder.decode('gbk'))
            # exit()
            time_start = time.time()
            # print(folder.decode('gbk'))
            shp_dir = fdir + folder + '\\'

            shp_list = os.listdir(shp_dir)

            for shp in shp_list:
                # print(shp.decode('gbk'))
                # exit()
                if shp.endswith(shp_type + '.shp'):
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
                        # geom.
                        Points = geom.GetPoints()
                        val1 = feature.GetField("val1")
                        val2 = feature.GetField("val2")
                        val3 = feature.GetField("val3")
                        # print(Points)
                        # exit()
                        if Points:
                            # print(Points)
                            # exit()
                            for i in range(len(Points)):
                                # if i == len(Points)-1:
                                #     break
                                inlist.append([Points[i][0], Points[i][1], val1, val2, val3, '', ''])
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
            log_process.process_bar(time_i, len(flist), time_init, time_start, time_end)
        output_fn = out_dir + shp_type + '_merge.shp'.decode('gbk').encode('utf-8')
        print('exporting line shp...')
        # print(inlist)
        # exit()
        point_to_shp(inlist, output_fn)

        pass

    def merge_daoxian(self, indir, outdir):

        fdir = indir + '\\'
        output_fn = outdir + '\\merge_dwg_Polyline.shp'
        inlist = []
        for folder in os.listdir(fdir):
            for f in os.listdir(fdir + folder):
                if f.endswith('_dwg_Polyline.shp'):
                    print(f.decode('gbk'))
                    # print(shp.decode('gbk'))
                    # exit()
                    daShapefile = fdir + folder + '\\' + f
                    print(daShapefile.decode('gbk'))
                    # print(daShapefile.decode('gbk'))
                    driver = ogr.GetDriverByName("ESRI Shapefile")
                    dataSource = driver.Open(daShapefile.decode('gbk'), 0)
                    layer = dataSource.GetLayer()
                    # inlist_i = []
                    for feature in layer:
                        geom = feature.GetGeometryRef()
                        # val1 = feature.GetField("val1")
                        # val2 = feature.GetField("val2")
                        # val3 = feature.GetField("val3")
                        # val4 = feature.GetField("val4")
                        Points = geom.GetPoints()
                        if Points:
                            for i in range(len(Points)):
                                if i == len(Points) - 1:
                                    break
                                inlist.append([Points[i], Points[i + 1], '', '', '', '', ''])

        print('exporting line shp...')
        # print(inlist)
        # exit()
        # print(inlist)
        line_to_shp(inlist, output_fn)


def main():
    Merge().run()
    pass


if __name__ == '__main__':
    main()