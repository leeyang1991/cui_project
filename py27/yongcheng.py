# coding=gbk

from osgeo import ogr
import os
# this_root = os.getcwd()+'\\..\\'
this_root = 'E:\\cui\\'
import gdal
import xlrd
import time
import re
import numpy as np
import log_process
from matplotlib import pyplot as plt

#
# f_excel = this_root+'190905/台账数据.xlsx'
# f_excel = 'E:\\cui\\190905\\台账数据.xls'.decode('gbk')


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




class GenLayer:

    def __init__(self):
        fdir = 'E:\\cui\\190905\\'
        flist = os.listdir(fdir)
        for f in flist:
            if f.endswith('xls'):
                self.f_excel = fdir + f
        pass

    def gen_naizhang_ganta_excel(self):

        # f_excel = this_root+'190509\\民权线路cad\\民权台账 - 以此为准.xls'
        bk = xlrd.open_workbook(self.f_excel)
        sh = bk.sheet_by_name('杆塔台账'.decode('gbk'))
        nrows = sh.nrows
        ganta = []
        ganta_num = []
        for i in range(nrows):
            ganta_attrib = sh.cell_value(i,4)
            ganta_name = sh.cell(i,0)
            ganta_num_i = sh.cell_value(i,1)
            if ganta_attrib == '耐张'.decode('gbk'):
                ganta.append(ganta_name.value)
                ganta_num.append(ganta_num_i)
        ganta_dic = {}
        for i in range(len(ganta)):
            ganta_dic[ganta[i]] = ganta_num[i]
        return ganta_dic


    def gen_naizhang_ganta_shp(self,daShapefile,out_shp):

        # daShapefile = this_root+'123_dwg_Annotation.shp'
        driver = ogr.GetDriverByName("ESRI Shapefile")
        dataSource = driver.Open(daShapefile, 0)
        layer = dataSource.GetLayer()
        ganta = self.gen_naizhang_ganta_excel()
        out_list = []
        for feature in layer:
            geom = feature.GetGeometryRef()
            x = geom.GetX()
            y = geom.GetY()
            name = feature.GetField("RefName")
            name_gbk = name.decode('utf-8')
            if name_gbk in ganta:
                out_list.append([x,y,name_gbk,ganta[name_gbk],''])
        point_to_shp(out_list,out_shp)



    def gen_zhushang_bianyaqi_excel(self):
        # f_excel = this_root + '190509\\民权线路cad\\民权台账 - 以此为准.xls'
        bk = xlrd.open_workbook(self.f_excel)
        sh = bk.sheet_by_name('柱上变压器'.decode('gbk'))
        nrows = sh.nrows
        gongbian = {}
        zhuanbian = {}
        for i in range(nrows):
            bianyaqi_name = sh.cell_value(i,0)
            bianyaqi_xinghao = sh.cell_value(i,3)

            bianyaqi_attrib = sh.cell_value(i,4)
            if bianyaqi_attrib == '公变'.decode('gbk'):
                gongbian[bianyaqi_name] = bianyaqi_xinghao
            elif bianyaqi_attrib == '专变'.decode('gbk'):
                zhuanbian[bianyaqi_name] = bianyaqi_xinghao
        return gongbian,zhuanbian
        pass


    def gen_zhushang_bianyaqi_shp(self):
        gongbian, zhuanbian = self.gen_zhushang_bianyaqi_excel()
        daShapefile = this_root + '123_dwg_Annotation.shp'
        driver = ogr.GetDriverByName("ESRI Shapefile")
        dataSource = driver.Open(daShapefile, 0)
        layer = dataSource.GetLayer()
        # ganta = gen_naizhang_ganta_excel()
        out_list_gongbian = []
        out_list_zhuanbian = []
        for feature in layer:
            geom = feature.GetGeometryRef()
            x = geom.GetX()
            y = geom.GetY()
            name = feature.GetField("RefName")
            name_gbk = name.decode('utf-8')
            if name_gbk in gongbian:
                out_list_gongbian.append([x, y, name_gbk, gongbian[name_gbk],''])
            elif name_gbk in zhuanbian:
                out_list_zhuanbian.append([x, y, name_gbk, zhuanbian[name_gbk],''])
        point_to_shp(out_list_gongbian, 'gongbian.shp')
        point_to_shp(out_list_zhuanbian, 'zhuanbian.shp')
        pass


    def gen_xiangshi_biandianzhan_excel(self):
        # f_excel = this_root + '190509\\民权线路cad\\民权台账 - 以此为准.xls'
        bk = xlrd.open_workbook(self.f_excel)
        sh = bk.sheet_by_name('箱式变电站'.decode('gbk'))
        nrows = sh.nrows
        biandianzhan = {}
        for i in range(nrows):
            biandianzhan_name = sh.cell_value(i, 0)

            if len(sh.cell_value(i,3)) > 1:
                biandianzhan_xinghao = sh.cell_value(i, 2)+' '+sh.cell_value(i,3)
            else:
                biandianzhan_xinghao = sh.cell_value(i, 2)

            bianyaqi_biaozhu = sh.cell_value(i, 1)
            biandianzhan[biandianzhan_name] = [bianyaqi_biaozhu,biandianzhan_xinghao]
        return biandianzhan
        pass


    def gen_xiangshi_biandianzhan_shp(self,daShapefile,out_shp):
        biandianzhan = self.gen_xiangshi_biandianzhan_excel()
        # daShapefile = this_root + '123_dwg_Annotation.shp'
        driver = ogr.GetDriverByName("ESRI Shapefile")
        dataSource = driver.Open(daShapefile, 0)
        layer = dataSource.GetLayer()
        out_list_biandianzhan = []
        for feature in layer:
            geom = feature.GetGeometryRef()
            x = geom.GetX()
            y = geom.GetY()
            name = feature.GetField("RefName")
            name_gbk = name.decode('utf-8')
            if name_gbk in biandianzhan:
                out_list_biandianzhan.append([x, y, name_gbk, biandianzhan[name_gbk][0], biandianzhan[name_gbk][1]])

        point_to_shp(out_list_biandianzhan, out_shp)
        # point_to_shp(out_list_zhuanbian, 'zhuanbian.shp')
        pass


    def gen_duanluqi_excel(self):
        # f_excel = this_root + '190509\\民权线路cad\\民权台账 - 以此为准.xls'
        bk = xlrd.open_workbook(self.f_excel)
        sh = bk.sheet_by_name('断路器'.decode('gbk'))
        nrows = sh.nrows
        rongduanqi = {}
        for i in range(nrows):
            rongduanqi_name = sh.cell_value(i, 0)
            rongduanqi_attrib = sh.cell_value(i, 4)
            rongduanqi[rongduanqi_name] = rongduanqi_attrib
        return rongduanqi
        pass


    def gen_duanluqi_shp(self,daShapefile,out_shp):
        rongduanqi = self.gen_duanluqi_excel()
        # daShapefile = this_root + '123_dwg_Annotation.shp'
        driver = ogr.GetDriverByName("ESRI Shapefile")
        dataSource = driver.Open(daShapefile, 0)
        layer = dataSource.GetLayer()
        out_list_biandianzhan = []
        for feature in layer:
            geom = feature.GetGeometryRef()
            x = geom.GetX()
            y = geom.GetY()
            name = feature.GetField("RefName")
            name_gbk = name.decode('utf-8')


            if name_gbk in rongduanqi:
                out_list_biandianzhan.append([x, y, name_gbk, rongduanqi[name_gbk],''])

            # if '熔断器'.decode('gbk') in name_gbk:
            #     out_list_biandianzhan.append([x, y, name_gbk, '',''])

        point_to_shp(out_list_biandianzhan, out_shp)

        pass



    def gen_gongbian_excel(self):
        # f_excel = this_root + '190509\\民权线路cad\\民权台账 - 以此为准.xls'
        bk = xlrd.open_workbook(self.f_excel)
        sh = bk.sheet_by_name('柱上变压器'.decode('gbk'))
        nrows = sh.nrows
        gongbian = {}
        for i in range(nrows):
            bianyaqi_name = sh.cell_value(i, 0)
            bianyaqi_xinghao = sh.cell_value(i, 4)
            gongbian[bianyaqi_name] = bianyaqi_xinghao

        # for i in gongbian:
        #     print i,gongbian[i]

        return gongbian
        pass

        pass



    def gen_gongbian_shp(self,daShapefile,out_shp):
        gongbian = self.gen_gongbian_excel()
        # daShapefile = this_root + '123_dwg_Annotation.shp'
        driver = ogr.GetDriverByName("ESRI Shapefile")
        dataSource = driver.Open(daShapefile, 0)
        layer = dataSource.GetLayer()
        # ganta = gen_naizhang_ganta_excel()
        out_list_gongbian = []

        for feature in layer:
            geom = feature.GetGeometryRef()
            x = geom.GetX()
            y = geom.GetY()
            name = feature.GetField("RefName")
            name_gbk = name.decode('utf-8')
            if name_gbk in gongbian:
                out_list_gongbian.append([x, y, name_gbk, gongbian[name_gbk],''])

        point_to_shp(out_list_gongbian, out_shp)
        pass




    def gen_zhuanbian_excel(self):
        # f_excel = this_root + '190509\\民权线路cad\\民权台账 - 以此为准.xls'
        bk = xlrd.open_workbook(self.f_excel)
        sh = bk.sheet_by_name('专变'.decode('gbk'))
        nrows = sh.nrows
        zhuanbian = {}
        for i in range(nrows):
            if i+1 == nrows:
                continue
            bianyaqi_name = sh.cell_value(i+1, 1)
            bianyaqi_xinghao = sh.cell_value(i+1, 3)
            bianyaqi_rongliang = sh.cell_value(i+1, 2)
            # bianyaqi_rongliang = int(float(unicode(bianyaqi_rongliang)))
            bianyaqi_rongliang = unicode(bianyaqi_rongliang)
            # print(bianyaqi_rongliang)
            bianyaqi_rongliang = int(float(bianyaqi_rongliang))
            bianyaqi_rongliang = str(bianyaqi_rongliang)
            zhuanbian[bianyaqi_name] = bianyaqi_xinghao+' '+bianyaqi_rongliang

        # for i in zhuanbian:
        #     print i,'\n',zhuanbian[i]

        return zhuanbian
        pass

        pass



    def gen_zhuanbian_shp(self,daShapefile,out_shp):
        zhuanbian = self.gen_zhuanbian_excel()
        # daShapefile = this_root + '123_dwg_Annotation.shp'
        driver = ogr.GetDriverByName("ESRI Shapefile")
        dataSource = driver.Open(daShapefile, 0)
        layer = dataSource.GetLayer()
        # ganta = gen_naizhang_ganta_excel()
        out_list_gongbian = []

        for feature in layer:
            geom = feature.GetGeometryRef()
            x = geom.GetX()
            y = geom.GetY()
            name = feature.GetField("RefName")
            name_gbk = name.decode('utf-8')
            if name_gbk in zhuanbian:
                out_list_gongbian.append([x, y, name_gbk, zhuanbian[name_gbk],''])

        point_to_shp(out_list_gongbian, out_shp)
        pass


    def gen_biandianzhan_excel(self):
        # f_excel = this_root + u'190714\\张桥所线路设备明细.xls'
        bk = xlrd.open_workbook(self.f_excel)
        sh = bk.sheet_by_name(u'电站')
        nrows = sh.nrows
        biandiamzhan = {}
        for i in range(nrows):
            if i + 1 == nrows:
                continue
            biandiamzhan_name = sh.cell_value(i + 1, 2)
            val1 = biandiamzhan_name
            biandiamzhan[val1] = val1

        return biandiamzhan

    def gen_biandianzhan_shp(self,daShapefile,out_shp):
        biandianzhan = self.gen_biandianzhan_excel()
        str_num = []
        for key in biandianzhan:
            # print key,len(key)
            str_num.append(len(key))
        # print(min(str_num))
        # print(max(str_num))
        # exit()
        driver = ogr.GetDriverByName("ESRI Shapefile")
        dataSource = driver.Open(daShapefile, 0)
        layer = dataSource.GetLayer()
        # ganta = gen_naizhang_ganta_excel()
        out_list_biandianzhan = []

        name_list = []
        xy_list = []
        for feature in layer:
            geom = feature.GetGeometryRef()
            x = geom.GetX()
            y = geom.GetY()
            xy_list.append([x,y])
            name = feature.GetField("RefName")
            name_gbk = name.decode('utf-8')

            name_list.append(name_gbk)
        # print(xy_list)
        out_dic = {}
        for i in range(len(name_list)):
            sum_str = ''
            selected_num = []
            for j in range(max(str_num)):
                try:
                    sum_str += name_list[i+j]
                    selected_num.append(i+j)
                    # print(sum_str)
                    if sum_str in biandianzhan:
                        # print(sum_str)
                        selected_x = []
                        selected_y = []
                        for k in selected_num:
                            selected_x.append(xy_list[k][0])
                            selected_y.append(xy_list[k][1])
                        x = np.mean(selected_x)
                        y = np.mean(selected_y)
                        out_dic[sum_str] = []
                        out_list_biandianzhan.append([x,y,sum_str,'',''])
                    str_num_ = len(sum_str)
                    if str_num_ > max(str_num):
                        break
                except:
                    pass
        for i in out_list_biandianzhan:
            name = i[2]
            out_dic[name].append([i[0],i[1]])

        out_list_biandianzhan = []
        for name in out_dic:
            # print(name)
            x = []
            y = []
            for xy in out_dic[name]:
                x.append(xy[0])
                y.append(xy[1])
            x_mean = np.mean(x)
            y_mean = np.mean(y)
            out_list_biandianzhan.append([x_mean,y_mean,name,'',''])
            # print(x_mean)
            # print(y_mean)
            # for i in zhuanbian:
            #     print(i)
            # exit()
            # continue
            # if name_gbk in biandianzhan:
            #     out_list_biandianzhan.append([x, y, name_gbk, biandianzhan[name_gbk],''])
        #
        point_to_shp(out_list_biandianzhan, out_shp)
        # pass

    def gen_zoom_layer(self,daShapefile, out_shp):
        '''
        生成zoom layer shp
        生成横竖config
        :param daShapefile:
        :param out_shp:
        :return:
        '''
        driver = ogr.GetDriverByName("ESRI Shapefile")
        dataSource = driver.Open(daShapefile, 0)
        layer = dataSource.GetLayer()
        x_list = []
        y_list = []

        for feature in layer:
            geom = feature.GetGeometryRef()
            x = geom.GetX()
            y = geom.GetY()
            x_list.append(x)
            y_list.append(y)
        xmin = min(x_list) - 0.003
        xmax = max(x_list) + 0.003
        ymin = min(y_list) - 0.003
        ymax = max(y_list) + 0.003

        a = [xmin, ymin, '', '', '']
        b = [xmin, ymax, '', '', '']
        c = [xmax, ymin, '', '', '']
        d = [xmax, ymax, '', '', '']
        # plt.scatter(a[0],a[1])
        # plt.scatter(b[0],b[1])
        # plt.scatter(c[0],c[1])
        # plt.scatter(d[0],d[1])
        # plt.show()
        inlist = [a, b, c, d]
        point_to_shp(inlist, out_shp)

        # 横竖config
        x_range = xmax - xmin
        y_range = ymax - ymin

        # print(x_range)
        # print(y_range)
        file_name = '\\'.join(daShapefile.split('\\')[:-1]) + '\\config.txt'
        # print(file_name.decode('utf-8'))
        fw = open(file_name.decode('utf-8'), 'w')
        if x_range > y_range:
            fw.write('heng')
        elif y_range > x_range:
            fw.write('shu')
        else:
            fw.write('error')

        pass



class LineAnnotation:

    def __init__(self):
        # 1 gen_shp_dic
        # 2 line_annotation
        # 3 gen_line_annotation
        pass

    def gen_shp_dic(self):

        fdir = 'E:\\cui\\190905\\dwg_to_shp\\'
        flist = os.listdir(fdir)
        shp_dic = {}

        time_init = time.time()
        flag = 0
        for folder in flist:
            time_start = time.time()
            folder = folder.decode('gbk')
            print(folder)
            shp_list = os.listdir(fdir+folder)
            for shp in shp_list:
                if shp.endswith('Annotation.shp'):
                    daShapefile = fdir+folder+'\\'+shp
                    print(daShapefile)
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
        # np.save(this_root+'190509\\民权线路cad\\shp_dic',shp_dic)
        np.save(r'E:\cui\190905\shp_dic',shp_dic)


    def line_annotation(self):
        '''
        生成离散的line shp
        1、生成shp dic
        根据所有的dwg to shp，生成电线杆的坐标，存入字典
        2、打开excel，re筛选汉字，将汉字作为字典键值，生成电线杆字典，将设备名称，所属线路等属性加入字典
        3、xian_dic 1-100排序 生成 xian_dic_sort
        4、根据xian_dic_sort生成shp
        5、下一步合成shp（功能不在此函数）
        :return:
        '''
        # npy = this_root + '190509\\民权线路cad\\shp_dic.npy'
        this_dir = r'E:\cui\190905\dwg_to_shp\\'
        npy = r'E:\cui\190905\shp_dic.npy'
        shp_dic = np.load(npy).item()
        shp_dic = dict(shp_dic)


        # f_excel = this_root + '190509\\民权线路cad\\民权台账 - 以此为准.xls'
        bk = xlrd.open_workbook(f_excel)
        sh = bk.sheet_by_index(0)
        nrows = sh.nrows
        #  筛选汉字
        xx = u"([\u4e00-\u9fff]+)"
        pattern = re.compile(xx)
        # 建立字典
        xian_list = []
        for i in range(nrows):
            shebei_name = sh.cell_value(i, 0)
            results = pattern.findall(shebei_name)
            xian_list.append(''.join(results))
        xian_list = set(xian_list)
        xian_dic = {}
        xian_dic_sort = {}
        for x in xian_list:
            xian_dic[x] = []
            xian_dic_sort[x] = []

        # 读excel 信息
        for i in range(nrows):
            shebei_name = sh.cell_value(i, 0)
            suoshu_xianlu = sh.cell_value(i,2)
            zhixian_mingcheng = sh.cell_value(i, 5)
            daoxian_xinghao = sh.cell_value(i, 6)
            results = pattern.findall(shebei_name)
            hanzi_xian = ''.join(results)
            xian_dic[hanzi_xian].append([shebei_name,suoshu_xianlu,zhixian_mingcheng,daoxian_xinghao])

        # sort 1-100
        error = 0
        for key in xian_dic:
            # line_list = []
            for i in range(len(xian_dic[key])):
                for j in range(len(xian_dic[key])):
                    # print(u'线#'+'%03d'%(i+1)+u'号')
                    if u'线#'+'%03d'%(i+1)+u'号' in xian_dic[key][j][0] or \
                        u'支#' + '%03d'%(i+1) + u'杆' in xian_dic[key][j][0] or \
                            u'线#' + '%03d'%(i+1) in xian_dic[key][j][0]:

                        try:
                            a = shp_dic[xian_dic[key][j][0]]
                            xian_dic_sort[key].append(xian_dic[key][j])
                        except:
                            error+=1
                            pass

        flag = 0
        time_init = time.time()

        for key in xian_dic_sort:
            time_start = time.time()
            inlist = []
            for i in range(len(xian_dic_sort[key])):
                if i == len(xian_dic_sort[key])-1:
                    continue
                x1 = shp_dic[xian_dic_sort[key][i][0]][0]
                y1 = shp_dic[xian_dic_sort[key][i][0]][1]

                x2 = shp_dic[xian_dic_sort[key][i+1][0]][0]
                y2 = shp_dic[xian_dic_sort[key][i+1][0]][1]

                start = [x1,y1]
                end = [x2,y2]
                # name_gbk = name.decode('utf-8')
                val1 = xian_dic_sort[key][i][0].encode('utf-8')+'-'+xian_dic_sort[key][i+1][0].encode('utf-8')
                val2 = xian_dic_sort[key][i][1].encode('utf-8')
                val3 = xian_dic_sort[key][i][2].encode('utf-8')
                val4 = xian_dic_sort[key][i][3].encode('utf-8')
                # print(val4)
                inlist.append([start,end,val1,val2,val3,val4,''])


            # print(key.encode('utf-8'))
            # for i in inlist:
            #     print(i)
            try:
                dwg_dir = this_dir+val2.decode('utf-8')+'\\annotation\\'
                mk_dir(dwg_dir)
                dwg_dir = dwg_dir.encode('utf-8')
                try:
                    line_to_shp(inlist, dwg_dir + '\\' + key.encode('utf-8') + 'line_annotation.shp')
                except:
                    pass
            except Exception as e:
                print(e)
            # print(dwg_dir)
            # exit()

            time_end = time.time()
            log_process.process_bar(flag,len(xian_dic_sort),time_init,time_start,time_end)
            flag += 1

        pass

    def gen_line_annotation(self):
        '''
        composite shp
        将xian_dic_sort生成的shp合成为1个shp
        作为annotation
        :return:
        '''
        gdal.SetConfigOption("SHAPE_ENCODING", "GBK")
        fdir = r'E:\cui\190905\dwg_to_shp\\'
        flist = os.listdir(fdir)
        time_init = time.time()
        time_i = 0
        for folder in flist:
            time_start = time.time()
            print(folder.decode('gbk'))
            shp_dir = fdir + folder + '\\annotation\\'
            try:
                shp_list = os.listdir(shp_dir)
                inlist = []
                for shp in shp_list:
                    # print(shp)
                    if shp.endswith('.shp'):
                        # print(shp.decode('gbk'))
                        daShapefile = shp_dir + shp
                        # print(daShapefile.decode('gbk'))
                        driver = ogr.GetDriverByName("ESRI Shapefile")
                        dataSource = driver.Open(daShapefile.decode('gbk'), 0)
                        layer = dataSource.GetLayer()
                        for feature in layer:
                            geom = feature.GetGeometryRef()
                            Points = geom.GetPoints()
                            val1 = feature.GetField("val1")
                            val2 = feature.GetField("val2")
                            val3 = feature.GetField("val3")
                            val4 = feature.GetField("val4")
                            inlist.append([Points[0], Points[1], val1, val2, val3, val4, ''])
                # print(fdir.decode('gbk') + folder.decode('gbk') + '\\test.shp')
                output_fn = fdir.decode('gbk').encode('utf-8') + folder.decode('gbk').encode(
                    'utf-8') + '\\line_annotation.shp'
                # print output_fn
                line_to_shp(inlist, output_fn)
                time_end = time.time()
                time_i += 1
                log_process.process_bar(time_i, len(flist), time_init, time_start, time_end)
            except:
                time_i += 1
                pass

            # exit()
        pass

        # exit()


def foo():
    '''
    看shp_dic字典里是啥
    :return:
    '''
    npy = this_root+'190509\\民权线路cad\\shp_dic.npy'
    shp_dic = np.load(npy).item()
    shp_dic = dict(shp_dic)
    # print(len(shp_dic))
    for i in shp_dic:
        print i,shp_dic[i]
        time.sleep(0.5)




def delete_shp():
    fdir = this_root+'\\190509\\民权线路cad\\dwg_to_shp\\'
    flist = os.listdir(fdir)
    for folder in flist:
        try:
            for f in os.listdir(fdir+folder+'\\annotation\\'):
                print(fdir.decode('gbk')+folder.decode('gbk')+'\\annotation\\'+f.decode('gbk'))
            # print(folder.decode('gbk'))
                # if 'line_annotation' in f:
                #     print(f.decode('gbk'))
                os.remove(fdir+folder+'\\annotation\\'+f)
        except:
            pass
        # exit()




def main():
    fdir = this_root+'190905\\dwg_to_shp\\'
    flist = os.listdir(fdir)
    genlayer = GenLayer()
    for folder in flist:
        shp_dir = fdir+folder+'\\'
        shp_list = os.listdir(shp_dir)
        for shp in shp_list:
            if shp.endswith('Annotation.shp'):
                fname = (shp_dir+shp).decode('gbk')
                print(fname)
                # print((shp_dir+'naizhang_ganta.shp').decode('gbk'))
                genlayer.gen_naizhang_ganta_shp(fname.encode('utf-8'),(shp_dir+'naizhang_ganta.shp').decode('gbk').encode('utf-8'))
                # gen_xiangshi_biandianzhan_shp(fname.encode('utf-8'),(shp_dir+'xiangshi_biandianzhan.shp').decode('gbk').encode('utf-8'))
                genlayer.gen_duanluqi_shp(fname.encode('utf-8'),(shp_dir+'duanluqi.shp').decode('gbk').encode('utf-8'))
                genlayer.gen_gongbian_shp(fname.encode('utf-8'),(shp_dir+'gongbian.shp').decode('gbk').encode('utf-8'))
                genlayer.gen_zhuanbian_shp(fname.encode('utf-8'),(shp_dir+'zhuanbian.shp').decode('gbk').encode('utf-8'))
                genlayer.gen_zoom_layer(fname.encode('utf-8'),(shp_dir+'zoom_layer.shp').decode('gbk').encode('utf-8'))


        # exit()


if __name__ == '__main__':

    #1 gendic
    # gen_shp_dic()
    #2 gen line annotation
    # line_annotation()
    #3 gen_line_annotation
    gen_line_annotation()
    # main()