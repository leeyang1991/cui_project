# coding=utf-8
import os
import arcpy
import time
import codecs

this_root = os.getcwd()+'\\..\\'

output_mxd = {'zhuanbian':'zhuanbian',
              'naizhang_ganta':'naizhang_ganta',
              'xiangshi_biandianzhan':'xiangshi_biandianzhan',
              'duanluqi':'duanluqi',
              'gongbian':'gongbian',
              'dwg_Polyline':'***********',
              'line_annotation':'line_annotation',
              'zoom_layer':'zoom_layer'
               }

def mk_dir(dir):
    if not os.path.isdir(dir):
        os.makedirs(dir)


def mapping(dir):
    # 横竖选择
    f=open(dir+'\\'+'config.txt','r')
    line = f.readline()
    if line == 'heng':
        # mxd_file = 'D:\\project13\\heng_ver5.mxd'
        mxd_file = this_root+'mxd\\template_heng.mxd'
    elif line == 'shu':
        mxd_file = this_root + 'mxd\\template_shu.mxd'
    else:
        mxd_file = None
    # mxd_file = 'D:\\project13\\新制作shu.mxd'
    mxd = arcpy.mapping.MapDocument(mxd_file)
    df0 = arcpy.mapping.ListDataFrames(mxd)[0]

    workplace = 'SHAPEFILE_WORKSPACE'
    output_mxd['dwg_Polyline'] = (dir.split('\\')[-2]+'_dwg_Polyline').decode('gbk')
    # print (dir.split('\\')[-2]+'_dwg_Polyline').decode('gbk')
    # exit()
    for i in output_mxd:

        try:
            lyr = arcpy.mapping.ListLayers(mxd, i, df0)[0]
            lyr.replaceDataSource(dir,workplace,output_mxd[i])
            print '绘制完成'.decode('gbk') + i.decode('gbk')
        except:
            print 'no '+i.decode('gbk')

    extent_lyr = arcpy.mapping.ListLayers(mxd,output_mxd['zoom_layer'],df0)[0]
    df0.extent = extent_lyr.getSelectedExtent()

    # dir = dir.encode('gbk')
    print 'saving mxd to',(dir+'\\mxd.mxd').decode('gbk')
    dir = dir.decode('gbk')
    if os.path.isfile(dir+'\\mxd.mxd'):
        os.remove(dir+'\\mxd.mxd')
    mxd.saveACopy(dir+'\\mxd.mxd','9.2')
    mk_dir(this_root+'output_pic\\')
    outjpeg = this_root+'output_pic\\'+dir.split('\\')[-2]

    arcpy.mapping.ExportToJPEG(mxd,outjpeg,data_frame='PAGE_LAYOUT',df_export_width=mxd.pageSize.width,df_export_height=mxd.pageSize.height,color_mode='24-BIT_TRUE_COLOR',resolution=300,jpeg_quality=100)
    print 'done'


def main():

    dir = this_root+'190509\\民权线路cad\\dwg_to_shp\\10kV鲁10Ⅱ鲁西线\\'
    mapping(dir)
    pass


if __name__ == '__main__':
    main()
