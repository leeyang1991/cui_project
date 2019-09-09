# coding=gbk
import os
import arcpy
import time
import codecs

this_root = 'E:\\cui\\'

output_mxd = {'�����û���ѹ��':'zhuanbian',
              '���Ÿ���':'naizhang_ganta',
              'xiangshi_biandianzhan':'xiangshi_biandianzhan',
              '��·��':'duanluqi',
              '���ϱ�ѹ��':'gongbian',
              # 'dwg_Polyline':'***********',
              'line_annotation':'line_annotation',
              'zoom_layer':'zoom_layer',
              '��վ':'biandianzhan'
               }

def mk_dir(dir):
    if not os.path.isdir(dir):
        os.makedirs(dir)


def mapping(dir):
    # ����ѡ��
    f=open(dir+'\\'+'config.txt','r')
    line = f.readline()
    if line == 'heng':
        # mxd_file = 'D:\\project13\\heng_ver5.mxd'
        # mxd_file = this_root+'mxd\\template_heng.mxd'
        mxd_file = r'E:\cui\190905\template_heng.mxd'
    elif line == 'shu':
        # mxd_file = this_root + 'mxd\\template_shu.mxd'
        mxd_file = r'E:\cui\190905\template_shu.mxd'
    else:
        mxd_file = None
    # mxd_file = 'D:\\project13\\������shu.mxd'
    title = ''
    # print(dir)
    # exit()

    mxd = arcpy.mapping.MapDocument(mxd_file)
    df0 = arcpy.mapping.ListDataFrames(mxd)[0]

    workplace = 'SHAPEFILE_WORKSPACE'
    output_mxd['����'] = (dir.split('\\')[-2]+'_dwg_Polyline').decode('gbk')
    title = output_mxd['����'].split('_')[0]+' ��·�ز�ͼ'.decode('gbk')

    for textElement in arcpy.mapping.ListLayoutElements(mxd, "TEXT_ELEMENT"):
        if textElement.name == 'title':
            textElement.text = (title)



    for i in output_mxd:

        try:
            lyr = arcpy.mapping.ListLayers(mxd, i, df0)[0]
            lyr.replaceDataSource(dir,workplace,output_mxd[i])
            print '�������'.decode('gbk') + i.decode('gbk')
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

    # dir = this_root+'190509\\��Ȩ��·cad\\dwg_to_shp\\10kV³10��³����\\'
    # dir = 'E:\\cui\\190905\\dwg_to_shp\\��1������\\'
    # mapping(dir)
    dirs = 'E:\\cui\\190905\\dwg_to_shp\\'
    dir_list = os.listdir(dirs)
    for dir in dir_list:
        # dir = dirs+'\\'+dir.decode('gbk')+'\\'
        dir = dirs+'\\'+dir+'\\'
        print(dir)
        mapping(dir)
    pass


if __name__ == '__main__':
    main()
