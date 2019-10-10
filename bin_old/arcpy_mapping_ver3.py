# coding=gbk
import os
import arcpy
import time
import codecs
import sys
args = sys.argv
# print(args)
# exit()
# root = 'D:\\project13\\output\\��ľ��\\S2017-03-22 1008��ׯ����ѹ��1\\'

# ver = 3

output_mxd3 = {'transparent_jiliangxiang':'transparent',
              '��ѹ��':'bianyaqi','��ѹǽ֧��':'qiangzhijia',
              '����':'dianxiangan',
              '�пձ�λ������':'red',
              '�޿ձ�λ������':'black',
              'dianxiangan_line':'dianxiangan_line',
              '����':'qiangzhijia_line',
              'extent_lyr':'extent_lyr',
              'qiangzhijia_line2':'qiangzhijia_line2',
              '��ѹ���·�֧��':'fenzhixiang',
              '��ѹ����':'dianlan',
              'yonghujierudian_line':'yonghujierudian_line',
              'yonghujierudian_jiliangxiang_line':'yonghujierudian_jiliangxiang_line'
              }

output_mxd1 = {'transparent_jiliangxiang':'transparent_jiliangxiang',
              '��ѹ��':'bianyaqi','��ѹǽ֧��':'qiangzhijia',
              '����':'dianxiangan',
              '�пձ�λ������':'red_jiliangxiang',
              '�޿ձ�λ������':'black_jiliangxiang',
              'dianxiangan_line':'dianxiangan_line',
              '����':'qiangzhijia_line',
              'extent_lyr':'extent_lyr',
              'qiangzhijia_line2':'qiangzhijia_line2'}

output_mxd = output_mxd3

def mapping(dir,outjpgdir):
    # if ver == 1:
    #     dir = 'data\\'+dir
    # elif ver == 3:
    #     dir = dir
    f=open(dir+'\\'+'select.txt','r')
    line = f.readline()
    if line == 'heng':
        mxd_file = args[-2]
    elif line == 'shu':
        mxd_file = args[-1]
    else:
        mxd_file = None
    # mxd_file = 'D:\\project13\\������shu.mxd'
    mxd = arcpy.mapping.MapDocument(mxd_file)
    df0 = arcpy.mapping.ListDataFrames(mxd)[0]

    workplace = 'SHAPEFILE_WORKSPACE'

    for i in output_mxd:
        print '����'.decode('gbk')+i.decode('gbk')
        lyr = arcpy.mapping.ListLayers(mxd,i,df0)[0]
        # print lyr.decode('gbk')
        # print lyr.name
        # print output_mxd[i],'output_mxd[i]'
        try:
            lyr.replaceDataSource(dir,workplace,output_mxd[i])
        except:
            print 'no '+i.decode('gbk')

    extent_lyr = arcpy.mapping.ListLayers(mxd,output_mxd['extent_lyr'],df0)[0]
    df0.extent = extent_lyr.getSelectedExtent()

    try:
        info_text=codecs.open(dir+'\\'+'info.txt','r')
        info = info_text.readline()
        info = info.split(',')
        info1 = []
        for i in info:
            if i == '':
                i = ' '
            info1.append(i)
        info = info1
    except:
        info = [' ']*100
    for textElement in arcpy.mapping.ListLayoutElements(mxd,'TEXT_ELEMENT'):
        try:
            if textElement.name == '̨�����'.decode('gbk'):
                textElement.text=(info[0].decode('gbk'))
            if textElement.name == '̨������'.decode('gbk'):
                textElement.text=(info[1].decode('utf-8'))
            if textElement.name == '��ѹ���ͺ�'.decode('gbk'):
                textElement.text=(info[2].decode('gbk'))
            if textElement.name == '��ѹ������'.decode('gbk'):
                textElement.text=(info[3].decode('gbk'))
            if textElement.name == '��·����'.decode('gbk'):
                textElement.text=(info[4].decode('gbk'))
            if textElement.name == '���ߵ����ͺ�'.decode('gbk'):
                textElement.text=(info[5].decode('gbk'))
            if textElement.name == '֧�ߵ����ͺ�'.decode('gbk'):
                textElement.text=(info[6].decode('gbk'))
            if textElement.name == '�������λ'.decode('gbk'):
                textElement.text=(info[7].decode('gbk'))
            if textElement.name == 'ʵ��װ��λ��'.decode('gbk'):
                textElement.text=(info[8].decode('gbk'))
        except Exception,e:
            print Exception,e
    # dir = dir
    print 'saving mxd to',(dir+'\\mxd.mxd').decode('gbk')
    dir = dir.decode('gbk')
    if os.path.isfile(dir+'\\mxd.mxd'):
        os.remove(dir+'\\mxd.mxd')
    mxd.saveACopy(dir+'\\mxd.mxd','9.2')

    # if ver == 3:
    if info[1] == ' ':
        outjpeg = outjpgdir+dir.split('\\')[-2]
    else:
        outjpeg = outjpgdir+info[1].decode('utf-8')

    arcpy.mapping.ExportToJPEG(mxd,outjpeg,data_frame='PAGE_LAYOUT',df_export_width=mxd.pageSize.width,df_export_height=mxd.pageSize.height,color_mode='8-BIT_GRAYSCALE',resolution=300,jpeg_quality=100)
    print 'done'




if __name__ == '__main__':
    # dir = r'E:\cui\191007\1007��һ����ͼ\1007��һ����ͼ\���--��ľ\��һ��\0418-����ׯ505-��Сǿ ����ΰ\shp\\'
    # outjpg = r'E:\cui\191007\\'
    # print('******')
    fdir = args[1]
    # for i in args:
    #     print(i)
    # print(fdir)
    # exit()
    out_pic_dir = args[2]+'\\'

    for dir in os.listdir(fdir):
        print(fdir+'\\'+dir).decode('gbk')
        try:
            mapping(fdir+'\\'+dir+'\\shp', out_pic_dir)
        except Exception as e:
            print(e)


