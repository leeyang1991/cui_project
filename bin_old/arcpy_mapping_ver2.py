# coding=gbk
import os
import arcpy
import time
import codecs

# root = 'D:\\project13\\output\\��ľ��\\S2017-03-22 1008��ׯ����ѹ��1\\'
output_mxd = {'transparent_jiliangxiang':'transparent',
              '��ѹ��':'bianyaqi','��ѹǽ֧��':'qiangzhijia',
              '����':'dianxiangan',
              '�пձ�λ������':'red',
              '�޿ձ�λ������':'black',
              'dianxiangan_line':'dianxiangan_line',
              '����':'qiangzhijia_line',
              'extent_lyr':'extent_lyr',
              'qiangzhijia_line2':'qiangzhijia_line2'}
# arcpy.env.workspace = r'D:\\project13\\output\\��ľ��\\S2017-03-22 1008��ׯ����ѹ��1\\'

if not os.path.isdir('D:\\project13\\output'):
    os.mkdir('D:\\project13\\output')
if not os.path.isdir('D:\\project13\\output_pic'):
    os.mkdir('D:\\project13\\output_pic')
def mapping(dir):

    f=open(r'D:\\project13\\output\\\\'+dir+'\\'+'select.txt','r')
    line = f.readline()
    if line == 'heng':
        mxd_file = 'D:\\project13\\heng.mxd'
    elif line == 'shu':
        mxd_file = 'D:\\project13\\shu.mxd'
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
        lyr.replaceDataSource(r'D:\\project13\\output\\\\'+dir,workplace,output_mxd[i])

    extent_lyr = arcpy.mapping.ListLayers(mxd,output_mxd['extent_lyr'],df0)[0]
    df0.extent = extent_lyr.getSelectedExtent()

    info_text=codecs.open(r'D:\\project13\\output\\\\'+dir+'\\'+'info.txt','r')
    info = info_text.readline()
    info = info.split(',')
    info1 = []
    for i in info:
        if i == '':
            i = ' '
        info1.append(i)
    info = info1
    for textElement in arcpy.mapping.ListLayoutElements(mxd,'TEXT_ELEMENT'):
        try:
            if textElement.name == '̨�����'.decode('gbk'):
                textElement.text=(info[0].decode('gbk'))
            elif textElement.name == '̨������'.decode('gbk'):
                textElement.text=(info[1].decode('utf-8'))
            elif textElement.name == '��ѹ���ͺ�'.decode('gbk'):
                textElement.text=(info[2].decode('gbk'))
            elif textElement.name == '��ѹ������'.decode('gbk'):
                textElement.text=(info[3].decode('gbk'))
            elif textElement.name == '��·����'.decode('gbk'):
                textElement.text=(info[4].decode('gbk'))
            elif textElement.name == '���ߵ����ͺ�'.decode('gbk'):
                textElement.text=(info[5].decode('gbk'))
            elif textElement.name == '֧�ߵ����ͺ�'.decode('gbk'):
                textElement.text=(info[6].decode('gbk'))
            elif textElement.name == '�������λ'.decode('gbk'):
                textElement.text=(info[7].decode('gbk'))
            elif textElement.name == 'ʵ��װ��λ��'.decode('gbk'):
                textElement.text=(info[8].decode('gbk'))
            elif textElement.name == '������'.decode('gbk'):
                textElement.text=(info[9].decode('utf-8'))
        except Exception,e:
            print Exception,e
    dir = dir.encode('gbk')
    print 'saving mxd to',('D:\\project13\\output\\\\'+dir+'\\mxd.mxd').decode('gbk')
    dir = dir.decode('gbk')
    if os.path.isfile('D:\\project13\\output\\\\'+dir+'\\mxd.mxd'):
        os.remove('D:\\project13\\output\\\\'+dir+'\\mxd.mxd')
    mxd.saveACopy('D:\\project13\\output\\\\'+dir+'\\mxd.mxd','9.2')
    arcpy.mapping.ExportToJPEG(mxd,'D:\\project13\\output_pic\\'+dir.replace('.','_'),data_frame='PAGE_LAYOUT',df_export_width=mxd.pageSize.width,df_export_height=mxd.pageSize.height,color_mode='8-BIT_GRAYSCALE',resolution=300,jpeg_quality=100)
    print 'done'
# mapping()

folder_list = os.listdir('D:\\project13\\output\\\\')

# if os.path.isfile('../mapping_log.txt'):
#     f = open('../log.txt','r')
#     lines = f.readlines()
#     f.close()
#
# else:
#     lines = []
# #
# f = open('../mapping_log.txt','w')
# f.write(''.join(lines))

j = 0
total = len(folder_list)
for folder in folder_list:
    j+=1
    print j,'/',total
    print folder.decode('gbk')

    try:
        mapping(folder.decode('gbk'))
    except Exception,e:
        print Exception,e
        print 'error'
        print 'error'
        print 'error'
        print 'error'
        print 'error'
        print folder.decode('gbk'),'����ʧ��'.decode('gbk')
        continue

    # #debug
    # mapping(folder)
    print (folder+' �������').decode('gbk')
    print '-------------------------------------'
    print '-------------------------------------'
    # break
    # except Exception,e:
    # f.write('error\t'+time.asctime(time.localtime(time.time()))+'\t'+folder+'\t'+str(e)+'\n')
    # print folder.decode('gbk')
    # print Exception,e


