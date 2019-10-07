# coding=gbk
import os
import arcpy
import time
import codecs
import sys
argvs = sys.argv
# print(argvs)
# this_root = 'E:\\cui\\'

output_mxd = {'柱上用户变压器':'zhuanbian',
              '耐张杆塔':'naizhang_ganta',
              'xiangshi_biandianzhan':'xiangshi_biandianzhan',
              '断路器':'duanluqi',
              '柱上变压器':'zhushangbianyaqi',
              # 'dwg_Polyline':'***********',
              'line_annotation':'line_annotation1',
              'zoom_layer':'zoom_layer',
              '电站':'biandianzhan'
               }

def mk_dir(dir):
    if not os.path.isdir(dir):
        os.makedirs(dir)


def mapping(dir,out_pic_dir):
    # 横竖选择
    f=open(dir+'\\'+'config.txt','r')
    line = f.readline()
    if line == 'heng':
        # mxd_file = 'D:\\project13\\heng_ver5.mxd'
        # mxd_file = this_root+'mxd\\template_heng.mxd'
        # mxd_file = r'E:\cui\190905\template_heng.mxd'
        mxd_file = sys.argv[3]
    elif line == 'shu':
        # mxd_file = this_root + 'mxd\\template_shu.mxd'
        # mxd_file = r'E:\cui\190905\template_shu.mxd'
        mxd_file = sys.argv[4]
    else:
        mxd_file = None
    # mxd_file = 'D:\\project13\\新制作shu.mxd'
    title = ''
    # print(dir)
    # exit()

    mxd = arcpy.mapping.MapDocument(mxd_file)
    df0 = arcpy.mapping.ListDataFrames(mxd)[0]

    workplace = 'SHAPEFILE_WORKSPACE'
    output_mxd['导线'] = (dir.split('\\')[-2]+'_dwg_Polyline').decode('gbk')
    title = output_mxd['导线'].split('_')[0]+' 线路沿布图'.decode('gbk')

    for textElement in arcpy.mapping.ListLayoutElements(mxd, "TEXT_ELEMENT"):
        if textElement.name == 'title':
            textElement.text = (title)



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
    mk_dir(out_pic_dir.decode('gbk')+'\\')
    # outjpeg = this_root+'output_pic\\'+dir.split('\\')[-2]
    outjpeg = out_pic_dir.decode('gbk')+'\\'+dir.split('\\')[-2]
    # print(outjpeg)
    arcpy.mapping.ExportToJPEG(mxd,outjpeg,data_frame='PAGE_LAYOUT',df_export_width=mxd.pageSize.width,df_export_height=mxd.pageSize.height,color_mode='24-BIT_TRUE_COLOR',resolution=300,jpeg_quality=100)
    # arcpy.mapping.ExportToAI(mxd,outjpeg+'.ai')
    print 'done'


def main():

    # dir = this_root+'190509\\民权线路cad\\dwg_to_shp\\10kV鲁10Ⅱ鲁西线\\'
    # dir = 'E:\\cui\\190905\\dwg_to_shp\\蒋1蒋蒋线\\'
    # mapping(dir)
    # dirs = 'E:\\cui\\190905\\dwg_to_shp\\'
    # for i in sys.argv:
    #     print(i)
    # exit()
    dirs = sys.argv[1]
    out_pic_dir = sys.argv[2]
    # print(dirs)
    # print(out_pic_dir)
    # print(dirs)
    dir_list = os.listdir(dirs)
    for dir in dir_list:
        # dir = dirs+'\\'+dir.decode('gbk')+'\\'
        dir = dirs+'\\'+dir+'\\'
        print(dir.decode('gbk'))
        mapping(dir,out_pic_dir)
    # print(sys.argv)
    # pass


if __name__ == '__main__':
    main()
