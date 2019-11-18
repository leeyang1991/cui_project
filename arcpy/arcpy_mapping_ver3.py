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
              'xiangshi_biandianzhan':'xiangbian',
              '断路器':'duanluqi',
              '柱上变压器':'zhushangbianyaqi',
              # 'dwg_Polyline':'***********',
              'line_annotation':'line_annotation1',
              'zoom_layer':'zoom_layer',
              '电站':'biandianzhan',
              '导线':'daoxian',
              '电缆':'dianlan',
              'huanwang':'huanwang',
              'peidian':'peidian',
               }

def mk_dir(dir):
    if not os.path.isdir(dir):
        os.makedirs(dir)

def huanhang(txt,n):
    n=int(n)
    newtxt=''
    for i,v in enumerate(txt):
        # if i == 0:
        #     continue
        if i % n == 0 and i != 0:
            newtxt+='\n'
        newtxt+=v

    return newtxt

def new_huanhang(txt,n):
    new_text = ''
    # print([txt])
    if '\\n' in txt:
        txt = txt.split('\\n')
        # print(123123)
        for t in txt:
            if len(t) > n:
                nt = huanhang(t,n)
                new_text += nt + '\\n'
            else:
                new_text+=t+'\\n'
    else:
        new_text = huanhang(txt,n)
    return new_text


def mapping(dir,out_pic_dir,ditu_path):
    # 横竖选择
    f=open(dir+'\\'+'config.txt','r')
    line = f.readline()
    template = line.split(',')[0]
    level = line.split(',')[1]
    if template == 'heng':
        # mxd_file = 'D:\\project13\\heng_ver5.mxd'
        # mxd_file = this_root+'mxd\\template_heng.mxd'
        # mxd_file = r'E:\cui\190905\template_heng.mxd'
        mxd_file = sys.argv[4]
        x = 35.0339
    elif template == 'shu':
        # mxd_file = this_root + 'mxd\\template_shu.mxd'
        # mxd_file = r'E:\cui\190905\template_shu.mxd'
        mxd_file = sys.argv[5]
        x = 22.9125
    else:
        mxd_file = None
        x = None

    mxd = arcpy.mapping.MapDocument(mxd_file)
    df0 = arcpy.mapping.ListDataFrames(mxd)[0]

    workplace = 'SHAPEFILE_WORKSPACE'
    # title = dir.split('\\')[-2]+'线路沿布图'
    text_f = open(dir+'\\'+'info.txt','r')
    line = text_f.readline()
    shebeimingcheng, qidiandianzhan, weihubanzu,\
    xianluzongchangdu,jiakong, dianlan, gongbian, \
    zhuanbian, duanluqi,title,beizhu = line.split(',')
    if len(beizhu) == 0:
        beizhu = ' '
    # print(beizhu)
    beizhu = new_huanhang(beizhu.decode('utf-8'),30)
    print(beizhu)
    # exit()
    for textElement in arcpy.mapping.ListLayoutElements(mxd, "TEXT_ELEMENT"):
        if textElement.name == 'title':
            textElement.text = (title)
        elif textElement.name == 'shebeimingcheng':
            textElement.text = (shebeimingcheng)
        elif textElement.name == 'qidiandianzhan':
            textElement.text = (qidiandianzhan)
        elif textElement.name == 'weihubanzu':
            textElement.text = (weihubanzu)
        elif textElement.name == 'xianluzongchangdu':
            textElement.text = (xianluzongchangdu)
        elif textElement.name == 'jiakongxianluchangdu':
            textElement.text = (jiakong)
        elif textElement.name == 'dianlanxianluchangdu':
            textElement.text = (dianlan)
        elif textElement.name == 'gongbianshuliang':
            textElement.text = (gongbian)
        elif textElement.name == 'zhuanbianshuliang':
            textElement.text = (zhuanbian)
        elif textElement.name == 'duanluqi':
            textElement.text = (duanluqi)
        elif textElement.name == 'beizhu':
            textElement.text = (beizhu)
            textElement.elementPositionY = 0.7
            textElement.elementPositionX = x
        else:
            pass


    # 替换 84.tif 底图
    ditu_dir = ditu_path
    # ditu_dir = '/'.join(ditu_dir)
    ditu_tif = level+'.tif'
    print(ditu_dir)
    print(ditu_tif)
    lyr_84 = arcpy.mapping.ListLayers(mxd, '84.tif', df0)[0]
    lyr_84.replaceDataSource(ditu_dir, 'RASTER_WORKSPACE', ditu_tif)

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
    # outjpeg = out_pic_dir.decode('gbk')+'\\'+dir.split('\\')[-2]
    outjpeg = out_pic_dir.decode('gbk')+'\\'+title.decode('utf-8')
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
    ditu_path = sys.argv[3]
    # print(dirs)
    # print(out_pic_dir)
    # print(dirs)
    dir_list = os.listdir(dirs)
    for dir in dir_list:
        # dir = dirs+'\\'+dir.decode('gbk')+'\\'
        dir = dirs+'\\'+dir+'\\'
        print(dir.decode('gbk'))
        mapping(dir,out_pic_dir,ditu_path)
    # print(sys.argv)
    # pass


if __name__ == '__main__':
    main()
