# coding=gbk
import os
import arcpy
import sys
import multiprocessing
import copy_reg
import types
from multiprocessing.pool import ThreadPool as TPool
# print(argvs)
# this_root = 'E:\\cui\\'
argvs = sys.argv

output_mxd = {'柱上用户变压器':'zhuanbian',
              '耐张杆塔':'naizhang_ganta',
              'xiangshi_biandianzhan':'xiangbian',
              # '断路器':'duanluqi',
              'duanluqi_changkai':'duanluqi_changkai',
              'duanluqi_changbi':'duanluqi_changbi',
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



class MUTIPROCESS:
    '''
    可对类内的函数进行多进程并行
    由于GIL，多线程无法跑满CPU，对于不占用CPU的计算函数可用多线程
    并行计算加入进度条
    '''

    def __init__(self, func, params):
        self.func = func
        self.params = params
        copy_reg.pickle(types.MethodType, self._pickle_method)
        pass

    def _pickle_method(self, m):
        if m.im_self is None:
            return getattr, (m.im_class, m.im_func.func_name)
        else:
            return getattr, (m.im_self, m.im_func.func_name)

    def run(self, process=6, process_or_thread='p', **kwargs):
        '''
        # 并行计算加进度条
        :param func: input a kenel_function
        :param params: para1,para2,para3... = params
        :param process: number of cpu
        :param thread_or_process: multi-thread or multi-process,'p' or 't'
        :param kwargs: tqdm kwargs
        :return:
        '''
        if 'text' in kwargs:
            kwargs['desc'] = kwargs['text']
            del kwargs['text']

        if process_or_thread == 'p':
            pool = multiprocessing.Pool(process)
        elif process_or_thread == 't':
            pool = TPool(process)
        else:
            raise IOError('process_or_thread key error, input keyword such as "p" or "t"')
        results = pool.imap(self.func, self.params)
        # results = list(tqdm(pool.imap(self.func, self.params), total=len(self.params), **kwargs))
        pool.close()
        pool.join()
        return results





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


def define_projection(fdir):
    # prj_f = fdir + 'extent_lyr.prj'
    prj_content = '''
    GEOGCS["GCS_WGS_1984",DATUM["D_WGS_1984",SPHEROID["WGS_1984",6378137.0,298.257223563]],PRIMEM["Greenwich",0.0],UNIT["Degree",0.0174532925199433],AUTHORITY["EPSG",4326]]
    '''
    fdir = fdir+'\\'
    prj_f = fdir+'projection.prj'
    if not os.path.isfile(prj_f):
        fw = open(prj_f,'w')
        fw.write(prj_content)
        fw.close()
    for f in os.listdir(fdir):
        if f.endswith('.shp'):
            print 'project '+f
            sr = arcpy.SpatialReference(prj_f)
            arcpy.DefineProjection_management(fdir+f, sr)
            # exit()
            pass

    pass



def mapping(dir,out_pic_dir,ditu_path):
    # 横竖选择
    size = sys.argv[6]
    f=open(dir+'\\'+'config.txt','r')
    line = f.readline()
    template = line.split(',')[0]
    level = line.split(',')[1]
    if template == 'heng':
        # mxd_file = 'D:\\project13\\heng_ver5.mxd'
        # mxd_file = this_root+'mxd\\template_heng.mxd'
        # mxd_file = r'E:\cui\190905\template_heng.mxd'
        mxd_file = sys.argv[4]
        if size == 'a0':
            x_beizhu = 8.1702
            # x_tuli = 10.3992
        elif size == 'a3':
            x_beizhu = 7.6115
            # x_tuli = 10.3992
    elif template == 'shu':
        # mxd_file = this_root + 'mxd\\template_shu.mxd'
        # mxd_file = r'E:\cui\190905\template_shu.mxd'
        mxd_file = sys.argv[5]
        if size == 'a0':
            x_beizhu = 8.1702
            # x_tuli = 10.3164
        elif size == 'a3':
            x_beizhu = 6.5893
            # x_tuli = 9.6794
    else:
        mxd_file = None
        x_beizhu = None
        x_tuli = None

    mxd = arcpy.mapping.MapDocument(mxd_file)
    df0 = arcpy.mapping.ListDataFrames(mxd)[0]

    workplace = 'SHAPEFILE_WORKSPACE'
    # title = dir.split('\\')[-2]+'线路沿布图'
    # text_f = open(dir+'\\'+'info.txt','r')
    # line = text_f.readline()
    # shebeimingcheng, qidiandianzhan, weihubanzu,\
    # xianluzongchangdu,jiakong, dianlan, gongbian, \
    # zhuanbian, duanluqi,title,beizhu = line.split(',')

    # print(beizhu)
    title = open(dir+'\\'+'info_title.txt','r').read()
    beizhu = open(dir+'\\'+'info_beizhu.txt','r').read()
    tuli = open(dir+'\\'+'info_tuli.txt','r').read()

    if len(tuli) == 0:
        tuli = ' '

    if len(beizhu) == 0:
        beizhu = ' '

    renyuan = open(dir+'\\'+'info_huizhi.txt','r').read()
    try:
        renyuan_split = renyuan.split('__')
        huizhiren,shenherenyuan,huizhishijian = renyuan_split
    except:
        huizhiren, shenherenyuan, huizhishijian = ' ',' ',' '

    beizhu = new_huanhang(beizhu.decode('utf-8'),30)
    if len(beizhu) == 0:
        beizhu = ' '
    for textElement in arcpy.mapping.ListLayoutElements(mxd, "TEXT_ELEMENT"):
        if textElement.name == 'title':
            textElement.text = (title)
        # elif textElement.name == 'tuli':
        #     textElement.text = (tuli)
        #     textElement.elementPositionY = 1.7
        #     textElement.elementPositionX = x_tuli
        elif textElement.name == 'beizhu':
            textElement.text = (beizhu)
            textElement.elementPositionY = 1.7
            textElement.elementPositionX = x_beizhu


        elif textElement.name == 'huizhiren':
            textElement.text = (huizhiren)
        elif textElement.name == 'shenherenyuan':
            textElement.text = (shenherenyuan)
        elif textElement.name == 'huizhishijian':
            textElement.text = (huizhishijian)
        else:
            pass


    # 替换 84.tif 底图
    ditu_dir = ditu_path
    # ditu_dir = '/'.join(ditu_dir)
    if size == 'a0':
        level = '17'
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
    # arcpy.mapping.ExportToPDF(mxd,outjpeg+'.pdf')
    print 'done'


def kernel_main(params):

    dirs,fdir,out_pic_dir,ditu_path = params
    # dir = dirs+'\\'+dir.decode('gbk')+'\\'
    dir = dirs + '\\' + fdir + '\\'
    # print(dir.decode('gbk'))
    define_projection(dir)
    mapping(dir, out_pic_dir, ditu_path)

    pass


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
    params = []
    for fdir in dir_list:
        params.append([dirs,fdir,out_pic_dir,ditu_path])
        # kernel_main([dirs,fdir,out_pic_dir,ditu_path])
    MUTIPROCESS(kernel_main,params).run(10,'p')
    # print(sys.argv)
    # pass


if __name__ == '__main__':
    main()
