# coding=gbk

import arcpy
import os
import log_process
import time
import sys



# this_root = os.getcwd()+'\\..\\'
this_root = 'e:\\cui\\'

# in_features_point = this_root+'new_test_data\\123.dwg\\Annotation'
# in_features_line = this_root+'new_test_data\\123.dwg\\Polyline'


def mk_dir(dir):
    if not os.path.isdir(dir):
        os.makedirs(dir)

def dwg_to_shp(in_features, output_dir):
    arcpy.FeatureClassToShapefile_conversion(in_features, output_dir)


def check():
    fdir = this_root + '190714\\dwg_to_shp\\'
    flist = os.listdir(fdir)
    for folder in flist:
        file_count = len(os.listdir(fdir+folder))
        if file_count != 12:
            print(folder)
    pass

def define_projection(fdir):
    # prj_f = fdir + 'extent_lyr.prj'
    prj_content = '''
    GEOGCS["GCS_WGS_1984",DATUM["D_WGS_1984",SPHEROID["WGS_1984",6378137.0,298.257223563]],PRIMEM["Greenwich",0.0],UNIT["Degree",0.0174532925199433],AUTHORITY["EPSG",4326]]
    '''
    prj_f = fdir+'\\projection.prj'
    if not os.path.isfile(prj_f):
        fw = open(prj_f,'w')
        fw.write(prj_content)
        fw.close()
    for f in os.listdir(fdir):
        if f.endswith('.shp'):
            print 'project '+f
            sr = arcpy.SpatialReference(prj_f)
            arcpy.DefineProjection_management(fdir+'\\'+f, sr)
            # exit()
            pass

    pass



def run_dwg_to_shp(fdir,out_dir_):
    # log = log_process.Logger('log.log')
    # fdir = this_root+'190905\\永城\\'
    flist = os.listdir(fdir)
    init_time = time.time()
    flag = 0
    for f in flist:
        start = time.time()
        out_dir = out_dir_+f.split('.')[0]
        # print(f)
        # print(out_dir)
        # exit()
        if os.path.isdir(out_dir):
            end = time.time()
            log_process.process_bar(flag, len(flist), init_time, start, end,
                                    str(flag + 1) + '/' + str(len(flist)) + '\n')
            flag += 1
            print(out_dir+'is existed')
            continue
        mk_dir(out_dir)
        # try:
        print(out_dir.decode('gbk'))
        # exit()
        dwg_to_shp(fdir+f+'\\Annotation',out_dir)
        dwg_to_shp(fdir+f+'\\Polyline',out_dir)
        # log.logger.info('\\Annotation')
        # log.logger.info('\\Polyline')
        # except Exception as e:
        # log.logger.error(e)
        end = time.time()
        log_process.process_bar(flag,len(flist),init_time,start,end,str(flag+1)+'/'+str(len(flist))+'\n')
        flag += 1


def rename(fdir):
    # 去除# (  )号
    # fdir = this_root + '190905\\永城\\'
    flist = os.listdir(fdir)
    for f in flist:
        f_new = f.replace('#',' ')
        f_new = f_new.replace('(','')
        f_new = f_new.replace(')','')
        # f_new = f_new.replace('_','.')
        os.rename(fdir+f,fdir+f_new)

def main():
    # rename()
    fdir = sys.argv[1]+'\\'
    rename(fdir)
    out_dir = sys.argv[2]+'\\'
    run_dwg_to_shp(fdir,out_dir)

if __name__ == '__main__':
    main()
