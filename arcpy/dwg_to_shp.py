# coding=utf-8

import arcpy
import os
import log_process
import time

this_root = os.getcwd()+'\\..\\'

# in_features_point = this_root+'new_test_data\\123.dwg\\Annotation'
# in_features_line = this_root+'new_test_data\\123.dwg\\Polyline'

def mk_dir(dir):
    if not os.path.isdir(dir):
        os.makedirs(dir)
# 10kVµ¤1µ¤³£Ïß

def dwg_to_shp(in_features, output_dir):
    arcpy.FeatureClassToShapefile_conversion(in_features, output_dir)


def check():
    fdir = this_root + u'190509\\ÃñÈ¨ÏßÂ·cad\\dwg_to_shp\\'
    flist = os.listdir(fdir)
    for folder in flist:
        file_count = len(os.listdir(fdir+folder))
        if file_count != 12:
            print(folder)
    pass

def run_dwg_to_shp():
    log = log_process.Logger('log.log')
    fdir = this_root+u'190509\\35kVÏßÂ·\\35kVÏßÂ·\\'
    flist = os.listdir(fdir)

    init_time = time.time()
    flag = 0
    for f in flist:
        start = time.time()
        out_dir = this_root+u'190509\\35kVÏßÂ·\\dwg_to_shp\\'+f.split('.')[0]
        if os.path.isdir(out_dir):
            flag += 1
            continue
        mk_dir(out_dir)
        try:
            dwg_to_shp(fdir+f+'\\Annotation',out_dir)
            dwg_to_shp(fdir+f+'\\Polyline',out_dir)
            log.logger.info('\\Annotation')
            log.logger.info('\\Polyline')
        except Exception as e:
            log.logger.error(e)
        end = time.time()
        log_process.process_bar(flag,len(flist),init_time,start,end,str(flag+1)+'/'+str(len(flist))+'\n')
        flag += 1
        # print(fdir+f)
    # dwg_to_shp(in_features_line,this_root)


def main():
    run_dwg_to_shp()

if __name__ == '__main__':
    main()
