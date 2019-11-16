# coding=gbk

import simple_tkinter as sg
import os
import sys
import codecs
import yongcheng
import multiprocessing


def kernel_dwg_to_shp(params):
    arcgis_python,input_python_script,input_dir,out_dir = params

    os.system(
        arcgis_python.encode('gbk') + ' ' +
        input_python_script.encode('gbk') + ' ' +
        input_dir.encode('gbk') + ' ' +
        out_dir.encode('gbk')
    )
    sg.Popup('dwg转shp完毕！\n按OK结束'.decode('gbk'))
    pass


def dwg_to_shp():
    if os.path.isfile(os.getcwd() + '/config_dwg_to_shp.cfg'):
        config_r = open(os.getcwd() + '/config_dwg_to_shp.cfg', 'r')
        lines = config_r.readlines()
        param_dic = {}
        for line in lines:
            line = line.split('\n')
            para = line[0].split('=')[0]
            val = line[0].split('=')[1]
            param_dic[para] = val
        # if 'fdir' in param_dic:
        try:
            sg_input_dir = sg.Input(param_dic['input_dir'].decode('gbk'))
            sg_arcgis_python = sg.Input(param_dic['arcgis_python'].decode('gbk'))
            sg_input_python_script = sg.Input(param_dic['input_python_scrip'].decode('gbk'))
            sg_out_dir = sg.Input(param_dic['out_dir'].decode('gbk'))

        except:
            sg_input_dir = sg.Input('')
            sg_arcgis_python = sg.Input('')
            sg_input_python_script = sg.Input('')
            sg_out_dir = sg.Input('')
    else:
        sg_input_dir = sg.Input('')
        sg_arcgis_python = sg.Input('')
        sg_input_python_script = sg.Input('')
        sg_out_dir = sg.Input('')

    layout1 = [
        [sg.Text('arcgis 的 python.exe'.decode('gbk'))],
        [sg_arcgis_python, sg.FileBrowse()],

        [sg.Text('dwg转shp脚本.py'.decode('gbk'))],
        [sg_input_python_script, sg.FileBrowse()],

        [sg.Text('dwg目录'.decode('gbk'))],
        [sg_input_dir, sg.FolderBrowse()],

        [sg.Text('dwg转shp输出目录'.decode('gbk'))],
        [sg_out_dir, sg.FolderBrowse()],
        [sg.OK()]
    ]

    window1 = sg.Window('dwg转shp'.decode('gbk'),font=("Helvetica", 20)).Layout(layout1)
    while 1:

        event1, values1 = window1.Read()
        if event1 is None:
            break
        arcgis_python = values1[0]
        input_python_script = values1[1]
        input_dir = values1[2]
        out_dir = values1[3]

        config = codecs.open(os.getcwd() + '/' + 'config_dwg_to_shp.cfg', 'w')
        config.write('arcgis_python=' + arcgis_python.encode('gbk') + '\n')
        config.write('input_python_scrip=' + input_python_script.encode('gbk') + '\n')
        config.write('input_dir=' + input_dir.encode('gbk') + '\n')
        config.write('out_dir=' + out_dir.encode('gbk') + '\n')

        config.close()
        input_dir = input_dir.replace('/','\\')
        out_dir = out_dir.replace('/','\\')

        params = [arcgis_python,input_python_script,input_dir,out_dir]
        p = multiprocessing.Process(target=kernel_dwg_to_shp, args=[params])
        p.start()
        window1.Close()
        break

def kernel_mapping(params):
    arcgis_python,mapping_script,fdir,out_pic_dir,mapping_input_ditu,heng_template,shu_template = params
    print(arcgis_python)
    os.system(
        arcgis_python.encode('gbk') + ' ' +
        mapping_script.encode('gbk') + ' ' +
        fdir.encode('gbk') + ' ' +
        out_pic_dir.encode('gbk') + ' ' +
        mapping_input_ditu.encode('gbk') + ' ' +
        heng_template.encode('gbk') + ' ' +
        shu_template.encode('gbk')
    )
    sg.Popup('制图完毕！\n按OK结束'.decode('gbk'))
    pass



def mapping():
    if os.path.isfile(os.getcwd() + '/config_mapping.cfg'):
        config_r = open(os.getcwd() + '/config_mapping.cfg', 'r')
        lines = config_r.readlines()
        param_dic = {}
        for line in lines:
            line = line.split('\n')
            para = line[0].split('=')[0]
            val = line[0].split('=')[1]
            param_dic[para] = val
        # if 'fdir' in param_dic:
        try:
            sg_input_dir = sg.Input(param_dic['fdir'].decode('gbk'))
            sg_arcgis_python = sg.Input(param_dic['arcgis_python'].decode('gbk'))
            sg_input_python_script = sg.Input(param_dic['mapping_script'].decode('gbk'))
            sg_input_ditu = sg.Input(param_dic['input_ditu'].decode('gbk'))
            sg_out_pic_dir = sg.Input(param_dic['out_pic_dir'].decode('gbk'))
            sg_heng_template = sg.Input(param_dic['heng_template'].decode('gbk'))
            sg_shu_template = sg.Input(param_dic['shu_template'].decode('gbk'))
        except:
            sg_arcgis_python = sg.Input('')
            sg_input_python_script = sg.Input('')
            sg_input_ditu = sg.Input('')
            sg_input_dir = sg.Input('')
            sg_out_pic_dir = sg.Input('')
            sg_heng_template = sg.Input('')
            sg_shu_template = sg.Input('')
    else:
        sg_arcgis_python = sg.Input('')
        sg_input_dir = sg.Input('')
        sg_input_python_script = sg.Input('')
        sg_input_ditu = sg.Input('')
        sg_out_pic_dir = sg.Input('')
        sg_heng_template = sg.Input('')
        sg_shu_template = sg.Input('')
    layout1 = [
            [sg.Text('arcgis 的 python.exe'.decode('gbk'))],
            [sg_arcgis_python, sg.FileBrowse()],
            [sg.Text('制图脚本.py'.decode('gbk'))],
            [sg_input_python_script, sg.FileBrowse()],
            [sg.Text('底图'.decode('gbk'))],
            [sg_input_ditu, sg.FolderBrowse()],
            [sg.Text('横mxd模板.mxd'.decode('gbk'))],
            [sg_heng_template, sg.FileBrowse()],
            [sg.Text('竖mxd模板.mxd'.decode('gbk'))],
            [sg_shu_template, sg.FileBrowse()],
            [sg.Text('制图目录'.decode('gbk'))],
            [sg_input_dir, sg.FolderBrowse()],
            [sg.Text('图片输出目录'.decode('gbk'))],
            [sg_out_pic_dir, sg.FolderBrowse()],
            [sg.OK()]
    ]
    window1 = sg.Window('arcpy制图'.decode('gbk'),font=("Helvetica", 20)).Layout(layout1)
    while 1:

        event1, values1 = window1.Read()
        if event1 is None:
            break
        arcgis_python = values1[0]
        mapping_script = values1[1]
        mapping_input_ditu = values1[2]
        heng_template = values1[3]
        shu_template = values1[4]
        fdir = values1[5]
        out_pic_dir = values1[6]

        config = codecs.open(os.getcwd() + '/' + 'config_mapping.cfg', 'w')
        config.write('fdir=' + fdir.encode('gbk') + '\n')
        config.write('mapping_script=' + mapping_script.encode('gbk') + '\n')
        config.write('input_ditu=' + mapping_input_ditu.encode('gbk') + '\n')
        config.write('arcgis_python=' + arcgis_python.encode('gbk') + '\n')
        config.write('out_pic_dir=' + out_pic_dir.encode('gbk') + '\n')
        config.write('heng_template=' + heng_template.encode('gbk') + '\n')
        config.write('shu_template=' + shu_template.encode('gbk') + '\n')
        config.close()
        # print(arcgis_python+' '+mapping_script)
        fdir = fdir.replace('/','\\')
        out_pic_dir = out_pic_dir.replace('/','\\')

        params = [arcgis_python,mapping_script,fdir,out_pic_dir,mapping_input_ditu,heng_template,shu_template]
        p = multiprocessing.Process(target=kernel_mapping, args=[params])
        p.start()

        window1.Close()
        break


def kernel_gen_layer(fdir,f_excel):
    yongcheng.main(fdir + '/', f_excel)
    sg.Popup('图层生成完毕！\n按OK结束'.decode('gbk'))

    pass



def gen_layer():

    if os.path.isfile(os.getcwd() + '\\config.cfg'):
        config_r = open(os.getcwd() + '\\config.cfg', 'r')
        lines = config_r.readlines()
        param_dic = {}
        for line in lines:
            line = line.split('\n')
            para = line[0].split('=')[0]
            val = line[0].split('=')[1]
            param_dic[para] = val
        # if 'fdir' in param_dic:
        try:
            sg_input_dir = sg.Input(param_dic['fdir'].decode('gbk'))
            sg_input_excel = sg.Input(param_dic['f_excel'].decode('gbk'))
        except:
            sg_input_excel = sg.Input('')
            sg_input_dir = sg.Input('')
    else:
        sg_input_excel = sg.Input('')
        sg_input_dir = sg.Input('')
    layout1 = [[sg.Text('输入台账Excel'.decode('gbk'))],
               [sg_input_excel, sg.FileBrowse()],
               [sg.Text('定义目录文件'.decode('gbk'))],
               [sg_input_dir, sg.FolderBrowse()],
               [sg.OK()] ]
    window1 = sg.Window('生成图层'.decode('gbk'),font=("Helvetica", 20)).Layout(layout1)

    while 1:

        event1, values1 = window1.Read()
        if event1 is None:
            break
        # print(values1)
        f_excel = values1[0]
        fdir = values1[1]
        print(fdir)
        print(f_excel)
        config = codecs.open(os.getcwd() + '\\' + 'config.cfg', 'w')
        config.write('fdir=' + fdir.encode('gbk') + '\n')
        config.write('f_excel=' + f_excel.encode('gbk') + '\n')
        config.close()

        p = multiprocessing.Process(target=kernel_gen_layer, args=(fdir,f_excel))
        p.start()
        window1.Close()
        break

def merge():
    if os.path.isfile(os.getcwd() + '\\merge_config.cfg'):
        config_r = open(os.getcwd() + '\\merge_config.cfg', 'r')
        lines = config_r.readlines()
        param_dic = {}
        for line in lines:
            line = line.split('\n')
            para = line[0].split('=')[0]
            val = line[0].split('=')[1]
            param_dic[para] = val
        # if 'fdir' in param_dic:
        try:
            sg_input_dir = sg.Input(param_dic['fdir'].decode('gbk'))
            sg_input_out_dir = sg.Input(param_dic['out_dir'].decode('gbk'))
        except:
            sg_input_out_dir = sg.Input('')
            sg_input_dir = sg.Input('')
    else:
        sg_input_out_dir = sg.Input('')
        sg_input_dir = sg.Input('')
    layout1 = [[sg.Text('输入需要合并的文件夹'.decode('gbk'))],
               [sg_input_dir, sg.FolderBrowse()],
               [sg.Text('输出目录'.decode('gbk'))],
               [sg_input_out_dir, sg.FolderBrowse()],
               [sg.OK()] ]
    window1 = sg.Window('生成图层'.decode('gbk'),font=("Helvetica", 20)).Layout(layout1)

    while 1:

        event1, values1 = window1.Read()
        if event1 is None:
            break
        # print(values1)
        indir = values1[0]
        outdir = values1[1]
        print(indir)
        print(outdir)
        # exit()
        config = codecs.open(os.getcwd() + '\\' + 'merge_config.cfg', 'w')
        config.write('fdir=' + indir.encode('gbk') + '\n')
        config.write('out_dir=' + outdir.encode('gbk') + '\n')
        config.close()
        indir = indir.encode('gbk')
        outdir = outdir.encode('gbk')
        yongcheng.Merge().merge_point_annotation_shp(indir,outdir)
        yongcheng.Merge().merge_daoxian(indir,outdir)
        # yongcheng.main(fdir+'/',f_excel)
        sg.Popup('图层生成完毕！\n按OK结束'.decode('gbk'))
        window1.Close()
        break



def main():
    # layout1 = [[sg.Text('1.dwg转shp'.decode('gbk'))],
    #            [sg.Text('2.生成layer'.decode('gbk'))],
    #            [sg.Text('3.出图'.decode('gbk'))],
    #            [sg.OK()]]
    layout1 = [[sg.Radio('1.dwg转shp'.decode('gbk'), "RADIO1")],
              [sg.Radio('2.生成layer'.decode('gbk'), "RADIO1")],
              [sg.Radio('3.出图'.decode('gbk'), "RADIO1")],
               [sg.Radio('4.合并图层'.decode('gbk'), "RADIO1")],
               [sg.OK()]
               ]

    window1 = sg.Window('自动制图'.decode('gbk'),layout1,font=("Helvetica", 20))
    while 1:
        event1, values1 = window1.Read()
        # print(values1)
        if event1 is None:
            break
        if values1[0]:
            dwg_to_shp()
        if values1[1]:
            gen_layer()
        if values1[2]:
            mapping()
        if values1[3]:
            merge()
        # window1.Close()
if __name__ == '__main__':
    main()