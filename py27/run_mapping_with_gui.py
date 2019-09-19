# coding=gbk

import simple_tkinter as sg
import os
import sys
import codecs
import yongcheng

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

        os.system(
            arcgis_python.encode('gbk') + ' ' +
            input_python_script.encode('gbk') + ' ' +
            input_dir.encode('gbk') + ' ' +
            out_dir.encode('gbk')
        )
        sg.Popup('dwg转shp完毕！\n按OK结束'.decode('gbk'))
        window1.Close()
        break


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
            sg_out_pic_dir = sg.Input(param_dic['out_pic_dir'].decode('gbk'))
            sg_heng_template = sg.Input(param_dic['heng_template'].decode('gbk'))
            sg_shu_template = sg.Input(param_dic['shu_template'].decode('gbk'))
        except:
            sg_arcgis_python = sg.Input('')
            sg_input_python_script = sg.Input('')
            sg_input_dir = sg.Input('')
            sg_out_pic_dir = sg.Input('')
            sg_heng_template = sg.Input('')
            sg_shu_template = sg.Input('')
    else:
        sg_arcgis_python = sg.Input('')
        sg_input_dir = sg.Input('')
        sg_input_python_script = sg.Input('')
        sg_out_pic_dir = sg.Input('')
        sg_heng_template = sg.Input('')
        sg_shu_template = sg.Input('')
    layout1 = [
            [sg.Text('arcgis 的 python.exe'.decode('gbk'))],
            [sg_arcgis_python, sg.FileBrowse()],
            [sg.Text('制图脚本.py'.decode('gbk'))],
            [sg_input_python_script, sg.FileBrowse()],
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
        # print(values1)
        arcgis_python = values1[0]
        mapping_script = values1[1]
        heng_template = values1[2]
        shu_template = values1[3]
        fdir = values1[4]
        out_pic_dir = values1[5]

        # print(fdir)
        # print(mapping_script)
        # print(arcgis_python)
        # print(out_pic_dir)
        config = codecs.open(os.getcwd() + '/' + 'config_mapping.cfg', 'w')
        config.write('fdir=' + fdir.encode('gbk') + '\n')
        config.write('mapping_script=' + mapping_script.encode('gbk') + '\n')
        config.write('arcgis_python=' + arcgis_python.encode('gbk') + '\n')
        config.write('out_pic_dir=' + out_pic_dir.encode('gbk') + '\n')
        config.write('heng_template=' + heng_template.encode('gbk') + '\n')
        config.write('shu_template=' + shu_template.encode('gbk') + '\n')
        config.close()
        # print(arcgis_python+' '+mapping_script)
        fdir = fdir.replace('/','\\')
        out_pic_dir = out_pic_dir.replace('/','\\')
        os.system(
                arcgis_python.encode('gbk')+' '+
                mapping_script.encode('gbk')+' '+
                fdir.encode('gbk')+' '+
                out_pic_dir.encode('gbk')+' '+
                heng_template.encode('gbk')+' '+
                shu_template.encode('gbk')
        )
        sg.Popup('制图完毕！\n按OK结束'.decode('gbk'))
        window1.Close()
        break


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

        yongcheng.main(fdir+'/',f_excel)
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
        # window1.Close()
if __name__ == '__main__':
    main()