# coding=gbk

import xlrd
import xlwt
import codecs
import os
import pandas as pd
import numpy as np
from tqdm import tqdm
import multiprocessing


def mkdir(fdir,force_mkdir=False):
    if not os.path.isdir(fdir):
        if force_mkdir == False:
            os.mkdir(fdir)
        elif force_mkdir == True:
            os.makedirs(fdir)
        else:
            raise IOError('force_mkdir type error, expect bool')

def excel_to_pd(f):
    df = pd.read_excel(f, header=1,na_values='')

    title = '供电单位,台区编号,台区名称,用户编号,用户名称,电能表计资产号,供电电压,用电类别,用户分类'
    title = title.split(',')

    dic = {}
    # for t in title:
    #     dic[t.decode('gbk')] = []
    # print(dic)
    title_gbk = []
    for t in title:
        # print(t.decode('gbk'))
        title_gbk.append(t.decode('gbk'))
        if t.decode('gbk') in df:
            series = df[t.decode('gbk')]
            # df0.append(series)
            dic[t.decode('gbk')] = series
    # dic['编号'.decode('gbk')] = range(1,len(series)+1)
    df0 = pd.DataFrame(dic,columns=title_gbk)
    return df0

def split(df):
    # 1 get unique
    # df._ixs()
    invalid_chr = list('\/:*?"<>|')
    df = df.fillna(value='')
    name_title = '台区名称'.decode('gbk')
    series = df[name_title]
    unique_vals = []
    for i in set(series):
        if i == '':
            continue
        unique_vals.append(i)
    split_dic = {}
    # flag = 0
    for name in unique_vals:
        # flag+=1
        # if flag == 2:
        #     break
        indexes = []
        for i in range(len(df)):
            if df[name_title][i] == name:
                indexes.append(i)
                # print(df['编号'.decode('gbk')][i])
        df_new = df._ixs(indexes)
        df_new.insert(0,'编号'.decode('gbk'),range(1,len(indexes)+1))
        for char in name:
            if char in invalid_chr:
                name = name.replace(char,'')
        # df_new.to_excel(out_dir+'\\{}.xlsx'.format(name.encode('gbk')),index=False,index_label='编号'.decode('gbk'))
        split_dic[name] = df_new
    return split_dic


def write_excel(title,content_dic,output_file):
    # add new excel
    # title: merge title
    # content_dic =
    # [
    # ['1','2','3',···，'10'],
    # ['1', '2', '3',···，'10']
    # ]
    book = xlwt.Workbook()
    sheet = book.add_sheet('Sheet1')
    # set width
    # 256*20 256是基准单位 20是字符
    sheet.col(0).width = 256*5
    sheet.col(1).width = 256*10
    sheet.col(2).width = 256*((len('0000804828'))+2)
    sheet.col(3).width = 256*10
    sheet.col(4).width = 256*((len('7390311437'))+2)
    sheet.col(5).width = 256*((len('4130001000000229954226')))
    sheet.col(6).width = 256*((len('4130001000000229954226'))+2)
    sheet.col(7).width = 256*10
    sheet.col(8).width = 256*16
    sheet.col(9).width = 256*10

    # style
    style = xlwt.XFStyle()
    # font
    fnt = xlwt.Font()
    fnt.name = '宋体'.decode('gbk')
    fnt.height = 320
    fnt.bold = True
    style.font = fnt
    # alignment
    align = xlwt.Alignment()
    align.horz = 0x02
    style.alignment = align
    # border
    border = xlwt.Borders()
    border.left = 0x01
    border.right = 0x01
    border.top = 0x01
    border.bottom = 0x01
    style.borders = border
    # print content_dic[2][2].decode('gbk')

    sheet.write_merge(0, 0, 0, 9, (title + '用户信息').decode('gbk'), style)

    # style1
    style1 = xlwt.XFStyle()
    # font
    fnt = xlwt.Font()
    fnt.name = '宋体'.decode('gbk')
    fnt.height = 200
    fnt.bold = True
    style1.font = fnt
    # alignment
    align = xlwt.Alignment()
    align.horz = 0x02
    align.wrap = 0x01
    style1.alignment = align
    # border
    border = xlwt.Borders()
    border.left = 0x01
    border.right = 0x01
    border.top = 0x01
    border.bottom = 0x01
    style1.borders = border

    sheet.write(1, 0, '序号'.decode('gbk'), style1)
    sheet.write(1, 1, '供电单位'.decode('gbk'), style1)
    sheet.write(1, 2, '台区编号'.decode('gbk'), style1)
    sheet.write(1, 3, '台区名称'.decode('gbk'), style1)
    sheet.write(1, 4, '用户编号'.decode('gbk'), style1)
    sheet.write(1, 5, '用户名称'.decode('gbk'), style1)
    sheet.write(1, 6, '电能表资产编号'.decode('gbk'), style1)
    sheet.write(1, 7, '供电电压'.decode('gbk'), style1)
    sheet.write(1, 8, '用电类别'.decode('gbk'), style1)
    sheet.write(1, 9, '用户分类'.decode('gbk'), style1)

    # style2
    style2 = xlwt.XFStyle()
    # font
    fnt = xlwt.Font()
    fnt.name = '宋体'.decode('gbk')
    fnt.height = 200
    # fnt.bold = True
    style2.font = fnt
    # alignment
    align = xlwt.Alignment()
    align.horz = 0x02
    align.wrap = 0x01
    style2.alignment = align
    # border
    border = xlwt.Borders()
    border.left = 0x01
    border.right = 0x01
    border.top = 0x01
    border.bottom = 0x01
    style2.borders = border

    j = 0
    for i in range(len(content_dic)):
        sheet.write(j + 2, 0, content_dic[i][0], style2)
        sheet.write(j + 2, 1, content_dic[i][1], style2)
        sheet.write(j + 2, 2, content_dic[i][2], style2)
        sheet.write(j + 2, 3, content_dic[i][3], style2)
        sheet.write(j + 2, 4, content_dic[i][4], style2)
        sheet.write(j + 2, 5, content_dic[i][5], style2)
        sheet.write(j + 2, 6, content_dic[i][6], style2)
        sheet.write(j + 2, 7, content_dic[i][7], style2)
        sheet.write(j + 2, 8, content_dic[i][8], style2)
        sheet.write(j + 2, 9, content_dic[i][9], style2)
        j += 1

    book.save(output_file)

def pd_to_excel(df_dic,outdir):
    # outdir = r'C:\Users\ly\PycharmProjects\split_excel\test\\'
    for name in df_dic:
        output_file = outdir + name+'.xls'
        # print(df_dic[name])
        # content
        df = df_dic[name]
        df = df.astype('unicode')
        content_list = []
        for i in range(len(df)):
            content = []
            for c in df:
                # print(c)
                val = df[c]._ixs(i)
                # print(val)
                content.append(val)
            content_list.append(content)
        # print(content_list)
        name = name.encode('gbk')
        write_excel(name,content_list,output_file)



def run(params):
    f, out_dir = params
    mkdir(out_dir,force_mkdir=True)
    df = excel_to_pd(f)
    split_dic = split(df)
    pd_to_excel(split_dic,out_dir)

def main():
    # f = r'E:\cui\表格分台区\表格分台区\表格分台区\用户信息\板木所.xlsx'
    # out_dir = ''
    # df = excel_to_pd(f)
    # split_dic = split(df)
    # pd_to_excel(split_dic,out_dir)
    fdir = r'E:\cui\表格分台区\表格分台区\表格分台区\用户信息\\'
    outdir = r'E:\cui\表格分台区\表格分台区\表格分台区\output\\'
    flist = []
    out_dir_list = []
    for f in os.listdir(fdir):
        flist.append(fdir+f)
        out_dir_list.append(outdir+f.split('.')[0]+'\\')
    params = []
    for i in range(len(flist)):
        params.append([flist[i],out_dir_list[i].decode('gbk')])
        # print flist[i]
    # for j in params:
    #     run(j)
        # exit()
    pool = multiprocessing.Pool()
    # pool.map(run,params)
    list(tqdm(pool.imap(run, params), total=len(params), ncols=50))
    pool.close()
    pool.join()

if __name__ == '__main__':
    main()
