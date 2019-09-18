# coding=gbk
import os
import read_excel_ver3
import write_point as wp
import write_line as wl
import time


fname3 = 'D:\\project13\\input\\����.xlsx'
fname4 = 'D:\\project13\\input\\ͼ����Ϣ.xls'

yonghujierudian = 1
mapping_break_flag = 0


#1���ɱ�ѹ��shp��
def bianyaqi_shp(r,f_dir):
    bianyaqi_lon, bianyaqi_lat, bianyaqi_label = r.didianyapeidianxiang()
    bianyaqi_label = bianyaqi_label
    directory = 'D:/project13/output//'+f_dir
    if not os.path.isdir(directory):
        os.mkdir(directory)
    fname = directory+'\\bianyaqi.shp'

    wp.point_to_shp([[bianyaqi_lon,bianyaqi_lat,str(bianyaqi_lon),str(bianyaqi_lat),bianyaqi_label,'','']],fname)

#2���ɵ��߸�shp��
def dianxiangan_dian_shp(r,f_dir):
    if r.dianxiangan():
        dianxiangan_lon, dianxiangan_lat, dianxiangan_name, dianxiangan_qianduan = r.dianxiangan()
        directory = 'D:\\project13\\output\\\\'+f_dir
        dianxiangan_coor_dic = {}
        for i in range(len(dianxiangan_lon)):
            dianxiangan_coor_dic[i] = [dianxiangan_lon[i],dianxiangan_lat[i],dianxiangan_qianduan[i],dianxiangan_name[i]]
        if not os.path.isdir(directory):
            os.mkdir(directory)
        fname = directory+'\\dianxiangan'

        point=[]
        for i in dianxiangan_coor_dic:
            point.append([dianxiangan_coor_dic[i][0],dianxiangan_coor_dic[i][1],dianxiangan_coor_dic[i][3],dianxiangan_coor_dic[i][2],'','',''])
        wp.point_to_shp(point,fname+'.shp')


#3����ǽ֧��shp��
def qiangzhijia_shp(r,f_dir):
    if r.diyaqiangzhijia():
        lon_list,lat_list,qianduan,_ = r.diyaqiangzhijia()
        directory = 'D:\\project13\\output\\\\'+f_dir
        if not os.path.isdir(directory):
            os.mkdir(directory)
        qiangzhijia_coor_dic={}
        for i in range(len(lon_list)):
            qiangzhijia_coor_dic[i] = [lon_list[i],lat_list[i],qianduan[i]]
        fname = directory+'\\qiangzhijia'

        point=[]
        for i in qiangzhijia_coor_dic:

            point.append([qiangzhijia_coor_dic[i][0],qiangzhijia_coor_dic[i][1],qiangzhijia_coor_dic[i][2],'','','',''])

        wp.point_to_shp(point,fname+'.shp')

#4���ɺ�ɫ��ɫ͸��������
def jiliangxiang_shp(r,f_dir):
    transparent, red, black = r.full_and_spare_and_transparent_jiliangxiang()
    directory = 'D:\\project13\\output\\\\'+f_dir
    if not os.path.isdir(directory):
        os.mkdir(directory)
    fname_trans = directory+'\\transparent'
    fname_red = directory+'\\red'
    fname_black = directory+'\\black'

    point_trans = []
    point_red = []
    point_black = []

    for i in transparent:
        point_trans.append([transparent[i][0],transparent[i][1],' '.join(transparent[i][2]),i,'','','',''])
    for i in range(len(red)):
        point_red.append([red[i][1],red[i][2],'',red[i][0],red[i][3],'',''])
    for i in range(len(black)):
        point_black.append([black[i][1],black[i][2],'',black[i][0],black[i][3],'',''])
    # print len(point_trans)
    # print len(red)
    # print len(black)
    wp.point_to_shp(point_trans,fname_trans+'.shp')
    wp.point_to_shp(point_red,fname_red+'.shp')
    wp.point_to_shp(point_black,fname_black+'.shp')


#5���ӵ��߸�
#5.1���ɵ��߸�������������
def dianxiangan_line(r):
    if r.dianxiangan():
        lon_list,lat_list,name_list,qianduan_list = r.dianxiangan()
        dianxiangan_coor_dic = {}
        for i in range(len(lon_list)):
            dianxiangan_coor_dic[i] = lon_list[i],lat_list[i],name_list[i],qianduan_list[i]
        p1_list=[]
        p2_list=[]

        for i in dianxiangan_coor_dic:
            if '�����'.decode('gbk') in dianxiangan_coor_dic[i][3]:
                # print dianxiangan_coor_dic[i][3]
                bianyaqi_lon, bianyaqi_lat, bianyaqi_label = r.didianyapeidianxiang()
                # print 'peidianxiang',bianyaqi_lon, bianyaqi_lat
                p2_list.append((bianyaqi_lon, bianyaqi_lat))
            elif 'ǽ֧��'.decode('gbk') in dianxiangan_coor_dic[i][3]:
                lon_list,lat_list,qianduan,qiangzhijia_name = r.diyaqiangzhijia()
                qiangzhijia_coor_dic={}
                for j in range(len(lon_list)):
                    qiangzhijia_coor_dic[j] = [lon_list[j],lat_list[j],qiangzhijia_name[j]]
                for k in qiangzhijia_coor_dic:
                    if dianxiangan_coor_dic[i][3] == qiangzhijia_coor_dic[k][2]:
                        p2_list.append((qiangzhijia_coor_dic[k][0],qiangzhijia_coor_dic[k][1]))
            elif '��֧��'.decode('gbk') in dianxiangan_coor_dic[i][3]:
                fenzhixiang_coor_dic = r.fenzhiiang()
                for j in fenzhixiang_coor_dic:
                    if dianxiangan_coor_dic[i][3] == fenzhixiang_coor_dic[j][2]:
                        p2_list.append((fenzhixiang_coor_dic[j][0],fenzhixiang_coor_dic[j][1]))


            p1_list.append(dianxiangan_coor_dic[i][:2])
            for j in dianxiangan_coor_dic:
                if dianxiangan_coor_dic[i][3] == dianxiangan_coor_dic[j][2]:
                    p2_list.append(dianxiangan_coor_dic[j][:2])

        temp_p1 = []
        temp_p2 = []

        for i in dianxiangan_coor_dic:
            temp_p1.append(dianxiangan_coor_dic[i][2])
            temp_p2.append(dianxiangan_coor_dic[i][3])
        for i in temp_p2:
            if not i in temp_p1:
                print '########',i,'########'


        # print p1_list
        # print p2_list
        # print len(p1_list)
        # print len(p2_list)
        distancestr = []
        distance = []
        # print p1_list
        # print p2_list
        # print len(p1_list)
        # print len(p2_list)
        for i in range(len(p1_list)):

            distancestr.append(str(round(wl.GetDistance(p1_list[i][0],p1_list[i][1],p2_list[i][0],p2_list[i][1]),2)))
            distance.append(wl.GetDistance(p1_list[i][0],p1_list[i][1],p2_list[i][0],p2_list[i][1]))
        total_distance = round(sum(distance),2)


        return p1_list,p2_list,distancestr,total_distance
#5.2���ɵ��߸���shp
def dianxiangan_line_shp(r,f_dir):
    if dianxiangan_line(r):
        p1_list,p2_list,distance,total_distance = dianxiangan_line(r)

        directory = 'D:\\project13\\output\\\\'+f_dir
        if not os.path.isdir(directory):
            os.mkdir(directory)
        fname = directory+'\\dianxiangan_line'

        lines = []
        for i in range(len(p1_list)):
            lines.append([p1_list[i],p2_list[i],distance[i],'','','',''])
        wl.line_to_shp(lines,fname+'.shp')


#6����ǽ֧��
#6.1����ǽ֧������߸�������������
def qiangzhijia_line(r):
    if r.dianxiangan():
        dianxiangan_lon,dianxiangan_lat,dianxiangan_name,_ = r.dianxiangan()
        qiangzhijia_lon,qiangzhijia_lat,qianduan_name,_ = r.diyaqiangzhijia()
        dianxiangan_coor_dic = {}
        for i in range(len(dianxiangan_lon)):
            dianxiangan_coor_dic[i] = [dianxiangan_lon[i],dianxiangan_lat[i],dianxiangan_name[i]]
        qiangzhijia_coor_dic ={}
        for i in range(len(qiangzhijia_lon)):
            qiangzhijia_coor_dic[i] = [qiangzhijia_lon[i],qiangzhijia_lat[i],qianduan_name[i]]

        p1_list=[];p2_list=[]
        for i in qiangzhijia_coor_dic:
            for j in dianxiangan_coor_dic:
                if qiangzhijia_coor_dic[i][2] == dianxiangan_coor_dic[j][2]:
                    p1_list.append([qiangzhijia_coor_dic[i][0],qiangzhijia_coor_dic[i][1]])
                    p2_list.append([dianxiangan_coor_dic[j][0],dianxiangan_coor_dic[j][1]])

        distancestr = []
        distance = []
        for i in range(len(p1_list)):

            distancestr.append(str(round(wl.GetDistance(p1_list[i][0],p1_list[i][1],p2_list[i][0],p2_list[i][1]),2)))
            distance.append(wl.GetDistance(p1_list[i][0],p1_list[i][1],p2_list[i][0],p2_list[i][1]))
        total_distance = round(sum(distance),2)
        return p1_list,p2_list,total_distance,distancestr

#6.2��ǽ֧������߸�����shp
def qiangzhijia_line_shp(r,f_dir):
    if qiangzhijia_line(r):
        p1_list,p2_list,total_distance,distancestr = qiangzhijia_line(r)
        directory = 'D:\\project13\\output\\\\'+f_dir
        if not os.path.isdir(directory):
            os.mkdir(directory)
        fname = directory+'\\qiangzhijia_line.shp'
        lines = []
        for i in range(len(p1_list)):
            lines.append([p1_list[i],p2_list[i],distancestr[i],'','','',''])
        wl.line_to_shp(lines,fname)


#6.2����ǽ֧��֮��������������
def qiangzhijia_line2(r):
    if r.diyaqiangzhijia():
        qiangzhijia_lon,qiangzhijia_lat,qianduan_name,name = r.diyaqiangzhijia()
        qiangzhijia_coor_dic = {}
        for i in range(len(qiangzhijia_lon)):
            qiangzhijia_coor_dic[i] = [qiangzhijia_lon[i],qiangzhijia_lat[i],qianduan_name[i],name[i]]
        p1_list=[];p2_list=[]
        for i in qiangzhijia_coor_dic:
            for j in qiangzhijia_coor_dic:
                if qiangzhijia_coor_dic[i][2] == qiangzhijia_coor_dic[j][3]:
                    p1_list.append([qiangzhijia_coor_dic[i][0],qiangzhijia_coor_dic[i][1]])
                    p2_list.append([qiangzhijia_coor_dic[j][0],qiangzhijia_coor_dic[j][1]])

        distancestr = []
        distance = []
        for i in range(len(p1_list)):

            distancestr.append(str(round(wl.GetDistance(p1_list[i][0],p1_list[i][1],p2_list[i][0],p2_list[i][1]),2)))
            distance.append(wl.GetDistance(p1_list[i][0],p1_list[i][1],p2_list[i][0],p2_list[i][1]))
        total_distance = round(sum(distance),2)

        return p1_list,p2_list,total_distance,distancestr

#6.3����ǽ֧��֮������shp
def qiangzhijia_line2_shp(r,f_dir):
    if qiangzhijia_line2(r):
        p1_list,p2_list,_,distancestr = qiangzhijia_line2(r)
        # print p1_list
        # print p2_list
        directory = 'D:\\project13\\output\\\\'+f_dir
        if not os.path.isdir(directory):
            os.mkdir(directory)
        fname = directory+'\\qiangzhijia_line2.shp'
        lines = []
        for i in range(len(p1_list)):
            lines.append([p1_list[i],p2_list[i],distancestr[i],'','','',''])
        wl.line_to_shp(lines,fname)


#7���ɵ�����
def dianlanxian_line_shp(r,f_dir):
    p1_list, p2_list = r.dianlan()
    directory = 'D:\\project13\\output\\\\'+f_dir
    if not os.path.isdir(directory):
        os.mkdir(directory)
    fname = directory+'\\dianlan.shp'
    lines = []
    for i in range(len(p1_list)):
        distancestr = (str(round(wl.GetDistance(p1_list[i][0],p1_list[i][1],p2_list[i][0],p2_list[i][1]),2)))
        lines.append([p1_list[i],p2_list[i],distancestr,'','','',''])
    wl.line_to_shp(lines,fname)

def dianlanian_line_distance(r):
    p1_list, p2_list = r.dianlan()
    distance = []
    for i in range(len(p1_list)):
        distance.append(round(wl.GetDistance(p1_list[i][0],p1_list[i][1],p2_list[i][0],p2_list[i][1]),2))
    total_distance = sum(distance)
    return total_distance

#8���ɷ�֧��
def fenzhixiang_points(r,f_dir):
    coor_dic = r.fenzhiiang()
    directory = 'D:\\project13\\output\\\\'+f_dir
    if not os.path.isdir(directory):
        os.mkdir(directory)
    fname = directory+'\\fenzhixiang'
    point=[]
    if coor_dic:
        for i in coor_dic:
            point.append((coor_dic[i][0],coor_dic[i][1],coor_dic[i][2],'','','',''))
        # print point
        wp.point_to_shp(point,fname+'.shp')




#9����
#9.1�жϳ�ͼ��������
def select_vertical_horizontal(r,f_dir):
    # transparent = transparent_jiliangxiang(r)

    transparent, red, black = r.full_and_spare_and_transparent_jiliangxiang()
    lon=[]
    lat=[]
    if r.dianxiangan():
        lon_list,lat_list,name_list,qianduan_list = r.dianxiangan()
        for i in range(len(lon_list)):
            lon.append(lon_list[i])
            lat.append(lat_list[i])
    bianyaqi_lon,bianyaqi_lat,_ = r.didianyapeidianxiang()
    directory = 'D:\\project13\\output\\'+f_dir
    if not os.path.isdir(directory):
        os.mkdir(directory)
    fname = directory+'\\extent_lyr.shp'

    for i in transparent:
        lon.append(transparent[i][0])
        lat.append(transparent[i][1])

    lon.append(bianyaqi_lon)
    lat.append(bianyaqi_lat)
    # print max(lon),min(lon)
    # print max(lat),min(lat)
    # print (max(lon)-min(lon))/(max(lat)-min(lat))
    fw = open(directory+'\\select.txt','w')

    if (max(lon)-min(lon))/(max(lat)-min(lat)) >1:
        fw.write('heng')
    else:
        fw.write('shu')
    fw.close()
    lon_offset=(max(lon)-min(lon))*.1
    lat_offset=(max(lat)-min(lat))*.1
    lon_list=[max(lon)+lon_offset,min(lon)-lon_offset]
    lat_list=[max(lat)+lat_offset,min(lat)-lon_offset]
    points=[]
    for i in range(2):
        points.append([lon_list[i],lat_list[i],'','','','',''])
    wp.point_to_shp(points,fname)

#9.2����ͼ����Ϣ�ı�
def gen_text_info(r,f_dir):

    #�������λ�Ͱ�װ����
    total = r.count_total_jiliangxiang()
    installed = r.count_intalled_jiliangxiang()

    #��·����
    if dianxiangan_line(r):
        p1_list, p2_list, distancestr, distance1 = dianxiangan_line(r)
    else:
        distance1 = 0

    if qiangzhijia_line2(r):
        p1_list,p2_list,distance2,_ = qiangzhijia_line2(r)
    else:
        distance2 = 0

    if qiangzhijia_line(r):
        p1_list,p2_list,distance3,_ = qiangzhijia_line(r)
    else:
        distance3 = 0

    distance4 = dianlanian_line_distance(r)
    total_distance = distance1 + distance2 + distance3 + distance4

    #PMS_ID
    info_dic = r.info()
    PMS_ID = r.legend_info()

    for i in info_dic:
        if PMS_ID == info_dic[i][0]:
            taiqu_code,taiqu_name,bianyaqi_type,bianyaqi_content,prime_line_type,branch_line_type = info_dic[i][1],info_dic[i][2],info_dic[i][3],info_dic[i][4],info_dic[i][5],info_dic[i][6]
            import codecs
            fw = codecs.open('D:\\project13\\output\\\\'+f_dir+'\\info.txt','w','utf-8')
            fw.write(taiqu_code+',')
            fw.write(taiqu_name+',')
            fw.write(str(bianyaqi_type)+',')
            fw.write(str(bianyaqi_content)+',')
            fw.write(str(round(total_distance/1000,2))+' KM,')
            fw.write(str(prime_line_type)+',')
            fw.write(str(branch_line_type)+',')
            fw.write(str(total)+',')
            fw.write(str(installed))
            fw.close()



def yonghujierudian_line(r):
    if r.diyaqiangzhijia():
        lon_list,lat_list,qianduan_list,_ = r.yonghujierudian()
        yonghujierudian_coor_dic = {}
        for i in range(len(lon_list)):
            yonghujierudian_coor_dic[i] = [lon_list[i],lat_list[i],qianduan_list[i]]

        lon_list,lat_list,qianduan,qiangzhijia_name = r.diyaqiangzhijia()
        qiangzhijia_coor_dic={}
        for j in range(len(lon_list)):
            qiangzhijia_coor_dic[j] = [lon_list[j],lat_list[j],qiangzhijia_name[j]]


        p1_list = []
        p2_list = []
        for i in yonghujierudian_coor_dic:
            for j in qiangzhijia_coor_dic:
                if yonghujierudian_coor_dic[i][2] == qiangzhijia_coor_dic[j][2]:
                    p1_list.append(yonghujierudian_coor_dic[i][:2])
                    p2_list.append(qiangzhijia_coor_dic[j][:2])

        return p1_list,p2_list


def yonghujierudian_line_shp(r,f_dir):
    if yonghujierudian_line(r):
        p1_list,p2_list = yonghujierudian_line(r)

        directory = 'D:\\project13\\output\\\\'+f_dir
        if not os.path.isdir(directory):
            os.mkdir(directory)
        fname = directory+'\\yonghujierudian_line'

        lines = []
        for i in range(len(p1_list)):
            lines.append([p1_list[i],p2_list[i],'','','','',''])
        wl.line_to_shp(lines,fname+'.shp')


def yonghujierudian_jiliangxiang_line(r):
    jiliangxiang_dic = r.jiliangxiang()
    lon_list,lat_list,qianduan_list,shebei_list = r.yonghujierudian()
    yonghujierudian_dic = {}
    for i in range(len(lon_list)):
        yonghujierudian_dic[i] = [lon_list[i],lat_list[i],shebei_list[i]]

    p1_list = []
    p2_list = []
    for i in jiliangxiang_dic:
        for j in yonghujierudian_dic:
            if jiliangxiang_dic[i][4] == yonghujierudian_dic[j][2]:
                p1_list.append([jiliangxiang_dic[i][0],jiliangxiang_dic[i][1]])
                p2_list.append([yonghujierudian_dic[j][0],yonghujierudian_dic[j][1]])
    return p1_list,p2_list


def yonghujierudian_jiliangxiang_line_shp(r,f_dir):
    p1_list,p2_list = yonghujierudian_jiliangxiang_line(r)

    directory = 'D:\\project13\\output\\\\'+f_dir
    if not os.path.isdir(directory):
        os.mkdir(directory)
    fname = directory+'\\yonghujierudian_jiliangxiang_line'

    lines = []
    for i in range(len(p1_list)):
        lines.append([p1_list[i],p2_list[i],'','','','',''])
    wl.line_to_shp(lines,fname+'.shp')



def check_files():
    root = 'D:/project13/input/data/'
    flist_root = os.listdir(root)
    fname1_list = []
    fname2_list = []
    for i in flist_root:
        flist_dir = os.listdir(root+i)
        for j in flist_dir:
            # print(j.decode('gbk'))
            # print(j.decode('gbk'))
            # print((j.decode('gbk')).encode('gbk'))
            # print '����������ܱ��Ĺ�ϵ.xls'.decode('gbk').encode('gbk')
            # print(j.decode('gbk') == '����������ܱ��Ĺ�ϵ.xls'.decode('gbk'))
            if '̨��' not in j and 'CAD' not in j and '������' not in j and (j.endswith('xls') or j.endswith('xlsx')):
                fname1_list.append(root+i+'/'+j)


            if j.decode('gbk') == '����������ܱ��Ĺ�ϵ.xls'.decode('gbk') or j.decode('gbk') == '����������ܱ��Ĺ�ϵ.xlsx'.decode('gbk'):
                # print(j.decode('gbk'),1321321)
                fname2_list.append(root + i + '/' + j)


        # exit()
    j = 0

    for i in range(len(fname1_list)):
        j +=1
        print j
        fname1 = fname1_list[i]
        print fname1.decode('gbk')

        fname2 = fname2_list[i]
        print fname2.decode('gbk')
        print '---------------------'
        # exit()
        break_flag = 0
        if fname1.split('/')[4] != fname2.split('/')[4]:
            print fname1.decode('gbk').split('/')[4]
            print 'error'
            print fname2.decode('gbk').split('/')[4]
            break_flag = 1
        if break_flag == 1:
            raise IOError('this directory is invalid')

def delete_output():
    fdir = r'd:\\project13\\output\\'
    if os.path.isdir(fdir):
        folders = os.listdir(fdir)
        for folder in folders:
            file_list = os.listdir(fdir+folder)
            for f in file_list:
                os.remove(fdir+folder+'\\'+f)
            os.removedirs(fdir+folder)



def main(fname1,fname2,fname3,fname4,output_dir):
    #��������ļ���
    if not os.path.isdir('d:\\project13\\output_pic\\'):
        os.mkdir('d:\\project13\\output_pic\\')
    if not os.path.isdir('d:\\project13\\output\\'):
        os.mkdir('d:\\project13\\output\\')


    for i in os.listdir('d:/project13/input/data/'):
        if not os.path.isdir('d:/project13/output/'+i):
            os.mkdir('d:\\project13\\output\\'+i)
            # print i.decode('gbk')



    r = read_excel_ver3.ReadExcel(fname1,fname2,fname3,fname4)
    bianyaqi_shp(r,output_dir)
    print '��ѹ���������'.decode('gbk')
    dianxiangan_dian_shp(r,output_dir)
    print '���߸˻������'.decode('gbk')
    qiangzhijia_shp(r,output_dir)
    print 'ǽ֧�ܻ������'.decode('gbk')
    jiliangxiang_shp(r,output_dir)
    print '������������'.decode('gbk')
    dianxiangan_line_shp(r,output_dir)
    print '���߸��������'.decode('gbk')
    qiangzhijia_line_shp(r,output_dir)
    print 'ǽ֧������߸��������'.decode('gbk')
    qiangzhijia_line2_shp(r,output_dir)
    print 'ǽ֧����ǽ֧���������'.decode('gbk')
    select_vertical_horizontal(r,output_dir)
    print 'extent layer�������'.decode('gbk')
    gen_text_info(r,output_dir)
    print '��Ϣ�ļ��������'.decode('gbk')
    dianlanxian_line_shp(r,output_dir)
    print '�������������'.decode('gbk')
    fenzhixiang_points(r,output_dir)
    print '��֧��������'.decode('gbk')
    if yonghujierudian:
        yonghujierudian_line_shp(r,output_dir)
        yonghujierudian_jiliangxiang_line_shp(r,output_dir)
        print '�û������'.decode('gbk')

    print '--------------------------------------------'




if __name__ == '__main__':
    delete_output()
    check_files()
    print 'files are all valid'
    import time
    start = time.time()
    #asdf
    root = 'D:/project13/input/data/'
    flist_root = os.listdir(root)
    fname1_list = []
    fname2_list = []
    for i in flist_root:
        flist_dir = os.listdir(root+i)
        for j in flist_dir:
            # print(j.decode('gbk'))
            # print(j.decode('gbk'))
            # print((j.decode('gbk')).encode('gbk'))
            # print '����������ܱ��Ĺ�ϵ.xls'.decode('gbk').encode('gbk')
            # print(j.decode('gbk') == '����������ܱ��Ĺ�ϵ.xls'.decode('gbk'))
            if '̨��' not in j and 'CAD' not in j and '������' not in j and (j.endswith('xls') or j.endswith('xlsx')):
                fname1_list.append(root+i+'/'+j)


            if j.decode('gbk') == '����������ܱ��Ĺ�ϵ.xls'.decode('gbk') or j.decode('gbk') == '����������ܱ��Ĺ�ϵ.xlsx'.decode('gbk'):
                # print(j.decode('gbk'),1321321)
                fname2_list.append(root + i + '/' + j)
    # print len(fname1_list)
    # print len(fname2_list)
    # for i in range(len(fname1_list)):
    #     print fname1_list[i].decode('gbk')
    #     print fname2_list[i].decode('gbk')
    #     print '-----------------------------'


    j = 0
    mapping_break_flag = 0
    total = len(fname1_list)
    for i in range(len(fname1_list)):
        j +=1
        print j,'/',total
        fname1 = fname1_list[i]
        print fname1.decode('gbk')

        fname2 = fname2_list[i]
        print fname2.decode('gbk')
        print '---------------------'

        #��������ļ��У����Ч�ʱ����ظ�����
        try:
            files_number = [44,41,42,32,26,35]
            check_files_list = os.listdir('d:\\project13\\output\\'+flist_root[i])
            if len(check_files_list) in files_number:
                print flist_root[i].decode('gbk'),'�Ѿ�����'.decode('gbk')
                continue
        except:
            pass


        # debug
        # if fname1 == 'D:/project13/input/data/��˰���ſ�/��˰���ſ�(�ѽ�ģ).xls'\
        #     or fname1 == 'D:/project13/input/data/����ڱ�/����ڱ�.xls'\
        #     or fname1 == 'D:/project13/input/data/�ǽ�һ���ſ�/�ǽ�һ���ſ�̨��.xls'\
        #     or fname1 == 'D:/project13/input/data/����101/GIS.xls'\
        #     or fname1 == 'D:/project13/input/data/�Ų˴��ۺ�1#/GIS.xls'\
        #     or fname1 == 'D:/project13/input/data/�߸�103/�߸�103.xls':
        #
        #     continue
        try:
            main(fname1.decode('gbk'),fname2.decode('gbk'),fname3.decode('gbk'),fname4.decode('gbk'),flist_root[i])
        except Exception as e:
            print(e)





        # try:
        #     main(fname1.decode('gbk'),fname2.decode('gbk'),fname3.decode('gbk'),fname4.decode('gbk'),flist_root[i])
        #     if fname1.split('/')[4] != fname2.split('/')[4]:
        #         print fname1.decode('gbk').split('/')[4]
        #         print 'error'
        #         print fname2.decode('gbk').split('/')[4]
        #         break_flag = 1
        #         break
        # except Exception,e:
        #     print Exception,e
        #     print 'error','ͼ�����ɴ���'.decode('gbk')


    # if mapping_break_flag == 0:
    os.system('C:\\Python27_for_arcgis\\ArcGIS10.2\\python.exe D:\\project13\\bin\\arcpy_mapping_ver3.py')
    end = time.time()
    print 'running duration',round((end-start),2)