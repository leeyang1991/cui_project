# coding=gbk
import os
import read_excel_ver3
import write_point as wp
import write_line as wl
import time
import sys
import xlrd

args = sys.argv


# fname3 = 'D:\\project13\\input\\城内.xlsx'
# fname4 = r'E:\cui\191007\1007第一批出图\1007第一批出图\图例.xlsx'
# this_root = r'E:\cui\191007\1007第一批出图\1007第一批出图\翠翠--板木\第一批\\'
# out_dir = r''

yonghujierudian = 0
mapping_break_flag = 0


#1生成变压器shp点
def bianyaqi_shp(r,f_dir):
    bianyaqi_lon, bianyaqi_lat, bianyaqi_label = r.didianyapeidianxiang()
    # print(bianyaqi_lon, bianyaqi_lat, bianyaqi_label)
    bianyaqi_label = bianyaqi_label
    # print(f_dir)
    # exit()
    directory = f_dir
    if not os.path.isdir(directory):
        os.mkdir(directory)
    fname = directory+'bianyaqi.shp'
    # fname = fname.encode('gbk')
    # print(directory)
    # exit()
    # print(bianyaqi_lon, bianyaqi_lat, bianyaqi_label)
    wp.point_to_shp([[bianyaqi_lon,bianyaqi_lat,str(bianyaqi_lon),str(bianyaqi_lat),bianyaqi_label,'','']],fname)
    # wp.point_to_shp([[bianyaqi_lon,bianyaqi_lat,'','','','','']],fname)

#2生成电线杆shp点
def dianxiangan_dian_shp(r,f_dir):
    if r.dianxiangan():
        dianxiangan_lon, dianxiangan_lat, dianxiangan_name, dianxiangan_qianduan,line_type_list, distance_str_list, ganta_type_list = r.dianxiangan()
        directory = f_dir
        dianxiangan_coor_dic = {}
        for i in range(len(dianxiangan_lon)):
            dianxiangan_coor_dic[i] = [dianxiangan_lon[i],dianxiangan_lat[i],dianxiangan_qianduan[i],dianxiangan_name[i],ganta_type_list[i]]
        if not os.path.isdir(directory):
            os.mkdir(directory)
        fname = directory+'\\dianxiangan'

        point=[]
        for i in dianxiangan_coor_dic:
            point.append([dianxiangan_coor_dic[i][0],dianxiangan_coor_dic[i][1],dianxiangan_coor_dic[i][3],dianxiangan_coor_dic[i][2],dianxiangan_coor_dic[i][4],'',''])
        wp.point_to_shp(point,fname+'.shp')


#3生成墙支架shp点
def qiangzhijia_shp(r,f_dir):
    if r.diyaqiangzhijia():
        # lon_list, lat_list, qianduan_list, name_list, types_list
        lon_list,lat_list,qianduan,name,types = r.diyaqiangzhijia()
        directory = f_dir
        if not os.path.isdir(directory):
            os.mkdir(directory)
        qiangzhijia_coor_dic={}
        for i in range(len(lon_list)):
            qiangzhijia_coor_dic[i] = [lon_list[i],lat_list[i],qianduan[i],name[i],types[i]]
        fname = directory+'\\qiangzhijia'

        point=[]
        for i in qiangzhijia_coor_dic:

            point.append([qiangzhijia_coor_dic[i][0],qiangzhijia_coor_dic[i][1],qiangzhijia_coor_dic[i][2],qiangzhijia_coor_dic[i][3],qiangzhijia_coor_dic[i][4],'',''])

        wp.point_to_shp(point,fname+'.shp')

#4生成红色黑色透明计量箱
def jiliangxiang_shp(r,f_dir):
    transparent, red, black = r.full_and_spare_and_transparent_jiliangxiang()
    directory = f_dir
    if not os.path.isdir(directory):
        os.mkdir(directory)
    fname_trans = directory+'\\transparent'
    fname_red = directory+'\\red'
    fname_black = directory+'\\black'

    point_trans = []
    point_red = []
    point_black = []


    for i in transparent:
        # 姓名换行
        new_name = ''
        # print '***'.join(transparent[i][2])
        break_point = int(len(transparent[i][2]) / 2)
        if break_point == 0:
            break_point = 1
        for ii,name in enumerate(transparent[i][2]):
            new_name += name + ' '
            if (ii + 1) % break_point == 0 and not ii == 0:
                new_name += '\n'
        # point_trans.append([transparent[56i][0],transparent[i][1],' '.join(transparent[i][2]),i,'','','',''])
        point_trans.append([transparent[i][0],transparent[i][1],new_name,i,'','','',''])
    # exit()
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


#5连接电线杆
#5.1生成电线杆连线两点坐标
def dianxiangan_line(r):
    # print(r)
    if r.dianxiangan():
        lon_list,lat_list,name_list,qianduan_list,line_type_list,distance_str_list, ganta_type_list = r.dianxiangan()
        dianxiangan_coor_dic = []
        for i in range(len(lon_list)):
            dianxiangan_coor_dic.append([lon_list[i],lat_list[i],name_list[i],qianduan_list[i],line_type_list[i]])
        p1_list=[]
        p2_list=[]

        for i in range(len(dianxiangan_coor_dic)):
            if '配电箱'.decode('gbk') in dianxiangan_coor_dic[i][3]:
                # print dianxiangan_coor_dic[i][3]
                bianyaqi_lon, bianyaqi_lat, bianyaqi_label = r.didianyapeidianxiang()
                # print 'peidianxiang',bianyaqi_lon, bianyaqi_lat
                p2_list.append((bianyaqi_lon, bianyaqi_lat))
            elif '墙支架'.decode('gbk') in dianxiangan_coor_dic[i][3]:
                lon_list,lat_list,qianduan,qiangzhijia_name,qiangzhijia_types= r.diyaqiangzhijia()
                qiangzhijia_coor_dic={}
                for j in range(len(lon_list)):
                    qiangzhijia_coor_dic[j] = [lon_list[j],lat_list[j],qiangzhijia_name[j]]
                for k in qiangzhijia_coor_dic:
                    if dianxiangan_coor_dic[i][3] == qiangzhijia_coor_dic[k][2]:
                        p2_list.append((qiangzhijia_coor_dic[k][0],qiangzhijia_coor_dic[k][1]))
            elif '分支箱'.decode('gbk') in dianxiangan_coor_dic[i][3]:
                fenzhixiang_coor_dic = r.fenzhiiang()
                for j in fenzhixiang_coor_dic:
                    if dianxiangan_coor_dic[i][3] == fenzhixiang_coor_dic[j][2]:
                        p2_list.append((fenzhixiang_coor_dic[j][0],fenzhixiang_coor_dic[j][1]))


            p1_list.append(dianxiangan_coor_dic[i][:2])
            for j in range(len(dianxiangan_coor_dic)):
                if dianxiangan_coor_dic[i][3] == dianxiangan_coor_dic[j][2]:
                    p2_list.append(dianxiangan_coor_dic[j][:2])

        temp_p1 = []
        temp_p2 = []

        for i in range(len(dianxiangan_coor_dic)):
            temp_p1.append(dianxiangan_coor_dic[i][2])
            temp_p2.append(dianxiangan_coor_dic[i][3])
        for i in temp_p2:
            if not i in temp_p1:
                print '########',i,'########'

        distancestr = []
        for i in range(len(p2_list)):
            distancestr.append(str(round(wl.GetDistance(p1_list[i][0],p1_list[i][1],p2_list[i][0],p2_list[i][1]),2)))

        total_distance = []
        new_distance_str = []
        for i in range(len(distancestr)):
            dist = distance_str_list[i]
            if len(dist) > 0:
                new_distance_str.append(dist)
                total_distance.append(float(dist))
            else:
                new_distance_str.append(distancestr[i])
                total_distance.append(float(distancestr[i]))

        total_distance = round(sum(total_distance),2)

        return p1_list,p2_list,new_distance_str,total_distance,line_type_list


#5.2生成电线杆线shp
def dianxiangan_line_shp(r,f_dir):
    if dianxiangan_line(r):
        p1_list,p2_list,distance,total_distance,line_type_list = dianxiangan_line(r)

        directory = f_dir
        if not os.path.isdir(directory):
            os.mkdir(directory)
        fname = directory+'\\dianxiangan_line'

        lines = []
        for i in range(len(p2_list)):
            lines.append([p1_list[i],p2_list[i],distance[i],line_type_list[i],'','',''])
        wl.line_to_shp(lines,fname+'.shp')


#6连接墙支架
#6.1生成墙支架与电线杆连线两点坐标
def qiangzhijia_line(r):
    if r.dianxiangan():
        dianxiangan_lon,dianxiangan_lat,dianxiangan_name,_,_,_ ,_= r.dianxiangan()
        qiangzhijia_lon,qiangzhijia_lat,qianduan_name,name_list,types_list = r.diyaqiangzhijia()
        dianxiangan_coor_dic = {}
        for i in range(len(dianxiangan_lon)):
            dianxiangan_coor_dic[i] = [dianxiangan_lon[i],dianxiangan_lat[i],dianxiangan_name[i]]
        qiangzhijia_coor_dic ={}
        for i in range(len(qiangzhijia_lon)):
            qiangzhijia_coor_dic[i] = [qiangzhijia_lon[i],qiangzhijia_lat[i],qianduan_name[i],types_list[i]]

        p1_list=[];p2_list=[]
        for i in qiangzhijia_coor_dic:
            for j in dianxiangan_coor_dic:
                if qiangzhijia_coor_dic[i][2] == dianxiangan_coor_dic[j][2]:
                    p1_list.append([qiangzhijia_coor_dic[i][0],qiangzhijia_coor_dic[i][1],qiangzhijia_coor_dic[i][3]])
                    p2_list.append([dianxiangan_coor_dic[j][0],dianxiangan_coor_dic[j][1]])

        distancestr = []
        distance = []
        type_str = []
        for i in range(len(p1_list)):

            distancestr.append(str(round(wl.GetDistance(p1_list[i][0],p1_list[i][1],p2_list[i][0],p2_list[i][1]),2)))
            distance.append(wl.GetDistance(p1_list[i][0],p1_list[i][1],p2_list[i][0],p2_list[i][1]))
            type_str.append(p1_list[i][2])
        total_distance = round(sum(distance),2)
        return p1_list,p2_list,total_distance,distancestr,type_str

#6.2画墙支架与电线杆连线shp
def qiangzhijia_line_shp(r,f_dir):
    if qiangzhijia_line(r):
        p1_list,p2_list,total_distance,distancestr,type_str = qiangzhijia_line(r)
        directory = f_dir
        if not os.path.isdir(directory):
            os.mkdir(directory)
        fname = directory+'\\qiangzhijia_line.shp'
        lines = []
        for i in range(len(p1_list)):
            lines.append([p1_list[i],p2_list[i],distancestr[i],type_str[i],'','',''])
        wl.line_to_shp(lines,fname)


#6.2生成墙支架之间连线两点坐标
def qiangzhijia_line2(r):
    if r.diyaqiangzhijia():
        qiangzhijia_lon,qiangzhijia_lat,qianduan_name,name,types_list= r.diyaqiangzhijia()
        qiangzhijia_coor_dic = {}
        for i in range(len(qiangzhijia_lon)):
            qiangzhijia_coor_dic[i] = [qiangzhijia_lon[i],qiangzhijia_lat[i],qianduan_name[i],name[i],types_list[i]]
        p1_list=[];p2_list=[]
        for i in qiangzhijia_coor_dic:
            for j in qiangzhijia_coor_dic:
                if qiangzhijia_coor_dic[i][2] == qiangzhijia_coor_dic[j][3]:
                    p1_list.append([qiangzhijia_coor_dic[i][0],qiangzhijia_coor_dic[i][1],qiangzhijia_coor_dic[i][4]])
                    p2_list.append([qiangzhijia_coor_dic[j][0],qiangzhijia_coor_dic[j][1]])

        distancestr = []
        distance = []
        types = []
        for i in range(len(p1_list)):
            types.append(p1_list[i][2])
            distancestr.append(str(round(wl.GetDistance(p1_list[i][0],p1_list[i][1],p2_list[i][0],p2_list[i][1]),2)))
            distance.append(wl.GetDistance(p1_list[i][0],p1_list[i][1],p2_list[i][0],p2_list[i][1]))
        total_distance = round(sum(distance),2)

        return p1_list,p2_list,total_distance,distancestr,types

#6.3生成墙支架之间连线shp
def qiangzhijia_line2_shp(r,f_dir):
    if qiangzhijia_line2(r):
        p1_list,p2_list,_,distancestr,types = qiangzhijia_line2(r)
        # print p1_list
        # print p2_list
        directory = f_dir
        if not os.path.isdir(directory):
            os.mkdir(directory)
        fname = directory+'\\qiangzhijia_line2.shp'
        lines = []
        for i in range(len(p1_list)):
            lines.append([p1_list[i],p2_list[i],distancestr[i],types[i],'','',''])
        wl.line_to_shp(lines,fname)


#7生成电缆线
def dianlanxian_line_shp(r,f_dir):
    if r.dianlan() == None:
        return None
    p1_list, p2_list = r.dianlan()
    directory = f_dir
    if not os.path.isdir(directory):
        os.mkdir(directory)
    fname = directory+'\\dianlan.shp'
    lines = []
    for i in range(len(p1_list)):
        distancestr = (str(round(wl.GetDistance(p1_list[i][0],p1_list[i][1],p2_list[i][0],p2_list[i][1]),2)))
        lines.append([p1_list[i],p2_list[i],distancestr,'','','',''])
    wl.line_to_shp(lines,fname)

def dianlanian_line_distance(r):
    if r.dianlan() == None:
        return 0
    p1_list, p2_list = r.dianlan()
    distance = []
    for i in range(len(p1_list)):
        distance.append(round(wl.GetDistance(p1_list[i][0],p1_list[i][1],p2_list[i][0],p2_list[i][1]),2))
    total_distance = sum(distance)
    return total_distance

#8生成分支箱
def fenzhixiang_points(r,f_dir):
    coor_dic = r.fenzhiiang()
    directory = f_dir
    if not os.path.isdir(directory):
        os.mkdir(directory)
    fname = directory+'\\fenzhixiang'
    point=[]
    if coor_dic:
        for i in coor_dic:
            point.append((coor_dic[i][0],coor_dic[i][1],coor_dic[i][2],'','','',''))
        # print point
        wp.point_to_shp(point,fname+'.shp')




#9后处理
#9.1判断出图横竖画幅
def select_vertical_horizontal(r,f_dir):
    # transparent = transparent_jiliangxiang(r)

    transparent, red, black = r.full_and_spare_and_transparent_jiliangxiang()
    dianbiao_dic = r.jiliangxiang()
    lon=[]
    lat=[]
    if r.dianxiangan():
        lon_list,lat_list,name_list,qianduan_list,_,_,_ = r.dianxiangan()
        for i in range(len(lon_list)):
            lon.append(lon_list[i])
            lat.append(lat_list[i])
    bianyaqi_lon,bianyaqi_lat,_ = r.didianyapeidianxiang()
    directory = f_dir
    if not os.path.isdir(directory):
        os.mkdir(directory)
    fname = directory+'\\extent_lyr.shp'


    for i in dianbiao_dic:
        lon.append(dianbiao_dic[i][0])
        lat.append(dianbiao_dic[i][1])
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

#9.2生成图形信息文本
def gen_text_info(r,f_dir):

    #计量箱表位和安装数量
    total = r.count_total_jiliangxiang()
    installed = r.count_intalled_jiliangxiang()

    #线路长度
    if dianxiangan_line(r):
        p1_list, p2_list, distancestr, distance1,_ = dianxiangan_line(r)
    else:
        distance1 = 0

    if qiangzhijia_line2(r):
        p1_list,p2_list,distance2,_,_ = qiangzhijia_line2(r)
    else:
        distance2 = 0

    if qiangzhijia_line(r):
        p1_list,p2_list,distance3,_,_ = qiangzhijia_line(r)
    else:
        distance3 = 0

    distance4 = dianlanian_line_distance(r)
    # total_distance = distance1 + distance2 + distance3 + distance4
    total_distance = distance1

    #PMS_ID
    info_dic = r.info()
    PMS_ID = r.legend_info()

    for i in info_dic:
        if PMS_ID == info_dic[i][0]:
            taiqu_code,taiqu_name,bianyaqi_type,bianyaqi_content,prime_line_type,branch_line_type = info_dic[i][1],info_dic[i][2],info_dic[i][3],info_dic[i][4],info_dic[i][5],info_dic[i][6]
            import codecs
            fw = codecs.open(f_dir+'\\info.txt','w','utf-8')
            content = '说明：台区编号：{}，台区名称：{}，变压器类型：{}，变压器容量：{}，导线总长度：{}，主线导线类型：{}，支线导线类型：{}，计量箱表位：{}，实际装表位数：{}'.decode('gbk').format(taiqu_code,taiqu_name,bianyaqi_type,bianyaqi_content,str(round(total_distance/1000,2))+' KM',prime_line_type,branch_line_type,total,installed)

            content_warp = ''
            for i in range(len(content)):
                content_warp += content[i]
                if i % 47 == 0 and i > 5:
                    content_warp += '\n'
            fw.write(content_warp)
            # fw.write(taiqu_code+',')
            # fw.write(taiqu_name+',')
            # fw.write(str(bianyaqi_type)+',')
            # fw.write(str(bianyaqi_content)+',')
            # fw.write(str(round(total_distance/1000,2))+' KM,')
            # fw.write(str(prime_line_type)+',')
            # fw.write(str(branch_line_type)+',')
            # fw.write(str(total)+',')
            # fw.write(str(installed))
            fw.close()

            fw_name = codecs.open(f_dir+'\\name.txt','w','utf-8')
            fw_name.write('{}台区低压沿布图'.decode('gbk').format(taiqu_name))
            fw_name.close()


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

        directory = f_dir
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

    directory = f_dir
    if not os.path.isdir(directory):
        os.mkdir(directory)
    fname = directory+'\\yonghujierudian_jiliangxiang_line'

    lines = []
    for i in range(len(p1_list)):
        lines.append([p1_list[i],p2_list[i],'','','','',''])
    wl.line_to_shp(lines,fname+'.shp')



def check_files(root):
    # root = this_root
    flist_root = os.listdir(root)
    fname1_list = []
    fname2_list = []
    for i in flist_root:
        # fname1_list
        flist_dir = os.listdir(root+i)
        fname2_list.append(root + i + '/台账' + '计量箱与电能表的关系.xls')
        for j in flist_dir:
            # print(j.decode('gbk'))
            # print(j.decode('gbk'))
            # print((j.decode('gbk')).encode('gbk'))
            # print '计量箱与电能表的关系.xls'.decode('gbk').encode('gbk')
            # print(j.decode('gbk') == '计量箱与电能表的关系.xls'.decode('gbk'))
            if '台账' not in j and 'CAD' not in j and '计量箱' not in j and (j.endswith('xls') or j.endswith('xlsx')):
                fname1_list.append(root+i+'/'+j)

            # print(j.decode('gbk'))

            # if j.decode('gbk') == '计量箱与电能表的关系.xls'.decode('gbk') or j.decode('gbk') == '计量箱与电能表的关系.xlsx'.decode('gbk'):
            #     # print(j.decode('gbk'),1321321)
            #     fname2_list.append(root + i + '/' + j)
            # else:

    # fname2_list = set(fname2_list)
    # print(fname1_list)
    # print(fname2_list)
    # for i in fname1_list:
    #     print(i.decode('gbk'))
    # for i in fname2_list:
    #     print(i.decode('gbk'))
    # exit()
    j = 0
        # exit()

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
        if fname1.split('/')[-2] != fname2.split('/')[-2]:
            print fname1.decode('gbk').split('/')[-2]
            print 'error'
            print fname2.decode('gbk').split('/')[-2]
            break_flag = 1
        if break_flag == 1:
            time.sleep(1)
            raise IOError('this directory is invalid')

def delete_output(fdir):
    # fdir = r'd:\\project13\\output\\'
    fdir = fdir+'\\shp\\'
    fdir = fdir.decode('gbk')
    # print(fdir)
    if os.path.isdir(fdir):
        for f in os.listdir(fdir):
            os.remove(fdir+f)
    # if os.path.isdir(fdir):
    #     folders = os.listdir(fdir)
    #     for folder in folders:
    #         file_list = os.listdir(fdir+folder)
    #         # for f in file_list:
    #         #     os.remove(fdir+folder+'\\'+f)
    #         os.removedirs(fdir+folder)



def main(fname1,fname2,fname4,output_dir,bk_additional):
    #创建输出文件夹
    # if not os.path.isdir('d:\\project13\\output_pic\\'):
    #     os.mkdir('d:\\project13\\output_pic\\')
    # if not os.path.isdir('d:\\project13\\output\\'):
    #     os.mkdir('d:\\project13\\output\\')
    #
    #
    # for i in os.listdir('d:/project13/input/data/'):
    #     if not os.path.isdir('d:/project13/output/'+i):
    #         os.mkdir('d:\\project13\\output\\'+i)
    #         # print i.decode('gbk')


    print(fname1)
    print(fname2)
    print(fname4)
    # bk_additional = xlrd
    r = read_excel_ver3.ReadExcel(fname1,fname2,fname4,bk_additional)
    # try:
    # print(output_dir)
    output_dir = output_dir.encode('gbk')
    # exit()
# try:
    bianyaqi_shp(r,output_dir)
    print '变压器绘制完成'.decode('gbk')

    dianxiangan_dian_shp(r,output_dir)
    print '电线杆绘制完成'.decode('gbk')

    qiangzhijia_shp(r,output_dir)
    print '墙支架绘制完成'.decode('gbk')

    jiliangxiang_shp(r,output_dir)
    print '计量箱绘制完成'.decode('gbk')

    dianxiangan_line_shp(r,output_dir)
    print '电线杆连线完成'.decode('gbk')

    qiangzhijia_line_shp(r,output_dir)
    print '墙支架与电线杆连线完成'.decode('gbk')

    qiangzhijia_line2_shp(r,output_dir)
    print '墙支架与墙支架连线完成'.decode('gbk')

    select_vertical_horizontal(r,output_dir)
    print 'extent layer绘制完成'.decode('gbk')

    gen_text_info(r,output_dir)
    print '信息文件生成完成'.decode('gbk')

    dianlanxian_line_shp(r,output_dir)
    print '电缆线连线完成'.decode('gbk')

    fenzhixiang_points(r,output_dir)
    print '分支箱绘制完成'.decode('gbk')


    if yonghujierudian:
        yonghujierudian_line_shp(r,output_dir)
        yonghujierudian_jiliangxiang_line_shp(r,output_dir)
        print '用户接入点'.decode('gbk')

    print '--------------------------------------------'




def gen_layer(root,fname4,additional=''):

    root = root+'/'
    if len(additional) > 0:
        bk_additional = xlrd.open_workbook(additional)
    else:
        bk_additional = None
    flist_root = os.listdir(root)
    fname1_list = []
    fname2_list = []
    for i in flist_root:
        print(root+'/'+i)

        delete_output(root+'\\'+i)
        # exit()
        flist_dir = os.listdir(root+i)
        fname2_list.append(root + i + '/台账/' + '计量箱与电能表的关系.xls')
        for j in flist_dir:
            # print(j.decode('gbk'))
            # print(j.decode('gbk'))
            # print((j.decode('gbk')).encode('gbk'))
            # print '计量箱与电能表的关系.xls'.decode('gbk').encode('gbk')
            # print(j.decode('gbk') == '计量箱与电能表的关系.xls'.decode('gbk'))
            if '台账' not in j and 'CAD' not in j and '计量箱' not in j and (j.endswith('xls') or j.endswith('xlsx')):
                fname1_list.append(root+i+'/'+j)

    j = 0

    total = len(fname1_list)
    # for i in fname1_list:
    #     print(i)
    # exit()
    for i in range(len(fname1_list)):
        j +=1
        print j,'/',total
        fname1 = fname1_list[i]
        print fname1.decode('gbk')

        fname2 = fname2_list[i]
        print fname2.decode('gbk')
        print '---------------------'

        #检查已做文件夹，提高效率避免重复计算
        # try:
        #     files_number = [44,41,42,32,26,35]
        #     check_files_list = os.listdir('d:\\project13\\output\\'+flist_root[i])
        #     if len(check_files_list) in files_number:
        #         print flist_root[i].decode('gbk'),'已经制作'.decode('gbk')
        #         continue
        # except:
        #     pass


        # debug
        # if fname1 == 'D:/project13/input/data/国税局门口/国税局门口(已建模).xls'\
        #     or fname1 == 'D:/project13/input/data/七里岗北/七里岗北.xls'\
        #     or fname1 == 'D:/project13/input/data/城郊一中门口/城郊一中门口台区.xls'\
        #     or fname1 == 'D:/project13/input/data/安桥101/GIS.xls'\
        #     or fname1 == 'D:/project13/input/data/张菜村综合1#/GIS.xls'\
        #     or fname1 == 'D:/project13/input/data/瓦岗103/瓦岗103.xls':

        out_dir = root+flist_root[i]+'\\shp\\'
        out_dir = out_dir.decode('gbk')
        # exit()
        print(out_dir)
        # exit()
        if not os.path.isdir(out_dir):
            os.mkdir(out_dir)
        # print(out_dir.decode('gbk'))
        # print(flist_root[i])
        # exit()
        main(fname1.decode('gbk'),fname2.decode('gbk'),fname4.decode('gbk'),out_dir,bk_additional)
        # exit()
        # except Exception as e:
        #     print(e)





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
        #     print 'error','图层生成错误'.decode('gbk')


    # if mapping_break_flag == 0:
    # os.system('C:\\Python27_for_arcgis\\ArcGIS10.2\\python.exe D:\\project13\\bin\\arcpy_mapping_ver3.py')
    # end = time.time()
    # print 'running duration',round((end-start),2)


if __name__ == '__main__':
    for i in args:
        print(i)
    exit()
    fname4 = r'E:\cui\191007\1007第一批出图\1007第一批出图\图例.xlsx'
    # gen_layer(this_root,fname4)