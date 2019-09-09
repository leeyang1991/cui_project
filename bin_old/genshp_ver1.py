# coding=gbk
import os
import read_excel_ver1
import write_point as wp
import write_line as wl
import time

#1生成变压器shp点
def bianyaqi_shp(r,f_dir):
    bianyaqi_lon, bianyaqi_lat, bianyaqi_label = r.didianyapeidianxiang()
    bianyaqi_label = bianyaqi_label
    directory = 'D:/project13/output/data/'+f_dir
    if not os.path.isdir(directory):
        os.mkdir(directory)
    fname = directory+'\\bianyaqi.shp'
    # print type(bianyaqi_lon)
    wp.point_to_shp([[bianyaqi_lon,bianyaqi_lat,str(bianyaqi_lon),str(bianyaqi_lat),bianyaqi_label,'','']],fname)
    # wp.write_shp_attrib(fname,bianyaqi_label)
    # wp.write_shp_attrib(fname,str(bianyaqi_lon))
    # wp.write_shp_attrib(fname,str(bianyaqi_lat))
# bianyaqi_shp()

#2生成电线杆shp点
def dianxiangan_dian_shp(r,f_dir):
    dianxiangan_lon, dianxiangan_lat, dianxiangan_name, dianxiangan_qianduan = r.dianxiangan()
    directory = 'D:\\project13\\output\\data\\'+f_dir
    dianxiangan_coor_dic = {}
    for i in range(len(dianxiangan_lon)):
        dianxiangan_coor_dic[i] = [dianxiangan_lon[i],dianxiangan_lat[i],dianxiangan_qianduan[i],dianxiangan_name[i]]
    if not os.path.isdir(directory):
        os.mkdir(directory)
    fname = directory+'\\dianxiangan'
    # lonlist=[];latlist=[];att1=[];att2=[];att3=[];att4=[];att5=[]
    point=[]
    for i in dianxiangan_coor_dic:
        point.append([dianxiangan_coor_dic[i][0],dianxiangan_coor_dic[i][1],dianxiangan_coor_dic[i][2],dianxiangan_coor_dic[i][3],'','',''])
        # # print type(dianxiangan_coor_dic[i][2])
        # lonlist.append(dianxiangan_coor_dic[i][0])
        # # print dianxiangan_coor_dic[i][0]
        # latlist.append(dianxiangan_coor_dic[i][1])
        # # print dianxiangan_coor_dic[i][1]
        # # att1.append(dianxiangan_coor_dic[i][2])
        # att1.append('')
        # # print dianxiangan_coor_dic[i][2]
        # att2.append(dianxiangan_coor_dic[i][3])
        # # print dianxiangan_coor_dic[i][3]
        # att3.append('')
        # att4.append('')
        # att5.append('')
    # print att1
    # print point[0]
    wp.point_to_shp(point,fname+'.shp')
        # print dianxiangan_coor_dic[i][2],dianxiangan_coor_dic[i][3]
    # lon_str=[]
    # lat_str=[]
    # for i in range(len(dianxiangan_lon)):
    #     lon_str.append(str(dianxiangan_lon[i]))
    #     lat_str.append(str(dianxiangan_lat[i]))

    # wp.merge_point(directory+'..\\dianxiangan_merge.shp',directory)
    # wp.write_shp_attrib(directory+'..\\dianxiangan_merge.shp',dianxiangan_name)
    # wp.write_shp_attrib(directory+'..\\dianxiangan_merge.shp',lon_str)
    # wp.write_shp_attrib(directory+'..\\dianxiangan_merge.shp',lat_str)

#3生成墙支架shp点
def qiangzhijia_shp(r,f_dir):
    lon_list,lat_list,qianduan,_ = r.diyaqiangzhijia()
    directory = 'D:\\project13\\output\\data\\'+f_dir
    if not os.path.isdir(directory):
        os.mkdir(directory)
    qiangzhijia_coor_dic={}
    for i in range(len(lon_list)):
        qiangzhijia_coor_dic[i] = [lon_list[i],lat_list[i],qianduan[i]]
    fname = directory+'\\qiangzhijia'


    # lonlist=[];latlist=[];att1=[];att2=[];att3=[];att4=[];att5=[]
    point=[]
    for i in qiangzhijia_coor_dic:
        # lonlist.append(qiangzhijia_coor_dic[i][0])
        # latlist.append(qiangzhijia_coor_dic[i][1])
        # att1.append(qiangzhijia_coor_dic[i][2])
        # att2.append('')
        # att3.append('')
        # att4.append('')
        # att5.append('')
        point.append([qiangzhijia_coor_dic[i][0],qiangzhijia_coor_dic[i][1],qiangzhijia_coor_dic[i][2],'','','',''])

    wp.point_to_shp(point,fname+'.shp')


#4生成计量箱shp点
#4.1选择出黑色和红色计量箱
def select_red_black_points(r):
    dianbiao_dic, spare_list = r.select_free_jiliangxiang()
    lon_list,lat_list,hang_list,lie_list,code_list = r.jiliangxiang()

    red_lon=[];red_lat=[];red_hang_lie=[];red_code=[]
    black_lon=[];black_lat=[];black_hang_lie=[];black_code=[]
    for i in dianbiao_dic:
        if i in spare_list:
            # print i,'red'
            index = code_list.index(i)
            red_lon.append(lon_list[index])
            red_lat.append(lat_list[index])
            red_hang_lie.append(str(hang_list[index])+','+str(lie_list[index]))
            red_code.append(i)
        else:
            try:
                index = code_list.index(i)
                black_lon.append(lon_list[index])
                black_lat.append(lat_list[index])
                black_hang_lie.append(str(hang_list[index])+','+str(lie_list[index]))
                black_code.append(i)
            except:
                pass
    return red_lon,red_lat,red_hang_lie,red_code,black_lon,black_lat,black_hang_lie,black_code

#4.2生成黑色计量箱shp点
def black_jiliangxiang(r,f_dir):
    _,_,_,_,black_lon,black_lat,black_hang_lie,black_code = select_red_black_points(r)
    # print black_lon
    # print black_lat
    # print black_hang_lie
    directory = 'D:\\project13\\output\\data\\'+f_dir
    if not os.path.isdir(directory):
        os.mkdir(directory)
    fname = directory+'\\black_jiliangxiang'
    lon_str=[]
    lat_str=[]
    for i in range(len(black_lon)):
        lon_str.append(str(black_lon[i]))
        lat_str.append(str(black_lat[i]))


        # print 'black_code_type',type(str(black_code[i]))
        # print 'int',str(int(black_code[i]))

    # lonlist=[];latlist=[];att1=[];att2=[];att3=[];att4=[];att5=[]
    lonlist = black_lon
    latlist = black_lat
    # att1 = lon_str
    # att2 = lat_str
    att3 = black_hang_lie
    point=[]
    for i in range(len(lonlist)):
        point.append([lonlist[i],latlist[i],'','',black_hang_lie[i],'',''])

    wp.point_to_shp(point,fname+'.shp')
    # wp.merge_point(directory+'..\\black_jiliangxiang_merge.shp',directory)
    # wp.write_shp_attrib(directory+'..\\black_jiliangxiang_merge.shp',black_hang_lie)
    # wp.write_shp_attrib(directory+'..\\black_jiliangxiang_merge.shp',black_code)
    # wp.write_shp_attrib(directory+'..\\black_jiliangxiang_merge.shp',lon_str)
    # wp.write_shp_attrib(directory+'..\\black_jiliangxiang_merge.shp',lat_str)


#4.3生成红色计量箱shp点
def red_jiliangxiang(r,f_dir):
    red_lon,red_lat,red_hang_lie,red_code,_,_,_,_ = select_red_black_points(r)
    directory = 'D:\\project13\\output\\data\\'+f_dir
    if not os.path.isdir(directory):
        os.mkdir(directory)
    fname = directory+'\\red_jiliangxiang'
    lon_str=[]
    lat_str=[]
    for i in range(len(red_lon)):
        lon_str.append(str(red_lon[i]))
        lat_str.append(str(red_lat[i]))

    # j=0
    # for i in range(len(red_lon)):
    #     j+=1
    #     wp.point_to_shp(red_lon[i],red_lat[i],fname+str(j)+'.shp',str(red_lon[i]),str(red_lat[i]),str(red_hang_lie[i]),'','')

    lonlist=[];latlist=[];att1=[];att2=[];att3=[];att4=[];att5=[]
    lonlist = red_lon
    latlist = red_lat
    att1 = lon_str
    att2 = lat_str
    att3 = red_hang_lie
    point=[]
    for i in range(len(lonlist)):
        point.append([red_lon[i],red_lat[i],'','',red_hang_lie[i],'',''])
        # att4.append('')
        # att5.append('')
    wp.point_to_shp(point,fname+'.shp')

    # wp.merge_point(directory+'..\\red_jiliangxiang_merge.shp',directory)
    # wp.write_shp_attrib(directory+'..\\red_jiliangxiang_merge.shp',red_hang_lie)
    # wp.write_shp_attrib(directory+'..\\red_jiliangxiang_merge.shp',red_code)
    # wp.write_shp_attrib(directory+'..\\red_jiliangxiang_merge.shp',lon_str)
    # wp.write_shp_attrib(directory+'..\\red_jiliangxiang_merge.shp',lat_str)


#4.4获取透明计量箱显示名字
def transparent_jiliangxiang(r):
    #error
    dianbiao_dic, _ = r.select_free_jiliangxiang()
    lon_list,lat_list,hang_list,lie_list,code_list= r.jiliangxiang()
    jiliangxiang_lon=[];jiliangxiang_lat=[];name_list=[];code_list1=[]
    for i in dianbiao_dic:
        try:
            index = code_list.index(i)
            jiliangxiang_lon.append(lon_list[index])
            jiliangxiang_lat.append(lat_list[index])
            name_list.append(dianbiao_dic[i])
            code_list1.append(i)
        except:
            pass
        #
        # print i
        # print lon_list[index]
        # print lat_list[index]
        # for j in dianbiao_dic[i]:
        #     print j
        # print '----------------'
        # code_list1.append(i)
    trans={}
    for i in range(len(code_list1)):
        trans[code_list1[i]] = [jiliangxiang_lon[i],jiliangxiang_lat[i],name_list[i]]

    return trans


#4.5生成透明计量箱shp点
def transparent_jiliangxiang_shp(r,f_dir):

    transparent = transparent_jiliangxiang(r)
    # print transparent['4130005000000054522636'][0]
    # print transparent
    directory = 'D:\\project13\\output\\data\\'+f_dir
    if not os.path.isdir(directory):
        os.mkdir(directory)
    fname = directory+'\\transparent_jiliangxiang'
    j=0
    # name_list=[]
    # code_list=[]
    # lon_str=[]
    # lat_str=[]
    # for i in transparent:
    #     j+=1
    #     lon_str.append(str(transparent[i][0]))
    #     lat_str.append(str(transparent[i][1]))
    #     str_name = ' '.join(transparent[i][2])
    #     # print str_name
    #     wp.point_to_shp(transparent[i][0],transparent[i][1],fname+str(j)+'.shp',str(transparent[i][0]),str(transparent[i][1]),str_name,'','')
    #     name_list.append(transparent[i][2])
    #     code_list.append(i)

    # lonlist=[];latlist=[];att1=[];att2=[];att3=[];att4=[];att5=[]
    point = []
    for i in transparent:
        point.append([transparent[i][0],transparent[i][1],' '.join(transparent[i][2]),'','','',''])
        # lonlist.append(transparent[i][0])
        # latlist.append(transparent[i][1])
        # att1.append(transparent[i][2])
        # att2.append('')
        # att3.append('')
        # att4.append('')
        # att5.append('')

    wp.point_to_shp(point,fname+'.shp')
    # name_list1=[]
    # for i in name_list:
    #     name_list1.append(' '.join(i))
    # wp.merge_point(directory+'..\\transparent_jiliangxiang_merge.shp',directory)
    # wp.write_shp_attrib(directory+'..\\transparent_jiliangxiang_merge.shp',name_list1)
    # wp.write_shp_attrib(directory+'..\\transparent_jiliangxiang_merge.shp',code_list)
    # wp.write_shp_attrib(directory+'..\\transparent_jiliangxiang_merge.shp',lon_str)
    # wp.write_shp_attrib(directory+'..\\transparent_jiliangxiang_merge.shp',lat_str)


#5连接电线杆
#5.1生成电线杆连线两点坐标
def dianxiangan_line(r):
    lon_list,lat_list,name_list,qianduan_list = r.dianxiangan()
    dianxiangan_coor_dic = {}
    for i in range(len(lon_list)):
        dianxiangan_coor_dic[i] = lon_list[i],lat_list[i],name_list[i],qianduan_list[i]
    p1_list=[]
    p2_list=[]

    for i in dianxiangan_coor_dic:
        if '配电箱'.decode('gbk') in dianxiangan_coor_dic[i][3]:
            # print dianxiangan_coor_dic[i][3]
            bianyaqi_lon, bianyaqi_lat, bianyaqi_label = r.didianyapeidianxiang()
            # print 'peidianxiang',bianyaqi_lon, bianyaqi_lat
            p2_list.append((bianyaqi_lon, bianyaqi_lat))
        elif '墙支架'.decode('gbk') in dianxiangan_coor_dic[i][3]:
            lon_list,lat_list,qianduan,qiangzhijia_name = r.diyaqiangzhijia()
            qiangzhijia_coor_dic={}
            for j in range(len(lon_list)):
                qiangzhijia_coor_dic[j] = [lon_list[j],lat_list[j],qiangzhijia_name[j]]
            for k in qiangzhijia_coor_dic:
                if dianxiangan_coor_dic[i][3] == qiangzhijia_coor_dic[k][2]:
                    p2_list.append((qiangzhijia_coor_dic[k][0],qiangzhijia_coor_dic[k][1]))



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
            print '###',i,'###'
    distancestr = []
    distance = []
    for i in range(len(p1_list)):

        distancestr.append(str(round(wl.GetDistance(p1_list[i][0],p1_list[i][1],p2_list[i][0],p2_list[i][1]),2)))
        distance.append(wl.GetDistance(p1_list[i][0],p1_list[i][1],p2_list[i][0],p2_list[i][1]))
    total_distance = round(sum(distance),2)
    # print distance
    # print distance2
    # for i in range(len(distance)):
    #     print distance[i] - distance2[i]
    # return p1_list,p2_list



    # for i in range(len(lon_list)):
    #     dianxiangan_coor_dic[name_list[i]] = [lon_list[i],lat_list[i]]
    #
    # p1_list=[];p2_list=[]
    # # for i in dianxiangan_coor_dic:
    # #     for j in qianduan_list:
    # #         if dianxiangan_coor_dic[i][2] == j:
    # #             try:
    # #                 p1_list.append([dianxiangan_coor_dic[i][0],dianxiangan_coor_dic[i][1]])
    # #                 p2_list.append([dianxiangan_coor_dic[j][0],dianxiangan_coor_dic[j][1]])
    # #
    # #                 # print '\n'
    # #             except:
    # #                 pass
    # key1 = name_list
    # key2 = qianduan_list
    # # print key1
    # for i in key1:
    #     p1_list.append(dianxiangan_coor_dic[i])
    # # print p1_list
    # # print len(p1_list)
    # for i in key2:
    #     try:
    #         p2_list.append(dianxiangan_coor_dic[i])
    #     except:
    #         p2_list.append([])
    # # print p2_list
    # # print len(p2_list)
    # # print p1_list
    # # print p2_list


    return p1_list,p2_list,distancestr,total_distance
#5.2生成电线杆线shp
def dianxiangan_line_shp(r,f_dir):
    p1_list,p2_list,distance,total_distance = dianxiangan_line(r)

    directory = 'D:\\project13\\output\\data\\'+f_dir
    if not os.path.isdir(directory):
        os.mkdir(directory)
    fname = directory+'\\dianxiangan_line'

    lines = []
    for i in range(len(p1_list)):
        lines.append([p1_list[i],p2_list[i],distance[i],'','','',''])
    wl.line_to_shp(lines,fname+'.shp')

    # for i in range(len(p1_list)):
    #     j+=1
    #     # print p1_list[i][0],p1_list[i][1],p2_list[i][0],p2_list[i][1]
    #     # if i+1 > len(p1_list):
    #     #     break
    #     try:
    #         distance = wl.haversine(p1_list[i+1][0],p1_list[i+1][1],p2_list[i+1][0],p1_list[i+1][1])
    #         # print distance
    #         # print p1_list[i+1][0],p1_list[i+1][1],p2_list[i+1][0],p1_list[i+1][1]
    #         wl.line_to_shp(p1_list[i],p2_list[i],fname+str(j)+'.shp',str(distance),'a','a')
    #
    #     except:
    #         distance = wl.haversine(p1_list[i][0],p1_list[i][1],p2_list[i][0],p1_list[i][1])
    #         wl.line_to_shp(p1_list[i],p2_list[i],fname+str(j)+'.shp',str(distance),'a','a')

#6连接墙支架
#6.1生成墙支架与电线杆连线两点坐标
def qiangzhijia_line(r):
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

#6.2画墙支架与电线杆连线shp
def qiangzhijia_line_shp(r,f_dir):
    p1_list,p2_list,total_distance,distancestr = qiangzhijia_line(r)
    directory = 'D:\\project13\\output\\data\\'+f_dir
    if not os.path.isdir(directory):
        os.mkdir(directory)
    fname = directory+'\\qiangzhijia_line.shp'
    lines = []
    for i in range(len(p1_list)):
        lines.append([p1_list[i],p2_list[i],distancestr[i],'','','',''])
    wl.line_to_shp(lines,fname)



    # j=0
    # for i in range(len(p1_list)):
    #     j+=1
    #     try:
    #         distance = wl.haversine(p1_list[i+1][0],p1_list[i+1][1],p2_list[i+1][0],p1_list[i+1][1])
    #         # print distance
    #         # print p1_list[i+1][0],p1_list[i+1][1],p2_list[i+1][0],p1_list[i+1][1]
    #         wl.line_to_shp(p1_list[i],p2_list[i],fname+str(j)+'.shp',str(distance),'a','a')
    #
    #     except:
    #         distance = wl.haversine(p1_list[i][0],p1_list[i][1],p2_list[i][0],p1_list[i][1])
    #         wl.line_to_shp(p1_list[i],p2_list[i],fname+str(j)+'.shp',str(distance),'a','a')

#6.2生成墙支架之间连线两点坐标
def qiangzhijia_line2(r):
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

#6.3生成墙支架之间连线shp
def qiangzhijia_line2_shp(r,f_dir):
    p1_list,p2_list,_,distancestr = qiangzhijia_line2(r)
    # print p1_list
    # print p2_list
    directory = 'D:\\project13\\output\\data\\'+f_dir
    if not os.path.isdir(directory):
        os.mkdir(directory)
    fname = directory+'\\qiangzhijia_line2.shp'
    lines = []
    for i in range(len(p1_list)):
        lines.append([p1_list[i],p2_list[i],distancestr[i],'','','',''])
    wl.line_to_shp(lines,fname)


    # for i in range(len(p1_list)):
    #     j+=1
    #     try:
    #         distance = wl.haversine(p1_list[i+1][0],p1_list[i+1][1],p2_list[i+1][0],p1_list[i+1][1])
    #         # print distance
    #         # print p1_list[i+1][0],p1_list[i+1][1],p2_list[i+1][0],p1_list[i+1][1]
    #         wl.line_to_shp(p1_list[i],p2_list[i],fname+'000'+str(j)+'.shp',str(distance),'a','a')
    #     except:
    #         distance = wl.haversine(p1_list[i][0],p1_list[i][1],p2_list[i][0],p1_list[i][1])
    #         wl.line_to_shp(p1_list[i],p2_list[i],fname+'000'+str(j)+'.shp',str(distance),'a','a')

#7后处理
#7.1判断出图横竖画幅
def select_vertical_horizontal(r,f_dir):
    transparent = transparent_jiliangxiang(r)
    lon_list,lat_list,name_list,qianduan_list = r.dianxiangan()
    directory = 'D:\\project13\\output\\data\\'+f_dir
    if not os.path.isdir(directory):
        os.mkdir(directory)
    fname = directory+'\\extent_lyr.shp'
    lon=[]
    lat=[]
    for i in transparent:
        lon.append(transparent[i][0])
        lat.append(transparent[i][1])
    for i in range(len(lon_list)):
        lon.append(lon_list[i])
        lat.append(lat_list[i])
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


#7.2生成图形信息文本
def gen_text_info(r,f_dir,sheet_name):
    #计量箱表位和安装数量
    total, install = r.count_free_jiliangxiang()
    #线路长度
    p1_list,p2_list,distancestr,distance1 = dianxiangan_line(r)
    p1_list,p2_list,distance2,_ = qiangzhijia_line2(r)
    p1_list,p2_list,distance3,_ = qiangzhijia_line(r)
    total_distance = distance1 + distance2 + distance3
    #PMS_ID
    bianyaqi_code = r.legend_info()
    import xlrd

    legend_file = 'D:\\project13\\input图例信息\\图例信息.xlsx'
    bk = xlrd.open_workbook(legend_file)
    # sh = bk.sheet_by_name(sheet_name.decode('gbk'))
    sh = bk.sheet_by_index(0)
    nrow = sh.nrows
    PSM_ID_list = []
    taiqu_code_list = []
    taiqu_name_list = []
    bianyaqi_type_list = []
    bianyaqi_content_list = []
    prime_line_list=[]
    branch_line_list=[]
    for i in range(nrow):
        if i == nrow - 1:
            break
        PSM_ID_list.append(str(sh.cell_value(i+1,1)))
        taiqu_code_list.append(str(sh.cell_value(i+1,2)))
        taiqu_name_list.append(sh.cell_value(i+1,3))
        bianyaqi_type_list.append(str(sh.cell_value(i+1,4)))
        bianyaqi_content_list.append(str(int(sh.cell_value(i+1,5))))
        prime_line_list.append(sh.cell_value(i+1,6))
        branch_line_list.append(sh.cell_value(i+1,7))
        if i == nrow - 1:
            break
    # print prime_line_list
    # print branch_line_list
    legend_dic = {}
    for i in range(len(PSM_ID_list)):
        legend_dic[i] = [PSM_ID_list[i],taiqu_code_list[i],taiqu_name_list[i],bianyaqi_type_list[i],bianyaqi_content_list[i],prime_line_list[i],branch_line_list[i]]


    for i in legend_dic:
        if bianyaqi_code == legend_dic[i][0]:
            taiqu_code = legend_dic[i][1]
            taiqu_name = legend_dic[i][2]
            bianyaqi_type = legend_dic[i][3]
            bianyaqi_content = legend_dic[i][4]
            prime_line = legend_dic[i][5]
            branch_line = legend_dic[i][6]

            import codecs
            fw = codecs.open('D:\\project13\\output\\data\\'+f_dir+'\\info.txt','w','utf-8')
            fw.write(taiqu_code+',')
            fw.write(taiqu_name+',')
            # print 1111111111111111111
            fw.write(str(bianyaqi_type)+',')
            fw.write(str(bianyaqi_content)+',')
            fw.write(str(round(total_distance/1000,2))+' KM'+',')
            fw.write(unicode(prime_line)+',')
            fw.write(unicode(branch_line)+',')
            fw.write(str(total)+',')
            fw.write(str(install))

            fw.close()

        # else:
        #     taiqu_code = ''
        #     taiqu_name = ''
        #     bianyaqi_type = ''
        #     bianyaqi_content = ''





def main(fname1,fname2,output_dir):
    #创建输出文件夹
    if not os.path.isdir('d:\\project13\\output_pic\\'):
        os.mkdir('d:\\project13\\output_pic\\')
    if not os.path.isdir('d:\\project13\\output\\'):
        os.mkdir('d:\\project13\\output\\')


    for i in os.listdir('d:/project13/input'):
        if not os.path.isdir('d:/project13/output/'+i):
            os.mkdir('d:\\project13\\output\\'+i)
        for j in os.listdir('d:/project13/input/'+i):
            if not os.path.isdir('d:/project13/output/'+i+'/'+j):
                os.mkdir('d:/project13/output/'+i+'/'+j)
                print j.decode('gbk')



    r = read_excel_ver1.ReadExcel(fname1,fname2)
    bianyaqi_shp(r,output_dir)
    print '变压器绘制完成'.decode('gbk')
    dianxiangan_dian_shp(r,output_dir)
    print '电线杆绘制完成'.decode('gbk')
    qiangzhijia_shp(r,output_dir)
    print '墙支架绘制完成'.decode('gbk')
    black_jiliangxiang(r,output_dir)
    print '黑色计量箱绘制完成'.decode('gbk')
    red_jiliangxiang(r,output_dir)
    print '红色计量箱绘制完成'.decode('gbk')
    transparent_jiliangxiang_shp(r,output_dir)
    print '透明计量箱绘制完成'.decode('gbk')
    dianxiangan_line_shp(r,output_dir)
    print '电线杆连线完成'.decode('gbk')
    qiangzhijia_line_shp(r,output_dir)
    print '墙支架与电线杆连线完成'.decode('gbk')
    qiangzhijia_line2_shp(r,output_dir)
    print '墙支架与墙支架连线完成'.decode('gbk')
    select_vertical_horizontal(r,output_dir)
    print 'extent layer绘制完成'.decode('gbk')
    gen_text_info(r,output_dir,'Sheet1')
    print '信息文件生成完成'.decode('gbk')
    print '--------------------------------------------'
if __name__ == '__main__':
    import time
    start = time.time()
    #asdf
    root = 'D:/project13/input/data/'
    flist_root = os.listdir(root)
    fname1_list = []
    fname2_list = []
    # for i in flist_root:
    #     print i.decode('gbk')
    for i in flist_root:
        flist_dir = os.listdir(root+i)

        for j in flist_dir:
            if j.endswith('.xls') and not j.startswith('4'):
                fname1_list.append(root+i+'/'+j)
            elif j == '台账':
                taizhang_dir = os.listdir(root+i+'/'+j)
                for k in taizhang_dir:
                    if k == '计量箱与电能表的关系.xls':
                        fname2_list.append(root+i+'/'+j+'/'+k)


    #test




    #添加日志文件
    # if os.path.isfile('../log.txt'):
    #     f = open('../log.txt','r')
    #     lines = f.readlines()
    #     f.close()
    #
    # else:
    #     lines = []
    # #
    # f = open('../log.txt','w')
    # f.write(''.join(lines))
    # for i in range(len(fname1_list)):
    #     fname1 = fname1_list[i]
    #     print fname1.decode('gbk')
    #     fname2 = fname2_list[i]
    #     print fname2.decode('gbk')
    #     print '---------------------'
    #     try:
    #         main(fname1.decode('gbk'),fname2.decode('gbk'),flist_root[i])
    #         # f.write('done\t'+time.asctime(time.localtime(time.time()))+'\t'+fname1+'\n')
    #     except Exception,e:
    #         # raise Exception,e
    #         # f.write('error\t'+time.asctime(time.localtime(time.time()))+'\t'+fname1+'\t'+str(e)+'\n')
    #         print e
    #         for i in range(5):
    #             print 'error'
    # f.close()

    # pass

    #执行main
    j = 0
    for i in range(len(fname1_list)):
        j +=1
        print j
        fname1 = fname1_list[i]
        print fname1.decode('gbk')

        fname2 = fname2_list[i]
        print fname2.decode('gbk')
        print '---------------------'

        # debug

        # if fname1 == 'D:/project13/input/data/西黄庄#台区/西黄庄#台区.xls':
        #     main(fname1.decode('gbk'),fname2.decode('gbk'),flist_root[i])
        #     break*
        # else:
        #     continue
        # if fname1 == 'D:/project13/input/data/七里岗北/七里岗北.xls'\
        #         or fname1 == 'D:/project13/input/data/十里铺安置小区1/十里铺安置小区1.xls'\
        #         or fname1 == 'D:/project13/input/data/十里铺安置小区2/十里铺安置小区2.xls'\
        #         or fname1 == 'D:/project13/input/data/十里铺安置小区3/十里铺安置小区3.xls'\
        #         or fname1 == 'D:/project13/input/data/十里铺安置小区4/十里铺安置小区4.xls'\
        #         or fname1 == 'D:/project13/input/data/十里铺安置小区5/十里铺安置小区5.xls'\
        #         or fname1 == 'D:/project13/input/data/城郊一中门口/城郊一中门口台区.xls'\
        #         or fname1 == 'D:/project13/input/data/星城湾二期1台区/星城湾二期1台区.xls'\
        #         or fname1 == 'D:/project13/input/data/星城湾二期2台区/星城湾二期2台区.xls'\
        #         or fname1 == 'D:/project13/input/data/美景鸿城1/美景鸿城1台区.xls'\
        #         or fname1 == 'D:/project13/input/data/美景鸿城2/美景鸿城2台区.xls'\
        #         or fname1 == 'D:/project13/input/data/马头社区1/马头社区1.xls'\
        #         or fname1 == 'D:/project13/input/data/马庄台区/马庄台区.xls'\
        #         or fname1 == 'D:/project13/input/data/西黄庄#台区/西黄庄#台区.xls'\
        #         or fname1 == 'D:/project13/input/data/高阳309台区/高阳309台区.xls'\
        #         or fname1 == 'D:/project13/input/data/东王庄西/东王庄西.xls'\
        #         or fname1 == 'D:/project13/input/data/马房村501台区/马房501台区.xls'\
        #         or fname1 == 'D:/project13/input/data/大河湾201台区/大河湾201台区.xls'\
        #         or fname1 == 'D:/project13/input/data/小集203/小集203台区.xls'\
        #         or fname1 == 'D:/project13/input/data/崔林202台区/崔林202台区.xls'\
        #         or fname1 == 'D:/project13/input/data/潘楼/潘楼.xls'\
        #         or fname1 == 'D:/project13/input/data/东铁401台区/东铁401台区.xls'\
        #         or fname1 == 'D:/project13/input/data/南村403台区/南村403台区.xls'\
        #         or fname1 == 'D:/project13/input/data/小岗401台区/小岗401台区.xls'\
        #         or fname1 == 'D:/project13/input/data/阳固镇史马房台区/阳固镇史马房台区.xls':
        #
        #
        #     continue
        main(fname1.decode('gbk'),fname2.decode('gbk'),flist_root[i])


        # try:
        #     main(fname1.decode('gbk'),fname2.decode('gbk'),flist_root[i])
        # except Exception,e:
        #     print Exception,e
        #     print 'error'
        #     print fname1.decode('gbk')
        #     print fname2.decode('gbk')




        #     f.write('done\t'+time.asctime(time.localtime(time.time()))+'\t'+fname1+'\n')


    os.system('C:\\Python27_for_arcgis\\ArcGIS10.2\\python.exe D:\\project13\\bin\\arcpy_mapping_ver3.py')


    end = time.time()
    print 'running duration:',round((end-start),2)