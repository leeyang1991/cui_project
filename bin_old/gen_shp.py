# coding=gbk
import os
import read_excel
import write_point as wp
import write_line as wl
import time

#1生成变压器shp点
def bianyaqi_shp(r,f_dir):
    bianyaqi_lon, bianyaqi_lat, bianyaqi_label = r.didianyapeidianxiang()
    bianyaqi_label = bianyaqi_label
    directory = 'D:/project13/output/葛岗所/'+f_dir+'/bianyaqi'
    if not os.path.isdir(directory):
        os.mkdir(directory)
    fname = directory+'\\bianyaqi.shp'



    wp.point_to_shp(bianyaqi_lon,bianyaqi_lat,fname,str(bianyaqi_lon),str(bianyaqi_lat),bianyaqi_label,'','')
    # wp.write_shp_attrib(fname,bianyaqi_label)
    # wp.write_shp_attrib(fname,str(bianyaqi_lon))
    # wp.write_shp_attrib(fname,str(bianyaqi_lat))
# bianyaqi_shp()

#2生成电线杆shp点
def dianxiangan_dian_shp(r,f_dir):
    dianxiangan_lon, dianxiangan_lat, dianxiangan_name, dianxiangan_qianduan = r.dianxiangan()
    directory = 'D:\\project13\\output\\葛岗所\\'+f_dir+'\\dianxiangan'
    dianxiangan_coor_dic = {}
    for i in range(len(dianxiangan_lon)):
        dianxiangan_coor_dic[i] = [dianxiangan_lon[i],dianxiangan_lat[i],dianxiangan_qianduan[i],dianxiangan_name[i]]
    if not os.path.isdir(directory):
        os.mkdir(directory)
    fname = directory+'\\dianxiangan'
    j=0
    for i in dianxiangan_coor_dic:
        j+=1
        wp.point_to_shp(dianxiangan_coor_dic[i][0],dianxiangan_coor_dic[i][1],fname+str(j)+'.shp',str(dianxiangan_coor_dic[i][0]),str(dianxiangan_coor_dic[i][1]),dianxiangan_coor_dic[i][3],dianxiangan_coor_dic[i][2],'')
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
    directory = 'D:\\project13\\output\\葛岗所\\'+f_dir+'\\qiangzhijia'
    if not os.path.isdir(directory):
        os.mkdir(directory)
    qiangzhijia_coor_dic={}
    for i in range(len(lon_list)):
        qiangzhijia_coor_dic[i] = [lon_list[i],lat_list[i],qianduan[i]]
    fname = directory+'\\qiangzhijia'
    j=0
    for i in qiangzhijia_coor_dic:
        j+=1
        wp.point_to_shp(qiangzhijia_coor_dic[i][0],qiangzhijia_coor_dic[i][1],fname+str(j)+'.shp',str(qiangzhijia_coor_dic[i][0]),str(qiangzhijia_coor_dic[i][1]),qiangzhijia_coor_dic[i][2],'','')


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
            index = code_list.index(i)
            black_lon.append(lon_list[index])
            black_lat.append(lat_list[index])
            black_hang_lie.append(str(hang_list[index])+','+str(lie_list[index]))
            black_code.append(i)
    return red_lon,red_lat,red_hang_lie,red_code,black_lon,black_lat,black_hang_lie,black_code

#4.2生成黑色计量箱shp点
def black_jiliangxiang(r,f_dir):
    _,_,_,_,black_lon,black_lat,black_hang_lie,black_code = select_red_black_points(r)
    # print black_lon
    # print black_lat
    # print black_hang_lie
    directory = 'D:\\project13\\output\\葛岗所\\'+f_dir+'\\black_jiliangxiang'
    if not os.path.isdir(directory):
        os.mkdir(directory)
    fname = directory+'\\black_jiliangxiang'
    lon_str=[]
    lat_str=[]
    for i in range(len(black_lon)):
        lon_str.append(str(black_lon[i]))
        lat_str.append(str(black_lat[i]))

    j=0
    for i in range(len(black_lon)):
        j+=1
        # print 'black_code_type',type(str(black_code[i]))
        # print 'int',str(int(black_code[i]))
        wp.point_to_shp(black_lon[i],black_lat[i],fname+str(j)+'.shp',str(black_lon[i]),str(black_lat[i]),str(black_hang_lie[i]),'','')
    # wp.merge_point(directory+'..\\black_jiliangxiang_merge.shp',directory)
    # wp.write_shp_attrib(directory+'..\\black_jiliangxiang_merge.shp',black_hang_lie)
    # wp.write_shp_attrib(directory+'..\\black_jiliangxiang_merge.shp',black_code)
    # wp.write_shp_attrib(directory+'..\\black_jiliangxiang_merge.shp',lon_str)
    # wp.write_shp_attrib(directory+'..\\black_jiliangxiang_merge.shp',lat_str)


#4.3生成红色计量箱shp点
def red_jiliangxiang(r,f_dir):
    red_lon,red_lat,red_hang_lie,red_code,_,_,_,_ = select_red_black_points(r)
    directory = 'D:\\project13\\output\\葛岗所\\'+f_dir+'\\red_jiliangxiang\\'
    if not os.path.isdir(directory):
        os.mkdir(directory)
    fname = directory+'red_jiliangxiang'
    lon_str=[]
    lat_str=[]
    for i in range(len(red_lon)):
        lon_str.append(str(red_lon[i]))
        lat_str.append(str(red_lat[i]))

    j=0
    for i in range(len(red_lon)):
        j+=1
        wp.point_to_shp(red_lon[i],red_lat[i],fname+str(j)+'.shp',str(red_lon[i]),str(red_lat[i]),str(red_hang_lie[i]),'','')
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
        index = code_list.index(i)
        jiliangxiang_lon.append(lon_list[index])
        jiliangxiang_lat.append(lat_list[index])
        name_list.append(dianbiao_dic[i])
        code_list1.append(i)
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
    directory = 'D:\\project13\\output\\葛岗所\\'+f_dir+'\\transparent_jiliangxiang\\'
    if not os.path.isdir(directory):
        os.mkdir(directory)
    fname = directory+'transparent_jiliangxiang'
    j=0
    name_list=[]
    code_list=[]
    lon_str=[]
    lat_str=[]
    for i in transparent:
        j+=1
        lon_str.append(str(transparent[i][0]))
        lat_str.append(str(transparent[i][1]))
        str_name = ' '.join(transparent[i][2])
        # print str_name
        wp.point_to_shp(transparent[i][0],transparent[i][1],fname+str(j)+'.shp',str(transparent[i][0]),str(transparent[i][1]),str_name,'','')
        name_list.append(transparent[i][2])
        code_list.append(i)
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
        dianxiangan_coor_dic[name_list[i]] = [lon_list[i],lat_list[i]]

    p1_list=[];p2_list=[]
    # for i in dianxiangan_coor_dic:
    #     for j in qianduan_list:
    #         if dianxiangan_coor_dic[i][2] == j:
    #             try:
    #                 p1_list.append([dianxiangan_coor_dic[i][0],dianxiangan_coor_dic[i][1]])
    #                 p2_list.append([dianxiangan_coor_dic[j][0],dianxiangan_coor_dic[j][1]])
    #
    #                 # print '\n'
    #             except:
    #                 pass
    key1 = name_list
    key2 = qianduan_list
    # print key1
    for i in key1:
        p1_list.append(dianxiangan_coor_dic[i])
    # print p1_list
    # print len(p1_list)
    for i in key2:
        try:
            p2_list.append(dianxiangan_coor_dic[i])
        except:
            p2_list.append([])
    # print p2_list
    # print len(p2_list)
    # print p1_list
    # print p2_list


    return p1_list,p2_list
#5.2生成电线杆线shp
def dianxiangan_line_shp(r,f_dir):
    p1_list,p2_list = dianxiangan_line(r)
    directory = 'D:\\project13\\output\\葛岗所\\'+f_dir+'\\dianxiangan_line\\'
    if not os.path.isdir(directory):
        os.mkdir(directory)
    fname = directory+'dianxiangan_line'
    j=0
    for i in range(len(p1_list)):
        j+=1
        # print p1_list[i][0],p1_list[i][1],p2_list[i][0],p2_list[i][1]
        # if i+1 > len(p1_list):
        #     break
        try:
            distance = wl.haversine(p1_list[i+1][0],p1_list[i+1][1],p2_list[i+1][0],p1_list[i+1][1])
            # print distance
            # print p1_list[i+1][0],p1_list[i+1][1],p2_list[i+1][0],p1_list[i+1][1]
            wl.line_to_shp(p1_list[i],p2_list[i],fname+str(j)+'.shp',str(distance),'a','a')

        except:
            distance = wl.haversine(p1_list[i][0],p1_list[i][1],p2_list[i][0],p1_list[i][1])
            wl.line_to_shp(p1_list[i],p2_list[i],fname+str(j)+'.shp',str(distance),'a','a')

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
    return p1_list,p2_list

#6.2画墙支架与电线杆连线shp
def qiangzhijia_line_shp(r,f_dir):
    p1_list,p2_list = qiangzhijia_line(r)
    directory = 'D:\\project13\\output\\葛岗所\\'+f_dir+'\\qiangzhijia_line\\'
    if not os.path.isdir(directory):
        os.mkdir(directory)
    fname = directory+'qiangzhijia_line'
    j=0
    for i in range(len(p1_list)):
        j+=1
        try:
            distance = wl.haversine(p1_list[i+1][0],p1_list[i+1][1],p2_list[i+1][0],p1_list[i+1][1])
            # print distance
            # print p1_list[i+1][0],p1_list[i+1][1],p2_list[i+1][0],p1_list[i+1][1]
            wl.line_to_shp(p1_list[i],p2_list[i],fname+str(j)+'.shp',str(distance),'a','a')

        except:
            distance = wl.haversine(p1_list[i][0],p1_list[i][1],p2_list[i][0],p1_list[i][1])
            wl.line_to_shp(p1_list[i],p2_list[i],fname+str(j)+'.shp',str(distance),'a','a')

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
    return p1_list,p2_list

#6.3生成墙支架之间连线shp
def qiangzhijia_line2_shp(r,f_dir):
    p1_list,p2_list = qiangzhijia_line2(r)
    # print p1_list
    # print p2_list
    directory = 'D:\\project13\\output\\葛岗所\\'+f_dir+'\\qiangzhijia_line\\'
    if not os.path.isdir(directory):
        os.mkdir(directory)
    fname = directory+'qiangzhijia_line'
    j=0
    for i in range(len(p1_list)):
        j+=1
        try:
            distance = wl.haversine(p1_list[i+1][0],p1_list[i+1][1],p2_list[i+1][0],p1_list[i+1][1])
            # print distance
            # print p1_list[i+1][0],p1_list[i+1][1],p2_list[i+1][0],p1_list[i+1][1]
            wl.line_to_shp(p1_list[i],p2_list[i],fname+'000'+str(j)+'.shp',str(distance),'a','a')
        except:
            distance = wl.haversine(p1_list[i][0],p1_list[i][1],p2_list[i][0],p1_list[i][1])
            wl.line_to_shp(p1_list[i],p2_list[i],fname+'000'+str(j)+'.shp',str(distance),'a','a')

def main(fname1,fname2,output_dir):
    #创建输出文件夹
    for i in os.listdir('d:/project13/数据1'):
        if not os.path.isdir('d:/project13/output/'+i):
            os.mkdir('d:\\project13\\output\\'+i)
        for j in os.listdir('d:/project13/数据1/'+i):
            if not os.path.isdir('d:/project13/output/'+i+'/'+j):
                os.mkdir('d:/project13/output/'+i+'/'+j)
                print j.decode('gbk')



    r = read_excel.ReadExcel(fname1,fname2)
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
    print '--------------------------------------------'
if __name__ == '__main__':
    # main()
    root = 'D:/project13/数据1/葛岗所/'
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

    print fname1_list
    print fname2_list
    #test
    for i in fname1_list:
        print i.decode('gbk')




    if os.path.isfile('../log.txt'):
        f = open('../log.txt','r')
        lines = f.readlines()
        f.close()

    else:
        lines = []
    #
    f = open('../log.txt','w')
    f.write(''.join(lines))
    for i in range(len(fname1_list)):
        fname1 = fname1_list[i]
        print fname1.decode('gbk')
        fname2 = fname2_list[i]
        print fname2.decode('gbk')
        print '---------------------'
        try:
            main(fname1.decode('gbk'),fname2.decode('gbk'),flist_root[i])
            f.write('done\t'+time.asctime(time.localtime(time.time()))+'\t'+fname1+'\n')
        except Exception,e:
            # raise Exception,e
            f.write('error\t'+time.asctime(time.localtime(time.time()))+'\t'+fname1+'\t'+str(e)+'\n')
            print e
            for i in range(5):
                print 'error'
    f.close()
    #
    # pass
