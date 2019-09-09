# coding=gbk

import xlrd
from collections import defaultdict


class ReadExcel:

    def __init__(self,fname1,fname2):
        # fname1 = 'D:\\project13\\数据1\\板木所\\S2017-03-22 1008马庄西变压器1\\马庄西台区.xls'
        # fname2 = 'D:\\project13\\数据1\\板木所\\S2017-03-22 1008马庄西变压器1\\台账\\计量箱与电能表的关系.xls'
        self.fname1 = fname1
        self.fname2 = fname2
        self.bk = xlrd.open_workbook(self.fname1)
        self.bk1 = xlrd.open_workbook(self.fname2)


    def didianyapeidianxiang(self):
        #配电箱只需要点一个点
        sh = self.bk.sheet_by_name("37.低压配电箱".decode('gbk'))
        nrows = sh.nrows
        lon = float(sh.cell_value(3,6))
        lat = float(sh.cell_value(3,7))
        value = sh.cell_value(3,4)
        return lon,lat,value

    def diyaqiangzhijia(self):
        #墙支架需要连线
        sh = self.bk.sheet_by_name("70.低压-墙支架".decode('gbk'))
        nrows = sh.nrows
        lon_list=[]
        lat_list=[]
        value_list=[]
        name_list=[]
        for i in range(nrows):
            lon = float(sh.cell_value(i+3,6))
            lat = float(sh.cell_value(i+3,7))
            value = sh.cell_value(i+3,10)
            name = sh.cell_value(i+3,2)
            lon_list.append(lon)
            lat_list.append(lat)
            value_list.append(value)
            name_list.append(name)
            if i == nrows - 4:
                break
        return lon_list,lat_list,value_list,name_list

    def dianxiangan(self):
        #电线杆
        sh = self.bk.sheet_by_name("33-1.低压杆塔（运行杆）".decode('gbk'))
        nrows = sh.nrows
        lon_list=[]
        lat_list=[]
        qianduan_list=[]
        name_list=[]
        for i in range(nrows):
            # print float(sh.cell_value(i+3,9))
            lon = float(sh.cell_value(i+3,9))
            lat = float(sh.cell_value(i+3,10))
            name = sh.cell_value(i+3,2)
            qianduan = sh.cell_value(i+3,12)
            lon_list.append(lon)
            lat_list.append(lat)
            qianduan_list.append(qianduan)
            name_list.append(name)
            if i == nrows - 4:
                break
        return lon_list,lat_list,name_list,qianduan_list

    def jiliangxiang(self):
        #计量箱
        sh = self.bk.sheet_by_name("40.计量箱".decode('gbk'))
        nrows = sh.nrows
        lon_list=[]
        lat_list=[]
        hang_list=[]
        lie_list=[]
        code_list=[]
        for i in range(nrows):
            lon = float(sh.cell_value(i+3,15))
            lat = float(sh.cell_value(i+3,16))
            hang = sh.cell_value(i+3,5)
            lie = sh.cell_value(i+3,6)
            code = sh.cell_value(i+3,1)
            lon_list.append(lon)
            lat_list.append(lat)
            hang_list.append(int(hang))
            lie_list.append(int(lie))
            code_list.append(code)
            if i == nrows - 4:
                break
        return lon_list,lat_list,hang_list,lie_list,code_list

    def yonghujierudian(self):
        #用户接入点
        sh = self.bk.sheet_by_name("69.用户接入点".decode('gbk'))
        nrows = sh.nrows
        lon_list=[]
        lat_list=[]
        # value_list=[]
        for i in range(nrows):
            lon = float(sh.cell_value(i+3,8))
            lat = float(sh.cell_value(i+3,9))
            value = sh.cell_value(i+3,1)
            lon_list.append(lon)
            lat_list.append(lat)
            # value_list.append(value)
            if i == nrows - 4:
                break
        return lon_list,lat_list


    def select_free_jiliangxiang(self):
        sh = self.bk1.sheet_by_name("4-1.计量箱与电能表的关系".decode('gbk'))
        nrows = sh.nrows
        code_list=[]
        name_list=[]
        for i in range(nrows):
            code = sh.cell_value(i+3,1)
            name = sh.cell_value(i+3,6)
            name_list.append(name)
            code_list.append(code)
            if i == nrows - 4:
                break

        #获取列表中重复元素的索引
        s = code_list
        d = defaultdict(list)
        for k,va in [(v,i) for i,v in enumerate(s)]:
            d[k].append(va)
        # print d


        dianbiao_dic = {}
        temp_list=[]
        for i in set(code_list):
            for j in d[i]:
                temp_list.append(name_list[j])
            dianbiao_dic[i]=temp_list
            temp_list = []
        # for i in dianbiao_dic:
        #     for j in dianbiao_dic[i]:
        #         print i,j

        #获取有空位的电表箱
        code_list1 = []
        for i in range(nrows):
            if sh.cell_value(i+3,2) == '':
                code1 = sh.cell_value(i+3,1)
                # print code1
                code_list1.append(code1)
            if i == nrows - 4:
                break
        spare_list = set(code_list1)

        return dianbiao_dic, spare_list

    def count_free_jiliangxiang(self):
        sh = self.bk1.sheet_by_name("4-1.计量箱与电能表的关系".decode('gbk'))
        nrows = sh.nrows
        # print nrows
        total = 0
        installed = 0
        for i in range(nrows):
            total += 1
            if sh.cell_value(i+3,2) != '':
                installed += 1
            if i == nrows - 4:
                break
        return total, installed


    def legend_info(self):
        sh = self.bk.sheet_by_name("39.低压线路".decode('gbk'))
        bianyaqi_code = sh.cell_value(3,6)

        return bianyaqi_code
if __name__ == '__main__':
    # select_free_jiliangxiang()
    #
    # # for i in select_free_jiliangxiang()[0]:
    # #     for j in select_free_jiliangxiang()[0][i]:
    # #         print i,j
    # #
    # # for i in select_free_jiliangxiang()[1]:
    # #     print i,'spare'
    #
    # dianbiao_dic, spare_list = select_free_jiliangxiang()
    # lon_list,lat_list,hang_list,lie_list,code_list = jiliangxiang()
    #
    # red_lon=[];red_lat=[];red_hang_lie=[]
    # black_lon=[];black_lat=[];black_hang_lie=[]
    # for i in dianbiao_dic:
    #     if i in spare_list:
    #         # print i,'red'
    #         index = code_list.index(i)
    #         red_lon.append(lon_list[index])
    #         red_lat.append(lat_list[index])
    #         red_hang_lie.append(str(hang_list[index])+','+str(lie_list[index]))
    #         print lon_list[index]
    #         print lat_list[index]
    #         print str(hang_list[index])+','+str(lie_list[index])
    #         print i
    #         print '-------------------------------red'
    #     else:
    #         index = code_list.index(i)
    #         black_lon.append(lon_list[index])
    #         black_lat.append(lat_list[index])
    #         black_hang_lie.append(str(hang_list[index])+','+str(lie_list[index]))
    #         print lon_list[index]
    #         print lat_list[index]
    #         print str(hang_list[index])+','+str(lie_list[index])
    #         print i
    #         print '-------------------------------black'
    pass





