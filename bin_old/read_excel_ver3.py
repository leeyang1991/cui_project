# coding=gbk

import xlrd
import xlwt
from collections import defaultdict


class ReadExcel:

    def __init__(self,fname1,fname2,fname3,fname4):
        '''
        :param fname1: D:\project13\input\data\����ũ��̨��\����ũ��̨��.xls
        :param fname2: D:\project13\input\data\����ũ��̨��\̨��\����������ܱ�Ĺ�ϵ.xls
        :param fname3: D:\project13\input\�Ͽ����û�����.xls
        :param fname4: D:\project13\input\�Ͽ�ͼ����Ϣ.xlsx
        '''
        self.legend = "Sheet1"
        self.username = "�ǹ�1�����ܱ���ϸ"
        self.fname1 = fname1
        self.fname2 = fname2
        self.fname3 = fname3
        self.fname4 = fname4
        self.bk1 = xlrd.open_workbook(self.fname1)
        self.bk2 = xlrd.open_workbook(self.fname2)
        self.bk3 = xlrd.open_workbook(self.fname3)
        self.bk4 = xlrd.open_workbook(self.fname4)
        self.write_new_excel()



    def didianyapeidianxiang(self):
        #�����ֻ��Ҫ��һ����
        sh = self.bk1.sheet_by_name("37.��ѹ�����".decode('gbk'))
        nrows = sh.nrows
        lon = float(sh.cell_value(3,6))
        lat = float(sh.cell_value(3,7))
        value = sh.cell_value(3,4)
        return lon,lat,value


    def diyaqiangzhijia(self):
        #ǽ֧����Ҫ����
        sh = self.bk1.sheet_by_name("70.��ѹ-ǽ֧��".decode('gbk'))
        nrows = sh.nrows
        if nrows == 3:
            return None
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
        #���߸�
        sh = self.bk1.sheet_by_name("33-1.��ѹ���������иˣ�".decode('gbk'))
        nrows = sh.nrows
        if nrows == 3:
            return None
        lon_list=[]
        lat_list=[]
        qianduan_list=[]
        name_list=[]
        for i in range(nrows):
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
        #������
        sh = self.bk1.sheet_by_name("40.������".decode('gbk'))
        nrows = sh.nrows
        lon_list=[]
        lat_list=[]
        code_list=[]
        num_list=[]
        hang_lie_str_list=[]
        tuopushebei_list = []
        for i in range(nrows):
            lon = float(sh.cell_value(i+3,15))
            lat = float(sh.cell_value(i+3,16))
            hang = int(sh.cell_value(i+3,5))
            lie = int(sh.cell_value(i+3,6))
            code = sh.cell_value(i+3,1)
            num = hang*lie
            hang_lie_str = str(hang)+','+str(lie)
            tuopushebei = sh.cell_value(i+3,19)

            lon_list.append(lon)
            lat_list.append(lat)
            code_list.append(code)
            num_list.append(num)
            hang_lie_str_list.append(hang_lie_str)
            tuopushebei_list.append(tuopushebei)
            if i == nrows - 4:
                break

        jiliangxiang_dic = {}
        for i in range(len(lon_list)):
            jiliangxiang_dic[code_list[i]] = [lon_list[i],lat_list[i],hang_lie_str_list[i],num_list[i],tuopushebei_list[i]]
        return jiliangxiang_dic


    def yonghujierudian(self):
        #�û������
        sh = self.bk1.sheet_by_name("69.�û������".decode('gbk'))
        nrows = sh.nrows
        lon_list=[]
        lat_list=[]
        qianduan_list=[]
        shebei_list = []
        for i in range(nrows):
            lon = float(sh.cell_value(i+3,8))
            lat = float(sh.cell_value(i+3,9))
            qianduan = sh.cell_value(i+3,6)
            shebei = sh.cell_value(i+3,2)

            lon_list.append(lon)
            lat_list.append(lat)
            qianduan_list.append(qianduan)
            shebei_list.append(shebei)
            if i == nrows - 4:
                break
        return lon_list,lat_list,qianduan_list,shebei_list


    def user_name(self):
        # D:\project13\input\�Ͽ����û�����.xls�ֵ�
        # sh = self.bk3.sheet_by_name(self.username.decode('gbk'))
        sh = self.bk3.sheet_by_index(0)
        nrows = sh.nrows
        dianbiao_code = []
        user_name = []
        for i in range(nrows):
            if i == nrows - 2:
                break
            dianbiao_code.append(sh.cell_value(i+2,6))
            # dianbiao_code.append(sh.cell_value(i+2,9))
            user_name.append(sh.cell_value(i+2,4))
            # user_name.append(sh.cell_value(i+2,1))

        user_name_dic = {}
        for i in range(len(user_name)):
            user_name_dic[dianbiao_code[i]] = user_name[i]

        return user_name_dic



    def write_new_excel(self):
        # sh = self.bk2.sheet_by_name("4-1.����������ܱ�Ĺ�ϵ".decode('gbk'))
        sh = self.bk2.sheet_by_index(0)
        user_name_dic = self.user_name()
        nrows = sh.nrows
        bk_w = xlwt.Workbook()
        sheet1 = bk_w.add_sheet('sheet 1')
        for i in range(nrows):
            if nrows - 2 == i:
                break
            jiliangxiang = (sh.cell_value(i+2,1))
            dianbiao = (sh.cell_value(i+2,3))
            sheet1.write(i+2,1,jiliangxiang)
            sheet1.write(i+2,3,dianbiao)
            try:
                sheet1.write(i+2,6,user_name_dic[dianbiao])
            except:
                sheet1.write(i+2,6,'')
        # fname2 = 'D:\project13\input\data\����ũ��̨��\̨��\����������ܱ�Ĺ�ϵ.xls'
        bk_w.save((self.fname2+'_new.xls'))


    def full_and_spare_and_transparent_jiliangxiang(self):
        jiliangxiang_dic = self.jiliangxiang()
        # print jiliangxiang_dic
        fname_new = (self.fname2+'_new.xls')
        bk_new = xlrd.open_workbook(fname_new)
        sh = bk_new.sheet_by_name('sheet 1')
        nrows = sh.nrows
        jiliangxiang_code = []
        name_list = []
        for i in range(nrows):
            if i == nrows - 2:
                break
            if not sh.cell_value(i+2,1) == '':
                jiliangxiang_code.append(sh.cell_value(i+2,1))
                name_list.append(sh.cell_value(i+2,6))
        #��ȡ�б����ظ�Ԫ�ص�����
        s = jiliangxiang_code
        d = defaultdict(list)
        for k,va in [(v,i) for i,v in enumerate(s)]:
            d[k].append(va)

        dianbiao_dic = {}
        temp_list=[]
        for i in set(jiliangxiang_code):
            for j in d[i]:
                temp_list.append(name_list[j])
            try:
                dianbiao_dic[i]=[jiliangxiang_dic[i][0],jiliangxiang_dic[i][1],temp_list]
                temp_list = []
            except:
                pass


        red = []
        black = []
        for i in d:
            try:
                if len(d[i]) < jiliangxiang_dic[i][3]:
                    red.append([i,jiliangxiang_dic[i][0],jiliangxiang_dic[i][1],jiliangxiang_dic[i][2]])
                else:
                    black.append([i,jiliangxiang_dic[i][0],jiliangxiang_dic[i][1],jiliangxiang_dic[i][2]])
            except:
                pass
        # print dianbiao_dic
        # print red
        # print black
        return dianbiao_dic,red,black

    def info(self):
        # sh = self.bk4.sheet_by_name(self.legend.decode('gbk'))
        sh = self.bk4.sheet_by_index(0)
        nrows = sh.nrows
        PMS = []
        taiqu_code = []
        taiqu_name = []
        bianyaqi_type = []
        bianyaqi_content = []
        prime_line_type = []
        branch_line_type = []
        for i in range(nrows):
            if i == nrows - 1:
                break
            PMS.append(sh.cell_value(i+1,1))
            taiqu_code.append(sh.cell_value(i+1,2))
            taiqu_name.append(sh.cell_value(i+1,3))
            bianyaqi_type.append(sh.cell_value(i+1,4))
            bianyaqi_content.append(int(sh.cell_value(i+1,5)))
            prime_line_type.append(sh.cell_value(i+1,6))
            branch_line_type.append('')
        info_dic = {}
        for i in range(len(PMS)):
            info_dic[i] = [PMS[i],taiqu_code[i],taiqu_name[i],bianyaqi_type[i],bianyaqi_content[i],prime_line_type[i],branch_line_type[i]]
        return info_dic


    def count_intalled_jiliangxiang(self):
        jiliangxiang_dic = self.jiliangxiang()
        fname_new = self.fname2+'_new.xls'
        bk_new = xlrd.open_workbook(fname_new)
        sh = bk_new.sheet_by_name('sheet 1')
        nrows = sh.nrows
        installed = 0
        for i in range(nrows):
            if not sh.cell_value(i,3) == '' and sh.cell_value(i,1) in jiliangxiang_dic:
                installed += 1
        return installed

    def count_total_jiliangxiang(self):
        jiliangxiang_dic = self.jiliangxiang()
        total = 0
        for i in jiliangxiang_dic:
            total += jiliangxiang_dic[i][3]
        return total


    def legend_info(self):
        sh = self.bk1.sheet_by_name("39.��ѹ��·".decode('gbk'))
        bianyaqi_code = sh.cell_value(3,6)
        return bianyaqi_code


    def dianlan(self):
        try:
            sh = self.bk1.sheet_by_name("35.��ѹ���¶Σ���ѡ��".decode('gbk'))
        except:
            sh = self.bk1.sheet_by_name("35.��ѹ���¶�".decode('gbk'))
        nrows = sh.nrows
        lon_list = []
        lat_list = []
        num_list = []
        for i in range(nrows):
            if i == nrows - 3:
                break
            lon_list.append(float(sh.cell_value(i+3,15)))
            lat_list.append(float(sh.cell_value(i+3,16)))
            num_list.append(int(sh.cell_value(i+3,0)))
        #��ȡ�б����ظ�Ԫ�ص�����
        s = num_list
        d = defaultdict(list)
        for k,va in [(v,i) for i,v in enumerate(s)]:
            d[k].append(va)

        # for i in d:
        #     print i
        #     print d[i]
        lines = []
        for i in d:
            temp = []
            for j in d[i]:
                temp.append([lon_list[j],lat_list[j]])
            lines.append(temp)
        p1_list = []
        p2_list = []
        for i in lines:
            for j in range(len(i)):
                if j + 1 == len(i):
                    break
                p1_list.append(i[j])
                p2_list.append(i[j+1])
                j += 1

        # print p1_list
        # print p2_list
        # print 'p1 len',len(p1_list)
        # print 'p2 len',len(p2_list)
        # print p1_list[-3:]
        # print p2_list[-3:]
        return p1_list,p2_list

    def fenzhiiang(self):
        sh = self.bk1.sheet_by_name("34.��ѹ���·�֧��".decode('gbk'))
        nrows = sh.nrows
        coor_dic = {}
        for i in range(nrows):
            if nrows == 3:
                return None
            if i == nrows - 3:
                break
            lon = float(sh.cell_value(i+3,6))
            lat = float(sh.cell_value(i+3,7))
            name = sh.cell_value(i+3,2)
            coor_dic[i] = [lon,lat,name]
        return coor_dic




if __name__ == '__main__':
    fname1 = 'D:\project13\input\data\���19#���ش�̱�\���19#���ش�̱�.xls'
    fname2 = 'D:\project13\input\data\���19#���ش�̱�\̨��\����������ܱ�Ĺ�ϵ.xls'
    fname3 = 'D:\project13\input\��1��.xlsx'
    fname4 = 'D:\project13\input\ͼ����Ϣ.xlsx'

    r = ReadExcel(fname1,fname2,fname3,fname4)
    # print r.full_and_spare_and_transparent_jiliangxiang()[0]
    # print r.full_and_spare_and_transparent_jiliangxiang()[1]
    # print r.full_and_spare_and_transparent_jiliangxiang()[2]
    # print r.count_intalled_jiliangxiang()
    # print r.count_total_jiliangxiang()
    r.dianlan()