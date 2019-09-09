# coding=gbk
import os
import arcpy
import time
# root = 'D:\\project13\\output\\板木所\\S2017-03-22 1008马庄西变压器1\\'


#
# print folder_list

def merge(indir,out_name):
    rootdir = indir
    flist = os.listdir(rootdir)
    li = []
    for file in flist:
        if file.endswith('.shp'):
            l = rootdir+'\\'+file
            li.append(l)
    shp = out_name+'.shp'
    # if os.path.isfile(shp):
    #     os.remove(shp)
    #     os.remove(shp+'.xml')
    arcpy.Merge_management(li,shp)
    # arcpy.AddXY_management(shp)


father_root = 'D:\\project13\\output\\葛岗所\\'
if os.path.isfile('../log.txt'):
    f = open('../log.txt','r')
    lines = f.readlines()
    f.close()

else:
    lines = []
#
f = open('../log.txt','w')
f.write(''.join(lines))



for i in os.listdir(father_root):
    print i.decode('gbk')
    root = os.listdir(father_root+i)
    for j in root:
        s = 'merging'+i+'\\'+j
        print s.decode('gbk')
        try:
            merge(father_root+i+'\\'+j,father_root+i+'\\'+j)
            f.write('done\t'+time.asctime(time.localtime(time.time()))+'\t'+i+'\\'+j+'\n')
        except Exception,e:
            print Exception,':',e
            f.write('done\t'+time.asctime(time.localtime(time.time()))+'\t'+i+'\\'+j+str(e)+'\n')
f.close()

# for i in folder_list:
#     print 'merging',i
#     merge(root+i,root+i)