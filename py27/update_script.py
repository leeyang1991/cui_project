# coding=gbk
# python 3
import os
# import urllib
# urllib.request
from urllib import request
# from urllib import
import zipfile
import shutil
import datetime
from tqdm import tqdm
# import urllib.request

this_root = os.getcwd()+'\\'
def mk_dir(dir):
    if not os.path.isdir(dir):
        os.makedirs(dir)

def update_scrpit():
    url = 'https://codeload.github.com/leeyang1991/cui_project/zip/master'
    now = datetime.datetime.now()
    year = now.year
    mon = now.month
    day = now.day
    zip_file_name =this_root+'{}_{}_{}.zip'.format(year,mon,day)
    date = '{}_{}_{}'.format(year, mon, day)
    downloadFILE(url,zip_file_name)
    return zip_file_name,date


def downloadFILE(url,name):
    resp = request.urlopen(url=url)
    resp = resp.read()
    # print(resp)
    # content_size = int(resp.headers['Content-Length'])/1024
    # content_size = 231
    with open(name, "wb") as f:
        f.write(resp)
        # for data in tqdm(iterable=resp.iter_content(1024),ncols=100,total=content_size,unit='k',desc=name):
        #     f.write(data)


def unzip(zip,move_dst_folder):
    mk_dir(move_dst_folder)
    path_to_zip_file = zip
    directory_to_extract_to = move_dst_folder
    zip_ref = zipfile.ZipFile(path_to_zip_file, 'r')
    zip_ref.extractall(directory_to_extract_to)
    zip_ref.close()
    os.remove(zip)


def replace(date):

    update = this_root+date+'\\cui_project-master\\'
    dest_dir = this_root+'..\\'
    for folder in os.listdir(update):
        for f in os.listdir(update+folder):
            try:
                if os.path.isfile(dest_dir+folder+'\\'+f):
                    print(f)
                    os.remove(dest_dir+folder+'\\'+f)
                    shutil.copy(update+folder+'\\'+f,dest_dir+folder+'\\'+f)
                else:
                    shutil.copy(update + folder + '\\' + f, dest_dir + folder + '\\' + f)
            except:
                pass


def main():
    zip_file_name, date = update_scrpit()
    unzip(zip_file_name,this_root+date)
    replace(date)

if __name__ == '__main__':
    main()