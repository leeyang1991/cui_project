# coding=gbk

import os
import requests
import zipfile
import shutil
import datetime
from tqdm import tqdm


this_root = os.getcwd()+'\\'
def mk_dir(dir):
    if not os.path.isdir(dir):
        os.makedirs(dir)

def update_scrpit():
    url = 'https://codeload.github.com/leeyang1991/cui_project/zip/master'
    # video_i = requests.get(url)
    now = datetime.datetime.now()
    year = now.year
    mon = now.month
    day = now.day
    zip_file_name =this_root+'{}_{}_{}.zip'.format(year,mon,day)

    downloadFILE(url,zip_file_name)

    date = '{}_{}_{}'.format(year,mon,day)
    return zip_file_name,date


def downloadFILE(url,name):
    resp = requests.get(url=url,stream=True)
    # print(resp.headers)
    # content_size = int(resp.headers['Content-Length'])/1024
    content_size = 231
    with open(name, "wb") as f:
        for data in tqdm(iterable=resp.iter_content(1024),ncols=100,total=content_size,unit='k',desc=name):
            f.write(data)


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
            if os.path.isfile(dest_dir+folder+'\\'+f):
                print(f)
                os.remove(dest_dir+folder+'\\'+f)
                shutil.copy(update+folder+'\\'+f,dest_dir+folder+'\\'+f)
            else:
                shutil.copy(update + folder + '\\' + f, dest_dir + folder + '\\' + f)



def main():
    zip_file_name, date = update_scrpit()
    unzip(zip_file_name,this_root+date)
    replace(date)

if __name__ == '__main__':
    main()