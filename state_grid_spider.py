# coding=utf-8
import pandas as pd
from selenium import webdriver
import time
import re
import numpy as np
from bs4 import BeautifulSoup
import hashlib
from os.path import *
import os

url = "http://b.esgcc.com.cn/SearchMaterial/showSearchPage?m=TlRBd01UUXdOalU0&chanel=13&isIndex=false&record=true&mSite=MjEwMDY=&first_category=&second_category=&third_category=&title=5rKz5Y2X56uZLeWKnuWFrOeUqOWTgeWPiumdnueUtee9kembtuaYn+eJqei1hOmAiei0reS4k+WMug==&$isHot=false"
chrome_drive_path = '/Users/liyang/Downloads/chromedriver'
cookie_f = 'cookies'
silent = False

def pause():
    # ANSI colors: https://gist.github.com/rene-d/9e584a7dd2935d0f461904b9f2950007
    input('\33[7m'+"PRESS ENTER TO CONTINUE."+'\33[0m')

def get_all_page(html):
    p = re.findall('id="allPageNum">.*?</i></span>', html)
    all_page = p[0].replace('id="allPageNum">', '')
    all_page = all_page.replace('</i></span>', '')
    all_page = int(all_page)
    return all_page

def hash_key(str_in):
    str_in = str_in.encode('utf-8')
    readable_hash = hashlib.sha256(str_in).hexdigest()
    return readable_hash

def get_cookie():
    f = cookie_f
    if isfile(f):
        cookie = read_cookie(f)
        return cookie
    else:
        get_new_cookie()
        cookie = read_cookie(f)
        return cookie

def dic_to_df(dic, key_col_str='__key__'):
    '''
    :param dic:
    {
    row1:{col1:val1, col2:val2},
    row2:{col1:val1, col2:val2},
    row3:{col1:val1, col2:val2},
    }
    :param key_col_str: define a Dataframe column to store keys of dict
    :return: Dataframe
    '''
    data = []
    columns = []
    index = []
    all_cols = []
    for key in dic:
        vals = dic[key]
        for col in vals:
            all_cols.append(col)
    all_cols = list(set(all_cols))
    all_cols.sort()
    for key in dic:
        vals = dic[key]
        if len(vals) == 0:
            continue
        vals_list = []
        col_list = []
        vals_list.append(key)
        col_list.append(key_col_str)
        for col in all_cols:
            if not col in vals:
                val = np.nan
            else:
                val = vals[col]
            vals_list.append(val)
            col_list.append(col)
        data.append(vals_list)
        columns.append(col_list)
        index.append(key)
    df = pd.DataFrame(data=data, columns=columns[0], index=index)
    return df

def df_to_excel(df, dff, n=None, random=False):
    if n == None:
        df.to_excel('{}.xlsx'.format(dff))
    else:
        if random:
            df = df.sample(n=n, random_state=1)
            df.to_excel('{}.xlsx'.format(dff))
        else:
            df = df.head(n)
            df.to_excel('{}.xlsx'.format(dff))

def str_strip(in_str):
    in_str = str(in_str)
    in_str = in_str.strip()
    in_str = in_str.replace(' ', '')
    return in_str
    pass

def parse_html(html,product_number):
    soup = BeautifulSoup(html, 'html5lib')
    product_item_obj_list = soup.find('ul',{'class':'product_list'})
    result_dic = {}
    for product_obj in product_item_obj_list:
        try:
            price = product_obj.find_all('div',{'class':'item_price'})[0]
            item_name = product_obj.find_all('input',{'class':'item_name','type':'hidden',})[0]
            company_name = product_obj.find_all('div',{'class':'item_store'})[0]
            sell_number = product_obj.find_all('span',{'class':'fl'})[0]
            url = product_obj.find_all('a',{'class':'item_img'})[0]
        except:
            continue
        key = hash_key(str(product_obj))

        price = str_strip(price)
        item_name = str_strip(item_name)
        company_name = str_strip(company_name)
        sell_number = str_strip(sell_number)
        url = str_strip(url)

        price_split = price.split('\n')
        price = price_split[1]
        # print(price)
        # print('====')
        p = re.findall('value=".*?"/>',item_name)[0]
        item_name = p.split('"')[1]
        # print(item_name)

        # print('====')
        company_name = company_name.split('\n')[1]
        p = re.findall('title=".*?">',company_name)[0]
        company_name = p.split('"')[1]
        # print(company_name)
        # print('====')
        p = re.findall('<em>.*?</em>', sell_number)[0]
        sell_number = p.replace('<em>','')
        sell_number = sell_number.replace('</em>','')
        # print(sell_number)
        # print('====')
        p = re.findall('href=".*?"target', url)[0]
        url = p.split('"')[1]

        result_dic[key] = {
            '价格':price,
            '名称':item_name,
            '公司':company_name,
            '销售数量':sell_number,
            '网址':url,
            '物料编码':product_number,
        }
        text = '\n'.join([item_name,company_name,price])
        print('----------------------------')
        print(text)
    df = dic_to_df(result_dic,'唯一编码')
    return df

def start_spder(driver,product_number):

    # product_number = '500140592'
    # product_number = '500140658'

    input = driver.find_elements_by_id("to_seach_id")
    for i in input:
        if i.get_attribute("class") == "searchInput":
            i.clear()
            i.send_keys(product_number)
            driver.find_element_by_class_name("searchBtn").click()
    html = driver.page_source
    if 'login-box' in html:
        print('cookie is expired')
        os.remove(cookie_f)
        driver.quit()
        driver = init_driver(silent)
        html = driver.page_source

    df = parse_html(html,product_number)

    all_page = get_all_page(html)
    # for page in tqdm(range(all_page)):
    for page in range(all_page):
        try:
            next_pate = driver.find_elements_by_class_name('nextPag')[0]
        except:
            continue
        next_pate.click()
        html = driver.page_source
        df_next_page = parse_html(html, product_number)
        df = df.append(df_next_page)
    df = df.drop_duplicates(subset=['唯一编码'])
    df_to_excel(df,product_number)

def get_new_cookie():
    driver = webdriver.Chrome(chrome_drive_path)
    driver.get(url)
    # pause()
    while 1:
        status = get_status(driver)
        time.sleep(1)
        if status == 'Dead':
            return

def get_status(driver):
    try:
        cookies = driver.get_cookies()
        new_cookies = []
        for dic_i in cookies:
            new_dic_i = {}
            for key in dic_i:
                val = dic_i[key]
                try:
                    if 'esgcc.com.cn' in val:
                        val = '.esgcc.com.cn'
                except:
                    pass
                new_dic_i[key] = val
            new_cookies.append(new_dic_i)
        save_cookie_to_txt(new_cookies, cookie_f)
        return "Alive"
    except:
        return 'Dead'

def save_cookie_to_txt(cookie, outf):
    fw = outf
    fw = open(fw, 'w')
    fw.write(str(cookie))
    fw.close()

def read_cookie(f):
    fr = open(f).read()
    dic = eval(fr)
    return dic


def init_driver(silent=True):
    cookie = get_cookie()
    opts = webdriver.ChromeOptions()
    # opts.headless = True
    opts.headless = silent
    driver = webdriver.Chrome(chrome_drive_path, options=opts)
    driver.get(url)
    for c in cookie:
        c = dict(c)
        driver.add_cookie(c)
    driver.get(url)
    return driver

def main():
    driver = init_driver(silent=silent)
    # product_number = '500140592'
    # product_number = '500140658'
    product_list = ['500140592','500140658'][::-1]
    for product in product_list:
        start_spder(driver,product)



if __name__ == '__main__':
    main()