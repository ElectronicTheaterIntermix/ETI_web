from  selenium import webdriver
from BeautifulSoup import BeautifulSoup
import sys
import urllib2
import xlrd
import re
import socket

global nrows
nrows = 0
#remember to change it with your own path
path = "/home/kevinkao/python_program/Selenium/20140428_renew網頁改版finalETI網頁檔20130823.xls"

def open_excel():
        book = xlrd.open_workbook(path)
        sh = book.sheet_by_index(0)
        print('num col')
        print sh.ncols
        print('num row')
        print sh.nrows
        global nrows
        nrows = sh.nrows
        return sh

def excel(row_n, sh):
        data_list = []
        cell = sh.row(row_n)
        up = "(上)"
        up = up.decode('utf8')
        mid = "(中)"
        mid = mid.decode('utf8')
        down = "(下)"
        down = down.decode('utf8')
        up2 = "〈上〉"
        up2 = up2.decode('utf8')
        mid2 = "〈中〉"
        mid2 = mid2.decode("utf8")
        down2 = "〈下〉"
        down2 = down2.decode("utf8")
        up3 = "（上）"
        up3 = up3.decode("utf8")
        mid3 = "（中）"
        mid3 = mid3.decode("utf8")
        down3 = "（下）"
        down3 = down3.decode("utf8")
        if up in cell[2].value or mid in cell[2].value or down in cell[2].value or up2 in cell[2].value or mid2 in cell[2].value or down2 in cell[2].value or up3 in cell[2].value or mid3 in cell[2].value or down3 in cell[2].value:
                return cell[2].value
        elif cell[4].value == "":
                data_list.append("NULL")
                return data_list
        else :
                for i in range(16):
                        data_list.append(cell[i+2].value)
                return data_list
          
###################################

def split_list(temp_list):
        answer_list = []
        temp = []

        for i in range(len(temp_list)):
                if temp_list[i] == "":
                        continue
                else:
                        temp = temp_list[i].encode('utf8').split("；")
                        for j in range(len(temp)):
                                temp_temp = temp[j].split("：")
                                for z in range(len(temp_temp)):
                                        answer_list.append(temp_temp[z])
        return answer_list

###################################
def parse_html(temp_list):
        print(len(temp_list))
        answer_list = []
        temp = []
        for i in range(len(temp_list)):
                temp = str(temp_list[i]).decode("string_escape")[1:-1].split(" : ")
                answer_list.append(temp[0])
                answer_list.append(temp[1])

        return answer_list
###################################
dri = webdriver.Firefox()
dri.get("http://www.eti-tw.com/eti_new/")
all_tag_a = dri.find_elements_by_tag_name("a")

url_list = []
url_name = []
for i in range(len(all_tag_a)):
        url_list.append(str(all_tag_a[i].get_attribute('href')))
        url_name.append(all_tag_a[i].text)

error_list = []
error_name=[]
useable_list = []
for i in range(len(url_list)):
        try:
                code = urllib2.urlopen(url_list[i],timeout = 1)
                if code.getcode() is 200:
                        useable_list.append(url_list[i])
                else:
                        error_list.append(url_list[i])
                        error_name.append(url_name[i])
        except urllib2.URLError,e:
                error_list.append(url_list[i])
                error_name.append(url_name[i])
        except socket.timeout:
                print("time out: " + url_list[i])
                #print(url_list[i])

for i in range(len(error_list)):
        print(error_name[i])
        print(error_list[i])
##  URL can not work-------------

###
sh = open_excel()
print("num row for nrows")
print nrows
divided_list = []
no_videos = []
report = []

for i in range(5, nrows):
    if i == 81 or i == 134 or i == 179:
            continue
    dri.get("http://www.eti-tw.com/eti_new/")
    parse_list = excel(i, sh)
    if type(parse_list) is not list:
            divided_list.append(parse_list)
            continue
    if parse_list[0] == "NULL":
            print("NULL")
            continue
    input_tag = dri.find_elements_by_tag_name("input")
    input_tag[2].send_keys(parse_list[0])
    input_tag[2].submit()
    
    new_tag = dri.find_elements_by_tag_name("a")
    next_url = new_tag[20].get_attribute('href')
    url = next_url.encode("utf8")
    dri.get(url)
    diff_list = []
    diff_list.append(parse_list[0])
    #print(parse_list[0])
    #video test
    try:
        all_new_tag = dri.find_elements_by_tag_name("source")
        video = str(all_new_tag[0].get_attribute('src'))
        code = urllib2.urlopen(video)
        if code.getcode() is 200:
                print("it work")
    except urllib2.URLError,e:
        print("ERROR : %r"%e)
    except socket.timeout:
        print("timeout")
    except IndexError:
        no_videos.append(parse_list[0])
        print("No Video!")
    #video test

    # compare_list is the source html from website  in order to compare  tag  <p>
    # parse_list  are the data from excel
    HTML = dri.page_source
    soup = BeautifulSoup(HTML)
    r_td = soup.findAll("td")
    r_p = soup.findAll("p")
    p = re.compile(r'>.*<')
    compare_list = []
    for i in range(len(r_p)):
            print r_p[i].text
    for i in range(len(r_p)):
        comp = str(r_p[i])
        add = re.findall(p,comp)
        compare_list.append(add[0])
    compare_list.pop()  # in order to remove 多餘的部份
    parse_list.remove(parse_list[0]) # in order to remove 多餘的部份
    parse_list.remove(parse_list[0]) # in order to remove 多餘的部份
    the_answer = []
    parse_new = split_list(parse_list)
    the_answer = parse_html(compare_list)
    print("From the websire")
    for i in range(len(the_answer)):
        print(the_answer[i].decode("string_escape"))
    print("from EXCEL")
    for i in range(len(parse_new)):
        print(parse_new[i])
    ### compare p web
    for i in range(len(the_answer)):
        if str(the_answer[i].decode("string_escape")) in parse_new[i]:
            print("WORK!!!")
        else:
            print("Failed")
            print(parse_new[i])
            diff_list.append(parse_new[i].decode("utf8"))
    print("##############################")
    report.append(diff_list)

"""for i in range(len(divided_list)):
        print divided_list[i]"""

"""print("The List for No Videos:")
if len(no_videos) == 0:
        print("Every video is ok.")
else:
        for i in range(len(no_videos)):
                print no_videos[i]
print("***********************")
for row in report:
        for col in row:
                if len(row) == 1:
                        print col
                        print("No Problem!")
                else:
                        print col
        print("#################")"""

###########write file
result = open('report.txt', 'w')
result.write("The List for No Videos:\n")
if len(no_videos) == 0:
        result.write("Every video exists.\n")
else:
        for i in range(len(no_videos)):
                result.write(no_videos[i].encode("utf8"))
                result.write("\n")
result.write("************************\n")
for row in report:
        for col in row:
                print col
                result.write(col.encode("utf8"))
                result.write("\n")
                if len(row) == 1:
                        result.write("No Problem!\n")
        result.write("###################\n")
result.close()
#

