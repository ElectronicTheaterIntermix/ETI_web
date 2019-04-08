#!/usr/bin/python
# -*- coding: utf8 -*-
from selenium import webdriver
from BeautifulSoup import BeautifulSoup
import sys
import urllib2
import xlrd
import re		# regular expression
import socket
import xlwt

global nrows
nrows = 0 #initial value

#remember to change it with your own path
p = "E:/renew網頁編碼20140806.xls"
path = p.decode('utf8')

def open_excel():
    book = xlrd.open_workbook(path)
    sh = book.sheet_by_index(1)		# 第二個工作表

    print('col: ')
    print sh.ncols
    print('row: ')
    print sh.nrows
    global nrows
    nrows = sh.nrows
    return sh

def read_excel(row_num, sh):
    data_list = []
    cell = sh.row(row_num)
    #string = str(cell[4]).encode('unicode')

    if cell[0].value == u"ETI原始編碼" or cell[0].value == "":
        data_list.append("")
        return data_list		#return null(?)
    else:
        for i in range(9):
            data_list.append(cell[i].value)
        return data_list		#return column "ETI原始編碼" ~ "表演者"

def check_link(dri):
    # check the url in the homepage
    tag_a = dri.find_elements_by_tag_name("a")
    url_list = [] # the url_list with tag_a
    url_name = [] # the text in the url
    for i in range(len(tag_a)):
        url_list.append(str(tag_a[i].get_attribute('href')))
        url_name.append(tag_a[i].text)

    error_list = []
    error_name = []
    useable_list = []

    for i in range(len(url_list)):
        try:
            code = urllib2.urlopen(url_list[i], timeout = 1)		#
            if code.getcode() is 200:								#
                useable_list.append(url_list[i])
            else:
                error_list.append(url_list[i])
                error_name.append(url_name[i])
        except urllib2.URLError,e:
            error_list.append(url_list[i])
            error_name.append(url_name[i])
        except socket.timeout:
            print("timeout: " + url_list[i])

    for i in range(len(error_list)):
        print(error_name[i])
        print(error_list[i])

def split_list(temp_list):		# 分解excel的"製作群"欄
    answer_list = []
    temp = []

    for i in range(len(temp_list)):
        if (temp_list[i] is ''):
            continue
        else:
            temp = temp_list[i].encode('utf8').split("；")
            for j in range(len(temp)):
                temp_temp = temp[j].split("：")
                for z in range(len(temp_temp)):
                    answer_list.append(temp_temp[z])
    return answer_list

def parse_html(temp_list):
    #print(len(temp_list))
    answer_list = []
    temp = []
    for i in range(len(temp_list)):
        temp = str(temp_list[i]).decode("string_escape")[1:-1].split(" : ")		#
        temp[0] = temp[0].strip(" \n")
        temp[1] = temp[1].strip(" \n")
        answer_list.append(temp[0])
        answer_list.append(temp[1])
    return answer_list

def search_drama_and_get_in( dri, excel_data ):
    # 
    result = ""
    input_drama = ""
    dri.get("http://www.eti-tw.com")
    input_tag = dri.find_elements_by_tag_name("input")
    input_drama = excel_data[2].replace("\n", "")
    input_tag[2].send_keys(input_drama) 			# excel_data[2] is the name of the drama
    input_tag[2].submit()
    dri.implicitly_wait(60)
    # new_tag = drama_name

    # but what if there are multiple options?
    dd = dri.find_elements_by_tag_name("h5")		# 搜尋結果的標題(作品名稱)
    print("dd: ", dd)
    #drama_name = dd[0].find_element_by_tag_name("a")
    #next_url = drama_name.get_attribute('href')
    #print( "drama_name : " + drama_name.text )

    if not dd:
        dri.get("http://www.eti-tw.com")

		# 有些作品後面要填空格才找的到
        input_drama = ""
        input_drama = excel_data[2].replace("\n", " ")
        input_tag = dri.find_elements_by_tag_name("input")
        input_tag[2].send_keys(input_drama)
        input_tag[2].submit()
        dri.implicitly_wait(60)
        dd = dri.find_elements_by_tag_name("h5")

        print ("dd: ", dd)
        if not dd:
            print("not dd")
            return "fail"
        drama_name = dd[0].find_element_by_tag_name("a")
        next_url = drama_name.get_attribute('href')             # 取得該劇的連結網址，ex. show.php?itemID=XueMH001&id=67
        result = "success"
    else:
        drama_name = dd[0].find_element_by_tag_name("a")
        next_url = drama_name.get_attribute('href')
        result = "success"

    #print("drama_name" + drama_name[20].text)
    url = next_url.encode('utf8')
    dri.get(url)
    return result

def check_staff(split_web_data, split_excel_data):
    staff_diff_list = []
    #staff_diff_list.append(split_excel_data[2])
    entitle_diff_list =[]
    #entitle_diff_list.append( split_excel_data[2] )
    flag = 0

    # title要一樣才會比對
    # excel "製作群" 欄可能以 '\n' 結尾，下方程式碼會判定為different，ex. '娃娃活木工教室\n'
    for i in range(0, len(split_web_data), 2):
        flag = 0
        for j in range(3, len(split_excel_data), 2):
            if str(split_web_data[i].decode("string_escape")) == str(split_excel_data[j].decode("string_escape")):
                if str(split_web_data[i+1].decode("string_escape")) == str(split_excel_data[j+1].decode("string_escape")):
                    flag = 1
                    break
                else:
                    # It means that the title is the same, but the staffs are different
                    staff_diff_list.append(split_web_data[i].decode("string_escape"))
                    staff_diff_list.append(split_web_data[i+1].decode("string_escape"))
                    flag = 1
                    break
            else:
                continue
        #print("flag = ", flag)
        if flag == 0:
            # It means that the title and the staff are totally different
            entitle_diff_list.append(split_web_data[i].decode("string_escape"))
            entitle_diff_list.append(split_web_data[i+1].decode("string_escape"))

    return staff_diff_list, entitle_diff_list

def write_excel_init():
    # Create workbook and sheet
    workbook = xlwt.Workbook(encoding = 'utf8')
    sheet = workbook.add_sheet( "Result", cell_overwrite_ok = True )

    # Initialize the labels
    sheet.write( 0, 0, "ETI原始編碼" )
    sheet.write( 0, 1, "大寫編碼" )
    sheet.write( 0, 2, "作品名稱" )
    sheet.write( 0, 3, "職稱相同，內容不同" )
    sheet.write( 0, 4, "職稱和內容皆不同" )
    sheet.write( 0, 5, "影片是否正常")

    sheet.col(0).width = 256 * len("ETI原始編碼")
    sheet.col(1).width = 256 * len("大寫編碼")
    sheet.col(2).width = 256 * len("作品名稱")
    sheet.col(3).width = 256 * len("職稱相同，內容不同")
    sheet.col(4).width = 256 * len("職稱和內容皆不同")
    sheet.col(5).width = 256 * len("影片是否正常")

    workbook.save("./foobar.xls")

    return workbook, sheet

def main_check( total_rows, driver, sheet ):
    dri = driver
    nrows = total_rows
    sh = sheet

    out_workbook, out_sheet = write_excel_init()
    maxDramaLen = 0
    drama_name_len = 0
    count = 0

    for i in range( 380, nrows ):
        excel_data = []
        excel_data = read_excel(i, sh)		# return column "ETI原始編碼" to "表演者" if 原始編碼 exist

        print("############ Row Number " + str(i) + "################")

        if len(excel_data) == 1:			# not a work to check
            continue

        count = count + 1

        maxDramaLen = max( drama_name_len, len(excel_data[2]) )			# excel_data[2]是作品名稱
        drama_name_len = len(excel_data[2])

        search_result = ""
        search_result = search_drama_and_get_in( dri, excel_data )		# return "success" or "fail"

        # write excel
        ###
        out_sheet.write( count, 0, excel_data[0] )		# ETI原始編碼
        out_sheet.write( count, 1, excel_data[1] )		# 大寫編碼
        out_sheet.write( count, 2, excel_data[2] )		# 作品名稱

        if search_result == "fail":						# 搜尋不到
            out_sheet.write( count, 3, "not found")
            continue

        # take the content from webpage
        # web_list is the source html from website in order to compare tag <p>
        # excel_data is the data from excel
        HTML = dri.page_source
        soup = BeautifulSoup( HTML )
        r_td = soup.findAll( "td" )						# 找網頁中所有的<td>
        r_p = soup.findAll( "p" )						# 找網頁中所有的<p>
        p = re.compile( r'>.*<', re.DOTALL )			# ">" "<"間的所有char
        web_list = [] # web_list = compare_list

        #for i in range(len(r_p)):
        #    print(r_p[i].text)
        for i in range(len(r_p)):
            comp = str(r_p[i])              # <p>導演:薛美華<p>
            add = re.findall(p,comp)		# re.findall(pattern, string)    Return all non-overlapping matches of pattern in string
                                            # 型別:type list；內容:>\xe5\xb0\x8e\xe6\xbc\x94:\xe8\x96\x9b\xe7\xbe\x8e\xe8\x8f\
                                            #                     >導演:薛美華<
            #print("add : " + add[0])
            web_list.append(add[0])

        web_list.pop()

        the_answer = []
        split_excel_data = split_list(excel_data) # split_excel_data = parse_new
                                                  # 變數內容長這樣 http://imgur.com/QWUEU3P
        split_web_data = parse_html(web_list) # split_web_data = the_answer
                                              # 變數內容長這樣 http://imgur.com/m5DuCQx

        staff_diff_list, entitle_diff_list = check_staff( split_web_data, split_excel_data )



        string = ""
        for i in range(len(staff_diff_list)):
            string = string + staff_diff_list[i] + ';' + '\n'
        out_sheet.write( count, 3, string )				# 職稱相同，內容不同


        string = ""
        for i in range(len(entitle_diff_list)):
            string = string + entitle_diff_list[i] + ';' + '\n'
        out_sheet.write( count, 4, string)				# 職稱和內容皆不同

        out_sheet.row( count ).height = 256 * max( len(staff_diff_list), len(entitle_diff_list) ) / 2
        out_workbook.save("./foobar.xls")

        # check videos
        check_video( dri, out_sheet, out_workbook, count )

        ###
        # Print the web data that the title are the same, but the  staffs are different
        #print( "Staff Different" )
        #for i in range(len(staff_diff_list)):
        #    print(staff_diff_list[i])

        # Print the web data that the title and the staff are totally different
        #print( "Totally Different" )
        #for i in range(len(entitle_diff_list)):
        #    print(entitle_diff_list[i])
        dri.get("http://www.eti-tw.com")

    #print("maxDramaLen: " + str(maxDramaLen))
    out_sheet.col(2).width = 256 * 2 * maxDramaLen
    out_workbook.save("./foobar.xls")

def check_video( dri, sheet, workbook, count ):
    no_videos = []
    try:
        # all_new_tag = video_tag
        video_tag = dri.find_elements_by_tag_name("source")
        video = str(video_tag[0].get_attribute('src'))
        code = urllib2.urlopen(video)
        if code.getcode() is 200:
            print("it works")
            sheet.write( count, 5, "v")					# 影片是否正常

    except urllib2.URLError,e:
        print("ERROR : %r"%e)
        sheet.write( count, 5, "%r"%e)
    except socket.timeout:
        print("timeout")
        sheet.write( count, 5, "timeout")
    except IndexError:
        #no_videos.append(excel_data[2])
        print("No Video!")
        sheet.write( count, 5, "x")
    workbook.save("./foobar.xls")

if __name__ == "__main__":
    sh = open_excel()
    # check the global variable
    print("global nrows: ")
    print nrows

    #create the browser
    dri = webdriver.Firefox()
    dri.get("http://www.eti-tw.com")

    # check the links in homepage
    #check_link(dri)

    # check the data between the web and the excel file
    main_check( nrows, dri, sh )
    # Write excel



    '''style = xlwt.easyxf('font: bold 1, height 280') # height 280 = 14 * 20 (unit: 1/20 point)
    #sheet.col(0).width = 256 * ( len('foobar') + 1 )
    data = ['1', '2', '3546']
    #sheet.write(0, 0, 'foobar') # row, column, value
    string = ""
    for i in range(len(data)):
        string = string + data[i] + ';'
    sheet.write(0, 0, string)'''



    # check the videos
    '''
    # video test
    no_videos = []
    try:
        # all_new_tag = video_tag
        video_tag = dri.find_elements_by_tag_name("source")
        video = str(video_tag[0].get_attribute('src'))
        code = urllib2.urlopen(video)
        if code.getcode() is 200:
            print("it works")
    except urllib2.URLError,e:
        print("ERROR : %r"%e)
    except socket.timeout:
        print("timeout")
    except IndexError:
        no_videos.append(excel_data[2])
        print("No Video!")
    # video test'''
