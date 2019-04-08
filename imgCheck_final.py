from selenium import webdriver
import sys
import urllib2
import socket
driver = webdriver.Firefox()
driver.get("http://eti-tw.com/eti_new/")
search_box = driver.find_element_by_tag_name("input")
search_box.submit()

#get the total number of the pages
page = driver.find_element_by_class_name("paging")
page_link = page.find_elements_by_tag_name("a")
page_num = len(page_link)

noproblem = []
error_list = []
error_name = []

#check the image
for i in range(0, page_num):
    #get the images of current page
    name_list = driver.find_elements_by_tag_name("h5")
    img_class = driver.find_elements_by_class_name("c4")
    img_list = []
    for j in range(len(img_class)):
        #get the link and try whether it can work
        img_list.append(img_class[j].find_element_by_tag_name("img"))
        img_list[j] = img_list[j].get_attribute("src")
        try:
            code = urllib2.urlopen(img_list[j], timeout = 1)
            if code.getcode() is 200:
               noproblem.append(img_list[j])
            else:
                error_list.append(img_list[j])
                error_name.append(name_list[j].text)
        except urllib2.HTTPError, e:
            error_list.append(img_list[j])
            error_name.append(name_list[j].text)
            print( name_list[j].text + "httperror")
        except urllib2.URLError, e:
            print( name_list[j].text + "urlerror")
        except socket.timeout:
            print("time out:" + img_list[j] + "\n" + name_list[j].text)
    page = driver.find_element_by_class_name("paging")
    page_link = page.find_elements_by_tag_name("a")
    page_link[i].click()

#write file
result = open("imgReport.txt", 'w')
result.write("The List for No Images:\n")
if len(error_name) == 0:
    result.write("Every image exists.\n")
else:
    for i in range(len(error_name)):
        result.write(error_name[i].encode("utf8") + "\n")
        result.write(error_list[i].encode("utf8") + "\n")
result.close()




