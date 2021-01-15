#!/usr/bin/env python
# coding: utf-8

# In[9]:


from selenium import webdriver
from time import sleep
import requests
from lxml import etree
from PIL import Image
from hashlib import md5
import xlrd
import xlwt
from docxtpl import DocxTemplate
from openpyxl import load_workbook

for j in range(1,4):
    driver = webdriver.Chrome()
    driver.maximize_window()

    driver.get('http://jxcp.zjiet.edu.cn/#/')
    sleep(2)
    #标签定位
    search_input = driver.find_element_by_id('username')
    #标签交互
    search_input.send_keys('xxxxxx')
    #标签定位
    search_input = driver.find_element_by_id('password')
    #标签交互
    search_input.send_keys('xxxxxx')
    sleep(2)
    #截图获取验证码图片
    driver.save_screenshot('a.png')
    imgelement = driver.find_element_by_xpath('//*[@id="root"]/div/div/div/div/form/div[3]/span[1]/img')   #定位标签
    location = imgelement.location
    size = imgelement.size
    rangle = (int(location['x']), int(location['y']), int(location['x'] + size['width']),
              int(location['y'] + size['height']))  # 写成我们需要截取的位置坐标
    i = Image.open("a.png")  # 打开截图
    frame4 = i.crop(rangle)  # 使用Image的crop函数，从截图中再次截取我们需要的区域
    frame4.save('save.png') # 保存我们接下来的验证码图片 进行打码
    sleep(1)
    #超级鹰验证码识别
    class Chaojiying_Client(object):

        def __init__(self, username, password, soft_id):
            self.username = username
            password =  password.encode('utf8')
            self.password = md5(password).hexdigest()
            self.soft_id = soft_id
            self.base_params = {
                'user': self.username,
                'pass2': self.password,
                'softid': self.soft_id,
            }
            self.headers = {
                'Connection': 'Keep-Alive',
                'User-Agent': 'Mozilla/4.0 (compatible; MSIE 8.0; Windows NT 5.1; Trident/4.0)',
            }

        def PostPic(self, im, codetype):
            """
            im: 图片字节
            codetype: 题目类型 参考 http://www.chaojiying.com/price.html
            """
            params = {
                'codetype': codetype,
            }
            params.update(self.base_params)
            files = {'userfile': ('ccc.jpg', im)}
            result = requests.post('http://upload.chaojiying.net/Upload/Processing.php', data=params, files=files, headers=self.headers).json()
            return result

        def ReportError(self, im_id):
            """
            im_id:报错题目的图片ID
            """
            params = {
                'id': im_id,
            }
            params.update(self.base_params)
            r = requests.post('http://upload.chaojiying.net/Upload/ReportError.php', data=params, headers=self.headers)
            return r.json()



    chaojiying = Chaojiying_Client('1551969060', 'luoli860826', '909607')	#用户中心>>软件ID 生成一个替换 96001
    im = open('save.png', 'rb').read()													#本地图片文件路径 来替换 a.jpg 有时WIN系统须要//
    result = chaojiying.PostPic(im, 1004)
    print(result)
    print('验证码结果为:' + result['pic_str'])
    sleep(2)
    #标签定位
    search_input = driver.find_element_by_id('kaptcha')
    #标签交互
    search_input.send_keys(result['pic_str'])
    sleep(2)
    #点击登录
    btn = driver.find_element_by_css_selector('.ant-btn')
    btn.click()
    sleep(2)
    #标签到同行评价
    btn = driver.find_element_by_xpath('//*[@id="root"]/div/div[2]/div[1]/div[1]/div/div/ul/li/a')
    btn.click()
    #标签到同行评价
    btn = driver.find_element_by_xpath('//*[@id="root"]/div/div[2]/div[1]/div[1]/div/div/ul/li/ul/li/a')
    btn.click()
    sleep(2)
    #标签部门并填写
    driver.find_element_by_xpath("//div[@id='root']/div/div[3]/div[2]/div/form/div/div/div/div[2]/div/span/span/span/span/span").click()
    driver.find_element_by_xpath("//li/ul/li/span[2]/span").click()
    sleep(1)
    #点击取消清空
    driver.find_element_by_xpath("//div[@id='root']/div/div[3]/div[2]/div/form/div/div/div[4]/div[2]/div/span/div/div/span").click()
    driver.find_element_by_xpath("//div[@id='root']/div/div[3]/div[2]/div/form/div/div[2]/div[4]/div[2]/div/span/div/div/span").click()
    driver.find_element_by_xpath("//div[@id='root']/div/div[3]/div[2]/div/form/div/div[2]/div[3]/div[2]/div/span/div[2]/div/span").click()
    driver.find_element_by_xpath("//div[@id='root']/div/div[3]/div[2]/div/form/div/div[2]/div[3]/div[2]/div/span/div/div/span").click()
    sleep(1)
    #打开文件
    workbook = xlrd.open_workbook("E:\谷歌下载文件\自动化录音填表\自动听课笔记表.xlsx")
    Sheet = workbook.sheet_by_name('Sheet'+ str(j))
    #授课教师、所属系部、课程名称、听课班级列表
    four_list = []
    #听课时间列表
    time_list = []
    #学生状态列表
    student_status_list = []
    #得分列表
    score_list = []
    #小计列表
    subtotal_list = []
    #授课教师、所属系部、课程名称、听课班级
    for four in range(4):
        four_list.append(Sheet.cell_value(four,1))
    print(four_list)
    #听课时间
    for time in range(6):
        if time != 3:
            time_list.append(int(Sheet.cell_value(5,time+1)))
        else:
            time_list.append(Sheet.cell_value(5,time+1))
    print(time_list)
    #听课地点
    class_location = Sheet.cell_value(6,1)
    print(class_location)
    #第几周
    weeks = Sheet.cell_value(6,3)
    print(weeks)
    #教学准备
    teaching_preparation = Sheet.cell_value(7,1)
    print(teaching_preparation)
    #学生状态
    for student_status in range(3):
        if student_status != 2:
            student_status_list.append(int(Sheet.cell_value(9,student_status+1)))
        else:
            student_status_list.append(round(Sheet.cell_value(9,student_status+1)*100))
    print(student_status_list)
    #获取听课记录信息
    lecture_notes = Sheet.cell_value(11,0)
    print(lecture_notes)
    #获取得分
    for score in range(12):
        score_list.append(int(Sheet.cell_value(score+22,5)))
    print(score_list)
    #获得课程标准小计
    subtotal_list.append(int(Sheet.cell_value(22,6)))
    #获得教学态度小计
    subtotal_list.append(int(Sheet.cell_value(23,6)))
    #获取教学内容小计
    subtotal_list.append(int(Sheet.cell_value(26,6)))
    #获得教学组织小计
    subtotal_list.append(int(Sheet.cell_value(29,6)))
    #获得学生学习效果小计
    subtotal_list.append(int(Sheet.cell_value(31,6)))
    print(subtotal_list)
    #获取教师教学评分得分
    all_score = int(Sheet.cell_value(34,6))
    print(all_score)
    #获取评价意见
    evaluation_opinions = Sheet.cell_value(36,0)
    print(evaluation_opinions)
    #获取听课人
    class_one = Sheet.cell_value(42,1)
    print(class_one)
    sleep(2)
    # 加载要填入的数据
    wb = load_workbook(r"E:\谷歌下载文件\自动化录音填表\自动听课笔记表.xlsx")
    ws = wb['Sheet'+ str(j)]
    contexts = [] 
    context = {"teacher_name": four_list[0], "department": four_list[1], "class_name": four_list[2], "listen_class":four_list[3],
              "year":time_list[0],"mouth":time_list[1],"day":time_list[2],"weeky":time_list[3],"num1":time_list[4],"num2":time_list[5],
              "location":class_location,"preparation":teaching_preparation,"student_num1":student_status_list[0],"student_num2":student_status_list[1],
              "student_num3":student_status_list[2],"notes":lecture_notes,"score1":score_list[0],"score2":score_list[1],"score3":score_list[2],
              "score4":score_list[3],"score5":score_list[4],"score6":score_list[5],"score7":score_list[6],"score8":score_list[7],
              "score9":score_list[8],"score10":score_list[9],"score11":score_list[10],"score12":score_list[11],"subtotal1":subtotal_list[0],
              "subtotal2":subtotal_list[1],"subtotal3":subtotal_list[2],"subtotal4":subtotal_list[3],"subtotal5":subtotal_list[4],"all_score":all_score,
              "evaluation_opinions":evaluation_opinions,"class_one":class_one}
    contexts.append(context)
    contexts
    tpl = DocxTemplate(r"E:\谷歌下载文件\自动化录音填表\听课评价表单表2019上半年" + str(j) + ".docx")
    tpl.render(context)
    tpl.save("E:\谷歌下载文件\自动化录音填表\听课评价表单表2019上半年" +str(j) + ".docx")
    sleep(2)
    #获取点击老师名称方法2
    driver.find_element_by_xpath("//div[@id='root']/div/div[3]/div[2]/div/form/div/div/div[2]/div[2]/div/span/div/div/div/div").click()
    sleep(1)
    driver.find_element_by_xpath("//*[@class='ant-select-search__field']").send_keys(four_list[0])
    sleep(2)
    driver.find_element_by_xpath("//*[@id='root']/div/div[3]/div[2]/div[2]/div/div/div/div[2]/div/div/div/div/div/table/tr[2]/td[2]/textarea").click()
    sleep(2)
    #班级点击
    import  requests
    from  lxml  import  etree 
    driver.find_element_by_xpath("//*[@id='root']/div/div[3]/div[2]/div[1]/form/div/div[1]/div[3]/div[2]/div/span/div/div/div/div[1]").click()
    sleep(2)
    page_text = driver.page_source
    selector = etree.HTML(page_text)
    content = selector.xpath('/html/body/div[4]/div/div/div/ul/li')
    li_list = []
    for i in content:
        li = i.xpath("./text()")[0]
        if li != "无匹配结果":
            li_list.append(li)
    print(li_list)
    class_index = li_list.index(four_list[3])+1
    # #点击班级
    # driver.find_element_by_xpath("//div[@id='root']/div/div[3]/div[2]/div/form/div/div/div[3]/div[2]/div/span/div/div/div/div").click()
    #获取第几个班级并点击
    if li_list != []:
        driver.find_elements_by_class_name("ant-select-dropdown-menu-item")[class_index].click()
    sleep(2)
    #课程点击
    import  requests
    from  lxml  import  etree 
    driver.find_element_by_xpath("//*[@id='root']/div/div[3]/div[2]/div[1]/form/div/div[2]/div[1]/div[2]/div/span/div/div/div/div[1]").click()
    sleep(2)
    page_text = driver.page_source
    selector = etree.HTML(page_text)
    content = selector.xpath('/html/body/div[5]/div/div/div/ul/li')
    curriculum_list = []
    for i in content:
        curriculum = i.xpath("./text()")[0]
        curriculum_list.append(curriculum)
    print(curriculum_list)
    curriculum_index = curriculum_list.index(four_list[2])+1+len(li_list)

    # #点击班级
    # driver.find_element_by_xpath("//div[@id='root']/div/div[3]/div[2]/div/form/div/div/div[3]/div[2]/div/span/div/div/div/div").click()
    #获取第几个班级并点击
    driver.find_elements_by_class_name("ant-select-dropdown-menu-item")[curriculum_index].click()
    sleep(2)
    #点击第几周
    import  requests
    from  lxml  import  etree 
    driver.find_element_by_xpath("//*[@id='root']/div/div[3]/div[2]/div[1]/form/div/div[1]/div[4]/div[2]/div/span/div/div/div/div").click()
    sleep(2)
    page_text = driver.page_source
    selector = etree.HTML(page_text)
    content = selector.xpath('/html/body/div[6]/div/div/div/ul/li')
    week_list = []
    for i in content:
        week = i.xpath("./text()")[0]
        week_list.append(week)
    print(week_list)
    week_index = week_list.index(weeks)+1+len(li_list)+len(curriculum_list)
    # #点击班级
    # driver.find_element_by_xpath("//div[@id='root']/div/div[3]/div[2]/div/form/div/div/div[3]/div[2]/div/span/div/div/div/div").click()
    #获取第几个班级并点击
    driver.find_elements_by_class_name("ant-select-dropdown-menu-item")[week_index].click()
    sleep(2)
    #上课地点点击
    import  requests
    from  lxml  import  etree 
    driver.find_element_by_xpath("//*[@id='root']/div/div[3]/div[2]/div[1]/form/div/div[2]/div[2]/div[2]/div/span/div/div/div/div[1]").click()
    sleep(2)
    page_text = driver.page_source
    selector = etree.HTML(page_text)
    content = selector.xpath('/html/body/div[7]/div/div/div/ul/li')
    class_location_list = []
    for i in content:
        if i.xpath("./text()") != []:
            class_locations = i.xpath("./text()")[0]
            class_location_list.append(class_locations)
    print(class_location_list)
    class_location_index = class_location_list.index(class_location)+1+len(li_list)+len(week_list)+len(curriculum_list)
    driver.find_elements_by_class_name("ant-select-dropdown-menu-item")[class_location_index].click()
    sleep(2)
    #开始节次的点击
    import  requests
    from  lxml  import  etree 
    driver.find_element_by_xpath("//*[@id='root']/div/div[3]/div[2]/div[1]/form/div/div[2]/div[3]/div[2]/div/span/div[1]/div/div/div").click()
    sleep(2)
    page_text = driver.page_source
    selector = etree.HTML(page_text)
    content = selector.xpath('/html/body/div[8]/div/div/div/ul/li')
    start_section_list = []
    for i in content:
        start_sections = i.xpath("./text()")[0]
        start_section_list.append(start_sections)
    print(start_section_list)
    start_section_index = start_section_list.index(str(time_list[4]))+2+len(li_list)+len(week_list)+len(curriculum_list)+len(class_location_list)
    #获取第几个班级并点击
    print(start_section_index)
    sleep(2)
    driver.find_elements_by_class_name("ant-select-dropdown-menu-item")[start_section_index].click()
    #结束节次的点击
    driver.find_element_by_xpath("//*[@id='root']/div/div[3]/div[2]/div[1]/form/div/div[2]/div[3]/div[2]/div/span/div[2]/div/div/div").click()
    sleep(2)
    end_section_index = start_section_list.index(str(time_list[5]))+2+len(li_list)+len(week_list)+len(curriculum_list)+len(class_location_list) +  len(start_section_list)
    driver.find_elements_by_class_name("ant-select-dropdown-menu-item")[end_section_index].click()
    sleep(2)
    #点击星期几
    driver.find_element_by_xpath("//*[@id='root']/div/div[3]/div[2]/div[1]/form/div/div[2]/div[4]/div[2]/div/span/div/div/div/div").click()
    sleep(2)
    page_text = driver.page_source
    selector = etree.HTML(page_text)
    content = selector.xpath('/html/body/div[10]/div/div/div/ul/li')
    week_day_list = []
    for i in content:
        week_days = i.xpath("./text()")[0]
        week_day_list.append(week_days)
    print(week_day_list)
    #获取第几个班级并点击
    week_day_index = week_day_list.index('星期' + time_list[3])+2+len(li_list)+len(week_list)+len(curriculum_list)+len(class_location_list) + 2 * len(start_section_list)
    driver.find_elements_by_class_name("ant-select-dropdown-menu-item")[week_day_index].click()
    sleep(2)
    #评测结果得分
    search_input = driver.find_element_by_xpath('//*[@id="root"]/div/div[3]/div[2]/div[2]/div/div/div/div[2]/div/div/div/div/div/table/tr[1]/td[3]/div[2]/input').clear()
    search_input = driver.find_element_by_xpath('//*[@id="root"]/div/div[3]/div[2]/div[2]/div/div/div/div[2]/div/div/div/div/div/table/tr[1]/td[3]/div[2]/input')
    search_input.send_keys(all_score)
    sleep(2)
    #获取同行测评
    btn = driver.find_element_by_xpath('//*[@id="root"]/div/div[3]/div[2]/div[2]/div/div/div/div[2]/div/div/div/div/div/table/tr[2]/td[2]/textarea')
    btn.send_keys(evaluation_opinions)
    sleep(2)
    #预提交点击
    driver.find_element_by_xpath("//*[@id='root']/div/div[3]/div[2]/div[1]/form/div/div[1]/button").click()
    sleep(2)
    #点击预提交确定
    driver.find_element_by_xpath("/html/body/div[12]/div/div[2]/div/div[1]/div/div/div[2]/button").click()
    sleep(2)
    #点击同行评价
    driver.find_element_by_xpath("//*[@id='root']/div/div[2]/div[1]/div[2]/h3/span").click()
    #点击评价明细
    driver.find_element_by_xpath("//*[@id='root']/div/div[2]/div[1]/div[2]/div/div/ul/li/a").click()
    driver.find_element_by_xpath("//*[@id='root']/div/div[2]/div[1]/div[2]/div/div/ul/li/ul/li/a").click()


# In[ ]:




