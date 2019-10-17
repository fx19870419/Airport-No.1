import os,openpyxl,datetime,time
from selenium import webdriver
from selenium.webdriver.support.select import Select

#用Excel维护简称与许可证号的对应字典及用户名
wh=openpyxl.load_workbook('用户名密码单位名称许可证号维护.xlsx')
whst=wh['Sheet1']
username=whst['B1'].value
password=whst['B2'].value
shopNN={}
for i in range(4,whst.max_row+1):
  shopNN[whst.cell(i,1).value]=whst.cell(i,2).value

#FileName=input("请输入要导入的xlsx文件名称：")
FileName="食品销售_20190930155211.xlsx"

#判断卫生监督类型赋值给typ_jd和typ_list变量并确定各项的分值
if '餐饮' in FileName:
  typ_jd='餐饮服务'
  typ_list=['※','※','※','2','5','※','2','2','2','2','※','※','※','2','2','1','5','1','2','5','5','2','※','5','10','5','5','1','2','5','2','5','2','※','5','2','※','2','5','2','2','5','5','5','2','※','2','2','2','2','2','1','2','5','2','5','2','2','2','2','2','2','2','※','2','2','5','5','※','2','5','5','5','2']
elif '生产' in FileName:
  typ_jd='食品生产'
  typ_list=['※','※','※','2','5','※','5','5','2','2','※','※','※','※','5','5','5','2','※','※','5','5','2','2','5','2','2','5','10','5','※','※','5','10','10','5','※','5','10','※','5','5','5','2','2','※','2','2','2','5','10','5','※','※','5','5','5']
elif '饮用水' in FileName:
  typ_jd='饮用水供应'
  typ_list=['※','※','※','5','10','5','5','5','5','5','※','※','5','5','5','※','5','2','5','5','2','5','5','10','10','5','※','2','※','5','5','※','2','2','10','2']
elif '住宿' in FileName:
  typ_jd='住宿业'
  typ_list=['※','※','※','5','10','5','5','2','※','5','10','3','5','3','※','2','10','※','10','10','5','5','※','5','5','5','2','3','3','3','3','3','2','3','5']
elif '候机' in FileName:
  typ_jd='候机'
  typ_list=['※','※','※','5','10','5','5','2','※','5','10','3','5','3','※','2','10','※','10','10','5','5','※','5','5','5','5','※','10']
elif '销售' in FileName:
  typ_jd='食品销售'
  typ_list=['※','※','※','5','5','5','5','5','5','※','※','10','10','5','5','5','※','5','5','2','5','2','2','※','5','5','10','5','5','5','5','10','5']

#判断符合或者不符合的函数:
def trueorfalse(x,y):
  if y=="符合":
    return typ_list[int(x/2-1)]
  if y=="合理缺项":
    return '99'
  if y=="不符合":
    return '0'

#填入结果的函数:
def result(score,i):
  if score!='0':
    el_score=browser.find_element_by_css_selector(name_score+value_score)
    browser.execute_script("arguments[0].scrollIntoView();",el_score)
    browser.execute_script("arguments[0].click();",el_score)
  else:
    el_score=browser.find_element_by_css_selector(name_score+value_score)
    browser.execute_script("arguments[0].scrollIntoView();",el_score)
    browser.execute_script("arguments[0].click();",el_score)
    el_explain=browser.find_element_by_css_selector(input_score)
    el_explain.send_keys(list[i+1])

#加载文件
wb=openpyxl.load_workbook(FileName)
sheet=wb['Sheet1']

#开浏览器、打开网页
browser = webdriver.Firefox()
browser.get('http://web3.prosas.hg.cn:8080/prosas/')

#登录账号密码
while 1:
  try:
    el_username=browser.find_element_by_id('username')
    el_username.send_keys(username)#输入用户名
    print('输入账号………………成功')
    el_password=browser.find_element_by_id('password')
    el_password.send_keys(password)#输入密码
    print('输入密码………………成功')
    submit=browser.find_element_by_name('submit')
    submit.click()#登录按钮
    print('登录……………………成功')
    break
  except:
    print('登录失败，请检查网络')
#找到监督评分→点击
el_ywjg=browser.find_element_by_id('heTab105')
el_ywjg.click()#点击“卫生监督”按钮
el_rcwsjd=browser.find_element_by_partial_link_text('日常卫生监督')
el_rcwsjd.click()#点击日常卫生监督按钮
el_jdpf=browser.find_element_by_partial_link_text('监督评分')
el_jdpf.click()#点击监督评分按钮




#将一次卫生监督结果存入list变量，然后将变量写入浏览器表单
for r in range(2,sheet.max_row+1):
  list=[]
  for c in range(1,sheet.max_column):
    list.append(sheet.cell(r,c).value)
    
  time.sleep(15)
  browser.switch_to.default_content()
  el_frame=browser.find_element_by_class_name('iframeClass')
  browser.switch_to.frame(el_frame)
  el_No=browser.find_element_by_name('cardNo')
  el_No.clear()
  el_No.send_keys(shopNN[list[1]])
  el_startDate=browser.find_element_by_name('startDate')
  sDate=datetime.datetime.now()-datetime.timedelta(days=200)#起始日期（当前时间往前推200天）
  browser.execute_script('arguments[0].removeAttribute(\"readonly\")',el_startDate)
  el_startDate.clear()
  el_startDate.send_keys(str(sDate.year)+'-'+str(sDate.month)+'-'+str(sDate.day))#输入起始日期（当前时间往前推200天）
  el_submit=browser.find_element_by_xpath("//input[@value='查询']")
  el_submit.click()
  time.sleep(5)
  el_add=browser.find_element_by_xpath("//i[@title='监督打分']")
  el_add.click()
  el_type=browser.find_element_by_id('itemCode')
  Select(el_type).select_by_visible_text(typ_jd)
  el_typechange=browser.find_element_by_xpath("//a[contains(text(),'修改监督评分表类型')]")
  el_typechange.click()
  browser.switch_to.default_content()
  el_frame_type=browser.find_element_by_css_selector("[src='/prosas/dailySup/listNoQuery.html?menuId=8B4C90F4861945B59DD330DA2378B103']")
  browser.switch_to.frame(el_frame_type)
  #el_typesubmit=browser.find_element_by_css_selector("button")
  el_typesubmit=browser.find_element_by_css_selector("button[class='aui_state_highlight'][type='button']")
  #el_typesubmit.click()
  browser.execute_script("$(arguments[0]).click()",el_typesubmit)
  
  
  data = {
     "餐饮服务":[9,23,31,35, 45, 53, 63, 67, 73, 91, 111, 123, 137, 143, 147, 149],
     "食品生产":[9,21,23,47,57,79,85,95,101,105,115],
     '饮用水供应':[9,21,23,25,33,37,43,63,65,75],
     '住宿业':[9,19,27,33,37,41,45,47,51,53,71],
     '候机':[9,19,27,33,37,41,45,51,53,77],
     '食品销售':[9,19,21,39,47,69]
  }
  for i in range(2,len(list),2):
    i=int(i)
    time.sleep(0.3)
    score=trueorfalse(i,list[i])
    score_list = data[typ_jd]
    
    if 0 < i <= score_list[0]:
        name_score="[name='score01"+(str(int(i/2)).rjust(2,'0'))+"']"
        value_score="[value='"+trueorfalse(i,list[i])+"']"
        input_score="[name='input01"+(str(int(i/2)).rjust(2,'0'))+"']"
        result(score,i)
    elif i <= score_list[-1]:
        pos = [i <= score for score in score_list].index(True)
        subscract = score_list[pos - 1] - 1
        num_str = str(pos + 1).zfill(2)
        name_score="[name='score" + num_str +(str(int((i-subscract)/2)).rjust(2,'0'))+"']"
        value_score="[value='"+trueorfalse(i,list[i])+"']"
        input_score="[name='input" + num_str +(str(int((i-subscract)/2)).rjust(2,'0'))+"']"
        result(score,i)
    else:
        pass


  el_sum=browser.find_element_by_xpath("//a[contains(text(),'计算结果')]")
  el_sum.click()
  #填入监督日期
  el_supdate=browser.find_element_by_name('supScores.supDate')
  browser.execute_script('arguments[0].removeAttribute(\"readonly\")',el_supdate)
  el_supdate.clear()
  #el_supdate.send_keys(list[152][0:4]+'-'+list[152][5:7]+'-'+list[152][8:10]+' '+) #输入监督日期（当前时间往前推200天）
  if typ_jd=='餐饮服务':
    el_supdate.send_keys(list[152])
  if typ_jd=='食品生产':
    el_supdate.send_keys(list[118])
  elif typ_jd=='饮用水供应':
    el_supdate.send_keys(list[78])
  elif typ_jd=='住宿业':
    el_supdate.send_keys(list[74])
  elif typ_jd=='候机':
    el_supdate.send_keys(list[80])
  elif typ_jd=='食品销售':
    el_supdate.send_keys(list[70])
  el_pfjgclick=browser.find_element_by_xpath("//label[contains(text(),'评分结果')]")
  el_pfjgclick.click()
  el_save=browser.find_element_by_xpath("//button[contains(text(),'保存')]")
  el_save.click()
  el_sumbit_2=browser.find_element_by_xpath("//a[contains(text(),'确定')]")
  el_sumbit_2.click()
