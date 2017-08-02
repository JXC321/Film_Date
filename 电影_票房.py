import requests
import re
from xlwt import Workbook
import xlrd
import time

R = requests.Session()
Url = 'http://58921.com'

def log_in():
    #模拟登录
    cookies = {}
    url = 'http://58921.com/user/login'
    raw_cookies = 'remember=MTEzNTI2LjIxNjM0Mi4xMDI4MTYuMTA3MTAwLjExMTM4NC4yMDc3NzQuMTE5OTUyLjExMTM4NC4xMDI4MTYuMA%3D%3D; DIDA642a4585eb3d6e32fdaa37b44468fb6c=msl4r0g46n4h70r5cviefs2672; time=MTEzNTI2LjIxNjM0Mi4xMDI4MTYuMTA3MTAwLjExMTM4NC4yMDc3NzQuMTE5OTUyLjExMTM4NC4xMDQ5NTguMTExMzg0LjEyMjA5NC4xMTc4MTAuMTIyMDk0LjExNzgxMC4xMDQ5NTguMTA5MjQyLjExNzgxMC4xMTEzODQuMA%3D%3D; Hm_lvt_e71d0b417f75981e161a94970becbb1b=1497971126; Hm_lpvt_e71d0b417f75981e161a94970becbb1b=1497971375'
    for lies in raw_cookies.split(';'):
        key,word = lies.split('=',1)
        cookies[key] = word
    
    res = R.post(url,cookies = cookies)

def get_page(year,number):
    #抓取索引页面的信息，病返回页面的html文本
    url = 'http://58921.com/alltime' + '/' + str(year)
    if(number == 0):
        pass
    else:
        url = url +'?page='+str(number)
    #print(url)    
    res = R.get(url)
    res.encoding = 'utf-8'
    return res.text
    

def get_number(Text):
    reg = r'<li class="pager_last"><a href="/alltime/.+?\?page=(.+?)"'
    reg = re.compile(reg)
    number = re.findall(reg,Text)
    return int(number[0])
    
def get_info(Text,sheet,k=1):
    #页面从文本中解析出电影名，并找到电影简介的链接，得到电影各种信息
    Soup = BeautifulSoup(Text,'html5lib')
    A = Soup.select('.odd')
    B = Soup.select('.even')
    
    #C储存一个页面的电影的排名
    C = []
    for i in range(len(A)):
        C.append((A[i]))
        try:
            C.append((B[i]))
        except:
            pass
    
    #遍历每个电影，获取电影的详情：导演、主演、上映时间、制片国家、电影时长、电影类型、票房
    for c in C:
        url = Url + c.a['href']
        try:
            name = c.a['title']            
            print(name,url)
        except:
            name = ''
            print('获取电影名失败！')
            continue
        sheet.write(k,0,name)
        
        m = R.get(url)
        m.encoding = 'utf-8'
        
        try:
            reg = r'>导演：<.+?title="(.+?)"'
            reg = re.compile(reg)
            Dir = re.findall(reg,m.text)
            Dir = Dir[0]
            print(Dir)
        except:
            Dir = ''
            print('获取导演数据失败')
        sheet.write(k,1,Dir)
        
        
        try:
            reg = r'>主演：<.+?title="(.+?)"'
            reg = re.compile(reg)
            Actor = re.findall(reg,m.text)
            Actor = Actor[0]
            print(Actor)
        except:
            Actor = ''
            print('获取主演数据失败')
        sheet.write(k,2,Actor)  
        
        try:
            reg = r'>上映时间：<.+?>(.+?)<'
            reg = re.compile(reg)
            Time = re.findall(reg,m.text)
            Time = Time[0]
            print(Time)
        except:
            Time = ''
            print('获取上映时间失败')
        sheet.write(k,3,Time) 
        
        
        try:
            reg = r'>片长：<.+?>(.+?)<'
            reg = re.compile(reg)
            Long = re.findall(reg,m.text)
            Long = Long[0]
            print(Long)
        except:
            Long = ''
            print('获取片长间失败')
        sheet.write(k,4,Long)
        
        try:
            reg = r'>制作国家/地区：<.+?title="(.+?)"'
            reg = re.compile(reg)
            Country = re.findall(reg,m.text)
            Country = Country[0]
            print(Country)
        except:
            COuntry = ''
            print('获取制作国家失败')
        sheet.write(k,5,Country)
        
        
        try:
            reg = r'>类型：<.+?title="(.+?)"'
            reg = re.compile(reg)
            Type = re.findall(reg,m.text)
            Type = Type[0]
            print(Type)
        except:
            Long = ''
            print('获取类型失败')
        sheet.write(k,6,Type)
        
        #获取电影的票房数据
        url_2 = url + '/boxoffice'
        res = R.get(url_2)
        res.encoding = 'utf-8'
        
        reg = r'\(最新票房 (.+?)\)'
        reg = re.compile(reg)
        money = re.findall(reg,res.text)
        print(money[0])
        sheet.write(k,7,money)
        k = k + 1
        #print(res.text)
        print('\n')
        
    return k
        
        
def main():
    book = Workbook()
    log_in()
    k = 1
    for i in range(2012,2016+1):
        
        Text = get_page(i,0)
       
        sheet = book.add_sheet(str(i) + '年的电影票房')
        sheet.write(0,0,'电影名')
        sheet.write(0,1,'导演')
        sheet.write(0,2,'主演')
        sheet.write(0,3,'上映时间')
        sheet.write(0,4,'片长')
        sheet.write(0,5,'制片国家')
        sheet.write(0,6,'类型')
        sheet.write(0,7,'票房')

        number = get_number(Text)
        
        k = get_info(Text,sheet)
        book.save('电影数据.xls')
        #break
        for j in range(1,number+1):
            Text = get_page(i,j)
            k = get_info(Text,sheet,k)
            
        time.sleep(3)
        book.save('电影数据.xls')
        #break
if __name__ == '__main__':
    main()
    
