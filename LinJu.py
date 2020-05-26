import gevent
from gevent import monkey
monkey.patch_all()
from gevent.queue import Queue
import requests
import json
import time,random,os,sys
sys.setrecursionlimit(1000000)
import openpyxl
import socket
import urllib3
import quopri
import datetime
import string
from bs4 import BeautifulSoup

global CurrDat,savetel
CurrTime =  datetime.datetime.now()
CurrDate = str(CurrTime.year)+str(CurrTime.month)+str(CurrTime.day)
ContactsList = []
path = os.path.abspath(os.path.dirname(sys.argv[0]))
    

class Contacts():
    #一个联系人类
    def __init__(self,name,phone,address):
        self.name = name
        self.phone = phone
        self.addr = address

    def setName(self,phone):
    #设置联系人电话
        self.phone = phone

    def setName(self,name):
    #设置联系人姓名
        self.name = name

socket.setdefaulttimeout(30)  # 设置socket层的超时时间为20秒
baidu_url = 'https://map.baidu.com/'
baidu_headers = {
        'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/75.0.3770.100 Safari/537.36'
        
}

def getCurrCity(StartLocation,SearchType):
    
    ParamsSearchCityCode = {

    }
    ParamsStart = {
        'newmap':'1',
        'reqflag':'pcmap',
        'biz':'1',
        'from':'webmap',
        'da_par':'baidu',
        'pcevaname':'pc4.1',
        'qt':'s',
        'da_src':'searchBox.button',
        'wd':StartLocation,
        'c':'142',
        'src':'0',
        'wd2':'',
        'pn':'0',
        'sug':'0',
        'l':'12',
        'b':'(12250687.563613862,2901166.0561386137;12259165.170049507,2908756.033861386)',
        'from':'webmap',
        'biz_forward':'{%22scaler%22:1,%22styles%22:%22pl%22}',
        'sug_forward':'',
        'device_ratio':'1',
        'ie':'utf-8',
        'tn':'B_NORMAL_MAP',
        'nn':'0',
        'auth':'4ZvVzDVyD%3DJa0JTe7KNFeYeS40OJzY%3DSuxHRTEBHNBEtAmk5zC88yy1uVt1GgvPUDZYOYIZuxNtQs3WJ9ILiidiB9APWv3GuExt58Jv7uUvhgMZSguxzBEHLNRTVtcEWe1GD8zvAUu0f7D2LChyBxf0wd0vyISUFAFAOOuuyWWJOfIuLmSfU2K473Fkk0H3L&device_ratio=1&tn=B_NORMAL_MAP&nn=0&u_loc=12278571,2888140&ie=utf-8',
        #'t':'1589435741201',
    }

    try:
        ResStart = requests.get(baidu_url,timeout=(4,6),params = ParamsStart, headers = baidu_headers)

    except KeyError:
        print('key error')
    except requests.exceptions.ReadTimeout:
        print('requests read time error,HttpConnectionPool Error')
    except urllib3.exceptions.ReadTimeoutError:
        print('INFO:connection error!')
    except socket.timeout:
        print('Info: socket time out,read operation timeout!') 
    else:
        time.sleep(random.randint(2,4))
        JsonStart = ResStart.json()
        UidStart = JsonStart['content'][0]['uid']
        CurrCityCode = str(JsonStart['current_city']['code'])
        print('当前搜索地点和城市代码：')
        print(StartLocation + CurrCityCode + '\n')
        ResStart.close()
        time.sleep(2)
        print(UidStart)
    return [UidStart,CurrCityCode,StartLocation,SearchType]

def chahao(SearchPhone):
    UrlChahao = 'https://www.chahaoba.com/index.php'
    ParamsChahao = {
        'title':'%E7%89%B9%E6%AE%8A%3A%E6%90%9C%E7%B4%A2',
        'profile':'default',
        'fulltext':'Search',
        'search':SearchPhone,
    }

    HeadChahao = {
        'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/75.0.3770.100 Safari/537.36'
    }
    ResChahao =  requests.get(UrlChahao, params = ParamsChahao, headers= HeadChahao)
    if(ResChahao.status_code != 200):
        print('Requests error')
    BsChahao = BeautifulSoup(ResChahao.text,'html.parser')
    BsLi = BsChahao.find_all('a',class_ = 'extiw')

    IspProvince = BsLi[1].text
    IspCity = BsLi[2].text
    IspType = BsLi[3].text
    IspInfo = IspProvince+IspCity+IspType
    time.sleep(3)
    return IspInfo
    

def saveExcel(ContactsList,StartLocation,SearchType):
    wb = openpyxl.Workbook()
    sheet = wb.active
    sheet['A1']='名称'
    sheet['B1']='联系方式'
    sheet['C1']='地址'
    for item in ContactsList:
        if item.phone == None:
            continue
        sheet.append([item.name,item.phone,item.addr]) 
    ExcelPath = path + '\\' + CurrDate + StartLocation + SearchType + '.xlsx'               
    wb.save(ExcelPath)
    print(ExcelPath + '已经处理完成')
    pass

def make_vcf_file(a):
    global data
    data = ''
    for Contact in a:
        
        name = Contact.name
        if Contact.phone == None:
            continue
        elif savetel == '1':
            if Contact.phone.count(',') == 1 :
                Tel = Contact.phone.split(',')
                tel1 = Tel[0]
                tel2 = Tel[1]
                tel = '' + tel1 +  '''
TEL;CELL:'''+tel2
            else:
                tel = Contact.phone
        elif savetel != '1':
            if Contact.phone.count(',') ==1:
                Tel = Contact.phone.split(',')
                tel1 = Tel[0]
                tel2 = Tel[1]
                if tel1.count('(') == 1:
                    tel1 = ''
                if tel2.count('(') == 1:
                    tel2 = ''
                tel = '' + tel1 +  '''
TEL;CELL:'''+tel2 
            else:
                if Contact.phone.count('(')==1:
                    continue
                else:
                    tel = Contact.phone
                    
        name = quopri.encodestring(name.encode())
        s = '''BEGIN:VCARD
VERSION:2.1
N;CHARSET=UTF-8;ENCODING=QUOTED-PRINTABLE:;'''+str(name, 'utf-8')+''';;;
FN;CHARSET=UTF-8;ENCODING=QUOTED-PRINTABLE:;'''+str(name, 'utf-8')+'''
TEL;CELL:'''+tel+'''
END:VCARD
'''
        data = data + s
        #print('"' + str(name) + '" : "' + str(tel) + '",')
    return data


def getPhone(UidStart,CurrCityCode,StartLocation,SearchType):
    
    #CityIdDict = {'桂林':'142','南宁':'261'}
    #City = input('请输入搜索城市：')

    for j in range(100):
        pagenumber = j*10
        ParamsNeighbor = {
            'newmap': '1',
            'reqflag': 'pcmap',
            'biz': '1',
            'from': 'webmap',
            'da_par': 'after_baidu',
            'pcevaname': 'pc4.1',
            'qt':'nb',
            'r': '1000',
            'c': CurrCityCode,
            'wd': SearchType,
            'uid': UidStart,
            'b': '(13059415.65,3724579.3;13061679.65,3727895.3)',
            'l': '16',
            'gr_radius': '1000',
            'pn': '0',
            'auth': 'IS2gRCfF5Y1ddM5CIvXfESBycVKdINcVuxHTVzTVRRNtAmk5zC88yy1uVt1GgvPUDZYOYIZuztFexLwWvGccZcuVtPWv3GuHt0A=H73uzRyUbNB9AUvhgMZSguxzBEHLNRTVtcEWe1GD8zvAua@QwBvgw4vjlBhlADMYZEYZSJ5zjnOOADJzEjjg2P',
            'device_ratio': '1',
            'tn': 'B_NORMAL_MAP',
            'nn': pagenumber,
            'u_loc': '12281086,2889132',
            'ie': 'utf-8',
            't': '1590290890571'
            }
        
        try:
            res = requests.get(baidu_url,timeout=(4,6),params=ParamsNeighbor,headers=baidu_headers)
            res.encoding ='gbk'
            if res.status_code != 200:
                print("error")
            pos_json = res.json()
            
        except KeyError:
            print('key error')
        except requests.exceptions.ReadTimeout:
            print('requests read time error,HttpConnectionPool Error')
        except urllib3.exceptions.ReadTimeoutError:
            print('INFO:connection error!')
        except socket.timeout:
            print('Info: socket time out,read operation timeout!') 
        else:
            res.close()
            time.sleep(2)
        
        if(pos_json == None):
            break
        else:
            ListLength = len(pos_json['content'])
        
        for i in range(len(pos_json['content'])-1):
            
            name1 = pos_json['content'][i]['name']
            uid1 = pos_json['content'][i]['uid']
            print(name1)
            print(uid1)

            
            ParamsDetail = {
            'uid':uid1,
            'ugc_type':'3',
            'ugc_ver':'1',
            'qt':'detailConInfo',
            'device_ratio':'1',
            'compat':'1',
            #'t':'1589434352045',
            'auth':'4ZvVzDVyD%3DJa0JTe7KNFeYeS40OJzY%3DSuxHRTEBEBzEtDpnSCE%40%40By1uVt1GgvPUDZYOYIZuxHtPqIVH82LiidiB9APWv3GuBHt9iTHf2UvhgMZSguxzBEHLNRTVtcEWe1GD8zvAOufdUBFFAexZFTHrwzBvpFcEegvcguxHRTEzEzNNtuyWWJ49Iydd%3DB11'
            }
            try:
                print('ResDetail ok?')
                ResDetail = requests.get(baidu_url, timeout=(4,6),params = ParamsDetail, headers = baidu_headers)
                print(ResDetail.status_code)
                if (ResDetail.status_code != 200):
                    continue
                JsonDetail =  ResDetail.json()
                
                DetailName =  CurrDate + StartLocation + JsonDetail['content']['name']
                DetailPhone = JsonDetail['content']['phone']
                DetailAddr = JsonDetail['content']['addr']
                if DetailPhone !=  None and len(DetailPhone) ==11:
                    DetailName = chahao(DetailPhone) + DetailName
                elif DetailPhone !=  None and len(DetailPhone) == 23:
                    Tel = DetailPhone.split(',')
                    tel1 = Tel[0]
                    tel2 = Tel[1]
                    DetailName = chahao(tel1) + chahao(tel2) + DetailName
                elif DetailPhone !=  None and DetailPhone.count('(') == 1 and DetailPhone.count(',') == 1:
                    Tel = DetailPhone.split(',')
                    tel1 = Tel[0]
                    tel2 = Tel[1]
                    if tel1[0] == '(':
                        DetailName = chahao(tel2) + DetailName
                    elif tel2[0] == '(':
                        DetailName = chahao(tel1) +DetailName
                Contact =  Contacts(DetailName,DetailPhone,DetailAddr)
                ContactsList.append(Contact)
                print(DetailPhone)
                
            except KeyError:
                print('key error')
                pass
            except requests.exceptions.ReadTimeout:
                print('requests read time error,HttpConnectionPool Error')
            except urllib3.exceptions.ReadTimeoutError:
                print('INFO:connection error!')
            except socket.timeout:
                print('Info: socket time out,read operation timeout!') 
            else:
                ResDetail.close()
                time.sleep(random.randint(2,4))
                
        if (ListLength < 11):
            break 
        
    return ContactsList

def OneTask(StartLocation,SearchType):

    #根据输入的地点和信息类型，输出第一个地点的UID，城市代码，搜索地域，搜索类型
    SearchParams=getCurrCity(StartLocation,SearchType)
    print('下面是搜索参数列表')
    print(SearchParams)
    print('\n')
    #ContactSList是根据输入的一个地点和类型输出所有相关的电话信息。
    ContactsList = getPhone(SearchParams[0],SearchParams[1],SearchParams[2],SearchParams[3])
    
    
if __name__ == '__main__':
    print('欢迎使用一键通讯录导入功能，导入手机通讯录非常方便！')
 
    SearchParams = []
    while(True):
        StartLocation =  input('请输入搜索区域,键入stop停止输入：')
        if(StartLocation == 'stop' ):
            break
        SearchType = input('请输入搜索类型，键入stop停止输入：')
        if(SearchType == 'stop'):
            break
        SearchParams.append([StartLocation,SearchType])
    savetel = input('输入1保存固定电话号码到通讯录，输入其他不保存：')
    work = Queue()
    for item in SearchParams:
        work.put_nowait(item)

    def crawler():
        while not work.empty():
            SearchParam = work.get_nowait()
            OneTask(SearchParam[0],SearchParam[1])
    
    TaskLists = []
    for x in range(2):
        task =  gevent.spawn(crawler)
        TaskLists.append(task)
    gevent.joinall(TaskLists)

    saveExcel(ContactsList,SearchParams[0][0],SearchParams[0][1])
    VcfData = make_vcf_file(ContactsList)
    
    VcfPath = path + '\\' + CurrDate + SearchParams[0][0] + SearchParams[0][1] + '.vcf' 
    with open(VcfPath,'w') as fileObject:
        fileObject.write(data)  
        fileObject.write('\n') 
        fileObject.close() 
    print('一个通讯录VCF文件任务已经保存！')

    