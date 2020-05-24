import requests
import json
import time,random,os,sys
import openpyxl
import socket
import urllib3
import quopri
import datetime
from bs4 import BeautifulSoup

global CurrDat
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

socket.setdefaulttimeout(20)  # 设置socket层的超时时间为20秒
baidu_url = 'https://map.baidu.com/'
baidu_headers = {
        'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/75.0.3770.100 Safari/537.36'
        
}

def getCurrCity():
    StartLocation =  input('请输入搜索区域：')
    SearchType = input('请输入搜索类型：')
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
        sheet.append([item.name,item.phone,item.addr]) 
    ExcelPath = path + '\\' + CurrDate + SearchParams[2] + SearchParams[3] + '.xlsx'               
    wb.save(ExcelPath)
    print(ExcelPath + '已经处理完成')
    pass

def make_vcf_file(a):
    global data
    data = ''
    for Contact in a:
        
        name = Contact.name
        if Contact.phone == None:
            tel = 'None'
        else:
            tel = '' + Contact.phone
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
                
                DetailName =  CurrDate + JsonDetail['content']['name']
                DetailPhone = JsonDetail['content']['phone']
                DetailAddr = JsonDetail['content']['addr']
                if DetailPhone !=  None and len(DetailPhone) ==11:
                    DetailName = chahao(DetailPhone) + DetailName
                 
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

if __name__ == '__main__':
    
    print('欢迎使用一键通讯录导入功能，导入手机通讯录非常方便！')
  
    SearchParams=getCurrCity()
    print('下面是搜索参数列表')
    print(SearchParams)
    print('\n')
    ContactsList = getPhone(SearchParams[0],SearchParams[1],SearchParams[2],SearchParams[3])
    
    saveExcel(ContactsList,SearchParams[2],SearchParams[3])
    VcfData = make_vcf_file(ContactsList)
    
    VcfPath = path + '\\' + CurrDate + SearchParams[2] + SearchParams[3] + '.vcf' 
    with open(VcfPath,'w') as fileObject:
        fileObject.write(data)  
        fileObject.write('\n') 
        fileObject.close() 
    input('通讯录VCF文件已经保存！')
