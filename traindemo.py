from xml.etree import ElementTree as et
import requests
import openpyxl

# 创建一个Excel
wb = openpyxl.Workbook()
sheet = wb.active
sheet.title = 'Train'
sheet["A1"].value = '车次'
sheet["B1"].value = '发车站'
sheet["C1"].value = '发车时间'
sheet["D1"].value = '到达站'
sheet["E1"].value = '到达时间'
# 获取深圳至上海的火车时刻表
r = requests.get('http://www.webxml.com.cn/WebServices/TrainTimeWebService.asmx/getStationAndTimeByStationName?StartStation=深圳&ArriveStation=上海&UserID=')
result = r.text
# 解析接口返回的xml
root = et.XML(result)
i = 1
for node in root.iter('TimeTable'):
    i = i + 1
    sheet["A%d" % (i)].value = node.find('TrainCode').text
    sheet["B%d" % (i)].value = node.find('StartStation').text
    sheet["C%d" % (i)].value = node.find('StartTime').text
    sheet["D%d" % (i)].value = node.find('ArriveStation').text
    sheet["E%d" % (i)].value = node.find('ArriveTime').text
    train = '车次:' + node.find('TrainCode').text
    start = '发车站:' + node.find('StartStation').text
    startt = '发车时间:' + node.find('StartTime').text
    arrive = '到达站:' + node.find('ArriveStation').text
    arrivet = '到达时间:' + node.find('ArriveTime').text
    print(train, '----', start, '----', startt, '----', arrive, '----', arrivet)
wb.save('traintime.xlsx')