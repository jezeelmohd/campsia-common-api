import xlwt
import requests
import json
import ast
#r = requests.get("http://gdata.youtube.com/feeds/api/standardfeeds/top_rated?v=2&alt=jsonc")
#r.text
txt=[
        {
            'titlelist':['name','id','address','phone'],

            'name': 'Zulu',
            'id': 'LBS Property',
            'address': 'MH 1, LBS College of Engineering, Kasaragod',
            'phone': '0485-1234897',

            'columnlist':['rooms'],
            'roomsattr':['number','capacity','rent','avail'],
            
            'rooms': [
                {
                    'number': 'A1',
                    'capacity': 5,
                    'rent': 2000,
                    'avail':'No'
                },
                {
                    'number': 'A2',
                    'capacity': 6,
                    'rent': 5000,
                    'avail':'Yes'
                },
                {
                    'number': 'A3',
                    'capacity':7,
                    'rent': 8000,
                    'avail':'Yes'
                }
            ]
        }
    ]

json_string = json.dumps(txt)
data = json.loads(json_string)

book = xlwt.Workbook(encoding="utf-8")

sheet1 = book.add_sheet("Sheet 1")

for dat in data:
    titlelist=dat['titlelist']
    columnlist=dat['columnlist']
    #write titles
    i=0
    for n in titlelist:
        sheet1.write(i, 0, dat[n])
        i = i+1

    i=i+1
    #
    for column in columnlist:#taking each columntitle
        j=0
        print column
        sheet1.write(i, 0, column)#writng each column title
        i=i+1
        #accessing each columns
        colattr=column+"attr"
        k=0#the pointer to listing column0 elements

        
        for catt in dat[colattr]:
            locali=i
            sheet1.write(locali, k, catt)
            locali=locali+1
            for elemnt in dat[column]:
                #print elemnt[catt]
                sheet1.write(locali, k, elemnt[catt])
                locali=locali+1
            k=k+1


book.save("Hostel1.xls")