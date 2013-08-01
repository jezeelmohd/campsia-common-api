import xlwt
import requests
import json
import ast
#r = requests.get("http://gdata.youtube.com/feeds/api/standardfeeds/top_rated?v=2&alt=jsonc")
#r.text
spacebtweencolumns=1
titlecolumn=5
txt=[
        {
            'titlelist':['name','id','address','phone'],

            'name': 'Zulu',
            'id': 'LBS Property',
            'address': 'MH 1, LBS College of Engineering, Kasaragod',
            'phone': '0485-1234897',

            'columnlist':['rooms','buk'],
            'roomsattr':['number','capacity','rent','avail'],
            'bukattr':['name','age'],
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
            ],
            'buk':[{'name':'jez','age':14}]

        },
        {
            'titlelist':['name','id','address','phone'],

            'name': 'Zulu',
            'id': 'LBS Property',
            'address': 'MH 1, LBS College of Engineering, Kasaragod',
            'phone': '0485-1234897',

            'columnlist':['rooms','buk'],
            'roomsattr':['number','capacity','rent','avail'],
            'bukattr':['name','age'],
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
            ],
            'buk':[{'name':'jasim','age':23}]

        }
    ]

json_string = json.dumps(txt)
data = json.loads(json_string)
"""
book = xlwt.Workbook(encoding="utf-8")

sheet1 = book.add_sheet("Sheet 1")
sheet1.row(0).height = 256 * (len(key))
"""

styletitle=xlwt.easyxf('font: bold 1,height 350,color red')
styletitle2=xlwt.easyxf('font: bold 1,height 210')
stylecheader=xlwt.easyxf('font: bold 1,height 250,color blue')
stylerheader=xlwt.easyxf('font: bold 1,height 220,color blue')
stylectext=xlwt.easyxf('font: bold 1,height 200')



if data:       
    fileno=1
    for dat in data:
        
        book = xlwt.Workbook(encoding="utf-8")
        sheet1 = book.add_sheet("Sheet 1")
        key='Zu'
        sheet1.row(0).height = 256 * (len(key))

        titlelist=dat['titlelist']
        columnlist=dat['columnlist']
        #write titles
        i=0
        locali=0
        if titlelist:
            for n in titlelist:
                if i==0:
                    sheet1.write(i,titlecolumn, dat[n],styletitle)
                else:
                    sheet1.write(i,titlecolumn-1, dat[n],styletitle2)
                i = i+1

        i=i+1
        #
        if columnlist:
            for column in columnlist:#taking each columntitle
                j=0
                print column,i
                sheet1.write(i, 0, column,stylecheader)#writng each column title
                i=i+1
                #accessing each columns
                colattr=column+"attr"
                k=0#the pointer to listing column0 elements

                locali_adder=0
                for catt in dat[colattr]:
                    locali=i
                    print catt
                    sheet1.write(locali, k, catt,stylerheader)
                    locali=locali+1
                    locali_adder=locali_adder+1
                    for elemnt in dat[column]:
                        print elemnt[catt]
                        sheet1.write(locali, k, elemnt[catt],stylectext)
                        locali=locali+1
                    k=k+1
                print i,locali,locali_adder

                i=i+locali_adder+spacebtweencolumns

        filenam="Hostel"+str(fileno)+".xls"
        #print '%%%%%%%%%%%%%%%%%%%',filenam
        book.save(filenam)
        fileno=fileno+1