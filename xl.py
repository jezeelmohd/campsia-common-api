import xlwt
import requests
import json
import ast
#r = requests.get("http://gdata.youtube.com/feeds/api/standardfeeds/top_rated?v=2&alt=jsonc")
#r.text


#false='false'
#x={"name":"Zulu Mens Hostel","address":"Zulu Mens Hostel, MH 3\nLBS College of Engineering\nPovval, Kasaragod","phone":"0458-36968755","number_of_rooms":"2","rooms":[{"number":"A-1","capacity":"2","rent":"800","editEnabled":false},{"number":"A-2","capacity":"2","rent":"800","editEnabled":false},{"number":"A-3","capacity":"2","rent":"800","editEnabled":false},{"number":"A-4","capacity":"2","rent":"800","editEnabled":false},{"number":"A-5","capacity":"2","rent":"800","editEnabled":false},{"number":"A-6","capacity":"2","rent":"800","editEnabled":false},{"number":"A-7","capacity":"2","rent":"800","editEnabled":false},{"number":"A-8","capacity":"2","rent":"800","editEnabled":false},{"number":"A-9","capacity":"2","rent":"800","editEnabled":false},{"number":"A-10","capacity":"2","rent":"800","editEnabled":false},{"number":"B-1","capacity":"1","rent":"1000","editEnabled":false},{"number":"B-2","capacity":"1","rent":"1000","editEnabled":false},{"number":"B-3","capacity":"1","rent":"1000","editEnabled":false},{"number":"B-4","capacity":"1","rent":"1000","editEnabled":false},{"number":"B-5","capacity":"1","rent":"1000","editEnabled":false},{"number":"B-6","capacity":"1","rent":"1000","editEnabled":false},{"number":"B-7","capacity":"1","rent":"1000","editEnabled":false},{"number":"B-8","capacity":"1","rent":"1000","editEnabled":false},{"number":"B-9","capacity":"1","rent":"1000","editEnabled":false},{"number":"B-10","capacity":"1","rent":"1000","editEnabled":false}]}
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
#print data['titlelist']


#for dat in data:
    
#   print columnlist
#for dat in data:#  print dat
    #for key in dat.items():
    #   print 'Key=',key #,'value=',value

#print data['name'],data['address'],data['phone'],data['number_of_rooms']
dataa=data[0]
titlelist=dataa['titlelist']
columnlist=dataa['columnlist']
print columnlist


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