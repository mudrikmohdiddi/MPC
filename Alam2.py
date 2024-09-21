from requests import *
from time import *
from openpyxl import *
from winsound import *
n=0
while(n==0):
    try:
        book=load_workbook('store.xlsx')
        n=1
    except FileNotFoundError:
        print("Please open main file for your first time")
        sleep(2)
book=load_workbook('store.xlsx')
sheet=book.active
while(sheet['AL2'].value==None):
    print("Please open main file")
    sleep(2)
    sheet['AL2']=sheet['AL1'].value
    book.save('store.xlsx')
while(sheet['NM1'].value==None):
    print("Please open main file and fill your name")
    sleep(2)
print(f"{sheet['NM1'].value}\nN.B: Message is only 8 word")
while(True):
    try:
        while(sheet['RN1'].value!=sheet['AL1'].value):
            print("Please wait...")
        if(sheet['AL2'].value!=sheet['AL1'].value):
            sheet['AL2']=sheet['AL1'].value
            book.save('store.xlsx')

        def receive(ids):
            t1=['q','w','e','r','t','y','u','i','o','p','a','s','d','f','g','h','j','k','l','z','x','c','v','b','n','m']
            dt1={}
            n=0
            for m in range(26):
                if(m<10):
                    n=30+m
                    dt1.update({n:t1[m]})
                else:
                    n=m
                    dt1.update({n:t1[m]})
            #recever
            ids=ids-1
            word2=[]
            receve = get(f'https://api.thingspeak.com/channels/{sheet['AP3'].value}/feeds.json?api_key={sheet['AP2'].value}&results')
            data = receve.json()
            value1 = int(data['feeds'][ids]['field1'])
            if(value1!=1000):
                word2.insert(0,str(value1))
                t=data['feeds'][ids]['created_at']
                t=str(t)
                w=f"{t[t.index("T")+1]}{t[t.index("T")+2]}"
                w=int(w)+3
                if(w==24):
                    w='(+1)T 00'
                elif(w==25):
                    w='(+1)T 01'
                elif(w==26):
                    w='(+1)T 02'
                else:
                    w=f'T {w}'
                h=f"T{t[t.index("T")+1]}{t[t.index("T")+2]}"
                t=t.replace(h,w)
            value2 = int(data['feeds'][ids]['field2'])
            if(value2!=1000):
                word2.insert(1,str(value2))
            value3 = int(data['feeds'][ids]['field3'])
            if(value3!=1000):
                word2.insert(2,str(value3))
            value4 = int(data['feeds'][ids]['field4'])
            if(value4!=1000):
                word2.insert(3,str(value4))
            value5 = int(data['feeds'][ids]['field5'])
            if(value5!=1000):
                word2.insert(4,str(value5))
            value6 = int(data['feeds'][ids]['field6'])
            if(value6!=1000):
                word2.insert(5,str(value6))
            value7 = int(data['feeds'][ids]['field7'])
            if(value7!=1000):
                word2.insert(6,str(value7))
            value8 = int(data['feeds'][ids]['field8'])
            if(value8!=1000):
                word2.insert(7,str(value8))
            list=[]
            p=''
            for m in word2:
                no=0
                for n in m:
                    no+=1
                    p+=n
                    if(no==2):
                        list.append(p)
                        p=''
                        no=0
                list.append('and')
            if(len(list)!=0 and sheet['MMM3'].value!=t):
                list.pop()
                n_w=''
                for m in list:
                    if(m=='and'):
                        n_w+=' '
                    else:
                        n_w+=dt1[int(m)]
                sheet['MMM3']=t
                book.save('store.xlsx')
                Beep(500,2000)
                
            else:
                return ' '
        def control():
            receve = get(f'https://api.thingspeak.com/channels/{sheet['AP3'].value}/feeds.json?api_key={sheet['AP2'].value}&results')
            data = receve.json()
            if(sheet['MMM4'].value==None):
                last=int(data['channel']['last_entry_id'])
                sheet['MMM4']=last
                book.save('store.xlsx')
            last=int(data['channel']['last_entry_id'])
            begin=int(sheet['MMM4'].value)
            if(last==0):
                print()
            elif(last!=begin):
                for m in range(begin,last+1):
                    print(receive(m))
                begin=last
                sheet['MMM4']=begin
                book.save('store.xlsx')
            else:
                print()
        control()
    except ConnectionError:
        print("Please connect internet")
        sleep(2)

