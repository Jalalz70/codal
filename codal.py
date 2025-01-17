import  json
import os
import openpyxl
import  requests
NumberOfPages = 2
howmanyreport =2
distance = 10
addresslist = []
i=0
namad = "وبملت"
address =  f'https://search.codal.ir/api/search/v2/q?&Audited=true&AuditorRef=-1&Category=3&Childs=false&CompanyState=0&CompanyType=-1&Consolidatable=true&IsNotAudited=false&Isic=272006&Length=-1&LetterType=-1&Mains=true&NotAudited=true&NotConsolidatable=true&PageNumber=1&Publisher=false&Symbol={namad}&TracingNo=-1&search=true'
header = {'User-Agent' : 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36'}
URLs =[]
fileexcel = openpyxl.Workbook()
for a in range(NumberOfPages):
    first = (address.split('PageNumber=1')[0:])
    firstmain = str(first[0])
    end = (address.split('PageNumber=1')[1:])
    endmaim = str(end[0])
    number = str(a+1)
    pagenumber = "PageNumber="
    mid = pagenumber + number
    addresss = firstmain + mid + endmaim
    addresslist.append(addresss)
k =0
for c in range(NumberOfPages):
    darkhast = requests.get(url = addresslist[k],headers = header).text.split('SuperVision":{')[1:]
    k = k+1
    for a in darkhast :
        AdressAgahi = a.split('Url":"/Reports/Decision.')[1].split('"')[0]
        i = i+1
        MyUrl = f'https://codal.ir/Reports/Decision.{AdressAgahi}'
        URLs.append(MyUrl)
#################################################################################
###################################EXEL PART#####################################
#################################################################################
ii =0
for aa in URLs [:howmanyreport]:
    ii = ii+1
    print(ii,": ",aa)
    address = aa
    darkhast = requests.get(url=address, headers=header).text
    darkhast = darkhast.split('datasource = ')[1].splitlines()[0][:-1]
    #print("darkhast: ",darkhast)
    jasondarkhast = json.loads(darkhast)['sheets'][0]
    titrefarsi = jasondarkhast['title_Fa']

    jadval = jasondarkhast['tables']

    fileexcel.create_sheet(str(ii))
    loop = 0
    lastaddress = 0
    for a in jadval:
        #print("1111111111111111111111111111111111111111111111111111111111111111111111111")
        loop = loop + 1
        #print("loop: ", loop)

        titrejadval = a['title_Fa']
        #print(titrejadval)

        if (loop == 1):
            celolha = a['cells']
            for b in celolha:
                adres = b['address']
                stringpart = adres.rstrip('0123456789')
                numberpart = tail = adres[len(stringpart):]
                if (int(numberpart) > int(lastaddress)):
                    lastaddress = numberpart
                value = b['value']
                #print(f'{adres} = {value}')

                try:
                    value = int(value)
                except:
                    try:
                        value = float(value)
                    except:
                        pass

                fileexcel[str(ii)][f'{adres}'] = value
                # print("lastaddress: ",lastaddress)

        if (loop == 2):
            celolha = a['cells']
            for b in celolha:
                adres = b['address']
                stringpart = adres.rstrip('0123456789')
                numberpart = tail = adres[len(stringpart):]
                numberpart = int(numberpart) + int(lastaddress) + distance
                #print("numberpart: ", numberpart)
                adres = stringpart + str(numberpart)
                #print("adres: ", adres)
                value = b['value']
                #print(f'{adres} = {value}')

                try:
                    value = int(value)
                except:
                    try:
                        value = float(value)
                    except:
                        pass
                #print("ii: ",ii)
                fileexcel[str(ii)][f'{adres}'] = value
    #fileexcel.save('bours.xlsx')
fileexcel.save('bours.xlsx')
os.startfile('bours.xlsx')
