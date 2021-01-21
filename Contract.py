import docx2txt
import os
import re
from docx2python import docx2python

class Contract:
#this contract object extract all useful info from docx file and store them in appropraite variables

    def __init__(self, path):
        #path is the location of the file, given when obj is created
        path = path + "\\"
        self.path = path
        self.error = False
        self.errorMsg = '错误提示：' + path
        #find file name to contract and inspection file
        os.chdir(path)
        dirEntry = os.listdir(path)
        contractName = ''
        shipmentName = ''
        for entry in dirEntry:
            if '合同' in entry and '$' not in entry and 'pdf' not in entry and 'jpg' not in entry and 'jpeg' not in entry and '~$' not in entry:
                contractName = entry
            if '发货单' in entry:
                shipmentName = entry
        if len(contractName) == 0 or len(shipmentName) == 0:
            foundFile  = False
            self.error = True
            self.errorMsg += '\n 无法找到合同或者发货单文件，无法录入'
        else:
            foundFile = True
        if foundFile:

            #use docx2txt to find contract name and company name
            contractInPy = docx2txt.process(contractName)
            if '合同编号：' in contractInPy:
                self.contractNum = contractInPy[contractInPy.find('合同编号：') + 5:contractInPy.find('合同编号：') + 12]
            else:
                self.contractNum = 0
                self.error = True
                self.errorMsg += '\n 合同docx文件，无法提取合同编号'

            if '需  方:' in contractInPy and '产品名称' in contractInPy:
                fullnamecomp_re = r"(?<=需  方: )(.*)(?=产品名称)"
                match2 = re.search(fullnamecomp_re, contractInPy, flags=re.DOTALL)
                self.companyFullName = match2[0].strip()
            else:
                self.companyFullName = 0
                self.error = True
                self.errorMsg += '\n 合同docx文件，无法提取需方信息'
            #use docx2txt to fetch info in shipment info file
            shipmentinPy = docx2txt.process(shipmentName)
            if '用户：' in shipmentinPy:
                compnam_re = r"(?<=用户：)([^\s]+)"
                match = re.search(compnam_re, shipmentinPy)
                self.companyName = match[0].strip()
            else:
                self.error = True
                self.companyName = 0
                self.errorMsg += '\n 发货单内无法提取用户简称'

            if '收货单位地址：' in shipmentinPy:
                shipadd_re = r'(?<=收货单位地址：)(.*)(\s)'
                match4 = re.search(shipadd_re, shipmentinPy)
                self.address = match4[0].strip()
            elif '收货地址：' in shipmentinPy:
                shipadd_re = r'(?<=收货地址：)(.*)(\s)'
                match4 = re.search(shipadd_re, shipmentinPy)
                self.address = match4[0].strip()
            else:
                self.error = True
                self.address = 0
                self.errorMsg += '\n发货单内无法提取收货单位地址'
            phone_re = r"(?<=电话：)\s*([0-9]{3}|[0-9]{4})-*[0-9]{4}\s*[0-9]{4}"
            match3 = re.search(phone_re, shipmentinPy)
            if match3 == None:
                self.phone = '无'
            else:
                self.phone = match3[0].strip()


            # function for docx2python, remove empty element from returned list
            def remove_empty(table):
                # remove empty element of list
                return list(filter(lambda x: not isinstance(x, (str, list, tuple)) or x,
                                   (remove_empty(x) if isinstance(x, (tuple, list)) else x for x in table)))


            # use docx2python to generate list and use that list to find price ,model count, and model number
            contractInList = docx2python(path + contractName)
            table =  remove_empty(contractInList.body)
            self.modelNumber = []
            self.modelCount  =[]
            self.price = []

            for row in table[1][1:]:
                if len(row) == 5:

                    self.modelNumber.append(row[1][0])

                    if row[3][0].find("台") == -1:
                        self.modelCount.append(int(row[3][0]))
                    else:
                        self.modelCount.append(int(row[3][0][:row[3][0].find("台")]))

                    if row[2][0].find("元") == -1:
                        self.price.append(int(row[2][0]))
                    else:
                        self.price.append(int(row[2][0][:row[2][0].find("元")]))

#function that take the unti price, model count, then generate a string in the format of
#unit price * model count = total
    def getFormattedPrice(self):
        formattedPrice = []
        for i in range(len(self.price)):
            formattedPrice.append(str(self.price[i]) + ' x ' + str(self.modelCount[i]) + ' = ' + str((int(self.price[i]) * int(self.modelCount[i]))))
        return formattedPrice


