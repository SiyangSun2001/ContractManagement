from win32com import client as wc
import os

convertError = ''
progress = 0

#converts doc file to docx using MS word and win32com
#converts the desired doc file into docx and then deletes the original file

def convert2docx(path):
    global convertError
    global progress1
    convertError = ''
    path = path + '\\' + '前台拷贝'
    os.chdir(path)
    entry = os.listdir()
    numOfFile = len(entry)
    numProcessed = 0
    for c in entry:
        os.chdir(path+'\\'+c)
        files = os.listdir()
        foundcontract = False
        foundShipment = False
        contractConverted  =False
        shipmentConverted = False
        for file in files:
            if '合同' in file and '$' not in file and 'pdf' not in file and 'jpg' not in file and 'jpeg' not in file and '~$' not in file and '.docx' not in file:
                w = wc.Dispatch('Word.Application')
                doc=w.Documents.Open(path+'\\'+c + '\\' + file)
                doc.SaveAs(path+'\\'+ c + '\\' + file[0:-4]+'.docx',16)
                doc.Close()
                foundcontract = True
            if '发货单' in file and '$' not in file and 'pdf' not in file and 'jpg' not in file and 'jpeg' not in file and '~$' not in file and '.docx' not in file:
                w = wc.Dispatch('Word.Application')
                doc=w.Documents.Open(path+'\\'+c + '\\' + file)
                doc.SaveAs(path+'\\'+ c + '\\' + file[0:-4]+'.docx',16)
                doc.Close()
                foundShipment = True
            if '合同' in file and '$' not in file and 'pdf' not in file and 'jpg' not in file and 'jpeg' not in file and '~$' not in file and '.docx'  in file:
                contractConverted = True
            if '发货单' in file and '$' not in file and 'pdf' not in file and 'jpg' not in file and 'jpeg' not in file and '~$' not in file and '.docx' in file:
                shipmentConverted = True
        if foundShipment and not foundcontract and not contractConverted:
            convertError = convertError + '\n' + c + '无法找到.doc合同文件'
        elif foundcontract and not foundShipment and not shipmentConverted:
            convertError = convertError + '\n' + c + '无法找到.doc发货单文件'
        elif not foundShipment and not foundcontract and not shipmentConverted and not contractConverted:
            convertError = convertError + '\n' + c + '无法找到.doc发货单与合同文件'
        elif contractConverted:
            convertError = convertError + '\n' + c + '合同已转换'
        elif shipmentConverted:
            convertError = convertError + '\n' + c + '发货单已转换'
        else:
            convertError = convertError + '\n' + c + '发货单与合同已转换'

        numProcessed += 1
        progress = round((numProcessed/numOfFile)*100)
    for c in entry:
        os.chdir(path+'\\'+c)
        files = os.listdir()
        for file in files:
            if '合同' in file and '.doc' in file and '.docx' not in file:
                os.remove(path+'\\'+ c + '\\' + file)
            if '发货单' in file and '.doc' in file and '.docx' not in file:
                os.remove(path+'\\'+ c + '\\' + file)


