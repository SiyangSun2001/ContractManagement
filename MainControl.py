import SalesRecordAssistant
import CustomerInfoAssistant
import MoveContractFiles

errorList = []
correctList = []
error = ''
def EnterData(path):

    #make contract OBJ
    #call function in salesRecordAssistant to make list of Contract OBJ

    list = SalesRecordAssistant.makeContractList(path)
    #Finally Move Contract into the storage folder
    MoveContractFiles.moveContract(path,list)
    global errorList
    global correctList
    global error
    for contract in list:
        if contract.error:
            errorList.append(contract)
        else:
            correctList.append(contract)


    #Input the extracted data into the Sales Record
    SalesRecordAssistant.checkNewSheets(path)
    try:
        SalesRecordAssistant.enterData(path,correctList)
    except:
        error = '销售统计录入失败！'
    #Input the customer info
    for contract in correctList:
        # try:
            #try to enter data directly assuming there is an exisiting sheet
        try:
            CustomerInfoAssistant.enterCustomerInfo(path,contract)
        except FileNotFoundError:
            #if failed, create new sheet and then enter
            CustomerInfoAssistant.createNewWBandEnter(path,contract)




