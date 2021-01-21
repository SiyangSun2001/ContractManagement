import os
import shutil

def moveContract(path,listOfContract):
    os.chdir(path + '\\' + '\用户合同')
    for aContract in listOfContract:
        storageEntry = os.listdir()
        setOfStorageEntry = set(storageEntry)
        if aContract.companyName not in setOfStorageEntry:
            os.mkdir(path + '\\' + '用户合同' + '\\' + aContract.companyName)
            shutil.move(aContract.path,path + '\\' + '用户合同' + '\\' + aContract.companyName)
        else:
            try:
                shutil.move(aContract.path,path + '\\' + '用户合同' + '\\' + aContract.companyName)
            except shutil.Error:
                aContract.error = True
                aContract.errorMsg += '\n 合同已录入'

