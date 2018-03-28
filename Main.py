import win32com.client
import pythoncom
import time
import csv

from Login import *
from API import *
from collections import defaultdict
    
if __name__ == "__main__":
    
    def getAllStocksInfo():
        xq_t8430 = XQuery_t8430.getInstance()
        xq_t8430.start()
        xq_t8430.singleRequest(0) # 0:전체 1:코스피 2:코스닥
    
        while xq_t8430.is_data_received == False:
            pythoncom.PumpWaitingMessages()
        
        return [xq_t8430.nameList,xq_t8430.shcodeList,xq_t8430.gubun]
    
    def getRealData():
        
        #코스피호가 인스턴스
        xreal_KOSPI = XReal_H1_.get_instance()
        xreal_KOSPI.start()
        
        #코스닥호가 인스턴스
        xreal_KOSDAQ = XReal_HA_.get_instance()
        xreal_KOSDAQ.start()
        
        #코스피체결 인스턴스
        xreal_KOSPI_excuted = XReal_S3_.get_instance()
        xreal_KOSPI_excuted.start()
        
        #코스닥체결 인스턴스
        xreal_KOSDAQ_excuted = XReal_K3_.get_instance()
        xreal_KOSDAQ_excuted.start()
               
        for shcode_ in stockCodeDict['KOSPI']:
            xreal_KOSPI.add_item(shcode_)
            xreal_KOSPI_excuted.add_item(shcode_)
            
        for shcode_ in stockCodeDict['KOSDAQ']:
            xreal_KOSDAQ.add_item(shcode_)
            xreal_KOSDAQ_excuted.add_item(shcode_)
                
        while True:
            pythoncom.PumpWaitingMessages()
            
    ## 로그인
    Login()
    
    ## 주식 정보 받아옴
    stockNameList,stockCodeList,stockGubunList = getAllStocksInfo()
        
    ## Dictionary 형태로 코스피/코스닥 나눠서 저장 => KEY값 : KOSPI , KOSDAQ
    stockCodeDict = defaultdict(list)
    for i,g in enumerate(stockGubunList):
        if g == '1':
            stockCodeDict['KOSPI'].append(stockCodeList[i])
        elif g == '2':
            stockCodeDict['KOSDAQ'].append(stockCodeList[i])
    
    ## 실시간 데이터 수집
    getRealData()