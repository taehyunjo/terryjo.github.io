import win32com.client
import pythoncom
import time
import csv

from collections import defaultdict

class XQuery_t8430: ## 종목 정보 가져오기
    
    def __init__(self):
        self.is_data_received = False
        
        self.nameList = []
        self.shcodeList = []
        self.gubun = []
        
    def OnReceiveData(self, tr_code):
        self.is_data_received = True
        self.count = self.GetBlockCount("t8430OutBlock")

        for i in range(self.count):
            self.nameList.append(self.GetFieldData("t8430OutBlock", "hname", i))
            self.shcodeList.append(self.GetFieldData("t8430OutBlock", "shcode", i))
            self.gubun.append(self.GetFieldData("t8430OutBlock", "gubun", i))
            
    def start(self):
        self.ResFileName = "C:\\eBEST\\xingAPI\\Res\\t8430.res"
    
    def singleRequest(self,gubun):
        self.SetFieldData("t8430InBlock", "gubun",0,0)
        self.Request(False)
    
    @classmethod
    def getInstance(cls):
        xq_t8430 = win32com.client.DispatchWithEvents("XA_DataSet.XAQuery", cls)
        return xq_t8430
    
class XReal_H1_: ## 코스피 호가
    
    def __init__(self):
        super().__init__()
        self._open_file()

    def OnReceiveRealData(self, tr_code):  # event handler
        """
        이베스트 서버에서 ReceiveRealData 이벤트 받으면 실행되는 event handler
        """        
        
        self.data = [] 
        
        """
        종목코드 / 시간(etrade) / 시간(time.time()) 
        매수호가[10] / 매수호가량[10] / 매도호가[10] / 매도호가량[10]
        총매수호가잔량 / 총매도호가잔량 / 동시호가구분 / 배분적용구분
        """   
        
        self.data.append(self.GetFieldData("OutBlock", "shcode"))
        self.data.append(self.GetFieldData("OutBlock", "hotime"))
        self.data.append(time.time())
        
        for i in range(1,11): # 1~10
            self.data.append(self.GetFieldData("OutBlock", "offerho" + str(i)))
        for i in range(1,11): # 1~10
            self.data.append(self.GetFieldData("OutBlock", "offerrem" + str(i)))
        for i in range(1,11): # 1~10
            self.data.append(self.GetFieldData("OutBlock", "bidho" + str(i)))
        for i in range(1,11): # 1~10
            self.data.append(self.GetFieldData("OutBlock", "bidrem" + str(i)))   
            
        self.data.append(self.GetFieldData("OutBlock", "totbidrem"))
        self.data.append(self.GetFieldData("OutBlock", "totofferrem"))
        self.data.append(self.GetFieldData("OutBlock", "donsigubun"))
        self.data.append(self.GetFieldData("OutBlock", "alloc_gubun")) 
        
        

        self.writer.writerow(self.data)
        
    def _open_file(self): #I/O
        f = open('KOSPI_HOGA.csv','a',encoding='utf-8',newline='')
        self.writer = csv.writer(f)

    def start(self):
        self.ResFileName = "C:\\eBEST\\xingAPI\\Res\\H1_.res"  # RES 파일 등록

    def add_item(self, stockcode):
        # 실시간데이터 요청 종목 추가
        self.SetFieldData("InBlock", "shcode", stockcode)
        self.AdviseRealData()

    def remove_item(self, stockcode):
        # stockcode 종목만 실시간데이터 요청 취소
        self.UnadviseRealDataWithKey(stockcode)

    def stop(self):
        self.UnadviseRealData()  # 실시간데이터 요청 모두 취소
        del self.writer

    @classmethod
    def get_instance(cls):
        xreal = win32com.client.DispatchWithEvents("XA_DataSet.XAReal", cls)
        return xreal

    
class XReal_HA_: ## 코스닥 호가
    
    def __init__(self):
        super().__init__()
        self._open_file()

    def OnReceiveRealData(self, tr_code):  # event handler
        """
        이베스트 서버에서 ReceiveRealData 이벤트 받으면 실행되는 event handler
        """        
        
        self.data = [] 
        
        """
        종목코드 / 시간(etrade) / 시간(time.time()) 
        매수호가[10] / 매수호가량[10] / 매도호가[10] / 매도호가량[10]
        총매수호가잔량 / 총매도호가잔량 / 동시호가구분 / 배분적용구분
        """   
        
        self.data.append(self.GetFieldData("OutBlock", "shcode"))
        self.data.append(self.GetFieldData("OutBlock", "hotime"))
        self.data.append(time.time())
        
        for i in range(1,11): # 1~10
            self.data.append(self.GetFieldData("OutBlock", "offerho" + str(i)))
        for i in range(1,11): # 1~10
            self.data.append(self.GetFieldData("OutBlock", "offerrem" + str(i)))
        for i in range(1,11): # 1~10
            self.data.append(self.GetFieldData("OutBlock", "bidho" + str(i)))
        for i in range(1,11): # 1~10
            self.data.append(self.GetFieldData("OutBlock", "bidrem" + str(i)))   
            
        self.data.append(self.GetFieldData("OutBlock", "totbidrem"))
        self.data.append(self.GetFieldData("OutBlock", "totofferrem"))
        self.data.append(self.GetFieldData("OutBlock", "donsigubun"))
        self.data.append(self.GetFieldData("OutBlock", "alloc_gubun"))          

        self.writer.writerow(self.data)
        
    def _open_file(self): #I/O
        f = open('KOSDAQ_HOGA.csv','a',encoding='utf-8',newline='')
        self.writer = csv.writer(f)

    def start(self):
        self.ResFileName = "C:\\eBEST\\xingAPI\\Res\\HA_.res"  # RES 파일 등록

    def add_item(self, stockcode):
        # 실시간데이터 요청 종목 추가
        self.SetFieldData("InBlock", "shcode", stockcode)
        self.AdviseRealData()

    def remove_item(self, stockcode):
        # stockcode 종목만 실시간데이터 요청 취소
        self.UnadviseRealDataWithKey(stockcode)

    def stop(self):
        self.UnadviseRealData()  # 실시간데이터 요청 모두 취소
        del self.writer

    @classmethod
    def get_instance(cls):
        xreal = win32com.client.DispatchWithEvents("XA_DataSet.XAReal", cls)
        return xreal
    

class XReal_S3_: ## 코스피 체결
    
    def __init__(self):
        super().__init__()
        self._open_file()

    def OnReceiveRealData(self, tr_code):  # event handler
        """
        이베스트 서버에서 ReceiveRealData 이벤트 받으면 실행되는 event handler
        """        
        
        self.data = [] 
        
        """
        종목코드 / 체결시간(ETRADE) / 체결시간(time.time())
        체결가격 / 체결구분(매수/매도) / 체결량
        시가 / 시가시간 / 고가 / 고가시간 / 저가 / 저가시간
        누적거래량 / 누적거래대금
        
        ## 쓸모 없어 보이는 정보 ##
        전일대비 / 등락율 / 매도누적체결량 / 매도누적체결건수 / 매수누적체결량 / 매수누적체결건수
        체결강도 / 가중평균가 / 매도호가 / 매수호가 / 장정보 / 전일동시간대거래량
        
        
        
        """
    
        self.data.append(self.GetFieldData("OutBlock", "shcode"))
        self.data.append(self.GetFieldData("OutBlock", "chetime"))
        self.data.append(time.time())
        
        self.data.append(self.GetFieldData("OutBlock", "price"))
        self.data.append(self.GetFieldData("OutBlock", "cgubun"))
        self.data.append(self.GetFieldData("OutBlock", "cvolume"))
        
        self.data.append(self.GetFieldData("OutBlock", "open"))
        self.data.append(self.GetFieldData("OutBlock", "opentime"))
        
        self.data.append(self.GetFieldData("OutBlock", "high"))
        self.data.append(self.GetFieldData("OutBlock", "hightime"))
        
        self.data.append(self.GetFieldData("OutBlock", "low"))
        self.data.append(self.GetFieldData("OutBlock", "lowtime"))
        
        self.data.append(self.GetFieldData("OutBlock", "volume"))
        self.data.append(self.GetFieldData("OutBlock", "value")) 

        self.data.append(self.GetFieldData("OutBlock", "change")) 
        self.data.append(self.GetFieldData("OutBlock", "drate")) 
        self.data.append(self.GetFieldData("OutBlock", "mdvolume")) 
        self.data.append(self.GetFieldData("OutBlock", "mdchecnt")) 
        self.data.append(self.GetFieldData("OutBlock", "msvolume")) 
        self.data.append(self.GetFieldData("OutBlock", "mschecnt")) 
        self.data.append(self.GetFieldData("OutBlock", "cpower")) 
        self.data.append(self.GetFieldData("OutBlock", "w_avrg")) 
        self.data.append(self.GetFieldData("OutBlock", "offerho")) 
        self.data.append(self.GetFieldData("OutBlock", "bidho")) 
        self.data.append(self.GetFieldData("OutBlock", "status")) 
        self.data.append(self.GetFieldData("OutBlock", "jnilvolume")) 
        
        
        self.writer.writerow(self.data)
        
    def _open_file(self): #I/O
        f = open('KOSPI_EXECUTED.csv','a',encoding='utf-8',newline='')
        self.writer = csv.writer(f)

    def start(self):
        self.ResFileName = "C:\\eBEST\\xingAPI\\Res\\S3_.res"  # RES 파일 등록

    def add_item(self, stockcode):
        # 실시간데이터 요청 종목 추가
        self.SetFieldData("InBlock", "shcode", stockcode)
        self.AdviseRealData()

    def remove_item(self, stockcode):
        # stockcode 종목만 실시간데이터 요청 취소
        self.UnadviseRealDataWithKey(stockcode)

    def stop(self):
        self.UnadviseRealData()  # 실시간데이터 요청 모두 취소
        del self.writer

    @classmethod
    def get_instance(cls):
        xreal = win32com.client.DispatchWithEvents("XA_DataSet.XAReal", cls)
        return xreal
    
class XReal_K3_: ## 코스피 체결
    
    def __init__(self):
        super().__init__()
        self._open_file()

    def OnReceiveRealData(self, tr_code):  # event handler
        """
        이베스트 서버에서 ReceiveRealData 이벤트 받으면 실행되는 event handler
        """        
        
        self.data = [] 
        
        """
        종목코드 / 체결시간(ETRADE) / 체결시간(time.time())
        체결가격 / 체결구분(매수/매도) / 체결량
        시가 / 시가시간 / 고가 / 고가시간 / 저가 / 저가시간
        누적거래량 / 누적거래대금
        
        ## 쓸모 없어 보이는 정보 ##
        전일대비 / 등락율 / 매도누적체결량 / 매도누적체결건수 / 매수누적체결량 / 매수누적체결건수
        체결강도 / 가중평균가 / 매도호가 / 매수호가 / 장정보 / 전일동시간대거래량
        
        """
    
        self.data.append(self.GetFieldData("OutBlock", "shcode"))
        self.data.append(self.GetFieldData("OutBlock", "chetime"))
        self.data.append(time.time())
        
        self.data.append(self.GetFieldData("OutBlock", "price"))
        self.data.append(self.GetFieldData("OutBlock", "cgubun"))
        self.data.append(self.GetFieldData("OutBlock", "cvolume"))
        
        self.data.append(self.GetFieldData("OutBlock", "open"))
        self.data.append(self.GetFieldData("OutBlock", "opentime"))
        
        self.data.append(self.GetFieldData("OutBlock", "high"))
        self.data.append(self.GetFieldData("OutBlock", "hightime"))
        
        self.data.append(self.GetFieldData("OutBlock", "low"))
        self.data.append(self.GetFieldData("OutBlock", "lowtime"))
        
        self.data.append(self.GetFieldData("OutBlock", "volume"))
        self.data.append(self.GetFieldData("OutBlock", "value")) 

        self.data.append(self.GetFieldData("OutBlock", "change")) 
        self.data.append(self.GetFieldData("OutBlock", "drate")) 
        self.data.append(self.GetFieldData("OutBlock", "mdvolume")) 
        self.data.append(self.GetFieldData("OutBlock", "mdchecnt")) 
        self.data.append(self.GetFieldData("OutBlock", "msvolume")) 
        self.data.append(self.GetFieldData("OutBlock", "mschecnt")) 
        self.data.append(self.GetFieldData("OutBlock", "cpower")) 
        self.data.append(self.GetFieldData("OutBlock", "w_avrg")) 
        self.data.append(self.GetFieldData("OutBlock", "offerho")) 
        self.data.append(self.GetFieldData("OutBlock", "bidho")) 
        self.data.append(self.GetFieldData("OutBlock", "status")) 
        self.data.append(self.GetFieldData("OutBlock", "jnilvolume"))              
        
        self.writer.writerow(self.data)
        
    def _open_file(self): #I/O
        f = open('KOSDAQ_EXECUTED.csv','a',encoding='utf-8',newline='')
        self.writer = csv.writer(f)

    def start(self):
        self.ResFileName = "C:\\eBEST\\xingAPI\\Res\\K3_.res"  # RES 파일 등록

    def add_item(self, stockcode):
        # 실시간데이터 요청 종목 추가
        self.SetFieldData("InBlock", "shcode", stockcode)
        self.AdviseRealData()

    def remove_item(self, stockcode):
        # stockcode 종목만 실시간데이터 요청 취소
        self.UnadviseRealDataWithKey(stockcode)

    def stop(self):
        self.UnadviseRealData()  # 실시간데이터 요청 모두 취소
        del self.writer

    @classmethod
    def get_instance(cls):
        xreal = win32com.client.DispatchWithEvents("XA_DataSet.XAReal", cls)
        return xreal


