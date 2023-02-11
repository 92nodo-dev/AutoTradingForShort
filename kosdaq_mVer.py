import win32com.client
from openpyxl import Workbook
from enum import Enum
from pykiwoom.kiwoom import *
import time
import threading

########################### 대신증권 CYBOS PLUS 연결 ###########################
kiwoom = Kiwoom()
kiwoom.CommConnect(block=True)

objCpCybos = win32com.client.Dispatch("CpUtil.CpCybos")
print("블록킹 로그인 완료")
bConnect = objCpCybos.IsConnect
if (bConnect == 0):
    print("PLUS가 정상적으로 연결되지 않음. ")
    exit()

################################################################################

# 차트 객체 구하기
objStockChart = win32com.client.Dispatch("CpSysDib.StockChart")
g_objCodeMgr = win32com.client.Dispatch('CpUtil.CpCodeMgr')
objStockWeek = win32com.client.Dispatch("DsCbo1.StockWeek")
codeList = g_objCodeMgr.GetStockListByMarket(1) # 
codeList2 = g_objCodeMgr.GetStockListByMarket(2) # 
allcodelist = codeList+codeList2

targetList = []
excelIndex = 1

stockList = []
checkList = []
percentage = 0

class stepState(Enum) :
    phase1 = 0 # 2percent over stock buy
    phase2 = 1
    
class State(Enum) :
    UNDER_2_PERCENT = 0
    HAVING_1 = 1
    HAVING_2 = 2
    HAVING_4 = 3
    
class Stock :
    def __init__(self, code) :
        self.code = code
        self.state = State.UNDER_2_PERCENT
        self.buyPrice = 0
        self.nowPercent = 0
        self.buyCount = 0
        
   # def push(self, num) :




executeJustFirstTime = True
class AsyncTask:
    def __init__(self):
        pass

    def run(self) :
        step = stepState.phase1
        totalHavingCnt = 0
        print("running")
        objStockWeek.SetInputValue(0, "U201")
        objStockWeek.BlockRequest()
        while True :
            now = time.localtime()
            if now.tm_hour < 9 :
                continue
            
            #if now.tm_hour == 9 :
            #    if now.tm_min > 25 :
            #        step = stepState.phase2

            if (step == stepState.phase1) :
                for i in range(len(targetList)) :
                    if totalHavingCnt > 5 :
                        continue
                    
                    if (targetList[i].state != State.UNDER_2_PERCENT) :
                        continue
                    
                    objStockWeek.SetInputValue(0, targetList[i].code)
                    objStockWeek.BlockRequest()
                    start = objStockWeek.GetDataValue(1, 0)
                    high = objStockWeek.GetDataValue(2, 0)
                    close = objStockWeek.GetDataValue(4, 0)
                    if (close > (start*1.02)):
                        if (((high - start)/(close - start)) > 1.6) :
                            cd = targetList[i].code[1:]
                            accounts = kiwoom.GetLoginInfo("ACCNO")
                            stock_account = accounts[0]
                            kiwoom.SendOrder("시장가매수", "0101", stock_account, 1, str(cd), int(1840000/close), 0, "03", "")
                            targetList[i].buyPrice = close
                            targetList[i].buyCount = int(1840000/close)
                            
                            print("[매수] " + g_objCodeMgr.CodeToName(targetList[i].code) + " : " + str(close))
                            time.sleep(5)
                            totalHavingCnt += 1
                            targetList[i].state = State.HAVING_1
            
            #if step == stepState.phase2 :
                
                #if (totalHavingCnt < 99) & (executeJustFirstTime == True) :
                #    moneyForStock = (9500000 - (totalHavingCnt*95000))/totalHavingCnt
                #    for i in range(len(targetList)) :
                #        if targetList[i].state == State.HAVING_1 :
                #            objStockWeek.SetInputValue(0, targetList[i].code)
                #            objStockWeek.BlockRequest()
                #            close = objStockWeek.GetDataValue(4, 0)
                #            
                #            cd = targetList[i].code[1:]
                #            accounts = kiwoom.GetLoginInfo("ACCNO")
                #            stock_account = accounts[0]
                #            kiwoom.SendOrder("시장가매수", "0101", stock_account, 1, str(cd), int(moneyForStock/close), 0, "03", "")
                #            targetList[i].buyCount += int(moneyForStock/close)
                #        
                #            print("[추가 매수] " + g_objCodeMgr.CodeToName(targetList[i].code) + " : " + str(close))
                #            time.sleep(5)
                #    executeJustFirstTime = False
                            
                for i in range(len(targetList)) :
                    objStockWeek.SetInputValue(0, targetList[i].code)
                    objStockWeek.BlockRequest()
                    start = objStockWeek.GetDataValue(1, 0)
                    close = objStockWeek.GetDataValue(4, 0)
                    
                    if ((targetList[i].state == State.HAVING_1) & ((((close - start)*100/start) < -5) | (((close - start)*100/start) > 2.7))) | ((targetList[i].state == State.HAVING_1) & (now.tm_hour == 15)) :
                        accounts = kiwoom.GetLoginInfo("ACCNO")
                        cd = targetList[i].code[1:]
                        stock_account = accounts[0]
                        totalHavingCnt -= 1
                        kiwoom.SendOrder("시장가매도", "0101", stock_account, 2, str(cd), targetList[i].buyCount, 0, "03", "")
                        print("[매도] " + g_objCodeMgr.CodeToName(targetList[i].code) + " : " + str(close))
                        targetList[i].state = State.UNDER_2_PERCENT
                        #targetList[i].nowPercent = (targetList[i].buyPrice-close)*100/targetList[i].buyPrice

                #targetList.sort(key=lambda x:x.nowPercent, reverse=True)
                # 
                #for i in range(len(targetList)) :
                #    objStockWeek.SetInputValue(0, targetList[i].code)
                #    objStockWeek.BlockRequest()
                #    close = objStockWeek.GetDataValue(4, 0)
                #    start = objStockWeek.GetDataValue(1, 0)
                #    high = objStockWeek.GetDataValue(2, 0)
                #    
                #    if ((close-start/start) < (high-start/start)*0.6) & (targetList[i].state == State.HAVING_1) & (i < totalHavingCnt*0.5) :
                        
                    
def main():
    loadList()
    at = AsyncTask()
    at.run()
        
def loadList() :
    bannedList = []
    
    myFile = open('bannedList.txt', 'r')
    while True :
        tmp = myFile.readline()[:-1]
        if tmp == 'X' :
            break
        if tmp == "" :
            break
        bannedList.append("A" + tmp)

    myFile.close()

    for i in range(len(allcodelist)) :
        
        skind = g_objCodeMgr.GetStockSectionKind(allcodelist[i])

        if (skind==10) | (skind==12) :
            continue

        if allcodelist[i][0] != 'A' :
            continue

        if allcodelist[i][6].isalpha() :
            continue

        namestr = g_objCodeMgr.CodeToName(allcodelist[i])
        if namestr == "" :
            continue
        if namestr[len(namestr)-1]=='우' :
            continue
        if "스팩" in namestr :
            continue
        if allcodelist[i] in bannedList :
            continue

        objStockChart.SetInputValue(0, allcodelist[i])
        objStockChart.SetInputValue(1, ord('2')) # 개수로 조회
        objStockChart.SetInputValue(4, 10) # 최근 n일 치
        objStockChart.SetInputValue(5, [0,2,3,4,5, 8]) #날짜,시가,고가,저가,종가,거래량
        objStockChart.SetInputValue(6, ord('D')) # '차트 주가 - 일간 차트 요청
        objStockChart.SetInputValue(9, ord('1')) # 수정주가 사용
        objStockChart.BlockRequest()
        sum3=0
        sum5=0
        sum10 = 0
        try :
            for j in range(10) :
                sum10 += objStockChart.GetDataValue(4, 9-j)
                if j>4 :
                    sum5 += objStockChart.GetDataValue(4, 9-j)
                if j>6 :
                    sum3 += objStockChart.GetDataValue(4, 9-j)
        except :
            continue
        recent = objStockChart.GetDataValue(4, 0)
        if (((sum3/3) > recent) | ((sum5/5) > recent) | ((sum10/10) > recent)) :
            continue
        
        targetList.append(Stock(allcodelist[i]))

    print(str(len(targetList)) + "개의 종목이 살아남았습니다.")
            
if __name__ == '__main__':
    main()
