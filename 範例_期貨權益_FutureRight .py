### GetFutureRight 範例
from comtypes.client import GetModule, GetEvents, CreateObject

### GetModule 將 API 打包成 python 可以呼叫的 package
# 第一次執行時會在 comtypes.gen 建立 SKCOMLib wrap package, 之後可以不用再執行，若 API 元件升級時需要至comtypes.gen資料夾刪除 skcom 相關檔案，再重新執行 GetModule
GetModule(r"C:\skcom\CapitalAPI_2.13.43\x64\SKCOM.dll")

from comtypes.gen import SKCOMLib as sk
import pythoncom
import asyncio
import pandas as pd

### 建立 COM 元件
if 'skC' not in globals(): skC = CreateObject(sk.SKCenterLib, interface=sk.ISKCenterLib)
if 'skR' not in globals(): skR = CreateObject(sk.SKReplyLib, interface=sk.ISKReplyLib)
if 'skO' not in globals(): skO = CreateObject(sk.SKOrderLib, interface=sk.ISKOrderLib)

### 輸入身分證與密碼
ID = '你的身分證號'
PW = '你的帳戶密碼'

### 建立 skcom 事件類別
# ReplyLib事件類別
class skR_events:
    def OnReplyMessage(self, bstrUserID, bstrMessage):
        '''API 2.13.17 以上一定要返回 sConfirmCode=-1'''
        sConfirmCode=-1
        print('skR_OnReplyMessage', bstrMessage)
        return sConfirmCode

# OrderLib 事件類別
class skO_events:
    def __init__(self):
        self.future_right = []
        # 現貨帳號
        self.accTS = ''
        # 期貨帳號
        self.accTF = ''
    
    def OnAccount(self, bstrLogInID, bstrAccountData):
        '''取得使用者帳號資訊'''
        strI = bstrAccountData.split(',')
        # bstrAccountData: 0市場,1分公司,2分公司代號,3帳號,4身份證字號,5姓名
        if strI[0] == 'TS': # 現貨帳號
            self.accTS = strI[1]+strI[3]
        elif strI[0] == 'TF': # 期貨帳號
            self.accTF = strI[1]+strI[3]
        # print('skO_OnAccount', bstrLogInID, bstrAccountData)
        
    def OnFutureRights(self, bstrData):
        if '##' not in bstrData:
            self.future_right.append(bstrData)
        # print('skO_OnFutureRights', bstrData)
        
        
# Event sink, 事件實體
EventR = skR_events()
EventO = skO_events()

# 連結 com 元件與事件callback
ConnR = GetEvents(skR, EventR)
ConnO = GetEvents(skO, EventO) 

print('Setting event handeller done!!')


### 適用於 Jupyter lab 的 event loop, 持續監測 skcom api 的事件函數
# working functions, async coruntime to pump events
async def pump_task():
    while True:
        pythoncom.PumpWaitingMessages()
        await asyncio.sleep(0.1)
        
# get an event loop 
loop = asyncio.get_event_loop()
pumping_loop = loop.create_task(pump_task())
print('Event pumping ready!')


### 登入 與 OrderLib 初始化
# login
nCode=skC.SKCenterLib_Login(ID,PW)
print('Login', skC.SKCenterLib_GetReturnCodeMessage(nCode))

# SKOrderLib 初始化
ncode = skO.SKOrderLib_Initialize()
print('SKOrderLib 初始化', skC.SKCenterLib_GetReturnCodeMessage(nCode))

# 取得帳號
ncode = skO.GetUserAccount()
print('取得帳號', skC.SKCenterLib_GetReturnCodeMessage(nCode))


### 取得期貨權益
ncode = skO.GetFutureRights(ID, EventO.accTF, 1)
print('取得期貨權益', skC.SKCenterLib_GetReturnCodeMessage(nCode))

# OnFutureRight 回傳欄位
futurRight_columns = [
    '帳戶餘額', '浮動損益', '已實現費用','交易稅','預扣權利金','權利金收付',
    '權益數','超額保證金','存提款','買方市值','賣方市值','期貨平倉損益',
    '盤中未實現','原始保證金','維持保證金','部位原始保證金','部位維持保證金',
    '委託保證金','超額最佳保證金','權利總值','預扣費用','原始保證金2',
    '昨日餘額','選擇權組合單加不加收保證金','維持率','幣別','足額原始保證金',
    '足額維持保證金', '足額可用','抵繳金額','有價可用','可用餘額',
    '足額現金可用','有價價值','風險指標','選擇權到期差異','選擇權到期差損',
    '期貨到期損益','加收保證金','LOGIN_ID','ACCOUNT_NO']
    
### 存入 dataframe 再轉存 excel 格式
# 存入 pandas dataframe
df = pd.DataFrame([EventO.future_right[0].split(',')], columns=futurRight_columns)

# 存成 excel 
df.to_excel('futureRight.xlsx')