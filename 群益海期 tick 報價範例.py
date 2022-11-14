# # 海期 tick 報價範例
import pythoncom
import asyncio
import datetime
import pandas as pd
import comtypes.client as cc
import plotly.graph_objects

# 只有第一次使用 api ，或是更新 api 版本時，才需要呼叫 GetModule
# 會將 SKCOM api 包裝成 python 可用的 package ，並存放在 comtypes.gen 資料夾下
# 更新 api 版本時，記得將 comtypes.gen 資料夾 SKCOMLib 相關檔案刪除，再重新呼叫 GetModule 
cc.GetModule('C:\\skcom\\CapitalAPI_2.13.39\\x64\\SKCOM.dll')
import comtypes.gen.SKCOMLib as sk


# login ID and PW
# 身份證
ID = ''
# 密碼
PW = ''
print(datetime.datetime.now().strftime("%Y/%m/%d %H:%M:%S,"), 'Set ID and PW')


# # 建立 event pump and event loop
# 新版的jupyterlab event pump 機制好像有改變，因此自行打造一個 event pump機制，
# 目前在 jupyterlab 環境下使用，也有在 spyder IDE 下測試過，都可以正常運行

# working functions, async coruntime to pump events
async def pump_task():
    '''在背景裡定時 pump windows messages'''
    while True:
        pythoncom.PumpWaitingMessages()
        # 想要反應更快 可以將 0.1 取更小值
        await asyncio.sleep(0.1)

# 將ticks 轉為Kline
def convert_to_kline(query_stock, freq):
    '''將ticks 轉為Kline
    query_stock: 欲查詢的商品代號 ex. "YM2212"
    freq: 請參考 pandas resample 用法, "T" 為分, "S"為秒
          "5T" 為 5分Kline, "30S" 為30秒Kline
    return a kline dataframe
    '''
    # 只保留成交時間,成交價與量資料
    df = EventOSQ.ticks.query(f'bstrStockNo == "{query_stock}"').copy()
    df = df.filter(['Datetime', 'price', 'volume'], axis=1)

    # 將成交時間欄位資料按格式轉換為 datetime 資料
    df['Datetime'] = pd.to_datetime(df['Datetime'], format='%Y%m%d%H%M%S')

    # 設定資料以成交時間欄位為序列索引
    df = df.set_index('Datetime')

    # return OHLCV Kline
    kline = df.resample(rule=freq).agg({'price': 'ohlc', 'volume': 'sum'}).dropna()
    kline.columns = kline.columns.get_level_values(1)
    return kline

# plot kline
def plot_candlestick(df):
    figure = plotly.graph_objects.Figure(
        data=[
            plotly.graph_objects.Candlestick(
                x=df.index,
                open=df['open'],
                high=df['high'],
                low=df['low'],
                close=df['close'],
                name='K line',
            ),
        ],
        # 設定 XY 顯示格式
        layout=plotly.graph_objects.Layout(
            xaxis=plotly.graph_objects.layout.XAxis(
                tickformat='%Y-%m-%d %H:%M'
            ),
            yaxis=plotly.graph_objects.layout.YAxis(
                tickformat='.2f'
            )
        )
    )
    figure.show()


# get an event loop
loop = asyncio.get_event_loop()
pumping_loop = loop.create_task(pump_task())
print(datetime.datetime.now().strftime("%Y/%m/%d %H:%M:%S,"), "Event pumping is ready!")


# 建立物件，避免重複 createObject
# 登錄物件
if 'skC' not in globals(): 
    skC = cc.CreateObject(sk.SKCenterLib, interface=sk.ISKCenterLib)
# 海期報價物件
if 'skOSQ' not in globals(): 
    skOSQ = cc.CreateObject(sk.SKOSQuoteLib , interface=sk.ISKOSQuoteLib)
# 回報物件
if 'skR' not in globals(): 
    skR = cc.CreateObject(sk.SKReplyLib, interface=sk.ISKReplyLib)


# # # 建立 event handler
# SKOSQ event handler
class skOSQ_events:
    def __init__(self):
        self.OverseaProductsDetail = []
        # 以dataframe方式存放ticks 
        self.ticks = pd.DataFrame(
            {'Datetime': pd.Series(dtype='str'), 
             'price': pd.Series(dtype='float'), 
             'volume': pd.Series(dtype='int')},

            index = pd.MultiIndex(levels=[[],[],[],[]],
                                  codes=[[],[],[],[]],
                                 names=['bstrStockNo', 'nPtr', 'nDate', 'nTime']),
                                 )

    def OnConnect(self, nKind, nCode):
        '''連線海期主機狀況回報'''
        print(f'skOSQ_OnConnect nCode={nCode}, nKind={nKind}')

    def OnOverseaProductsDetail(self, bstrValue):
        '''查詢海期/報價下單商品代號'''
        if "##" not in self.OverseaProductsDetail:
            self.OverseaProductsDetail.append(bstrValue.split(','))
        else:
            print("skOSQ_OverseaProductsDetail downloading is completed.")

    def OnNotifyQuoteLONG(self, sIndex):
        '''requestStock 報價回報'''
        # 儘量避免在這裡使用繁複的運算，這裡僅在 console 端印出報價
        ts = sk.SKFOREIGNLONG()
        skOSQ.SKOSQuoteLib_GetStockByIndexLONG(sIndex, ts)
        print(ts.bstrExchangeNo, ts.bstrStockNo, ts.nClose, ts.nTickQty)

    def OnNotifyTicksNineDigitLONG (self, nIndex, nPtr, nDate, nTime, 
                                    nClose, nQty):
        '''requestTick 回報'''
        # 儘量避免在這裡使用繁複的運算
        ts = sk.SKFOREIGN_9LONG()
        skOSQ.SKOSQuoteLib_GetStockByIndexNineDigitLONG(nIndex, ts)
        self.ticks.loc[(ts.bstrStockNo, nPtr, nDate, nTime), 
                       ["Datetime", "price", "volume"]] = [f"{nDate}{nTime:06}", 
                                                          nClose/10**ts.sDecimal, 
                                                          nQty]

    def OnNotifyHistoryTicksNineDigitLONG (self, nIndex, nPtr, 
        nDate, nTime, nClose, nQty):
        ''' History tick 回報'''
        # 儘量避免在這裡使用繁複的運算
        ts = sk.SKFOREIGN_9LONG()
        ncode = skOSQ.SKOSQuoteLib_GetStockByIndexNineDigitLONG(nIndex, ts)
        self.ticks.loc[(ts.bstrStockNo, nPtr, nDate, nTime), 
                       ["Datetime", "price", "volume"]] = [f"{nDate}{nTime:06}", 
                                                          nClose/10**ts.sDecimal, 
                                                          nQty]


# SKReplyLib event handler
class skR_events:
    def OnReplyMessage(self, bstrUserID, bstrMessage):
        '''API 2.13.17 以上一定要返回 sConfirmCode=-1'''
        sConfirmCode = -1
        print('skR_OnReplyMessage ok')
        return sConfirmCode


# # 建立 event 跟 event handler 的連結

# Event sink, 事件實體化
EventOSQ = skOSQ_events()
EventR = skR_events()

# 建立 event 跟 event handler 的連結
ConnOSQ = cc.GetEvents(skOSQ, EventOSQ)
ConnR = cc.GetEvents(skR, EventR)


# # 登入及各項初始化作業

# login
print('Login', skC.SKCenterLib_GetReturnCodeMessage(skC.SKCenterLib_Login(ID,PW)))

# 海期商品初始化
nCode = skOSQ.SKOSQuoteLib_Initialize()
print("SKOSQuoteLib_Initialize", skC.SKCenterLib_GetReturnCodeMessage(nCode))


###################################################################################
# 以下皆以手動輸入
# 登入海期報價主機
nCode = skOSQ.SKOSQuoteLib_LeaveMonitor()
nCode = skOSQ.SKOSQuoteLib_EnterMonitorLONG()
print('SKOSQuoteLib_EnterMonitorLONG()', skC.SKCenterLib_GetReturnCodeMessage(nCode))

# 登入海期報價主機，確認 OnConnect 出現 3001 回報後
# 才可 requesttick
StockNo ="CBOT,YM2212"
nCode = skOSQ.SKOSQuoteLib_RequestTicks(0, StockNo)
print(f"Requesting ticks, {StockNo}", skC.SKCenterLib_GetReturnCodeMessage(nCode[1]))

# 檢視 ticks 狀態，熱門商品資料數量很多，可能要等一下資料回傳完畢
EventOSQ.ticks

# 轉換為 1分K,並畫出來，pandas 轉 kline 數據一多，好像有點慢
df_1k = convert_to_kline(query_stock="YM2212" ,freq="1T")
plot_candlestick(df_1k)

# 轉換為 5分K,並畫出來
df_5k = convert_to_kline(query_stock="YM2212",freq="5T")
plot_candlestick(df_5k)