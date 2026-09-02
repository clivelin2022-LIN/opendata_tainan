#載入所需套用工具
import os #處理路徑
import logging #記錄日誌(Log)
import requests as rq #向server提出請求
import pandas as pandas #處理資料表
from sqlalchemy import create_engine, text #處理資料庫DB engine (SqlAlchemy)
import pymysql #MySql資料庫連接
import openpyxl #處理excel檔案
import urllib3 #處理url
from dotenv import load_dotenv #抓取.env檔案

#1.載入.env設定檔
load_dotenv()

#2.設定log 機制(純寫到file &輸出console終端機上)
logging.basicConfig(
    level=logging.INFO, #INFO等級以上的資訊都列印
    format='%(asctime)s [%(levelname)s] %(filename)s(行:%(lineno)d:%(message)s',
    datefmt='%Y-%m-%d %H:%M:%S',#設定時間格式
handlers=[
        logging.FileHandler("scraper.log", encoding="utf-8"),
        logging.StreamHandler()
    ]
)

#設定變數
DB_HOST = os.environ.get("DB_HOST", "localhost")
DB_USER = os.environ.get("DB_USER", "peter")
DB_PASSWORD = os.environ.get("DB_PASSWORD", "[PASSWORD]")
DB_PORT = int(os.environ.get("DB_PORT", 3306))
DB_NAME = os.environ.get("DB_NAME", "tainan")
DB_CHARSET = os.environ.get("DB_CHARSET", "utf8mb4")


# API 與檔案設定
api_url= os.environ.get("https://soa.tainan.gov.tw/Api/Service/Get/ea68f0dd-a0e9-4200-be8a-a4b76c3f87a4")
EXCEL_FILENAME= os.environ.get("EXCEL_FILENAME", "tainan_house.xlsx")

#透過api抓取 opendata
def fetch_data(api_url, params=None):
    #檢查 api_url
    if not api_url:
        logging.error("API URL 為空")
        return None
try:
    logging.info(f"正在從API get data...")
    # get opendata by API
    res = rq.get(api_url, verify=False)

    # 檢查回傳狀態
    if res.ok: # 200
        # 轉換資料型態 json -> python
        data = res.json()
        logging.info(f"成功取得資料, 共 {len(data.get('data',[]))}筆")
        return data
    else:
        logging.warning(f"回應錯誤資料: {res.status_code}")
        return None
except rq.exceptions.SSLError as ssl_err:
    logging.error(f"SSL認證錯誤: {ssl_err}")
    return None
except Exception as error:
    logging.error(f"執行擷取open時，發生錯誤: {error}")
    return None
