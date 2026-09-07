#載入所需套用工具
import logging
import os #處理路徑
import logging #記錄日誌(Log)
import requests as rq #向server提出請求
import pandas as pd #處理資料表
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
api_url= os.environ.get("API_URL", "https://soa.tainan.gov.tw/Api/Service/Get/ea68f0dd-a0e9-4200-be8a-a4b76c3f87a4")
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
        response = rq.get(api_url, verify=False)

        # 檢查回傳狀態
        if response.ok: # 200
            # 轉換資料型態 json -> python
            data = response.json()
            logging.info(f"成功取得資料, 共 {len(data.get('data', []))}筆")
            return data
        else:
            logging.warning(f"回應錯誤資料: {response.status_code}")
            return None
    except rq.exceptions.SSLError as ssl_err:
        logging.error(f"SSL認證錯誤: {ssl_err}")
        return None
    except Exception as error:
        logging.error(f"執行擷取open時，發生錯誤: {error}")
        return None

def save_to_excel(data, filename):
    try:
        #將資料轉換成pandasd的資料型態 pandas dataFrame, excel(2D):dataframe,openpyxl為excel的驅動(用來讀取寫入excel)
        df = pd.DataFrame(data)
        df.to_excel(filename, index=False, engine="openpyxl")
        logging.info(f"資料已寫入成功: excel {filename}")
    except Exception as e:
        #exc_info:True->將錯誤的 traceback 一併記錄進log檔中
        logging.error(f"儲存excel時發生錯誤: {e}", exc_info=True)

def save_to_mysql(data):
    #建立通道
    conn = None
    #SQL 指令物件
    cursor = None
    try:
        #建立SQL的連線通道
        df = pd.DataFrame(data)
        #建立資料庫
        logging.info(f"連線至MySQL ({DB_HOST}): {DB_PORT}) 檢查資料庫狀態...")
        conn = pymysql.connect(
            host=DB_HOST,
            user=DB_USER,
            password=DB_PASSWORD,
            port=DB_PORT,
            charset=DB_CHARSET,
        )
        cursor = conn.cursor()
        #建立資料庫
        cursor.execute(f"CREATE DATABASE IF NOT EXISTS {DB_NAME} CHARACTER SET {DB_CHARSET} COLLATE utf8mb4_unicode_ci;")
        conn.commit() #commit 告訴mysql 已更新
        cursor.close() #關閉通道
        conn.close() #關閉通道
        conn = None #切斷MySQL連線
        cursor = None #切斷MySQL指令物件

        #使用sqlalchemy建立table建立與資料庫寫入
        #將mysql連線字串轉換成sqlalchemy
        db_url = f"mysql+pymysql://{DB_USER}:{DB_PASSWORD}@{DB_HOST}:{DB_PORT}/{DB_NAME}?charset={DB_CHARSET}" 
        #建立資料庫引擎(engine)
        engine = create_engine(db_url)

        #建立資料表table==>產生通道 -> 執行 --> 關閉通道
        #產生通道
        with engine.begin() as sql_conn:
            #傳遞sql指令
            sql_conn.execute(
                text("""
                CREATE TABLE IF NOT EXISTS tainan_house(
                    Seq BIGINT,
                    鄉鎮市區別 TEXT,
                    區段數合計 TEXT,
                    一般區段數 TEXT,
                    繁榮街道路線價區段數 TEXT,
                    一般區段價最高 TEXT,
                    一般區段價最低 TEXT,
                    最高繁榮街道路線價 TEXT, 
                    最低繁榮街道路線價 TEXT,
                    最高一般路線價 TEXT,
                    最低一般路線價 TEXT,
                    最高宗地地價 TEXT
                    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;
                """)
            )
            #清空TABLE舊資料
            sql_conn.execute(text("TRUNCATE TABLE tainan_house"))
        logging.info("MySQL資料表已建立並清空舊資料")
        #寫入資料
        logging.info("正將資料寫入 MySQL table 中...")
        df.to_sql(
            name = 'house',
            con=engine,
            if_exists='append',
            index=False,
            chunksize=1000,
            method='multi',
        )
        logging.info("資料已成功寫入 MySQL")
    except Exception as e:
        if conn:
            try:
                conn.rollback()
            except Exception:
                pass
        logging.error(f"MySQL 操作錯誤: {e}", exc_info=True)
    finally:
        if cursor:
            cursor.close()
        if conn:
            conn.close()
        logging.info("MySQL 連線已關閉")

def read_from_excel(filename):
    try:
        df = pd.read_excel(filename, engine="openpyxl")
        logging.info(f"已成功讀取excel檔案, 共{len(df)}筆資料")
        return df
    except Exception as e:
        logging.error(f"讀取excel時發生錯誤，失敗原因: {e}")
        return None

def read_from_mysql():
    try:
        db_uri = f"mysql+pymysql://{DB_USER}:{DB_PASSWORD}@{DB_HOST}:{DB_PORT}/{DB_NAME}?charset={DB_CHARSET}" 
        engine = create_engine(db_uri)
        df = pd.read_sql_table("house", con=engine)
        logging.info(f"已成功讀取mysql table資料,共{len(df)}筆資料")
        return df
    except Exception as e:
        logging.error(f"讀取mysql table時發生錯誤，失敗原因: {e}")
        return None

if __name__ == '__main__':
    urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

    data = fetch_data(api_url)
    if data and isinstance(data, dict) and 'data' in data:
        records = data.get("data")

        if records:
            logging.info(f"準備處理{len(records)}筆資料")
            #儲存資料
            save_to_excel(records, EXCEL_FILENAME)
            save_to_mysql(records)

            #儲取資料
            logging.info("執行讀取資料……")
            excel_df = read_from_excel(EXCEL_FILENAME)
            mysql_df = read_from_mysql()

            if excel_df is not None:
                print(f"前2筆記錄: ({excel_df.head(2)})")
            if mysql_df is not None:
                print(f"前2筆記錄: ({excel_df.head(2)})")
        else:
            logging.warning(f"API回傳成功,但data是空的")
    else:
        logging.info(f"程式執行結果")

            
            
        
            




        



