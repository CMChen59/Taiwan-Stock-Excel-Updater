# -*- coding: utf-8 -*-
"""
台灣股票收盤價 Excel 自動更新工具 (Excel Native COM 引擎 + 超高速列區塊寫入)
寫入效能優化說明：
1. 暫停 Excel 自動重算 (xlCalculationManual)：避免每寫入一格就觸發全檔數十個 Sheet 連動重算。
2. 行區塊一次性寫入 (Row-Range Assignment)：將 I~M 欄 (張數、昨收、開盤、最高、最低) 包成陣列一次寫入，降低 80% COM RPC 呼叫負擔。
3. 採用 Microsoft Excel 原生 COM 引擎，100% 絕不毀損繪圖物件 (drawing1.xml)、按鈕或 VBA 巨集。
"""

import os
import sys
import glob
import json
import time
import shutil
import logging
import traceback
import urllib.request
from datetime import datetime
import win32com.client
import win32com.client.gencache
import win32com.client.dynamic

if getattr(sys, 'frozen', False):
    SCRIPT_DIR = os.path.dirname(sys.executable)
else:
    SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

LOG_FILE = os.path.join(SCRIPT_DIR, "update_stocks.log")

logger = logging.getLogger("StockUpdater")
logger.setLevel(logging.INFO)

if logger.hasHandlers():
    logger.handlers.clear()

file_handler = logging.FileHandler(LOG_FILE, mode='a', encoding='utf-8')
file_handler.setFormatter(logging.Formatter('[%(asctime)s] [%(levelname)s] %(message)s', datefmt='%Y-%m-%d %H:%M:%S'))

console_handler = logging.StreamHandler(sys.stdout)
console_handler.setFormatter(logging.Formatter('%(message)s'))

logger.addHandler(file_handler)
logger.addHandler(console_handler)

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(line_buffering=True)

import ssl

def safe_save_wb(wb):
    for attempt in range(5):
        try:
            wb.Save()
            return
        except Exception as e:
            if attempt < 4:
                time.sleep(0.5)
            else:
                raise

def safe_close_wb(wb, save_changes=False):
    for attempt in range(5):
        try:
            wb.Close(save_changes)
            return
        except Exception:
            time.sleep(0.3)

def safe_com_call(func, max_retries=5, delay=0.3):
    for attempt in range(max_retries):
        try:
            return func()
        except Exception as e:
            err_str = str(e)
            if ("-2147418111" in err_str or "rejected" in err_str.lower() or "80010001" in err_str) and attempt < max_retries - 1:
                time.sleep(delay)
            else:
                raise

def fetch_stock_data(tickers):
    if not tickers:
        return {}
    
    channels = []
    for t in tickers:
        t_clean = str(t).strip().zfill(4 if len(str(t).strip()) <= 4 else len(str(t).strip()))
        channels.append(f"tse_{t_clean}.tw")
        channels.append(f"otc_{t_clean}.tw")
    
    url = f"https://mis.twse.com.tw/stock/api/getStockInfo.jsp?ex_ch={'|'.join(channels)}"
    logger.info(f"發送證交所 API 請求: {url}")
    
    req = urllib.request.Request(
        url,
        headers={
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
            'Accept': 'application/json, text/javascript, */*; q=0.01',
        }
    )
    
    stock_dict = {}
    ssl_context = ssl._create_unverified_context()
    try:
        with urllib.request.urlopen(req, timeout=10, context=ssl_context) as response:
            content = response.read().decode('utf-8')
            data = json.loads(content)
            
            if 'msgArray' in data:
                for item in data['msgArray']:
                    code = item.get('c', '').strip()
                    if not code:
                        continue
                    
                    def safe_float(val):
                        try:
                            if val and val != '--':
                                return float(val)
                        except (ValueError, TypeError):
                            pass
                        return None

                    def safe_int(val):
                        try:
                            if val and val != '--':
                                return int(str(val).replace(',', ''))
                        except (ValueError, TypeError):
                            pass
                        return None

                    price = safe_float(item.get('z'))
                    prev_close = safe_float(item.get('y'))
                    open_price = safe_float(item.get('o'))
                    high_price = safe_float(item.get('h'))
                    low_price = safe_float(item.get('l'))
                    volume_lots = safe_int(item.get('v'))

                    if code not in stock_dict or price is not None:
                        stock_dict[code] = {
                            'name': item.get('n', ''),
                            'price': price,
                            'prev_close': prev_close,
                            'open': open_price,
                            'high': high_price,
                            'low': low_price,
                            'volume_lots': volume_lots
                        }
    except Exception as e:
        logger.error(f"抓取證交所 API 時發生例外異常: {e}")
        logger.error(traceback.format_exc())

    missing_tickers = [t for t in tickers if t not in stock_dict or stock_dict[t]['price'] is None]
    if missing_tickers:
        logger.info(f"嘗試使用 Yahoo Finance API 備援查詢股票: {', '.join(missing_tickers)}")
        for t in missing_tickers:
            fallback_data = fetch_yahoo_fallback(t)
            if fallback_data:
                stock_dict[t] = fallback_data

    return stock_dict

def fetch_yahoo_fallback(ticker):
    ssl_context = ssl._create_unverified_context()
    ticker_clean = str(ticker).strip().zfill(4 if len(str(ticker).strip()) <= 4 else len(str(ticker).strip()))
    suffix_list = ['.TW', '.TWO']
    
    for suffix in suffix_list:
        symbol = f"{ticker_clean}{suffix}"
        url = f"https://query1.finance.yahoo.com/v8/finance/chart/{symbol}?interval=1d&range=1d"
        req = urllib.request.Request(
            url,
            headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'}
        )
        try:
            with urllib.request.urlopen(req, timeout=5, context=ssl_context) as resp:
                data = json.loads(resp.read().decode('utf-8'))
                result = data.get('chart', {}).get('result', [])
                if result:
                    meta = result[0].get('meta', {})
                    price = meta.get('regularMarketPrice')
                    prev_close = meta.get('chartPreviousClose') or meta.get('previousClose')
                    open_p = meta.get('regularMarketDayHigh')
                    high_p = meta.get('regularMarketDayHigh')
                    low_p = meta.get('regularMarketDayLow')
                    vol_shares = meta.get('regularMarketVolume')
                    vol_lots = int(vol_shares / 1000) if vol_shares else None

                    indicators = result[0].get('indicators', {}).get('quote', [{}])[0]
                    if indicators.get('open'):
                        open_p = indicators['open'][-1]
                    
                    if price is not None:
                        logger.info(f"Yahoo API 備援成功獲取 {ticker_clean}: 成交={price}, 張數={vol_lots}")
                        return {
                            'name': meta.get('shortName', ticker),
                            'price': float(price) if price else None,
                            'prev_close': float(prev_close) if prev_close else None,
                            'open': float(open_p) if open_p else None,
                            'high': float(high_p) if high_p else None,
                            'low': float(low_p) if low_p else None,
                            'volume_lots': vol_lots
                        }
        except Exception as e:
            logger.warning(f"Yahoo API 查詢 {symbol} 失敗: {e}")
            continue
    return None

def find_target_excel_files(directory):
    pattern = os.path.join(directory, "*.xls*")
    all_files = glob.glob(pattern)
    target_files = []
    
    for fpath in all_files:
        fname = os.path.basename(fpath)
        if fname.startswith("~$"):
            continue
        
        if ("定存股" in fname) or ("試算表" in fname) or ("股票" in fname):
            target_files.append(os.path.abspath(fpath))
            
    return target_files

def clear_gen_py():
    try:
        import win32com.client.gencache
        for mod_name in list(sys.modules.keys()):
            if "win32com.gen_py" in mod_name:
                del sys.modules[mod_name]
        gen_py_dir = win32com.client.gencache.GetGeneratePath()
        if gen_py_dir and os.path.exists(gen_py_dir):
            shutil.rmtree(gen_py_dir, ignore_errors=True)
            logger.info(f"已自動清除損壞之 pywin32 快取目錄: {gen_py_dir}")
    except Exception as e:
        logger.warning(f"清理 gen_py 快取時發生例外: {e}")

def get_excel_app():
    for attempt in range(2):
        try:
            try:
                raw_app = win32com.client.GetActiveObject("Excel.Application")
                app = win32com.client.gencache.EnsureDispatch(raw_app)
                logger.info("已附加連線至現有運作中之 Excel.Application 實例 (EnsureDispatch)")
                return app, False
            except Exception:
                app = win32com.client.gencache.EnsureDispatch("Excel.Application")
                app.Visible = False
                app.DisplayAlerts = False
                app.ScreenUpdating = False
                logger.info("已建立全新背景 Excel.Application COM 實例 (EnsureDispatch)")
                return app, True
        except (AttributeError, Exception) as e:
            logger.warning(f"win32com gencache 發生例外 ({e})，正在嘗試清除 gen_py 重置快取...")
            clear_gen_py()

    try:
        try:
            raw_app = win32com.client.GetActiveObject("Excel.Application")
            app = win32com.client.dynamic.Dispatch(raw_app)
            logger.info("已附加連線至現有運作中之 Excel.Application 實例 (Dynamic Dispatch 備援)")
            return app, False
        except Exception:
            app = win32com.client.dynamic.Dispatch("Excel.Application")
            app.Visible = False
            app.DisplayAlerts = False
            app.ScreenUpdating = False
            logger.info("已建立全新背景 Excel.Application COM 實例 (Dynamic Dispatch 備援)")
            return app, True
    except Exception as e:
        logger.error(f"建立或連線 Excel.Application COM 實例完全失敗: {e}")
        raise

def update_excel_file_com(excel_app, file_path, stock_data_dict):
    fname = os.path.basename(file_path)
    logger.info(f"[正在更新] 檔案: {fname}")
    
    wb = None
    is_already_open = False
    
    try:
        try:
            for open_wb in excel_app.Workbooks:
                if open_wb.FullName.lower() == file_path.lower():
                    wb = open_wb
                    is_already_open = True
                    logger.info(f"  └ 檔案 {fname} 目前已在 Excel 中開啟，將進行現場即時寫入")
                    break
        except Exception:
            pass

        if not wb:
            wb = safe_com_call(lambda: excel_app.Workbooks.Open(file_path, 0, False))

        # 暫停自動計算以提升寫入速度
        old_calc = None
        try:
            old_calc = excel_app.Calculation
            excel_app.Calculation = -4135 # xlCalculationManual
        except Exception:
            pass

        ws = None
        for sheet in wb.Sheets:
            if sheet.Name == "一覽表":
                ws = sheet
                break
                
        if not ws:
            logger.warning(f"  └ [跳過] 檔案 {fname} 內找不到 '一覽表' 工作表")
            if not is_already_open:
                safe_com_call(lambda: wb.Close(False))
            return False

        now_str = datetime.now().strftime("%Y/%m/%d %H:%M:%S")
        safe_com_call(lambda: ws.Cells(3, 2).__setattr__('Value', f"※更新時間：{now_str}"))
        
        updated_count = 0
        for r in range(5, 51):
            code_val = safe_com_call(lambda: ws.Cells(r, 3).Value)
            if code_val is None:
                continue
                
            code_str = str(code_val).strip()
            if code_str.endswith('.0'):
                code_str = code_str[:-2]
                
            if code_str.isdigit():
                code_str = code_str.zfill(4 if len(code_str) <= 4 else len(code_str))
                
            if code_str in stock_data_dict:
                info = stock_data_dict[code_str]
                
                # E欄：成交價
                if info['price'] is not None:
                    safe_com_call(lambda: ws.Cells(r, 5).__setattr__('Value', info['price']))
                
                # I~M欄：張數, 昨收, 開盤, 最高, 最低 (一次性 Range 區塊寫入)
                sub_vals = [[
                    info['volume_lots'] if info['volume_lots'] is not None else ws.Cells(r, 9).Value,
                    info['prev_close'] if info['prev_close'] is not None else ws.Cells(r, 10).Value,
                    info['open'] if info['open'] is not None else ws.Cells(r, 11).Value,
                    info['high'] if info['high'] is not None else ws.Cells(r, 12).Value,
                    info['low'] if info['low'] is not None else ws.Cells(r, 13).Value,
                ]]
                safe_com_call(lambda: ws.Range(ws.Cells(r, 9), ws.Cells(r, 13)).__setattr__('Value', sub_vals))
                
                stock_name = info.get('name') or ws.Cells(r, 4).Value
                logger.info(f"  ├─ 列 {r:2d} [{code_str} {stock_name}]: 成交={info['price']}, 昨收={info['prev_close']}, 張數={info['volume_lots']}")
                updated_count += 1

        # 恢復自動計算
        if old_calc is not None:
            try:
                excel_app.Calculation = old_calc
            except Exception:
                pass

        safe_save_wb(wb)
        if not is_already_open:
            safe_close_wb(wb, False)
        logger.info(f"  └ [成功] 檔案 {fname} 已完成 {updated_count} 支股票數據寫入與存檔！")
        return True
    except Exception as e:
        logger.error(f"  └ [錯誤] 檔案 {fname} 更新失敗: {e}")
        logger.error(traceback.format_exc())
        if wb and not is_already_open:
            safe_close_wb(wb, False)
        return False

def collect_all_tickers(target_files):
    import openpyxl
    tickers = set()
    for fpath in target_files:
        try:
            wb = openpyxl.load_workbook(fpath, read_only=True)
            if "一覽表" in wb.sheetnames:
                ws = wb["一覽表"]
                for r in range(5, 51):
                    val = ws.cell(row=r, column=3).value
                    if val is not None:
                        val_str = str(val).strip()
                        if val_str.endswith('.0'):
                            val_str = val_str[:-2]
                        if val_str.isdigit():
                            val_str = val_str.zfill(4 if len(val_str) <= 4 else len(val_str))
                            tickers.add(val_str)
            wb.close()
        except Exception as e:
            logger.warning(f"讀取 {os.path.basename(fpath)} 股票代號時發生錯誤: {e}")
    return sorted(list(tickers))

def main():
    logger.info("======================================================================")
    logger.info("      台灣股票收盤價 Excel 自動更新工具 (Excel Native COM 引擎)")
    logger.info("======================================================================")
    logger.info(f"Log 記錄檔位置: {LOG_FILE}")
    
    target_files = find_target_excel_files(SCRIPT_DIR)
    
    if not target_files:
        logger.warning("[資訊] 當前資料夾下未找到包含『一覽表』的股票試算表。")
        input("按 Enter 鍵結束...")
        return

    selected_files = []
    if len(sys.argv) > 1:
        arg = sys.argv[1].strip()
        if arg in ["--all", "all", "0"]:
            selected_files = target_files
        else:
            for f in target_files:
                if arg.lower() in os.path.basename(f).lower():
                    selected_files.append(f)
            if not selected_files:
                selected_files = target_files
    else:
        logger.info("實時掃描到的股票試算表清單：")
        logger.info("  [0] 更新所有股票試算表 (共 {} 個檔案)".format(len(target_files)))
        for idx, fpath in enumerate(target_files, 1):
            logger.info(f"  [{idx}] 僅更新: {os.path.basename(fpath)}")
        logger.info("======================================================================")
        logger.info("(已自動過濾非股票檔案如 '支出明細.xlsx' 及 Excel 暫存檔)")
        
        user_choice = input(f"請輸入選項 [0-{len(target_files)}] (直接按 Enter 預設全部更新): ").strip()
        
        if user_choice == "" or user_choice == "0":
            selected_files = target_files
        else:
            try:
                choice_idx = int(user_choice)
                if 1 <= choice_idx <= len(target_files):
                    selected_files = [target_files[choice_idx - 1]]
                else:
                    logger.info("[提示] 無效的選項，自動選擇更新全部檔案。")
                    selected_files = target_files
            except ValueError:
                logger.info("[提示] 輸入非數字，自動選擇更新全部檔案。")
                selected_files = target_files

    logger.info("[步驟 1/2] 正在收集目標股票代號並向證交所 API 查詢收盤數據...")
    all_tickers = collect_all_tickers(selected_files)
    logger.info(f"  ├─ 需查詢的股票代號 ({len(all_tickers)} 支): {', '.join(all_tickers)}")
    
    start_time = time.time()
    stock_data = fetch_stock_data(all_tickers)
    elapsed = time.time() - start_time
    logger.info(f"  └─ API 數據抓取完成！耗時: {elapsed:.2f} 秒")

    excel_app = None
    is_newly_created = False

    try:
        excel_app, is_newly_created = get_excel_app()
        logger.info("[步驟 2/2] 正在經由 Excel 原生引擎寫入儲存格 (100% 原生保留圖片與公式)...")
        success_count = 0
        for fpath in selected_files:
            if update_excel_file_com(excel_app, fpath, stock_data):
                success_count += 1
                
        logger.info("======================================================================")
        logger.info(f"  全數更新完成！成功更新 {success_count} / {len(selected_files)} 個檔案。")
        logger.info("======================================================================")
    except Exception as global_e:
        logger.critical(f"全局執行發生嚴重例外錯誤: {global_e}")
        logger.critical(traceback.format_exc())
    finally:
        if excel_app and is_newly_created:
            try:
                safe_com_call(lambda: excel_app.Quit())
            except Exception:
                pass
        
    if len(sys.argv) == 1:
        input("按 Enter 鍵關閉視窗...")

if __name__ == "__main__":
    main()
