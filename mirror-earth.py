import requests
from bs4 import BeautifulSoup
import pandas as pd
import time
import os
import sys
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import date, timedelta
from tqdm import tqdm
import itertools 
import re 

# =========================================================================
# 配置 (CONFIG)
# =========================================================================
# 动态获取脚本所在目录，确保文件路径是相对路径
SCRIPT_DIR = os.path.dirname(os.path.abspath(sys.argv[0]))

EXCEL_FILE_PATH = os.path.join(SCRIPT_DIR, 'weather_district_id.xlsx') 
SHEET_NAME = 'zzw' 
MAX_THREADS = 10 
FINAL_OUTPUT_PATH = os.path.join(SCRIPT_DIR, 'weather.csv') 

# 列定义
# 修改：新的最终输出列名，包含单位
FINAL_OUTPUT_COLUMNS_WITH_NAME = [
    '城市名', 
    '日期', 
    '最高温(℃)', 
    '最低温(℃)', 
    '天气', 
    '风力风向', 
    '日照时长(h)'
] 
# 抓取和处理阶段使用的列名 (不含单位)
SCRAPE_PROCESSING_COLUMNS_WITH_CODE = ['城市代码', '日期', '最高温', '最低温', '天气', '风力风向', '日照时长']


# =========================================================================
# 辅助函数 (UTILITIES)
# =========================================================================

def get_required_days(year, month):
    """计算给定年月的历史数据应包含多少天 (排除今天及未来)。"""
    today = date.today()
    if year > today.year or (year == today.year and month > today.month):
        return 0 
    elif year == today.year and month == today.month:
        return today.day - 1
    else:
        if month == 2:
            return 29 if year % 4 == 0 and (year % 100 != 0 or year % 400 == 0) else 28
        elif month in [4, 6, 9, 11]:
            return 30
        else:
            return 31

def generate_month_range(start_y, end_y, start_m, end_m):
    """统一的生成器：生成 (year, month) 元组，用于任务计数和执行。"""
    yesterday = date.today() - timedelta(days=1)
    
    for year in range(start_y, end_y + 1):
        start_m_curr = start_m if year == start_y else 1
        end_m_curr = end_m if year == end_y else 12
        
        for month in range(start_m_curr, end_m_curr + 1):
            # 排除未来月份
            if year > yesterday.year or (year == yesterday.year and month > yesterday.month):
                continue
            # 排除 required_days <= 0 的情况
            if get_required_days(year, month) > 0:
                 yield year, month

def generate_expected_dates(start_year, end_year, start_month, end_month):
    """生成查询范围内的所有期望日期列表 (用于完整性校验)。"""
    dates_list = []
    for year, month in generate_month_range(start_year, end_year, start_month, end_month):
        last_day = get_required_days(year, month)
        for day in range(1, last_day + 1):
            dates_list.append(date(year, month, day).strftime('%Y-%m-%d'))
    return dates_list

def get_date_input(prompt, limit_date_check):
    """通用输入校验函数，处理 YYYYMM 格式的日期输入。"""
    while True:
        date_str = input(prompt).strip()
        if not re.fullmatch(r'\d{6}', date_str):
            print("❌ 输入错误：日期格式必须是 YYYYMM (例如 '202501')。请重新输入。")
            continue
        try:
            year, month = int(date_str[:4]), int(date_str[4:])
            if not 1 <= month <= 12:
                print("❌ 输入错误：月份必须在 01 到 12 之间。请重新输入。")
                continue
            input_date = date(year, month, 1)
            if input_date > limit_date_check:
                print(f"❌ 逻辑错误：输入日期 {date_str} 超过最大可爬取月份 {limit_date_check.strftime('%Y年%m月')}。请重新输入。")
                continue
            return year, month
        except ValueError:
            print("❌ 输入错误：年份或月份不合法。请重新输入。")
            
def clean_numeric_column(series):
    """新增函数：从包含单位的字符串中提取纯数字。"""
    # 匹配可选的负号/减号，后跟一个或多个数字，小数点，以及可选的更多数字
    # 例如: '32℃' -> '32', '-5℃' -> '-5', '8.5h' -> '8.5'
    return series.astype(str).str.extract(r'(-?\d+\.?\d*)', expand=False).fillna('')

# =========================================================================
# 文件操作 (FILE IO) 
# =========================================================================

def read_city_info_from_excel(file_path, sheet_name):
    """从文件读取城市名和代码。"""
    print(f"尝试从文件读取城市信息: {file_path}")
    if not os.path.exists(file_path):
        print(f"错误: 找不到文件 {file_path}。请检查路径是否正确。")
        return [], {}
    try:
        # 优化读取和清理
        df = pd.read_excel(file_path, sheet_name=sheet_name, header=None, usecols=[0, 1])
        df[0] = df[0].astype(str).str.strip()
        df[1] = df[1].astype(str).str.strip()
        start_index = 1 if not df.empty and df.iloc[0, 1].upper() == '编码' else 0
        df_valid = df.iloc[start_index:].dropna(subset=[0, 1])
        df_valid = df_valid[df_valid[1].str.match(r'^\d{5,6}$')]

        city_codes_list = df_valid[1].tolist()
        code_to_name_map = dict(zip(df_valid[1], df_valid[0]))
        
        print(f"成功读取 {len(city_codes_list)} 个城市代码及其名称。")
        return city_codes_list, code_to_name_map
    except Exception as e:
        print(f"读取文件失败: {e}")
        return [], {}

def load_existing_data(file_path, code_to_name_map):
    """读取本地已存在的 CSV 文件。"""
    if not os.path.exists(file_path) or os.path.getsize(file_path) == 0:
        return pd.DataFrame()
        
    try:
        # 定义本地文件列名与处理列名的映射关系（注意：本地文件可能包含带单位的表头）
        # 这里使用新的 FINAL_OUTPUT_COLUMNS_WITH_NAME 作为读取时的列名，以匹配旧的/新的文件格式
        read_columns = FINAL_OUTPUT_COLUMNS_WITH_NAME
        
        df_local = pd.read_csv(file_path, encoding='utf-8-sig', usecols=read_columns) 
        
        # 统一列名为处理阶段的名称（不带单位），方便后续去重合并
        reverse_rename_map = {
            '最高温(℃)': '最高温',
            '最低温(℃)': '最低温',
            '日照时长(h)': '日照时长'
        }
        df_local.rename(columns=reverse_rename_map, inplace=True)
        
        # 1. 清洗数据，确保它们是纯数字，即使它们是从带单位的列名中读入的
        df_local['最高温'] = clean_numeric_column(df_local['最高温'])
        df_local['最低温'] = clean_numeric_column(df_local['最低温'])
        df_local['日照时长'] = clean_numeric_column(df_local['日照时长'])
        
        
        # 2. 转换城市名回城市代码 (用作内部处理键)
        code_to_name_map_reverse = {v: k for k, v in code_to_name_map.items()}
        df_local.rename(columns={'城市名': '城市代码'}, inplace=True)
        df_local['城市代码'] = df_local['城市代码'].astype(str).str.strip().map(code_to_name_map_reverse)
        
        df_local.dropna(subset=['城市代码'], inplace=True) 
        df_local['日期'] = pd.to_datetime(df_local['日期'], errors='coerce').dt.strftime('%Y-%m-%d')
        df_local.dropna(subset=['日期'], inplace=True) 
             
        # 确保列顺序和名称与 SCRAPE_PROCESSING_COLUMNS_WITH_CODE 一致
        return df_local.sort_values(by=['城市代码', '日期']).reset_index(drop=True)[SCRAPE_PROCESSING_COLUMNS_WITH_CODE]
    except Exception as e:
        print(f"⚠️ 读取本地文件 {file_path} 失败或格式不匹配: {e}. 将从零开始爬取。")
        return pd.DataFrame()

# =========================================================================
# 爬虫核心 (CORE SCRAPER) 
# =========================================================================

def get_monthly_weather(city_code, year, month):
    """尝试爬取指定城市、年份和月份的历史天气数据。"""
    url = f'https://mirror-earth.com/wea_history/{city_code}/{str(year)}-{str(month).zfill(2)}'
    headers = {'User-Agent': 'Mozilla/5.0'}
    COLUMNS = ['日期', '最高温', '最低温', '天气', '风力风向', '日照时长']
    try:
        response = requests.get(url, headers=headers, timeout=15)
        response.raise_for_status() 
    except requests.exceptions.RequestException:
        return None
        
    soup = BeautifulSoup(response.content, 'html.parser')
    table = soup.find('table', class_='table-striped') 
    if not table:
        return None
        
    all_rows = [[col.text.strip() for col in row.find_all('td')] for row in table.find_all('tr')[1:]]
    if not all_rows:
        return None
        
    df = pd.DataFrame(all_rows, columns=COLUMNS)
    
    # 日期修正逻辑 
    df['日期'] = df['日期'].astype(str).str.split(' ').str[0]
    date_day_series = df['日期'].str.extract(r'(\d+)$', expand=False).str.zfill(2)
    df['日期'] = f'{year}-{str(month).zfill(2)}-' + date_day_series.astype(str)
    df['日期'] = pd.to_datetime(df['日期'], errors='coerce').dt.strftime('%Y-%m-%d')
    
    df.insert(0, '城市代码', city_code)
    return df

def scrape_city_weather_range(city_code, start_year, end_year, start_month, end_month, df_local_code, code_to_name_map, pbar):
    """城市 worker 函数：处理指定范围内的所有月份任务。"""
    all_data = []
    city_name = code_to_name_map.get(city_code, city_code)
    
    # 准备城市本地缓存 (用于快速查找)
    df_local_city_dt = pd.DataFrame(columns=SCRAPE_PROCESSING_COLUMNS_WITH_CODE + ['日期_DT']).astype({'日期_DT': 'datetime64[ns]'})
    if not df_local_code.empty:
        df_local_city = df_local_code[df_local_code['城市代码'] == city_code].copy()
        df_local_city_dt = df_local_city
        df_local_city_dt['日期_DT'] = pd.to_datetime(df_local_city_dt['日期'], errors='coerce')
    
    # 迭代月份任务 
    for year, month in generate_month_range(start_year, end_year, start_month, end_month):
        
        required_days = get_required_days(year, month)
        
        # 1. 检查缓存
        df_city_month_local = df_local_city_dt[
            (df_local_city_dt['日期_DT'].dt.year == year) & 
            (df_local_city_dt['日期_DT'].dt.month == month)
        ].drop(columns='日期_DT')
        
        if len(df_city_month_local) >= required_days:
            status_msg = f"🔍 缓存命中: {city_name} ({city_code}) - {year}年{month}月"
            all_data.append(df_city_month_local[SCRAPE_PROCESSING_COLUMNS_WITH_CODE])
        else:
            # 2. 网络拉取
            status_msg = f"💻 正在拉取: {city_name} ({city_code}) - {year}年{month}月"
            df_month = get_monthly_weather(city_code, year, month)
            if df_month is not None and not df_month.empty:
                all_data.append(df_month)
            else:
                 status_msg = f"❌ 失败/空: {city_name} ({city_code}) - {year}年{month}月 (无数据返回)"
                 
        pbar.set_description(f"🌐 正在校验和拉取数据: {status_msg}")
        pbar.update(1) 
                 
    return pd.concat(all_data, ignore_index=True) if all_data else None

# =========================================================================
# 主程序入口 (MAIN EXECUTION)
# =========================================================================
if __name__ == '__main__':
    
    print("--------------------------------------")
    print(" 🚀 脚本版本: V4.3 (已包含数据清洗与表头更新)")
    print("--------------------------------------")
    
    # 1. 初始化和加载数据
    CITY_CODES_LIST, CODE_TO_NAME_MAP = read_city_info_from_excel(EXCEL_FILE_PATH, SHEET_NAME)
    if not CITY_CODES_LIST: sys.exit(1)
    NUM_CITIES = len(CITY_CODES_LIST)
    print(f"准备处理 {NUM_CITIES} 个城市/地区的历史数据。")

    print("=" * 50)
    print(f"文件将保存到路径: {FINAL_OUTPUT_PATH}")
    
    DF_LOCAL_CACHE_WITH_CODE = load_existing_data(FINAL_OUTPUT_PATH, CODE_TO_NAME_MAP)
    
    if not DF_LOCAL_CACHE_WITH_CODE.empty:
        print(f"✅ 已加载本地 {len(DF_LOCAL_CACHE_WITH_CODE)} 条历史数据用于校验和去重。")
        print("ℹ️ 操作策略：优先使用本地缓存，对查询范围内缺失的数据进行网络爬取和更新。")
    else:
        print("ℹ️ 本地文件不存在或无法读取，将全部进行网络爬取。")
    print("=" * 50)

    # 2. 日期输入和任务计算
    yesterday = date.today() - timedelta(days=1)
    MAX_DATE_CHECK = date(yesterday.year, yesterday.month, 1)

    print(f"ℹ️ 历史数据最大可爬取月份为 {MAX_DATE_CHECK.strftime('%Y年%m月')}。")
    start_year, start_month = get_date_input(f"请输入爬取起始年月份 (YYYYMM): ", MAX_DATE_CHECK)
    end_year, end_month = get_date_input(f"请输入爬取结束年月份 (YYYYMM): ", MAX_DATE_CHECK)
    
    print("-" * 50)
    start_dt = date(start_year, start_month, 1)
    end_dt = date(end_year, end_month, 1)
    if start_dt > end_dt:
        print("❌ 逻辑错误：起始日期不能晚于结束日期。脚本退出。")
        sys.exit(1)
        
    time_period_str = f"{start_year}年{start_month}月 到 {end_year}年{end_month}月"
    print(f"✅ 爬取范围确定：从 {time_period_str}。")
    
    TOTAL_TASKS = len(list(generate_month_range(start_year, end_year, start_month, end_month))) * NUM_CITIES
    
    if TOTAL_TASKS == 0:
         print("❌ 目标时间范围内没有可爬取的月份，脚本退出。")
         sys.exit(0)
    print(f"✅ 总计需要处理 {TOTAL_TASKS} 个城市-月份任务。")
    print("-" * 50)
    
    # 3. 多线程并发处理
    all_final_data = []
    
    with ThreadPoolExecutor(max_workers=MAX_THREADS) as executor:
        with tqdm(total=TOTAL_TASKS, 
                  desc="🌐 正在校验和拉取数据", 
                  unit="个城市月份任务", 
                  bar_format="{desc}: {percentage:.1f}%|{bar}| {n_fmt}/{total_fmt} [{elapsed}<{remaining}, {rate_fmt}]") as pbar:
            
            futures = {
                executor.submit(
                    scrape_city_weather_range, 
                    code, start_year, end_year, start_month, end_month, DF_LOCAL_CACHE_WITH_CODE, CODE_TO_NAME_MAP, pbar
                ): code for code in CITY_CODES_LIST
            }
            
            for future in as_completed(futures):
                city_code = futures[future]
                city_name = CODE_TO_NAME_MAP.get(city_code, city_code)
                try:
                    df_city_data = future.result()
                    if df_city_data is not None:
                        all_final_data.append(df_city_data)
                    else:
                         pbar.write(f"ℹ️ {city_name} ({city_code}) 任务结束，查询范围内无新数据或无历史数据。")
                except Exception as e:
                    pbar.write(f"❌ 城市 {city_name} ({city_code}) 处理发生错误: {e}")
                
            pbar.set_description("✅ 所有城市数据处理完成")


    # 4. 合并、校验与保存
    integrity_report_str = "" 
    
    if all_final_data or not DF_LOCAL_CACHE_WITH_CODE.empty:
        
        df_new_processed_data = pd.concat(all_final_data, ignore_index=True)
        final_combined_df_code = df_new_processed_data.copy()

        if not DF_LOCAL_CACHE_WITH_CODE.empty:
            df_new_keys = df_new_processed_data[['城市代码', '日期']].drop_duplicates()
            df_new_keys['is_new'] = True
            
            # 将本地未被新爬取数据覆盖的历史数据重新纳入
            df_unqueried_local = DF_LOCAL_CACHE_WITH_CODE.merge(
                df_new_keys, on=['城市代码', '日期'], how='left'
            )
            
            df_unqueried_local = df_unqueried_local[df_unqueried_local['is_new'].isna()].drop(columns='is_new')
            df_unqueried_local = df_unqueried_local[SCRAPE_PROCESSING_COLUMNS_WITH_CODE] 

            final_combined_df_code = pd.concat([df_unqueried_local, df_new_processed_data], ignore_index=True)
            
        final_combined_df_code.drop_duplicates(subset=['城市代码', '日期'], inplace=True)
        
        # 显式内存优化：手动释放对旧缓存和中间数据框的引用
        if 'DF_LOCAL_CACHE_WITH_CODE' in locals():
            del DF_LOCAL_CACHE_WITH_CODE
        del df_new_processed_data 
        if 'df_unqueried_local' in locals():
            del df_unqueried_local
        
        # 完整性校验
        EXPECTED_DATES_LIST = generate_expected_dates(start_year, end_year, start_month, end_month)
        if EXPECTED_DATES_LIST:
            expected_keys = pd.DataFrame(list(itertools.product(CITY_CODES_LIST, EXPECTED_DATES_LIST)), columns=['城市代码', '日期'])
            actual_keys = final_combined_df_code[['城市代码', '日期']].drop_duplicates()
            df_missing = expected_keys.merge(actual_keys.assign(is_present=True), on=['城市代码', '日期'], how='left')
            df_missing = df_missing[df_missing['is_present'].isna()]
            
            if not df_missing.empty:
                df_missing['城市名'] = df_missing['城市代码'].map(CODE_TO_NAME_MAP)
                df_missing['年份月份'] = pd.to_datetime(df_missing['日期'], errors='coerce').dt.strftime('%Y年%m月')
                df_report = df_missing.groupby(['城市名', '年份月份']).size().reset_index(name='缺失天数')
                df_report.rename(columns={'城市名': '城市/区域'}, inplace=True)
                
                integrity_report_str = "\n" + "#" * 60 + "\n⚠️ **数据完整性校验报告：以下数据缺失！**\n"
                integrity_report_str += df_report.to_string(index=False)
                integrity_report_str += "\n" + "#" * 60
            else:
                integrity_report_str = "✅ **数据完整性校验通过：所有查询范围内的数据均存在。**"

        # 转换为最终格式并保存
        if not final_combined_df_code.empty:
            
            # --- START: 最终数据处理（数据清洗和列名修改） ---
            
            # 1. 清洗数据，提取纯数字 (在合并后对所有数据进行一次清洗)
            final_combined_df_code['最高温'] = clean_numeric_column(final_combined_df_code['最高温'])
            final_combined_df_code['最低温'] = clean_numeric_column(final_combined_df_code['最低温'])
            final_combined_df_code['日照时长'] = clean_numeric_column(final_combined_df_code['日照时长'])
            
            final_combined_df = final_combined_df_code.copy()
            
            # 2. 城市代码 -> 城市名 (先替换值)
            final_combined_df.rename(columns={'城市代码': '城市名'}, inplace=True)
            final_combined_df['城市名'] = final_combined_df['城市名'].map(CODE_TO_NAME_MAP)
            
            # 3. 列名添加单位
            column_rename_map = {
                '最高温': '最高温(℃)',
                '最低温': '最低温(℃)',
                '日照时长': '日照时长(h)'
            }
            final_combined_df.rename(columns=column_rename_map, inplace=True)
            
            # 4. 确保列的顺序和名称符合 FINAL_OUTPUT_COLUMNS_WITH_NAME
            final_combined_df = final_combined_df[FINAL_OUTPUT_COLUMNS_WITH_NAME]
            
            # --- END: 最终数据处理（数据清洗和列名修改） ---
            
            final_combined_df['日期'] = final_combined_df['日期'].astype(str).str.split(' ').str[0]
            
            # 保存重试逻辑
            while True:
                try:
                    final_combined_df.to_csv(FINAL_OUTPUT_PATH, index=False, encoding='utf-8-sig')
                    
                    num_cities = final_combined_df['城市名'].nunique()
                    num_dates = final_combined_df['日期'].nunique()
                    print("\n" + "=" * 60)
                    print("✅ **最终数据保存成功！**")
                    print(f"最终总记录数: **{len(final_combined_df)}** 条。涵盖城市总数: **{num_cities}** 个，独立日期总数: **{num_dates}** 天。")
                    print(f"文件保存路径: **{FINAL_OUTPUT_PATH}**")
                    print("=" * 60)
                    print("\n" + integrity_report_str) 
                    break 
                except PermissionError:
                    print(f"\n⚠️ **保存失败：文件 {FINAL_OUTPUT_PATH} 被占用！** 请关闭文件后按 ENTER 键重试...")
                    input()
                except Exception as e:
                    print(f"❌ 发生未知错误，无法保存文件: {e}")
                    sys.exit(1)
    else:
        print("\n❌ 爬取过程没有获取到任何有效数据，且本地无缓存数据。")
