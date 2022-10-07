import glob
import pandas as pd
import openpyxl as px
import os
import sys

def get_dirname():
    if getattr(sys, 'frozen', False):
        dirname = os.path.dirname((sys.executable))
    else:
        dirname = os.path.dirname(__file__)
    return dirname

# フォルダにあるエクセルファイルの絶対パスのリストを取得
def get_xlsx_file_paths(folder_path):
    file_path = folder_path + '\*.xlsx'
    return glob.glob(file_path)

# 成績を生成する選手を取得
def get_players_name(path: str):
    wb = px.load_workbook(path + "\選手登録.xlsx")
    ws = wb.worksheets[0]
    for col in ws.columns:
        players_name = []
        for cell in col:
            players_name.append(cell.value)

    return players_name

# 打撃成績を計算
def calc_bat_record(df: pd.DataFrame):
    divisor = df['打数'] + df['四球'] + df['死球'] + df['犠飛']
    wOBA_divided = 0.692*df['四球'] + 0.73*df['死球'] + 0.966*df['失策'] + 0.865 *df['単打'] + 1.334*df['二塁打'] + 1.725*df['三塁打'] + 2.065*df['本塁打']
    df['盗塁成功率'] = df['盗塁'] / df['盗塁企画']
    df['打率'] = df['安打'] / df['打数']
    df['出塁率'] = (df['安打'] + df['四球'] + df['死球']) / divisor
    df['長打率'] = df['塁打'] / df['打数']
    df['OPS'] = df['出塁率'] + df['長打率']
    df['BB/K'] = df['四球'] / df['三振']
    df['wOBA'] = wOBA_divided / divisor
    df.fillna(0, inplace=True)
    
# 投手成績を計算
def calc_pitch_record(df: pd.DataFrame):
    IP = df['奪アウト数']/3                 # 投球回（Innings pitched / IP）
    q, mod = divmod(df['奪アウト数'], 3)    # 奪アウト数　→　投球回数
    df['奪アウト数'] = q + 0.1*mod
    df.rename(columns={'奪アウト数':'投球回'}, inplace=True)
    df['防御率'] = df['自責点']*7 / IP
    df['奪三振率'] = df['奪三振']*7 / IP
    df['K%'] = df['奪三振'] / df['打者数']
    df['BB%'] = df['与四球'] / df['打者数']
    df['被打率'] = df['被安打'] / df['打者数']
    df['WHIP'] = (df['与四球'] + df['被安打']) / IP
    df['投球数/回'] = df['投球数']/IP
    df.fillna(0, inplace=True)

# 率を小数第三位までに設定    
def set_rate_format(ws: px.Workbook.worksheets, beginning: int):
    for col in ws.iter_cols(min_row=2, min_col=beginning):
        for cell in col:
            cell.number_format = "0.000"
            