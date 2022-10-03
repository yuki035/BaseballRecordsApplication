import os
import glob
import pandas as pd
import datetime
import subprocess
import openpyxl as px
from openpyxl.xml.constants import NAMESPACES
from openpyxl.styles import Alignment, PatternFill

# フォルダにあるエクセルファイルの絶対パスのリストを取得
def get_xlsx_file_paths(folder_path):
    file_path = folder_path + '\*.xlsx'
    return glob.glob(file_path)
    
# 試合結果の全データを打撃/投手別で統合
def concat_games(paths, df_bat, df_pitch):
    for path in paths:
        df_bat_read_excel   = pd.read_excel(path, sheet_name='打撃成績', index_col=0)
        df_pitch_read_excel = pd.read_excel(path, sheet_name='投手成績', index_col=0)
        df_bat   = pd.concat([df_bat, df_bat_read_excel])
        df_pitch = pd.concat([df_pitch, df_pitch_read_excel])
    return df_bat, df_pitch

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
def set_rate_format(beginning: int, ws: px.Workbook.worksheets):
    for col in ws.iter_cols(min_row=2, min_col=beginning):
        for cell in col:
            cell.number_format = "0.000"

# 列の幅を設定
def set_column_width(ws: px.Workbook.worksheets, beginning_rate: int):
    name_width = 14
    width = 5
    ws.column_dimensions['A'].width = name_width
    for col in ws.iter_cols(min_col=2):
        column_num = col[0].column
        column_str = col[0].column_letter
        if column_num == beginning_rate:
            width = 7
        ws.column_dimensions[column_str].width = width
        
# 一行目を縦書きに設定
def set_vertical_writing_row1(ws: px.Workbook.worksheets):
    for row in ws.iter_rows(min_col=2,max_row=1):
        for cell in row:
            cell.alignment = Alignment(vertical='center', textRotation=255)
    
#　セルの背景色を設定
def set_backgroud_color(ws: px.Workbook.worksheets, title_color: str, name_color: str):
    max_row = ws.max_row
    # name
    for col in ws.iter_cols(min_row=2, max_row=max_row, max_col=1):  
        for cell in col:
            cell.fill = PatternFill(patternType = 'solid', fgColor=name_color)
    # title
    for row in ws.iter_rows(max_row=1):
        for cell in row:
            cell.fill = PatternFill(patternType = 'solid', fgColor=title_color)
    for row in ws.iter_rows(min_row=max_row):
        for cell in row:
            cell.fill = PatternFill(patternType = 'solid', fgColor=title_color)

def main():
    
    print("start")
    
    # パスを設定
    dirname = os.path.dirname(__file__)
    import_foloder_path = dirname + '\試合結果'
    export_folder_path  = dirname + '\チーム成績'

    # 試合結果の全データを統合
    df_bat_concat   = pd.DataFrame()
    df_pitch_concat = pd.DataFrame()
    game_file_paths = get_xlsx_file_paths(folder_path=import_foloder_path)
    df_bat_concat, df_pitch_concat = concat_games(paths=game_file_paths, df_bat=df_bat_concat, df_pitch=df_pitch_concat)
    
    # カラムに試合or登板を追加
    df_bat_concat['試合'] = 1
    df_pitch_concat['登板'] = 1

    # 各項目を合計したデータフレームを作成
    df_bat_sum = df_bat_concat[['試合', '打席', '打数', '安打', '単打', '二塁打', '三塁打', '本塁打', '塁打',
                                '打点', '得点', '四球', '死球', '犠打', '犠飛', '打撃妨害', '失策', '野選',
                                '振り逃げ', '三振', '併殺', '盗塁企画', '盗塁']].groupby('名前').sum()

    df_pitch_sum = df_pitch_concat[['登板', '完封', '完投', '勝利', '敗戦', '引き分け', 'セーブ', '奪アウト数', '投球数', '打者数', '被安打',
                                    '与四球', '与死球', '奪三振', '失点', '自責点']].groupby('名前').sum()

    # 成績を生成する選手を抽出
    players_name = get_players_name(dirname)
    df_bat_sum   = df_bat_sum.filter(items=players_name, axis=0)
    df_pitch_sum = df_pitch_sum.filter(items=players_name, axis=0)

    # チーム総合をデータ化
    games = len(game_file_paths)
    df_bat_sum.loc['チーム総合']   = df_bat_sum.iloc[0:, 1:].sum()
    df_pitch_sum.loc['チーム総合'] = df_pitch_sum.iloc[0:, 1:].sum()
    df_bat_sum.loc['チーム総合', '試合']   = games
    df_pitch_sum.loc['チーム総合', '登板'] = games

    # 指標計算
    calc_bat_record(df_bat_sum)
    calc_pitch_record(df_pitch_sum)
    
    # ファイルに出力する
    date = str(datetime.date.today())
    with pd.ExcelWriter(export_folder_path + '/通算_' + date + '.xlsx') as writer:
        df_bat_sum.to_excel(writer, sheet_name='打撃成績')
        df_pitch_sum.to_excel(writer, sheet_name='投手成績')
        
    # 書式設定
    wb = px.load_workbook(export_folder_path+'/通算_'+date+'.xlsx')
    ws_bat = wb.worksheets[0]
    ws_pitch = wb.worksheets[1]
    
    # 列と行を固定化
    ws_bat.freeze_panes = 'B2'
    ws_pitch.freeze_panes = 'B2'
    
    # 試合数
    ws_bat['A1'] = str(games)+ '試合'
    ws_pitch['A1'] = str(games)+ '試合'
    
    # 率を小数第三位までに設定
    beginning_bat_rate = 25
    beginning_pitch_rate = 18
    set_rate_format(beginning=beginning_bat_rate, ws=ws_bat)
    set_rate_format(beginning=beginning_pitch_rate, ws=ws_pitch)
        
    # 一行目を縦書きに設定
    set_vertical_writing_row1(ws_bat)
    set_vertical_writing_row1(ws_pitch)       
        
    # 列の幅を設定
    set_column_width(ws=ws_bat, beginning_rate=beginning_bat_rate)
    set_column_width(ws=ws_pitch, beginning_rate=beginning_pitch_rate)
    
    # ['A1']は横書き，中央揃えに
    ws_bat['A1'].alignment = Alignment(horizontal = 'center', vertical='center')
    ws_pitch['A1'].alignment = Alignment(horizontal = 'center', vertical='center')
    
    #　セルの塗りつぶし　1列目→1行目→最終行目
    BAT_TITLE_CELL_COLOR = 'A4C6FF'
    BAT_NAME_CELL_COLOR = 'D9E5FF'

    PITCH_TITLE_CELL_COLOR = 'FFA3A3'
    PITCH_NAME_CELL_COLOR = 'FFD9D9'
    
    set_backgroud_color(ws=ws_bat, title_color=BAT_TITLE_CELL_COLOR, name_color= BAT_NAME_CELL_COLOR)
    set_backgroud_color(ws=ws_pitch, title_color=PITCH_TITLE_CELL_COLOR, name_color=PITCH_NAME_CELL_COLOR)
    
    #保存
    wb.save(export_folder_path+'/通算_'+date+'.xlsx')
    
    # 出力したフォルダを開く
    subprocess.Popen(["explorer", export_folder_path], shell=True)

    print("success!")

if __name__ == '__main__':
    main()
