import os
import pandas as pd
import mymodule
import subprocess

def concat_games(paths, df_bat, df_pitch):
    for path in paths:
        _game = path.split('\\')[-1]
        game  = _game.split('_')
        date  = game[0]
        style = game[1]
        team  = game[2].split('.')[0]
        # 打撃
        df_bat_read_excel=pd.read_excel(path, sheet_name='打撃成績')
        df_bat_read_excel.insert(0,'日付',date)
        df_bat_read_excel.insert(1,'試合形式',style)
        df_bat_read_excel.insert(2,'チーム',team)
        df_bat=pd.concat([df_bat, df_bat_read_excel])
        # 投手
        df_pitch_read_excel=pd.read_excel(path, sheet_name='投手成績')
        df_pitch_read_excel.insert(0,'日付',date)
        df_pitch_read_excel.insert(1,'試合形式',style)
        df_pitch_read_excel.insert(2,'チーム',team)
        df_pitch=pd.concat([df_pitch, df_pitch_read_excel])
    return df_bat, df_pitch

def main():
    
    print("start")
    
    # パスを設定
    dirname = os.path.dirname(__file__)
    import_foloder_path = dirname + '\試合結果'
    export_folder_path  = dirname + '\個人成績'
    
     # 試合結果の全データを統合
    df_bat_concat   = pd.DataFrame()
    df_pitch_concat = pd.DataFrame()
    game_file_paths = mymodule.get_xlsx_file_paths(folder_path=import_foloder_path)
    df_bat_concat, df_pitch_concat = concat_games(game_file_paths, df_bat_concat, df_pitch_concat)
    
    # 成績を生成する選手を抽出
    players_name = mymodule.get_players_name(dirname)
    df_bat_concat   = df_bat_concat[df_bat_concat['名前'].isin(players_name)]
    df_pitch_concat = df_pitch_concat[df_pitch_concat['名前'].isin(players_name)]
    
    #個人の名前を抽出
    bat_players_name   = df_bat_concat['名前'].unique()
    pitch_players_name = df_pitch_concat['名前'].unique()
    
    #個人ごとにファイルを作成
    for bat_player_name in bat_players_name:
        df_bat = pd.DataFrame()
        df_bat = df_bat_concat[df_bat_concat['名前']==bat_player_name]
        # 必要なカラムのみ抽出
        df_bat = df_bat[['日付','試合形式','チーム', '打席', '打数', '安打', '単打',
                                       '二塁打', '三塁打', '本塁打', '塁打','打点', '得点', '四球',
                                       '死球', '犠打', '犠飛', '打撃妨害', '失策', '野選', '振り逃げ',
                                       '三振', '併殺', '盗塁企画', '盗塁']]
        
        # 日付，試合形式でソート
        df_bat["日付"] = pd.to_datetime(df_bat['日付'])
        df_bat = df_bat.sort_values(['日付','試合形式'])
        df_bat["日付"]=df_bat['日付'].dt.strftime('%Y-%m-%d')
        # 試合数を保持
        bat_games = len(df_bat)
        # 合計行を作成
        df_bat.loc['合計'] = float('nan')
        df_bat.loc['合計', '打席':'盗塁'] = df_bat.iloc[:bat_games,3:].sum()
        # indexを設定
        df_bat = df_bat.set_index(['日付','試合形式','チーム'])
        
        mymodule.calc_bat_record(df_bat)
        
        is_pitch = False
        
        if bat_player_name in pitch_players_name:
            is_pitch = True
            df_pitch = df_pitch_concat[df_pitch_concat['名前']==bat_player_name]
            # 日付，試合形式でソート
            df_pitch["日付"] = pd.to_datetime(df_pitch['日付'])
            df_pitch = df_pitch.sort_values(['日付','試合形式'])
            df_pitch["日付"]=df_pitch['日付'].dt.strftime('%Y-%m-%d')
            # 試合数を保持
            pitch_games = len(df_pitch)
            # 合計行を作成
            df_pitch.loc['合計'] = float('nan')
            df_pitch.loc['合計', '完封':'自責点'] = df_pitch.iloc[0:pitch_games,3:].sum()
            # indexを設定
            df_pitch = df_pitch.set_index(['日付','試合形式','チーム'])
            df_pitch.drop('名前', axis=1, inplace=True)

            mymodule.calc_pitch_record(df_pitch)
            
        with pd.ExcelWriter(export_folder_path+'/'+bat_player_name+'.xlsx') as writer:
            df_bat.to_excel(writer, sheet_name='打撃成績')
            if is_pitch:
                df_pitch.to_excel(writer, sheet_name='投手成績')
                
    # 出力したフォルダを開く
    subprocess.Popen(["explorer", export_folder_path], shell=True)
    print("success!")
    
if __name__ == '__main__':
    main()