import discord
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import gspread
from gspread_dataframe import get_as_dataframe, set_with_dataframe

from oauth2client.service_account import ServiceAccountCredentials 

#2つのAPIを記述しないとリフレッシュトークンを3600秒毎に発行し続けなければならない
scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']

#認証情報設定
#ダウンロードしたjsonファイル名をクレデンシャル変数に設定（秘密鍵、Pythonファイルから読み込みしやすい位置に置く）
credentials = ServiceAccountCredentials.from_json_keyfile_name('/Users/daisukeide/Desktop/file/other/private_document/mk8dx/op-course-result-1517e476bebb.json', scope)

#OAuth2の資格情報を使用してGoogle APIにログインします。
gc = gspread.authorize(credentials)

#共有設定したスプレッドシートキーを変数[SPREADSHEET_KEY]に格納する。
## コース情報
SPREADSHEET_KEY_1 = '1KY3x9lhCjxPjO5fLDyJ1uloi4NVylvB4KGCZDQJr8Ws'
## 成績情報
SPREADSHEET_KEY_2 = '1-NP-9Ghuc9x_t9GCM-ITcmcGlNOvJI3sBog1ENqUFaY'

#共有設定したスプレッドシートのシート1を開く
worksheet = gc.open_by_key(SPREADSHEET_KEY_1).sheet1
worksheet2 = gc.open_by_key(SPREADSHEET_KEY_2).sheet1


# 自分のBotのアクセストークンに置き換えてください
TOKEN = 'NzcwNTkzMjk3MzU2OTQ3NDU2.X5f07w.o9iErwVjRAabmrULjzElFCbmlrc'

# 接続に必要なオブジェクトを生成
client = discord.Client()

# 起動時に動作する処理
@client.event
async def on_ready():
    # 起動したらターミナルにログイン通知が表示される
    print('ログインしました')

# メッセージ受信時に動作する処理
@client.event
async def on_message(message):
    
    # ローカルデータにアクセス
    course = pd.DataFrame(worksheet.get_all_values())
    # 整形
    course.columns = list(course.loc[0, :])
    course.drop(0, inplace=True)
    course.reset_index(inplace=True)
    course.drop('index', axis=1, inplace=True)

    course_list = list(course.iloc[0])
    course_type_list = list(course.iloc[1])
    course_label = list(course.columns)
    course_list_lower = []

    # コース名を小文字に変換
    for i in range(len(course_list)):
        course_list_lower.append(str.lower(course_list[i]))

    # メッセージ送信者がBotだった場合は無視する
    if message.author.bot:
        return

    # _aggコマンドでコースの集計
    elif message.content[:4] == '_agg':

        # コースの略記名で集計するコースを絞る
        course_name_lower = str.lower(message.content[5:])
        if course_name_lower in course_list_lower:
            # コース名を日本語に変換
            for k in range(len(course_label)):
                if course_name_lower == course_list_lower[k]:
                    cl = course_label[k]
                    cn = course_list[k]
            await message.channel.send('【コース別集計】 ' + cl)
            
            # データ読み込み
            course_result = get_as_dataframe(worksheet2, usecols=[0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 12, 13, 14])

            # 整形
            course_result = course_result.dropna(how='all')
            course_result = course_result.fillna(0)

            list_date = list(course_result['date'])
            count_battle = len(set(list_date))
            course_result_2 = course_result.loc[course_result['course'] == cn]

            if len(course_result_2) > 0:
                await message.channel.send('`データ数: ' + str(len(course_result_2)) + ' 出走確率: ' + str(round((len(course_result_2) / count_battle) * 100, 2)) + '%`')
                rank_x = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12]
                rank_y = []
                result_label = list(course_result_2.columns)
                result_label.remove('course')
                result_label.remove('date')
                result_label.remove('enemy')
                course_result_2 = course_result_2[result_label]
                for i in range(len(result_label)):
                    rank_y.append(course_result_2[result_label[i]].sum())

                point = [15, 12, 10, 9, 8, 7, 6, 5, 4, 3, 2, 1]
                print(rank_x)
                print(rank_y)
                rank_avg = round(np.dot(rank_x, rank_y) / (len(course_result_2) * 6), 2)
                await message.channel.send('`レースごとの順位平均: ' + str(rank_avg) + '`')
                point_avg = round(np.dot(point, rank_y) / len(course_result_2), 2)
                await message.channel.send('`レースごとの平均獲得pts: ' + str(point_avg) + '(' + str(round(2 * point_avg - 82.0, 2)) +')`')
                
                if len(course_result_2) > 1:
                    list_for_str = []
                    for j in range(len(course_result_2)):
                        list_for_str.append(np.dot(point, course_result_2.iloc[j]))
                    point_str = np.round(np.std(list_for_str), 2)
                    await message.channel.send('`獲得pts標準偏差: ' + str(point_str) + '`')
                else:
                    await message.channel.send('`獲得pts標準偏差: 算出には最低2レース分のデータが必要です。`')
                        
                plt.figure(figsize=(7, 5))
                plt.title(cl + ' ' + cn)
                plt.xlabel('順位')
                plt.ylabel('回数')
                plt.bar(rank_x, rank_y)
                save_path = '/Users/daisukeide/Desktop/file/other/private_document/mk8dx/data/op/' + cn + '.png'
                plt.savefig(save_path)

                pic_hoge = discord.File(save_path)
                await message.channel.send(file=pic_hoge)
                
                del course_result
                del course_result_2

            else:
                await message.channel.send('データがありません')
        
        # allで全コースのデータ集計
        elif message.content[5:] == 'all':
            await message.channel.send('【全コース集計】')
            # データ読み込み
            course_result = get_as_dataframe(worksheet2, usecols=[0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 12, 13, 14])

            # 整形
            course_result = course_result.dropna(how='all')
            course_result = course_result.fillna(0)
            list_date = list(course_result['date'])
            count_battle = len(set(list_date))

            if len(course_result) > 0:
                await message.channel.send('`総データ数: ' + str(len(course_result)) + ' 交流戦回数: ' + str(count_battle) + '回`')
                rank_x = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12]
                rank_y = []
                result_label = list(course_result.columns)
                result_label.remove('course')
                result_label.remove('date')
                result_label.remove('enemy')
                course_result = course_result[result_label]
                for i in range(len(result_label)):
                    rank_y.append(course_result[result_label[i]].sum())
                point = [15, 12, 10, 9, 8, 7, 6, 5, 4, 3, 2, 1]
                rank_avg = round(np.dot(rank_x, rank_y) / (len(course_result) * 6), 2)
                await message.channel.send('`レースごとの順位平均: ' + str(rank_avg) + '`')
                point_avg = round(np.dot(point, rank_y) / len(course_result), 2)
                await message.channel.send('`レースごとの平均獲得pts: ' + str(point_avg) + '(' + str(round(2 * point_avg - 82.0, 2)) +')`')
                
                if len(course_result) > 1:
                    list_for_str = []
                    for j in range(len(course_result)):
                        list_for_str.append(np.dot(point, course_result.iloc[j]))
                    point_str = np.round(np.std(list_for_str), 2)
                    await message.channel.send('`獲得pts標準偏差: ' + str(point_str) + '`')
                
                else:
                    await message.channel.send('`獲得pts標準偏差:  算出には最低2レース分のデータが必要です。`')

                plt.figure(figsize=(7, 5))
                plt.title('全コース集計 Aggregation of all courses')
                plt.xlabel('順位')
                plt.ylabel('回数')
                plt.bar(rank_x, rank_y)
                save_path = '/Users/daisukeide/Desktop/file/other/private_document/mk8dx/data/op/all_courses.png'
                plt.savefig(save_path)

                pic_hoge = discord.File(save_path)
                await message.channel.send(file=pic_hoge)

                del course_result

            else:
                await message.channel.send('データがありません')
    
    # 利用方法
    elif message.content == '_how2':
        await message.channel.send('【コマンド一覧 ver1.0】')
        await message.channel.send('`_agg xxx` : コース別集計。xxxの箇所には以下のコース名を入力してください。xxxの箇所にallと入力すると全コースの集計が見れます。')
        await message.channel.send('`_count x` : 出走回数のランキング表示。xの箇所には半角数字をを入力すると上位x位までのランキングを表示します。(省略可能。その場合は1回以上走った全てのコースを表示します。)')
        await message.channel.send('`_team xxx` : コース別集計。xxxの箇所にはチーム名を入力してください(大文字小文字も正確に)。_teamのみ入力すればデータとして残っているチームを表示してくれます。')
        await message.channel.send('`_rank x` : 出走回数のランキング表示。xの箇所にはf, m, bいずれかを入力するとそれぞれ前コ、中位コ、打開コ別に表示します。(省略可能。その場合は全てのコースを表示します。)')
        list_path = '/Users/daisukeide/Desktop/file/other/private_document/mk8dx/data/list.jpg'
        pic_hoge = discord.File(list_path)
        await message.channel.send(file=pic_hoge)

    # _teamコマンドでチームごとの戦績集計
    elif message.content[:5] == '_team':
        enemy_name = message.content[6:]
        enemy_name_lower = str.lower(enemy_name)
        # データ読み込み
        course_result = get_as_dataframe(worksheet2, usecols=[0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 12, 13, 14])

        # 整形
        course_result = course_result.dropna(how='all')
        course_result = course_result.fillna(0)
        enemy_list = course_result['enemy']
        enemy_list = list(set(enemy_list))
        enemy_lower_list = []
        for i in range(len(enemy_list)):
            enemy_lower_list.append(str.lower(enemy_list[i]))
        if enemy_name_lower in enemy_lower_list:
            # チーム名逆引き
            enemy_index = 0
            for i in range(len(enemy_lower_list)):
                if enemy_name_lower == enemy_lower_list[i]:
                    enemy_index = i
                    break
            enemy_name = enemy_list[enemy_index]
            course_result2 = course_result.loc[course_result['enemy'] == enemy_name]
            course_result_label = list(course_result.columns)
            course_result_label.remove('course')
            course_result_label.remove('date')
            course_result_label.remove('enemy')
            course_result2 = course_result2.fillna(0)
            # データ上の戦績集計
            date_list = list(set(list(course_result2['date'])))
            points = [15, 12, 10, 9, 8, 7, 6, 5, 4, 3, 2, 1]
            pts_list = []
            btpts_list = []
            wld = [0, 0, 0]

            # ヒストグラム用
            rank_x = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12]
            rank_y = []
            cr2 = course_result2[course_result_label]
            for i in range(len(course_result_label)):
                rank_y.append(cr2[course_result_label[i]].sum())

            for i in date_list:
                cr3 = course_result2.loc[course_result2['date'] == i]
                cr3 = cr3[course_result_label]
                cr3 = cr3.fillna(0)
                point = 0
                for j in range(len(cr3)):
                    pt = np.dot(points, list(cr3.iloc[j]))
                    pts_list.append(pt)
                    point = point + pt
                btpts_list.append(point)
                if point >= 493:
                    wld[0] = wld[0] + 1
                elif point <= 491:
                    wld[1] = wld[1] + 1
                else:
                    wld[2] = wld[2] + 1
            
            await message.channel.send('【チーム集計】 ' + enemy_name)
            if wld[0] + wld[1] > 0:
                await message.channel.send('`戦績: ' + str(wld[0]) + '勝' + str(wld[1]) + '負' + str(wld[2]) + '分  勝率' + str(np.round((wld[0] / (wld[0] + wld[1])) * 100, 2)) + '%`')
            else:
                await message.channel.send('`戦績: ' + str(wld[0]) + '勝' + str(wld[1]) + '負' + str(wld[2]) + '分`')
            btpts_avg = np.round(sum(btpts_list) / len(btpts_list), 1)
            await message.channel.send('`試合ごとの平均獲得pts: ' + str(btpts_avg) + '(' + str(round(2 * btpts_avg - 984, 1)) +')`')
            if len(btpts_list) > 1:
                btpts_str = np.round(np.std(btpts_list), 2)
                await message.channel.send('`試合ごとの獲得pts標準偏差: ' + str(btpts_str) + '`')
            
            else:
                await message.channel.send('`試合ごとの獲得pts標準偏差:  算出には最低2試合分のデータが必要です。`')
            pts_avg = np.round(sum(pts_list) / len(pts_list), 2)
            await message.channel.send('`レースごとの平均獲得pts: ' + str(pts_avg) + '(' + str(round(2 * pts_avg - 82.0, 2)) +')`')
            if len(pts_list) > 2:
                pts_str = np.round(np.std(pts_list), 2)
                await message.channel.send('`レースごとの獲得pts標準偏差: ' + str(pts_str) + '`')
            
            else:
                await message.channel.send('`獲得pts標準偏差:  算出には最低2レース分のデータが必要です。`')
            
            plt.figure(figsize=(7, 5))
            plt.title('vs ' + enemy_name)
            plt.xlabel('順位')
            plt.ylabel('回数')
            plt.bar(rank_x, rank_y)
            save_path = '/Users/daisukeide/Desktop/file/other/private_document/mk8dx/data/op/' + enemy_name + '_courses.png'
            plt.savefig(save_path)

            pic_hoge = discord.File(save_path)
            await message.channel.send(file=pic_hoge)

            del course_result2
            del cr2
            del cr3
            del course_result

        else:
            await message.channel.send('データがありません。現在登録されているチームの情報は以下の通りです。')
            await message.channel.send('`' + str(enemy_list) + '`')


    # _countコマンドでコースの走った回数をランキング形式で表示
    elif message.content[:6] == '_count':
        try:
            iteration_number = 48
            if message.content[7:] != '':
                iteration_number = int(message.content[7:])
            await message.channel.send('【走った回数が多いコースランキング】')
            # データ読み込み
            course_result = get_as_dataframe(worksheet2, usecols=[0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 12, 13, 14])

            # 整形
            course_result = course_result.dropna(how='all')
            course_result = course_result.fillna(0)

            list_date = list(course_result['date'])
            count_battle = len(set(list_date))
            await message.channel.send('交流戦回数(集計データ分): ' + str(count_battle))
            course_count_list = []
            for i in range(len(course_list)):
                course_result_2 = course_result.loc[course_result['course'] == course_list[i]]
                course_count_list.append(len(course_result_2))
            count_data = pd.DataFrame()
            count_data['course'] = course_label
            count_data['count'] = course_count_list
            count_data = count_data.sort_values('count', ascending=False)

            rank_list = []
            r = 0
            for i in range(len(count_data)):
                if i == 0:
                    rank_list.append(r)
                else:
                    if count_data['count'].iloc[i-1] == count_data['count'].iloc[i]:
                        rank_list.append(r)
                    else:
                        r = i
                        rank_list.append(r)
            count_data['rank'] = rank_list
            count_data = count_data.loc[count_data['count'] >= 1]
            for i in range(len(count_data)):
                if int(count_data['rank'].iloc[i] + 1) > iteration_number:
                    break
                await message.channel.send('`第' + str(count_data['rank'].iloc[i] + 1) + '位: ' + count_data['course'].iloc[i] + ' ' + str(count_data['count'].iloc[i]) + '回 ' + str(round((count_data['count'].iloc[i] / count_battle) * 100, 2)) + '%`')
                
            del course_result
            del count_data
        except ValueError:
            await message.channel.send('`_count (半角数字)`の形式で入力してください(半角数字は省略可)。入力した数字の分だけ上から集計結果を表示できます。')
            
    # _rankコマンドでコースの走った回数をランキング形式で表示
    elif message.content[:5] == '_rank':
        course_type = message.content[6:]

        # 集計対象コースの選定
        if course_type in ['f', 'm', 'b', '']:
            target_course = []
            target_course_label = []
            if course_type == 'f':
                await message.channel.send('【平均獲得ptsランキング 前コース編】')
                for i in range(len(course_list)):
                    if course_type_list[i] == 'f':
                        target_course.append(course_list[i])
                        target_course_label.append(course_label[i])
            elif course_type == 'm':
                await message.channel.send('【平均獲得ptsランキング 中位コース編】')
                for i in range(len(course_list)):
                    if course_type_list[i] == 'm':
                        target_course.append(course_list[i])
                        target_course_label.append(course_label[i])
            elif course_type == 'b':
                await message.channel.send('【平均獲得ptsランキング 打開コース編】')
                for i in range(len(course_list)):
                    if course_type_list[i] == 'b':
                        target_course.append(course_list[i])
                        target_course_label.append(course_label[i])
            else:
                await message.channel.send('【平均獲得ptsランキング】')
                target_course = course_list
                target_course_label = course_label

            # ランキング用テーブル作成
            avg_list = []
            size_list = []
            points = [15, 12, 10, 9, 8, 7, 6, 5, 4, 3, 2, 1]
            # データ読み込み
            course_result = get_as_dataframe(worksheet2, usecols=[0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 12, 13, 14])

            # 整形
            course_result = course_result.dropna(how='all')
            course_result = course_result.fillna(0)

            for i in range(len(target_course)):
                course_result_2 = course_result.loc[course_result['course'] == target_course[i]]
                size = len(course_result_2)
                size_list.append(size)
                if size > 0:
                    rank_y = []
                    result_label = list(course_result.columns)
                    result_label.remove('course')
                    result_label.remove('date')
                    result_label.remove('enemy')
                    course_result_2 = course_result_2[result_label]
                    for j in range(len(result_label)):
                        rank_y.append(course_result_2[result_label[j]].sum())
                    rank_avg = round(np.dot(points, rank_y) / (len(course_result_2)), 2)
                    avg_list.append(rank_avg)
                else:
                    avg_list.append(0)
            rank_data = pd.DataFrame()
            rank_data['course'] = target_course_label
            rank_data['count'] = size_list
            rank_data['avg'] = avg_list
            rank_data = rank_data.sort_values('avg', ascending=False)

            ranking_list = []
            r = 0
            for i in range(len(rank_data)):
                if i == 0:
                    ranking_list.append(r)
                else:
                    if rank_data['avg'].iloc[i-1] == rank_data['avg'].iloc[i]:
                        ranking_list.append(r)
                    else:
                        r = i
                        ranking_list.append(r)
            rank_data['rank'] = ranking_list

            for i in range(len(rank_data)):
                if rank_data['count'].iloc[i] > 0:
                    await message.channel.send('`第' + str(rank_data['rank'].iloc[i] + 1) + '位: ' + rank_data['course'].iloc[i] + ' ' + str(rank_data['avg'].iloc[i]) + '(' + str(np.round(2 * rank_data['avg'].iloc[i] - 82.0, 2)) + ') 出走回数: ' + str(rank_data['count'].iloc[i]) + '回`')
                else:
                    await message.channel.send('`第' + str(rank_data['rank'].iloc[i] + 1) + '位: ' + rank_data['course'].iloc[i] + ' ' + str(rank_data['avg'].iloc[i]) + ' ※未出走`')

            del course_result
            del rank_data

        else:
            await message.channel.send('コマンドが正しくありません。以下のコマンドを入力してください。')
            await message.channel.send('`_rank` :  全コースの平均獲得ptsランキング')
            await message.channel.send('`_rank f` : 前コース')
            await message.channel.send('`_rank m` : 中位コース')
            await message.channel.send('`_rank b` :  打開コース、')

    elif message.content[:5] == '_burn':
        await message.channel.send('そんな簡単に燃やされてたまるかよ')
    
    
# Botの起動とDiscordサーバーへの接続
client.run(TOKEN)