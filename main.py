import Message
import xlsContorller
import toukiController
from datetime import datetime

output_file_list = []

# 設定ファイル読込
import getConfig


def main():
    # 設定ファイル読込
    config = None
    setting = None
    try:
        config = getConfig.getConditionFromXls()
        setting = getConfig.getSettingFromXls()
    except Exception as e:
        errorMessage=f'設定ファイルの読み込みでエラーが発生しました。\nエラー内容[{e}]'
        print(errorMessage)
        Message.MessageForefrontShowinfo(errorMessage)
        return
    
    user_id = config['user_id']
    password = config['password']
    conditions_list = config['conditions_list']
    
    # 収集条件チェック(収集条件) 
    for conditions in conditions_list:
        if toukiController.checkConditions(conditions) == False:
            # エラーメッセージを表示し、処理終了
            errorMessage = '収集条件に誤りがあります。\nログを参照し、訂正後、再実行してください。'
            print(errorMessage)
            Message.MessageForefrontShowinfo(errorMessage)
            return

    # 収集開始メッセージ
    startMessage = f'不動産請求情報収集を実行しますか？\n実行ユーザはID番号【{user_id}】、パスワード【{password}】です。\n収集条件は以下の通りです。'
    for i, conditions in enumerate(conditions_list):
        startMessage += f'\n\n{i+1}：{xlsContorller.editCollectionCondition(conditions)}'
        if conditions[6] == True:
            startMessage += f' ※注意※ファイル出力なし、地番・家屋番号の選択状態で処理を一時停止します。'
    print(startMessage)
    if Message.MessageForefront(startMessage) == False:
        return
    
    # データ収集
    import ProcessStatus
    ps = ProcessStatus.ProcessStatus(setting)
    p_count = len(conditions_list)
    output_file_path = ""
    for i, conditions in enumerate(conditions_list):
        status = f"{i+1}/{p_count}"
        ps.showStatus(status)
        output_file_path = toukiController.collectData(conditions, user_id, password, setting=setting)
        # 処理終了時のメッセージ表示のため、出力ファイル名を追記 ToDo:選択のみのとき、ファイル名の返却なしを考慮
        if output_file_path != "":
            output_file_list.append(output_file_path)
        else:
            output_file_list.append("ファイル出力なし、地番・家屋番号の選択状態で一時停止しました")

    ps.close()

    
    # 収集終了メッセージ
    endMessage = '不動産請求情報収集が処理終了しました。\n収集結果は以下に出力されています。ご確認ください。'
    for i, output_file in enumerate(output_file_list):
        endMessage += f'\n\n{i+1}：【{output_file}】'
    print(endMessage)
    Message.MessageForefrontShowinfo(endMessage)


import os
import sys
# カレントディレクトリ変更
os.chdir(os.path.dirname(os.path.abspath(sys.argv[0])))

main()

# 終了時に自動的にコンソールを消さない
endMessage = '''
≪≪≪≪≪ コンソールを消すためには、「Enter」キーを押してください ≫≫≫≫≫
'''
input(endMessage)
