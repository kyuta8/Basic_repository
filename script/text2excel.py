# ターミナルで実行する場合、
# python text2excel.py < ファイルパス >
# python text2excel.py < ディレクトリパス > < -a or -l >
# python text2excel.py -f
# python text2excel.py --help
# のいずれかで実行してください。
# -a オプションは、指定したディレクトリパス内に存在する全てのテキストファイルをExcelファイルに変換します。
# -l オプションは、指定したディレクトリパス内のファイルを表示します。
# -f オプションは、まずカレントディレクトリのフォルダとテキストファイルを表示します。
# 変換したいテキストファイルが違う階層にある場合、フォルダ名を入力し、一つずつ階層を移動します。
# 変換したテキストファイルが見つかれば、テキストファイル名を入力するか、何も入力せずにエンターキーを押してください。
# 何も入力しなかった場合、確認が行われます。
# yを入力した場合は、ディレクトリ内の全てのテキストファイルがテキストファイルに変換されます。
# nを入力した場合は、もう一度入力フェーズに移ります。誤ってエンターキーを押した場合などに使用してください。
# --help オプションは、text2excel.pyの簡単な説明が表示されます。


import os
import sys
import traceback

import re
import pprint
import pandas as pd


# 存在しないオプションを指定した場合のエラー
class NoSuchOption(Exception):
    pass

# 拡張子がテキストでない場合のエラー
class FileExtensionError(Exception):
    pass

# 指定ファイルが存在しない場合のエラー
class NoSuchFile(Exception):
    pass


# テキストファイルをExcelファイルに変換するための関数
def text2excel(text_file, file_path):
    print('Loading...\n   =>> {}'.format(text_file))
    

    # パスの存在確認
    if os.path.exists(text_file):
        # 存在すれば拡張子をテキストからExcelに変換
        excel_file = re.sub(r'\.txt', '.xlsx', text_file)
        ##pass

    else:
        # なければエラーを出力
        raise NoSuchFile('{} is not found.'.format(text_file))
    
    with open(text_file, mode='rt', encoding='utf-8-sig') as f:
        # テキストデータをリスト化
        read_data = list(f)
        # リストをデータフレームに変換：テキスト置換処理高速化のため
        text_df = pd.DataFrame(read_data)

    # 欠損値行、改行コードなどの除去
    text_df = text_df[text_df != '']
    text_df = text_df[text_df != ' ']    
    text_df = text_df.replace('\n', '', regex = True)
    text_df.reset_index(drop = True, inplace = True)

    # 文書の一文判定、連結のためにリスト化
    text_data = text_df[0].tolist()

    # 処理後の文の格納用リスト
    rev_text_data = []

    # ページ番号の除去のためのパターン
    page_num = r'^[0-9]*$|^[0-9]* $|^[０-９]*$|^[０-９]* $|^[i-ⅹ]*$|^[i-ⅹ]* $|^[0-9]* - [0-9]* -$|^- [0-9]* -$|^[０-９]* - [０-９]* -$|^- [０-９]* -$'
    re_page_num = re.compile(page_num)

    # ページ番号の要素を除去
    for i, r in enumerate(text_data):
        judge_page_num = re.match(re_page_num, r)
        
        # ページ番号以外の要素をリストに格納
        if judge_page_num:
            continue
        else:
            rev_text_data.append(r)       
    

    S = '' # 文連結のための保存用変数
    S_tmp = ''
    final_text_data = [] # 最終加工済みデータ格納用
    trace_order = []
    flag = 0 # 表データ除去のためのフラグ

    # 。で終わっている文の判定パターン1
    end_keyword1 = r'.*。.*'
    com_end_keyword1 = re.compile(end_keyword1)

    # 。で終わっている文の判定パターン2
    end_keyword2 = r'.*。 $'
    com_end_keyword2 = re.compile(end_keyword2)

    # (1)や(a)など文の始まりを表すキーワード
    start_keyword1 = r'^\([0-9]*\)|^\([a-z]\)|^\([０-９]*\)|^\([ア-ン]*\)|^\（[0-9]*\）|^\（[a-z]*\）|^\（[０-９]*\）|^\（[ア-ン]*\）'
    com_start_keyword1 = re.compile(start_keyword1)

    # ア.やア など文の始まりを表すキーワード
    start_keyword2 = r'^[ア-ヲ]\. |^[ア-ヲ] '
    com_start_keyword2 = re.compile(start_keyword2)

    # 1.や1.1など文の始まりを表すキーワード
    start_keyword3 = r'^([0-9]*\.)*[0-9]* |^([０-９]*\.)*[０-９]* '
    com_start_keyword3 = re.compile(start_keyword3)

    # ・など文の始まりを表すキーワード
    start_keyword4 = r'^・|^▪|^･'
    com_start_keyword4 = re.compile(start_keyword4)

    # 1や１など文の始まりを表すキーワード
    start_keyword5 = r'^[0-9] |^[０-９] '
    com_start_keyword5 = re.compile(start_keyword5)

    # ①など文の始まりを表すキーワード
    start_keyword6 = r'^[①-⑨]'
    com_start_keyword6 = re.compile(start_keyword6)

    # 注意事項を表すキーワード
    attract_keyword = r'^※'
    com_attract_keyword = re.compile(attract_keyword)

    # 。を含むが文の終わりでないパターン
    # 現在使っていないパターン
    continue_keyword = r'.*という。\).*|.*。\).*|.*。\）.*'
    com_continue_keyword = re.compile(continue_keyword)
    
    # 表でたまに出でくるキーワード
    sum_keyword = r'^合計.*'
    com_sum_keyword = re.compile(sum_keyword)

    # 表の開始位置検出のためのパターン
    table_keyword = r'.*表.*[0-9]*-[0-9]*.*|.*表.*[０-９]*-[０-９]*.*'
    com_table_keyword = re.compile(table_keyword)

    # 表題は除去
    tablename_keyword = r'^表.*[0-9]*-[0-9]*.*|^表.*[０-９]*-[０-９]*.*'
    com_tablename_keyword = re.compile(tablename_keyword)

    # 先頭文字が数字であっても文の始まりでないパターン
    linkage_keyword = r'^[0-9]*時間.*|[0-9]*月.*|^[0-9]*日.*|^[0-9]*万.*^[0-9]*年.*|^[0-9]*ヶ月.*|^[0-9]*ヵ月.*|^[0-9]*,[0-9]*.*'
    com_linkage_keyword = re.compile(linkage_keyword)
    
    # 図の題名
    graph_keyword = r'^図.*'
    com_graph_keyword = re.compile(graph_keyword)

    # 結合判定
    space_key = r'.* $'
    com_space_key = re.compile(space_key)



    for i, r in enumerate(rev_text_data):

        # 各パターン・キーワードとマッチング
        result_start_keyword1 = re.match(com_start_keyword1, r) # (1)や(a)など文の始まりを表すキーワード
        result_start_keyword2 = re.match(com_start_keyword2, r) # ア.やア など文の始まりを表すキーワード
        result_start_keyword3 = re.match(com_start_keyword3, r) # ・など文の始まりを表すキーワード
        result_start_keyword4 = re.match(com_start_keyword4, r) # ・など文の始まりを表すキーワード
        result_start_keyword5 = re.match(com_start_keyword5, r) # 1や１など文の始まりを表すキーワード
        result_start_keyword6 = re.match(com_start_keyword6, r) # ①など文の始まりを表すキーワード
        #result_continue_keyword = re.match(com_continue_keyword, r) # 。を含むが文の終わりでないパターン  # 現在使っていないパターン
        result_attract_keyword = re.match(com_attract_keyword, r) # 表でたまに出でくるキーワード
        #result_sum_keyword = re.match(com_sum_keyword, r) # 表でたまに出teくるキーワード
        #result_table_keyword = re.match(com_table_keyword, r) # 表の開始位置検出のためのパターン
        #result_tablename_keyword = re.match(com_tablename_keyword, r) # 表題は除去
        #result_linkage_keyword = re.match(com_linkage_keyword, r) # 先頭文字が数字であっても文の始まりでないパターン
        #result_graph_keyword = re.match(com_graph_keyword, r) # 図の題名
        result_space_key = re.match(com_space_key, r)

        if result_start_keyword1:
            R = re.sub(r'^\([0-9]*\)|^\([a-z]\)|^\([０-９]*\)|^\([ア-ン]*\)|^\（[0-9]*\）|^\（[a-z]*\）|^\（[０-９]*\）|^\（[ア-ン]*\）', '', r)

        elif result_start_keyword2:
            R = re.sub(r'^[ア-ヲ]\.|^[ア-ヲ] ', '', r)
            
        elif result_start_keyword3:
            R = re.sub(r'[0-9]*\.|[0-9]*-[0-9]*|-[0-9]*|[０-９]*\.|[０-９]*-[０-９]*|-[０-９]*', '', r)

        elif result_start_keyword4:
            R = re.sub(r'^・|^▪|^･', '', r)

        elif result_start_keyword5:         
            R = re.sub(r'[0-9]* |[０-９]* ', '', r)

        elif result_start_keyword6:
            R = re.sub(r'[①-⑨]', '', r)

        elif result_attract_keyword:  
            R = re.sub(r'^※', '', r)

        else:
            R = r



        if result_space_key:
            S += R
            '''
            result_end_keyword1 = re.search(com_end_keyword1, S) # 。を含む文の判定パターン1
            result_end_keyword2 = re.search(com_end_keyword2, S) # 。を含む文の判定パターン2
            
            if result_end_keyword1:
                loc = r.count('。')
                if loc == 1:
                    if result_end_keyword2:
                        final_text_data.append(S)
                        S = ''
                    else:
                        F = S.split('。')
                        final_text_data.append(F[0])
                        final_text_data.append(F[1])
                        S = ''
                elif loc > 1:
                    F = S.split('。')
                    if result_end_keyword2:
                        for f_i in range(len(F)):
                            _f = F[f_i]
                            if f_i == 0:
                                S_tmp = _f
                                
                            elif f_i == len(F) - 1:
                                final_text_data.append(S_tmp)
                                S = ''
                                break
                            if _f[0] == ')':
                                S_tmp += _f
                            
                            else:
                                final_text_data.append(S_tmp)
                                S_tmp = _f
                    else:
                        for f_i in range(len(F)):
                            _f = F[f_i]
                            if f_i == 0:
                                S_tmp = _f
                            if _f[0] == ')':
                                S_tmp += _f
                            
                            else:
                                final_text_data.append(S_tmp)
                                S_tmp = _f
                        final_text_data.append(S_tmp)
                        S = ''
            else:
                final_text_data.append(S)
                S = ''
            '''
            final_text_data.append(S)
            S = ''

        else:
            S += R



    if not os.path.exists(file_path + '/excel_dataset'):
        os.mkdir(file_path + '/excel_dataset')
        print('Create directory \'excel_dataset\': path -> {}'.format(file_path + '/excel_dataset'))

    '''
    # 処理のデバッグデータ
    if not os.path.exists(file_path + '/trace'):
        os.mkdir(file_path + '/trace')
        print('Create directory \'trace\': path -> {}'.format(file_path + '/trace'))
    TRACE = pd.DataFrame(trace_order, columns = ['next_sentence'])
    TRACE.to_excel(file_path + '/trace/trace_' + os.path.basename(excel_file), index = None, encoding = 'utf_8_sig')
    '''

    NFR = pd.DataFrame(final_text_data, columns = ['text'])
    NFR = NFR.replace(' ', '', regex = True)
    NFR = NFR[NFR['text'] != '']

    # 仕様書件名のインデックス取得
    index_NFR = NFR[NFR['text'] == '調達件名'].index
    # 件名が記述されている場合とされていない場合の処理分岐
    if len(index_NFR) == 1: # 件名が記述されている場合
        index_NFR = index_NFR[0] + 1
    
    elif len(index_NFR) == 0: # 件名が記述されていない場合
        index_NFR = 0
    
    else: # 件名が複数記述されている場合
        index_NFR = index_NFR[1] + 1

    title_name = NFR['text'][index_NFR] + '.xlsx'

    #NFR.to_excel(file_path + '/excel_dataset/' + os.path.basename(excel_file), index = None, encoding = 'utf_8_sig')
    #print('Save file converted to excel \'{}\':'.format(os.path.basename(excel_file)))
    #print('   PATH =>> {}'.format(file_path + '/excel_dataset' + os.path.basename(excel_file)))

    print('Rename: {} =>> {}'.format(excel_file, title_name))

    NFR.to_excel(file_path + '/excel_dataset/' + title_name, index = None, encoding = 'utf_8_sig')
    print('Save file converted to excel \'{}\':'.format(title_name))
    print('   PATH =>> {}'.format(file_path + '/excel_dataset/' + title_name))




# 実行ファイルのオプション
_options = ['-a', '-l']

def main():
    try:

        # オプションの説明
        if '--help' in sys.argv:
            print('\nUsage: python pdf2excel.py < text file > < -a or -l > or python pdf2excel.py < -f or --help >')
            print()
            print('Some useful options are:')
            print('   -a  : All text files that exist in selected directory convert to excel files')
            print('   -l  : Show all text files that exist in selected directory')
            print('   -f  : Show all text files that exist in current directory and input directory or file path of file that you want to convert to')
            print('\nIf you want to read usage, you select \'--help\' option, please')
            print()

        elif '-f' in sys.argv:
            # カレントディレクトリの表示
            print('\nCurrent directory: {}'.format(os.getcwd()))
            
            # カレントディレクトリに含まれるフォルダとファイルの表示
            print('Directories or Files: \n')
            # カレントディレクトリのディレクトリ構造を取得
            files = os.listdir(os.getcwd())
            # テキストファイルとフォルダの取得
            k_words = re.compile(r'^.*\.txt$')
            ext_files = [f for f in files if re.match(k_words, f) or os.path.isdir(os.path.join(os.getcwd(), f))]
            ext_files = sorted(ext_files)
            pprint.pprint(ext_files)
            
            # フォルダやファイルの入力
            print('\nPlease input path:')
            file_path = input()

            # パス結合
            file_path = os.path.join(os.getcwd(), file_path) 

            # 変換したいファイルやファイル群が決定するまで入力を行う
            while True:
                print('\nCurrent directory: {}'.format(file_path))
                print('Directories or Files: \n')
                files = os.listdir(file_path)
                k_words1 = re.compile(r'^\.')
                ext_files = [f for f in files if re.match(k_words, f)]
                [ext_files.append(f) for f in files if not re.match(k_words1, f) and os.path.isdir(os.path.join(file_path, f))]
                ext_files = sorted(ext_files)
                pprint.pprint(ext_files)
                print('\nPlease input path:')
                path = input()

                confirm = ''

                # 何も入力されなかった場合
                if len(path) == 0:

                    # ファイル群一式変換したい場合はy、再入力を行う場合はn
                    while True:
                        print('Finish to input path (y/n):')
                        confirm = input()

                        if 'y' == confirm or 'n' == confirm:
                            break

                        else:
                            print('Please input y or n.')
                    
                    if 'y' == confirm:

                        if os.path.exists(file_path):
                            files = os.listdir(file_path)
                            k_words = re.compile(r'^.*\.txt$')
                            file_paths = [file_path + '/' + f for f in files if re.match(k_words, f)]
                        break
                    
                    elif 'n' == confirm:
                        pass

                elif '..' == path:

                    # ..を入力された場合は一つ階層を戻す
                    # ただしUsersの階層には戻れないようにしておく
                    if os.path.normpath(os.path.join(file_path, path)) == '/Users':
                        pass
                    
                    else:
                        file_path = os.path.join(file_path, path)
                        file_path = os.path.normpath(file_path)
                
                else:
                    
                    # ファイル名が入力された場合は、存在するファイルか確認し
                    # 存在すれば入力を終了し、なければ再入力を行う
                    if os.path.exists(os.path.join(file_path, path)):
                        file_path = os.path.join(file_path, path)

                        if re.match(r'^.*\.txt$', file_path):
                            file_paths = [file_path]
                            file_path = re.sub('/' + os.path.basename(file_path), '', file_path)
                            break
                    
                    else:
                        print('Incorrect path!!!')

        else:
            file_path = sys.argv[1]

            if len(sys.argv) > 2:
                option = sys.argv[2]

                if option in _options:

                    if option == _options[0]: # -a オプション
                        files = os.listdir(file_path)
                        k_words = re.compile(r'^.*\.txt$')
                        file_paths = [file_path + '/' + f for f in files if re.match(k_words, f)]
                        pprint.pprint(file_paths)

                    elif option == _options[1]: # -l オプション
                        print()
                        files = os.listdir(file_path)
                        k_words = re.compile(r'^.*\.txt$')
                        [print(f) for f in files if re.match(k_words, f)]
                        print('\nSelect file:')
                        file_paths = [file_path + '/' + input()]

                else:
                    # 存在しないオプションを入力された場合にエラー出力
                    raise NoSuchOption('Selected option is incorrect. Please select --help option and read usage.')

            else:

                if re.match(r'^.*\.txt$', file_path):
                    file_paths = [file_path]
                    file_path = re.sub('/' + os.path.basename(file_path), '', file_path)

                else:
                    # 拡張子が違う場合にエラーを出力
                    raise FileExtensionError('File extension must be text.')

        if not '--help' in sys.argv:
            for text_file in file_paths:
                text2excel(text_file, file_path)

    except:
        traceback.print_exc()



if __name__ == '__main__':
    main()

