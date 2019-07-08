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

    # パスの存在確認
    if os.path.exists(text_file):
        # 存在すれば拡張子をテキストからExcelに変換
        excel_file = re.sub(r'\.txt', '.xlsx', text_file)

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
    page_num = r'^[0-9]*$|^[0-9]* - [0-9]* -$|^- [0-9]* -$|^[０-９]*$|^[０-９]* - [０-９]* -$|^- [０-９]* -$'
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
    final_text_data = [] # 最終加工済みデータ格納用
    flag = 0 # 表データ除去のためのフラグ

    # 。で終わっている文の判定パターン1
    end_keyword1 = r'.*。.*'
    com_end_keyword1 = re.compile(end_keyword1)

    # 。で終わっている文の判定パターン2
    end_keyword2 = r'.*。$'
    com_end_keyword2 = re.compile(end_keyword2)

    # (1)や(a)など文の始まりを表すキーワード
    start_keyword1 = r'^\([0-9]*\).*|^\([a-z]\).*|^\([０-９]*\).*|^\([ア-ン]*\).*|^\（[0-9]*\）.*|^\（[a-z]*\）.*|^\（[０-９]*\）.*|^\（[ア-ン]*\）.*'
    com_start_keyword1 = re.compile(start_keyword1)

    # ア.やア など文の始まりを表すキーワード
    start_keyword2 = r'^[ア-ヲ]\. .*|^[ア-ヲ] .*'
    com_start_keyword2 = re.compile(start_keyword2)

    # 1.や１.など文の始まりを表すキーワード
    start_keyword3 = r'^[0-9]*\. .*|^[０-９]*\. .*'
    com_start_keyword3 = re.compile(start_keyword3)

    # ・など文の始まりを表すキーワード
    start_keyword4 = r'^・.*|^▪.*|^･.*'
    com_start_keyword4 = re.compile(start_keyword4)

    # 1 や１ など文の始まりを表すキーワード
    start_keyword5 = r'^[0-9] .*|^[０-９] .*'
    com_start_keyword5 = re.compile(start_keyword5)

    # 注意事項を表すキーワード
    attract_keyword = r'^※.*'
    com_attract_keyword = re.compile(attract_keyword)

    # 。を含むが文の終わりでないパターン
    continue_keyword = r'.*という。\).*|.*。\).*|.*。\）.*'
    com_continue_keyword = re.compile(continue_keyword)
    
    # 表でたまに出でくるキーワード
    sum_keyword = r'^合計.*'
    com_sum_keyword = re.compile(sum_keyword)

    # 表の開始位置検出のためのパターン
    table_keyword = r'.*表.*[0-9]*-[0-9]*.*|.*表.*[０-９]*-[０-９]*.*'
    com_table_keyword = re.compile(table_keyword)

    # 先頭文字が数字であっても文の始まりでないパターン
    linkage_keyword = r'^[0-9]*時間.*|[0-9]*月.*|^[0-9]*日.*|^[0-9]*万.*^[0-9]*年.*|^[0-9]*ヶ月.*|^[0-9]*ヵ月.*|^[0-9]*,[0-9]*.*'
    com_linkage_keyword = re.compile(linkage_keyword)
    
    # 図の題名
    graph_keyword = r'^図.*'
    com_graph_keyword = re.compile(graph_keyword)


    splt21 = r'.*\).*|.*\].*|.*」.*|.*\）.*'
    resplt21 = re.compile(splt21)
    splt23 = r'.*\(.*|.*\[.*|.*「.*|.*\（.*'
    resplt23 = re.compile(splt23)



    for i, r in enumerate(rev_text_data):

        # 各パターン・キーワードとマッチング
        result_end_keyword1 = re.search(com_end_keyword1, r)
        result_end_keyword2 = re.search(com_end_keyword2, r)
        result_start_keyword1 = re.match(com_start_keyword1, r)
        result_start_keyword2 = re.match(com_start_keyword2, r)
        result_start_keyword3 = re.match(com_start_keyword3, r)
        result_start_keyword4 = re.match(com_start_keyword4, r)
        result_start_keyword5 = re.match(com_start_keyword5, r)
        result_continue_keyword = re.match(com_continue_keyword, r)
        result_attract_keyword = re.match(com_attract_keyword, r)
        result_sum_keyword = re.match(com_sum_keyword, r)
        result_table_keyword = re.match(com_table_keyword, r)
        result_linkage_keyword = re.match(com_linkage_keyword, r)
        result_graph_keyword = re.match(com_graph_keyword, r)
        result21 = re.match(resplt21, r)
        result23 = re.match(resplt23, r)

        # 。で終わらない場合が存在するため、次のテキストが文始まりのキーワードかどうか判定
        if i < len(rev_text_data) - 1:
            next2 = re.match(com_start_keyword1, rev_text_data[i + 1])
            next3 = re.match(com_start_keyword2, rev_text_data[i + 1])
            next5 = re.match(com_start_keyword3, rev_text_data[i + 1])
            next6 = re.match(com_start_keyword4, rev_text_data[i + 1])

        # 表を示そうとしているかどうかの判定
        if result_table_keyword:

            # 表というワードを含んでいるが、文が終わっていない場合はflagを1,文が終わっていたらflagを2
            if result_end_keyword2:
                flag = 2

            else:
                flag = 1

        # result21、23の条件内の処理は改善が必要
        # とりあえず、下の処理と同様のものをクローンしている
        if result21 and (not result23) and flag == 0:
            
            if result_start_keyword1:

                if len(S) > 0:
                        final_text_data.append(S)
                        S = ''

                S += re.sub(r'^\([0-9]*\)|^\([a-z]\)|^\([０-９]*\)|^\([ア-ン]*\)|^\（[0-9]*\）|^\（[a-z]*\）|^\（[０-９]*\）|^\（[ア-ン]*\）', '', r)
                
                if next2 or next3 or next5 or next6:
                    final_text_data.append(S)
                    S = ''

                flag = 0

            elif result_start_keyword2:

                if len(S) > 0:
                    final_text_data.append(S)
                    S = ''

                S += re.sub(r'^[ア-ヲ]\.|^[ア-ヲ] ', '', r)
                
                if next2 or next3 or next5 or next6:
                    final_text_data.append(S)
                    S = ''

                flag = 0
                
            elif result_start_keyword3:

                if len(S) > 0:
                    final_text_data.append(S)
                    S = ''

                S += re.sub(r'[0-9]*\.|[0-9]*-[0-9]*|-[0-9]*|[０-９]*\.|[０-９]*-[０-９]*|-[０-９]*', '', r)
                
                if next2 or next3 or next5 or next6:
                    final_text_data.append(S)
                    S = ''

                flag = 0

            elif r == '目次':

                if len(S) > 0:
                    final_text_data.append(S)
                    S = ''

                final_text_data.append(r)

            elif result_start_keyword4:

                if flag == 2:
                    final_text_data.append(r)

                else:
                    if len(S) > 0:
                        final_text_data.append(S)
                        S = ''

                    S += re.sub(r'^・|^▪|^･', '', r)

                    if next2 or next3 or next5 or next6:
                        final_text_data.append(S)
                        S = ''

            elif result_start_keyword5:

                if flag == 2:
                    final_text_data.append(r)
                
                else:

                    if result_linkage_keyword:
                        S += r

                        if r[len(r) - 1] == '。':
                            final_text_data.append(S)
                            S = ''
                        

                    else:
                        if len(S) > 0:
                            final_text_data.append(S)
                            S = ''
                            
                        S += re.sub(r'[0-9]*\.|[0-9]*-[0-9]*|-[0-9]*|[０-９]*\.|[０-９]*-[０-９]*|-[０-９]*|^[0-9]|^[０-９]', '', r)

                        if next2 or next3 or next5 or next6:
                            final_text_data.append(S)
                            S = ''

            elif result_attract_keyword:

                if flag == 2:
                    final_text_data.append(r)

                else:

                    if len(S) > 0:
                        final_text_data.append(S)
                        S = ''
                        
                    S += re.sub(r'^※', '', r)

                    if next2 or next3 or next5 or next6:
                        final_text_data.append(S)
                        S = ''

            elif result_sum_keyword:

                if flag == 2:
                    final_text_data.append(r)

                else:
                    final_text_data.append(S)
                    final_text_data.append(r)
                    S = ''

            elif result_graph_keyword:
                pass

            elif result_end_keyword1:

                if flag == 2:
                    final_text_data.append(r)

                else:
                    loc = r.count('。')

                    if loc == 1:

                        if result_end_keyword2:
                            S += re.sub('。', '', r)
                            final_text_data.append(S)
                            S = ''

                        else:
                            F = r.split('。')
                            S += F[0]
                            final_text_data.append(S)
                            S = F[1]

                    elif loc > 1:
                    
                        if loc > 2:
                            F = r.split('。')

                            if result_end_keyword2:

                                for f_i in len(F):
                                    _f = F[f_i]

                                    if f_i == 0:
                                        S += _f
                                    
                                    elif f_i == len(F) - 1:
                                        S = ''
                                        break

                                    if _f[0] == ')':
                                        S += _f
                                    
                                    else:
                                        final_text_data.append(S)
                                        S = _f

                            else:

                                for f_i in len(F):
                                    _f = F[f_i]

                                    if f_i == 0:
                                        S += _f

                                    if _f[0] == ')':
                                        S += _f
                                    
                                    else:
                                        final_text_data.append(S)
                                        S = _f

                        else:
                            S += re.sub('。', '', r)
                            final_text_data.append(S)
                            S = ''

                    if flag == 1:
                        flag = 2

            continue


        if result_start_keyword1:

            if len(S) > 0:
                    final_text_data.append(S)
                    S = ''

            S += re.sub(r'^\([0-9]*\)|^\([a-z]\)|^\([０-９]*\)|^\([ア-ン]*\)|^\（[0-9]*\）|^\（[a-z]*\）|^\（[０-９]*\）|^\（[ア-ン]*\）', '', r)
            
            if next2 or next3 or next5 or next6:
                final_text_data.append(S)
                S = ''

            flag = 0

        elif result_start_keyword2:

            if len(S) > 0:
                final_text_data.append(S)
                S = ''

            S += re.sub(r'^[ア-ヲ]\.|^[ア-ヲ] ', '', r)
            
            if next2 or next3 or next5 or next6:
                final_text_data.append(S)
                S = ''

            flag = 0
            
        elif result_start_keyword3:

            if len(S) > 0:
                final_text_data.append(S)
                S = ''

            S += re.sub(r'[0-9]*\.|[0-9]*-[0-9]*|-[0-9]*|[０-９]*\.|[０-９]*-[０-９]*|-[０-９]*', '', r)
            
            if next2 or next3 or next5 or next6:
                final_text_data.append(S)
                S = ''

            flag = 0

        elif r == '目次':

            if len(S) > 0:
                final_text_data.append(S)
                S = ''

            final_text_data.append(r)

        elif result_start_keyword4:

            if flag == 2:
                final_text_data.append(r)

            else:
                if len(S) > 0:
                    final_text_data.append(S)
                    S = ''

                S += re.sub(r'^・|^▪|^･', '', r)

                if next2 or next3 or next5 or next6:
                    final_text_data.append(S)
                    S = ''

        elif result_start_keyword5:

            if flag == 2:
                final_text_data.append(r)
            
            else:

                if result_linkage_keyword:
                    S += r

                    if r[len(r) - 1] == '。':
                        final_text_data.append(S)
                        S = ''
                    

                else:
                    if len(S) > 0:
                        final_text_data.append(S)
                        S = ''
                        
                    S += re.sub(r'[0-9]*\.|[0-9]*-[0-9]*|-[0-9]*|[０-９]*\.|[０-９]*-[０-９]*|-[０-９]*|^[0-9]|^[０-９]', '', r)

                    if next2 or next3 or next5 or next6:
                        final_text_data.append(S)
                        S = ''

        elif result_attract_keyword:

            if flag == 2:
                final_text_data.append(r)

            else:

                if len(S) > 0:
                    final_text_data.append(S)
                    S = ''
                    
                S += re.sub(r'^※', '', r)

                if next2 or next3 or next5 or next6:
                    final_text_data.append(S)
                    S = ''

        elif result_sum_keyword:

            if flag == 2:
                final_text_data.append(r)

            else:
                final_text_data.append(S)
                final_text_data.append(r)
                S = ''

        elif result_graph_keyword:
            pass

        elif result_end_keyword1:

            if flag == 2:
                final_text_data.append(r)

            else:
                loc = r.count('。')

                if loc == 1:

                    if result_end_keyword2:
                        S += re.sub('。', '', r)
                        final_text_data.append(S)
                        S = ''

                    else:
                        F = r.split('。')
                        S += F[0]
                        final_text_data.append(S)
                        S = F[1]

                elif loc > 1:
                 
                    if loc > 2:
                        F = r.split('。')

                        if result_end_keyword2:

                            for f_i in len(F):
                                _f = F[f_i]

                                if f_i == 0:
                                    S += _f
                                
                                elif f_i == len(F) - 1:
                                    S = ''
                                    break

                                if _f[0] == ')':
                                    S += _f
                                
                                else:
                                    final_text_data.append(S)
                                    S = _f

                        else:

                            for f_i in len(F):
                                _f = F[f_i]

                                if f_i == 0:
                                    S += _f

                                if _f[0] == ')':
                                    S += _f
                                
                                else:
                                    final_text_data.append(S)
                                    S = _f

                    else:
                        S += re.sub('。', '', r)
                        final_text_data.append(S)
                        S = ''

                if flag == 1:
                    flag = 2

        else:
            if flag == 2:
                final_text_data.append(r)

            else:
                S += r

    if not os.path.exists(file_path + '/excel_dataset'):
        os.mkdir(file_path + '/excel_dataset')
        print('Create directory \'excel_dataset\': path -> {}'.format(file_path + '/excel_dataset'))

    NFR = pd.DataFrame(final_text_data, columns = ['text'])
    NFR = NFR.replace(' ', '', regex = True)
    NFR.to_excel(file_path + '/excel_dataset/' + os.path.basename(excel_file), index = None, encoding = 'utf_8_sig')
    print('Save file converted to excel \'{}\':'.format(os.path.basename(excel_file)))
    print('   PATH =>> {}'.format(file_path + '/excel_dataset' + os.path.basename(excel_file)))




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
            print('\nCurrent directory: {}'.format(os.getcwd()))
            print('Directories or Files: \n')
            files = os.listdir(os.getcwd())
            k_words = re.compile(r'^.*\.txt$')
            ext_files = [f for f in files if re.match(k_words, f) or os.path.isdir(os.path.join(os.getcwd(), f))]
            ext_files = sorted(ext_files)
            pprint.pprint(ext_files)
            print('\nPlease input path:')
            file_path = input()

            if './' == file_path[0:2] or '/Users/' == file_path[0:7]:
                pass
            
            else:

                if '/' == file_path[0]:
                    file_path = './' + file_path[1:]
                
                else:
                    file_path = './' + file_path

            while True:
                print('\nCurrent directory: {}'.format(file_path))
                print('Directories or Files: \n')
                files = os.listdir(file_path)
                k_words = re.compile(r'^.*\.txt$')
                ext_files = [f for f in files if re.match(k_words, f) or os.path.isdir(os.path.join(file_path, f))]
                ext_files = sorted(ext_files)
                pprint.pprint(ext_files)
                print('\nPlease input path:')
                path = input()

                confirm = ''

                if len(path) == 0:

                    while True:
                        print('Finish to input path (y/n):')
                        confirm = input()

                        if 'y' == confirm or 'n' == confirm:
                            break

                        else:
                            print('Input y or n')
                    
                    if 'y' == confirm:

                        if os.path.exists(file_path):
                            files = os.listdir(file_path)
                            k_words = re.compile(r'^.*\.txt$')
                            file_paths = [file_path + '/' + f for f in files if re.match(k_words, f)]
                        break
                    
                    elif 'n' == confirm:
                        pass

                elif '..' == path:
                    file_path = re.sub('/' + os.path.basename(file_path), '', file_path)
                
                else:

                    if os.path.exists(os.path.join(file_path, path)):
                        file_path = os.path.join(file_path, path)

                        if re.match(r'^.*\.txt$', file_path):
                            file_paths = [file_path]
                            file_path = re.sub('/' + os.path.basename(file_path), '', file_path)
                            break
                    
                    else:
                        print('Input correct path')

        else:
            file_path = sys.argv[1]

            if len(sys.argv) > 2:
                option = sys.argv[2]

                if option in _options:

                    if option == _options[0]:
                        files = os.listdir(file_path)
                        k_words = re.compile(r'^.*\.txt$')
                        file_paths = [file_path + '/' + f for f in files if re.match(k_words, f)]
                        pprint.pprint(file_paths)

                    elif option == _options[1]:
                        print()
                        files = os.listdir(file_path)
                        k_words = re.compile(r'^.*\.txt$')
                        [print(f) for f in files if re.match(k_words, f)]
                        print('\nSelect file:')
                        file_paths = [file_path + '/' + input()]

                else:
                    raise NoSuchOption('Selected option is incorrect. Please select --help option and read usage.')

            else:

                if re.match(r'^.*\.txt$', file_path):
                    file_paths = [file_path]
                    file_path = re.sub('/' + os.path.basename(file_path), '', file_path)

                else:
                    raise FileExtensionError('File extension must be text.')

        if not '--help' in sys.argv:
            for text_file in file_paths:
                text2excel(text_file, file_path)

    except:
        traceback.print_exc()



if __name__ == '__main__':
    main()



