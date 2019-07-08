# ターミナルで実行する場合、
# python pdf2text.py < ファイルパス >
# python pdf2text.py < ディレクトリパス > < -a or -l >
# python pdf2text.py -f
# python pdf2text.py --help
# のいずれかで実行してください。
# -a オプションは、指定したディレクトリパス内に存在する全てのPDFファイルをテキストファイルに変換します。
# -l オプションは、指定したディレクトリパス内のファイルを表示します。
# -f オプションは、まずカレントディレクトリのフォルダとPDFファイルを表示します。
# 変換したいPDFファイルが違う階層にある場合、フォルダ名を入力し、一つずつ階層を移動します。
# 変換したPDFファイルが見つかれば、PDFファイル名を入力するか、何も入力せずにエンターキーを押してください。
# 何も入力しなかった場合、確認が行われます。
# yを入力した場合は、ディレクトリ内の全てのPDFファイルがテキストファイルに変換されます。
# nを入力した場合は、もう一度入力フェーズに移ります。誤ってエンターキーを押した場合などに使用してください。
# --help オプションは、pdf2textの簡単な説明が表示されます。




import os
import sys
import re
import pandas as pd

# pip install tika
# Apache TikaのPython用モジュール
# PDFをテキストに変換する
from tika import parser


# 存在しないオプションを指定した場合のエラー
class NoSuchOption(Exception):
    pass

# 拡張子がテキストでない場合のエラー
class FileExtensionError(Exception):
    pass


# PDFをテキストに変換する関数
def pdf2text(pdf_file):
    
    # PDFのファイルパスを入れると勝手にパースしてくれる
    parsed = parser.from_file(pdf_file)
    save_path = re.sub(r'\.pdf', '.txt', pdf_file)
    
    # parsedは一行の文字列となっているため、
    # 一度テキストファイルに出力する
    with open(save_path, 'w') as f:
        f.write(parsed["content"])

    # 保存したテキストファイルを再度読み込む
    with open(save_path, encoding = 'utf-8-sig') as f:
        # データフレームで処理を行うためリストに変換
        read_text = list(f)
    
    # 一度改行コード除去する必要があるか不明（笑）
    text_df = pd.DataFrame(read_text)
    text_df.replace('\n', '', regex = True, inplace = True)
    #text_df.replace(' ', '', regex = True, inplace = True)
    text_df.replace(' $', '', regex = True, inplace = True)
    
    # 欠損値や空白の要素は除去
    text_df = text_df[text_df != '']
    text_df = text_df[text_df != ' ']
    text_df.dropna(inplace = True)
    
    # テキストファイルとして保存するために
    # 改行コードで連結して再び文字列へ変換
    text_list = text_df[0].tolist()
    text_data = '\n'.join(text_list)
    
    with open(save_path, 'w') as f:
        f.write(text_data)


# 実行ファイルのオプション
_options = ['-a', '-l']

def main():
    try:

        # オプションの説明
        if '--help' in sys.argv:
            print('\nUsage: python pdf2excel.py < text file > < -a or -l > or python pdf2excel.py < -f or --help >')
            print()
            print('Some useful options are:')
            print('   -a  : All PDF files that exist in selected directory convert to text files.')
            print('   -l  : Show all PDF files that exist in selected directory.')
            print('   -f  : Show all PDF files that exist in current directory and input directory or file path of file that you want to convert to.')
            print('\nIf you want to read usage, you select \'--help\' option, please.')
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
                k_words = re.compile(r'^.*\.pdf$')
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
                            print('Please input y or n.\n')
                    
                    if 'y' == confirm:

                        if os.path.exists(file_path):
                            files = os.listdir(file_path)
                            k_words = re.compile(r'^.*\.pdf$')
                            file_paths = [file_path + '/' + f for f in files if re.match(k_words, f)]
                        break
                    
                    elif 'n' == confirm:
                        pass

                elif '..' == path:
                    file_path = re.sub('/' + os.path.basename(file_path), '', file_path)
                
                else:

                    if os.path.exists(os.path.join(file_path, path)):
                        file_path = os.path.join(file_path, path)

                        if re.match(r'^.*\.pdf$', file_path):
                            file_paths = [file_path]
                            file_path = re.sub('/' + os.path.basename(file_path), '', file_path)
                            break
                    
                    else:
                        print('Incorrect path!!\n')

        else:
            file_path = sys.argv[1]

            if len(sys.argv) > 2:
                option = sys.argv[2]

                if option in _options:

                    if option == _options[0]:
                        files = os.listdir(file_path)
                        k_words = re.compile(r'^.*\.pdf$')
                        file_paths = [file_path + '/' + f for f in files if re.match(k_words, f)]
                        pprint.pprint(file_paths)

                    elif option == _options[1]:
                        print()
                        files = os.listdir(file_path)
                        k_words = re.compile(r'^.*\.pdf$')
                        [print(f) for f in files if re.match(k_words, f)]
                        print('\nSelect file:')
                        file_paths = [file_path + '/' + input()]

                else:
                    # 正しいオプションが選択されなかった時にエラーを出力
                    raise NoSuchOption('Selected option is incorrect. Please select --help option and read usage.')

            else:

                if re.match(r'^.*\.pdf$', file_path):
                    file_paths = [file_path]
                    file_path = re.sub('/' + os.path.basename(file_path), '', file_path)

                else:
                    # 指定されたPDFファイルが存在しなかった時にエラーを出力
                    raise FileExtensionError('File extension must be text.')

        if not '--help' in sys.argv:
            for pdf_file in file_paths:
                pdf2text(pdf_file)

    except:
        traceback.print_exc()



if __name__ == '__main__':
    main()
