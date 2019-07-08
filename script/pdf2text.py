import os
import sys
import re
import pandas as pd

from tika import parser


def pdf2text(pdf_file):
    parsed = parser.from_file(pdf_file)
    save_path = re.sub(r'\.pdf', '.txt', pdf_file)
    with open(save_path, 'w') as f:
        f.write(parsed["content"])

    with open(save_path, encoding = 'utf-8-sig') as f:
        read_text = list(f)
    text_df = pd.DataFrame(read_text)
    text_df.replace('\n', '', regex = True, inplace = True)
    #text_df.replace(' ', '', regex = True, inplace = True)
    text_df.replace(' $', '', regex = True, inplace = True)
    text_df = text_df[text_df != '']
    text_df = text_df[text_df != ' ']

    text_df.dropna(inplace = True)
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
                    raise NoSuchOption('Selected option is incorrect. Please select --help option and read usage.')

            else:

                if re.match(r'^.*\.pdf$', file_path):
                    file_paths = [file_path]
                    file_path = re.sub('/' + os.path.basename(file_path), '', file_path)

                else:
                    raise FileExtensionError('File extension must be text.')

        if not '--help' in sys.argv:
            for pdf_file in file_paths:
                pdf2text(pdf_file)

    except:
        traceback.print_exc()



if __name__ == '__main__':
    main()