import os
import sys
import re
import pandas as pd


file_path = os.getcwd() + '/dataset/pdf_dataset'
files = os.listdir(file_path)
k_words = re.compile(r'^.*\.txt$')
file_paths = [file_path + '/' + f for f in files if re.match(k_words, f)]

head_word = pd.DataFrame()

for path in file_paths:
    list_head_word = []
    with open(path, mode='rt', encoding='utf-8-sig') as f:
        text = list(f)
    for t in text:
        list_head_word.append(t[:5])
    text_df = pd.DataFrame(list_head_word, columns = ['header'])
    text_df.replace('\n', '', regex = True, inplace = True)
    text_df.replace(' $', '', regex = True, inplace = True)
    head_word = pd.concat([head_word, text_df])
    
dic_head_word = head_word['header'].value_counts().to_dict()
pd.DataFrame(list(dic_head_word.items()), columns = ['header', 'count']).to_csv('head_word_patterns.csv', encoding = 'utf_8_sig')