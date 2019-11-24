# -*- coding: utf-8 -*-
import json
import os
import time
import re
import sys
import unicodedata

from prepro import preprocessing_mark2
from synonym.wordVecMaker import wordVecMaker

from gensim import corpora, matutils

import pandas as pd
import numpy as np

import time

from collections import Counter

from preprocessings.ja.stopwords import get_stop_words
from preprocessings.ja.cleaning import clean_text
from preprocessings.ja.normalization import normalize
#from preprocessings.ja.tokenizer import MeCabTokenizer, JanomeTokenizer
from preprocessings.ja.tokenizer import NagisaTokenizer
from preprocessings.ja import zen2han
#DATA_DIR = os.path.join(os.getcwd(), 'data/processed')

from sklearn.model_selection import train_test_split
from sklearn.feature_extraction.text import CountVectorizer, TfidfVectorizer
from sklearn.feature_extraction.text import TfidfTransformer
from sklearn.model_selection import GridSearchCV
from sklearn.pipeline import Pipeline
from sklearn.metrics import classification_report, accuracy_score, precision_score, recall_score, f1_score
from sklearn.metrics import roc_curve
import matplotlib.pyplot as plt

sys.path.append('libsvm/python/')
from svm import svm_parameter, svm_problem
import svmutil

from sklearn.neighbors import KNeighborsClassifier
from sklearn.ensemble import ExtraTreesClassifier, RandomForestClassifier
from sklearn.naive_bayes import GaussianNB

from preprocessings.ja.cleaning import clean_text
from preprocessings.ja.normalization import normalize
from preprocessings.ja.tokenizer import MeCabTokenizer, JanomeTokenizer
from preprocessings.ja.tokenizer import NagisaTokenizer

from multiprocessing import Pool



# 形態素解析（分かち書き）
def with_preprocess(text, phrase):
    # Nagisa形態素解析ツール
    tokenizer = NagisaTokenizer(phrase)
    words = clean_text(text)
    words = tokenizer.tokenize(words)

    return words


# RandomForest
def build_pipeline(stopwords):
    parameters = {'n_estimators': [10, 30, 50, 70, 90, 110, 130, 150], 'max_features': ['auto', 'sqrt', 'log2', None]}
    text_clf = Pipeline([('vect', TfidfVectorizer(stop_words=stopwords)),
                         ('tfidf', TfidfTransformer()),
                         ('clf', GridSearchCV(RandomForestClassifier(), parameters, cv=2, scoring='accuracy', n_jobs=-1)),
                         ])
    return text_clf

# この関数何かわからん
def isnan(value):
    try:
        import math
        return math.isnan(float(value))
    except:
        return


def classify(classification_model, stopwords, Normalization=True, pickup_pos=[], load_model='word2vec'):

    NFR=["機能適合性", "性能効率性", "互換性", "使用性", "信頼性", "セキュリティ", "保守性", "移植性"]
    # モデル学習用のデータ
    train = pd.read_csv('text.csv')

    # テキストのクリーニング
    train['要件'] = train['要件'].replace(r'[【】]', ' ', regex = True)
    train['要件'] = train['要件'].replace(r'[（）()]', ' ', regex = True)
    train['要件'] = train['要件'].replace(r'[［］\[\]]', ' ', regex = True)
    train['要件'] = train['要件'].replace(r'[@＠]\w+', '', regex = True)
    train['要件'] = train['要件'].replace(r'https?:\/\/.*?[\r\n ]', '', regex = True)
    train['要件'] = train['要件'].replace(r'\n', ' ', regex = True)
    train['要件'] = train['要件'].replace(r'　', '', regex = True)
    train['要件'] = train['要件'].replace(' ', '', regex = True)
    train['要件'] = train['要件'].replace(r'・|/', '、', regex = True)
    train['要件'] = train['要件'].replace(r',', '', regex = True)
    train['要件'] = train['要件'].replace(r'^[0-9]+', '', regex = True)
    train['要件'] = train['要件'].replace(r'[0-9]+', '0', regex = True)

    train_text = train["要件"].tolist()
    train_label = train["最終判断"].tolist()

    train_wakati = []
    # テキストから分かち書きに変換
    for line in train_text:
        pre = preprocessing_mark2(type='wakati')
        train_wakati.extend([with_preprocess(line, [])])

    # ストップワードを除去した単語リスト、単語の重み付け（Tf-Idf）
    w_train, train_feature = pre.rem_stopwords(text = train_wakati)
    train_matrix = pd.DataFrame(w_train.toarray(), columns = train_feature)

    train_docs, docs_tmp1 = [], []
    # 各文書におけるストップワードの除去（積演算）
    docs_tmp1 = [set(t_w.split(' ')) & set(train_feature) for t_w in train_wakati]

    # 重要語の抽出（Tf-Idfの重み）
    for i, doc in enumerate(docs_tmp1):
        docs_tmp2 = []

        for d in doc:
            # 閾値を超えた単語のみ抽出
            if train_matrix[d][i] >= 0.0:
                docs_tmp2.append(d)

        train_docs.append(docs_tmp2)

    labels = []
    # 空白がないリストに変換
    for num, label in enumerate(train_label):
        if isnan(label) != True:
            labels.append((label))

    # 自動分類処理用ファイルの読み込み
    path = './dataset/pdf_dataset/excel_dataset/'
    files = os.listdir(path)
    k_words = re.compile(r'^.*\.xlsx$')
    ext_files = [os.path.join(path, f) for f in files if re.match(k_words, f)]

    for file_name in ext_files:
        print('Open file: {}'.format(file_name))
        b_file_name = os.path.basename(file_name)
        b_file_name = re.sub(r'xlsx', 'csv', b_file_name)
        save_file_name = '/Users/yuta/oss/Text Mining/dataset/pdf_dataset/excel_dataset/predict_' + b_file_name

        if os.path.exists(save_file_name):
            continue

        original_predict = pd.read_excel(file_name, encoding='UTF-8')
        predict = original_predict.copy()
        predict['text'] = predict['text'].replace(r'[【】]', ' ', regex = True)
        predict['text'] = predict['text'].replace(r'[（）()]', ' ', regex = True)
        predict['text'] = predict['text'].replace(r'[［］\[\]]', ' ', regex = True)
        predict['text'] = predict['text'].replace(r'[@＠]\w+', '', regex = True)
        predict['text'] = predict['text'].replace(r'https?:\/\/.*?[\r\n ]', '', regex = True)
        predict['text'] = predict['text'].replace(r'\n', ' ', regex = True)
        predict['text'] = predict['text'].replace(r'　', '', regex = True)
        predict['text'] = predict['text'].replace(' ', '', regex = True)
        predict['text'] = predict['text'].replace(r'・|/', '、', regex = True)
        predict['text'] = predict['text'].replace(r',', '', regex = True)
        predict['text'] = predict['text'].replace(r'^[0-9]+', '', regex = True)
        predict['text'] = predict['text'].replace(r'[0-9]+', '0', regex = True)
        predict_text = predict['text'].tolist()
        
        predict_wakati = []
        for line2 in predict_text:
            pre = preprocessing_mark2(type='wakati')
            predict_wakati.extend([with_preprocess(line2, [])])

        w_predict, predict_feature = pre.rem_stopwords(text = predict_wakati)
        predict_matrix = pd.DataFrame(w_predict.toarray(), columns = predict_feature)

        predict_docs, docs_tmp1 = [], []
        docs_tmp1 = [set(p_w.split(' ')) & set(predict_feature) for p_w in predict_wakati]

        # 重要語の抽出（Tf-Idfの重み）
        for i, doc in enumerate(docs_tmp1):
            docs_tmp2 = []

            for d in doc:
                # 閾値を超えた単語のみ抽出
                if predict_matrix[d][i] >= 0.0:
                    docs_tmp2.append(d)

            predict_docs.append(docs_tmp2)

        sta = 0
        cnt = 1

        if load_model == 'word2vec':
            model_path = 'synonym/entity_vector/entity_vector.model.bin'
        elif load_model == 'fasttext':
            model_path = 'synonym/fasttext_vector/fasttext_model.bin'

        final_result = pd.DataFrame()
        for i, nfr in enumerate(NFR, 1):

            for count in range(sta, cnt):  # 処理のループ回数
                train_docs_, labels_, _labels_ = [], [], []

                for num, label in enumerate(labels):
                    # if nfr in label:
                    
                    if labels[num] in nfr:
                        _labels_.append(i)
                        ex = [train_text[num], train_docs[num]]
                        #ex = [wakati[num], docs[num]]
                        train_docs_.append(ex)

                    elif labels[num] not in nfr:
                        _labels_.append(0)
                        ex = [train_text[num], train_docs[num]]
                        #ex = [wakati[num], docs[num]]
                        train_docs_.append(ex)

                labels_train_ = _labels_

                '''
                data_train_, data_test_, labels_train_, labels_test_ = train_test_split(docs_, _labels_, train_size=rate, stratify = _labels_)

                train, test = [], []


                for j in range(len(data_train_)):  # data_train ▶︎ train
                    train.append(data_train_[j][1])

                for k in range(len(data_test_)):  # data_test ▶︎ test
                    test.append(data_test_[k][1])
                '''
                train = train_docs
                test = predict_docs

                dense_all_test = []
                dense_all_train = []

                # siki = [機能適合性・性能効率性・互換性・使用性・信頼性・セキュリティ・保守性・移植性]
                if classification_model == 'K-NN':
                    siki = [0.0, 0.4, 0.4, 0.1, 0.35, 0.35, 0.2, 0.3]
                elif classification_model == 'NB':
                    siki = [0.3, 0.05, 0.3, 0.0, 0.05, 0.0, 0.1, 0.0]
                elif classification_model == 'SMO':
                    siki = [0.3, 0.3, 0.2, 0.3, 0.2, 0.0, 0.2, 0.05]
                elif classification_model == 'RandomForest':
                    siki = [0.35, 0.0, 0.3, 0.4, 0.3, 0.3, 0.15, 0.4]

                dictionary = corpora.Dictionary(train)

                if len(pickup_pos) == 0:
                    add_path = 'pro_'
                elif len(pickup_pos) > 0:
                    add_path = 'spe_'

                if Normalization == False:
                    docs_train = train
                    docs_test = test
                    add_path = "正規化なし_"
                    load_model = 'NO'
                elif Normalization == True:
                    w1 = wordVecMaker(tokens=train, threshold=siki[i-1], nfr=nfr, count=count, classify=classification_model, path=add_path, load_model=load_model)
                    docs_train = w1.synonimTransfer(sentences=train, synonyms=w1.get_synonym(model_path=model_path))
                    # w2 = get_synonym(test, siki)
                    docs_test = w1.synonimTransfer(sentences=test, synonyms=w1.get_synonym(model_path=model_path))
                    add_path = "正規化あり_"
                
                
                bow_corpus_train = [dictionary.doc2bow(d) for d in docs_train]
                bow_corpus_test = [dictionary.doc2bow(d) for d in docs_test]

                for bow in bow_corpus_train:
                    dense = list(matutils.corpus2dense([bow], num_terms=len(dictionary)).T[0])
                    dense_all_train.append(dense)

                for bow2 in bow_corpus_test:
                    dense2 = list(matutils.corpus2dense([bow2], num_terms=len(dictionary)).T[0])
                    dense_all_test.append(dense2)
                
                #label_predict_

                # estimator.fit(dense_all_train, labels_train_)
                # label_predict_ = estimator.predict(dense_all_test)

                if classification_model == 'K-NN':
                    knc = KNeighborsClassifier(n_neighbors=1)
                    knc.fit(dense_all_train, labels_train_)
                    label_predict_ = knc.predict(dense_all_test)
                if classification_model == 'SMO':
                    prob = svm_problem(labels_train_, dense_all_train)
                    param = svm_parameter("-s 0 -t 0")
                    model = svmutil.svm_train(prob, param)
                    label_predict_, accuracy, dec_values = svmutil.svm_predict([], dense_all_test, model)
                if classification_model == 'NB':
                    clf = GaussianNB()
                    clf.fit(dense_all_train, labels_train_)
                    label_predict_ = clf.predict(dense_all_test)
                if classification_model == 'RandomForest':
                    """
                    search_params = {
                        'n_estimators': [5, 10, 20, 30, 50, 100, 300],
                        'max_features': [3, 5, 10, 15, 20],
                        'random_state': [0, 2525],
                        'n_jobs': [1],
                        'min_samples_split': [3, 5, 10, 15, 20, 25, 30, 40, 50, 100],
                        'max_depth': [3, 5, 10, 15, 20, 25, 30, 40, 50, 100]
                    }
                    gs = GridSearchCV(RandomForestClassifier(), search_params, cv=3, verbose=True, n_jobs=-1)
                    gs.fit(dense_all_train, labels_train_)
                    print(gs.best_estimator_)
                    """
                    clf = RandomForestClassifier()
                    clf.fit(dense_all_train, labels_train_)
                    label_predict_ = clf.predict(dense_all_test)
                """
                if self.classification_model == 'DeepForest':
                    dense_all_train = np.array(dense_all_train)
                    labels_train_ = np.array(labels_train_)
                    mgc_forest = MGCForest(
                        estimators_config={
                            'mgs': [{
                                'estimator_class': ExtraTreesClassifier,
                                'estimator_params': {
                                    'n_estimators': 30,
                                    'min_samples_split': 21,
                                    'n_jobs': -1,
                                }
                            }, {
                                'estimator_class': RandomForestClassifier,
                                'estimator_params': {
                                    'n_estimators': 30,
                                    'min_samples_split': 21,
                                    'n_jobs': -1,
                                }
                            }],
                            'cascade': [{
                                'estimator_class': ExtraTreesClassifier,
                                'estimator_params': {
                                    'n_estimators': 1000,
                                    'min_samples_split': 11,
                                    'max_features': 1,
                                    'n_jobs': -1,
                                }
                            }, {
                                'estimator_class': ExtraTreesClassifier,
                                'estimator_params': {
                                    'n_estimators': 1000,
                                    'min_samples_split': 11,
                                    'max_features': 'sqrt',
                                    'n_jobs': -1,
                                }
                            }, {
                                'estimator_class': RandomForestClassifier,
                                'estimator_params': {
                                    'n_estimators': 1000,
                                    'min_samples_split': 11,
                                    'max_features': 1,
                                    'n_jobs': -1,
                                }
                            }, {
                                'estimator_class': RandomForestClassifier,
                                'estimator_params': {
                                    'n_estimators': 1000,
                                    'min_samples_split': 11,
                                    'max_features': 'sqrt',
                                    'n_jobs': -1,
                                }
                            }]
                        }
                    )
                    mgc_forest.fit(dense_all_train, labels_train_)
                    label_predict_ = mgc_forest.predict(dense_all_test)
                """

                '''
                acc_score_ = accuracy_score(labels_test_, label_predict_)
                pre_score_ = precision_score(labels_test_, label_predict_, pos_label = i) #, average=None)
                recall_score_ = recall_score(labels_test_, label_predict_, pos_label = i) #, average=None)
                f1_score_ = f1_score(labels_test_, label_predict_, pos_label = i) #, average=None)
                '''
                #plane_csv = [x[1] for x in data_train_]
                #plane_csv_test = [y[1] for y in data_test_]

                #df1 = pd.DataFrame({'元ラベル': labels_train_, '原文': plane_csv, '使用した単語': train, '変換後': docs_train}, columns=["元ラベル", "予測ラベル", "原文", "使用した単語", "変換後"])

                """
                df2 = pd.DataFrame({'元ラベル': labels_test_, '予測ラベル': label_predict_, '原文': plane_csv_test, '使用した単語': test,'変換後':docs_test},columns=["元ラベル", "予測ラベル", "原文", "使用した単語","変換後"])
                df3 = pd.concat([df1, df2])
                df3.to_csv('/Users/s206050/NFR分類/実験結果/list/RQ1/' + classify + '/' + nfr + '/list' + nfr + str(count)+'___'+ No + '.csv')  # csv書き出し先設定
                """  # 必ず戻す！！

                '''
                print(str(count) +  " : " + classification_model +  " : " + nfr)
                print("正答率　　  : " + str(acc_score_))  # accuracyの出力
                print("Precision　: " + str(pre_score_))  # precisionの出力
                print("Recall　　 : " + str(recall_score_))  # recallの出力
                print("F値　      : " + str(f1_score_))
                print("--------------------------------------------")
                

                # dense_all_train = []
                # dense_all_test = []

                pre_sum += pre_score_
                rec_sum += recall_score_
                f1_sum += f1_score_

                pre_list.append(pre_score_)
                rec_list.append(recall_score_)
                f1_list.append(f1_score_)

                # print(count + 1)
                '''
            
            '''
            pre_list.append(pre_sum / (cnt - sta))
            rec_list.append(rec_sum / (cnt - sta))
            f1_list.append(f1_sum / (cnt - sta))
            result = pd.DataFrame({'Precision': pre_list, 'Recall': rec_list, 'F1_score': f1_list})
            
            # df4.to_csv('/Users/s206050/NFR分類/実験結果/分類精度/RQ1/' + classify + '/'+ nfr + '分類精度_' + No + '.csv')  # csv書き出し先設定
            tmp_df = result[result.index == cnt]
            tmp_df.index = [rate]
            final_result = pd.concat([final_result, tmp_df])
            '''

            '''
            dir_path = '実験/NFR分類/閾値評価/' + classification_model + '/' + load_model + '/' + nfr + '/'
            if not os.path.exists(dir_path):
                os.makedirs(dir_path)
            result.to_csv(dir_path + add_path + 'preprocessing_' + nfr + str(siki[i-1]) + '_ver2.0.csv', encoding = 'utf_8_sig')
            '''

            nfr_df = pd.DataFrame({nfr: label_predict_})
            nfr_df[nfr] = nfr_df[nfr].replace(i, nfr, regex=True)
            nfr_df[nfr] = nfr_df[nfr].replace(0, '', regex=True)
            final_result = pd.concat([final_result, nfr_df], axis=1)
        #file_name = 'data/result/' + classification_model + '_' + nfr + '.csv'
        #final_result.to_csv(file_name, encoding = 'utf_8_sig')
        csv_output = pd.DataFrame({'要件': predict_text})
        csv_output = pd.concat([csv_output, final_result], axis=1)
        csv_output.to_csv(save_file_name, encoding = 'utf_8_sig')
        print('Finish...')

        
        
"""
並列処理をするためのメソッド
"""
def wrapper(arg):
    classify(*arg)

if __name__ == "__main__":
    #classification_model = ["NB", "`SMO", "K-NN", "RandomForest"]
    classification_model = ['SMO']#["SMO", "RandomForest"]
    pu_pos = ["名詞"]
    time_list = []
    parameter_list = []

    for model in classification_model:
        classify(model, stopwords = [])