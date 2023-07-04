import pandas as pd
import os
import jieba
import math



origin_fold_path = 'D:\\Code\\Python\\转置数据\\origin_data'
trans_fold_path = 'D:\\Code\\Python\\转置数据\\trans_data'


items = ['资产', '负债', '损失', '利益']
items_have = [0, 0, 0, 0]

sim_threshold = 0.8

class item:
    def __init__(self, itm):
        self.itm = itm
        self.names = []

def cosineSimilarity(list_A, list_B):
    listA_cut = [i for i in jieba.cut(list_A, cut_all=True) if i != '']
    listB_cut = [i for i in jieba.cut(list_B, cut_all=True) if i != '']
    word_set = set(listA_cut).union(set(listB_cut))
    word_dict = dict()
    for word, i in zip(word_set, range(len(word_set))):
        word_dict[word] = i

    listA_cut_code = [0]*len(word_dict)
    listB_cut_code = [0]*len(word_dict)

    for word in listA_cut:
        listA_cut_code[word_dict[word]]+=1

    for word in listB_cut:
        listB_cut_code[word_dict[word]]+=1

    sum, sqA, sqB = [0, 0, 0]
    for i in range(len(word_dict)):
        sum += listA_cut_code[i] * listB_cut_code[i]
        sqA += pow(listA_cut_code[i], 2)
        sqB += pow(listB_cut_code[i], 2)

    try:
        result = round(float(sum)/(math.sqrt(sqA)*math.sqrt(sqB)), 2)
    except ZeroDivisionError:
        result = 0.0

    return result



if __name__ == "__main__":
    origin_files = [file for file in os.listdir(origin_fold_path) if file.endswith(".xlsx")]
    for file in origin_files:
        file_data = pd.read_excel(origin_fold_path+'\\'+file, header=0, keep_default_na=False)
        output_data = {'银行名':[], '年份':[]}
        bank_name = ''
        time_list = []

        item_list = []
        for itm in items:
            item_list.append(item(itm))

        for row in range(file_data.shape[0]):
            if(file_data.iloc[row][1] == 'BanksName'):
                bank_name = file_data.iloc[row][2]

            for i in range(len(items)):
                if(file_data.iloc[row][0] == items[i]):
                    if(file_data.iloc[row][1] not in item_list[i].names and (file_data.iloc[row][4] == '' or str(file_data.iloc[row][4])[-1] != '部')):
                        sim = False
                        for head in item_list[i].names:
                           if(cosineSimilarity(head, file_data.iloc[row][1]) > sim_threshold):
                                sim = True
                                break
                        if(not sim):
                            item_list[i].names.append(file_data.iloc[row][1])  
            
            if(file_data.iloc[row][0] in items  and  file_data.iloc[row][3] not in time_list):
                output_data['银行名'].append(bank_name)
                output_data['年份'].append(file_data.iloc[row][3])
                time_list.append(file_data.iloc[row][3])


        
        for itm in item_list:
            for head in itm.names:
                if head == '合计':
                    head = itm.itm + ' 合计'
                output_data[head] = []

            for row in range(file_data.shape[0]):
                if(file_data.iloc[row][0] == itm.itm and (file_data.iloc[row][4] == '' or str(file_data.iloc[row][4])[-1] != '部')):
                    head = file_data.iloc[row][1]
                    data = file_data.iloc[row][2]
                    time = file_data.iloc[row][3]
                    data_num = time_list.index(time)
                    
                    for name in itm.names:
                        if(head != name and  cosineSimilarity(head, name) > sim_threshold):
                            print(head, name)
                            head = name
                            
                            break

                    if head == '合计':
                        head = itm.itm + ' 合计'

                    for _ in range(data_num - len(output_data[head])):
                        output_data[head].append('')
                    output_data[head].append(data)



        max_len = 0
        for key in output_data:
            if(len(output_data[key]) > max_len):
                max_len = len(output_data[key])
        print(max_len)

        for key in output_data:
            for _ in range(max_len - len(output_data[key])):
                output_data[key].append('')


        df = pd.DataFrame.from_dict(output_data)
        df.to_excel(trans_fold_path+'\\'+file)

