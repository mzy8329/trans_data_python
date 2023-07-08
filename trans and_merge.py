import pandas as pd
import os



origin_fold_path = 'D:\\Code\\Python\\转置数据\\origin_data'
trans_fold_path = 'D:\\Code\\Python\\转置数据\\trans_data'

merged_file_name = 'merged_file.xlsx'

items = ['资产', '负债', '损失', '利益']
file_inf = ['ID', '全行员生总数', '分行总数', '支行总数', '办事处和寄庄总数量', 
            '初设年份', '初设资本', '初设资本实收', '现资本', '现资本实收', ]

class item:
    def __init__(self, itm):
        self.itm = itm
        self.names = []

class file_c:
    def __init__(self):
        self.name = ''
        self.inf = ['' for inf in file_inf]
        self.time_list = []
        self.item_list = [item(itm) for itm in items]
        self.trouble = False

if __name__ == "__main__":
    origin_files = [file for file in os.listdir(origin_fold_path) if file.endswith(".xlsx")]
    file_c_lists = [file_c() for _ in range(len(origin_files))]

# trans
    for file, file_index in zip(origin_files, range(len(origin_files))):
        file_data = pd.read_excel(origin_fold_path+'\\'+file, header=0, keep_default_na=False)
        output_data = {'银行名':[], '年份':[]}

        for row in range(file_data.shape[0]):
            if(file_data.iloc[row][1] == 'BanksName'):
                    file_c_lists[file_index].name = file_data.iloc[row][2]
            for inf_index in range(len(file_inf)):
                if(file_data.iloc[row][1] == file_inf[inf_index]):
                    if(inf_index >= 5 and str(file_data.iloc[row][2])[-1] == '万'):
                        data_tmp = float(file_data.iloc[row][2][:-1])*10000
                        file_c_lists[file_index].inf[inf_index]  = str(data_tmp)[:-2]
                    else:
                        file_c_lists[file_index].inf[inf_index]  = file_data.iloc[row][2]

            
            for i in range(len(items)):
                if(file_data.iloc[row][0] == items[i]):
                    if(file_data.iloc[row][1] not in file_c_lists[file_index].item_list[i].names and (file_data.iloc[row][4] == '' or str(file_data.iloc[row][4])[-1] != '部')):
                        file_c_lists[file_index].item_list[i].names.append(file_data.iloc[row][1])  
            
            if(file_data.iloc[row][0] in items  and  file_data.iloc[row][3] not in file_c_lists[file_index].time_list):
                output_data['银行名'].append(file_c_lists[file_index].name)
                output_data['年份'].append(file_data.iloc[row][3])
                file_c_lists[file_index].time_list.append(file_data.iloc[row][3])

        for itm in file_c_lists[file_index].item_list:
            for head in itm.names:
                if head == '合计':
                    head = itm.itm + ' 合计'
                output_data[head] = []

            for row in range(file_data.shape[0]):
                if(file_data.iloc[row][0] == itm.itm and (file_data.iloc[row][4] == '' or str(file_data.iloc[row][4])[-1] != '部')):
                    head = file_data.iloc[row][1]
                    data = file_data.iloc[row][2]
                    time = file_data.iloc[row][3]
                    data_num = file_c_lists[file_index].time_list.index(time)

                    if head == '合计':
                        head = itm.itm + ' 合计'

                    for _ in range(data_num - len(output_data[head])):
                        output_data[head].append('')
                    output_data[head].append(data)

            if(itm.itm + ' 合计' in  output_data.keys()):
                value_tmp = output_data[itm.itm + ' 合计']
                output_data.pop(itm.itm + ' 合计', None)
                output_data[itm.itm + ' 合计'] = value_tmp

        max_len = 0
        for key in output_data:
            if(len(output_data[key]) > max_len):
                max_len = len(output_data[key])
        if(max_len > len(file_c_lists[file_index].time_list)):
            file_c_lists[file_index].trouble = True
            print(file+' 有多的行可以合并, 需要手动操作')

        for key in output_data:
            for _ in range(max_len - len(output_data[key])):
                output_data[key].append('')

        df = pd.DataFrame.from_dict(output_data)
        df.to_excel(trans_fold_path+'\\'+file, index = False)
    
# merge
    merged_data = {'银行名':[], 'Bank_ID':[]}
    for inf in file_inf[1:]:
        merged_data[inf] = []

    all_col = []

    for file, file_index in zip(origin_files, range(len(origin_files))):
        if(file_c_lists[file_index].trouble):
            continue
        data_fy = pd.read_excel(os.path.join(trans_fold_path, file), header=0 ,keep_default_na=False)
        for col in data_fy.keys()[2:]:
            if col not in all_col:
                all_col.append(col)

    for col in all_col:
        merged_data[col] = []

    for file, file_index in zip(origin_files, range(len(origin_files))):
        data_fy = pd.read_excel(os.path.join(trans_fold_path, file), header=0 ,keep_default_na=False)
        if(file_c_lists[file_index].trouble):
            continue
        for row in range(data_fy.shape[0]):
            merged_data['银行名'].append(file_c_lists[file_index].name)

            merged_data['Bank_ID'].append(file_c_lists[file_index].inf[0])
            for inf_index in range(1, len(file_inf)):
                merged_data[file_inf[inf_index]].append(file_c_lists[file_index].inf[inf_index])

            for merged_key in merged_data.keys():
                if(merged_key in data_fy.keys() and merged_key != '银行名'):
                    merged_data[merged_key].append(data_fy[merged_key][row])
                elif(merged_key == '银行名' or merged_key == 'Bank_ID' or merged_key in file_inf):
                    pass
                else:
                    merged_data[merged_key].append('')

    df = pd.DataFrame.from_dict(merged_data)
    df.to_excel(trans_fold_path+'\\'+merged_file_name, index = False)


# rename
    data_fy = pd.read_excel(os.path.join(trans_fold_path, merged_file_name), header=None ,keep_default_na=False)
    for itm in items:
        if(itm+' 合计' in data_fy.iloc[0].to_list()):     
            data_fy.iloc[0][data_fy.iloc[0].tolist().index(itm+' 合计')] = '合计'
    data_fy.to_excel(trans_fold_path+'\\'+merged_file_name, index = False, header=False)
    
    # for file in origin_files:
    #     data_fy = pd.read_excel(os.path.join(trans_fold_path, file), header=None ,keep_default_na=False)
    #     for itm in items:
    #         if(itm+' 合计' in data_fy.iloc[0].to_list()):     
    #             data_fy.iloc[0][data_fy.iloc[0].tolist().index(itm+' 合计')] = '合计'
    #     data_fy.to_excel(trans_fold_path+'\\'+file, index = False, header=False)
