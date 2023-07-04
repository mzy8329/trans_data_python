import jieba
import math




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



a = '存款提现'
b = '提现_存款'

print(cosineSimilarity(a,b))