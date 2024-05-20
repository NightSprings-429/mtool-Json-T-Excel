from itertools import zip_longest
import json
import pandas as pd
import os
import re
import string

# 获取当前工作目录
current_directory = os.getcwd()

# 指定包含 Excel 文件的文件夹路径（当前目录）
folder_path = current_directory

# 读取原始 JSON 文件
with open("翻译文件.json", "r", encoding="utf-8") as file:
    original_data = json.load(file)

error_data = {}  # 存储处理错误的数据
processed_data = {}  # 存储处理后的数据

# 头字符匹配
custom_punctuation = set(string.punctuation)
custom_punctuation.update(string.ascii_letters)
custom_punctuation.update(string.digits)
custom_punctuation.update("【】「」『』（）…。、？！：；《》")

# 定义分隔符
separators = ['\n ','\n', '\\', '] ', ']', '-', ' ', '/']
separators.extend(string.punctuation)
separators.extend(string.digits)

for original_text, translated_text in original_data.items():
    # 替换全角空格为半角空格
    original_text = original_text.replace("　", " ")
    translated_text = translated_text.replace("　", " ")

    # 使用换行符分割文本
    paragraphs_original = re.split(f'({"|".join(map(re.escape, separators))})', original_text)
    paragraphs_translated = re.split(f'({"|".join(map(re.escape, separators))})', translated_text)

    # 如果原文比翻译短，将键值对添加到 error_data 中，同时将翻译装入第一个键值对
    if len(paragraphs_original) < len(paragraphs_translated):  
        error_data[original_text] = translated_text
        paragraphs_translated = [str(translated_text)]

    # 使用 zip_longest 补全较短的列表
    for item, trans in zip_longest(paragraphs_original, paragraphs_translated, fillvalue=None):
        # 更新 processed_data 字典，使用 strip() 移除首尾空白字符
        if item is not None:
            processed_data[item.strip()] = trans.strip() if trans is not None else None



with open("字典.json", "w", encoding="utf-8") as file:
    json.dump(processed_data, file, ensure_ascii=False, indent=2)

# 将错误数据写入 JSON 文件
with open("字典错误.json", "w", encoding="utf-8") as file:
    json.dump(error_data, file, ensure_ascii=False, indent=2)

trans_error_file = "翻译错误.txt"

# 获取文件夹中所有的文件名
file_names = [file for file in os.listdir(folder_path) if file.endswith('.xlsx')]


# 定义一个翻译函数
def translate_text(text_in):

    text = text_in.replace("　", " ")   # 全角空格转换半角
    text_parts = re.split(f'({"|".join(map(re.escape, separators))})', text)    # 分隔源文本

    pass_num = 0
    translated_parts = []
    for part in text_parts:
        contains_non_custom_punctuation = any(char not in custom_punctuation for char in part)
        if part in separators:  # 部分为换行符，跳过翻译
            translation = None 
            translated_parts.append(part)
            pass_num += 1
            continue
        else:
            translation = processed_data.get(part.strip(), None) # 翻译每个部分，忽略头尾空格
        if translation: # 正常翻译
            translated_parts.append(translation)
            pass_num += 1
            continue 
        if translation is None and not contains_non_custom_punctuation :  # 无翻译，头字符匹配英文、数字和符号
            translated_parts.append(part)
            pass_num += 1
            continue
        if translation is None and pass_num == 0: # 跳过翻译，输出错误信息
            trans_error[part] = "" 
            pass_num += 1
            continue
        
    if pass_num > 0 :
        pass
    else :
        trans_error[str(text_parts)] = "" # 输出错误信息


    # 合并翻译后的部分
    translated_text = ''.join(translated_parts)

    return translated_text

# 初始化错误字典
trans_error = {}

# 遍历每个文件并进行翻译
for file_name in file_names:
    # 读取 Excel 文件
    file_path = os.path.join(folder_path, file_name)
    df = pd.read_excel(file_path, header=None, converters={0: str})  # 假设没有表头

    # 对第一列进行翻译，将翻译结果写入第二列
    df[1] = df[0].apply(translate_text)

    # 将修改后的 DataFrame 写回 Excel 文件
    df.to_excel(file_path, index=False, header=False)

    print(f"翻译后的数据已保存到 Excel 文件: {file_path}")

# 将错误保存到文件
with open(trans_error_file, 'w', encoding='utf-8') as error_file:
    for key, value in trans_error.items():
        error_file.write(f"{key}: {value}\n")

# 删除文件中的重复项和空行
lines = set()
with open(trans_error_file, 'r', encoding='utf-8') as error_file:
    lines = error_file.readlines()

# 使用 filter 过滤掉空行
non_empty_lines = filter(lambda line: line.strip(), lines)

# 删除行末尾的冒号，保留换行符
modified_lines = [line.rstrip(':') for line in non_empty_lines]
end_lines = [line.replace(':', '') for line in modified_lines]

# 打印翻译错误的部分
with open(trans_error_file, 'w', encoding='utf-8') as error_file:
    error_file.writelines(end_lines)
   
print("字典错误的部分已保存到文件: 字典错误.json")
print("翻译错误的部分已保存到文件: 翻译错误.txt")
input("按任意键继续……")
