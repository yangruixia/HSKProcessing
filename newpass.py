import pandas as pd
import re

def process_text1(input_text):
    if isinstance(input_text, str):
        matches = re.findall(r'\[BQ(.*?)\]', input_text)
        
        for match in matches:
            input_text = input_text.replace(f'[BQ{match}]', match)
        
        input_text = re.sub(r'\[.*?\]', '', input_text)
        
        return input_text
    else:
        return input_text
def process_text2(input_text):
    # 使用正则表达式匹配花括号中的内容
    pattern = r'\{([^{}]+)\}'
    
    def replace(match):
        # 检查匹配到的内容中是否包含"CQ"或"CJ-"
        if 'CQ' in match.group(1) or 'CJ-' in match.group(1):
            # 如果包含"CQ"或"CJ-"，则保留中文字符，删除英文字符和花括号
            return re.sub(r'[^\u4e00-\u9fa5]+', '', match.group(1))
        else:
            # 如果不包含"CQ"或"CJ-"，则删除整个花括号中的内容
            return ''

    # 使用re.sub函数替换匹配到的内容
    result = re.sub(pattern, replace, input_text)
    
    return result

    # 使用re.sub函数替换匹配到的内容
    result = re.sub(pattern, replace, input_text)
    
    return result
df = pd.read_excel('original.xlsx', usecols='O', names=['O'])
df['O'] = df['O'].astype(str)
df['T'] = df['O'].replace(to_replace=r'\d+[^\d.]+\.\d+|\(.*?\)|\bTitle\b', value='', regex=True)
df['T'] = df['T'].astype(str)
df['S1'] = df['T'].apply(process_text1)
df['S1'] = df['S1'].astype(str)
df['S2'] = df['S1'].apply(process_text2)

with pd.ExcelWriter('output.xlsx') as writer:
    df.to_excel(writer, sheet_name='Sheet1', index=False, columns=['S2'])
