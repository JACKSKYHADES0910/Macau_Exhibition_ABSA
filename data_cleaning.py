import pandas as pd
import re
import os

def clean_xhs_text(text):
    if pd.isna(text):
        return ""
    text = str(text)
    # 移除 URL
    text = re.sub(r'http[s]?://\S+', '', text)
    # 移除 @提及
    text = re.sub(r'@[^\s]+', '', text)
    # 处理末尾堆砌的话题标签：移除结尾处连串的 #话题
    text = re.sub(r'(#\S+\s*)+$', '', text)
    # 剥离句子中保留的话题的 # 号
    text = text.replace('#', '')
    # 移除不可见控制字符（保留 Emoji 等）
    text = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f-\x9f]', '', text)
    # 将连续换行替换为单换行
    text = re.sub(r'\n+', '\n', text)
    # 将连续空格替换为单空格
    text = re.sub(r' +', ' ', text)
    
    text = text.strip()
    # 强制截断在前 500 个字符
    return text[:500]

def main():
    file_path_excel = 'd:/Project/Macau_Sentiment_Analysis/艺术类ESG举措小红书UGC.xlsx'
    
    if os.path.exists(file_path_excel):
        print(f"Loading data from {file_path_excel}...")
        df = pd.read_excel(file_path_excel)
    else:
        print("Data file not found.")
        return

    print("Executing Data Preprocessing (XHS Deep Cleaning)...")
    if "笔记内容" not in df.columns:
        print("Error: '笔记内容' column not found.")
        return

    initial_len = len(df)
    df = df.dropna(how='all')
    df = df.dropna(subset=['笔记内容'])
    df = df.drop_duplicates(subset=['笔记内容']).copy()

    df['cleaned_text'] = df['笔记内容'].apply(clean_xhs_text)
    df = df[df['cleaned_text'] != ""]
    df = df.reset_index(drop=True)
    
    print(f"Preprocessing completed. Filtered from {initial_len} down to {len(df)} valid, deduplicated records.")
    
    output_file = 'd:/Project/Macau_Sentiment_Analysis/艺术类ESG举措小红书UGC_已清洗.xlsx'
    df.to_excel(output_file, index=False)
    print(f"Successfully saved cleaned data to {output_file}")

if __name__ == "__main__":
    main()
