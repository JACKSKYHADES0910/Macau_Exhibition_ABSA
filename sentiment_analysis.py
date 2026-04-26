import pandas as pd
import json
import re
import os
import time
from tqdm import tqdm

# 忽略本地代理，防止由于系统开启代理软件(如Clash)导致的 502 Bad Gateway 错误
os.environ['NO_PROXY'] = 'localhost,127.0.0.1'
os.environ['no_proxy'] = 'localhost,127.0.0.1'

from ollama import Client
import httpx

def main():
    # ==========================
    # 运行配置区
    # ==========================
    TEST_MODE = False   # True=测试模式, False=全量
    TEST_LIMIT = 5
    USE_CHECKPOINT = False  # 关闭断点续传，全量重跑（修复后）
    MODEL = 'qwen2.5:7b'   # qwen3.5:4b 在当前环境返回空值，改用 qwen2.5:7b
    # ==========================

    print(f"Initializing Ollama Client (Model: {MODEL})...")
    client = Client(
        host='http://127.0.0.1:11434',
        timeout=httpx.Timeout(120.0, connect=10.0)
    )

    # 读取已经清洗完毕的数据集
    file_path_excel = 'd:/Project/Macau_Sentiment_Analysis/艺术类ESG举措小红书UGC_已清洗.xlsx'
    
    if os.path.exists(file_path_excel):
        print(f"Loading cleaned data from {file_path_excel}...")
        df = pd.read_excel(file_path_excel)
    else:
        print(f"Data file not found: {file_path_excel}\nPlease run data_cleaning.py first.")
        return

    if "cleaned_text" not in df.columns:
        print("Error: 'cleaned_text' column not found.")
        return
    
    df = df[df['cleaned_text'].notna() & (df['cleaned_text'] != "")]
    df = df.reset_index(drop=True)
    print(f"Loaded {len(df)} valid records.")

    if TEST_MODE:
        print(f"\n[!] 测试模式: 只处理前 {TEST_LIMIT} 条 [!]\n")
        df = df.head(TEST_LIMIT).copy()

    checkpoint_file = 'd:/Project/Macau_Sentiment_Analysis/澳门展览_情感分析_中间态_v2.csv'
    
    processed_indices = set()
    results_map = {}
    
    # 七个维度 (一字不变)
    dimensions = [
        "文化认知价值",
        "体验审美价值",
        "伦理公共价值",
        "象征价值",
        "感知文化真实性",
        "忠诚",
        "声誉"
    ]

    if USE_CHECKPOINT and os.path.exists(checkpoint_file):
        try:
            chk_df = pd.read_csv(checkpoint_file, encoding='utf-8-sig')
            for _, row in chk_df.iterrows():
                if 'original_index' in row:
                    idx_str = str(row['original_index'])
                    processed_indices.add(idx_str)
                    dim_dict = {}
                    for dim in dimensions:
                        dim_dict[dim] = {
                            "情感极性": row.get(f"{dim}_情感极性", "未提及"),
                            "判定原因提取": row.get(f"{dim}_判定原因提取", "无")
                        }
                    results_map[idx_str] = dim_dict
            print(f"Loaded {len(processed_indices)} records from checkpoint.")
        except Exception as e:
            print(f"Checkpoint load failed, starting fresh. Error: {e}")

    # 优化后的 prompt：放宽"伦理公共价值""忠诚""声誉"的判定标准，捕获隐性表达
    system_prompt = """你是数据分析助手。请分析小红书笔记，判断是否提及以下7个维度。
请注意：用户表达往往是隐性的、口语化的，请从宽判定，只要文本中有相关暗示即判为"提及"。

维度定义（请从宽理解）：
1. 文化认知价值：提到澳门文化、本地历史、艺术语境、社会意义、文化学习等。
2. 体验审美价值：提到参观愉悦、好看、震撼、沉浸感、美学享受、视觉效果等。
3. 伦理公共价值：提到免费开放、公益、社会责任、文化普及、教育意义、造福公众、值得推广等。也包括提到"免费""免门票""公共开放""公众可参观"等表述。
4. 象征价值：提到度假村/品牌的文化形象提升、品味、高端定位、文化地标等。
5. 感知文化真实性：提到度假村不仅是博彩场所，也有真正的文化内容、艺术氛围、不像赌场等。
6. 忠诚：提到想再去、下次还来、推荐朋友去、值得一去、必去、打卡推荐、安利、种草等。只要有任何推荐或再访意愿的暗示，即判为提及。
7. 声誉：提到度假村/品牌整体评价、口碑、名气、知名度、大牌、高级感、国际化等。包括任何对品牌形象的正面或负面评价。

情感极性只能填"提及"或"未提及"。判定原因提取填原因摘要，未提及则填"无"。
仅返回JSON：{"文化认知价值":{"情感极性":"...","判定原因提取":"..."},"体验审美价值":{"情感极性":"...","判定原因提取":"..."},"伦理公共价值":{"情感极性":"...","判定原因提取":"..."},"象征价值":{"情感极性":"...","判定原因提取":"..."},"感知文化真实性":{"情感极性":"...","判定原因提取":"..."},"忠诚":{"情感极性":"...","判定原因提取":"..."},"声誉":{"情感极性":"...","判定原因提取":"..."}}"""

    for dim in dimensions:
        df[f"{dim}_情感极性"] = "未提及"
        df[f"{dim}_判定原因提取"] = "无"

    success_count = 0
    fail_count = 0

    for idx, row in tqdm(df.iterrows(), total=len(df), desc="Analyzing Sentiment"):
        idx_str = str(idx)
        if idx_str in processed_indices:
            res = results_map[idx_str]
            for dim in dimensions:
                if isinstance(res.get(dim), dict):
                    df.at[idx, f"{dim}_情感极性"] = res[dim].get("情感极性", "未提及")
                    df.at[idx, f"{dim}_判定原因提取"] = res[dim].get("判定原因提取", "无")
            continue

        content = str(row['cleaned_text'])[:400]
        messages = [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": f"笔记内容：{content}"}
        ]
        
        parsed_json = None
        max_retries = 3
        for attempt in range(max_retries):
            try:
                start_time = time.time()
                response = client.chat(
                    model=MODEL,
                    messages=messages,
                    options={
                        'num_gpu': 99,
                        'temperature': 0.1,
                        'num_predict': 512,
                    }
                )
                elapsed = time.time() - start_time
                
                result_text = response['message']['content'].strip()
                
                # 清理可能的 markdown 代码块
                result_text = re.sub(r'```json\s*', '', result_text)
                result_text = re.sub(r'```\s*', '', result_text).strip()
                
                # 修复中文模型常见的标点问题：中文标点→英文标点
                result_text = result_text.replace('：', ':').replace('，', ',')
                result_text = result_text.replace('“', '"').replace('”', '"')
                result_text = result_text.replace('‘', "'").replace('’', "'")
                result_text = result_text.replace('；', ';').replace('。', '.')
                
                # 提取 JSON 对象
                match = re.search(r'\{[\s\S]*\}', result_text)
                if not match:
                    raise ValueError(f"No JSON found in: {result_text[:80]}...")
                json_str = match.group(0)
                
                # 尝试解析，如果失败则自动修复常见格式错误
                try:
                    parsed_json = json.loads(json_str)
                except json.JSONDecodeError:
                    # 模型常见错误：省略了 "情感极性" key，直接写 {"未提及", ...}
                    # 修复：把 {"提及" 或 {"未提及" 补全为 {"情感极性":"提及" 或 {"情感极性":"未提及"
                    fixed = re.sub(
                        r'\{\s*"(提及|未提及)"\s*,',
                        r'{"情感极性":"\1",',
                        json_str
                    )
                    # 另一个常见错误：省略了 "判定原因提取" key
                    fixed = re.sub(
                        r',\s*"(无|[^"]*?)"\s*\}',
                        r',"判定原因提取":"\1"}',
                        fixed
                    )
                    parsed_json = json.loads(fixed)
                
                tqdm.write(f"  [#{idx+1}] OK ({elapsed:.1f}s)")
                break
            except Exception as e:
                tqdm.write(f"  [#{idx+1}] Attempt {attempt+1} failed: {e}")
                if attempt < max_retries - 1:
                    time.sleep(1)
                    
        if parsed_json:
            for dim in dimensions:
                dim_val = parsed_json.get(dim, {})
                if isinstance(dim_val, dict):
                    df.at[idx, f"{dim}_情感极性"] = dim_val.get("情感极性", "未提及")
                    df.at[idx, f"{dim}_判定原因提取"] = dim_val.get("判定原因提取", "无")
                else:
                    df.at[idx, f"{dim}_情感极性"] = "未提及"
                    df.at[idx, f"{dim}_判定原因提取"] = "无"
            results_map[idx_str] = parsed_json
            success_count += 1
        else:
            for dim in dimensions:
                df.at[idx, f"{dim}_情感极性"] = "解析失败"
                df.at[idx, f"{dim}_判定原因提取"] = "解析失败"
            failed_dict = {dim: {"情感极性": "解析失败", "判定原因提取": "解析失败"} for dim in dimensions}
            results_map[idx_str] = failed_dict
            fail_count += 1
            
        processed_indices.add(idx_str)
        
        if len(processed_indices) % 10 == 0:
            try:
                processed_df = df.iloc[:idx+1].copy()
                processed_df['original_index'] = processed_df.index
                processed_df.to_csv(checkpoint_file, index=False, encoding='utf-8-sig')
            except PermissionError:
                tqdm.write("  [!] Checkpoint file locked (close Excel?), skipping save...")

    print(f"\nDone! Success: {success_count}, Failed: {fail_count}")
    df['original_index'] = df.index
    try:
        df.to_csv(checkpoint_file, index=False, encoding='utf-8-sig')
    except PermissionError:
        alt = checkpoint_file.replace('.csv', '_backup.csv')
        df.to_csv(alt, index=False, encoding='utf-8-sig')
        print(f"[!] Checkpoint locked, saved to {alt}")
    
    output_file = 'd:/Project/Macau_Sentiment_Analysis/澳门展览_清洗分析版_v2.xlsx'
    try:
        df.to_excel(output_file, index=False)
        print(f"Saved to {output_file}")
    except PermissionError:
        alt = output_file.replace('.xlsx', '_backup.xlsx')
        df.to_excel(alt, index=False)
        print(f"[!] Output locked, saved to {alt}")

if __name__ == "__main__":
    main()
