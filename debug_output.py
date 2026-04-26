import os, re, json, time
os.environ['NO_PROXY'] = 'localhost,127.0.0.1'
os.environ['no_proxy'] = 'localhost,127.0.0.1'

from ollama import Client
import httpx

client = Client(
    host='http://127.0.0.1:11434',
    timeout=httpx.Timeout(120.0, connect=10.0)
)

# 测试 qwen3.5:9b，降低 num_gpu
for num_gpu in [40, 30, 20]:
    print(f"\n=== TEST: qwen3.5:9b (num_gpu={num_gpu}) ===")
    try:
        start = time.time()
        response = client.chat(
            model='qwen3.5:9b',
            messages=[
                {"role": "system", "content": "只返回JSON，不要解释。"},
                {"role": "user", "content": '分析: "澳门威尼斯人的蔡国强展太震撼了，强烈推荐大家去看！"\n返回: {"忠诚":{"情感极性":"提及或未提及","判定原因提取":"原因或无"}}'}
            ],
            options={'num_gpu': num_gpu, 'temperature': 0.1, 'num_predict': 512}
        )
        elapsed = time.time() - start
        raw = response['message']['content']
        cleaned = re.sub(r'<think>[\s\S]*?</think>', '', raw).strip()
        cleaned = re.sub(r'```json\s*', '', cleaned)
        cleaned = re.sub(r'```\s*', '', cleaned).strip()
        print(f"  Time: {elapsed:.1f}s, Length: {len(cleaned)}")
        print(f"  Output: {cleaned[:300]}")
        
        match = re.search(r'\{[\s\S]*\}', cleaned)
        if match:
            parsed = json.loads(match.group(0))
            print(f"  Parsed: {json.dumps(parsed, ensure_ascii=False)}")
        print("  SUCCESS!")
        break
    except Exception as e:
        print(f"  Error: {e}")
