import os
import time
from ollama import Client

# 【终极修复1】屏蔽所有系统代理和加速器的干扰
os.environ['HTTP_PROXY'] = ''
os.environ['HTTPS_PROXY'] = ''
os.environ['NO_PROXY'] = 'localhost,127.0.0.1'

client = Client(host='http://127.0.0.1:11434')

my_question = "你好，请用一句话告诉我，澳门有什么好玩的？"
print(f"正在向大模型提问: {my_question}")
print("正在等待大模型思考，请稍候...\n")

start_time = time.time()

# 【终极修复2】开启 stream=True（流式输出）
response = client.chat(
    model='qwen3.5:4b',  # 你的主力模型
    messages=[{'role': 'user', 'content': my_question}],
    keep_alive='1h',
    stream=True  # 像打字机一样实时输出
)

print("--- 大模型的回答 ---")
first_token_time = None

# 挨个字抓取大模型的输出
for chunk in response:
    if first_token_time is None:
        first_token_time = time.time()
        print(f"\n[雷达报告：收到第一个字！前置等待耗时: {first_token_time - start_time:.2f} 秒]\n")
    
    # 实时打印每个字到屏幕上
    print(chunk['message']['content'], end='', flush=True)

end_time = time.time()
print(f"\n\n[回答完毕！本次总耗时: {end_time - start_time:.2f} 秒]")