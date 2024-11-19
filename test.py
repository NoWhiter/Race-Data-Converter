with open("Form1.cs", "rb") as f:
    raw_data = f.read()

try:
    # 强制用 UTF-8 解码
    content = raw_data.decode("utf-8")
    print(content[:300])
except UnicodeDecodeError as e:
    print(f"UTF-8 解码失败: {e}")
