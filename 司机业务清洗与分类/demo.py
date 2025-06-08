import sys
print("当前Python解释器路径：", sys.executable)
try:
    import openai
    print("openai 版本：", openai.__version__)
except ImportError:
    print("没有 openai 包")
