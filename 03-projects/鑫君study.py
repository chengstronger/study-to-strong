# import json  # 处理json文件
# import queue    # 创建一个新的线程，避免数据冲突
# import sys   # 系统的相关功能，在系统加载失败的时候强制关闭程序
# import sounddevice as sd   # 获取系统麦克风的使用权限，实时使用麦克风录制音频数据
# from vosk import Model,KaldiRecognizer  # 准备识别的环境，执行语言识别
#
# models = "D:\\浏览器下载\Edg\\vosk-model-small-cn-0.22"     #模型的文件路径
# Sample = 1600
# block_size = 8000 #每次处理的音频块的大小，8000是大概每次读取0.25秒的音频
# ears = 1    #这里可以修改是单声道还是双声道
#
# # --------加载模型——-------
# try:
#     model = Model(models)
# except Exception as e:
#     print(f"模型加载失败，请检查路径 '{models}' 是否正确。错误详情：{e}")
#     sys.exit(1)   # 退出程序
#
#
# # 创建语言识别器
# rec = KaldiRecognizer(model,Sample)   # 模型和音频的采样率
# rec.SetWords(False)  # 设置为 False 可以减少资源消耗，如果不需要单个词的时间戳
# # 若要获得单个词级别结果，可设置为 True: rec.SetWords(True)
#



import os
print(os.path.exists("D:\\浏览器下载\\Edg\\model"))












