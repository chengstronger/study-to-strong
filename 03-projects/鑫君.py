# import json
# import queue
# import sys
# import sounddevice as sd
# from vosk import Model, KaldiRecognizer
# import pyttsx3
#
# # --- 配置参数 (请根据实际情况修改) ---
# MODEL_PATH = r"D:\python\model"  # 改成你的模型路径
# SAMPLE_RATE = 16000  # 采样率 (Hz)
# BLOCK_SIZE = 8000    # 每次处理的音频块大小 (帧)
# CHANNELS = 1         # 单声道
#
# # --- 初始化 Vosk ---
# # 1. 加载模型
# try:
#     model = Model(MODEL_PATH)
# except Exception as e:
#     print(f"模型加载失败，请检查路径 '{MODEL_PATH}' 是否正确。错误详情：{e}")
#     sys.exit(1)
#
# # 2. 创建识别器
# rec = KaldiRecognizer(model, SAMPLE_RATE)
# rec.SetWords(True)  # 设置为 False 可以减少资源消耗，如果不需要单个词的时间戳
# # 若要获得单个词级别结果，可设置为 True: rec.SetWords(True)
#
# # 3. 创建音频队列 (用于解耦音频采集与识别)
# q = queue.Queue()
#
# # --- 音频回调函数 (在新线程中被调用) ---
# def audio_callback(indata, frames, time, status):
#     # 音频处理函数，当有新的音频块时被调用
#     if status:
#         print(f"音频状态异常: {status}", file=sys.stderr)
#     q.put(bytes(indata))
#
# # --- 开始识别 ---
# try:
#     with sd.RawInputStream(samplerate=SAMPLE_RATE, blocksize=BLOCK_SIZE,
#                            device=None, dtype='int16',
#                            channels=CHANNELS, callback=audio_callback):
#         print("正在监听... (按 Ctrl+C 停止)")
#
#         while True:
#             # 从队列中获取音频数据
#             data = q.get()
#             # 使用识别器处理音频数据
#             if rec.AcceptWaveform(data):
#                 # 当识别到一个完整的句子时，返回结果
#                 result_json = json.loads(rec.Result())
#                 result_text = result_json.get('text', '')
#                 if result_text:
#                     print(f"识别结果: {result_text}")
#                     engine = pyttsx3.init()
#
#                     engine.say(result_text)
#                     engine.runAndWait()
#
#
#             else:
#                 # 识别中间结果（部分结果）
#                 partial_json = json.loads(rec.PartialResult())
#                 partial_text = partial_json.get('partial', '')
#                 if partial_text:
#                     # 可以打印，或用于实时反馈，如实时字幕等
#                     # print(f"中间结果: {partial_text}", end="\r")
#                     pass
#
# except KeyboardInterrupt:
#     print("\n程序已停止。")
# except Exception as e:
#     print(f"发生错误: {e}")

import json
import queue
import sys
import sounddevice as sd
from vosk import Model, KaldiRecognizer
import pyttsx3

# ========== 配置 ==========
MODEL_PATH = r"D:\python\model"   # 模型路径
SAMPLE_RATE = 16000
BLOCK_SIZE = 8000
CHANNELS = 1

# ========== 初始化 Vosk ==========
try:
    model = Model(MODEL_PATH)
except Exception as e:
    print(f"模型加载失败: {e}")
    sys.exit(1)

rec = KaldiRecognizer(model, SAMPLE_RATE)
rec.SetWords(True)

# ========== 初始化语音引擎（只初始化一次）==========
engine = pyttsx3.init()
engine.setProperty('rate', 150)   # 语速 150 词/分钟
engine.setProperty('volume', 0.9) # 音量 90%

def speak(text):
    """朗读文本并等待完成"""
    print(f"[助手] {text}")
    engine.say(text)
    engine.runAndWait()   # 阻塞直到播完

# ========== 音频队列 ==========
audio_queue = queue.Queue()

def audio_callback(indata, frames, time, status):
    if status:
        print(f"音频状态异常: {status}", file=sys.stderr)
    audio_queue.put(bytes(indata))

# ========== 主循环 ==========
print("正在监听... (按 Ctrl+C 停止)")
try:
    with sd.RawInputStream(samplerate=SAMPLE_RATE, blocksize=BLOCK_SIZE,
                           device=None, dtype='int16',
                           channels=CHANNELS, callback=audio_callback):
        while True:
            data = audio_queue.get()
            if rec.AcceptWaveform(data):
                # 完整句子识别结果
                result = json.loads(rec.Result())
                text = result.get('text', '').strip()
                if text:
                    print(f"[用户] {text}")
                    # 可选：如果用户说“退出”，结束程序
                    if text in ["退出", "再见", "拜拜"]:
                        speak("好的，再见")
                        break
                    # 回复用户（这里用回显，你也可以接入大模型）
                    reply = f"你说的是：{text}"
                    speak(reply)
                # 继续下一轮监听（循环自动继续）
            else:
                # 可选：打印实时中间结果（调试用）
                # partial = json.loads(rec.PartialResult()).get('partial', '')
                # if partial:
                #     print(f"\r[实时] {partial}", end='')
                pass
except KeyboardInterrupt:
    print("\n程序已停止。")
except Exception as e:
    print(f"发生错误: {e}")

