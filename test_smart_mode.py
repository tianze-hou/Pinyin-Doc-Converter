#!/usr/bin/env python3
"""
测试智能转换模式的效果
"""
import sys
import os

# 添加当前目录到Python路径
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from run import convert_text, TERM_MAPPING

# 测试智能转换模式
def test_smart_conversion():
    print("=== 智能转换模式测试 ===")
    
    # 测试单个词语
    test_words = ['计算机', '打印机', '互联网', '手机', '软件', '硬件', '电子邮件']
    
    print("\n1. 单个词语转换测试：")
    print(f"{'原词':<8} {'智能模式':<8} {'转换后字数':<10}")
    print("-" * 40)
    
    for word in test_words:
        result = convert_text(word, 1, smart_mode=True)
        print(f"{word:<8} {result:<8} {len(result):<10}")
        
        # 验证字数是否保持一致
        if len(result) != len(word):
            print(f"   WARNING: 字数发生变化！{len(word)} → {len(result)}")
    
    # 测试句子转换
    print("\n2. 句子转换测试：")
    test_sentences = [
        '计算机软件和硬件是现代科技的重要组成部分。',
        '中国人民正在努力发展经济和文化。',
        '电子邮件是互联网时代的重要通信工具。',
        '浏览器可以访问服务器上的各种资源。'
    ]
    
    for sentence in test_sentences:
        result = convert_text(sentence, 1, smart_mode=True)
        print(f"\n原句：{sentence}")
        print(f"转换：{result}")
        print(f"字数：{len(sentence)} → {len(result)}")
    
    # 测试智能模式的容错能力
    print("\n3. 智能模式容错能力测试：")
    # 假设有些词在转换后会改变字数（实际中可能不会发生，但测试代码逻辑）
    test_cases = [
        '这是一个测试句子，包含计算机和打印机等术语。',
        '中国的科技发展迅速，互联网应用广泛。',
        '软件工程师需要掌握硬件知识。'
    ]
    
    for text in test_cases:
        # 模拟OpenCC可能产生的字数变化情况
        # 注意：实际上s2hk配置通常不会改变字数
        print(f"\n处理文本：{text}")
        result = convert_text(text, 1, smart_mode=True)
        print(f"转换结果：{result}")
        print(f"字数保持：{'是' if len(result) == len(text) else '否'}")

# 测试术语映射表
def test_term_mapping():
    print("\n=== 术语映射表测试 ===")
    print(f"{'简体术语':<8} {'繁体术语':<8}")
    print("-" * 30)
    
    for term_s, term_t in TERM_MAPPING.items():
        print(f"{term_s:<8} {term_t:<8}")

if __name__ == "__main__":
    test_smart_conversion()
    test_term_mapping()