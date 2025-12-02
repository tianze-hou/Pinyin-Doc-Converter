#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试不同地区(中国香港/中国台湾)的繁体转换效果
"""

import sys
import os

# 添加当前目录到Python路径
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from run import convert_text

def test_region_specific_terms():
    """测试特定术语在不同地区的转换效果"""
    print("=== 测试特定术语在不同地区的转换效果 ===")
    
    test_terms = [
        "计算机",
        "互联网",
        "软件",
        "硬件",
        "服务器",
        "客户端",
        "数据库",
        "电子邮件"
    ]
    
    print("术语\t中国香港繁体\t中国台湾繁体")
    print("-" * 50)
    
    for term in test_terms:
        hk_result = convert_text(term, 1, smart_mode=True, region="hk")
        tw_result = convert_text(term, 1, smart_mode=True, region="tw")
        print(f"{term}\t{hk_result}\t{tw_result}")

def test_complete_sentences():
    """测试完整句子在不同地区的转换效果"""
    print("\n=== 测试完整句子在不同地区的转换效果 ===")
    
    test_sentences = [
        "计算机软件和硬件是现代科技的重要组成部分。",
        "互联网改变了人们的生活方式。",
        "服务器是网络中的重要设备。",
        "中国的科技发展非常迅速。",
        "用户可以通过客户端访问数据库。"
    ]
    
    for i, sentence in enumerate(test_sentences, 1):
        hk_result = convert_text(sentence, 1, smart_mode=True, region="hk")
        tw_result = convert_text(sentence, 1, smart_mode=True, region="tw")
        
        print(f"\n句子 {i}:")
        print(f"简体:    {sentence}")
        print(f"中国香港繁体: {hk_result}")
        print(f"中国台湾繁体: {tw_result}")
        
        # 检查字数是否一致
        if len(sentence) == len(hk_result) == len(tw_result):
            print(f"✓ 字数一致: {len(sentence)} 个字符")
        else:
            print(f"✗ 字数不一致: 简体({len(sentence)}) → 中国香港繁体({len(hk_result)}) → 中国台湾繁体({len(tw_result)})")

def test_smart_mode_vs_regular():
    """测试智能模式在不同地区的效果"""
    print("\n=== 测试智能模式在不同地区的效果 ===")
    
    test_text = "软件是计算机的重要组成部分。"
    
    # 中国香港繁体测试
    hk_smart = convert_text(test_text, 1, smart_mode=True, region="hk")
    hk_regular = convert_text(test_text, 1, smart_mode=False, region="hk")
    
    # 中国台湾繁体测试
    tw_smart = convert_text(test_text, 1, smart_mode=True, region="tw")
    tw_regular = convert_text(test_text, 1, smart_mode=False, region="tw")
    
    print(f"测试文本: {test_text}")
    print("\n中国香港繁体:")
    print(f"  智能模式: {hk_smart}")
    print(f"  普通模式: {hk_regular}")
    
    print("\n中国台湾繁体:")
    print(f"  智能模式: {tw_smart}")
    print(f"  普通模式: {tw_regular}")

if __name__ == "__main__":
    test_region_specific_terms()
    test_complete_sentences()
    test_smart_mode_vs_regular()
