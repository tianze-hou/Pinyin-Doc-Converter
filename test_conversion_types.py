#!/usr/bin/env python3
"""
测试简繁转换的字符级和词语级效果
"""
from opencc import OpenCC

# 测试转换效果
def test_conversion_effects():
    # 初始化不同类型的转换器
    cc_s2hk_word = OpenCC('s2hk')  # 词语级简转繁
    cc_s2hk_char = OpenCC('s2t')   # 字符级简转繁
    cc_t2s = OpenCC('t2s')         # 繁转简（同时支持字符级和词语级）
    
    # 测试词汇
    test_words = ['计算机', '打印机', '互联网', '手机', '软件', '硬件', '电子邮件']
    
    print("=== 转换效果对比 ===")
    print(f"{'原词':<8} {'词语级转换':<10} {'字符级转换':<10} {'繁转简':<8}")
    print("-" * 50)
    
    for word in test_words:
        word_conv = cc_s2hk_word.convert(word)
        char_conv = cc_s2hk_char.convert(word)
        back_conv = cc_t2s.convert(word_conv)
        print(f"{word:<8} {word_conv:<10} {char_conv:<10} {back_conv:<8}")
    
    print("\n=== 完整句子转换对比 ===")
    test_sentence = "计算机是现代科技的重要发明，打印机可以将数字内容打印出来。"
    print(f"原句：{test_sentence}")
    print(f"词语级：{cc_s2hk_word.convert(test_sentence)}")
    print(f"字符级：{cc_s2hk_char.convert(test_sentence)}")

if __name__ == "__main__":
    test_conversion_effects()