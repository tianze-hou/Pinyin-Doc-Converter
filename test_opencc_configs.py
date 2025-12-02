#!/usr/bin/env python3
"""
测试不同OpenCC配置的转换效果，特别是专业术语的处理
"""
from opencc import OpenCC

# 测试不同配置的转换效果
def test_opencc_configs():
    # 初始化不同配置的转换器
    configs = {
        's2t': OpenCC('s2t'),     # 简体到繁体（字符级）
        's2hk': OpenCC('s2hk'),   # 简体到中国香港繁体（词语级）
        's2tw': OpenCC('s2tw'),   # 简体到中国台湾繁体（词语级）
        's2twp': OpenCC('s2twp'), # 简体到中国台湾繁体（词语级，带注音符号）
    }
    
    # 测试词汇，特别关注IT术语
    test_words = [
        # IT术语
        '计算机', '打印机', '互联网', '手机', '软件', '硬件', 
        '电子邮件', '浏览器', '服务器', '客户端', '数据库',
        # 普通词汇
        '中国', '人民', '科技', '发展', '经济', '文化'
    ]
    
    print("=== 不同OpenCC配置的转换效果 ===")
    print(f"{'原词':<8} {'s2t':<8} {'s2hk':<8} {'s2tw':<12} {'s2twp':<16}")
    print("-" * 60)
    
    for word in test_words:
        results = [cc.convert(word) for cc in configs.values()]
        print(f"{word:<8} {' '.join(f'{r:<8}' for r in results)}")
    
    # 测试完整句子
    print("\n=== 完整句子转换测试 ===")
    test_sentence = "计算机软件和硬件是现代科技的重要组成部分。"
    print(f"原句：{test_sentence}")
    for name, cc in configs.items():
        print(f"{name:<6}: {cc.convert(test_sentence)}")

# 测试词语级转换的上下文影响
def test_context_effects():
    cc_s2hk = OpenCC('s2hk')
    cc_s2t = OpenCC('s2t')
    
    print("\n=== 上下文对转换的影响测试 ===")
    
    # 单个词语测试
    print("\n单个词语转换：")
    words = ['软件', '软体', '计算机', '电脑']
    for word in words:
        print(f"{word:<8} → s2hk: {cc_s2hk.convert(word):<8}, s2t: {cc_s2t.convert(word):<8}")
    
    # 不同上下文中的相同词语
    print("\n不同上下文中的相同词语：")
    phrases = [
        '开发软件',
        '这个软件很好用',
        '软件工程',
        '软件公司'
    ]
    for phrase in phrases:
        print(f"{phrase:<12} → s2hk: {cc_s2hk.convert(phrase):<16}, s2t: {cc_s2t.convert(phrase):<16}")

if __name__ == "__main__":
    test_opencc_configs()
    test_context_effects()