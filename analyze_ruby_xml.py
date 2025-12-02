#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
分析带有拼音指南的Word文档的XML结构，了解拼音与文字的对应关系
"""

from docx import Document
import xml.etree.ElementTree as ET

# 创建一个简单的带有拼音指南的测试文档
def create_test_doc():
    doc = Document()
    p = doc.add_paragraph()
    run = p.add_run()
    
    # 直接操作XML添加拼音指南
    run._r.append(ET.Element(ET.QName('http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'ruby')))
    ruby = run._r[0]
    
    # 添加rubyPr
    rubyPr = ET.SubElement(ruby, ET.QName('http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'rubyPr'))
    rubyAlign = ET.SubElement(rubyPr, ET.QName('http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'rubyAlign'), {'w:val': 'center'})
    hps = ET.SubElement(rubyPr, ET.QName('http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'hps'), {'w:rubyBase': '21', 'w:rubyText': '11'})
    lid = ET.SubElement(rubyPr, ET.QName('http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'lid'), {'w:val': 'zh-CN'})
    
    # 添加rubyBase (基文本)
    rubyBase = ET.SubElement(ruby, ET.QName('http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'rubyBase'))
    r_base = ET.SubElement(rubyBase, ET.QName('http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'r'))
    t_base = ET.SubElement(r_base, ET.QName('http://schemas.openxmlformats.org/wordprocessingml/2006/main', 't'))
    t_base.text = '计算机'
    
    # 添加rubyText (拼音)
    rubyText = ET.SubElement(ruby, ET.QName('http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'rubyText'))
    r_text = ET.SubElement(rubyText, ET.QName('http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'r'))
    t_text = ET.SubElement(r_text, ET.QName('http://schemas.openxmlformats.org/wordprocessingml/2006/main', 't'))
    t_text.text = 'Jìsuànjī'
    
    doc.save('test_ruby_doc.docx')
    print("测试文档创建完成: test_ruby_doc.docx")

# 分析文档的XML结构
def analyze_doc(doc_path):
    print(f"\n=== 分析文档: {doc_path} ===")
    doc = Document(doc_path)
    
    for i, p in enumerate(doc.paragraphs):
        print(f"\n段落 {i+1}:")
        for j, run in enumerate(p.runs):
            print(f"  Run {j+1}:")
            print(f"    文本内容: '{run.text}'")
            
            # 检查是否有拼音指南
            if hasattr(run, '_r') and run._r is not None:
                ruby_xml = run._r
                print("    包含拼音指南:")
                
                # 查找ruby元素
                for elem in ruby_xml.iter():
                    if elem.tag.endswith('ruby'):
                        print("      找到ruby元素")
                        
                        # 查找rubyBase和rubyText
                        for child in elem:
                            if child.tag.endswith('rubyBase'):
                                print("        rubyBase (基文本):")
                                for text_elem in child.iter():
                                    if text_elem.tag.endswith('t') and text_elem.text:
                                        print(f"          文本: '{text_elem.text}'")
                            elif child.tag.endswith('rubyText'):
                                print("        rubyText (拼音):")
                                for text_elem in child.iter():
                                    if text_elem.tag.endswith('t') and text_elem.text:
                                        print(f"          文本: '{text_elem.text}'")

if __name__ == "__main__":
    create_test_doc()
    analyze_doc('test_ruby_doc.docx')
