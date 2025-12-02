import os
import logging
import yaml
from docx import Document
from opencc import OpenCC

logging.basicConfig(filename='conversion.log', level=logging.INFO,
                    format='%(asctime)s - %(levelname)s - %(message)s')

def load_config(config_file):
    try:
        with open(config_file, 'r', encoding='utf-8') as file:
            config = yaml.safe_load(file)
        return config
    except Exception as e:
        logging.error(f"Failed to load config file: {e}")
        raise

# 初始化OpenCC转换器实例（仅创建一次以提高效率）
cc_s2hk = OpenCC('s2hk')      # 简体到中国香港繁体（词语级）
cc_s2twp = OpenCC('s2twp')    # 简体到中国台湾繁体（词语级，更全面的术语转换）
cc_t2s = OpenCC('t2s')        # 繁体到简体（同时支持字符级和词语级）

# 术语映射表，用于手动指定特定术语的转换（解决拼音匹配问题）
# 键：地区(hk/tw) -> 简体术语 -> 繁体术语
TERM_MAPPING = {
    'hk': {  # 中国香港繁体术语映射
        # IT术语
        '计算机': '計算機',
        '打印机': '打印機',
        '互联网': '互聯網',
        '手机': '手機',
        '软件': '軟件',  # 中国香港用"軟件"
        '硬件': '硬件',
        '电子邮件': '電子郵件',
        '浏览器': '瀏覽器',
        '服务器': '服務器',  # 中国香港用"服務器"
        '客户端': '客戶端',
        '数据库': '數據庫',  # 中国香港用"數據庫"
        # 其他常用术语
        '中国': '中國',
        '人民': '人民',
        '科技': '科技',
        '发展': '發展',
        '经济': '經濟',
        '文化': '文化'
    },
    'tw': {  # 中国台湾繁体术语映射
        # IT术语
        '计算机': '計算機',
        '打印机': '打印機',
        '互联网': '網際網路',  # 中国台湾用"網際網路"
        '手机': '手機',
        '软件': '軟體',  # 中国台湾用"軟體"
        '硬件': '硬體',
        '电子邮件': '電子郵件',
        '浏览器': '瀏覽器',
        '服务器': '伺服器',  # 中国台湾用"伺服器"
        '客户端': '用戶端',  # 中国台湾用"用戶端"
        '数据库': '資料庫',  # 中国台湾用"資料庫"
        # 其他常用术语
        '中国': '中國',
        '人民': '人民',
        '科技': '科技',
        '发展': '發展',
        '经济': '經濟',
        '文化': '文化'
    }
}

def convert_text(text, conversion_type, smart_mode=True, region="hk"):
    """智能文本转换
    
    Args:
        text: 要转换的文本
        conversion_type: 转换类型（1: 简转繁, 2: 繁转简）
        smart_mode: 是否启用智能模式
            True: 使用词语级转换，但保持拼音文本不变，并且使用术语映射表确保转换后字数一致
            False: 根据配置使用字符级或词语级转换
        region: 地区差异选项（仅在conversion_type=1时生效）
            hk: 中国香港繁体
            tw: 中国台湾繁体
    """
    if conversion_type == 2:  # 繁转简
        return cc_t2s.convert(text)
    
    # 简转繁，根据地区选择转换器
    selected_converter = cc_s2hk if region == "hk" else cc_s2twp
    
    if not smart_mode:
        return selected_converter.convert(text)
    
    # 智能模式：使用词语级转换，但确保转换后字数一致
    result = text
    
    # 先使用OpenCC进行词语级转换
    opencc_result = selected_converter.convert(text)
    
    # 如果转换后字数发生变化，尝试使用术语映射表进行精确转换
    if len(opencc_result) != len(text):
        # 简单的词语替换实现
        # 注意：这是一个简化实现，实际应用中可能需要更复杂的分词和替换逻辑
        for term_s, term_t in TERM_MAPPING[region].items():
            if term_s in text:
                result = result.replace(term_s, term_t)
        
        # 对剩余字符进行字符级转换
        for char in text:
            if char in result:  # 只转换未被术语替换的字符
                # 根据地区选择不同的转换方式
                if region == "hk":
                    conv_char = cc_s2hk.convert(char)  # 转换为中国香港繁体
                else:  # tw
                    conv_char = cc_s2twp.convert(char)  # 转换为中国台湾繁体
                if conv_char != char:
                    result = result.replace(char, conv_char)
    else:
        # 如果字数未变化，直接使用OpenCC的转换结果
        result = opencc_result
    
    return result

def process_run(run, conversion_type, smart_mode=True, region="hk"):
    ruby_xml = run._r
    if ruby_xml is not None:
        for elem in ruby_xml.iter():
            # 检查是否为文本元素
            if elem.tag.endswith('t') and elem.text:
                # 检查该文本元素是否属于拼音文本(rubyText)
                is_ruby_text = False
                parent = elem.getparent()
                while parent is not None:
                    if parent.tag.endswith('rubyText'):
                        is_ruby_text = True
                        break
                    parent = parent.getparent()
                
                # 只转换基文本，不转换拼音文本
                if not is_ruby_text:
                    elem.text = convert_text(elem.text, conversion_type, smart_mode, region)

def main():
    try:
        config = load_config('config.yaml')
        input_file = config.get("input_file")
        output_file = config.get("output_file", "")
        conversion_type = config.get("conversion_type")
        smart_mode = config.get("smart_mode", True)  # 默认启用智能模式
        region = config.get("region", "hk")  # 默认使用中国香港繁体

        if not input_file or conversion_type not in (1, 2):
            logging.error("Invalid configuration: input_file or conversion_type is missing or invalid")
            return
        
        # Generate default output file path if not specified
        if not output_file:
            base, ext = os.path.splitext(input_file)
            if conversion_type == 1:
                if region == "tw":
                    output_file = f"{base}-TW{ext}"
                else:  # 默认或hk
                    output_file = f"{base}-HK{ext}"
            elif conversion_type == 2:
                output_file = f"{base}-SC{ext}"

        logging.info(f"Starting conversion: {input_file} to {output_file}, type: {conversion_type}, smart_mode: {smart_mode}, region: {region}")

        doc = Document(input_file)

        for p in doc.paragraphs:
            for run in p.runs:
                process_run(run, conversion_type, smart_mode, region)

        doc.save(output_file)
        logging.info("Conversion completed successfully")

    except Exception as e:
        logging.error(f"An error occurred: {e}")

if __name__ == "__main__":
    main()
