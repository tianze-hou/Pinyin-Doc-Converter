import os
import logging
import yaml
from docx import Document
import zhconv

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

def convert_to_hk_traditional(text):
    return zhconv.convert(text, 'zh-hant')

def convert_to_simplified(text):
    return zhconv.convert(text, 'zh-hans')

def process_run(run, conversion_type):
    ruby_xml = run._r
    if ruby_xml is not None:
        for elem in ruby_xml.iter():
            if elem.tag.endswith('t') and elem.text:
                if conversion_type == 1:
                    elem.text = convert_to_hk_traditional(elem.text)
                elif conversion_type == 2:
                    elem.text = convert_to_simplified(elem.text)

def main():
    try:
        config = load_config('config.yaml')
        input_file = config.get("input_file")
        output_file = config.get("output_file", "")
        conversion_type = config.get("conversion_type")

        if not input_file or conversion_type not in (1, 2):
            logging.error("Invalid configuration: input_file or conversion_type is missing or invalid")
            return
        
        # Generate default output file path if not specified
        if not output_file:
            base, ext = os.path.splitext(input_file)
            if conversion_type == 1:
                output_file = f"{base}-TC{ext}"
            elif conversion_type == 2:
                output_file = f"{base}-SC{ext}"

        logging.info(f"Starting conversion: {input_file} to {output_file}, type: {conversion_type}")

        doc = Document(input_file)

        for p in doc.paragraphs:
            for run in p.runs:
                process_run(run, conversion_type)

        doc.save(output_file)
        logging.info("Conversion completed successfully")

    except Exception as e:
        logging.error(f"An error occurred: {e}")

if __name__ == "__main__":
    main()
