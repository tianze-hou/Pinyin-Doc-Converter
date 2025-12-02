# Pinyin-Doc-Converter：拼音Word文档简繁转换

这是一个能够将「带有拼音指南的 Word 文档」进行简繁转换的 Python 工具。

## 功能

- 处理带有「拼音指南」的Word文档（`.doc`, `.docx`），实现简繁互转。
 
## 缘起：闲言碎语

在之前的一份工作中，甲方要求提供带注音的简繁文稿。在处理文稿的过程中我们发现，Word 自带的「拼音指南」功能体验极差，带有拼音指南的文本无法在 Word 界面中连续选中，因而不能实现整篇的简繁转换。为了解决这个难题，我掘地三尺查到了这个 Ruby 类的实现，进而写下了这个代码。

~~目前版本的简繁转换是逐字转换，部分多对一简化字可能存在错误，转换结果仅供参考。~~
2025.11.25 摸鱼的时候又捡起这个项目，使用`opencc-python-reimplemented`库取代原来的`zhconv`库，实现词语级别的转换。

「拼音指南」这个功能在汉语和日语中并不是一个低频需求，许多学习者和像我的甲方一样普通话不好的人都会需要使用。然而印象中仿佛小时候用的 Office 2003 就在用这个方式实现，体验完全没有任何改变。

**谨对微软长期以来不尽人意的实现表示遗憾与担忧**。



## 环境要求

- Python 3.x
- 依赖库：
  - `python-docx`
  - `pyyaml`
  - `opencc-python-reimplemented`

## 安装

使用以下命令安装依赖库：

```bash
pip install python-docx pyyaml opencc-python-reimplemented
```

## 配置

使用 `config.yaml` 定义输入文件路径、输出文件路径和转换类型。示例如下：

```yaml
# 输入文件路径
input_file: '/path/to/your/input.docx'

# 输出文件路径，留空则默认保存到同一路径
output_file: ''  

# 转换类型：1为简转繁，2为繁转简
conversion_type: 1
```

## 使用方法

运行 `run.py` 文件以开始转换：

```bash
python run.py
```

确保 `config.yaml` 文件与 `run.py` 在同一目录下。

## Ruby 类介绍

- Ruby（ルビ）是 Microsoft Word 中为东亚语言文字提供的拼音指南功能。
- 该工具中使用了 Microsoft 的 Ruby 类（定义在 `DocumentFormat.OpenXml.Wordprocessing` 命名空间），用于处理拼音指南。Ruby 类继承自 `DocumentFormat.OpenXml.OpenXmlCompositeElement`，主要功能包括：

- **子元素**：包括 `rt`（拼音指南文本）和 `rubyBase`（拼音指南基文本）
- **构造函数**：支持多种方式初始化，包括从外部 XML 初始化
- **属性和方法**：提供丰富的属性和方法以便于操作 Ruby 元素

更多信息请参考 [Microsoft 官方文档](https://learn.microsoft.com/zh-cn/dotnet/api/documentformat.openxml.wordprocessing.ruby)。

## TODO

### 性能优化

- [ ] 支持批量处理多个文件或目录
- [ ] 实现并行处理，充分利用多核CPU资源
- [ ] 优化内存使用，支持处理大型文档
- [ ] 提升转换效率，考虑批量转换和更高效的转换库

### 功能优化

- [x] 优化转换精度，实现词语级别的转换
- [x] 扩展转换选项，支持地区差异转换（大陆简体、中国台湾繁体、中国香港繁体）
- [ ] 增强拼音处理，确保转换后拼音与文字正确对应
- [ ] 支持更多文件格式（如.doc、PDF、TXT等）
- [ ] 添加GUI界面，提升易用性
- [ ] 添加转换预览功能和转换历史记录
- [ ] 增强错误处理，提供更详细的错误信息
