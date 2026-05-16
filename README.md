# 液氧安全知识抽取工具

从 PDF/TXT 格式的安全标准规范中自动抽取液氧相关安全知识点，汇总成 Word 文档。

## 快速使用

```bash
pip install -r requirements.txt
python3 pipeline.py -i samples/ -o output/合集.docx -t 液氧 -m mock
```

## 模式

- **mock**: 关键词匹配（无需 LLM，开箱即用）
- **ollama**: 本地 LLM 智能抽取（需安装 Ollama）
