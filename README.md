# 🧩 Excel 图片压缩工具
本人能力不足全靠搜索和Ai，此脚本仅为自己日常会用到的工具。（主打一个能用就行）
## 📘 脚本简介

`compress_xlsx.py` 是一个用于 **压缩 Excel (.xlsx) 文件中嵌入图片** 的实用脚本。  
它会自动：

- 解压 `.xlsx` 文件（本质是 ZIP 压缩包）
- 找到图片文件 (`xl/media/` 下)
- 根据设定的参数压缩图片（调整分辨率与质量）
- 重新打包生成一个新的 `_compressed.xlsx` 文件  
  👉 原始文件不会被修改

使用后可显著减小 Excel 文件体积，适合图片较多的报表、日志或图文汇总类 Excel。

---

## 🧰 环境要求

- Python 3.8.12
- Pillow 10.4.0

安装依赖：
```bash
pip install -r requirements.txt
```

使用方法：
```bash
python compress_xlsx.py <文件路径> [--resize 比例] [--quality 质量]
```

使用示例：
```bash
# 默认参数
python compress_xlsx.py report.xlsx
# 自定义参数
python compress_xlsx.py report.xlsx --resize 0.5 --quality 75
```

## 推荐参数

- 追求画质，轻微压缩：--resize 0.8 --quality 85
- 尽量压小，但可看清：--resize 0.5 --quality 75
- 仅需预览，无需细节：--resize 0.4 --quality 60