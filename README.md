# 🚄 国铁电子客票解析工具

[![Python 3.8+](https://img.shields.io/badge/python-3.8+-blue.svg)](https://www.python.org/downloads/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Flask](https://img.shields.io/badge/Flask-2.0+-green.svg)](https://flask.palletsprojects.com/)

批量解析国铁电子客票 PDF，自动提取车次、票价、行程等信息，支持导出 CSV / Excel。

![演示截图](https://via.placeholder.com/800x400/1677ff/ffffff?text=Web+界面预览)

---

## ✨ 功能特性

- **📁 批量解析** — 一次性上传多个 PDF 文件，后端依次处理
- **📊 逐文件进度** — 每解析完一个文件，进度条实时前进，状态一目了然
- **🔍 排序与筛选** — 点击表头排序、按车次类型/席别筛选、关键词搜索
- **📈 图表统计** — 按月票价柱图、热门路线图、车次类型分布、席别分布
- **📋 详情弹窗** — 点击任意行查看完整 16 个字段
- **📥 双格式导出** — CSV / Excel（含格式化的表头、奇偶行隔色、自适应列宽）
- **🛏️ 席别兼容** — 支持高铁/动车（二等座）和卧铺（上铺/下铺/中铺）

---

## 📂 项目结构

```
guotie-ticket-parser/
├── tickets.csv             # 离线批量输出示例（含脱敏数据）
├── extract_tickets.py      # CLI 离线批量工具
├── requirements.txt        # Python 依赖
├── .gitignore              # Git 忽略配置（PDF/CSV 等不上传）
├── LICENSE                 # MIT 开源协议
├── README.md               # 项目说明
├── CONTRIBUTING.md         # 贡献指南
├── CHANGELOG.md            # 版本变更日志
│
└── web_app/
    ├── app.py              # Flask 后端服务
    └── templates/
        └── index.html      # 前端页面
```

> 💡 示例 PDF 文件夹（001.pdf ~ 017.pdf）不在本仓库中，请自行准备发票 PDF。

---

## 🚀 快速开始

### 环境要求

- Python 3.8+
- pip 包管理器

### 安装依赖

```bash
pip install -r requirements.txt
```

### 启动 Web 服务

```bash
cd web_app
python app.py
# 访问 http://localhost:5678
```

### 使用离线批量工具（CLI）

```bash
# 将 PDF 文件放在项目根目录，然后运行
python extract_tickets.py
# 输出 tickets.csv
```

---

## 📊 提取字段

| 字段 | 说明 | 示例 |
|------|------|------|
| 文件名 | 原始 PDF 文件名 | `001.pdf` |
| 发票号码 | 18 位发票号 | `26379118647000002974` |
| 开票日期 | 格式：YYYY年M月D日 | `2026年02月11日` |
| 出发站 | 火车站名 | `威海南海` |
| 到达站 | 火车站名 | `青岛北` |
| 车次 | 车次代码 | `G6989`、`C6494` |
| 乘车日期 | 格式：YYYY年M月D日 | `2025年10月18日` |
| 开车时间 | 24 小时制 | `10:47` |
| 车厢号 | 数字 | `08` |
| 座位号 | 含铺位标注 | `10F`、`上铺` |
| 席别 | 座位类型 | `二等座`、`软卧` |
| 票价（元） | 数字，保留两位小数 | `110.00` |
| 乘车人 | 乘客姓名 | `张三` |
| 证件号（脱敏） | 中间四位用 * 替换 | `********0812` |
| 电子客票号 | 21 位电子客票号 | `1864780086101690009372025` |
| 备注 | 始发改签、学生票等 | `始发改签` |

---

## 🛠️ 技术栈

- **后端**: Python + Flask
- **PDF 解析**: pdfplumber
- **前端**: 原生 HTML5 + CSS3 + JavaScript
- **图表**: Chart.js
- **Excel 导出**: openpyxl

---

## ⚠️ 已知限制

- 图片型 PDF（文字经 OCR 嵌入的扫描件）无法直接解析，需 OCR 处理
- 仅支持中国铁路 12306 电子客票格式

---

## 🤝 贡献指南

欢迎提交 Issue 和 Pull Request！请参阅 [CONTRIBUTING.md](./CONTRIBUTING.md) 了解详情。

---

## 📜 许可证

本项目采用 [MIT 许可证](./LICENSE) 开源。

---

## 🙏 致谢

- [pdfplumber](https://github.com/jsvine/pdfplumber) - PDF 解析库
- [Chart.js](https://www.chartjs.org/) - 图表库
- [Flask](https://flask.palletsprojects.com/) - Web 框架
