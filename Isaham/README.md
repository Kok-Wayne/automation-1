# Bursa Malaysia Contact Info Scraper

自动从 isaham.my 抓取上市公司官网，并提取电话、Email、地址，导出为 Excel 文件。

---

## 安装步骤

### 1. 安装 Python（需要 3.10+）
https://www.python.org/downloads/

### 2. 安装依赖
打开终端（cmd / PowerShell / Terminal），进入脚本所在文件夹：

```bash
pip install -r requirements.txt
```

### 3. 安装 Google Chrome
脚本使用 Chrome 浏览器自动抓取。若没装，请先安装：
https://www.google.com/chrome/

---

## 使用方式

### 抓取所有 Sector 的公司（完整模式，可能需要数小时）
```bash
python scraper.py
```

### 只抓一个 Sector（推荐先测试）
```bash
python scraper.py --sector semiconductors --max 10
```

### 常用参数

| 参数 | 说明 | 例子 |
|------|------|------|
| `--sector` | 指定 sector slug | `--sector rubber-gloves` |
| `--max` | 每个 sector 最多抓几家（0=全部） | `--max 20` |
| `--output` | 输出文件名 | `--output mydata.xlsx` |
| `--visible` | 显示浏览器窗口（方便调试） | `--visible` |

### 例子

```bash
# 抓取半导体行业前20家公司，显示浏览器窗口
python scraper.py --sector semiconductors --max 20 --visible

# 抓取橡胶手套行业全部公司
python scraper.py --sector rubber-gloves

# 抓取所有行业（完整运行）
python scraper.py --output bursa_contacts_2026.xlsx
```

---

## Sector Slug 参考

常用 sector 的 slug（URL 末段）：

| 行业 | Slug |
|------|------|
| 半导体 | `semiconductors` |
| 橡胶手套 | `rubber-gloves` |
| 银行 | `banking` |
| 科技 | `technology` |
| 房地产 | `property` |
| 医疗 | `healthcare` |
| 种植 | `plantation` |
| 汽车 | `automotive` |
| 石油天然气 | `oil-and-gas` |
| 人工智能 | `artificial-intelligence` |

---

## 输出 Excel 格式

| 栏位 | 说明 |
|------|------|
| Stock Code | 股票代号 |
| Company Name | 公司名称 |
| Sector | 所属行业 |
| Website | 公司官网 |
| Phone | 联系电话 |
| Email | 电子邮件 |
| Address | 公司地址 |
| Status | 抓取状态 |

---

## 注意事项

- 脚本每家公司之间有随机延迟（1-2秒），避免被网站封锁
- 抓取完每个 sector 后自动保存 Excel（断点续抓）
- 如遇到某公司官网无法访问，Status 栏会显示原因
- 如需抓取全部公司，建议分 sector 分批运行
