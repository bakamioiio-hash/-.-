# Skill: 01-daily-leads-extractor (海外客户自动化线索挖掘与清洗)

## 技能定位
调用本地脚本 `tools/extract_leads.py`。通过搜索引擎关键词轮换检索（或读取 `data/urls.txt`），自动抓取提炼海外独立站的商业邮箱、电话/WhatsApp、社媒链接及主营简介，去重后追加写入 `data/leads.xlsx`。

## 触发指令
当用户要求「跑爬虫」「爬取今日客户」「谷歌搜线索」「运行技能 05」时触发。

## 执行流程
1. 在项目根目录下运行清洗脚本：
```bash
python tools/extract_leads.py
```
2. 脚本运行完成后，自动读取 `data/leads.xlsx` 最新追加的记录。

## 汇报规范
在对话框中以 Markdown 表格汇报今日新增线索：
| 目标独立站 | 商业邮箱 (Emails) | 电话 / WhatsApp | 核心品类简介 |
| :--- | :--- | :--- | :--- |
| 示例域名 | contact@brand.com | +1 ... | 劳保产品系列 |