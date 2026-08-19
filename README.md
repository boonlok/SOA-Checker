# SOA 对账小助手（Streamlit 版，简洁粉色风）

比对 Xero 导出的 Invoice Excel 报表 和 供应商 Statement of Account (SOA) PDF。
**只比对发票号码和金额，不比对日期**。发票号码会自动标准化后再匹配，
格式上的小差异（比如 Excel 里的 `TB C2 (20059764)` 和 PDF 里的 `20059764`、
或者带不带前导0）都会被认成同一张发票，然后检查两边金额差了多少。

支持**一次上传多份 PDF**（比如每个月一份 SOA，最多十几份都没问题），
同一份 Excel 会分别跟每一份 PDF 对比，结果按 PDF 分标签页显示，
并可以打包下载全部报告。

结果展示方便一眼看出问题：
- 有差异/未对账的发票自动排在最前面，用颜色标出（玫红=金额不一致，琥珀=未对账，绿=匹配没问题）
- 每一行都会显示 Excel 金额、SOA 金额、以及差额是多少
- 顶部有一句大白话总结，比如"有2张金额对不上，需要看一下"
- 总览页汇总了所有PDF里需要关注的发票（含具体是哪份SOA文件），不用逐个点进去找
- 每个PDF页面里可以勾选"只看有问题的"、按发票号码搜索

UI 做了简化，只用一个粉色色系（不再是花花绿绿的渐变），状态颜色只保留
绿/玫红/琥珀三种，界面更干净、留白更多。

已用真实文件（REDFUSE ELECTRIC 的 Excel + PDF）以及模拟的"发票号码格式不同、
金额有细微差异"场景测试通过，一次上传多份PDF、逐份对比、打包下载CSV也都正常工作。

## 部署到 Streamlit Community Cloud（免费，约3分钟）

1. 去 https://share.streamlit.io 用 GitHub 账号登录。
2. 把这个文件夹推送到一个 GitHub 仓库（public 或 private 都可以）：
   ```
   git init
   git add .
   git commit -m "SOA checker (streamlit)"
   git branch -M main
   git remote add origin <你的仓库地址>
   git push -u origin main
   ```
3. 在 Streamlit Cloud 点击 **New app**，选择这个仓库，Main file path 填 `app.py`，点 **Deploy**。
4. 部署完成后，在该 App 的 **⋮ (右上角) → Settings → Secrets** 里粘贴：
   ```
   site_password = "你自己设定的密码"
   ```
   保存后 App 会自动重启并生效。
5. 把 Streamlit 给你的网址（类似 `https://soa-checker-xxxx.streamlit.app`）和密码
   发给需要使用的人即可，不需要额外注册账号。

> 免费额度对个人/小范围内部使用完全够用。App 长时间没人访问会自动休眠，
> 下次打开时等几秒钟重新唤醒即可，之后使用就正常了。

## 本地测试

```
pip install -r requirements.txt
cp .streamlit/secrets.toml.example .streamlit/secrets.toml
# 编辑 .streamlit/secrets.toml 改成你想要的密码
streamlit run app.py
```
打开 http://localhost:8501

## 文件说明

```
app.py                          主程序（含解析、比对逻辑、密码保护、简洁粉色界面）
requirements.txt                Python依赖
.streamlit/config.toml          粉色主题配色
.streamlit/secrets.toml.example 密码配置示例（真正的secrets.toml不要提交到公开仓库）
```

## 关于识别准确度

- Excel：需为 Xero 标准导出格式，表头包含 `Invoice Date`、`Reference`、`Gross` 三列
  （其他栏位比如 Project 会自动忽略，不影响比对）。同一张发票如果因为四舍五入
  等原因在 Excel 里拆成多行，系统会自动加总后再比对。
- PDF：目前仅支持**文字型PDF**（不是扫描/拍照得到的图片型PDF）。系统按
  `DATE TYPE DOCNO DUEDATE REF CUR DEBIT CREDIT` 格式逐行识别，SOA 里的单据类型：
  - `IV`（Invoice）和 `AD`（Adjustment）会计入该发票的金额
  - `CN`（Credit Note）如果发票号码跟某张 invoice 一样，会自动从该发票金额里扣掉
  - `OR`（Official Receipt，付款/收据记录）会被完全忽略，不计入发票金额
- 发票号码比对时会自动标准化（去掉括号、连字符、空格、前导0等），
  Excel 的 `TB C2 (20057646)` / `TB C2-20057646` 和 PDF 的 `20057646`
  都会被认成同一张发票，然后比较金额差多少。
- 如果供应商PDF格式不一样，把样本发给我，我可以针对性调整解析规则。
