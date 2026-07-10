# 自动化订单清洗系统

这个项目是一个基于 Streamlit 的 Excel 订单自动清洗与映射工具，支持上传订单表格后进行字段清洗、年级映射、渠道补齐和结果导出。

## 在线使用（Streamlit Cloud）

应用已部署到 Streamlit Community Cloud，可以直接在线使用：

👉 **[点击打开在线应用](https://share.streamlit.io/1482619383-stack/update2/main/app.py)**

## Streamlit Cloud 部署步骤

1. 确保仓库已推送到 GitHub（`app.py` 和 `requirements.txt` 在仓库根目录）。
2. 访问 [share.streamlit.io](https://share.streamlit.io)，使用 GitHub 账号登录。
3. 点击 **New app**，选择仓库 `1482619383-stack/update2`，主文件路径填 `app.py`。
4. 点击 **Deploy**，等待几分钟后即可获得在线访问链接。

## 本地运行

```bash
pip install -r requirements.txt
streamlit run app.py
```

## GitHub Pages 静态介绍页

仓库同时配置了 GitHub Pages 静态介绍页（位于 `docs/` 目录），通过 GitHub Actions 自动部署。这只是一个项目介绍页面，实际的数据处理功能在 Streamlit 应用中运行。

### 部署步骤

1. 进入仓库 Settings → Pages。
2. 在 Source 中选择 GitHub Actions。
3. 推送到 main 分支后，GitHub Actions 会自动发布静态页面。
