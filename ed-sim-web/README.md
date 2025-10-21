# 急诊排班仿真评估（Streamlit）

本项目提供一个网页应用，上传 `res_doctor.xlsx`，并读取同目录下：
- `2025 服务系统问题-问题数据.xlsx`（到达率）
- `optimized_schedule_IDs_01_matrix.xlsx`（优化方案，内置对比）

随后运行事件驱动仿真（周内等待口径A），输出：
- 总等待时间(人·小时) [平均±标准差 over N runs]
- 总工作成本、借调成本、总成本

---

## 一键本地运行

```bash
pip install -r requirements.txt
streamlit run app.py
```

## 部署方案 A：Streamlit Community Cloud（免服务器）

1. 将本仓库推到 GitHub（含 `app.py`、`requirements.txt`、两份 Excel：
   `2025 服务系统问题-问题数据.xlsx` 与 `optimized_schedule_IDs_01_matrix.xlsx`）。
2. 登录 https://streamlit.io → **Deploy an app** → 选择你的仓库、分支、入口文件 `app.py`。
3. 部署完成后，访问给定链接即可在线使用；来访者在网页里上传 `res_doctor.xlsx` 即可得到结果。

> 如需更新 Excel，直接更新 GitHub 仓库即可。也可在左侧上传自定义到达文件临时覆盖。

## 部署方案 B：Hugging Face Spaces（免服务器）

1. 登录 https://huggingface.co → **Create new Space** → 选择 **Streamlit** 框架。
2. 将本项目文件（含两个 Excel）上传到 Space 仓库。
3. Space 会自动安装 `requirements.txt` 并启动。访问 Space URL 即可。

## 方案 C：临时分享（本地运行 + 隧道）

如果只是临时让别人访问：
```bash
streamlit run app.py    # 本地起服务（默认 8501）
# 另开一个终端：
# 使用 cloudflared 或 ngrok 打开外网访问隧道（二选一）
cloudflared tunnel --url http://localhost:8501
# 或
ngrok http 8501
```
将隧道输出的网址发给对方即可临时访问（注意不要泄露机密数据）。

---

## 文件格式约定

- `res_doctor.xlsx`：
  - 第 2~19 行 × 第 2~169 列为 0/1 排班矩阵（18 × 168）。
  - 第 13 行及以后用于统计“借调医生数”（任一工时>0 计 1 人）。
- `2025 服务系统问题-问题数据.xlsx`：工作表名为“数据”，第 6~12 行 × 第 2~25 列为到达率（7×24）。
- `optimized_schedule_IDs_01_matrix.xlsx`：与 `res_doctor.xlsx` 相同格式。

