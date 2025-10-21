# app.py
# -*- coding: utf-8 -*-
"""
急诊仿真网页（Streamlit）
功能：
1) 上传 res_doctor.xlsx（外部来访者的排班表）。
2) 读取本地 arrival 数据文件：2025 服务系统问题-问题数据.xlsx（默认放在与本脚本同目录）。
3) 读取本地优化方案 optimized_schedule_IDs_01_matrix.xlsx（与来访者方案对比）。
4) 运行给定的事件驱动仿真 1000 次，输出：
   总等待时间(人·小时) [平均±标准差 over N runs]
   总工作成本、借调成本、总成本
5) 可选：下载结果 CSV。

请将以下文件放到本脚本同目录：
- 2025 服务系统问题-问题数据.xlsx   （到达率）
- optimized_schedule_IDs_01_matrix.xlsx （我们给出的优化方案）

运行：
  pip install -r requirements.txt
  streamlit run app.py
"""

import io
import time
import heapq
from collections import deque
from typing import Tuple, Optional

import numpy as np
import pandas as pd
import streamlit as st
from pathlib import Path

# -------------------------
# 页面设置
# -------------------------
st.set_page_config(page_title="ED 仿真评估", page_icon="⚕️", layout="wide")
st.title("⚕️ 急诊排班仿真评估（事件驱动 · 周内等待口径A）")
st.caption("上传排班表，快速得到等待与成本指标；并与内置优化排班进行对比。（已移除自定义到达率上传接口，固定读取仓库内的到达率与优化方案。）")

# -------------------------
# 读取/解析工具（用脚本所在目录作为基准，避免工作目录不同导致找不到文件）
# -------------------------
BASE_DIR = Path(__file__).resolve().parent
ARRIVAL_DEFAULT_PATH = BASE_DIR / "2025 服务系统问题-问题数据.xlsx"
OPTIMIZED_SCHEDULE_PATH = BASE_DIR / "optimized_schedule_IDs_01_matrix.xlsx"

@st.cache_data(show_spinner=False)
def load_arrival_rates_from_excel(path: str, sheet_name: str = "数据") -> np.ndarray:
    """读取到达率：7天 × 24小时，返回 shape (7,24) 的 float ndarray。"""
    df = pd.read_excel(path, sheet_name=sheet_name, header=None)
    # 与用户原代码一致：第 6~12 行（0-based 5:12），第 2~25 列（1:25）
    arr = df.iloc[5:12, 1:25].values.astype(float)
    if arr.shape != (7, 24):
        raise ValueError(f"到达率矩阵应为 (7,24)，当前为 {arr.shape}。请检查文件/切片。")
    return arr


def _borrow_count_like_user(df: pd.DataFrame) -> int:
    """按用户口径：取出第13行及之后的子表，统计借调医生数（是否有任意工时）"""
    borrow_part = df.iloc[12:, 1:]  # 第13行起、去掉第一列索引列
    return int((borrow_part.sum(axis=1) > 0).sum())


def _schedule_matrix_from_df(df: pd.DataFrame) -> np.ndarray:
    """解析 18×168 的 0/1 矩阵（第2~19行、第2~169列），与用户代码对齐。"""
    sched = df.iloc[1:19, 1:169].values.astype(int)
    if sched.shape != (18, 168):
        raise ValueError(f"排班矩阵应为 (18,168)，当前为 {sched.shape}。")
    return sched


@st.cache_data(show_spinner=False)
def load_schedule_from_excel(content: bytes, filename: str = "upload.xlsx") -> Tuple[np.ndarray, int]:
    """从上传内容解析排班矩阵和借调人数。自动适配 xlsx/xls/csv/xlsb。
    返回 (schedule_matrix, borrow_count)。"""
    from pathlib import Path
    suffix = Path(filename).suffix.lower()
    bio = io.BytesIO(content)

    df = None
    try:
        if suffix in {".xlsx", ".xlsm", ".xltx", ".xltm"}:
            xl = pd.ExcelFile(bio, engine="openpyxl")
            if not xl.sheet_names:
                raise ValueError("上传文件未检测到工作表（xlsx/xlsm）。请确认不是空白或受保护文件。")
            df = pd.read_excel(xl, sheet_name=0, header=None)
        elif suffix == ".xls":
            xl = pd.ExcelFile(bio, engine="xlrd")
            if not xl.sheet_names:
                raise ValueError("上传的 .xls 文件未检测到工作表。")
            df = pd.read_excel(xl, sheet_name=0, header=None, engine="xlrd")
        elif suffix == ".xlsb":
            xl = pd.ExcelFile(bio, engine="pyxlsb")
            if not xl.sheet_names:
                raise ValueError("上传的 .xlsb 文件未检测到工作表。")
            df = pd.read_excel(xl, sheet_name=0, header=None, engine="pyxlsb")
        elif suffix == ".csv":
            bio.seek(0)
            df = pd.read_csv(bio, header=None)
        else:
            bio.seek(0)
            xl = pd.ExcelFile(bio)
            if not xl.sheet_names:
                raise ValueError("上传文件未检测到任何工作表。")
            df = pd.read_excel(xl, sheet_name=0, header=None)
    except Exception as e:
        raise ValueError(f"无法读取排班文件：{e}")

    borrow = _borrow_count_like_user(df)
    sched = _schedule_matrix_from_df(df)
    return sched, borrow


@st.cache_data(show_spinner=False)
def load_optimized_schedule(path: str = OPTIMIZED_SCHEDULE_PATH) -> Tuple[np.ndarray, int]:
    df = pd.read_excel(path, sheet_name=0, header=None)
    borrow = _borrow_count_like_user(df)
    sched = _schedule_matrix_from_df(df)
    return sched, borrow


# -------------------------
# 仿真核心（与用户给定代码一致的逻辑，做了函数化和小修正）
# -------------------------
class Simulator:
    def __init__(self,
                 arrival_rates: np.ndarray,  # (7,24)
                 service_rate_per_doctor: float = 6.0,
                 n_simulations: int = 1000,
                 seed_base: int = 0):
        self.arrival_rates = arrival_rates
        self.service_rate_per_doctor = float(service_rate_per_doctor)
        self.n_simulations = int(n_simulations)
        self.seed_base = int(seed_base)

    def run_for_schedule(self, doctor_schedule: np.ndarray, borrow_doctor_count: int) -> dict:
        n_doctors, n_hours = doctor_schedule.shape
        T_end = float(n_hours)

        # 计算总到达、总服务能力（用于 sanity check 文案）
        total_arrivals = float(self.arrival_rates.sum())  # 7*24 小时总到达率之和（人）
        # 每小时能力 = 在班医生数 * μ；总能力 = Σ 小时能力
        hourly_doctors = doctor_schedule.sum(axis=0)  # (168,)
        total_capacity = float(np.dot(hourly_doctors, np.full(n_hours, self.service_rate_per_doctor)))

        # 事件优先级
        EVENT_ARRIVAL = 0
        EVENT_CHECK   = 1
        EVENT_FINISH  = 2

        def on_shift(d, t):
            idx = int(t)
            return (0 <= idx < n_hours) and (doctor_schedule[d, idx] == 1)

        def simulate_once(seed: Optional[int] = None) -> float:
            if seed is not None:
                np.random.seed(seed)

            waiting_queue = deque()
            total_wait_time = 0.0
            doctor_busy_until = np.zeros(n_doctors, dtype=float)
            events = []

            # 到达事件生成（piecewise-constant NHPP）
            for hour in range(n_hours):
                day = hour // 24
                hr_in_day = hour % 24
                lam = float(self.arrival_rates[day % 7, hr_in_day])
                if lam <= 0:
                    continue
                t = 0.0
                while True:
                    inter_arrival = np.random.exponential(1.0 / lam)
                    t += inter_arrival
                    if t >= 1.0:
                        break
                    arrival_time = hour + t
                    service_time = np.random.exponential(1.0 / self.service_rate_per_doctor)
                    heapq.heappush(events, (arrival_time, EVENT_ARRIVAL, "arrival", None, (arrival_time, service_time)))

            # 整点换班检查
            for hour in range(n_hours):
                heapq.heappush(events, (float(hour), EVENT_CHECK, "check_shift", None, None))

            while events:
                current_time, _, event_type, doctor_id, payload = heapq.heappop(events)
                if event_type == "arrival":
                    arrival_time, service_time = payload
                    waiting_queue.append((arrival_time, service_time))
                    for d in range(n_doctors):
                        if not waiting_queue:
                            break
                        if on_shift(d, current_time) and doctor_busy_until[d] <= current_time:
                            start_time = max(current_time, doctor_busy_until[d])
                            if on_shift(d, start_time):
                                pat_arrival, pat_service = waiting_queue.popleft()
                                finish_time = start_time + pat_service
                                doctor_busy_until[d] = finish_time
                                total_wait_time += (start_time - pat_arrival)
                                heapq.heappush(events, (finish_time, EVENT_FINISH, "finish", d, None))

                elif event_type == "finish":
                    d = doctor_id
                    if d is None:
                        continue
                    if on_shift(d, current_time) and waiting_queue:
                        start_time = max(current_time, doctor_busy_until[d])
                        if on_shift(d, start_time):
                            pat_arrival, pat_service = waiting_queue.popleft()
                            finish_time = start_time + pat_service
                            doctor_busy_until[d] = finish_time
                            total_wait_time += (start_time - pat_arrival)
                            heapq.heappush(events, (finish_time, EVENT_FINISH, "finish", d, None))

                elif event_type == "check_shift":
                    for d in range(n_doctors):
                        if not waiting_queue:
                            break
                        if on_shift(d, current_time) and doctor_busy_until[d] <= current_time:
                            start_time = max(current_time, doctor_busy_until[d])
                            if on_shift(d, start_time):
                                pat_arrival, pat_service = waiting_queue.popleft()
                                finish_time = start_time + pat_service
                                doctor_busy_until[d] = finish_time
                                total_wait_time += (start_time - pat_arrival)
                                heapq.heappush(events, (finish_time, EVENT_FINISH, "finish", d, None))

            # 口径A：只统计周内等待（末尾未开诊者等到 T_end）
            total_wait_time += sum(max(0.0, T_end - arr) for (arr, _) in waiting_queue)
            return total_wait_time

        # 多次仿真
        all_total_wait = np.zeros(self.n_simulations, dtype=float)
        for i in range(self.n_simulations):
            all_total_wait[i] = simulate_once(seed=self.seed_base + i)

        # 成本
        total_doctor_hours = float(doctor_schedule.sum())
        avg_total_wait = float(all_total_wait.mean())
        std_total_wait = float(all_total_wait.std(ddof=1)) if self.n_simulations > 1 else 0.0

        # 文案提示（供显示）
        supply_demand_note = (
            "⚠️ 供给 < 需求：周末会积压，口径A的尾巴等待会较大。"
            if total_capacity < total_arrivals else
            "✅ 供给 ≥ 需求：尾巴不应过大；若等待仍陡增，需排查事件逻辑或数据切片。"
        )

        return {
            "avg_total_wait": avg_total_wait,
            "std_total_wait": std_total_wait,
            "total_doctor_hours": total_doctor_hours,
            "borrow_doctor_count": borrow_doctor_count,
            "total_cost": avg_total_wait + total_doctor_hours * 1.3 + borrow_doctor_count * 20.0,
            "note": supply_demand_note
        }


# -------------------------
# 侧边栏参数
# -------------------------
st.sidebar.header("参数设置")
mu = st.sidebar.number_input("单医生服务率 μ (人/小时)", min_value=0.1, max_value=60.0, value=6.0, step=0.1)
nruns = st.sidebar.number_input("仿真次数", min_value=1, max_value=5000, value=1000, step=100)
seed0 = st.sidebar.number_input("随机种子基数", min_value=0, max_value=10_000_000, value=0, step=1)

# （按用户要求）去掉自定义到达率上传接口；一律使用仓库内默认到达文件

# 必传：res_doctor.xlsx
st.subheader("上传来访者排班表（res_doctor.xlsx）")
upload = st.file_uploader("选择 Excel 文件", type=["xlsx", "xls"], accept_multiple_files=False)

col_run, col_info = st.columns([1, 1])

with col_info:
    st.markdown("""
**文件格式要求**
- **res_doctor.xlsx**：
  - 第 2~19 行 × 第 2~169 列为 0/1 排班矩阵（18 × 168）。
  - 第 13 行及以后用于统计“借调医生数”（任一工时>0 计 1 人）。
- **到达文件**：固定从仓库内 `2025 服务系统问题-问题数据.xlsx` 读取（工作表名“数据”，第 6~12 行 × 第 2~25 列为到达率 7×24）。
- **优化方案**：固定从仓库内 `optimized_schedule_IDs_01_matrix.xlsx` 读取相同格式矩阵。
""")

# -------------------------
# 主逻辑
# -------------------------
if upload is not None:
    try:
        # 到达率（固定本地文件）
        if not ARRIVAL_DEFAULT_PATH.exists():
            raise FileNotFoundError(f"未找到到达率文件：{ARRIVAL_DEFAULT_PATH}")
        arrival_rates = load_arrival_rates_from_excel(ARRIVAL_DEFAULT_PATH)

        # 来访者排班
        if upload.size == 0:
            raise ValueError("上传的排班文件为空（0 字节）。请检查导出方式并重试。")
        sched_user, borrow_user = load_schedule_from_excel(upload.getvalue(), filename=upload.name)
        # 内置优化排班
        sched_opt, borrow_opt = load_optimized_schedule(OPTIMIZED_SCHEDULE_PATH)

        # 运行仿真
        sim = Simulator(arrival_rates, service_rate_per_doctor=mu, n_simulations=nruns, seed_base=seed0)

        with st.spinner("正在运行来访者排班仿真……"):
            res_user = sim.run_for_schedule(sched_user, borrow_user)
        with st.spinner("正在运行优化排班仿真……"):
            res_opt = sim.run_for_schedule(sched_opt, borrow_opt)

        # 展示结果
        st.subheader("结果")
        c1, c2 = st.columns(2)
        def _pretty_block(container, title, res):
            container.markdown(f"### {title}")
            container.write(res["note"])
            container.code(
                f"总等待时间(人·小时) [平均±标准差 over {nruns} runs]: "
                f"{res['avg_total_wait']:.2f} ± {res['std_total_wait']:.2f}\n"
                f"总工作成本: {res['total_doctor_hours']*1.3:.2f}\n"
                f"借调成本: {res['borrow_doctor_count']*20.0:.2f}\n"
                f"总成本: {res['total_cost']:.2f}",
                language="text"
            )

        _pretty_block(c1, "来访者上传排班", res_user)
        _pretty_block(c2, "内置优化排班", res_opt)

        # 汇总表 & 下载
        out_df = pd.DataFrame([
            {
                "方案": "来访者",
                "平均总等待(人·小时)": res_user["avg_total_wait"],
                "等待标准差": res_user["std_total_wait"],
                "总工作时长": res_user["total_doctor_hours"],
                "借调人数": res_user["borrow_doctor_count"],
                "总工作成本": res_user["total_doctor_hours"] * 1.3,
                "借调成本": res_user["borrow_doctor_count"] * 20.0,
                "总成本": res_user["total_cost"],
            },
            {
                "方案": "优化方案",
                "平均总等待(人·小时)": res_opt["avg_total_wait"],
                "等待标准差": res_opt["std_total_wait"],
                "总工作时长": res_opt["total_doctor_hours"],
                "借调人数": res_opt["borrow_doctor_count"],
                "总工作成本": res_opt["total_doctor_hours"] * 1.3,
                "借调成本": res_opt["borrow_doctor_count"] * 20.0,
                "总成本": res_opt["total_cost"],
            },
        ])
        st.dataframe(out_df, use_container_width=True)

        csv = out_df.to_csv(index=False).encode("utf-8-sig")
        st.download_button("下载结果 CSV", data=csv, file_name="simulation_results.csv", mime="text/csv")

    except Exception as e:
        # 显示更详细的错误上下文，帮助定位是哪一个文件导致
        st.error("出错了：" + str(e))
        st.info(
            f"调试信息：
"
            f"- 上传文件名：{upload.name if upload else '（未上传）'}
"
            f"- 固定到达率路径：{ARRIVAL_DEFAULT_PATH}
"
            f"- 固定优化方案路径：{OPTIMIZED_SCHEDULE_PATH}"
        )
else:
    st.info("请上传 res_doctor.xlsx（支持 .xlsx/.xls）。")
