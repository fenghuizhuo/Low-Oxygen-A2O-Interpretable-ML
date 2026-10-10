# -*- coding: utf-8 -*-
r"""
A²/O 四指标：DO≤2.0 mg/L 条件下的补充运行参数筛选
=====================================================

用途
----
本脚本用于在已经完成的“四指标一级A参考目标反向筛选”基础上，
进一步回答论文修改中的一个具体问题：

    在保持原有筛选标准不变的前提下，
    如果从低氧运行角度进一步要求好氧区 DO ≤ 2.0 mg/L，
    低、中、高三个进水 TP 水平下，原来表现较好的运行参数组合还剩多少？
    对应的 HRT 和厌氧区进水分配比范围是否发生变化？

重要说明
--------
1. 本脚本不是重新建模，也不重新进行正向预测。
2. 直接读取已经生成的：
       四指标_目标出水反向优化结果.xlsx
   中的“12_各组参数组合”工作表。
3. 原筛选标准完全不变：
   每个进水 TP 组内，仍以“达到本组最佳表现 90% 以上”作为表现较好组合的标准。
4. 本脚本只是在原来已经判定为“达到本组最佳90%以上”的组合中，
   再增加一个低氧条件：
       Aerobic_DO ≤ 2.0 mg/L
5. 不在 DO≤2.0 mg/L 后重新计算一个新的 90% 标准，
   避免因为缩小 DO 范围而改变原来的比较基准。
6. 正式粗网格的 DO 水平中没有正好 2.00 mg/L；
   按当前网格，DO≤2.0 mg/L 时实际纳入的最高 DO 水平应为 1.95 mg/L。
   因此论文中可将其表述为“DO 约控制在 2.0 mg/L 以下”。
7. 输出的 HRT、DO 和厌氧区进水分配比范围，
   是筛选后组合中实际出现的最小值和最大值，
   不表示三个参数区间内任意自由组合都能达到要求。
8. 所有结果仍属于冻结模型的一步条件预测筛选，
   不是长期稳态模拟，也不等同于法规意义上的日均达标判定。

默认输入
--------
${PROJECT_ROOT}/outputs/screening/inverse/四指标_目标出水反向优化结果.xlsx

默认输出目录
------------
${PROJECT_ROOT}/outputs/screening/inverse/low_do

正式输出
--------
1. 四指标_DO不高于2mg_L补充筛选结果.xlsx
2. 四指标_DO不高于2mg_L补充筛选说明.txt

可选环境变量
------------
A2O_LOWDO_INPUT_FILE
A2O_LOWDO_OUTPUT_DIR
A2O_LOWDO_MAX=2.0
A2O_LOWDO_STRICT=1

依赖
----
pip install pandas numpy openpyxl xlsxwriter
"""

from __future__ import annotations

import hashlib
import math
import os
import time
import traceback
from pathlib import Path
from typing import Any, Dict, List, Mapping, Optional, Sequence, Tuple

import numpy as np
import pandas as pd


# =============================================================================
# 1. 路径与固定设置
# =============================================================================

PROJECT_ROOT = Path(__file__).resolve().parents[1]

INPUT_FILE = Path(
    os.environ.get(
        "A2O_LOWDO_INPUT_FILE",
        str(PROJECT_ROOT / "outputs/screening/inverse/四指标_目标出水反向优化结果.xlsx"),
    )
)

OUTPUT_DIR = Path(
    os.environ.get(
        "A2O_LOWDO_OUTPUT_DIR",
        str(PROJECT_ROOT / "outputs/screening/inverse/low_do"),
    )
)

OUTPUT_EXCEL = OUTPUT_DIR / "四指标_DO不高于2mg_L补充筛选结果.xlsx"
OUTPUT_TEXT = OUTPUT_DIR / "四指标_DO不高于2mg_L补充筛选说明.txt"

SCRIPT_VERSION = "A2O_LOW_DO_SUPPLEMENT_v1.0_20261005"

STRICT_MODE = os.environ.get("A2O_LOWDO_STRICT", "1").strip() != "0"
LOW_DO_MAX = float(os.environ.get("A2O_LOWDO_MAX", "2.0"))

# 必须沿用原反向筛选中的 90% 标准。
NEAR_BEST_RATIO = 0.90

# 当前正式反向结果中，每个进水TP组均有 3780 组统一参数组合。
EXPECTED_COMBINATION_N = 3780

# 数值比较容差。
TOL = 1e-12

GROUP_ORDER = ["低进水TP", "中等进水TP", "高进水TP"]

SOURCE_SHEET = "12_各组参数组合"

REQUIRED_COLUMNS = [
    "进水TP水平",
    "分组规则",
    "本组历史背景数",
    "组合序号",
    "候选组合ID",
    "候选_DO_mg_L",
    "候选_HRT_h",
    "候选_Feed",
    "候选_Feed_%",
    "通过物理和数值检查背景数",
    "TP满足背景数",
    "TN满足背景数",
    "NH4满足背景数",
    "COD满足背景数",
    "四个出水指标同时达到要求背景数",
    "四个出水指标同时达到要求比例%",
    "相对本组最佳表现%",
    "是否达到本组最佳90%以上",
    "是否本组并列最佳组合",
]

NUMERIC_COLUMNS = [
    "本组历史背景数",
    "组合序号",
    "候选_DO_mg_L",
    "候选_HRT_h",
    "候选_Feed",
    "候选_Feed_%",
    "通过物理和数值检查背景数",
    "TP满足背景数",
    "TN满足背景数",
    "NH4满足背景数",
    "COD满足背景数",
    "四个出水指标同时达到要求背景数",
    "四个出水指标同时达到要求比例%",
    "相对本组最佳表现%",
]


# =============================================================================
# 2. 通用函数
# =============================================================================

def now_str() -> str:
    return time.strftime("%Y-%m-%d %H:%M:%S")


def log(message: str) -> None:
    print(f"[{now_str()}] {message}", flush=True)


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def safe_float(value: Any) -> float:
    try:
        number = float(value)
        return number if np.isfinite(number) else np.nan
    except Exception:
        return np.nan


def yes_mask(series: pd.Series) -> np.ndarray:
    return series.fillna("").astype(str).str.strip().eq("是").to_numpy(dtype=bool)


def group_sort_key(series: pd.Series) -> pd.Series:
    order = {name: i for i, name in enumerate(GROUP_ORDER)}
    return series.astype(str).map(order).fillna(len(order)).astype(int)


def fmt(value: Any, digits: int = 6) -> str:
    number = safe_float(value)
    if not np.isfinite(number):
        return "—"
    return f"{number:.{digits}g}"


def atomic_write_text(path: Path, text: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_name(path.stem + ".part.txt")
    tmp.write_text(text, encoding="utf-8-sig")
    os.replace(tmp, path)


def parameter_range(
    frame: pd.DataFrame,
    column: str,
) -> Tuple[float, float, int]:
    if frame.empty:
        return np.nan, np.nan, 0
    values = pd.to_numeric(frame[column], errors="coerce")
    values = values[np.isfinite(values.to_numpy(dtype=float))]
    if len(values) == 0:
        return np.nan, np.nan, 0
    return float(values.min()), float(values.max()), int(values.nunique())


def select_representative_best(frame: pd.DataFrame) -> Optional[pd.Series]:
    """
    在给定组合中选择“达到要求背景数最多”的代表组合。
    如果并列，则沿用原反向筛选的顺序：
        较低 DO -> 较短 HRT -> 较小 Feed。
    """
    if frame.empty:
        return None

    counts = pd.to_numeric(
        frame["四个出水指标同时达到要求背景数"],
        errors="coerce",
    )
    if not np.isfinite(counts.to_numpy(dtype=float)).any():
        return None

    best_count = int(np.nanmax(counts.to_numpy(dtype=float)))
    if best_count <= 0:
        return None

    candidates = frame.loc[counts == best_count].copy()
    candidates = candidates.sort_values(
        ["候选_DO_mg_L", "候选_HRT_h", "候选_Feed"],
        ascending=[True, True, True],
        kind="mergesort",
    )
    return candidates.iloc[0]


def cleanup_previous_outputs() -> None:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    for path in [
        OUTPUT_EXCEL,
        OUTPUT_EXCEL.with_name(OUTPUT_EXCEL.stem + ".part.xlsx"),
        OUTPUT_TEXT,
        OUTPUT_TEXT.with_name(OUTPUT_TEXT.stem + ".part.txt"),
    ]:
        try:
            if path.exists():
                path.unlink()
        except FileNotFoundError:
            pass


# =============================================================================
# 3. 读取原反向筛选结果并严格核查
# =============================================================================

def read_source_table() -> pd.DataFrame:
    if not INPUT_FILE.exists():
        raise FileNotFoundError(f"未找到原反向筛选结果：{INPUT_FILE}")

    with pd.ExcelFile(INPUT_FILE, engine="openpyxl") as excel:
        if SOURCE_SHEET not in excel.sheet_names:
            raise ValueError(
                f"输入文件缺少工作表“{SOURCE_SHEET}”。"
                "本补充分析必须直接读取原反向筛选的各组参数组合结果。"
            )
        frame = pd.read_excel(excel, sheet_name=SOURCE_SHEET)

    frame.columns = [str(c).strip() for c in frame.columns]

    missing = [c for c in REQUIRED_COLUMNS if c not in frame.columns]
    if missing:
        raise ValueError(f"“{SOURCE_SHEET}”缺少必要字段：{missing}")

    frame = frame[REQUIRED_COLUMNS].copy()

    for col in NUMERIC_COLUMNS:
        frame[col] = pd.to_numeric(frame[col], errors="coerce")

    for col in ["候选_DO_mg_L", "候选_HRT_h", "候选_Feed", "四个出水指标同时达到要求背景数"]:
        if frame[col].isna().any():
            raise ValueError(f"字段“{col}”存在缺失或非数值，无法开展补充筛选。")

    frame["进水TP水平"] = frame["进水TP水平"].astype(str).str.strip()
    frame["分组规则"] = frame["分组规则"].astype(str).str.strip()
    frame["是否达到本组最佳90%以上"] = (
        frame["是否达到本组最佳90%以上"].fillna("").astype(str).str.strip()
    )
    frame["是否本组并列最佳组合"] = (
        frame["是否本组并列最佳组合"].fillna("").astype(str).str.strip()
    )

    unexpected_groups = sorted(set(frame["进水TP水平"]) - set(GROUP_ORDER))
    if unexpected_groups:
        raise ValueError(f"发现未预期的进水TP分组：{unexpected_groups}")

    missing_groups = [g for g in GROUP_ORDER if g not in set(frame["进水TP水平"])]
    if missing_groups:
        raise ValueError(f"缺少进水TP分组：{missing_groups}")

    return frame.reset_index(drop=True)


def validate_original_selection(frame: pd.DataFrame) -> pd.DataFrame:
    """
    核查原Excel中的“是否达到本组最佳90%以上”是否与原规则一致。
    补充筛选必须保持原标准，不重新定义新的90%门槛。
    """
    rows: List[Dict[str, Any]] = []

    for group_name in GROUP_ORDER:
        group = frame.loc[frame["进水TP水平"] == group_name].copy()

        duplicate_n = int(
            group[["候选_DO_mg_L", "候选_HRT_h", "候选_Feed"]]
            .duplicated()
            .sum()
        )

        count_values = group["四个出水指标同时达到要求背景数"].to_numpy(dtype=float)
        best_count = int(np.max(count_values)) if len(count_values) else 0
        threshold = (
            int(math.ceil(NEAR_BEST_RATIO * best_count - 1e-12))
            if best_count > 0
            else 0
        )

        expected_good = (
            (count_values >= threshold) & (count_values > 0)
            if best_count > 0
            else np.zeros(len(group), dtype=bool)
        )
        saved_good = yes_mask(group["是否达到本组最佳90%以上"])
        mismatch_n = int(np.sum(expected_good != saved_good))

        background_values = (
            pd.to_numeric(group["本组历史背景数"], errors="coerce")
            .dropna()
            .unique()
        )
        background_n = (
            int(background_values[0])
            if len(background_values) == 1
            else np.nan
        )

        rows.append(
            {
                "进水TP水平": group_name,
                "本组历史背景数": background_n,
                "本组参数组合数": int(len(group)),
                "重复DO_HRT_Feed组合数": duplicate_n,
                "本组最佳达到要求背景数": best_count,
                "原90%门槛背景数": threshold,
                "原Excel标记为达到最佳90%以上的组合数": int(np.sum(saved_good)),
                "重新核算达到最佳90%以上的组合数": int(np.sum(expected_good)),
                "原90%标记不一致数": mismatch_n,
                "核查结论": (
                    "通过"
                    if (
                        duplicate_n == 0
                        and mismatch_n == 0
                        and len(group) == EXPECTED_COMBINATION_N
                    )
                    else "需核对"
                ),
            }
        )

        if STRICT_MODE:
            if len(group) != EXPECTED_COMBINATION_N:
                raise ValueError(
                    f"{group_name}参数组合数={len(group)}，"
                    f"预期={EXPECTED_COMBINATION_N}。"
                )
            if duplicate_n > 0:
                raise ValueError(
                    f"{group_name}存在{duplicate_n}组重复DO×HRT×进水分配比组合。"
                )
            if mismatch_n > 0:
                raise ValueError(
                    f"{group_name}原Excel中的90%筛选标记有{mismatch_n}条"
                    "与原规则不一致。请先核对原反向筛选结果。"
                )

    return pd.DataFrame(rows)


# =============================================================================
# 4. DO≤2.0 mg/L 补充筛选
# =============================================================================

def build_analysis(
    frame: pd.DataFrame,
) -> Tuple[
    pd.DataFrame,
    pd.DataFrame,
    pd.DataFrame,
    pd.DataFrame,
    pd.DataFrame,
]:
    """
    返回：
    1) 三个进水TP组汇总
    2) 参数范围对比
    3) 低DO较优组合明细
    4) 低DO代表组合
    5) 论文可用简表
    """
    summary_rows: List[Dict[str, Any]] = []
    range_rows: List[Dict[str, Any]] = []
    selected_frames: List[pd.DataFrame] = []
    representative_rows: List[Dict[str, Any]] = []
    paper_rows: List[Dict[str, Any]] = []

    for group_name in GROUP_ORDER:
        group = frame.loc[frame["进水TP水平"] == group_name].copy()
        group = group.sort_values("组合序号", kind="mergesort").reset_index(drop=True)

        saved_good = yes_mask(group["是否达到本组最佳90%以上"])
        original_good = group.loc[saved_good].copy()

        low_do_all = group.loc[
            group["候选_DO_mg_L"] <= LOW_DO_MAX + TOL
        ].copy()

        low_do_good = group.loc[
            saved_good
            & (group["候选_DO_mg_L"].to_numpy(dtype=float) <= LOW_DO_MAX + TOL)
        ].copy()

        best_count = int(
            group["四个出水指标同时达到要求背景数"].max()
        )
        threshold = (
            int(math.ceil(NEAR_BEST_RATIO * best_count - 1e-12))
            if best_count > 0
            else 0
        )

        background_values = (
            group["本组历史背景数"].dropna().astype(int).unique()
        )
        background_n = (
            int(background_values[0])
            if len(background_values) == 1
            else np.nan
        )

        actual_low_do_max = (
            float(low_do_all["候选_DO_mg_L"].max())
            if not low_do_all.empty
            else np.nan
        )

        low_do_best_count = (
            int(low_do_all["四个出水指标同时达到要求背景数"].max())
            if not low_do_all.empty
            else 0
        )
        low_do_best_ratio_to_original = (
            100.0 * low_do_best_count / best_count
            if best_count > 0
            else np.nan
        )

        original_best_in_low_do = bool(
            best_count > 0
            and np.any(
                (group["四个出水指标同时达到要求背景数"].to_numpy(dtype=float) == best_count)
                & (group["候选_DO_mg_L"].to_numpy(dtype=float) <= LOW_DO_MAX + TOL)
            )
        )

        # 原较优组合范围。
        orig_do_min, orig_do_max, orig_do_levels = parameter_range(
            original_good, "候选_DO_mg_L"
        )
        orig_hrt_min, orig_hrt_max, orig_hrt_levels = parameter_range(
            original_good, "候选_HRT_h"
        )
        orig_feed_min, orig_feed_max, orig_feed_levels = parameter_range(
            original_good, "候选_Feed"
        )

        # 低DO较优组合范围。
        low_do_min, low_do_max, low_do_levels = parameter_range(
            low_do_good, "候选_DO_mg_L"
        )
        low_hrt_min, low_hrt_max, low_hrt_levels = parameter_range(
            low_do_good, "候选_HRT_h"
        )
        low_feed_min, low_feed_max, low_feed_levels = parameter_range(
            low_do_good, "候选_Feed"
        )

        if best_count <= 0:
            conclusion = (
                "原筛选中本组没有找到四项出水指标同时达到参考要求的参数组合；"
                f"增加DO≤{LOW_DO_MAX:g} mg/L条件后仍无可推荐组合。"
            )
        elif low_do_good.empty:
            conclusion = (
                f"原筛选存在表现较好组合，但增加DO≤{LOW_DO_MAX:g} mg/L条件后，"
                "没有组合能够继续达到原本组最佳表现的90%以上。"
            )
        elif low_do_best_count == best_count:
            conclusion = (
                f"增加DO≤{LOW_DO_MAX:g} mg/L条件后仍保留原本组最佳表现；"
                "说明不需要提高到更高DO才能获得本组最高的历史背景覆盖水平。"
            )
        else:
            conclusion = (
                f"增加DO≤{LOW_DO_MAX:g} mg/L条件后，本组最佳覆盖水平有所下降，"
                "论文中应同时说明降低DO上限带来的影响。"
            )

        summary_rows.append(
            {
                "进水TP水平": group_name,
                "分组规则": str(group["分组规则"].iloc[0]),
                "历史背景数": background_n,
                "原本组最佳达到要求背景数": best_count,
                "原90%门槛背景数": threshold,
                "原达到最佳90%以上组合数": int(len(original_good)),
                f"DO≤{LOW_DO_MAX:g}后达到原最佳90%以上组合数": int(len(low_do_good)),
                "低DO较优组合占原较优组合比例%": (
                    100.0 * len(low_do_good) / len(original_good)
                    if len(original_good) > 0
                    else np.nan
                ),
                f"DO≤{LOW_DO_MAX:g}范围内最佳达到要求背景数": low_do_best_count,
                "低DO最佳表现相对原最佳表现%": low_do_best_ratio_to_original,
                "原最佳组合是否在低DO范围内": (
                    "是" if original_best_in_low_do else ("否" if best_count > 0 else "不适用")
                ),
                f"DO≤{LOW_DO_MAX:g}时实际最高计算DO_mg_L": actual_low_do_max,
                "原较优_DO范围_mg_L": (
                    f"{fmt(orig_do_min)}~{fmt(orig_do_max)}"
                    if len(original_good) > 0 else "无"
                ),
                "原较优_HRT范围_h": (
                    f"{fmt(orig_hrt_min)}~{fmt(orig_hrt_max)}"
                    if len(original_good) > 0 else "无"
                ),
                "原较优_厌氧区进水分配比范围": (
                    f"{fmt(orig_feed_min)}~{fmt(orig_feed_max)}"
                    if len(original_good) > 0 else "无"
                ),
                "低DO较优_DO范围_mg_L": (
                    f"{fmt(low_do_min)}~{fmt(low_do_max)}"
                    if len(low_do_good) > 0 else "无"
                ),
                "低DO较优_HRT范围_h": (
                    f"{fmt(low_hrt_min)}~{fmt(low_hrt_max)}"
                    if len(low_do_good) > 0 else "无"
                ),
                "低DO较优_厌氧区进水分配比范围": (
                    f"{fmt(low_feed_min)}~{fmt(low_feed_max)}"
                    if len(low_do_good) > 0 else "无"
                ),
                "结果说明": conclusion,
            }
        )

        # 参数范围对比表：原较优组合与低DO较优组合并排整理。
        for range_type, selected in [
            ("原达到本组最佳90%以上的组合", original_good),
            (f"在原90%标准基础上再要求DO≤{LOW_DO_MAX:g} mg/L", low_do_good),
        ]:
            for parameter, chinese_name, column, unit in [
                ("Aerobic_DO", "好氧区DO", "候选_DO_mg_L", "mg/L"),
                ("HRT", "水力停留时间", "候选_HRT_h", "h"),
                (
                    "Anaerobic_Feed_Fraction",
                    "厌氧区进水分配比",
                    "候选_Feed",
                    "比例",
                ),
            ]:
                minimum, maximum, levels = parameter_range(selected, column)
                range_rows.append(
                    {
                        "进水TP水平": group_name,
                        "分组规则": str(group["分组规则"].iloc[0]),
                        "范围类型": range_type,
                        "运行参数": parameter,
                        "中文含义": chinese_name,
                        "最小值": minimum,
                        "最大值": maximum,
                        "单位": unit,
                        "实际出现水平数": levels,
                        "对应参数组合数": int(len(selected)),
                        "说明": (
                            "范围表示筛选后组合中实际出现的最小值和最大值，"
                            "不表示三个参数区间内任意自由组合都能达到要求。"
                        ),
                    }
                )

        # 保存低DO较优组合明细。
        if not low_do_good.empty:
            detail = low_do_good.copy()
            detail.insert(
                0,
                "低DO补充筛选说明",
                f"原达到本组最佳90%以上，且DO≤{LOW_DO_MAX:g} mg/L",
            )
            detail["_组顺序"] = group_sort_key(detail["进水TP水平"])
            detail = detail.sort_values(
                [
                    "_组顺序",
                    "四个出水指标同时达到要求背景数",
                    "候选_DO_mg_L",
                    "候选_HRT_h",
                    "候选_Feed",
                ],
                ascending=[True, False, True, True, True],
                kind="mergesort",
            ).drop(columns=["_组顺序"])
            detail.insert(
                1,
                "低DO组内排序",
                np.arange(1, len(detail) + 1, dtype=int),
            )
            selected_frames.append(detail)

        # 每个进水TP组给出一个“低DO条件下表现最好”的代表组合。
        representative = select_representative_best(low_do_all)
        if representative is not None:
            representative_rows.append(
                {
                    "进水TP水平": group_name,
                    "分组规则": str(group["分组规则"].iloc[0]),
                    "代表方案": f"DO≤{LOW_DO_MAX:g} mg/L条件下表现最好",
                    "DO_mg_L": float(representative["候选_DO_mg_L"]),
                    "HRT_h": float(representative["候选_HRT_h"]),
                    "厌氧区进水分配比": float(representative["候选_Feed"]),
                    "四项同时达到要求背景数": int(
                        representative["四个出水指标同时达到要求背景数"]
                    ),
                    "历史背景数": background_n,
                    "四项同时达到要求比例%": (
                        100.0
                        * float(representative["四个出水指标同时达到要求背景数"])
                        / float(background_n)
                        if np.isfinite(background_n) and background_n > 0
                        else np.nan
                    ),
                    "相对原本组最佳表现%": (
                        100.0
                        * float(representative["四个出水指标同时达到要求背景数"])
                        / float(best_count)
                        if best_count > 0
                        else np.nan
                    ),
                    "说明": (
                        "若有多个组合并列最佳，依次优先选择较低DO、较短HRT、较小进水分配比。"
                    ),
                }
            )
        else:
            representative_rows.append(
                {
                    "进水TP水平": group_name,
                    "分组规则": str(group["分组规则"].iloc[0]),
                    "代表方案": f"DO≤{LOW_DO_MAX:g} mg/L条件下表现最好",
                    "DO_mg_L": np.nan,
                    "HRT_h": np.nan,
                    "厌氧区进水分配比": np.nan,
                    "四项同时达到要求背景数": 0,
                    "历史背景数": background_n,
                    "四项同时达到要求比例%": 0.0,
                    "相对原本组最佳表现%": np.nan,
                    "说明": "本组没有找到四项同时达到参考要求的组合。",
                }
            )

        # 论文可直接参考的简表。
        if best_count <= 0:
            paper_text = (
                f"{group_name}：在原扫描范围内未筛得四项出水指标同时达到参考要求的组合；"
                f"增加DO≤{LOW_DO_MAX:g} mg/L条件后仍未筛得可行组合。"
            )
        elif not low_do_good.empty:
            paper_text = (
                f"{group_name}：在原筛选得到的{len(original_good)}组较优组合中，"
                f"有{len(low_do_good)}组同时满足DO≤{LOW_DO_MAX:g} mg/L；"
                f"实际低DO计算范围为{fmt(low_do_min)}~{fmt(low_do_max)} mg/L，"
                f"HRT为{fmt(low_hrt_min)}~{fmt(low_hrt_max)} h，"
                f"厌氧区进水分配比为{fmt(low_feed_min)}~{fmt(low_feed_max)}；"
                f"低DO条件下最佳组合仍可覆盖{low_do_best_count}/{background_n}个历史状态。"
            )
        else:
            paper_text = (
                f"{group_name}：原筛选存在较优组合，但加入DO≤{LOW_DO_MAX:g} mg/L条件后，"
                "未保留达到原最佳表现90%以上的组合。"
            )

        paper_rows.append(
            {
                "进水TP水平": group_name,
                "论文结果表述参考": paper_text,
            }
        )

    summary_df = pd.DataFrame(summary_rows)
    range_df = pd.DataFrame(range_rows)

    if selected_frames:
        selected_df = pd.concat(selected_frames, ignore_index=True)
    else:
        selected_df = pd.DataFrame(
            columns=[
                "低DO补充筛选说明",
                "低DO组内排序",
            ] + REQUIRED_COLUMNS
        )

    representative_df = pd.DataFrame(representative_rows)
    paper_df = pd.DataFrame(paper_rows)

    return summary_df, range_df, selected_df, representative_df, paper_df


# =============================================================================
# 5. 输出
# =============================================================================

def build_settings_table(
    source_frame: pd.DataFrame,
    validation_df: pd.DataFrame,
) -> pd.DataFrame:
    do_values = np.sort(
        pd.to_numeric(source_frame["候选_DO_mg_L"], errors="coerce")
        .dropna()
        .unique()
    )
    actual_max = (
        float(np.max(do_values[do_values <= LOW_DO_MAX + TOL]))
        if np.any(do_values <= LOW_DO_MAX + TOL)
        else np.nan
    )

    rows = [
        {
            "项目": "分析名称",
            "内容": "DO不高于2.0 mg/L条件下的补充运行参数筛选",
            "说明": "在原反向筛选结果基础上增加低DO条件，不重新建模、不重新预测。",
        },
        {
            "项目": "脚本版本",
            "内容": SCRIPT_VERSION,
            "说明": "",
        },
        {
            "项目": "输入文件",
            "内容": str(INPUT_FILE),
            "说明": "读取原反向筛选结果。",
        },
        {
            "项目": "输入文件SHA256",
            "内容": sha256_file(INPUT_FILE),
            "说明": "用于结果追踪。",
        },
        {
            "项目": "输入工作表",
            "内容": SOURCE_SHEET,
            "说明": "直接读取低、中、高进水TP组的3780组参数组合统计。",
        },
        {
            "项目": "低DO条件",
            "内容": f"好氧区DO ≤ {LOW_DO_MAX:g} mg/L",
            "说明": "论文中可表述为DO控制在约2.0 mg/L以下。",
        },
        {
            "项目": "低DO条件下实际最高计算DO",
            "内容": f"{fmt(actual_max)} mg/L",
            "说明": (
                "由于原正式粗网格没有正好2.00 mg/L，"
                "当前DO≤2.0 mg/L时实际最高计算水平为1.95 mg/L。"
            ),
        },
        {
            "项目": "原筛选标准",
            "内容": "每个进水TP组内保留达到本组最佳表现90%以上的组合",
            "说明": "本补充分析保持该标准不变。",
        },
        {
            "项目": "本补充分析的关键原则",
            "内容": "先沿用原90%较优组合，再要求DO≤2.0 mg/L",
            "说明": "不在缩小DO范围后重新定义新的90%门槛。",
        },
        {
            "项目": "解释边界",
            "内容": "冻结模型的一步条件预测筛选",
            "说明": "不是长期稳态模拟，也不等同于法规意义上的日均达标。",
        },
        {
            "项目": "参数范围含义",
            "内容": "筛选后组合中实际出现的最小值~最大值",
            "说明": "不表示DO、HRT和进水分配比区间内任意自由组合都能达到要求。",
        },
        {
            "项目": "严格核查",
            "内容": (
                "通过"
                if (validation_df["核查结论"] == "通过").all()
                else "存在需核对项目"
            ),
            "说明": "核查原90%标记、每组3780组合及重复组合情况。",
        },
    ]
    return pd.DataFrame(rows)


def build_checks_table(
    source_frame: pd.DataFrame,
    validation_df: pd.DataFrame,
) -> pd.DataFrame:
    rows: List[Dict[str, Any]] = []

    rows.append(
        {
            "核查项目": "输入文件存在",
            "核查结果": "通过" if INPUT_FILE.exists() else "失败",
            "数值或说明": str(INPUT_FILE),
        }
    )
    rows.append(
        {
            "核查项目": "输入总记录数",
            "核查结果": "通过",
            "数值或说明": int(len(source_frame)),
        }
    )

    for _, row in validation_df.iterrows():
        group_name = str(row["进水TP水平"])
        rows.extend(
            [
                {
                    "核查项目": f"{group_name}参数组合数",
                    "核查结果": (
                        "通过"
                        if int(row["本组参数组合数"]) == EXPECTED_COMBINATION_N
                        else "失败"
                    ),
                    "数值或说明": int(row["本组参数组合数"]),
                },
                {
                    "核查项目": f"{group_name}重复DO×HRT×进水分配比组合",
                    "核查结果": (
                        "通过"
                        if int(row["重复DO_HRT_Feed组合数"]) == 0
                        else "失败"
                    ),
                    "数值或说明": int(row["重复DO_HRT_Feed组合数"]),
                },
                {
                    "核查项目": f"{group_name}原90%标记一致性",
                    "核查结果": (
                        "通过"
                        if int(row["原90%标记不一致数"]) == 0
                        else "失败"
                    ),
                    "数值或说明": (
                        f"不一致数={int(row['原90%标记不一致数'])}；"
                        f"原最佳={int(row['本组最佳达到要求背景数'])}；"
                        f"90%门槛={int(row['原90%门槛背景数'])}"
                    ),
                },
            ]
        )

    do_values = np.sort(
        source_frame["候选_DO_mg_L"].dropna().to_numpy(dtype=float)
    )
    do_values = np.unique(do_values)
    included = do_values[do_values <= LOW_DO_MAX + TOL]

    rows.append(
        {
            "核查项目": f"DO≤{LOW_DO_MAX:g} mg/L时是否存在可用DO水平",
            "核查结果": "通过" if len(included) > 0 else "失败",
            "数值或说明": (
                f"纳入{len(included)}个DO水平；"
                f"最低={fmt(np.min(included) if len(included) else np.nan)}；"
                f"最高={fmt(np.max(included) if len(included) else np.nan)}"
            ),
        }
    )

    return pd.DataFrame(rows)


def write_excel(
    settings_df: pd.DataFrame,
    summary_df: pd.DataFrame,
    range_df: pd.DataFrame,
    selected_df: pd.DataFrame,
    representative_df: pd.DataFrame,
    checks_df: pd.DataFrame,
    paper_df: pd.DataFrame,
) -> None:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    tmp = OUTPUT_EXCEL.with_name(OUTPUT_EXCEL.stem + ".part.xlsx")
    if tmp.exists():
        tmp.unlink()

    with pd.ExcelWriter(
        tmp,
        engine="xlsxwriter",
        datetime_format="yyyy-mm-dd",
        engine_kwargs={"options": {"strings_to_urls": False, "nan_inf_to_errors": True}},
    ) as writer:

        tables = [
            ("01_说明与条件", settings_df),
            ("02_分组结果汇总", summary_df),
            ("03_参数范围对比", range_df),
            ("04_低DO较优组合", selected_df),
            ("05_低DO代表组合", representative_df),
            ("06_方法检查", checks_df),
            ("07_论文表述参考", paper_df),
        ]

        workbook = writer.book

        header_fmt = workbook.add_format(
            {
                "bold": True,
                "bg_color": "#D9EAF7",
                "border": 1,
                "align": "center",
                "valign": "vcenter",
                "text_wrap": True,
            }
        )
        number_fmt = workbook.add_format({"num_format": "0.000000"})
        integer_fmt = workbook.add_format({"num_format": "0"})
        highlight_fmt = workbook.add_format(
            {"bg_color": "#E2F0D9", "bold": True}
        )
        warning_fmt = workbook.add_format(
            {"bg_color": "#FFF2CC"}
        )

        for sheet_name, frame in tables:
            out = frame.copy()
            if out.empty:
                out = pd.DataFrame([{"说明": "本工作表当前没有记录。"}])

            out.to_excel(writer, sheet_name=sheet_name, index=False)
            ws = writer.sheets[sheet_name]
            ws.freeze_panes(1, 0)

            if len(out.columns) > 0:
                ws.autofilter(
                    0,
                    0,
                    max(len(out), 1),
                    len(out.columns) - 1,
                )

            for col_idx, col in enumerate(out.columns):
                ws.write(0, col_idx, col, header_fmt)

                sample = out[col].head(500).astype(str)
                width = max(12, len(str(col)) + 2)
                if len(sample):
                    try:
                        width = max(
                            width,
                            int(sample.map(len).quantile(0.95)) + 2,
                        )
                    except Exception:
                        pass
                width = min(width, 44)

                cell_fmt = None
                if pd.api.types.is_integer_dtype(out[col]):
                    cell_fmt = integer_fmt
                elif pd.api.types.is_float_dtype(out[col]):
                    cell_fmt = number_fmt

                ws.set_column(col_idx, col_idx, width, cell_fmt)

            if sheet_name == "02_分组结果汇总":
                ws.set_tab_color("#70AD47")
                # 对“原最佳组合是否在低DO范围内=是”做轻度提示。
                if "原最佳组合是否在低DO范围内" in out.columns:
                    col_idx = out.columns.get_loc("原最佳组合是否在低DO范围内")
                    ws.conditional_format(
                        1,
                        col_idx,
                        len(out),
                        col_idx,
                        {
                            "type": "text",
                            "criteria": "containing",
                            "value": "是",
                            "format": highlight_fmt,
                        },
                    )

            if sheet_name == "04_低DO较优组合":
                ws.set_tab_color("#5B9BD5")

            if sheet_name == "06_方法检查":
                if "核查结果" in out.columns:
                    col_idx = out.columns.get_loc("核查结果")
                    ws.conditional_format(
                        1,
                        col_idx,
                        len(out),
                        col_idx,
                        {
                            "type": "text",
                            "criteria": "containing",
                            "value": "失败",
                            "format": warning_fmt,
                        },
                    )

    os.replace(tmp, OUTPUT_EXCEL)


def build_text_report(
    summary_df: pd.DataFrame,
    representative_df: pd.DataFrame,
    checks_df: pd.DataFrame,
    elapsed_seconds: float,
) -> str:
    lines: List[str] = []

    lines.append("A²/O 四指标：DO≤2.0 mg/L 条件下的补充运行参数筛选")
    lines.append("=" * 66)
    lines.append("")
    lines.append("一、分析目的")
    lines.append(
        "本分析不重新建模、不重新预测，只在原反向筛选结果基础上，"
        f"进一步要求好氧区DO≤{LOW_DO_MAX:g} mg/L，"
        "比较低、中、高三个进水TP水平下表现较好的运行参数范围是否发生变化。"
    )
    lines.append("")
    lines.append("二、筛选原则")
    lines.append(
        "1. 原有“达到本组最佳表现90%以上”的标准保持不变；"
        "先按原标准确定较优组合，再增加低DO条件。"
    )
    lines.append(
        f"2. 不在DO≤{LOW_DO_MAX:g} mg/L后重新计算新的90%门槛。"
    )
    lines.append(
        "3. 当前正式粗网格中没有正好2.00 mg/L；"
        "因此DO≤2.0 mg/L时实际最高计算DO为1.95 mg/L。"
    )
    lines.append("")

    lines.append("三、三个进水TP水平的补充筛选结果")
    for _, row in summary_df.iterrows():
        lines.append("")
        lines.append(f"[{row['进水TP水平']}] {row['分组规则']}")
        lines.append(
            f"- 原达到本组最佳90%以上的组合："
            f"{int(row['原达到最佳90%以上组合数'])}组。"
        )
        lowdo_col = f"DO≤{LOW_DO_MAX:g}后达到原最佳90%以上组合数"
        lines.append(
            f"- 再要求DO≤{LOW_DO_MAX:g} mg/L后："
            f"{int(row[lowdo_col])}组。"
        )
        lines.append(
            f"- 原较优范围：DO {row['原较优_DO范围_mg_L']} mg/L；"
            f"HRT {row['原较优_HRT范围_h']} h；"
            f"厌氧区进水分配比 {row['原较优_厌氧区进水分配比范围']}。"
        )
        lines.append(
            f"- 低DO较优范围：DO {row['低DO较优_DO范围_mg_L']} mg/L；"
            f"HRT {row['低DO较优_HRT范围_h']} h；"
            f"厌氧区进水分配比 {row['低DO较优_厌氧区进水分配比范围']}。"
        )
        if int(row["原本组最佳达到要求背景数"]) > 0:
            lines.append(
                f"- 原最佳达到要求背景数："
                f"{int(row['原本组最佳达到要求背景数'])}；"
                f"低DO范围内最佳达到要求背景数："
                f"{int(row[f'DO≤{LOW_DO_MAX:g}范围内最佳达到要求背景数'])}。"
            )
            lines.append(
                f"- 低DO最佳表现相对原最佳表现："
                f"{fmt(row['低DO最佳表现相对原最佳表现%'], 5)}%。"
            )
        lines.append(f"- 结果说明：{row['结果说明']}")

    lines.append("")
    lines.append("四、低DO条件下的代表组合")
    for _, row in representative_df.iterrows():
        if int(row["四项同时达到要求背景数"]) <= 0:
            lines.append(
                f"- {row['进水TP水平']}：没有找到四项同时达到参考要求的低DO组合。"
            )
        else:
            lines.append(
                f"- {row['进水TP水平']}：DO={fmt(row['DO_mg_L'])} mg/L，"
                f"HRT={fmt(row['HRT_h'])} h，"
                f"厌氧区进水分配比={fmt(row['厌氧区进水分配比'])}；"
                f"可使{int(row['四项同时达到要求背景数'])}/"
                f"{int(row['历史背景数'])}个历史状态满足四项参考目标。"
            )

    lines.append("")
    lines.append("五、解释边界")
    lines.append(
        "这些结果是在冻结模型的一步条件预测基础上得到的运行参数筛选结果，"
        "用于比较低氧条件下不同运行参数组合的相对表现；"
        "不表示长期稳态结果，也不表示三个参数范围内任意自由组合均可达到要求。"
    )
    lines.append("")
    lines.append("六、方法核查")
    fail_n = int((checks_df["核查结果"].astype(str) == "失败").sum())
    lines.append(
        f"方法检查共{len(checks_df)}项，失败项={fail_n}。"
    )
    lines.append("")
    lines.append(f"运行耗时：{elapsed_seconds:.2f} s")
    lines.append(f"结果Excel：{OUTPUT_EXCEL}")
    lines.append(f"说明文件：{OUTPUT_TEXT}")

    return "\n".join(lines)


# =============================================================================
# 6. 主程序
# =============================================================================

def main() -> None:
    start = time.time()
    cleanup_previous_outputs()

    log("开始DO≤2.0 mg/L条件下的补充运行参数筛选。")
    log(f"原反向筛选结果：{INPUT_FILE}")
    log(f"输出目录：{OUTPUT_DIR}")
    log(
        "原则：保持原有“达到本组最佳表现90%以上”的标准不变，"
        f"在此基础上再要求DO≤{LOW_DO_MAX:g} mg/L。"
    )

    try:
        source = read_source_table()
        log(f"已读取“{SOURCE_SHEET}”：{len(source)}条参数组合记录。")

        validation_df = validate_original_selection(source)
        log("原反向筛选的90%标准与保存结果核查完成。")

        (
            summary_df,
            range_df,
            selected_df,
            representative_df,
            paper_df,
        ) = build_analysis(source)

        checks_df = build_checks_table(source, validation_df)
        settings_df = build_settings_table(source, validation_df)

        if STRICT_MODE and (checks_df["核查结果"].astype(str) == "失败").any():
            bad = checks_df.loc[
                checks_df["核查结果"].astype(str) == "失败"
            ]
            raise RuntimeError(
                "补充筛选方法检查未全部通过：\n"
                + bad.to_string(index=False)
            )

        write_excel(
            settings_df=settings_df,
            summary_df=summary_df,
            range_df=range_df,
            selected_df=selected_df,
            representative_df=representative_df,
            checks_df=checks_df,
            paper_df=paper_df,
        )

        elapsed = time.time() - start
        report = build_text_report(
            summary_df=summary_df,
            representative_df=representative_df,
            checks_df=checks_df,
            elapsed_seconds=elapsed,
        )
        atomic_write_text(OUTPUT_TEXT, report)

        log("补充筛选完成。")
        log(f"结果Excel：{OUTPUT_EXCEL}")
        log(f"说明文件：{OUTPUT_TEXT}")

        # 终端打印最关键结果，便于运行后立即核对。
        log("三个进水TP水平关键结果：")
        for _, row in summary_df.iterrows():
            lowdo_col = f"DO≤{LOW_DO_MAX:g}后达到原最佳90%以上组合数"
            log(
                f"{row['进水TP水平']}："
                f"原较优组合={int(row['原达到最佳90%以上组合数'])}；"
                f"低DO较优组合={int(row[lowdo_col])}；"
                f"低DO范围 DO={row['低DO较优_DO范围_mg_L']} mg/L，"
                f"HRT={row['低DO较优_HRT范围_h']} h，"
                f"Feed={row['低DO较优_厌氧区进水分配比范围']}。"
            )

    except Exception:
        log("运行失败。错误如下：")
        traceback.print_exc()
        raise


if __name__ == "__main__":
    main()
