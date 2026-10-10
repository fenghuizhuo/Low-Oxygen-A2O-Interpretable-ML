# -*- coding: utf-8 -*-
r"""
A²/O 四指标最终模型：二维 PDP 条件响应分析
============================================

用途
----
本脚本接在最新版“一维 PDP / ICE 运行响应分析”之后，直接读取一维结果 Excel
中的“后续二维候选”和“响应网格定义”，使用已经冻结的最终 F2 模型开展二维条件响应分析。

当前冻结模型固定为：
- TP       : F2 + SVR        / SVR_032
- TN       : F2 + SVR        / SVR_018
- NH4+-N   : F2 + ElasticNet / ElasticNet_025
- COD      : F2 + ElasticNet / ElasticNet_040

按当前一维结果，正式二维候选通常为：
- TP : HRT × Anaerobic_Volume_Fraction
- TP : Aerobic_DO × HRT
- TP : Aerobic_DO × Anaerobic_Volume_Fraction
- TN : Aerobic_DO × HRT
- TN : Aerobic_DO × Anoxic_Volume_Fraction
- TN : HRT × Anoxic_Volume_Fraction

重要原则
--------
1. 直接读取冻结 joblib，不重新选择模型、不重新调参、不重新训练最终模型。
2. 二维候选完全读取一维结果“后续二维候选”，不在本脚本中重新人为扩充组合。
3. 二维网格完全复用一维结果“响应网格定义”中的网格值，保证一维与二维范围一致。
4. 每个二维网格单元同时改变两个运行变量，其余进水条件、反应器结构变量以及
   对应目标的 lag1 / lag2 / roll3 保持每条历史参考样本原值。
5. 每个二维网格单元直接使用全部“基本物理条件有效”的历史参考样本计算经典二维 PDP。
   基本物理条件包括：
   - Aerobic_DO >= 0
   - 0 <= Anaerobic_Feed_Fraction <= 1
   - HRT > 0
   - 0 <= Anaerobic_Volume_Fraction <= 1
   - 0 <= Anoxic_Volume_Fraction <= 1
   - Anaerobic_Volume_Fraction + Anoxic_Volume_Fraction < 1
6. 同时输出绝对二维 PDP 平均预测值和中心化二维响应。中心化响应以同一历史样本
   原始状态预测值为基准，表示只改变当前两个运行变量后冻结模型预测值的平均变化。
7. 历史精确二元组合出现次数只作为结果解释和审计信息，不作为二维网格过滤条件。
8. 计算双向切片：
   - 固定变量2，扫描变量1；
   - 固定变量1，扫描变量2。
   切片响应方向沿用一维阶段的 Spearman 和近零响应判定规则。
9. 计算二维表面非加性指数，用于检查冻结模型二维表面偏离简单可加结构的程度。
   该指标只描述模型表面，不作为因果交互证据。
10. 100 次 Bootstrap 只对历史参考样本组成进行有放回重采样，不重新训练模型。
    输出每个二维网格单元的经验区间、二维表面幅度稳定性、表面形状 Spearman
    以及最低/最高中心化响应网格出现频率。
11. NH4+-N 和 COD 是否进入二维完全以最新版一维结果“后续二维候选”为准。
    当前一维结果若没有这两个目标的候选，本脚本不会强行构造二维表面。
12. 本阶段只输出 Excel 数据表和详细 TXT 说明，不生成 PNG/PDF/SVG 等图片。
13. 所有二维结果均为冻结预测模型在历史变量范围内的条件响应，不是实验干预因果效应，
    也不是运行条件改变后的长期稳态结果。

默认输入
--------
动态数据：
${PROJECT_ROOT}/outputs/preprocessing/dynamic/动态可用数据与插补比较结果.xlsx

最终模型目录：
${PROJECT_ROOT}/outputs/final_selection/moxingbaocun

SHAP + PI 结果：
${PROJECT_ROOT}/outputs/interpretation/shap_pi/四指标_SHAP与PermutationImportance综合结果.xlsx

一维响应结果：
${PROJECT_ROOT}/outputs/response/one_dimensional/四指标_一维PDP_ICE响应分析结果.xlsx

默认输出
--------
${PROJECT_ROOT}/outputs/response/two_dimensional

正式只输出：
1. 四指标_二维PDP响应分析结果.xlsx
2. 四指标_二维PDP响应分析说明.txt

不会输出图片。

可选环境变量
------------
A2O_2D_INPUT_FILE
A2O_2D_MODEL_DIR
A2O_2D_SHAP_PI_FILE
A2O_2D_ONE_D_FILE
A2O_2D_OUTPUT_DIR
A2O_2D_BOOTSTRAP_N=100
A2O_2D_INCLUDE_EXPLORATORY=0
A2O_2D_STRICT=1
A2O_2D_SMOKE_TEST=0
A2O_2D_DRY_RUN=0

依赖
----
pip install pandas numpy openpyxl xlsxwriter scipy scikit-learn joblib
"""

from __future__ import annotations

import hashlib
import json
import math
import os
import re
import time
import traceback
import warnings
from collections import Counter
from pathlib import Path
from typing import Any, Dict, List, Mapping, Optional, Sequence, Tuple

import joblib
import numpy as np
import pandas as pd
from scipy.stats import spearmanr


# =============================================================================
# 1. 路径、常量与最终模型配置锁定
# =============================================================================

PROJECT_ROOT = Path(__file__).resolve().parents[1]

INPUT_FILE = Path(
    os.environ.get(
        "A2O_2D_INPUT_FILE",
        str(PROJECT_ROOT / "outputs/preprocessing/dynamic/动态可用数据与插补比较结果.xlsx"),
    )
)

MODEL_DIR = Path(
    os.environ.get(
        "A2O_2D_MODEL_DIR",
        str(PROJECT_ROOT / "outputs/final_selection/moxingbaocun"),
    )
)

SHAP_PI_FILE = Path(
    os.environ.get(
        "A2O_2D_SHAP_PI_FILE",
        str(PROJECT_ROOT / "outputs/interpretation/shap_pi/四指标_SHAP与PermutationImportance综合结果.xlsx"),
    )
)

ONE_D_FILE = Path(
    os.environ.get(
        "A2O_2D_ONE_D_FILE",
        str(PROJECT_ROOT / "outputs/response/one_dimensional/四指标_一维PDP_ICE响应分析结果.xlsx"),
    )
)

OUTPUT_DIR = Path(
    os.environ.get(
        "A2O_2D_OUTPUT_DIR",
        str(PROJECT_ROOT / "outputs/response/two_dimensional"),
    )
)

OUTPUT_EXCEL = OUTPUT_DIR / "四指标_二维PDP响应分析结果.xlsx"
OUTPUT_TEXT = OUTPUT_DIR / "四指标_二维PDP响应分析说明.txt"

RANDOM_SEED = 20260620
ANALYSIS_VERSION = "A2O_2D_PDP_v4_bootstrap_nonadditivity_note"
STRICT_MODE = os.environ.get("A2O_2D_STRICT", "1").strip() != "0"
SMOKE_TEST = os.environ.get("A2O_2D_SMOKE_TEST", "0").strip() == "1"
DRY_RUN = os.environ.get("A2O_2D_DRY_RUN", "0").strip() == "1"
INCLUDE_EXPLORATORY = os.environ.get("A2O_2D_INCLUDE_EXPLORATORY", "0").strip() == "1"

BOOTSTRAP_N = int(os.environ.get("A2O_2D_BOOTSTRAP_N", "100"))
if BOOTSTRAP_N <= 0:
    raise ValueError("A2O_2D_BOOTSTRAP_N 必须为正整数。")
if SMOKE_TEST:
    BOOTSTRAP_N = min(BOOTSTRAP_N, 5)

MAX_INTERPOLATION_GAP = 7
ROUNDTRIP_METRIC_TOL = 1e-9
MEDIAN_TOL = 1e-10
FLOAT_LEVEL_ATOL = 1e-10

# 与最新版一维阶段保持一致。
MONOTONIC_SPEARMAN_THRESHOLD = 0.80
NEGLIGIBLE_NORMALIZED_AMPLITUDE = 0.02

TARGET_ORDER = ["TP", "TN", "NH4+-N", "COD"]

RAW_FEATURES = [
    "Aerobic_DO",
    "Anaerobic_Feed_Fraction",
    "Inf.TP",
    "Inf.COD",
    "Inf.NH4+-N",
    "Inf.TN",
    "HRT",
    "Anaerobic_Volume_Fraction",
    "Anoxic_Volume_Fraction",
]

INFLUENT_FEATURES = ["Inf.TP", "Inf.COD", "Inf.NH4+-N", "Inf.TN"]

CONTROLLABLE_FEATURES = [
    "Aerobic_DO",
    "Anaerobic_Feed_Fraction",
    "HRT",
    "Anaerobic_Volume_Fraction",
    "Anoxic_Volume_Fraction",
]

TARGET_CONFIG: Dict[str, Dict[str, Any]] = {
    "TP": {
        "sheet": "TP动态可用表",
        "target_col": "eff.TP",
        "history": ["eff.TP_lag1", "eff.TP_lag2", "eff.TP_roll3"],
        "joblib": "TP_最终模型.joblib",
    },
    "TN": {
        "sheet": "TN动态可用表",
        "target_col": "eff.TN",
        "history": ["eff.TN_lag1", "eff.TN_lag2", "eff.TN_roll3"],
        "joblib": "TN_最终模型.joblib",
    },
    "NH4+-N": {
        "sheet": "氨氮动态可用表",
        "target_col": "eff.NH4+-N",
        "history": ["eff.NH4+-N_lag1", "eff.NH4+-N_lag2", "eff.NH4+-N_roll3"],
        "joblib": "NH4_N_最终模型.joblib",
    },
    "COD": {
        "sheet": "COD动态可用表",
        "target_col": "eff.COD",
        "history": ["eff.COD_lag1", "eff.COD_lag2", "eff.COD_roll3"],
        "joblib": "COD_最终模型.joblib",
    },
}

EXPECTED_FINAL_CONFIG: Dict[str, Dict[str, Any]] = {
    "TP": {
        "feature_group": "F2",
        "algorithm": "SVR",
        "candidate_id": "SVR_032",
        "parameters": {"C": 10.0, "epsilon": 0.05, "gamma": 0.01, "kernel": "rbf"},
    },
    "TN": {
        "feature_group": "F2",
        "algorithm": "SVR",
        "candidate_id": "SVR_018",
        "parameters": {"C": 1.0, "epsilon": 0.1, "gamma": 0.03, "kernel": "rbf"},
    },
    "NH4+-N": {
        "feature_group": "F2",
        "algorithm": "ElasticNet",
        "candidate_id": "ElasticNet_025",
        "parameters": {"alpha": 0.01, "l1_ratio": 0.9},
    },
    "COD": {
        "feature_group": "F2",
        "algorithm": "ElasticNet",
        "candidate_id": "ElasticNet_040",
        "parameters": {"alpha": 0.3, "l1_ratio": 0.9},
    },
}

TEXTUAL_MISSING_VALUES = {
    "", " ", "NA", "N/A", "na", "n/a", "NaN", "nan", "None", "none",
    "NULL", "null", "-", "—", "–", "－", "/", "\\", "未测", "缺失",
}


# =============================================================================
# 2. 通用工具
# =============================================================================

def now_str() -> str:
    return time.strftime("%Y-%m-%d %H:%M:%S")


def log(message: str) -> None:
    print(f"[{now_str()}] {message}", flush=True)


def stable_seed(text: str) -> int:
    digest = hashlib.md5(str(text).encode("utf-8")).hexdigest()
    return int(digest[:8], 16)


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def clean_column_name(value: Any) -> str:
    text = str(value)
    text = text.replace("\ufeff", "").replace("\xa0", " ").replace("\u3000", " ")
    text = text.replace("\n", "").replace("\r", "").strip()
    return re.sub(r"\s+", "", text)


def convert_numeric(series: pd.Series) -> pd.Series:
    output = series.copy()
    if output.dtype == "object" or str(output.dtype).startswith("string"):
        output = output.astype("string").str.strip()
        output = output.str.replace(",", "", regex=False).str.replace("，", "", regex=False)
        output = output.mask(output.isin(TEXTUAL_MISSING_VALUES), np.nan)
    return pd.to_numeric(output, errors="coerce")


def safe_float(value: Any) -> float:
    try:
        number = float(value)
        return number if np.isfinite(number) else np.nan
    except Exception:
        return np.nan


def fmt(value: Any, digits: int = 6) -> str:
    number = safe_float(value)
    if not np.isfinite(number):
        return "不适用"
    return f"{number:.{digits}f}"


def params_equal(left: Mapping[str, Any], right: Mapping[str, Any]) -> bool:
    if set(left.keys()) != set(right.keys()):
        return False
    for key in left:
        a, b = left[key], right[key]
        if isinstance(a, (int, float, np.integer, np.floating)) and isinstance(
            b, (int, float, np.integer, np.floating)
        ):
            if not np.isclose(float(a), float(b), rtol=1e-12, atol=1e-12):
                return False
        elif str(a) != str(b):
            return False
    return True


def calc_metrics(y_true: Sequence[float], y_pred: Sequence[float]) -> Dict[str, float]:
    true = np.asarray(y_true, dtype=float)
    pred = np.asarray(y_pred, dtype=float)
    mask = np.isfinite(true) & np.isfinite(pred)
    true = true[mask]
    pred = pred[mask]
    if len(true) == 0:
        return {"n": 0, "R2": np.nan, "RMSE": np.nan, "MAE": np.nan}
    residual = true - pred
    rmse = float(np.sqrt(np.mean(residual**2)))
    mae = float(np.mean(np.abs(residual)))
    denominator = float(np.sum((true - np.mean(true)) ** 2))
    r2 = 1.0 - float(np.sum(residual**2)) / denominator if denominator > 0 else np.nan
    return {"n": int(len(true)), "R2": r2, "RMSE": rmse, "MAE": mae}


def robust_spearman(x: Sequence[float], y: Sequence[float]) -> float:
    x_arr = np.asarray(x, dtype=float)
    y_arr = np.asarray(y, dtype=float)
    mask = np.isfinite(x_arr) & np.isfinite(y_arr)
    x_arr = x_arr[mask]
    y_arr = y_arr[mask]
    if len(x_arr) < 2 or len(np.unique(x_arr)) < 2 or len(np.unique(y_arr)) < 2:
        return np.nan
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        result = spearmanr(x_arr, y_arr)
    return float(result.statistic) if np.isfinite(result.statistic) else np.nan


def classify_curve(
    grid: np.ndarray,
    response: np.ndarray,
    baseline_prediction_std: float,
) -> Dict[str, Any]:
    mask = np.isfinite(grid) & np.isfinite(response)
    x = np.asarray(grid, dtype=float)[mask]
    y = np.asarray(response, dtype=float)[mask]

    if len(x) < 2:
        return {
            "响应类型": "可计算网格不足，无法分类",
            "Spearman_rho": np.nan,
            "响应幅度": np.nan,
            "标准化响应幅度": np.nan,
            "高低水平响应差": np.nan,
        }

    amplitude = float(np.max(y) - np.min(y))
    denom = baseline_prediction_std if baseline_prediction_std > 1e-12 else np.nan
    normalized = amplitude / denom if np.isfinite(denom) else np.nan
    rho = robust_spearman(x, y)

    order = np.argsort(x)
    high_low = float(y[order[-1]] - y[order[0]])

    zero_threshold = max(
        1e-12,
        baseline_prediction_std * NEGLIGIBLE_NORMALIZED_AMPLITUDE,
    )

    if amplitude <= zero_threshold:
        response_type = "模型响应近零"
    elif len(x) == 2:
        response_type = "两水平比较：上升" if high_low > 0 else "两水平比较：下降"
    elif np.isfinite(rho) and rho >= MONOTONIC_SPEARMAN_THRESHOLD:
        response_type = "总体上升"
    elif np.isfinite(rho) and rho <= -MONOTONIC_SPEARMAN_THRESHOLD:
        response_type = "总体下降"
    else:
        response_type = "非单调/平台型"

    return {
        "响应类型": response_type,
        "Spearman_rho": rho,
        "响应幅度": amplitude,
        "标准化响应幅度": normalized,
        "高低水平响应差": high_low,
    }


def target_sort_key(series: pd.Series) -> pd.Series:
    order = {name: i for i, name in enumerate(TARGET_ORDER)}
    return series.astype(str).map(order).fillna(len(order)).astype(int)


def sort_table_by_target(
    frame: pd.DataFrame,
    extra_columns: Optional[List[str]] = None,
) -> pd.DataFrame:
    if frame is None or frame.empty or "目标指标" not in frame.columns:
        return frame
    out = frame.copy()
    out["_目标顺序"] = target_sort_key(out["目标指标"])
    columns = ["_目标顺序"] + (extra_columns or [])
    ascending = [True] * len(columns)
    out = out.sort_values(columns, ascending=ascending, kind="mergesort")
    return out.drop(columns=["_目标顺序"]).reset_index(drop=True)


def pair_cell_text(var1: str, x: float, var2: str, y: float) -> str:
    return f"{var1}={float(x):.6g}；{var2}={float(y):.6g}"


def atomic_write_text(path: Path, text: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_suffix(path.suffix + ".part")
    temporary.write_text(text, encoding="utf-8-sig")
    os.replace(temporary, path)


def sanitize_excel_value(value: Any) -> Any:
    if isinstance(value, str):
        return re.sub(r"[\x00-\x08\x0B\x0C\x0E-\x1F]", "", value)
    return value


def sanitize_frame(frame: pd.DataFrame) -> pd.DataFrame:
    out = frame.copy()
    for col in out.select_dtypes(include=["object", "string"]).columns:
        out[col] = out[col].map(sanitize_excel_value)
    return out


def atomic_write_excel(path: Path, tables: Mapping[str, pd.DataFrame]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_name(path.stem + ".part.xlsx")
    if temporary.exists():
        temporary.unlink()

    try:
        import xlsxwriter  # noqa: F401
        engine = "xlsxwriter"
    except Exception:
        engine = "openpyxl"

    with pd.ExcelWriter(
        temporary,
        engine=engine,
        date_format="yyyy-mm-dd",
        datetime_format="yyyy-mm-dd",
    ) as writer:
        for sheet_name, frame in tables.items():
            clean = sanitize_frame(frame if frame is not None else pd.DataFrame())
            if clean.empty:
                clean = pd.DataFrame([{"说明": "本工作表当前没有记录。"}])

            clean.to_excel(writer, sheet_name=sheet_name[:31], index=False)

            if engine == "xlsxwriter":
                workbook = writer.book
                worksheet = writer.sheets[sheet_name[:31]]

                header_format = workbook.add_format(
                    {
                        "bold": True,
                        "bg_color": "#DCE6F1",
                        "border": 1,
                        "align": "center",
                        "valign": "vcenter",
                        "text_wrap": True,
                    }
                )
                numeric_format = workbook.add_format({"num_format": "0.000000"})
                date_format = workbook.add_format({"num_format": "yyyy-mm-dd"})

                for column_index, column in enumerate(clean.columns):
                    worksheet.write(0, column_index, column, header_format)
                    sample = clean[column].head(1000).astype(str)
                    width = int(sample.map(len).quantile(0.95)) + 2 if len(sample) else 10
                    width = min(max(width, len(str(column)) + 2, 10), 44)

                    column_text = str(column)
                    if "Date" in column_text or "日期" in column_text:
                        worksheet.set_column(column_index, column_index, width, date_format)
                    elif any(
                        token in column_text
                        for token in [
                            "预测", "响应", "比例", "幅度", "指数", "RMSE", "R2",
                            "MAE", "分位", "均值", "中位数", "次数", "频率",
                            "Spearman", "标准差", "水平",
                        ]
                    ):
                        worksheet.set_column(column_index, column_index, width, numeric_format)
                    else:
                        worksheet.set_column(column_index, column_index, width)

                worksheet.freeze_panes(1, 0)
                worksheet.autofilter(
                    0,
                    0,
                    len(clean),
                    max(len(clean.columns) - 1, 0),
                )

    os.replace(temporary, path)


# =============================================================================
# 3. 动态数据、冻结模型、一维结果和 SHAP+PI 读取核查
# =============================================================================

def interpolate_short_internal_gaps(
    series: pd.Series,
    sample_ids: pd.Series,
    max_gap: int = MAX_INTERPOLATION_GAP,
) -> pd.Series:
    """
    与最终模型固化脚本保持一致：
    仅对原样本序号跨度不超过 max_gap 的内部连续缺口做线性插值。
    """
    output = pd.to_numeric(series, errors="coerce").copy()
    values = output.to_numpy(dtype=float)
    coordinates = pd.to_numeric(sample_ids, errors="coerce").to_numpy(dtype=float)

    n = len(values)
    i = 0

    while i < n:
        if np.isfinite(values[i]):
            i += 1
            continue

        start = i
        while i < n and not np.isfinite(values[i]):
            i += 1

        end = i - 1
        left = start - 1
        right = i

        if (
            left >= 0
            and right < n
            and np.isfinite(values[left])
            and np.isfinite(values[right])
            and np.isfinite(coordinates[left])
            and np.isfinite(coordinates[right])
        ):
            original_gap_span = coordinates[right] - coordinates[left] - 1.0
            coordinate_span = coordinates[right] - coordinates[left]

            if 0 <= original_gap_span <= max_gap and coordinate_span > 0:
                left_value = float(values[left])
                right_value = float(values[right])

                for index in range(start, end + 1):
                    fraction = (
                        coordinates[index] - coordinates[left]
                    ) / coordinate_span
                    values[index] = (
                        left_value
                        + fraction * (right_value - left_value)
                    )

    return pd.Series(values, index=output.index, name=output.name)


def load_model_package(target_name: str) -> Tuple[Path, Dict[str, Any]]:
    path = MODEL_DIR / TARGET_CONFIG[target_name]["joblib"]
    if not path.exists():
        raise FileNotFoundError(f"未找到 {target_name} 最终模型：{path}")

    package = joblib.load(path)
    if not isinstance(package, dict):
        raise TypeError(f"{target_name} joblib 不是字典型最终模型包。")

    required = {
        "package_type",
        "target_name",
        "target_column",
        "input_sheet",
        "selected_feature_group",
        "feature_names",
        "algorithm",
        "candidate_id",
        "parameters",
        "preprocessing",
        "scaler",
        "model",
        "source_validation_metrics",
        "full_fit_metrics",
        "training_data",
    }
    missing = sorted(required - set(package.keys()))
    if missing:
        raise KeyError(f"{target_name} 模型包缺少字段：{missing}")

    expected = EXPECTED_FINAL_CONFIG[target_name]

    checks = [
        str(package["target_name"]) == target_name,
        str(package["selected_feature_group"]) == expected["feature_group"],
        str(package["algorithm"]) == expected["algorithm"],
        str(package["candidate_id"]) == expected["candidate_id"],
        params_equal(package["parameters"], expected["parameters"]),
        len(package["feature_names"]) == 12,
        "Q" not in package["feature_names"],
    ]

    if STRICT_MODE and not all(checks):
        raise RuntimeError(
            f"{target_name} 模型包与锁定最终配置不一致："
            f"F={package['selected_feature_group']}，"
            f"算法={package['algorithm']}，"
            f"ID={package['candidate_id']}，"
            f"参数={package['parameters']}"
        )

    return path, package


def read_target_data(target_name: str) -> pd.DataFrame:
    if not INPUT_FILE.exists():
        raise FileNotFoundError(f"未找到动态数据：{INPUT_FILE}")

    cfg = TARGET_CONFIG[target_name]

    data = pd.read_excel(
        INPUT_FILE,
        sheet_name=cfg["sheet"],
        engine="openpyxl",
        na_values=list(TEXTUAL_MISSING_VALUES),
        keep_default_na=True,
    )
    data.columns = [clean_column_name(col) for col in data.columns]

    required = (
        ["样本序号", "Date"]
        + RAW_FEATURES
        + cfg["history"]
        + [cfg["target_col"]]
    )

    missing = [col for col in required if col not in data.columns]
    if missing:
        raise ValueError(f"{target_name} 动态表缺少字段：{missing}")

    data["Date"] = pd.to_datetime(data["Date"], errors="coerce")

    for col in ["样本序号"] + RAW_FEATURES + cfg["history"] + [cfg["target_col"]]:
        data[col] = convert_numeric(data[col])

    data = data.sort_values(
        ["样本序号", "Date"],
        kind="mergesort",
    ).reset_index(drop=True)

    if data["样本序号"].isna().any() or data["Date"].isna().any():
        raise ValueError(f"{target_name} 存在无效样本序号或日期。")

    if data["样本序号"].duplicated().any():
        raise ValueError(f"{target_name} 样本序号存在重复。")

    return data


def rebuild_training_data(
    target_name: str,
    raw_data: pd.DataFrame,
    package: Mapping[str, Any],
) -> Tuple[pd.DataFrame, pd.DataFrame]:
    data = raw_data.copy(deep=True)
    saved_medians = package["preprocessing"].get("training_medians", {})
    audit_rows: List[Dict[str, Any]] = []

    for col in INFLUENT_FEATURES:
        original = convert_numeric(data[col]).astype(float)
        current_median = float(original.median(skipna=True))
        saved_median = float(saved_medians[col])
        median_diff = abs(current_median - saved_median)

        if STRICT_MODE and median_diff > MEDIAN_TOL:
            raise RuntimeError(
                f"{target_name} {col} 当前动态数据中位数与模型包不一致："
                f"当前={current_median}，保存={saved_median}。"
            )

        interpolated = interpolate_short_internal_gaps(
            original,
            data["样本序号"],
            int(
                package["preprocessing"].get(
                    "max_interpolation_gap",
                    MAX_INTERPOLATION_GAP,
                )
            ),
        )

        linear_filled = int(original.isna().sum() - interpolated.isna().sum())
        median_filled = int(interpolated.isna().sum())

        data[col] = interpolated.fillna(saved_median).astype(float)

        audit_rows.append(
            {
                "目标指标": target_name,
                "变量": col,
                "原缺失数": int(original.isna().sum()),
                "短内部插值数": linear_filled,
                "中位数兜底数": median_filled,
                "模型包训练中位数": saved_median,
                "当前原始中位数": current_median,
                "中位数绝对差": median_diff,
            }
        )

    feature_names = list(package["feature_names"])
    missing_after = {
        col: int(data[col].isna().sum())
        for col in feature_names
        if data[col].isna().any()
    }

    if missing_after:
        raise RuntimeError(
            f"{target_name} 重建最终模型输入后仍有缺失：{missing_after}"
        )

    expected_ids = [
        int(x)
        for x in package["training_data"].get("sample_ids", [])
    ]
    current_ids = data["样本序号"].astype(int).tolist()

    if STRICT_MODE and expected_ids and expected_ids != current_ids:
        raise RuntimeError(
            f"{target_name} 动态表样本ID与模型包训练样本ID不一致。"
        )

    return data, pd.DataFrame(audit_rows)


def validate_full_fit(
    target_name: str,
    data: pd.DataFrame,
    package: Mapping[str, Any],
) -> Tuple[np.ndarray, Dict[str, float], Dict[str, float]]:
    feature_names = list(package["feature_names"])

    x = data[feature_names].to_numpy(dtype=float)
    y = data[TARGET_CONFIG[target_name]["target_col"]].to_numpy(dtype=float)

    transformed = package["scaler"].transform(x)
    pred = np.asarray(
        package["model"].predict(transformed),
        dtype=float,
    )

    metrics = calc_metrics(y, pred)
    saved = package.get("full_fit_metrics", {})

    diffs = {
        key: abs(float(metrics[key]) - float(saved[key]))
        for key in ["R2", "RMSE", "MAE"]
        if key in saved
        and np.isfinite(metrics[key])
        and np.isfinite(float(saved[key]))
    }

    if STRICT_MODE and diffs and max(diffs.values()) > ROUNDTRIP_METRIC_TOL:
        raise RuntimeError(
            f"{target_name} 全量拟合指标复现失败：{diffs}"
        )

    return pred, metrics, diffs


def parse_grid_json(value: Any) -> np.ndarray:
    if isinstance(value, str):
        raw = json.loads(value)
    elif isinstance(value, (list, tuple, np.ndarray)):
        raw = value
    else:
        raise ValueError(f"无法解析网格值：{value!r}")

    grid = np.asarray([float(x) for x in raw], dtype=float)
    grid = np.unique(grid[np.isfinite(grid)])
    grid.sort()

    if len(grid) == 0:
        raise ValueError("读取到的响应网格为空。")

    return grid


def read_one_dimensional_results(
    dynamic_hash: str,
    shap_hash: str,
) -> Dict[str, pd.DataFrame]:
    if not ONE_D_FILE.exists():
        raise FileNotFoundError(
            f"未找到最新版一维响应结果：{ONE_D_FILE}"
        )

    required_sheets = {
        "运行设置与模型核查",
        "响应网格定义",
        "响应方向与幅度",
        "后续二维候选",
    }

    with pd.ExcelFile(ONE_D_FILE, engine="openpyxl") as excel:
        missing = required_sheets - set(excel.sheet_names)
        if missing:
            raise ValueError(
                f"一维响应结果缺少工作表：{sorted(missing)}"
            )

        tables = {
            sheet: pd.read_excel(excel, sheet_name=sheet)
            for sheet in required_sheets
        }

    settings = tables["运行设置与模型核查"].copy()

    required_settings_columns = {
        "项目",
        "值",
        "记录类型",
        "目标指标",
        "最终特征组",
        "算法",
        "候选参数ID",
        "核查状态",
    }
    missing_settings_columns = required_settings_columns - set(settings.columns)
    if missing_settings_columns:
        raise ValueError(
            "一维‘运行设置与模型核查’缺少字段："
            f"{sorted(missing_settings_columns)}"
        )

    setting_rows = settings.loc[
        settings["记录类型"].astype(str) == "运行设置"
    ].copy()
    setting_map = {
        str(row["项目"]): row["值"]
        for _, row in setting_rows.iterrows()
        if pd.notna(row.get("项目"))
    }

    if STRICT_MODE:
        one_d_dynamic_hash = str(setting_map.get("动态数据SHA256", "")).strip()
        if one_d_dynamic_hash and one_d_dynamic_hash != dynamic_hash:
            raise RuntimeError(
                "一维结果记录的动态数据SHA256与当前动态数据不一致。"
            )

        one_d_shap_hash = str(setting_map.get("SHAP+PI结果SHA256", "")).strip()
        if one_d_shap_hash and one_d_shap_hash != shap_hash:
            raise RuntimeError(
                "一维结果记录的SHAP+PI结果SHA256与当前文件不一致。"
            )

        one_d_zero_threshold = safe_float(
            setting_map.get("响应近零标准化幅度阈值")
        )
        if np.isfinite(one_d_zero_threshold) and not np.isclose(
            one_d_zero_threshold,
            NEGLIGIBLE_NORMALIZED_AMPLITUDE,
            rtol=1e-12,
            atol=1e-12,
        ):
            raise RuntimeError(
                "一维结果的响应近零阈值与二维脚本不一致。"
            )

        one_d_rho_threshold = safe_float(
            setting_map.get("单调Spearman阈值")
        )
        if np.isfinite(one_d_rho_threshold) and not np.isclose(
            one_d_rho_threshold,
            MONOTONIC_SPEARMAN_THRESHOLD,
            rtol=1e-12,
            atol=1e-12,
        ):
            raise RuntimeError(
                "一维结果的Spearman阈值与二维脚本不一致。"
            )

    model_rows = settings.loc[
        settings["记录类型"].astype(str) == "模型核查"
    ].copy()

    for target_name in TARGET_ORDER:
        expected = EXPECTED_FINAL_CONFIG[target_name]
        sub = model_rows.loc[
            model_rows["目标指标"].astype(str) == target_name
        ]

        if sub.empty:
            raise RuntimeError(
                f"一维结果缺少 {target_name} 模型核查记录。"
            )

        row = sub.iloc[0]

        if STRICT_MODE:
            if str(row["最终特征组"]) != expected["feature_group"]:
                raise RuntimeError(
                    f"一维结果中 {target_name} 最终特征组不一致。"
                )
            if str(row["算法"]) != expected["algorithm"]:
                raise RuntimeError(
                    f"一维结果中 {target_name} 算法不一致。"
                )
            if str(row["候选参数ID"]) != expected["candidate_id"]:
                raise RuntimeError(
                    f"一维结果中 {target_name} 候选参数ID不一致。"
                )
            if str(row["核查状态"]) != "通过":
                raise RuntimeError(
                    f"一维结果中 {target_name} 模型核查未通过。"
                )

    grid_table = tables["响应网格定义"].copy()
    required_grid_columns = {
        "目标指标",
        "变量",
        "网格类型",
        "网格值JSON",
    }
    missing_grid_columns = required_grid_columns - set(grid_table.columns)
    if missing_grid_columns:
        raise ValueError(
            f"一维‘响应网格定义’缺少字段：{sorted(missing_grid_columns)}"
        )

    response_table = tables["响应方向与幅度"].copy()
    required_response_columns = {
        "目标指标",
        "变量",
        "证据等级",
        "响应类型",
        "标准化响应幅度",
    }
    missing_response_columns = (
        required_response_columns - set(response_table.columns)
    )
    if missing_response_columns:
        raise ValueError(
            "一维‘响应方向与幅度’缺少字段："
            f"{sorted(missing_response_columns)}"
        )

    candidate_table = tables["后续二维候选"].copy()
    required_candidate_columns = {
        "目标指标",
        "变量1",
        "变量2",
        "二维候选优先级",
        "一维前置条件结论",
        "预定义原因",
    }
    missing_candidate_columns = (
        required_candidate_columns - set(candidate_table.columns)
    )
    if missing_candidate_columns:
        raise ValueError(
            "一维‘后续二维候选’缺少字段："
            f"{sorted(missing_candidate_columns)}"
        )

    return tables


def validate_shap_pi_file() -> pd.DataFrame:
    if not SHAP_PI_FILE.exists():
        raise FileNotFoundError(
            f"未找到 SHAP+PI 结果：{SHAP_PI_FILE}"
        )

    with pd.ExcelFile(SHAP_PI_FILE, engine="openpyxl") as excel:
        if "分析与模型核查" not in excel.sheet_names:
            raise ValueError(
                "SHAP+PI结果缺少‘分析与模型核查’工作表。"
            )
        overview = pd.read_excel(
            excel,
            sheet_name="分析与模型核查",
        )

    for target_name in TARGET_ORDER:
        expected = EXPECTED_FINAL_CONFIG[target_name]
        sub = overview.loc[
            overview["目标指标"].astype(str) == target_name
        ]

        if sub.empty:
            raise RuntimeError(
                f"SHAP+PI结果缺少 {target_name} 模型核查记录。"
            )

        row = sub.iloc[0]

        if STRICT_MODE:
            if str(row.get("最终特征组")) != expected["feature_group"]:
                raise RuntimeError(
                    f"SHAP+PI中 {target_name} 最终特征组不一致。"
                )
            if str(row.get("算法")) != expected["algorithm"]:
                raise RuntimeError(
                    f"SHAP+PI中 {target_name} 算法不一致。"
                )
            if str(row.get("候选参数ID")) != expected["candidate_id"]:
                raise RuntimeError(
                    f"SHAP+PI中 {target_name} 候选参数ID不一致。"
                )

    return overview


# =============================================================================
# 4. 二维候选、网格与物理检查
# =============================================================================

def build_grid_map(
    grid_table: pd.DataFrame,
) -> Dict[Tuple[str, str], np.ndarray]:
    result: Dict[Tuple[str, str], np.ndarray] = {}

    for _, row in grid_table.iterrows():
        target_name = str(row["目标指标"])
        variable = str(row["变量"])

        if (
            target_name in TARGET_ORDER
            and variable in CONTROLLABLE_FEATURES
        ):
            result[(target_name, variable)] = parse_grid_json(
                row["网格值JSON"]
            )

    return result


def select_pairs(
    candidate_table: pd.DataFrame,
) -> pd.DataFrame:
    table = candidate_table.copy()

    allowed_priority = {"高", "中"}
    if INCLUDE_EXPLORATORY:
        allowed_priority.add("探索")

    selected = table.loc[
        table["二维候选优先级"].astype(str).isin(allowed_priority)
    ].copy()

    selected = selected.loc[
        ~selected["一维前置条件结论"]
        .astype(str)
        .str.contains("不进入|不足|缺少", regex=True)
    ].copy()

    if selected.empty:
        raise RuntimeError(
            "一维结果中没有符合当前设置的二维候选组合。"
        )

    for _, row in selected.iterrows():
        target_name = str(row["目标指标"])
        var1 = str(row["变量1"])
        var2 = str(row["变量2"])

        if target_name not in TARGET_ORDER:
            raise RuntimeError(
                f"二维候选出现未知目标：{target_name}"
            )

        if (
            var1 not in CONTROLLABLE_FEATURES
            or var2 not in CONTROLLABLE_FEATURES
        ):
            raise RuntimeError(
                f"二维候选包含非运行变量：{target_name} / {var1} × {var2}"
            )

        if var1 == var2:
            raise RuntimeError(
                f"二维候选两个变量相同：{target_name} / {var1}"
            )

    duplicate = selected.duplicated(
        ["目标指标", "变量1", "变量2"]
    )
    if duplicate.any():
        bad = selected.loc[
            duplicate,
            ["目标指标", "变量1", "变量2"],
        ].to_dict("records")
        raise RuntimeError(
            f"二维候选存在重复组合：{bad}"
        )

    selected = selected.reset_index(drop=True)
    selected.insert(
        0,
        "二维组合序号",
        np.arange(1, len(selected) + 1, dtype=int),
    )
    return selected


def one_d_variable_info(
    response_table: pd.DataFrame,
    target_name: str,
    variable: str,
) -> Dict[str, Any]:
    sub = response_table.loc[
        (response_table["目标指标"].astype(str) == target_name)
        & (response_table["变量"].astype(str) == variable)
    ]

    if sub.empty:
        return {}

    row = sub.iloc[0]

    keys = [
        "证据等级",
        "响应类型",
        "标准化响应幅度",
        "Spearman_rho",
        "ICE与总体方向一致比例百分比",
        "SHAP排名",
        "SHAP重要性占比百分比",
        "Time_PI可评估性",
        "Time_PI证据结论",
        "Expanding_PI可评估性",
        "Expanding_PI证据结论",
        "PDP_ICE解释建议",
    ]

    return {
        key: row.get(key)
        for key in keys
        if key in row.index
    }


def physical_valid_mask(data: pd.DataFrame) -> np.ndarray:
    valid = np.ones(len(data), dtype=bool)

    for col in CONTROLLABLE_FEATURES:
        values = pd.to_numeric(
            data[col],
            errors="coerce",
        ).to_numpy(dtype=float)
        valid &= np.isfinite(values)

    do = data["Aerobic_DO"].to_numpy(dtype=float)
    feed = data["Anaerobic_Feed_Fraction"].to_numpy(dtype=float)
    hrt = data["HRT"].to_numpy(dtype=float)
    anaer = data["Anaerobic_Volume_Fraction"].to_numpy(dtype=float)
    anoxic = data["Anoxic_Volume_Fraction"].to_numpy(dtype=float)

    valid &= do >= 0.0
    valid &= (feed >= 0.0) & (feed <= 1.0)
    valid &= hrt > 0.0
    valid &= (anaer >= 0.0) & (anaer <= 1.0)
    valid &= (anoxic >= 0.0) & (anoxic <= 1.0)
    valid &= anaer + anoxic < 1.0 - 1e-12

    return valid


def count_exact_pair(
    data: pd.DataFrame,
    var1: str,
    level1: float,
    var2: str,
    level2: float,
) -> int:
    first = np.isclose(
        data[var1].to_numpy(dtype=float),
        float(level1),
        rtol=0.0,
        atol=FLOAT_LEVEL_ATOL,
    )
    second = np.isclose(
        data[var2].to_numpy(dtype=float),
        float(level2),
        rtol=0.0,
        atol=FLOAT_LEVEL_ATOL,
    )
    return int(np.sum(first & second))


def observed_pair_frequency(
    target_name: str,
    var1: str,
    var2: str,
    data: pd.DataFrame,
    baseline_predictions: np.ndarray,
    grid1: np.ndarray,
    grid2: np.ndarray,
) -> pd.DataFrame:
    temp = pd.DataFrame(
        {
            var1: data[var1].to_numpy(dtype=float),
            var2: data[var2].to_numpy(dtype=float),
            "Date": data["Date"].to_numpy(),
            "实测目标": data[
                TARGET_CONFIG[target_name]["target_col"]
            ].to_numpy(dtype=float),
            "冻结模型原始状态预测": baseline_predictions,
        }
    )

    rows: List[Dict[str, Any]] = []
    n = len(temp)

    for (level1, level2), group in temp.groupby(
        [var1, var2],
        dropna=False,
        sort=True,
    ):
        level1 = float(level1)
        level2 = float(level2)

        in_grid1 = bool(
            np.any(
                np.isclose(
                    grid1,
                    level1,
                    rtol=0.0,
                    atol=FLOAT_LEVEL_ATOL,
                )
            )
        )
        in_grid2 = bool(
            np.any(
                np.isclose(
                    grid2,
                    level2,
                    rtol=0.0,
                    atol=FLOAT_LEVEL_ATOL,
                )
            )
        )

        rows.append(
            {
                "目标指标": target_name,
                "变量1": var1,
                "变量2": var2,
                "变量1历史水平": level1,
                "变量2历史水平": level2,
                "历史精确二元组合出现次数": int(len(group)),
                "历史精确二元组合出现比例百分比": (
                    100.0 * len(group) / max(n, 1)
                ),
                "是否同时位于本次二维网格": (
                    "是" if in_grid1 and in_grid2 else "否"
                ),
                "首次日期": group["Date"].min(),
                "末次日期": group["Date"].max(),
                "实测目标均值": float(group["实测目标"].mean()),
                "实测目标中位数": float(group["实测目标"].median()),
                "原始状态模型预测均值": float(
                    group["冻结模型原始状态预测"].mean()
                ),
            }
        )

    return pd.DataFrame(rows)


# =============================================================================
# 5. 二维表面、切片、非加性和 Bootstrap
# =============================================================================

def nonadditivity_index(matrix: np.ndarray) -> float:
    values = np.asarray(matrix, dtype=float)

    if (
        values.ndim != 2
        or values.shape[0] < 2
        or values.shape[1] < 2
        or not np.isfinite(values).all()
    ):
        return np.nan

    grand = float(np.mean(values))
    row_effect = np.mean(values, axis=1) - grand
    col_effect = np.mean(values, axis=0) - grand

    additive = (
        grand
        + row_effect[:, None]
        + col_effect[None, :]
    )

    residual = values - additive
    rms_residual = float(
        np.sqrt(np.mean(residual**2))
    )

    sd = float(np.std(values, ddof=1)) if values.size > 1 else 0.0

    if sd <= 1e-12:
        return 0.0 if rms_residual <= 1e-12 else np.nan

    return rms_residual / sd


def nonadditivity_label(value: float) -> str:
    if not np.isfinite(value):
        return "不可计算"
    if value < 0.10:
        return "接近可加"
    if value < 0.30:
        return "轻度非加性"
    if value < 0.50:
        return "中等非加性"
    return "较强非加性"


def nonadditivity_bootstrap_interpretation(
    original_value: float,
    bootstrap_low: float,
    bootstrap_high: float,
) -> str:
    """根据Bootstrap 95%经验区间判断非加性分类是否跨越人为分类界线。

    这里的Bootstrap只反映历史参考样本组成变化下的经验稳定性，
    因此文字中使用“经验区间”，不将其表述为模型参数置信区间。
    """
    values = [original_value, bootstrap_low, bootstrap_high]
    if not all(np.isfinite(v) for v in values):
        return "Bootstrap有效结果不足，暂不评价非加性程度分类的稳定性。"

    original_label = nonadditivity_label(original_value)
    low_label = nonadditivity_label(bootstrap_low)
    high_label = nonadditivity_label(bootstrap_high)

    label_order = [
        "接近可加",
        "轻度非加性",
        "中等非加性",
        "较强非加性",
    ]
    low_index = label_order.index(low_label)
    high_index = label_order.index(high_label)

    if low_label == high_label:
        if original_label == low_label:
            return (
                f"Bootstrap 95%经验区间未跨越非加性分类界线，"
                f"{original_label}的判断较稳定。"
            )
        return (
            f"Bootstrap 95%经验区间整体处于{low_label}范围，"
            f"但原始二维表面指数对应{original_label}；"
            "建议同时报告原始指数和Bootstrap经验区间。"
        )

    range_phrase_map = {
        (0, 1): "接近可加至轻度非加性",
        (1, 2): "轻度至中等非加性",
        (2, 3): "中等至较强非加性",
    }
    phrase = range_phrase_map.get((low_index, high_index))
    if phrase is not None:
        return (
            f"Bootstrap 95%经验区间跨越{low_label}与{high_label}的分类界线，"
            f"建议按{phrase}解释，而不只依据单一指数贴标签。"
        )

    return (
        f"Bootstrap 95%经验区间从{low_label}延伸至{high_label}，"
        "跨越多个非加性分类界线；建议直接结合原始指数和Bootstrap经验区间解释。"
    )


def build_slice_table(
    target_name: str,
    var1: str,
    var2: str,
    grid1: np.ndarray,
    grid2: np.ndarray,
    centered_surface: np.ndarray,
    baseline_sd: float,
) -> pd.DataFrame:
    rows: List[Dict[str, Any]] = []

    for j, fixed_level in enumerate(grid2):
        values = centered_surface[:, j]
        finite = np.isfinite(values)

        metrics = classify_curve(
            grid1[finite],
            values[finite],
            baseline_sd,
        )

        rows.append(
            {
                "目标指标": target_name,
                "变量1": var1,
                "变量2": var2,
                "切片方向": f"固定{var2}，扫描{var1}",
                "扫描变量": var1,
                "固定变量": var2,
                "固定水平": float(fixed_level),
                "总扫描网格点数": int(len(grid1)),
                "可计算扫描网格点数": int(np.sum(finite)),
                **metrics,
            }
        )

    for i, fixed_level in enumerate(grid1):
        values = centered_surface[i, :]
        finite = np.isfinite(values)

        metrics = classify_curve(
            grid2[finite],
            values[finite],
            baseline_sd,
        )

        rows.append(
            {
                "目标指标": target_name,
                "变量1": var1,
                "变量2": var2,
                "切片方向": f"固定{var1}，扫描{var2}",
                "扫描变量": var2,
                "固定变量": var1,
                "固定水平": float(fixed_level),
                "总扫描网格点数": int(len(grid2)),
                "可计算扫描网格点数": int(np.sum(finite)),
                **metrics,
            }
        )

    return pd.DataFrame(rows)


def bootstrap_surface(
    target_name: str,
    var1: str,
    var2: str,
    grid1: np.ndarray,
    grid2: np.ndarray,
    pred_matrix: np.ndarray,
    delta_matrix: np.ndarray,
    centered_surface_flat: np.ndarray,
    baseline_prediction_std: float,
) -> Tuple[pd.DataFrame, pd.DataFrame, Dict[str, Any]]:
    rng = np.random.default_rng(
        (
            RANDOM_SEED
            + stable_seed(
                f"{target_name}-{var1}-{var2}-2d-bootstrap"
            )
        )
        % (2**32 - 1)
    )

    n, cell_count = delta_matrix.shape

    boot_delta = np.full(
        (BOOTSTRAP_N, cell_count),
        np.nan,
        dtype=float,
    )
    boot_pred = np.full(
        (BOOTSTRAP_N, cell_count),
        np.nan,
        dtype=float,
    )

    detail_rows: List[Dict[str, Any]] = []

    for bootstrap_index in range(BOOTSTRAP_N):
        sampled = rng.choice(
            np.arange(n),
            size=n,
            replace=True,
        )

        for cell in range(cell_count):
            delta_values = delta_matrix[sampled, cell]
            delta_values = delta_values[
                np.isfinite(delta_values)
            ]

            pred_values = pred_matrix[sampled, cell]
            pred_values = pred_values[
                np.isfinite(pred_values)
            ]

            if len(delta_values):
                boot_delta[bootstrap_index, cell] = float(
                    np.mean(delta_values)
                )
            if len(pred_values):
                boot_pred[bootstrap_index, cell] = float(
                    np.mean(pred_values)
                )

        current = boot_delta[bootstrap_index]
        finite = np.isfinite(current)

        valid_values = current[finite]
        amplitude = (
            float(np.max(valid_values) - np.min(valid_values))
            if len(valid_values) >= 2
            else np.nan
        )

        comparable = (
            finite
            & np.isfinite(centered_surface_flat)
        )

        surface_spearman = (
            robust_spearman(
                centered_surface_flat[comparable],
                current[comparable],
            )
            if int(np.sum(comparable)) >= 4
            else np.nan
        )

        minimum_cell: Optional[int] = None
        maximum_cell: Optional[int] = None

        valid_indices = np.where(finite)[0]
        if len(valid_indices):
            minimum_cell = int(
                valid_indices[
                    np.argmin(current[valid_indices])
                ]
            )
            maximum_cell = int(
                valid_indices[
                    np.argmax(current[valid_indices])
                ]
            )

        current_matrix = current.reshape(
            len(grid1),
            len(grid2),
        )
        nonadd = nonadditivity_index(current_matrix)

        standardized_amplitude = (
            amplitude / baseline_prediction_std
            if np.isfinite(amplitude)
            and baseline_prediction_std > 1e-12
            else np.nan
        )

        if minimum_cell is not None:
            min_i = minimum_cell // len(grid2)
            min_j = minimum_cell % len(grid2)
            minimum_cell_text = pair_cell_text(
                var1,
                grid1[min_i],
                var2,
                grid2[min_j],
            )
        else:
            minimum_cell_text = "不适用"

        if maximum_cell is not None:
            max_i = maximum_cell // len(grid2)
            max_j = maximum_cell % len(grid2)
            maximum_cell_text = pair_cell_text(
                var1,
                grid1[max_i],
                var2,
                grid2[max_j],
            )
        else:
            maximum_cell_text = "不适用"

        detail_rows.append(
            {
                "目标指标": target_name,
                "变量1": var1,
                "变量2": var2,
                "Bootstrap编号": bootstrap_index + 1,
                "可计算二维单元数": int(np.sum(finite)),
                "二维表面响应幅度": amplitude,
                "二维表面标准化响应幅度": standardized_amplitude,
                "与原始二维表面Spearman": surface_spearman,
                "最低中心化响应网格": minimum_cell_text,
                "最高中心化响应网格": maximum_cell_text,
                "二维表面非加性指数": nonadd,
            }
        )

    point_rows: List[Dict[str, Any]] = []

    for i, level1 in enumerate(grid1):
        for j, level2 in enumerate(grid2):
            cell = i * len(grid2) + j

            delta_values = boot_delta[:, cell]
            delta_values = delta_values[
                np.isfinite(delta_values)
            ]

            pred_values = boot_pred[:, cell]
            pred_values = pred_values[
                np.isfinite(pred_values)
            ]

            point_rows.append(
                {
                    "目标指标": target_name,
                    "变量1": var1,
                    "变量2": var2,
                    "变量1水平": float(level1),
                    "变量2水平": float(level2),
                    "Bootstrap有效次数": int(
                        min(
                            len(delta_values),
                            len(pred_values),
                        )
                    ),
                    "平均预测值Bootstrap中位数": (
                        float(np.median(pred_values))
                        if len(pred_values)
                        else np.nan
                    ),
                    "平均预测值Bootstrap2.5%分位": (
                        float(np.quantile(pred_values, 0.025))
                        if len(pred_values)
                        else np.nan
                    ),
                    "平均预测值Bootstrap97.5%分位": (
                        float(np.quantile(pred_values, 0.975))
                        if len(pred_values)
                        else np.nan
                    ),
                    "中心化响应Bootstrap中位数": (
                        float(np.median(delta_values))
                        if len(delta_values)
                        else np.nan
                    ),
                    "中心化响应Bootstrap2.5%分位": (
                        float(np.quantile(delta_values, 0.025))
                        if len(delta_values)
                        else np.nan
                    ),
                    "中心化响应Bootstrap97.5%分位": (
                        float(np.quantile(delta_values, 0.975))
                        if len(delta_values)
                        else np.nan
                    ),
                }
            )

    detail = pd.DataFrame(detail_rows)

    valid_detail = detail.loc[
        detail["可计算二维单元数"] >= 2
    ].copy()

    minimum_counter: Counter[str] = Counter(
        str(value)
        for value in valid_detail["最低中心化响应网格"]
        if str(value) != "不适用"
    )
    maximum_counter: Counter[str] = Counter(
        str(value)
        for value in valid_detail["最高中心化响应网格"]
        if str(value) != "不适用"
    )

    minimum_text = (
        minimum_counter.most_common(1)[0][0]
        if minimum_counter
        else "不适用"
    )
    minimum_frequency = (
        100.0
        * minimum_counter.most_common(1)[0][1]
        / max(len(valid_detail), 1)
        if minimum_counter
        else np.nan
    )

    maximum_text = (
        maximum_counter.most_common(1)[0][0]
        if maximum_counter
        else "不适用"
    )
    maximum_frequency = (
        100.0
        * maximum_counter.most_common(1)[0][1]
        / max(len(valid_detail), 1)
        if maximum_counter
        else np.nan
    )

    def quantiles(column: str) -> Tuple[float, float, float]:
        values = pd.to_numeric(
            valid_detail[column],
            errors="coerce",
        ).dropna()

        if len(values) == 0:
            return np.nan, np.nan, np.nan

        return (
            float(values.quantile(0.5)),
            float(values.quantile(0.025)),
            float(values.quantile(0.975)),
        )

    amp_median, amp_low, amp_high = quantiles(
        "二维表面响应幅度"
    )
    std_amp_median, std_amp_low, std_amp_high = quantiles(
        "二维表面标准化响应幅度"
    )
    rho_median, rho_low, rho_high = quantiles(
        "与原始二维表面Spearman"
    )
    nonadd_median, nonadd_low, nonadd_high = quantiles(
        "二维表面非加性指数"
    )

    summary = {
        "Bootstrap次数": BOOTSTRAP_N,
        "Bootstrap有效表面次数": int(len(valid_detail)),
        "Bootstrap有效比例百分比": (
            100.0 * len(valid_detail) / max(BOOTSTRAP_N, 1)
        ),
        "二维表面响应幅度Bootstrap中位数": amp_median,
        "二维表面响应幅度Bootstrap2.5%分位": amp_low,
        "二维表面响应幅度Bootstrap97.5%分位": amp_high,
        "二维表面标准化响应幅度Bootstrap中位数": std_amp_median,
        "二维表面标准化响应幅度Bootstrap2.5%分位": std_amp_low,
        "二维表面标准化响应幅度Bootstrap97.5%分位": std_amp_high,
        "表面Spearman_Bootstrap中位数": rho_median,
        "表面Spearman_Bootstrap2.5%分位": rho_low,
        "表面Spearman_Bootstrap97.5%分位": rho_high,
        "非加性指数Bootstrap中位数": nonadd_median,
        "非加性指数Bootstrap2.5%分位": nonadd_low,
        "非加性指数Bootstrap97.5%分位": nonadd_high,
        "Bootstrap最常见最低中心化响应网格": minimum_text,
        "Bootstrap最低响应网格出现比例百分比": minimum_frequency,
        "Bootstrap最常见最高中心化响应网格": maximum_text,
        "Bootstrap最高响应网格出现比例百分比": maximum_frequency,
    }

    return pd.DataFrame(point_rows), detail, summary


def compute_pair_response(
    target_name: str,
    var1: str,
    var2: str,
    grid1: np.ndarray,
    grid2: np.ndarray,
    data: pd.DataFrame,
    package: Mapping[str, Any],
    baseline_predictions: np.ndarray,
    candidate_row: Mapping[str, Any],
    one_d_response_table: pd.DataFrame,
) -> Tuple[
    pd.DataFrame,
    pd.DataFrame,
    Dict[str, Any],
    pd.DataFrame,
    pd.DataFrame,
    pd.DataFrame,
]:
    n = len(data)
    cell_count = len(grid1) * len(grid2)

    baseline_sd = (
        float(np.std(baseline_predictions, ddof=1))
        if len(baseline_predictions) > 1
        else 0.0
    )

    pred_matrix = np.full(
        (n, cell_count),
        np.nan,
        dtype=float,
    )
    delta_matrix = np.full(
        (n, cell_count),
        np.nan,
        dtype=float,
    )
    physical_matrix = np.zeros(
        (n, cell_count),
        dtype=bool,
    )

    surface_rows: List[Dict[str, Any]] = []

    cell = 0

    for i, level1 in enumerate(grid1):
        for j, level2 in enumerate(grid2):
            scenarios = data.copy(deep=True)
            scenarios[var1] = float(level1)
            scenarios[var2] = float(level2)

            physical = physical_valid_mask(scenarios)
            physical_matrix[:, cell] = physical

            if np.any(physical):
                matrix = scenarios.loc[
                    physical,
                    package["feature_names"],
                ].to_numpy(dtype=float)

                transformed = package["scaler"].transform(matrix)

                predictions = np.asarray(
                    package["model"].predict(transformed),
                    dtype=float,
                )

                pred_matrix[physical, cell] = predictions
                delta_matrix[physical, cell] = (
                    predictions
                    - baseline_predictions[physical]
                )

            valid_pred = pred_matrix[:, cell]
            valid_delta = delta_matrix[:, cell]

            pred_values = valid_pred[np.isfinite(valid_pred)]
            delta_values = valid_delta[np.isfinite(valid_delta)]

            physical_count = int(np.sum(physical))
            exact_count = count_exact_pair(
                data,
                var1,
                float(level1),
                var2,
                float(level2),
            )

            surface_rows.append(
                {
                    "目标指标": target_name,
                    "变量1": var1,
                    "变量2": var2,
                    "二维候选优先级": candidate_row.get(
                        "二维候选优先级",
                        "",
                    ),
                    "变量1网格序号": i + 1,
                    "变量2网格序号": j + 1,
                    "变量1水平": float(level1),
                    "变量2水平": float(level2),
                    "总参考样本数": n,
                    "物理有效参考样本数": physical_count,
                    "物理无效参考样本数": n - physical_count,
                    "物理有效比例百分比": (
                        100.0 * physical_count / max(n, 1)
                    ),
                    "历史精确二元组合出现次数": exact_count,
                    "历史精确二元组合出现比例百分比": (
                        100.0 * exact_count / max(n, 1)
                    ),
                    "历史是否出现过该精确二元组合": (
                        "是" if exact_count > 0 else "否"
                    ),
                    "二维PDP平均预测值": (
                        float(np.mean(pred_values))
                        if len(pred_values)
                        else np.nan
                    ),
                    "中心化二维PDP响应": (
                        float(np.mean(delta_values))
                        if len(delta_values)
                        else np.nan
                    ),
                }
            )

            cell += 1

    surface = pd.DataFrame(surface_rows)

    centered_surface_flat = surface[
        "中心化二维PDP响应"
    ].to_numpy(dtype=float)

    prediction_surface_flat = surface[
        "二维PDP平均预测值"
    ].to_numpy(dtype=float)

    centered_surface = centered_surface_flat.reshape(
        len(grid1),
        len(grid2),
    )

    prediction_surface = prediction_surface_flat.reshape(
        len(grid1),
        len(grid2),
    )

    bootstrap_points, bootstrap_detail, bootstrap_summary = (
        bootstrap_surface(
            target_name,
            var1,
            var2,
            grid1,
            grid2,
            pred_matrix,
            delta_matrix,
            centered_surface_flat,
            baseline_sd,
        )
    )

    surface = surface.merge(
        bootstrap_points,
        on=[
            "目标指标",
            "变量1",
            "变量2",
            "变量1水平",
            "变量2水平",
        ],
        how="left",
        validate="one_to_one",
    )

    slices = build_slice_table(
        target_name,
        var1,
        var2,
        grid1,
        grid2,
        centered_surface,
        baseline_sd,
    )

    historical_frequency = observed_pair_frequency(
        target_name,
        var1,
        var2,
        data,
        baseline_predictions,
        grid1,
        grid2,
    )

    finite_centered = np.isfinite(centered_surface_flat)
    finite_prediction = np.isfinite(prediction_surface_flat)

    centered_values = centered_surface_flat[finite_centered]
    prediction_values = prediction_surface_flat[finite_prediction]

    response_amplitude = (
        float(np.max(centered_values) - np.min(centered_values))
        if len(centered_values) >= 2
        else np.nan
    )

    standardized_amplitude = (
        response_amplitude / baseline_sd
        if np.isfinite(response_amplitude) and baseline_sd > 1e-12
        else np.nan
    )

    if len(centered_values):
        valid_indices = np.where(finite_centered)[0]
        min_centered_cell = int(
            valid_indices[
                np.argmin(centered_surface_flat[valid_indices])
            ]
        )
        max_centered_cell = int(
            valid_indices[
                np.argmax(centered_surface_flat[valid_indices])
            ]
        )
    else:
        min_centered_cell = -1
        max_centered_cell = -1

    if len(prediction_values):
        valid_indices = np.where(finite_prediction)[0]
        min_prediction_cell = int(
            valid_indices[
                np.argmin(prediction_surface_flat[valid_indices])
            ]
        )
        max_prediction_cell = int(
            valid_indices[
                np.argmax(prediction_surface_flat[valid_indices])
            ]
        )
    else:
        min_prediction_cell = -1
        max_prediction_cell = -1

    def index_to_text(cell_index: int) -> str:
        if cell_index < 0:
            return "不适用"
        i = cell_index // len(grid2)
        j = cell_index % len(grid2)
        return pair_cell_text(
            var1,
            grid1[i],
            var2,
            grid2[j],
        )

    nonadd = nonadditivity_index(centered_surface)

    # 将原始二维表面的非加性指数与Bootstrap 95%经验区间结合解释。
    # 若经验区间跨越人为分类界线，文字说明不强行给出单一等级。
    bootstrap_summary = dict(bootstrap_summary)
    bootstrap_summary["Bootstrap下的非加性解释"] = (
        nonadditivity_bootstrap_interpretation(
            nonadd,
            safe_float(bootstrap_summary.get("非加性指数Bootstrap2.5%分位")),
            safe_float(bootstrap_summary.get("非加性指数Bootstrap97.5%分位")),
        )
    )

    variable1_info = one_d_variable_info(
        one_d_response_table,
        target_name,
        var1,
    )
    variable2_info = one_d_variable_info(
        one_d_response_table,
        target_name,
        var2,
    )

    physical_percentages = pd.to_numeric(
        surface["物理有效比例百分比"],
        errors="coerce",
    ).dropna()

    summary: Dict[str, Any] = {
        "目标指标": target_name,
        "变量1": var1,
        "变量2": var2,
        "二维候选优先级": candidate_row.get(
            "二维候选优先级",
            "",
        ),
        "一维前置条件结论": candidate_row.get(
            "一维前置条件结论",
            "",
        ),
        "预定义原因": candidate_row.get(
            "预定义原因",
            "",
        ),
        "变量1一维证据等级": variable1_info.get(
            "证据等级",
            "",
        ),
        "变量1一维响应类型": variable1_info.get(
            "响应类型",
            "",
        ),
        "变量1一维标准化响应幅度": safe_float(
            variable1_info.get("标准化响应幅度")
        ),
        "变量2一维证据等级": variable2_info.get(
            "证据等级",
            "",
        ),
        "变量2一维响应类型": variable2_info.get(
            "响应类型",
            "",
        ),
        "变量2一维标准化响应幅度": safe_float(
            variable2_info.get("标准化响应幅度")
        ),
        "变量1网格点数": int(len(grid1)),
        "变量2网格点数": int(len(grid2)),
        "二维网格单元总数": int(cell_count),
        "可计算二维单元数": int(np.sum(finite_centered)),
        "历史至少出现过一次的二维网格单元数": int(
            (
                surface["历史精确二元组合出现次数"]
                .astype(float)
                > 0
            ).sum()
        ),
        "物理有效比例最小值百分比": (
            float(physical_percentages.min())
            if len(physical_percentages)
            else np.nan
        ),
        "物理有效比例中位数百分比": (
            float(physical_percentages.median())
            if len(physical_percentages)
            else np.nan
        ),
        "物理有效比例最大值百分比": (
            float(physical_percentages.max())
            if len(physical_percentages)
            else np.nan
        ),
        "基线预测标准差": baseline_sd,
        "二维中心化表面响应幅度": response_amplitude,
        "二维中心化表面标准化幅度": standardized_amplitude,
        "最低中心化响应网格": index_to_text(
            min_centered_cell
        ),
        "最高中心化响应网格": index_to_text(
            max_centered_cell
        ),
        "最低平均预测值网格": index_to_text(
            min_prediction_cell
        ),
        "最高平均预测值网格": index_to_text(
            max_prediction_cell
        ),
        "二维表面非加性指数": nonadd,
        "二维表面非加性判断": nonadditivity_label(nonadd),
        **bootstrap_summary,
    }

    return (
        surface,
        slices,
        summary,
        bootstrap_points,
        bootstrap_detail,
        historical_frequency,
    )


# =============================================================================
# 6. 文字说明
# =============================================================================

def build_text_report(
    candidate_table: pd.DataFrame,
    model_checks: pd.DataFrame,
    pair_summary: pd.DataFrame,
    warning_table: pd.DataFrame,
    elapsed_minutes: float,
) -> str:
    lines: List[str] = []

    lines.append("A²/O 四指标二维 PDP 条件响应分析说明")
    lines.append("=" * 112)
    lines.append("")

    lines.append("一、任务与输出")
    lines.append(f"1. 完成时间：{now_str()}")
    lines.append(f"2. 总运行耗时：{elapsed_minutes:.2f} 分钟。")
    lines.append(f"3. 动态数据：{INPUT_FILE}")
    lines.append(f"4. 最终模型目录：{MODEL_DIR}")
    lines.append(f"5. SHAP + PI结果：{SHAP_PI_FILE}")
    lines.append(f"6. 一维响应结果：{ONE_D_FILE}")
    lines.append(f"7. 输出Excel：{OUTPUT_EXCEL}")
    lines.append(f"8. 输出TXT：{OUTPUT_TEXT}")
    lines.append("9. 本阶段不生成任何图片。")
    lines.append("")

    lines.append("二、二维响应分析方法")
    lines.append(
        "1. 二维候选直接读取最新版一维结果中的“后续二维候选”，"
        "不重新人为增加变量组合。"
    )
    lines.append(
        "2. 每个变量的二维网格完全复用一维“响应网格定义”，"
        "保证一维和二维分析范围一致。"
    )
    lines.append(
        "3. 对每个二维网格单元，同时改变两个运行变量；"
        "其余进水水质、结构条件和对应目标的lag1/lag2/roll3"
        "保持每条历史样本原值。"
    )
    lines.append(
        "4. 每个网格单元使用全部基本物理条件有效的历史参考样本计算经典二维PDP；"
        "同时保存绝对平均预测值和相对于各样本原始状态预测的中心化平均变化。"
    )
    lines.append(
        "5. 历史精确二元组合出现次数只作为解释背景和结果审计信息，"
        "不作为二维网格过滤门槛。"
    )
    lines.append(
        "6. 二维切片沿用一维阶段的响应分类方法："
        f"响应幅度不超过基线预测标准差的{NEGLIGIBLE_NORMALIZED_AMPLITUDE*100:g}%"
        "记为模型响应近零；"
        f"Spearman |rho|>={MONOTONIC_SPEARMAN_THRESHOLD:g}"
        "时才标记总体上升或下降。"
    )
    lines.append(
        "7. 二维表面非加性指数用于描述冻结模型二维表面偏离简单可加结构的程度，"
        "不解释为实验因果交互；同时结合Bootstrap 95%经验区间判断分类是否跨界，"
        "若跨越分类界线则采用范围式自然语言解释。"
    )
    lines.append(
        f"8. Bootstrap共{BOOTSTRAP_N}次，只对历史参考样本组成进行有放回重采样，"
        "不重新训练模型；所给区间为参考样本组成的经验稳定性区间。"
    )
    lines.append(
        "9. 所有二维结果均为一步条件预测响应，不代表运行条件改变后的长期稳态结果。"
    )
    lines.append("")

    lines.append("三、一维阶段进入二维的组合")
    for _, row in candidate_table.iterrows():
        lines.append(
            f"- {row['目标指标']}："
            f"{row['变量1']} × {row['变量2']}；"
            f"优先级={row['二维候选优先级']}；"
            f"{row['一维前置条件结论']}"
        )
    lines.append("")

    lines.append("四、模型与数据复现核查")
    for _, row in model_checks.iterrows():
        lines.append(
            f"- {row['目标指标']}："
            f"{row['最终特征组']} + {row['算法']}（{row['候选参数ID']}），"
            f"样本数={int(row['样本数'])}，特征数={int(row['特征数'])}；"
            f"全量拟合复现最大指标绝对差="
            f"{fmt(row['全量拟合复现最大指标绝对差'], 12)}；"
            f"动态数据SHA256匹配={row['动态数据SHA256匹配']}；"
            f"核查={row['核查状态']}。"
        )
    lines.append("")

    lines.append("五、二维组合结果摘要")
    for _, row in pair_summary.iterrows():
        lines.append(
            f"- {row['目标指标']}：{row['变量1']} × {row['变量2']}；"
            f"二维单元={int(row['可计算二维单元数'])}/"
            f"{int(row['二维网格单元总数'])}；"
            f"历史精确出现过的网格单元="
            f"{int(row['历史至少出现过一次的二维网格单元数'])}；"
            f"物理有效比例最低="
            f"{fmt(row['物理有效比例最小值百分比'], 1)}%；"
            f"标准化表面幅度="
            f"{fmt(row['二维中心化表面标准化幅度'], 3)}；"
            f"非加性={row['二维表面非加性判断']}"
            f"（指数={fmt(row['二维表面非加性指数'], 3)}）；"
            f"Bootstrap表面Spearman中位数="
            f"{fmt(row['表面Spearman_Bootstrap中位数'], 3)}。"
        )
        lines.append(
            f"  Bootstrap下的非加性解释："
            f"{row.get('Bootstrap下的非加性解释', '不适用')}"
        )
        lines.append(
            f"  最低中心化响应网格：{row['最低中心化响应网格']}；"
            f"Bootstrap最常见最低网格："
            f"{row['Bootstrap最常见最低中心化响应网格']}"
            f"（{fmt(row['Bootstrap最低响应网格出现比例百分比'], 1)}%）。"
        )
        lines.append(
            f"  最高中心化响应网格：{row['最高中心化响应网格']}；"
            f"Bootstrap最常见最高网格："
            f"{row['Bootstrap最常见最高中心化响应网格']}"
            f"（{fmt(row['Bootstrap最高响应网格出现比例百分比'], 1)}%）。"
        )
    lines.append("")

    lines.append("六、解释边界")
    lines.append(
        "1. 二维PDP展示的是冻结模型在历史变量范围内对两个运行变量同时变化的条件响应。"
    )
    lines.append(
        "2. 历史精确二元组合出现次数为0的网格仍可能得到模型预测；"
        "这种结果应表述为模型条件预测，而不是历史反应器已经直接验证过的工况。"
    )
    lines.append(
        "3. 二维表面非加性、切片方向和Bootstrap稳定性都属于模型解释结果，"
        "不能单独证明真实工艺中的因果作用或长期稳态交互。"
    )
    lines.append(
        "4. 后续正向预测和反向筛选应继续使用冻结模型本身；"
        "二维PDP结果不用于重新训练模型或修改模型参数。"
    )
    lines.append("")

    lines.append("七、Excel工作表说明")
    lines.append(
        "1. ‘运行设置与模型核查’：记录输入文件、版本、参数和冻结模型复现结果。"
    )
    lines.append(
        "2. ‘二维候选组合’：保存从一维结果读取的正式二维组合及一维前置解释。"
    )
    lines.append(
        "3. ‘二维PDP表面’：逐二维网格单元保存平均预测、中心化响应、"
        "物理有效样本数、历史精确组合出现次数和Bootstrap区间。"
    )
    lines.append(
        "4. ‘二维切片’：分别固定一个变量，查看另一个变量的模型响应方向和幅度。"
    )
    lines.append(
        "5. ‘二维组合摘要’：汇总表面幅度、非加性、Bootstrap下的非加性解释、极值网格和Bootstrap稳定性。"
    )
    lines.append(
        "6. ‘Bootstrap表面明细’：保存每次参考样本Bootstrap的表面幅度、"
        "表面Spearman、极值位置和非加性指数。"
    )
    lines.append(
        "7. ‘历史二元组合频次’：统计历史数据中实际出现过的二元运行组合及出现次数。"
    )
    lines.append(
        "8. ‘插补重建核查’：核对二维分析数据重建与冻结模型固化时的预处理是否一致。"
    )
    lines.append("")

    lines.append("八、完整性与警告")
    lines.append(
        f"1. 本次运行警告/错误记录数：{len(warning_table)}。"
    )
    lines.append(
        "2. 严格模式下，如模型配置、动态数据哈希、一维结果、SHAP解释对象"
        "或全量拟合复现结果不一致，脚本会停止，不生成正式分析结果。"
    )
    lines.append(
        "3. 本脚本没有重新训练模型、没有修改joblib、没有输出图片。"
    )

    return "\n".join(lines)


# =============================================================================
# 7. 主流程
# =============================================================================

def main() -> None:
    overall_start = time.time()
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    log("A²/O 四指标二维 PDP 条件响应分析开始。")
    log(f"动态数据：{INPUT_FILE}")
    log(f"最终模型目录：{MODEL_DIR}")
    log(f"SHAP + PI结果：{SHAP_PI_FILE}")
    log(f"一维响应结果：{ONE_D_FILE}")
    log(f"输出目录：{OUTPUT_DIR}")
    log(
        f"Bootstrap次数={BOOTSTRAP_N}；"
        f"IncludeExploratory={INCLUDE_EXPLORATORY}；"
        f"SmokeTest={SMOKE_TEST}；DryRun={DRY_RUN}"
    )

    warning_rows: List[Dict[str, Any]] = []

    try:
        for path, label in [
            (INPUT_FILE, "动态数据"),
            (SHAP_PI_FILE, "SHAP+PI结果"),
            (ONE_D_FILE, "一维响应结果"),
        ]:
            if not path.exists():
                raise FileNotFoundError(
                    f"未找到{label}：{path}"
                )

        dynamic_hash = sha256_file(INPUT_FILE)
        shap_hash = sha256_file(SHAP_PI_FILE)
        one_d_hash = sha256_file(ONE_D_FILE)

        validate_shap_pi_file()

        one_d_tables = read_one_dimensional_results(
            dynamic_hash,
            shap_hash,
        )

        grid_map = build_grid_map(
            one_d_tables["响应网格定义"]
        )
        candidate_table = select_pairs(
            one_d_tables["后续二维候选"]
        )
        one_d_response_table = one_d_tables[
            "响应方向与幅度"
        ]

        for _, row in candidate_table.iterrows():
            target_name = str(row["目标指标"])
            var1 = str(row["变量1"])
            var2 = str(row["变量2"])

            if (target_name, var1) not in grid_map:
                raise RuntimeError(
                    f"一维结果缺少 {target_name}-{var1} 网格。"
                )
            if (target_name, var2) not in grid_map:
                raise RuntimeError(
                    f"一维结果缺少 {target_name}-{var2} 网格。"
                )

        setting_rows: List[Dict[str, Any]] = [
            {"项目": "完成时间", "值": now_str()},
            {"项目": "分析版本", "值": ANALYSIS_VERSION},
            {"项目": "动态数据", "值": str(INPUT_FILE)},
            {"项目": "动态数据SHA256", "值": dynamic_hash},
            {"项目": "最终模型目录", "值": str(MODEL_DIR)},
            {"项目": "SHAP+PI结果", "值": str(SHAP_PI_FILE)},
            {"项目": "SHAP+PI结果SHA256", "值": shap_hash},
            {"项目": "一维响应结果", "值": str(ONE_D_FILE)},
            {"项目": "一维响应结果SHA256", "值": one_d_hash},
            {"项目": "输出目录", "值": str(OUTPUT_DIR)},
            {"项目": "参考样本Bootstrap次数", "值": BOOTSTRAP_N},
            {
                "项目": "二维候选来源",
                "值": "直接读取一维结果“后续二维候选”",
            },
            {
                "项目": "二维网格来源",
                "值": "完全复用一维结果“响应网格定义”",
            },
            {
                "项目": "二维PDP计算",
                "值": "每个网格单元使用全部基本物理条件有效的历史参考样本",
            },
            {
                "项目": "中心化基准",
                "值": "同一历史样本原始状态的冻结模型预测值",
            },
            {
                "项目": "历史精确二元组合出现次数",
                "值": "仅作为解释和审计信息，不作为网格过滤条件",
            },
            {
                "项目": "基本物理条件",
                "值": (
                    "DO>=0；0<=Feed<=1；HRT>0；"
                    "两个体积分数均在0-1之间；"
                    "厌氧区体积分数+缺氧区体积分数<1"
                ),
            },
            {
                "项目": "响应近零标准化幅度阈值",
                "值": NEGLIGIBLE_NORMALIZED_AMPLITUDE,
            },
            {
                "项目": "单调Spearman阈值",
                "值": MONOTONIC_SPEARMAN_THRESHOLD,
            },
            {
                "项目": "是否纳入探索优先级候选",
                "值": "是" if INCLUDE_EXPLORATORY else "否",
            },
            {"项目": "Smoke test", "值": "是" if SMOKE_TEST else "否"},
            {"项目": "Dry run", "值": "是" if DRY_RUN else "否"},
            {"项目": "严格核查", "值": "是" if STRICT_MODE else "否"},
        ]

        model_check_rows: List[Dict[str, Any]] = []
        imputation_tables: List[pd.DataFrame] = []
        surface_tables: List[pd.DataFrame] = []
        slice_tables: List[pd.DataFrame] = []
        pair_summary_rows: List[Dict[str, Any]] = []
        bootstrap_detail_tables: List[pd.DataFrame] = []
        historical_frequency_tables: List[pd.DataFrame] = []

        if not DRY_RUN:
            target_to_pairs: Dict[str, pd.DataFrame] = {
                target_name: candidate_table.loc[
                    candidate_table["目标指标"].astype(str)
                    == target_name
                ].copy()
                for target_name in TARGET_ORDER
            }

            for target_name in TARGET_ORDER:
                target_pairs = target_to_pairs[target_name]
                if target_pairs.empty:
                    continue

                target_start = time.time()
                log(
                    f"开始 {target_name} 二维分析，"
                    f"组合数={len(target_pairs)}。"
                )

                model_path, package = load_model_package(
                    target_name
                )
                raw = read_target_data(target_name)
                data, imputation = rebuild_training_data(
                    target_name,
                    raw,
                    package,
                )

                (
                    baseline_predictions,
                    full_metrics,
                    metric_diffs,
                ) = validate_full_fit(
                    target_name,
                    data,
                    package,
                )

                saved_dynamic_hash = str(
                    package["training_data"].get(
                        "dynamic_input_file_sha256",
                        "",
                    )
                ).strip()

                dynamic_hash_match = (
                    saved_dynamic_hash == dynamic_hash
                    if saved_dynamic_hash
                    else False
                )

                if STRICT_MODE and not dynamic_hash_match:
                    raise RuntimeError(
                        f"{target_name} 模型包记录的动态数据SHA256"
                        "与当前动态数据不一致。"
                    )

                imputation_tables.append(imputation)

                for _, candidate_row in target_pairs.iterrows():
                    var1 = str(candidate_row["变量1"])
                    var2 = str(candidate_row["变量2"])
                    grid1 = grid_map[(target_name, var1)]
                    grid2 = grid_map[(target_name, var2)]

                    log(
                        f"{target_name}：计算 "
                        f"{var1} × {var2} 二维PDP。"
                    )

                    (
                        surface,
                        slices,
                        pair_summary,
                        _bootstrap_points,
                        bootstrap_detail,
                        historical_frequency,
                    ) = compute_pair_response(
                        target_name,
                        var1,
                        var2,
                        grid1,
                        grid2,
                        data,
                        package,
                        baseline_predictions,
                        candidate_row,
                        one_d_response_table,
                    )

                    combination_index = int(
                        candidate_row["二维组合序号"]
                    )

                    for current_frame in [
                        surface,
                        slices,
                        bootstrap_detail,
                        historical_frequency,
                    ]:
                        current_frame.insert(
                            1,
                            "二维组合序号",
                            combination_index,
                        )

                    pair_summary["二维组合序号"] = (
                        combination_index
                    )

                    surface_tables.append(surface)
                    slice_tables.append(slices)
                    pair_summary_rows.append(pair_summary)
                    bootstrap_detail_tables.append(
                        bootstrap_detail
                    )
                    historical_frequency_tables.append(
                        historical_frequency
                    )

                model_check_rows.append(
                    {
                        "目标指标": target_name,
                        "最终特征组": package["selected_feature_group"],
                        "算法": package["algorithm"],
                        "候选参数ID": package["candidate_id"],
                        "特征数": len(package["feature_names"]),
                        "样本数": len(data),
                        "日期起点": data["Date"].min(),
                        "日期终点": data["Date"].max(),
                        "模型文件": str(model_path),
                        "模型SHA256": sha256_file(model_path),
                        "动态数据SHA256匹配": (
                            "是" if dynamic_hash_match else "否"
                        ),
                        "全量拟合R2_当前复现": full_metrics["R2"],
                        "全量拟合RMSE_当前复现": full_metrics["RMSE"],
                        "全量拟合MAE_当前复现": full_metrics["MAE"],
                        "全量拟合复现最大指标绝对差": (
                            max(metric_diffs.values())
                            if metric_diffs
                            else 0.0
                        ),
                        "二维组合数": len(target_pairs),
                        "核查状态": "通过",
                        "运行耗时分钟": (
                            (time.time() - target_start) / 60.0
                        ),
                    }
                )

        settings_df = pd.DataFrame(setting_rows)
        model_checks_df = pd.DataFrame(model_check_rows)

        # 运行设置与模型核查合并到同一个工作表。
        settings_df["记录类型"] = "运行设置"
        if not model_checks_df.empty:
            model_checks_df["记录类型"] = "模型核查"

        settings_and_checks = pd.concat(
            [settings_df, model_checks_df],
            ignore_index=True,
            sort=False,
        )

        surface_df = (
            pd.concat(surface_tables, ignore_index=True)
            if surface_tables
            else pd.DataFrame()
        )
        slices_df = (
            pd.concat(slice_tables, ignore_index=True)
            if slice_tables
            else pd.DataFrame()
        )
        pair_summary_df = pd.DataFrame(pair_summary_rows)
        bootstrap_detail_df = (
            pd.concat(
                bootstrap_detail_tables,
                ignore_index=True,
            )
            if bootstrap_detail_tables
            else pd.DataFrame()
        )
        historical_frequency_df = (
            pd.concat(
                historical_frequency_tables,
                ignore_index=True,
            )
            if historical_frequency_tables
            else pd.DataFrame()
        )
        imputation_df = (
            pd.concat(imputation_tables, ignore_index=True)
            if imputation_tables
            else pd.DataFrame()
        )
        warning_df = pd.DataFrame(warning_rows)

        candidate_output = candidate_table.copy()

        # 一维候选表中的“提醒”是面向下一阶段的旧说明。进入二维结果后，
        # 改为当前阶段的自然语言说明，避免出现“本脚本未计算二维响应”等过时文字。
        if "提醒" in candidate_output.columns:
            candidate_output = candidate_output.drop(columns=["提醒"])
        candidate_output["二维分析说明"] = (
            "该组合由一维响应分析筛选进入二维条件响应分析；"
            "二维网格沿用一维网格，历史精确二元组合出现次数仅用于解释和审计。"
        )

        # 固定目标顺序；组合内部保留一维候选原始顺序。
        if not candidate_output.empty:
            candidate_output = sort_table_by_target(
                candidate_output,
                ["二维组合序号"],
            )

        if not model_checks_df.empty:
            model_checks_df = sort_table_by_target(
                model_checks_df
            )

        if not surface_df.empty:
            surface_df = sort_table_by_target(
                surface_df,
                [
                    "二维组合序号",
                    "变量1网格序号",
                    "变量2网格序号",
                ],
            )

        if not slices_df.empty:
            slices_df = sort_table_by_target(
                slices_df,
                [
                    "二维组合序号",
                    "切片方向",
                    "固定水平",
                ],
            )

        if not pair_summary_df.empty:
            pair_summary_df = sort_table_by_target(
                pair_summary_df,
                ["二维组合序号"],
            )

        if not bootstrap_detail_df.empty:
            bootstrap_detail_df = sort_table_by_target(
                bootstrap_detail_df,
                ["二维组合序号", "Bootstrap编号"],
            )

        if not historical_frequency_df.empty:
            historical_frequency_df = sort_table_by_target(
                historical_frequency_df,
                [
                    "二维组合序号",
                    "变量1历史水平",
                    "变量2历史水平",
                ],
            )

        if not imputation_df.empty:
            imputation_df = sort_table_by_target(
                imputation_df,
                ["变量"],
            )

        elapsed_minutes = (
            time.time() - overall_start
        ) / 60.0

        tables = {
            "运行设置与模型核查": settings_and_checks,
            "二维候选组合": candidate_output,
            "二维PDP表面": surface_df,
            "二维切片": slices_df,
            "二维组合摘要": pair_summary_df,
            "Bootstrap表面明细": bootstrap_detail_df,
            "历史二元组合频次": historical_frequency_df,
            "插补重建核查": imputation_df,
            "错误与警告": warning_df,
        }

        atomic_write_excel(
            OUTPUT_EXCEL,
            tables,
        )

        text = build_text_report(
            candidate_output,
            model_checks_df,
            pair_summary_df,
            warning_df,
            elapsed_minutes,
        )

        atomic_write_text(
            OUTPUT_TEXT,
            text,
        )

        log("二维响应分析完成。")
        log(f"结果Excel：{OUTPUT_EXCEL}")
        log(f"说明TXT：{OUTPUT_TEXT}")
        log(
            f"二维候选组合数={len(candidate_output)}；"
            f"二维表面记录数={len(surface_df)}；"
            f"总耗时={elapsed_minutes:.2f}分钟。"
        )

    except Exception as exc:
        failure_text = "\n".join(
            [
                "A²/O 四指标二维 PDP 条件响应分析运行失败",
                "=" * 80,
                f"时间：{now_str()}",
                f"分析版本：{ANALYSIS_VERSION}",
                f"动态数据：{INPUT_FILE}",
                f"最终模型目录：{MODEL_DIR}",
                f"SHAP+PI结果：{SHAP_PI_FILE}",
                f"一维响应结果：{ONE_D_FILE}",
                f"输出目录：{OUTPUT_DIR}",
                "",
                f"错误类型：{type(exc).__name__}",
                f"错误信息：{exc}",
                "",
                "完整错误堆栈：",
                traceback.format_exc(),
                "",
                "说明：脚本失败后不会自动修改模型、数据或一维结果。"
                "请根据错误信息核查输入文件、冻结模型和一维响应结果版本。",
            ]
        )

        try:
            atomic_write_text(
                OUTPUT_TEXT,
                failure_text,
            )
        except Exception:
            pass

        temporary_excel = OUTPUT_EXCEL.with_name(
            OUTPUT_EXCEL.stem + ".part.xlsx"
        )
        try:
            if temporary_excel.exists():
                temporary_excel.unlink()
        except Exception:
            pass

        log(
            f"运行失败：{type(exc).__name__}: {exc}"
        )
        raise


if __name__ == "__main__":
    main()
