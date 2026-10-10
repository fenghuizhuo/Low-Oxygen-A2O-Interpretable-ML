# -*- coding: utf-8 -*-
r"""
A²/O 四指标最终模型：一维 PDP / ICE 运行响应分析
================================================================

用途
----
本脚本用于已经完成最终模型终选、全量固化以及 SHAP + Permutation Importance
解释之后的一维运行响应分析。

本阶段只解释 5 个可控运行/结构变量：
1. Aerobic_DO
2. Anaerobic_Feed_Fraction
3. HRT
4. Anaerobic_Volume_Fraction
5. Anoxic_Volume_Fraction

正式分析对象固定为：
- TP       : F2 + SVR          / SVR_032
- TN       : F2 + SVR          / SVR_018
- NH4+-N   : F2 + ElasticNet   / ElasticNet_025
- COD      : F2 + ElasticNet   / ElasticNet_040

重要原则
--------
1. 直接读取已经冻结的四个最终 joblib，不重新选择模型、不重新调参。
2. PDP / ICE 只扰动当前运行变量；进水条件以及 lag1/lag2/roll3 保持每条历史样本原值。
3. HRT、进水分配比例、厌氧/缺氧体积分数使用历史真实观测水平；
   DO 若历史唯一水平较少则使用真实观测水平，若水平较多则使用历史 5%–95% 分位范围内的分位数网格。
4. 每个网格点使用全部物理有效历史参考样本计算经典 PDP；
   中心化 PDP / ICE 以同一样本原始状态预测值为基准，表示改变当前运行变量后的预测变化。
5. 物理约束仅用于排除明显不合理的运行组合：
   Aerobic_DO >= 0；0 <= Anaerobic_Feed_Fraction <= 1；HRT > 0；
   0 <= Anaerobic_Volume_Fraction <= 1；0 <= Anoxic_Volume_Fraction <= 1；
   且 Anaerobic_Volume_Fraction + Anoxic_Volume_Fraction < 1。
6. ICE 保存按时间阶段和目标浓度水平分层抽取的代表样本，不生成图；
   后续论文阶段直接读取 Excel 制图。
7. 100 次 Bootstrap 仅对历史参考样本组成进行重采样，不重新训练模型；
   其区间反映参考样本组成变化下的响应稳定性，不等同于模型参数置信区间。
8. 读取最新 SHAP + PI 的“SHAP_PI单特征对照”，给运行变量附加证据等级；
   SHAP/PI 只用于决定解释优先级，不作为 PDP 数值计算的权重，也不会反向修改最终模型。
9. NH4+-N 若运行变量在最终 ElasticNet 中对应系数为 0 或整体响应接近 0，
    应如实输出“模型响应近零”，不强行构造运行参数效应。
10. 本阶段不计算二维响应面，不计算 H-like，不输出 PNG/PDF/SVG 等任何图件。
11. 所有响应均为冻结预测模型在历史变量范围内的条件响应，
    不是实验干预因果效应，也不是工况改变后的长期稳态机理结论。

默认输入
--------
动态数据：
${PROJECT_ROOT}/outputs/preprocessing/dynamic/动态可用数据与插补比较结果.xlsx

最终模型目录：
${PROJECT_ROOT}/outputs/final_selection/moxingbaocun

SHAP + PI 结果：
${PROJECT_ROOT}/outputs/interpretation/shap_pi/四指标_SHAP与PermutationImportance综合结果.xlsx

默认输出
--------
${PROJECT_ROOT}/outputs/response/one_dimensional

正式只输出：
1. 四指标_一维PDP_ICE响应分析结果.xlsx
2. 四指标_一维PDP_ICE响应分析说明.txt

不会输出图片。

可选环境变量
------------
A2O_1D_INPUT_FILE
A2O_1D_MODEL_DIR
A2O_1D_SHAP_PI_FILE
A2O_1D_OUTPUT_DIR
A2O_1D_BOOTSTRAP_N=100
A2O_1D_ICE_SAMPLE_N=120
A2O_1D_DO_GRID_N=21
A2O_1D_SMOKE_TEST=1
A2O_1D_STRICT=1

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
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, Iterable, List, Mapping, Optional, Sequence, Tuple

import joblib
import numpy as np
import pandas as pd
from scipy.stats import spearmanr
from sklearn.metrics import mean_absolute_error, mean_squared_error, r2_score


# =============================================================================
# 1. 路径、常量与最终配置锁定
# =============================================================================

PROJECT_ROOT = Path(__file__).resolve().parents[1]

INPUT_FILE = Path(
    os.environ.get(
        "A2O_1D_INPUT_FILE",
        str(PROJECT_ROOT / "outputs/preprocessing/dynamic/动态可用数据与插补比较结果.xlsx"),
    )
)

MODEL_DIR = Path(
    os.environ.get(
        "A2O_1D_MODEL_DIR",
        str(PROJECT_ROOT / "outputs/final_selection/moxingbaocun"),
    )
)

SHAP_PI_FILE = Path(
    os.environ.get(
        "A2O_1D_SHAP_PI_FILE",
        str(PROJECT_ROOT / "outputs/interpretation/shap_pi/四指标_SHAP与PermutationImportance综合结果.xlsx"),
    )
)

OUTPUT_DIR = Path(
    os.environ.get(
        "A2O_1D_OUTPUT_DIR",
        str(PROJECT_ROOT / "outputs/response/one_dimensional"),
    )
)

OUTPUT_EXCEL = OUTPUT_DIR / "四指标_一维PDP_ICE响应分析结果.xlsx"
OUTPUT_TEXT = OUTPUT_DIR / "四指标_一维PDP_ICE响应分析说明.txt"

RANDOM_SEED = 20260620
STRICT_MODE = os.environ.get("A2O_1D_STRICT", "1").strip() != "0"
SMOKE_TEST = os.environ.get("A2O_1D_SMOKE_TEST", "0").strip() == "1"

BOOTSTRAP_N = int(os.environ.get("A2O_1D_BOOTSTRAP_N", "100"))
ICE_SAMPLE_N = int(os.environ.get("A2O_1D_ICE_SAMPLE_N", "120"))
DO_GRID_N = int(os.environ.get("A2O_1D_DO_GRID_N", "21"))

if SMOKE_TEST:
    BOOTSTRAP_N = min(BOOTSTRAP_N, 5)
    ICE_SAMPLE_N = min(ICE_SAMPLE_N, 24)
    DO_GRID_N = min(DO_GRID_N, 7)

MAX_INTERPOLATION_GAP = 7
ROUNDTRIP_METRIC_TOL = 1e-9
MEDIAN_TOL = 1e-10


# 响应分类阈值。
MONOTONIC_SPEARMAN_THRESHOLD = 0.80
ICE_DIRECTION_RHO_THRESHOLD = 0.70
NEGLIGIBLE_NORMALIZED_AMPLITUDE = 0.02  # 响应范围 <= 基线预测SD的2%，记为模型响应近零。

# 网格。
# HRT、Feed和两个体积分数始终使用历史真实观测水平；
# DO唯一水平较少时使用真实观测水平，较多时使用历史5%–95%分位范围内的分位数网格。
AUTO_DISCRETE_UNIQUE_MAX = 10

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

# 这四类必须按真实历史水平比较；DO根据唯一值数量自动决定。
FORCE_OBSERVED_LEVELS = {
    "Anaerobic_Feed_Fraction",
    "HRT",
    "Anaerobic_Volume_Fraction",
    "Anoxic_Volume_Fraction",
}

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

# 二维阶段只做候选清单，本脚本不计算二维表面。
PREDEFINED_2D_PAIRS: Dict[str, List[Tuple[str, str, str]]] = {
    "TP": [
        ("HRT", "Anaerobic_Volume_Fraction", "HRT与厌氧区体积分数均在TP的一维解释中具有分析价值，适合进一步查看二维条件响应。"),
        ("Aerobic_DO", "HRT", "DO×HRT是典型曝气-停留时间组合。"),
        ("Aerobic_DO", "Anaerobic_Volume_Fraction", "检查曝气与厌氧结构组合响应。"),
    ],
    "TN": [
        ("Aerobic_DO", "HRT", "TN中DO具有最完整的SHAP/PI证据，HRT为模型响应候选。"),
        ("Aerobic_DO", "Anoxic_Volume_Fraction", "检查曝气与缺氧空间组合。"),
        ("HRT", "Anoxic_Volume_Fraction", "检查停留时间与缺氧空间组合。"),
    ],
    "NH4+-N": [],
    "COD": [
        ("Aerobic_DO", "Anoxic_Volume_Fraction", "COD中缺氧体积分数有SHAP响应，DO作为辅助变量。"),
    ],
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
        while True:
            block = handle.read(1024 * 1024)
            if not block:
                break
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
    mse = float(mean_squared_error(true, pred))
    return {
        "n": int(len(true)),
        "R2": float(r2_score(true, pred)) if len(true) >= 2 and np.std(true) > 0 else np.nan,
        "RMSE": float(math.sqrt(mse)),
        "MAE": float(mean_absolute_error(true, pred)),
    }


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


def atomic_write_text(path: Path, text: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(path.suffix + ".part")
    tmp.write_text(text, encoding="utf-8-sig")
    os.replace(tmp, path)


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
    tmp = path.with_name(path.stem + ".part.xlsx")
    if tmp.exists():
        tmp.unlink()

    try:
        import xlsxwriter  # noqa: F401
        engine = "xlsxwriter"
    except Exception:
        engine = "openpyxl"

    with pd.ExcelWriter(
        tmp,
        engine=engine,
        date_format="yyyy-mm-dd",
        datetime_format="yyyy-mm-dd",
    ) as writer:
        for sheet_name, frame in tables.items():
            if frame is None:
                continue
            clean = sanitize_frame(frame)
            if clean.empty:
                clean = pd.DataFrame([{"说明": "本工作表没有可用记录。"}])
            clean.to_excel(writer, sheet_name=sheet_name[:31], index=False)

            if engine == "xlsxwriter":
                wb = writer.book
                ws = writer.sheets[sheet_name[:31]]
                header_fmt = wb.add_format(
                    {
                        "bold": True,
                        "bg_color": "#DCE6F1",
                        "border": 1,
                        "align": "center",
                        "valign": "vcenter",
                        "text_wrap": True,
                    }
                )
                date_fmt = wb.add_format({"num_format": "yyyy-mm-dd"})
                num_fmt = wb.add_format({"num_format": "0.000000"})
                pct_fmt = wb.add_format({"num_format": "0.00"})

                for j, col in enumerate(clean.columns):
                    ws.write(0, j, col, header_fmt)
                    sample = clean[col].head(500).astype(str)
                    width = max(10, len(str(col)) + 2)
                    if len(sample):
                        width = max(width, int(sample.map(len).quantile(0.95)) + 2)
                    width = min(width, 42)
                    col_text = str(col)
                    if "Date" in col_text or "日期" in col_text:
                        ws.set_column(j, j, width, date_fmt)
                    elif any(
                        token in col_text
                        for token in [
                            "预测", "响应", "幅度", "比例", "距离", "阈值", "均值", "分位",
                            "RMSE", "R2", "MAE", "SHAP", "PI", "rho", "标准差",
                        ]
                    ):
                        ws.set_column(j, j, width, pct_fmt if "百分比" in col_text else num_fmt)
                    else:
                        ws.set_column(j, j, width)
                ws.freeze_panes(1, 0)
                ws.autofilter(0, 0, len(clean), max(len(clean.columns) - 1, 0))

    os.replace(tmp, path)


# =============================================================================
# 3. 模型、数据、SHAP+PI 读取与严格核查
# =============================================================================

def interpolate_short_internal_gaps(
    series: pd.Series,
    sample_ids: pd.Series,
    max_gap: int = MAX_INTERPOLATION_GAP,
) -> pd.Series:
    """与最终模型固化脚本保持一致：按原样本序号跨度判断短内部缺口。"""
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
                    fraction = (coordinates[index] - coordinates[left]) / coordinate_span
                    values[index] = left_value + fraction * (right_value - left_value)

    return pd.Series(values, index=output.index, name=output.name)


def load_model_package(target_name: str) -> Tuple[Path, Dict[str, Any]]:
    cfg = TARGET_CONFIG[target_name]
    path = MODEL_DIR / cfg["joblib"]
    if not path.exists():
        raise FileNotFoundError(f"未找到 {target_name} 最终模型：{path}")

    package = joblib.load(path)
    if not isinstance(package, dict):
        raise TypeError(f"{target_name} joblib不是字典型最终模型包。")

    required = {
        "package_type", "target_name", "target_column", "input_sheet",
        "selected_feature_group", "feature_names", "algorithm", "candidate_id",
        "parameters", "preprocessing", "scaler", "model", "source_validation_metrics",
        "full_fit_metrics", "training_data",
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
            f"F={package['selected_feature_group']}，算法={package['algorithm']}，"
            f"ID={package['candidate_id']}，参数={package['parameters']}"
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

    required = ["样本序号", "Date"] + RAW_FEATURES + cfg["history"] + [cfg["target_col"]]
    missing = [c for c in required if c not in data.columns]
    if missing:
        raise ValueError(f"{target_name} 动态表缺少字段：{missing}")

    data["Date"] = pd.to_datetime(data["Date"], errors="coerce")
    for col in ["样本序号"] + RAW_FEATURES + cfg["history"] + [cfg["target_col"]]:
        data[col] = convert_numeric(data[col])

    data = data.sort_values(["样本序号", "Date"], kind="mergesort").reset_index(drop=True)
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
        calculated_median = float(original.median(skipna=True))
        saved_median = float(saved_medians[col])
        median_diff = abs(calculated_median - saved_median)
        if STRICT_MODE and median_diff > MEDIAN_TOL:
            raise RuntimeError(
                f"{target_name} {col} 当前动态数据中位数与模型包不一致："
                f"当前={calculated_median}，保存={saved_median}。"
            )
        interpolated = interpolate_short_internal_gaps(
            original,
            data["样本序号"],
            int(package["preprocessing"].get("max_interpolation_gap", MAX_INTERPOLATION_GAP)),
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
                "当前原始中位数": calculated_median,
                "中位数绝对差": median_diff,
            }
        )

    feature_names = list(package["feature_names"])
    missing_after = {c: int(data[c].isna().sum()) for c in feature_names if data[c].isna().any()}
    if missing_after:
        raise RuntimeError(f"{target_name} 重建后仍有缺失：{missing_after}")

    expected_ids = [int(x) for x in package["training_data"].get("sample_ids", [])]
    current_ids = data["样本序号"].astype(int).tolist()
    if STRICT_MODE and expected_ids and expected_ids != current_ids:
        raise RuntimeError(f"{target_name} 动态表样本ID与模型包训练样本ID不一致。")

    return data, pd.DataFrame(audit_rows)


def validate_full_fit(
    target_name: str,
    data: pd.DataFrame,
    package: Mapping[str, Any],
) -> Tuple[np.ndarray, Dict[str, float], Dict[str, float]]:
    feature_names = list(package["feature_names"])
    x = data[feature_names].to_numpy(dtype=float)
    y = data[TARGET_CONFIG[target_name]["target_col"]].to_numpy(dtype=float)
    x_scaled = package["scaler"].transform(x)
    pred = np.asarray(package["model"].predict(x_scaled), dtype=float)
    metrics = calc_metrics(y, pred)
    saved = package.get("full_fit_metrics", {})
    diffs = {
        key: abs(float(metrics[key]) - float(saved[key]))
        for key in ["R2", "RMSE", "MAE"]
        if key in saved and np.isfinite(metrics[key]) and np.isfinite(float(saved[key]))
    }
    if STRICT_MODE and diffs and max(diffs.values()) > ROUNDTRIP_METRIC_TOL:
        raise RuntimeError(f"{target_name} 全量拟合指标复现失败：{diffs}")
    return pred, metrics, diffs


def read_shap_pi_evidence() -> Tuple[pd.DataFrame, pd.DataFrame]:
    if not SHAP_PI_FILE.exists():
        raise FileNotFoundError(f"未找到SHAP+PI结果：{SHAP_PI_FILE}")

    with pd.ExcelFile(SHAP_PI_FILE, engine="openpyxl") as excel:
        required_sheets = {"SHAP_PI单特征对照", "分析与模型核查"}
        missing = required_sheets - set(excel.sheet_names)
        if missing:
            raise ValueError(f"SHAP+PI结果缺少工作表：{sorted(missing)}")
        comparison = pd.read_excel(excel, sheet_name="SHAP_PI单特征对照")
        overview = pd.read_excel(excel, sheet_name="分析与模型核查")

    # 严格核对SHAP解释对象与最终模型一致。
    for target in TARGET_ORDER:
        expected = EXPECTED_FINAL_CONFIG[target]
        sub = overview.loc[overview["目标指标"].astype(str) == target]
        if sub.empty:
            raise RuntimeError(f"SHAP+PI结果缺少 {target} 模型核查记录。")
        row = sub.iloc[0]
        if STRICT_MODE:
            if str(row.get("最终特征组")) != expected["feature_group"]:
                raise RuntimeError(f"SHAP+PI中 {target} 特征组不是F2。")
            if str(row.get("算法")) != expected["algorithm"]:
                raise RuntimeError(f"SHAP+PI中 {target} 算法不一致。")
            if str(row.get("候选参数ID")) != expected["candidate_id"]:
                raise RuntimeError(f"SHAP+PI中 {target} 候选参数ID不一致。")

    return comparison, overview


# =============================================================================
# 4. SHAP+PI 运行变量证据等级
# =============================================================================

def has_positive_pi_signal(status: str, text: str) -> bool:
    """只把正式可评估项目中的正PI证据算作证据；"无明确正"不得误判。"""
    status_text = str(status)
    s = str(text)
    if status_text != "可评估":
        return False
    if "无明确正" in s or "未见明确正" in s:
        return False
    return ("稳健正" in s) or ("正重要性信号" in s) or ("正PI" in s)


def has_robust_pi_signal(status: str, text: str) -> bool:
    return str(status) == "可评估" and "稳健正" in str(text)


def evidence_grade(row: pd.Series) -> Tuple[str, str]:
    shap_effective = str(row.get("SHAP是否有效贡献", "")) == "是"
    shap_rank = safe_float(row.get("SHAP排名"))
    shap_top5 = np.isfinite(shap_rank) and shap_rank <= 5

    time_status = str(row.get("Time_PI可评估性", ""))
    exp_status = str(row.get("Expanding_PI可评估性", ""))
    time_conclusion = str(row.get("Time_PI证据结论", ""))
    exp_conclusion = str(row.get("Expanding_PI证据结论", ""))

    robust = has_robust_pi_signal(time_status, time_conclusion) or has_robust_pi_signal(exp_status, exp_conclusion)
    positive = has_positive_pi_signal(time_status, time_conclusion) or has_positive_pi_signal(exp_status, exp_conclusion)
    pi_limited = (
        time_status.startswith("不可评估")
        or time_status.startswith("有限可评估")
        or exp_status.startswith("不可评估")
        or exp_status.startswith("有限可评估")
    )

    if shap_effective and shap_top5 and robust:
        return "A", "SHAP Top5且至少一种时间PI为稳健正信号；可作为一维响应主分析变量。"
    if shap_effective and shap_top5 and (positive or pi_limited):
        return "B", "SHAP Top5，但PI仅部分支持或受时间工况可评估性限制；可重点查看，并结合响应幅度、ICE异质性和PI边界解释。"
    if shap_effective and robust:
        return "B", "SHAP存在有效贡献且时间PI有稳健正信号；可作为次重点响应变量。"
    if shap_effective:
        return "C", "SHAP存在模型贡献，但时间PI支持较弱/不足；仅作辅助模型响应解释。"
    if positive:
        return "C", "SHAP无有效贡献但PI出现正信号；方法间不一致，仅作探索。"
    return "D", "SHAP无有效贡献且无可靠正PI；不作为运行响应主解释。"


def build_operation_evidence(comparison: pd.DataFrame) -> pd.DataFrame:
    rows: List[Dict[str, Any]] = []
    for target in TARGET_ORDER:
        for variable in CONTROLLABLE_FEATURES:
            sub = comparison.loc[
                (comparison["目标指标"].astype(str) == target)
                & (comparison["特征"].astype(str) == variable)
            ]
            if sub.empty:
                if STRICT_MODE:
                    raise RuntimeError(f"SHAP_PI单特征对照缺少 {target}-{variable}。")
                row = pd.Series(dtype=object)
            else:
                row = sub.iloc[0]
            grade, note = evidence_grade(row)
            rows.append(
                {
                    "目标指标": target,
                    "变量": variable,
                    "证据等级": grade,
                    "证据等级说明": note,
                    "SHAP是否有效贡献": row.get("SHAP是否有效贡献", "缺少"),
                    "SHAP排名": safe_float(row.get("SHAP排名")),
                    "SHAP重要性占比百分比": safe_float(row.get("SHAP重要性占比百分比")),
                    "SHAP方向结论": row.get("SHAP方向结论", "缺少"),
                    "Time_PI可评估性": row.get("Time_PI可评估性", "缺少"),
                    "Time_PI证据结论": row.get("Time_PI证据结论", "缺少"),
                    "Time_PI排名": safe_float(row.get("Time_PI排名")),
                    "Time_RMSE增加均值": safe_float(row.get("Time_RMSE增加均值")),
                    "Expanding_PI可评估性": row.get("Expanding_PI可评估性", "缺少"),
                    "Expanding_PI证据结论": row.get("Expanding_PI证据结论", "缺少"),
                    "Expanding_PI排名": safe_float(row.get("Expanding_PI排名")),
                    "Expanding_RMSE增加均值": safe_float(row.get("Expanding_RMSE增加均值")),
                    "是否涉及高相关特征": row.get("是否涉及高相关特征", "缺少"),
                    "SHAP_PI一致性解释": row.get("一致性解释", "缺少"),
                }
            )
    return pd.DataFrame(rows)



# =============================================================================
# 5. 物理约束和响应网格
# =============================================================================

def physical_valid_mask(data: pd.DataFrame) -> np.ndarray:
    """
    检查一维扰动后的运行组合是否满足基本物理条件。

    这里只排除明显不合理的运行组合。
    """
    valid = np.ones(len(data), dtype=bool)

    for col in CONTROLLABLE_FEATURES:
        values = pd.to_numeric(data[col], errors="coerce").to_numpy(dtype=float)
        valid &= np.isfinite(values)

    do_values = pd.to_numeric(data["Aerobic_DO"], errors="coerce").to_numpy(dtype=float)
    feed = pd.to_numeric(data["Anaerobic_Feed_Fraction"], errors="coerce").to_numpy(dtype=float)
    hrt = pd.to_numeric(data["HRT"], errors="coerce").to_numpy(dtype=float)
    anaer = pd.to_numeric(data["Anaerobic_Volume_Fraction"], errors="coerce").to_numpy(dtype=float)
    anoxic = pd.to_numeric(data["Anoxic_Volume_Fraction"], errors="coerce").to_numpy(dtype=float)

    valid &= do_values >= 0.0
    valid &= (feed >= 0.0) & (feed <= 1.0)
    valid &= hrt > 0.0
    valid &= (anaer >= 0.0) & (anaer <= 1.0)
    valid &= (anoxic >= 0.0) & (anoxic <= 1.0)
    valid &= (anaer + anoxic) < 1.0 - 1e-12
    return valid


def build_grid(values: pd.Series, variable: str) -> Tuple[np.ndarray, Dict[str, Any]]:
    """
    构造一维响应网格。

    - HRT、Feed、厌氧/缺氧体积分数：使用历史真实观测水平；
    - DO：唯一水平较少时使用真实观测水平，较多时使用历史5%–95%分位范围内的分位数网格。
    """
    clean = pd.to_numeric(values, errors="coerce").dropna().to_numpy(dtype=float)
    if len(clean) == 0:
        raise ValueError(f"{variable} 没有有效观测值。")

    unique = np.unique(clean)
    unique.sort()

    stats: Dict[str, Any] = {
        "历史最小值": float(np.min(clean)),
        "历史5%分位值": float(np.quantile(clean, 0.05)),
        "历史P25": float(np.quantile(clean, 0.25)),
        "历史P50": float(np.quantile(clean, 0.50)),
        "历史P75": float(np.quantile(clean, 0.75)),
        "历史95%分位值": float(np.quantile(clean, 0.95)),
        "历史最大值": float(np.max(clean)),
        "历史唯一值数": int(len(unique)),
    }

    if variable in FORCE_OBSERVED_LEVELS:
        grid = unique.astype(float)
        grid_type = "历史真实观测水平"
        grid_rule = "使用该变量在当前目标动态数据中的全部真实观测水平"
    elif len(unique) <= AUTO_DISCRETE_UNIQUE_MAX:
        grid = unique.astype(float)
        grid_type = "历史真实观测水平"
        grid_rule = "DO历史唯一水平较少，直接使用全部真实观测水平"
    else:
        probabilities = np.linspace(0.05, 0.95, DO_GRID_N)
        grid = np.unique(np.quantile(clean, probabilities)).astype(float)
        grid_type = "历史5%–95%分位数网格"
        grid_rule = f"DO历史水平较多，在历史5%–95%分位范围内使用{len(grid)}个分位数网格点"

    if len(grid) < 2:
        raise ValueError(f"{variable} 可用于一维响应分析的网格点少于2个。")

    stats["网格类型"] = grid_type
    stats["网格构造说明"] = grid_rule
    stats["网格点数"] = int(len(grid))
    stats["网格值JSON"] = json.dumps([float(x) for x in grid], ensure_ascii=False)
    return grid, stats


def stratified_ice_indices(
    data: pd.DataFrame,
    target_col: str,
    n_samples: int,
    seed_text: str,
) -> np.ndarray:
    """
    按运行时间阶段 × 目标浓度水平分层抽取ICE代表样本。
    """
    n_samples = min(n_samples, len(data))
    if n_samples >= len(data):
        return np.arange(len(data), dtype=int)

    rng = np.random.default_rng((RANDOM_SEED + stable_seed(seed_text)) % (2**32 - 1))
    frame = pd.DataFrame(
        {
            "position": np.arange(len(data), dtype=int),
            "time_group": pd.qcut(np.arange(len(data)), q=3, labels=False, duplicates="drop"),
            "target_group": pd.qcut(
                data[target_col].rank(method="first"), q=3, labels=False, duplicates="drop"
            ),
        }
    )
    frame["stratum"] = frame["time_group"].astype(str) + "_" + frame["target_group"].astype(str)

    groups = list(frame.groupby("stratum", sort=True))
    per_group = max(1, n_samples // max(len(groups), 1))
    selected: List[int] = []

    for _, group in groups:
        take = min(per_group, len(group))
        selected.extend(
            rng.choice(group["position"].to_numpy(), size=take, replace=False).tolist()
        )

    selected = sorted(set(selected))
    remaining = n_samples - len(selected)
    if remaining > 0:
        available = np.setdiff1d(np.arange(len(data)), np.asarray(selected, dtype=int))
        if len(available):
            selected.extend(
                rng.choice(
                    available,
                    size=min(remaining, len(available)),
                    replace=False,
                ).tolist()
            )

    return np.asarray(sorted(set(selected))[:n_samples], dtype=int)

# =============================================================================
# 6. 一维 PDP / ICE 计算
# =============================================================================

@dataclass
class VariableResponseResult:
    target_name: str
    variable: str
    grid: np.ndarray
    summary: pd.DataFrame
    ice_long: pd.DataFrame
    response_metrics: Dict[str, Any]
    bootstrap_metrics: Dict[str, Any]


def classify_curve(
    grid: np.ndarray,
    response: np.ndarray,
    baseline_prediction_std: float,
) -> Dict[str, Any]:
    """
    根据中心化PDP曲线给出简洁的总体响应描述。

    说明：
    - 响应幅度用中心化PDP最大值减最小值表示；
    - 标准化响应幅度 = 响应幅度 / 原始状态预测标准差；
    - 幅度不超过基线预测标准差的2%时记为“模型响应近零”；
    - Spearman |rho| >= 0.8时才标记总体上升/下降。
    """
    mask = np.isfinite(grid) & np.isfinite(response)
    x = grid[mask]
    y = response[mask]

    if len(x) < 2:
        return {
            "响应类型": "可用网格不足，无法分类",
            "Spearman_rho": np.nan,
            "响应幅度": np.nan,
            "标准化响应幅度": np.nan,
            "最低网格响应": np.nan,
            "最高网格响应": np.nan,
            "高低水平响应差": np.nan,
        }

    amplitude = float(np.max(y) - np.min(y))
    denom = float(baseline_prediction_std) if baseline_prediction_std > 1e-12 else np.nan
    normalized = amplitude / denom if np.isfinite(denom) else np.nan
    rho = robust_spearman(x, y)

    low_response = float(y[np.argmin(x)])
    high_response = float(y[np.argmax(x)])
    high_low = high_response - low_response

    zero_threshold = max(
        1e-12,
        baseline_prediction_std * NEGLIGIBLE_NORMALIZED_AMPLITUDE,
    )

    if amplitude <= zero_threshold:
        response_type = "模型响应近零"
    elif len(x) == 2:
        if high_low > 0:
            response_type = "两水平比较：上升"
        elif high_low < 0:
            response_type = "两水平比较：下降"
        else:
            response_type = "模型响应近零"
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
        "最低网格响应": low_response,
        "最高网格响应": high_response,
        "高低水平响应差": high_low,
    }


def compute_reference_bootstrap(
    target_name: str,
    variable: str,
    grid: np.ndarray,
    delta_matrix: np.ndarray,
    baseline_prediction_std: float,
) -> Tuple[pd.DataFrame, Dict[str, Any]]:
    """
    对历史参考样本组成进行Bootstrap。

    不重新训练模型，不重新拟合scaler。
    每次仅对历史参考样本行有放回抽样，再重新计算各网格点的平均中心化响应。
    """
    rng = np.random.default_rng(
        (RANDOM_SEED + stable_seed(f"{target_name}-{variable}-reference-bootstrap"))
        % (2**32 - 1)
    )

    n, g = delta_matrix.shape
    boot_curves = np.full((BOOTSTRAP_N, g), np.nan, dtype=float)
    metric_rows: List[Dict[str, Any]] = []

    for b in range(BOOTSTRAP_N):
        sampled = rng.choice(np.arange(n), size=n, replace=True)

        for j in range(g):
            values = delta_matrix[sampled, j]
            values = values[np.isfinite(values)]
            if len(values) > 0:
                boot_curves[b, j] = float(np.mean(values))

        metrics = classify_curve(
            grid,
            boot_curves[b],
            baseline_prediction_std,
        )
        metric_rows.append(
            {
                "Bootstrap编号": b + 1,
                "响应类型": metrics["响应类型"],
                "Spearman_rho": metrics["Spearman_rho"],
                "响应幅度": metrics["响应幅度"],
                "标准化响应幅度": metrics["标准化响应幅度"],
                "高低水平响应差": metrics["高低水平响应差"],
            }
        )

    point_rows: List[Dict[str, Any]] = []
    for j, value in enumerate(grid):
        vals = boot_curves[:, j]
        vals = vals[np.isfinite(vals)]
        point_rows.append(
            {
                "目标指标": target_name,
                "变量": variable,
                "网格值": float(value),
                "Bootstrap有效次数": int(len(vals)),
                "中心化响应Bootstrap中位数": (
                    float(np.median(vals)) if len(vals) else np.nan
                ),
                "中心化响应Bootstrap2.5%分位": (
                    float(np.quantile(vals, 0.025)) if len(vals) else np.nan
                ),
                "中心化响应Bootstrap97.5%分位": (
                    float(np.quantile(vals, 0.975)) if len(vals) else np.nan
                ),
            }
        )

    metrics_df = pd.DataFrame(metric_rows)
    valid_metrics = metrics_df.loc[
        metrics_df["响应类型"] != "可用网格不足，无法分类"
    ].copy()

    type_series = (
        valid_metrics["响应类型"].astype(str)
        if len(valid_metrics)
        else pd.Series(dtype=str)
    )
    rise_count = (
        int(type_series.str.contains("上升", regex=False).sum())
        if len(type_series)
        else 0
    )
    fall_count = (
        int(type_series.str.contains("下降", regex=False).sum())
        if len(type_series)
        else 0
    )
    nonmono_count = (
        int((type_series == "非单调/平台型").sum())
        if len(type_series)
        else 0
    )
    zero_count = (
        int(type_series.str.contains("响应近零", regex=False).sum())
        if len(type_series)
        else 0
    )

    final_summary: Dict[str, Any] = {
        "Bootstrap次数": BOOTSTRAP_N,
        "Bootstrap有效响应次数": int(len(valid_metrics)),
        "Bootstrap有效比例百分比": (
            100.0 * len(valid_metrics) / max(BOOTSTRAP_N, 1)
        ),
        "Bootstrap总体上升比例百分比": (
            100.0 * rise_count / max(len(valid_metrics), 1)
        ),
        "Bootstrap总体下降比例百分比": (
            100.0 * fall_count / max(len(valid_metrics), 1)
        ),
        "Bootstrap非单调平台比例百分比": (
            100.0 * nonmono_count / max(len(valid_metrics), 1)
        ),
        "Bootstrap近零比例百分比": (
            100.0 * zero_count / max(len(valid_metrics), 1)
        ),
        "标准化响应幅度Bootstrap中位数": (
            float(valid_metrics["标准化响应幅度"].median())
            if len(valid_metrics)
            else np.nan
        ),
        "标准化响应幅度Bootstrap2.5%分位": (
            float(valid_metrics["标准化响应幅度"].quantile(0.025))
            if len(valid_metrics)
            else np.nan
        ),
        "标准化响应幅度Bootstrap97.5%分位": (
            float(valid_metrics["标准化响应幅度"].quantile(0.975))
            if len(valid_metrics)
            else np.nan
        ),
    }
    return pd.DataFrame(point_rows), final_summary


def summarize_ice_heterogeneity(
    ice_long: pd.DataFrame,
    baseline_prediction_std: float,
    final_response_type: str,
) -> Dict[str, Any]:
    """
    汇总代表样本ICE曲线的方向差异。
    """
    rows: List[Dict[str, Any]] = []
    zero_threshold = max(
        1e-12,
        baseline_prediction_std * NEGLIGIBLE_NORMALIZED_AMPLITUDE,
    )

    for sample_id, group in ice_long.groupby("样本序号", sort=False):
        sub = group.loc[
            (group["物理有效"] == "是")
            & group["中心化ICE响应"].notna()
        ].sort_values("网格值")

        if len(sub) < 2:
            continue

        x = sub["网格值"].to_numpy(dtype=float)
        y = sub["中心化ICE响应"].to_numpy(dtype=float)
        amplitude = float(np.max(y) - np.min(y))

        if amplitude <= zero_threshold:
            cls = "近零"
            rho = np.nan
        else:
            rho = robust_spearman(x, y)
            if np.isfinite(rho) and rho >= ICE_DIRECTION_RHO_THRESHOLD:
                cls = "上升"
            elif np.isfinite(rho) and rho <= -ICE_DIRECTION_RHO_THRESHOLD:
                cls = "下降"
            else:
                cls = "混合/非单调"

        rows.append(
            {
                "样本序号": sample_id,
                "ICE类型": cls,
                "rho": rho,
                "幅度": amplitude,
            }
        )

    detail = pd.DataFrame(rows)
    n = len(detail)

    if n == 0:
        return {
            "ICE可分类样本数": 0,
            "ICE上升比例百分比": np.nan,
            "ICE下降比例百分比": np.nan,
            "ICE混合非单调比例百分比": np.nan,
            "ICE近零比例百分比": np.nan,
            "ICE与总体方向一致比例百分比": np.nan,
        }

    counts = detail["ICE类型"].value_counts().to_dict()

    if "上升" in final_response_type:
        consistent = counts.get("上升", 0)
    elif "下降" in final_response_type:
        consistent = counts.get("下降", 0)
    elif "响应近零" in final_response_type:
        consistent = counts.get("近零", 0)
    else:
        consistent = counts.get("混合/非单调", 0)

    return {
        "ICE可分类样本数": n,
        "ICE上升比例百分比": 100.0 * counts.get("上升", 0) / n,
        "ICE下降比例百分比": 100.0 * counts.get("下降", 0) / n,
        "ICE混合非单调比例百分比": 100.0 * counts.get("混合/非单调", 0) / n,
        "ICE近零比例百分比": 100.0 * counts.get("近零", 0) / n,
        "ICE与总体方向一致比例百分比": 100.0 * consistent / n,
    }


def compute_variable_response(
    target_name: str,
    variable: str,
    data: pd.DataFrame,
    package: Mapping[str, Any],
    baseline_predictions: np.ndarray,
    ice_indices: np.ndarray,
) -> Tuple[VariableResponseResult, pd.DataFrame]:
    """
    计算单个目标、单个运行变量的一维经典PDP、中心化PDP和代表样本ICE。
    """
    grid, grid_stats = build_grid(data[variable], variable)

    n = len(data)
    g = len(grid)
    baseline_sd = (
        float(np.std(baseline_predictions, ddof=1))
        if n > 1
        else 0.0
    )

    pred_matrix = np.full((n, g), np.nan, dtype=float)
    delta_matrix = np.full((n, g), np.nan, dtype=float)
    physical_matrix = np.zeros((n, g), dtype=bool)

    summary_rows: List[Dict[str, Any]] = []

    for j, grid_value in enumerate(grid):
        scenarios = data.copy(deep=True)
        scenarios[variable] = float(grid_value)

        physical = physical_valid_mask(scenarios)
        physical_matrix[:, j] = physical

        if np.any(physical):
            matrix = scenarios.loc[
                physical,
                list(package["feature_names"]),
            ].to_numpy(dtype=float)

            transformed = package["scaler"].transform(matrix)
            predictions = np.asarray(
                package["model"].predict(transformed),
                dtype=float,
            )

            pred_matrix[physical, j] = predictions
            delta_matrix[physical, j] = (
                predictions - baseline_predictions[physical]
            )

        valid_pred = pred_matrix[physical, j]
        valid_delta = delta_matrix[physical, j]
        finite_pred = valid_pred[np.isfinite(valid_pred)]
        finite_delta = valid_delta[np.isfinite(valid_delta)]

        exact_count = int(
            np.sum(
                np.isclose(
                    data[variable].to_numpy(dtype=float),
                    float(grid_value),
                    rtol=0.0,
                    atol=1e-12,
                )
            )
        )

        summary_rows.append(
            {
                "目标指标": target_name,
                "变量": variable,
                "网格类型": grid_stats["网格类型"],
                "网格序号": j + 1,
                "网格值": float(grid_value),
                "历史该精确水平样本数": exact_count,
                "总参考样本数": n,
                "物理有效情景数": int(np.sum(physical)),
                "物理无效情景数": int(n - np.sum(physical)),
                "物理有效比例百分比": 100.0 * np.sum(physical) / max(n, 1),
                "经典PDP平均预测值": (
                    float(np.mean(finite_pred))
                    if len(finite_pred)
                    else np.nan
                ),
                "经典PDP预测标准差": (
                    float(np.std(finite_pred, ddof=1))
                    if len(finite_pred) > 1
                    else np.nan
                ),
                "中心化PDP响应": (
                    float(np.mean(finite_delta))
                    if len(finite_delta)
                    else np.nan
                ),
            }
        )

    summary = pd.DataFrame(summary_rows)

    curve_for_classification = summary["中心化PDP响应"].to_numpy(dtype=float)
    response_metrics = classify_curve(
        grid,
        curve_for_classification,
        baseline_sd,
    )

    response_metrics.update(
        {
            "目标指标": target_name,
            "变量": variable,
            "网格类型": grid_stats["网格类型"],
            "总网格点数": g,
            "可计算网格点数": int(
                np.isfinite(summary["中心化PDP响应"]).sum()
            ),
            "物理有效比例最小值百分比": float(
                summary["物理有效比例百分比"].min()
            ),
            "物理有效比例中位数百分比": float(
                summary["物理有效比例百分比"].median()
            ),
            "物理有效比例最大值百分比": float(
                summary["物理有效比例百分比"].max()
            ),
            "方向判定曲线基础": "全部物理有效历史参考样本的经典中心化PDP",
            "基线预测标准差": baseline_sd,
            "响应近零阈值_绝对值": (
                baseline_sd * NEGLIGIBLE_NORMALIZED_AMPLITUDE
            ),
        }
    )

    # 对ElasticNet补充模型结构信息。
    elasticnet_coef = np.nan
    elasticnet_coef_zero = False

    if (
        str(package.get("algorithm")) == "ElasticNet"
        and hasattr(package["model"], "coef_")
    ):
        feature_names = list(package["feature_names"])
        if variable in feature_names:
            elasticnet_coef = float(
                np.asarray(package["model"].coef_, dtype=float)[
                    feature_names.index(variable)
                ]
            )
            elasticnet_coef_zero = abs(elasticnet_coef) <= 1e-12

    response_metrics["ElasticNet标准化系数"] = elasticnet_coef
    response_metrics["ElasticNet对应系数是否为0"] = (
        "是" if elasticnet_coef_zero else "否"
    )

    if elasticnet_coef_zero:
        response_metrics["响应类型"] = "模型响应近零（ElasticNet对应系数为0）"
        response_metrics["Spearman_rho"] = np.nan
        response_metrics["响应幅度"] = 0.0
        response_metrics["标准化响应幅度"] = 0.0
        response_metrics["最低网格响应"] = 0.0
        response_metrics["最高网格响应"] = 0.0
        response_metrics["高低水平响应差"] = 0.0
        response_metrics["方向判定曲线基础"] = (
            "ElasticNet对应标准化系数为0，且经典PDP不随该变量变化"
        )

    # 代表样本ICE长表。
    ice_rows: List[Dict[str, Any]] = []

    for position in ice_indices:
        sample_id = int(data.iloc[position]["样本序号"])
        sample_date = data.iloc[position]["Date"]
        original_value = float(data.iloc[position][variable])
        baseline_pred = float(baseline_predictions[position])

        for j, grid_value in enumerate(grid):
            ice_rows.append(
                {
                    "目标指标": target_name,
                    "变量": variable,
                    "样本序号": sample_id,
                    "Date": sample_date,
                    "原始变量值": original_value,
                    "网格值": float(grid_value),
                    "原始状态预测值": baseline_pred,
                    "扰动后预测值": pred_matrix[position, j],
                    "中心化ICE响应": delta_matrix[position, j],
                    "物理有效": (
                        "是" if physical_matrix[position, j] else "否"
                    ),
                }
            )

    ice_long = pd.DataFrame(ice_rows)

    ice_metrics = summarize_ice_heterogeneity(
        ice_long,
        baseline_sd,
        str(response_metrics["响应类型"]),
    )
    response_metrics.update(ice_metrics)

    # 历史参考样本Bootstrap。
    boot_points, boot_metrics = compute_reference_bootstrap(
        target_name,
        variable,
        grid,
        delta_matrix,
        baseline_sd,
    )
    response_metrics.update(boot_metrics)

    summary = summary.merge(
        boot_points,
        on=["目标指标", "变量", "网格值"],
        how="left",
    )

    grid_row = pd.DataFrame(
        [
            {
                "目标指标": target_name,
                "变量": variable,
                **grid_stats,
                "分析说明": (
                    "运行变量网格依据历史观测范围和实际水平确定；"
                    "扰动情景仅进行基本物理合理性检查。"
                ),
            }
        ]
    )

    return (
        VariableResponseResult(
            target_name=target_name,
            variable=variable,
            grid=grid,
            summary=summary,
            ice_long=ice_long,
            response_metrics=response_metrics,
            bootstrap_metrics=boot_metrics,
        ),
        grid_row,
    )

# =============================================================================
# 7. 结果整合、二维候选和文字说明
# =============================================================================

def response_recommendation(
    grade: str,
    response_type: str,
    calculable_points: int,
) -> str:
    if calculable_points < 2:
        return "可计算网格不足，不作为正式一维响应结论。"

    if "模型响应近零" in response_type:
        return (
            "冻结最终模型对该运行变量几乎没有预测响应；"
            "不建议作为PDP主图或后续运行优化的主要驱动变量。"
        )

    if grade == "A":
        return (
            "可作为主文重点变量：SHAP/PI证据较强，"
            "并可进一步结合PDP形状、ICE个体差异和Bootstrap稳定性解释。"
        )

    if grade == "B":
        return (
            "可作为重点变量，但应同时说明PI可评估性、响应幅度和ICE个体差异。"
        )

    if grade == "C":
        return (
            "作为辅助或探索性结果保留，不宜单独作为核心运行效应证据。"
        )

    return "不作为运行响应主解释。"


def build_combined_response_table(
    response_table: pd.DataFrame,
    evidence: pd.DataFrame,
) -> pd.DataFrame:
    out = response_table.merge(
        evidence,
        on=["目标指标", "变量"],
        how="left",
    )

    out["PDP_ICE解释建议"] = [
        response_recommendation(
            str(row.get("证据等级", "D")),
            str(row.get("响应类型", "")),
            int(row.get("可计算网格点数", 0)),
        )
        for _, row in out.iterrows()
    ]

    target_rank = {name: index for index, name in enumerate(TARGET_ORDER)}
    out["_目标顺序"] = out["目标指标"].map(target_rank).fillna(len(TARGET_ORDER))
    out = out.sort_values(
        ["_目标顺序", "证据等级", "标准化响应幅度"],
        ascending=[True, True, False],
        na_position="last",
        kind="mergesort",
    ).drop(columns=["_目标顺序"]).reset_index(drop=True)
    return out


def build_2d_candidates(
    response_with_evidence: pd.DataFrame,
) -> pd.DataFrame:
    """
    根据一维SHAP/PI证据和PDP响应结果给出下一阶段二维分析候选。
    只输出真正满足进入条件的组合；本脚本本身不计算二维响应。
    """
    rows: List[Dict[str, Any]] = []

    for target in TARGET_ORDER:
        target_response = response_with_evidence.loc[
            response_with_evidence["目标指标"] == target
        ].set_index("变量")

        for var1, var2, reason in PREDEFINED_2D_PAIRS[target]:
            if var1 not in target_response.index or var2 not in target_response.index:
                continue

            r1 = target_response.loc[var1]
            r2 = target_response.loc[var2]

            grade1 = str(r1["证据等级"])
            grade2 = str(r2["证据等级"])
            response1 = str(r1["响应类型"])
            response2 = str(r2["响应类型"])
            points1 = int(r1["可计算网格点数"])
            points2 = int(r2["可计算网格点数"])

            usable = (
                grade1 != "D"
                and grade2 != "D"
                and "模型响应近零" not in response1
                and "模型响应近零" not in response2
                and response1 != "可用网格不足，无法分类"
                and response2 != "可用网格不足，无法分类"
                and points1 >= 2
                and points2 >= 2
            )
            if not usable:
                continue

            if (
                "A" in {grade1, grade2}
                and {grade1, grade2}.issubset({"A", "B"})
            ):
                priority = "高"
            elif (
                grade1 in {"A", "B"}
                and grade2 in {"A", "B", "C"}
            ) or (
                grade2 in {"A", "B"}
                and grade1 in {"A", "B", "C"}
            ):
                priority = "中"
            else:
                priority = "探索"

            rows.append(
                {
                    "目标指标": target,
                    "变量1": var1,
                    "变量2": var2,
                    "二维候选优先级": priority,
                    "一维前置条件结论": (
                        "两个变量的一维结果均具备可解释响应，可进入下一步二维条件响应分析。"
                    ),
                    "预定义原因": reason,
                    "提醒": (
                        "本脚本未计算二维响应。二维阶段沿用一维网格，"
                        "并统计历史精确二元组合出现次数作为解释背景；"
                        "历史组合频次不作为过滤门槛。"
                    ),
                }
            )

    columns = [
        "目标指标", "变量1", "变量2", "二维候选优先级",
        "一维前置条件结论", "预定义原因", "提醒",
    ]
    return pd.DataFrame(rows, columns=columns)


def build_text_report(
    settings: pd.DataFrame,
    model_checks: pd.DataFrame,
    response: pd.DataFrame,
    candidates: pd.DataFrame,
    warnings_table: pd.DataFrame,
    elapsed_minutes: float,
) -> str:
    lines: List[str] = []

    lines.append("A²/O 四指标一维 PDP / ICE 运行响应分析说明")
    lines.append("=" * 112)
    lines.append("")

    lines.append("一、任务与输出")
    lines.append(f"1. 完成时间：{now_str()}")
    lines.append(f"2. 总运行耗时：{elapsed_minutes:.2f} 分钟")
    lines.append(f"3. 动态数据：{INPUT_FILE}")
    lines.append(f"4. 最终模型目录：{MODEL_DIR}")
    lines.append(f"5. SHAP + PI结果：{SHAP_PI_FILE}")
    lines.append(f"6. 输出Excel：{OUTPUT_EXCEL}")
    lines.append(f"7. 输出TXT：{OUTPUT_TEXT}")
    lines.append(
        "8. 本阶段只输出数据表和文字说明，不生成任何图片，也不计算二维响应。"
    )
    lines.append("")

    lines.append("二、方法与解释边界")
    lines.append(
        "1. 四个冻结joblib直接用于PDP/ICE，不重新选模型、不重新调参，"
        "也不因响应结果修改最终模型。"
    )
    lines.append(
        "2. 仅扫描5个运行/结构变量；进水条件与各目标lag1/lag2/roll3"
        "保持每条历史样本原值。"
    )
    lines.append(
        "3. HRT、进水分配比例和两个体积分数使用历史真实观测水平；"
        "DO在历史唯一水平较多时使用历史5%–95%分位范围内的分位数网格。"
    )
    lines.append(
        "4. 每个网格点直接使用全部物理有效历史参考样本计算经典PDP。"
    )
    lines.append(
        "5. 中心化PDP和中心化ICE均以同一样本原始状态预测值为基准，"
        "表示仅改变当前运行变量后冻结模型预测值的变化。"
    )
    lines.append(
        "6. 仅保留基本物理约束：DO>=0、Feed在0-1之间、HRT>0、"
        "两个体积分数均在0-1之间且二者之和<1。"
    )
    lines.append(
        f"7. 响应幅度不超过基线预测标准差的"
        f"{NEGLIGIBLE_NORMALIZED_AMPLITUDE*100:g}%时记为“模型响应近零”；"
        f"Spearman |rho|>={MONOTONIC_SPEARMAN_THRESHOLD:g}时才标记总体上升/下降。"
    )
    lines.append(
        f"8. Bootstrap次数={BOOTSTRAP_N}，只对历史参考样本组成进行重采样，"
        "不重新训练模型；经验区间不是模型参数置信区间。"
    )
    lines.append(
        "9. SHAP/PI证据等级只用于决定解释优先级，不作为PDP数值计算的权重。"
    )
    lines.append(
        "10. 历史精确运行水平出现次数仅作为结果解释和审计信息，"
        "不作为PDP网格筛选门槛。"
    )
    lines.append(
        "11. PDP/ICE描述冻结模型在历史变量范围内的条件响应，"
        "不是实验干预因果效应，也不是长期稳态过程响应。"
    )
    lines.append("")

    lines.append("三、模型与数据核查")
    for _, row in model_checks.iterrows():
        lines.append(
            f"- {row['目标指标']}："
            f"{row['最终特征组']} + {row['算法']}（{row['候选参数ID']}），"
            f"样本数={int(row['样本数'])}，特征数={int(row['特征数'])}；"
            f"全量拟合复现最大指标差="
            f"{fmt(row['全量拟合复现最大指标绝对差'], 12)}；"
            f"ICE代表样本数={int(row['ICE代表样本数'])}；"
            f"核查={row['核查状态']}。"
        )
    lines.append("")

    chapter_names = {
        "TP": "四",
        "TN": "五",
        "NH4+-N": "六",
        "COD": "七",
    }

    for target in TARGET_ORDER:
        lines.append(f"{chapter_names[target]}、{target} 一维运行响应")
        sub = response.loc[
            response["目标指标"] == target
        ].copy()

        if sub.empty:
            lines.append("- 无结果。")
            lines.append("")
            continue

        package_metrics = model_checks.loc[
            model_checks["目标指标"] == target
        ].iloc[0]

        lines.append(
            f"1. 冻结模型："
            f"{package_metrics['最终特征组']}-"
            f"{package_metrics['算法']}（{package_metrics['候选参数ID']}）。"
        )
        lines.append("2. 五个运行变量结果：")

        grade_order = {"A": 0, "B": 1, "C": 2, "D": 3}
        sub["_grade_order"] = sub["证据等级"].map(
            grade_order
        ).fillna(9)
        sub = sub.sort_values(
            ["_grade_order", "标准化响应幅度"],
            ascending=[True, False],
        )

        for _, row in sub.iterrows():
            lines.append(
                f"   - {row['变量']}："
                f"证据{row['证据等级']}；"
                f"网格={row['网格类型']}；"
                f"可计算网格={int(row['可计算网格点数'])}/"
                f"{int(row['总网格点数'])}；"
                f"响应={row['响应类型']}；"
                f"标准化幅度={fmt(row['标准化响应幅度'], 3)}；"
                f"rho={fmt(row['Spearman_rho'], 3)}；"
                f"物理有效比例最低="
                f"{fmt(row['物理有效比例最小值百分比'], 1)}%；"
                f"ICE方向一致="
                f"{fmt(row['ICE与总体方向一致比例百分比'], 1)}%；"
                f"Bootstrap有效="
                f"{fmt(row['Bootstrap有效比例百分比'], 1)}%；"
                f"建议：{row['PDP_ICE解释建议']}"
            )

        if target == "NH4+-N":
            lines.append(
                "3. NH4+-N特别边界：其最佳ML未在Time与Expanding中同时超过roll3。"
                "若五个运行变量响应接近零，表示当前最终ElasticNet没有学到"
                "可用于工况响应或优化的稳定运行信号，不能据此声称实际NH4+-N"
                "不受DO、HRT等运行参数影响。"
            )
        else:
            main_vars = sub.loc[
                sub["PDP_ICE解释建议"]
                .astype(str)
                .str.contains("主文重点变量|重点变量")
            ]

            if len(main_vars):
                lines.append(
                    "3. 当前较值得进入后续论文PDP/ICE图形筛选的变量："
                    + "、".join(main_vars["变量"].tolist())
                    + "。"
                )
            else:
                lines.append(
                    "3. 当前没有同时满足较高解释证据和明显模型响应的主变量，"
                    "建议保持探索性表述。"
                )

        lines.append("")

    lines.append("八、后续二维候选（本次未计算二维）")
    if candidates.empty:
        lines.append("- 无候选。")
    else:
        for _, row in candidates.iterrows():
            lines.append(
                f"- {row['目标指标']}："
                f"{row['变量1']} × {row['变量2']}；"
                f"优先级={row['二维候选优先级']}；"
                f"{row['一维前置条件结论']}"
            )
    lines.append(
        "- 二维阶段沿用一维网格，并结合二维PDP、切片、Bootstrap、非加性分析"
        "和历史精确二元组合出现次数进行解释。"
    )
    lines.append("")

    lines.append("九、Excel数据表用途")
    lines.append(
        "1. ‘运行设置与模型核查’：运行参数、模型文件哈希、"
        "动态数据一致性和全量拟合复现结果。"
    )
    lines.append(
        "2. ‘运行变量证据’：上一阶段SHAP、Time PI、Expanding PI"
        "及A-D证据等级。"
    )
    lines.append(
        "3. ‘响应网格定义’：每个运行变量的历史范围、唯一值数量和实际响应网格。"
    )
    lines.append(
        "4. ‘PDP汇总’：经典PDP、中心化PDP、物理有效样本数量、"
        "历史精确水平样本数和Bootstrap点位区间。"
    )
    lines.append(
        "5. ‘ICE抽样长表’：后续中心化ICE图的逐样本逐网格原始数据。"
    )
    lines.append(
        "6. ‘响应方向与幅度’：总体响应类型、标准化幅度、"
        "ICE异质性与Bootstrap稳定性。"
    )
    lines.append(
        "7. ‘PDP_SHAP_PI对照’：用于论文选择重点变量，不建立加权总分。"
    )
    lines.append(
        "8. ‘后续二维候选’：只列下一阶段候选，不在本脚本中计算二维响应。"
    )
    lines.append("")

    lines.append("十、完整性和警告")
    lines.append(
        f"1. 运行警告/错误记录数：{len(warnings_table)}。"
    )
    lines.append(
        "2. 本脚本如在严格模式发现模型配置、样本ID、动态数据哈希、"
        "全量拟合指标或SHAP解释对象不一致，会停止而不生成正式结果。"
    )
    lines.append(
        "3. 本阶段没有输出图片，没有修改joblib，"
        "也没有用PDP结果反向修改最终模型。"
    )

    return "\n".join(lines)

# =============================================================================
# 8. 主流程
# =============================================================================

def main() -> None:
    overall_start = time.time()
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    log("A²/O 四指标一维 PDP / ICE 响应分析开始。")
    log(f"动态数据：{INPUT_FILE}")
    log(f"最终模型目录：{MODEL_DIR}")
    log(f"SHAP + PI结果：{SHAP_PI_FILE}")
    log(f"输出目录：{OUTPUT_DIR}")
    log(
        f"Bootstrap次数：{BOOTSTRAP_N}；"
        f"ICE代表样本数：{ICE_SAMPLE_N}；"
        f"SmokeTest={SMOKE_TEST}"
    )

    if not INPUT_FILE.exists():
        raise FileNotFoundError(INPUT_FILE)
    if not SHAP_PI_FILE.exists():
        raise FileNotFoundError(SHAP_PI_FILE)

    dynamic_hash = sha256_file(INPUT_FILE)
    shap_hash = sha256_file(SHAP_PI_FILE)

    shap_comparison, _ = read_shap_pi_evidence()
    operation_evidence = build_operation_evidence(shap_comparison)

    setting_rows: List[Dict[str, Any]] = [
        {"项目": "完成时间", "值": now_str()},
        {"项目": "动态数据", "值": str(INPUT_FILE)},
        {"项目": "动态数据SHA256", "值": dynamic_hash},
        {"项目": "最终模型目录", "值": str(MODEL_DIR)},
        {"项目": "SHAP+PI结果", "值": str(SHAP_PI_FILE)},
        {"项目": "SHAP+PI结果SHA256", "值": shap_hash},
        {"项目": "输出目录", "值": str(OUTPUT_DIR)},
        {"项目": "参考样本Bootstrap次数", "值": BOOTSTRAP_N},
        {"项目": "ICE分层代表样本目标数", "值": ICE_SAMPLE_N},
        {"项目": "DO连续网格目标点数", "值": DO_GRID_N},
        {
            "项目": "离散运行变量网格",
            "值": "HRT、Feed、厌氧/缺氧体积分数使用全部历史真实观测水平",
        },
        {
            "项目": "DO网格",
            "值": (
                f"唯一水平<={AUTO_DISCRETE_UNIQUE_MAX}时使用历史真实水平；"
                f"否则使用历史5%–95%分位范围内{DO_GRID_N}个分位数网格点"
            ),
        },
        {
            "项目": "物理约束",
            "值": (
                "DO>=0；0<=Feed<=1；HRT>0；两个体积分数均在0-1之间；"
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
        {"项目": "Smoke test", "值": "是" if SMOKE_TEST else "否"},
        {"项目": "严格核查", "值": "是" if STRICT_MODE else "否"},
    ]

    model_check_rows: List[Dict[str, Any]] = []
    imputation_tables: List[pd.DataFrame] = []
    grid_tables: List[pd.DataFrame] = []
    pdp_tables: List[pd.DataFrame] = []
    ice_tables: List[pd.DataFrame] = []
    response_rows: List[Dict[str, Any]] = []
    warning_rows: List[Dict[str, Any]] = []

    for target_name in TARGET_ORDER:
        start = time.time()
        log(f"开始 {target_name} 一维响应分析。")

        try:
            model_path, package = load_model_package(target_name)
            raw = read_target_data(target_name)
            data, imputation = rebuild_training_data(
                target_name,
                raw,
                package,
            )
            baseline_pred, full_metrics, metric_diffs = validate_full_fit(
                target_name,
                data,
                package,
            )

            # 模型包记录的动态文件hash必须与当前动态数据一致。
            saved_dynamic_hash = str(
                package["training_data"].get(
                    "dynamic_input_file_sha256",
                    "",
                )
            )
            dynamic_hash_match = (
                saved_dynamic_hash == dynamic_hash
                if saved_dynamic_hash
                else False
            )

            if STRICT_MODE and not dynamic_hash_match:
                raise RuntimeError(
                    f"{target_name} 模型包动态数据SHA256与当前文件不一致。"
                )

            cfg = TARGET_CONFIG[target_name]
            ice_indices = stratified_ice_indices(
                data,
                cfg["target_col"],
                ICE_SAMPLE_N,
                f"{target_name}-ICE",
            )

            imputation_tables.append(imputation)

            for variable in CONTROLLABLE_FEATURES:
                log(
                    f"{target_name}：计算 {variable} 经典PDP/中心化ICE。"
                )

                result, grid_table = compute_variable_response(
                    target_name,
                    variable,
                    data,
                    package,
                    baseline_pred,
                    ice_indices,
                )

                pdp_tables.append(result.summary)
                ice_tables.append(result.ice_long)
                grid_tables.append(grid_table)
                response_rows.append(result.response_metrics)

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
                        else np.nan
                    ),
                    "ICE代表样本数": len(ice_indices),
                    "Bootstrap次数": BOOTSTRAP_N,
                    "核查状态": "通过",
                    "运行耗时分钟": (time.time() - start) / 60.0,
                }
            )

            log(
                f"{target_name} 完成，"
                f"耗时 {(time.time()-start)/60:.2f} 分钟。"
            )

        except Exception as exc:
            warning_rows.append(
                {
                    "目标指标": target_name,
                    "阶段": "一维响应分析",
                    "警告或错误": repr(exc),
                    "Traceback": traceback.format_exc(),
                }
            )
            log(f"{target_name} 失败：{exc}")

            if STRICT_MODE:
                break

    warnings_table = pd.DataFrame(warning_rows)

    if warning_rows and STRICT_MODE:
        error_path = OUTPUT_DIR / "一维PDP_ICE分析失败记录.txt"
        atomic_write_text(
            error_path,
            "\n\n".join(
                f"目标：{row['目标指标']}\n"
                f"{row['警告或错误']}\n"
                f"{row['Traceback']}"
                for row in warning_rows
            ),
        )
        raise RuntimeError(
            f"一维响应分析失败；详见 {error_path}"
        )

    model_checks = pd.DataFrame(model_check_rows)
    imputation_all = (
        pd.concat(imputation_tables, ignore_index=True)
        if imputation_tables
        else pd.DataFrame()
    )
    grid_all = (
        pd.concat(grid_tables, ignore_index=True)
        if grid_tables
        else pd.DataFrame()
    )
    pdp_all = (
        pd.concat(pdp_tables, ignore_index=True)
        if pdp_tables
        else pd.DataFrame()
    )
    ice_all = (
        pd.concat(ice_tables, ignore_index=True)
        if ice_tables
        else pd.DataFrame()
    )
    response_all = pd.DataFrame(response_rows)

    response_with_evidence = build_combined_response_table(
        response_all,
        operation_evidence,
    )
    candidates = build_2d_candidates(response_with_evidence)

    # PDP_SHAP_PI对照保留最适合论文筛选的字段。
    comparison_columns = [
        "目标指标",
        "变量",
        "证据等级",
        "证据等级说明",
        "SHAP是否有效贡献",
        "SHAP排名",
        "SHAP重要性占比百分比",
        "SHAP方向结论",
        "Time_PI可评估性",
        "Time_PI证据结论",
        "Time_PI排名",
        "Time_RMSE增加均值",
        "Expanding_PI可评估性",
        "Expanding_PI证据结论",
        "Expanding_PI排名",
        "Expanding_RMSE增加均值",
        "响应类型",
        "Spearman_rho",
        "响应幅度",
        "标准化响应幅度",
        "ElasticNet标准化系数",
        "ElasticNet对应系数是否为0",
        "可计算网格点数",
        "总网格点数",
        "物理有效比例最小值百分比",
        "ICE与总体方向一致比例百分比",
        "Bootstrap有效比例百分比",
        "PDP_ICE解释建议",
        "是否涉及高相关特征",
        "SHAP_PI一致性解释",
    ]

    comparison_out = response_with_evidence[
        [
            column
            for column in comparison_columns
            if column in response_with_evidence.columns
        ]
    ].copy()

    # 所有正式结果表统一按 TP -> TN -> NH4+-N -> COD 排列。
    target_rank = {name: index for index, name in enumerate(TARGET_ORDER)}

    def sort_target_rows(frame: pd.DataFrame) -> pd.DataFrame:
        if frame is None or frame.empty or "目标指标" not in frame.columns:
            return frame
        out = frame.copy()
        out["_目标顺序"] = out["目标指标"].map(target_rank).fillna(len(TARGET_ORDER))
        out = out.sort_values("_目标顺序", kind="mergesort").drop(columns=["_目标顺序"])
        return out.reset_index(drop=True)

    model_checks = sort_target_rows(model_checks)
    imputation_all = sort_target_rows(imputation_all)
    operation_evidence = sort_target_rows(operation_evidence)
    grid_all = sort_target_rows(grid_all)
    pdp_all = sort_target_rows(pdp_all)
    ice_all = sort_target_rows(ice_all)
    response_with_evidence = sort_target_rows(response_with_evidence)
    comparison_out = sort_target_rows(comparison_out)
    candidates = sort_target_rows(candidates)

    settings = pd.DataFrame(setting_rows)
    elapsed = (time.time() - overall_start) / 60.0

    tables = {
        "运行设置与模型核查": pd.concat(
            [
                settings.assign(记录类型="运行设置"),
                model_checks.assign(记录类型="模型核查"),
            ],
            ignore_index=True,
            sort=False,
        ),
        "插补重建核查": imputation_all,
        "运行变量证据": operation_evidence,
        "响应网格定义": grid_all,
        "PDP汇总": pdp_all,
        "ICE抽样长表": ice_all,
        "响应方向与幅度": response_with_evidence,
        "PDP_SHAP_PI对照": comparison_out,
        "后续二维候选": candidates,
        "错误与警告": warnings_table,
    }

    text_report = build_text_report(
        settings,
        model_checks,
        response_with_evidence,
        candidates,
        warnings_table,
        elapsed,
    )

    atomic_write_excel(
        OUTPUT_EXCEL,
        tables,
    )
    atomic_write_text(
        OUTPUT_TEXT,
        text_report,
    )

    log("四指标一维 PDP / ICE 响应分析完成。")
    print("\n正式输出：")
    print(f" - {OUTPUT_EXCEL}")
    print(f" - {OUTPUT_TEXT}")
    print("本阶段未生成任何图片。")


if __name__ == "__main__":
    main()
