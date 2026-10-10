# -*- coding: utf-8 -*-
r"""
A²/O F0-F3 最终模型统一终选与随机稳定性复核（目标级直选修正版）
================================================================

用途
----
本脚本在上一版“统一终选与随机稳定性复核”基础上进一步修正最终模型选择逻辑。

本版不重新执行已经完成的 1000 次 Bootstrap 正式复核，而是直接读取
F0、F1、F2、F3 共 16 个既有模型结果 Excel 中的“正式候选综合信息”。

本版解决的核心问题
------------------
1. 完整指标输出
   所有最终比较统一报告：
   - Bootstrap .632+：R²、RMSE
   - Time：R²、RMSE
   - Expanding：R²、RMSE

2. 随机模型稳定性真正参与比较
   - ElasticNet / SVR / GPR 为确定性模型，直接使用原正式验证结果；
   - RandomForest / XGBoost / MLP 对“所有进入正式复核的候选参数”
     使用完全相同的 5 个随机种子重新执行 Time + Expanding；
   - 随机模型最终使用 5 个随机种子的平均 Time/Expanding 指标，
     同时保存标准差，避免幸运随机种子决定最终排名。

3. 消除“两层 5% 筛选”造成的候选丢失
   - 仍然输出每个“目标 × F组 × 算法”的 24 个组内代表配置，
     供算法比较、F0-F3 消融和结果展示使用；
   - 但 24 个代表配置不再作为最终模型的硬性入口；
   - 每个污染物的最终机器学习模型，直接从该目标“全部正式复核候选”
     的统一评价结果中选择，因此不会出现某个参数在组内被替换后，
     虽然适合参加目标级决赛却失去最终竞争资格的问题。

四个目标完全一致的最终选定规则
------------------------------
TP、TN、NH4+-N、COD 四个目标严格使用同一套规则：

第一步：对该目标全部正式候选使用统一评价指标。
- 确定性模型：采用原正式 Time / Expanding；
- 随机模型：采用本次 5 随机种子平均 Time / Expanding；
- Bootstrap .632+ R²/RMSE 均沿用原 1000 次正式复核结果。

第二步：在该目标全部候选中找到最低 Eval_Expanding_RMSE。

第三步：保留
    Eval_Expanding_RMSE <= 最低值 × (1 + 5%)
的候选，视为目标级近优候选。

第四步：在目标级 5% 近优候选中按完全相同的顺序终选：
1) Eval_Time_RMSE 越低越好；
2) Time_RMSE 稳定性标准差越低越好；
3) Bootstrap_RMSE 越低越好；
4) Eval_Time_MAE 越低越好；
5) Eval_Expanding_MAE 越低越好；
6) Expanding_RMSE 稳定性标准差越低越好；
7) 特征数越少越好；
8) 候选参数ID仅用于最后确定性排序。

说明：
- 对确定性模型，稳定性标准差记为 0，因为相同数据与参数下结果确定；
- 上述规则对四个目标完全一致，不针对任何单独污染物调整阈值或权重；
- 5%仅表示“实践上接近”的筛选边界，不代表统计显著性检验。

第五步：最终最佳 ML 与 lag1 / roll3 朴素动态基线比较。
- 若 ML 在 Time 和 Expanding 两种验证中都优于各自最佳历史基线，
  则标记为最终预测模型；
- 否则保留“最佳机器学习模型 + 更强历史基线”的性能边界说明。
  最佳 ML 仍可用于后续 SHAP/PDP 和运行响应分析，但论文中应明确边界。

默认输入目录
------------
${PROJECT_ROOT}/outputs/model_comparison

默认动态数据
------------
${PROJECT_ROOT}/outputs/preprocessing/dynamic/动态可用数据与插补比较结果.xlsx

默认输出目录
------------
${PROJECT_ROOT}/outputs/final_selection

主要输出
--------
1. F0-F3_统一终选与随机稳定性复核结果.xlsx
2. F0-F3_最终模型选定完整说明.txt

依赖
----
pip install pandas numpy openpyxl scikit-learn xgboost

重要说明
--------
- 不改变已经完成的 Bootstrap 结果；
- 随机模型会比上一版补跑更多稳定性，因为本版对所有正式随机候选统一复核；
- 24 个组内代表仅用于比较/消融，不再限制最终模型候选池；
- Q 不进入模型。
"""

from __future__ import annotations

import json
import math
import os
import re
import time
import traceback
import warnings
from copy import deepcopy
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Iterable, List, Mapping, Sequence, Tuple

import numpy as np
import pandas as pd
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter
from sklearn.ensemble import RandomForestRegressor
from sklearn.exceptions import ConvergenceWarning
from sklearn.gaussian_process import GaussianProcessRegressor
from sklearn.gaussian_process.kernels import ConstantKernel, Matern, RBF
from sklearn.linear_model import ElasticNet
from sklearn.metrics import mean_absolute_error, mean_squared_error, r2_score
from sklearn.neural_network import MLPRegressor
from sklearn.preprocessing import StandardScaler
from sklearn.svm import SVR

try:
    from xgboost import XGBRegressor
    XGBOOST_AVAILABLE = True
except Exception:
    XGBOOST_AVAILABLE = False


# =============================================================================
# 1. 路径与统一设置
# =============================================================================

PROJECT_ROOT = Path(__file__).resolve().parents[1]

RESULT_ROOT = Path(
    os.environ.get(
        "FINAL_MODEL_RESULT_ROOT",
        str(PROJECT_ROOT / "outputs/model_comparison"),
    )
)

DYNAMIC_INPUT_FILE = Path(
    os.environ.get(
        "FINAL_MODEL_DYNAMIC_INPUT",
        str(PROJECT_ROOT / "outputs/preprocessing/dynamic/动态可用数据与插补比较结果.xlsx"),
    )
)

OUTPUT_DIR = Path(
    os.environ.get(
        "FINAL_MODEL_OUTPUT_DIR",
        str(PROJECT_ROOT / "outputs/final_selection"),
    )
)

OUTPUT_EXCEL = OUTPUT_DIR / "F0-F3_统一终选与随机稳定性复核结果.xlsx"
OUTPUT_TEXT = OUTPUT_DIR / "F0-F3_最终模型选定完整说明.txt"

RANDOM_SEED = 20260620
TIME_TRAIN_RATIO = 0.80
EXPANDING_INITIAL_TRAIN_RATIO = 0.60
EXPANDING_VALID_WINDOW_SIZE = 10
EXPANDING_STEP_SIZE = 10
MAX_INTERPOLATION_GAP = 7
EPS = 1e-6
N_JOBS_TREE = int(os.environ.get("FINAL_MODEL_N_JOBS", "-1"))

# “性能接近”统一定义：Expanding RMSE 不超过最低值 5%。
NEAR_OPTIMAL_REL_TOL = float(os.environ.get("FINAL_MODEL_NEAR_OPTIMAL_TOL", "0.05"))

# 随机模型补跑稳定性次数，正式默认 5。
STABILITY_REPEATS = int(os.environ.get("FINAL_MODEL_STABILITY_REPEATS", "5"))
if STABILITY_REPEATS < 1:
    raise ValueError("FINAL_MODEL_STABILITY_REPEATS 至少为 1。")

# 本版对所有进入正式复核的随机模型候选统一执行稳定性复核，
# 不再使用“每个算法最多3组”的预筛限制，避免最终候选被过早排除。

# 所有随机模型候选使用完全相同的随机种子。
# 不把候选ID、特征组或目标写入种子，避免“不同参数使用不同随机数”影响公平比较。
_SEED_OFFSETS = [1009, 7919, 15401, 23003, 31013, 39019, 47017, 55001, 63029, 71011]
if STABILITY_REPEATS > len(_SEED_OFFSETS):
    # 超过预置数量时，用确定性方式继续生成。
    rng_seed = np.random.default_rng(RANDOM_SEED + 99173)
    extra = rng_seed.integers(100000, 2_000_000_000, size=STABILITY_REPEATS - len(_SEED_OFFSETS)).tolist()
    STABILITY_SEEDS = [RANDOM_SEED + x for x in _SEED_OFFSETS] + [int(x) for x in extra]
else:
    STABILITY_SEEDS = [RANDOM_SEED + x for x in _SEED_OFFSETS[:STABILITY_REPEATS]]

TARGET_ORDER = ["TP", "TN", "NH4+-N", "COD"]
FEATURE_GROUP_ORDER = ["F0", "F1", "F2", "F3"]
MODEL_ORDER = ["ElasticNet", "SVR", "GPR", "RandomForest", "XGBoost", "MLP"]
DETERMINISTIC_MODELS = {"ElasticNet", "SVR", "GPR"}
STOCHASTIC_MODELS = {"RandomForest", "XGBoost", "MLP"}

TARGET_FILE_PREFIX = {
    "TP": "TP",
    "TN": "TN",
    "NH4+-N": "NH4_N",
    "COD": "COD",
}

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

IMPUTE_COLUMNS = ["Inf.TP", "Inf.COD", "Inf.NH4+-N", "Inf.TN"]

OPERATION_COLUMNS = [
    "Aerobic_DO",
    "Anaerobic_Feed_Fraction",
    "HRT",
    "Anaerobic_Volume_Fraction",
    "Anoxic_Volume_Fraction",
]

MECHANISM_FEATURES = [
    "Anaerobic_HRT",
    "Anoxic_HRT",
    "Aerobic_HRT",
    "C_N_ratio",
    "N_P_ratio",
    "COD_volumetric_load",
    "TN_volumetric_load",
    "NH4_N_volumetric_load",
    "TP_volumetric_load",
]

TARGET_CONFIGS: Dict[str, Dict[str, Any]] = {
    "TP": {
        "sheet": "TP动态可用表",
        "target_col": "eff.TP",
        "history": ["eff.TP_lag1", "eff.TP_lag2", "eff.TP_roll3"],
    },
    "TN": {
        "sheet": "TN动态可用表",
        "target_col": "eff.TN",
        "history": ["eff.TN_lag1", "eff.TN_lag2", "eff.TN_roll3"],
    },
    "NH4+-N": {
        "sheet": "氨氮动态可用表",
        "target_col": "eff.NH4+-N",
        "history": ["eff.NH4+-N_lag1", "eff.NH4+-N_lag2", "eff.NH4+-N_roll3"],
    },
    "COD": {
        "sheet": "COD动态可用表",
        "target_col": "eff.COD",
        "history": ["eff.COD_lag1", "eff.COD_lag2", "eff.COD_roll3"],
    },
}

TEXTUAL_MISSING_VALUES = {
    "", " ", "NA", "N/A", "na", "n/a", "NaN", "nan", "None", "none",
    "NULL", "null", "-", "—", "–", "－", "/", "\\", "未测", "缺失",
}

# 原结果 Excel 中必须存在的正式候选字段。
REQUIRED_FORMAL_COLUMNS = [
    "目标指标",
    "特征组",
    "特征数",
    "模型",
    "候选参数ID",
    "参数JSON",
    "R2_632plus均值",
    "RMSE_632plus均值",
    "MAE_632plus均值",
    "Time_R2",
    "Time_RMSE",
    "Time_MAE",
    "Expanding_R2",
    "Expanding_RMSE",
    "Expanding_MAE",
]


# =============================================================================
# 2. 通用工具
# =============================================================================

def now_str() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def log(message: str) -> None:
    print(f"[{now_str()}] {message}", flush=True)


def clean_column_name(value: Any) -> str:
    text = str(value)
    text = text.replace("\ufeff", "").replace("\xa0", " ").replace("\u3000", " ")
    text = text.replace("\n", "").replace("\r", "").strip()
    return re.sub(r"\s+", "", text)


def convert_to_numeric_series(series: pd.Series) -> pd.Series:
    output = series.copy()
    if output.dtype == "object" or str(output.dtype).startswith("string"):
        output = (
            output.astype("string")
            .str.strip()
            .str.replace(",", "", regex=False)
            .str.replace("，", "", regex=False)
        )
        output = output.mask(output.isin(TEXTUAL_MISSING_VALUES), np.nan)
    return pd.to_numeric(output, errors="coerce")


def safe_float(value: Any) -> float:
    try:
        x = float(value)
    except Exception:
        return np.nan
    return x if np.isfinite(x) else np.nan


def pct_improvement(model_rmse: float, baseline_rmse: float) -> float:
    if not np.isfinite(model_rmse) or not np.isfinite(baseline_rmse) or baseline_rmse <= 0:
        return np.nan
    return float((baseline_rmse - model_rmse) / baseline_rmse * 100.0)


def result_file_for(target_name: str, feature_group: str) -> Path:
    prefix = TARGET_FILE_PREFIX[target_name]
    expected = RESULT_ROOT / feature_group / f"{prefix}_{feature_group}_六模型比较结果.xlsx"
    if expected.exists():
        return expected

    # 兼容目录层级有变化的情况，但优先精确文件名；排除最终输出目录。
    matches = [
        p for p in RESULT_ROOT.rglob(f"{prefix}_{feature_group}_六模型比较结果.xlsx")
        if "zuizhongmoxingqueding" not in str(p).lower()
    ]
    if len(matches) == 1:
        return matches[0]
    if len(matches) > 1:
        raise RuntimeError(
            f"{target_name}-{feature_group} 找到多个同名结果文件，请整理目录后重试：\n"
            + "\n".join(str(x) for x in matches)
        )
    raise FileNotFoundError(
        f"未找到 {target_name}-{feature_group} 结果文件。期望路径：{expected}"
    )


def validate_environment() -> pd.DataFrame:
    rows: List[Dict[str, Any]] = []
    if not RESULT_ROOT.exists():
        raise FileNotFoundError(f"模型结果根目录不存在：{RESULT_ROOT}")
    if not DYNAMIC_INPUT_FILE.exists():
        raise FileNotFoundError(f"动态可用数据文件不存在：{DYNAMIC_INPUT_FILE}")

    for target_name in TARGET_ORDER:
        for fg in FEATURE_GROUP_ORDER:
            path = result_file_for(target_name, fg)
            excel = pd.ExcelFile(path, engine="openpyxl")
            needed_sheets = ["正式候选综合信息", "朴素动态基线"]
            missing_sheets = [s for s in needed_sheets if s not in excel.sheet_names]
            if missing_sheets:
                raise ValueError(f"{path} 缺少工作表：{missing_sheets}")
            header = pd.read_excel(path, sheet_name="正式候选综合信息", nrows=0, engine="openpyxl")
            missing_cols = [c for c in REQUIRED_FORMAL_COLUMNS if c not in header.columns]
            if missing_cols:
                raise ValueError(f"{path} 的“正式候选综合信息”缺少字段：{missing_cols}")
            rows.append({
                "目标指标": target_name,
                "特征组": fg,
                "结果文件": str(path),
                "正式候选工作表": "正常",
                "朴素动态基线工作表": "正常",
            })
    return pd.DataFrame(rows)


def feature_list_for_target(feature_group: str, history_features: Sequence[str]) -> List[str]:
    if feature_group == "F0":
        features = list(RAW_FEATURES)
    elif feature_group == "F1":
        features = list(RAW_FEATURES) + list(MECHANISM_FEATURES)
    elif feature_group == "F2":
        features = list(RAW_FEATURES) + list(history_features)
    elif feature_group == "F3":
        features = list(RAW_FEATURES) + list(MECHANISM_FEATURES) + list(history_features)
    else:
        raise ValueError(f"未知特征组：{feature_group}")

    expected = {"F0": 9, "F1": 18, "F2": 12, "F3": 21}[feature_group]
    if len(features) != expected or len(set(features)) != expected:
        raise RuntimeError(f"{feature_group} 特征数量或重复检查失败：{features}")
    if "Q" in features:
        raise RuntimeError("Q 不允许进入模型特征。")
    return features


# =============================================================================
# 3. 动态数据读取、训练内插补与特征构造
# =============================================================================

def read_target_table(target_name: str) -> pd.DataFrame:
    cfg = TARGET_CONFIGS[target_name]
    excel = pd.ExcelFile(DYNAMIC_INPUT_FILE, engine="openpyxl")
    if cfg["sheet"] not in excel.sheet_names:
        raise ValueError(
            f"动态可用数据缺少工作表 {cfg['sheet']}。现有工作表：{excel.sheet_names}"
        )

    df = pd.read_excel(
        DYNAMIC_INPUT_FILE,
        sheet_name=cfg["sheet"],
        engine="openpyxl",
        na_values=list(TEXTUAL_MISSING_VALUES),
        keep_default_na=True,
    )
    df.columns = [clean_column_name(c) for c in df.columns]

    required = ["样本序号", "Date"] + RAW_FEATURES + [cfg["target_col"]] + list(cfg["history"])
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"{target_name} 动态可用表缺少必要字段：{missing}")

    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    numeric_cols = ["样本序号"] + RAW_FEATURES + [cfg["target_col"]] + list(cfg["history"])
    for col in numeric_cols:
        df[col] = convert_to_numeric_series(df[col])

    df = df.sort_values(["样本序号", "Date"], kind="mergesort").reset_index(drop=True)
    df["_row_id"] = np.arange(len(df), dtype=int)

    if df.empty:
        raise ValueError(f"{target_name} 动态可用表为空。")
    if df["样本序号"].isna().any() or df["Date"].isna().any():
        raise ValueError(f"{target_name} 存在无效样本序号或日期。")
    if df["样本序号"].duplicated().any():
        raise ValueError(f"{target_name} 存在重复样本序号。")
    if not df["样本序号"].is_monotonic_increasing:
        raise ValueError(f"{target_name} 样本未按样本序号递增。")

    must_complete = [cfg["target_col"]] + list(cfg["history"]) + OPERATION_COLUMNS
    missing_counts = {c: int(df[c].isna().sum()) for c in must_complete if df[c].isna().any()}
    if missing_counts:
        raise ValueError(
            f"{target_name} 的目标、历史状态或运行工况存在缺失："
            + json.dumps(missing_counts, ensure_ascii=False)
        )
    if (df["HRT"] <= 0).any():
        raise ValueError(f"{target_name} 存在 HRT<=0。")
    aerobic_fraction = 1.0 - df["Anaerobic_Volume_Fraction"] - df["Anoxic_Volume_Fraction"]
    if (aerobic_fraction <= 0).any():
        raise ValueError(f"{target_name} 存在好氧区容积分数<=0。")
    return df


def interpolate_short_internal_gaps(
    series: pd.Series,
    sequence_coordinates: pd.Series,
    max_gap: int,
) -> pd.Series:
    """仅插补两侧均有训练观测、且按原始样本序号跨度不超过 max_gap 的内部缺口。"""
    output = convert_to_numeric_series(series).astype(float).copy()
    coordinates = convert_to_numeric_series(sequence_coordinates).astype(float).to_numpy()
    values = output.to_numpy(copy=True)
    n = len(values)
    position = 0

    while position < n:
        if np.isfinite(values[position]):
            position += 1
            continue
        start = position
        while position < n and not np.isfinite(values[position]):
            position += 1
        end = position - 1
        left = start - 1
        right = end + 1
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


def fit_train_imputer_and_apply(
    train_raw: pd.DataFrame,
    valid_raw: pd.DataFrame,
) -> Tuple[pd.DataFrame, pd.DataFrame, Dict[str, Any]]:
    """
    与原 F0-F3 正式脚本保持一致：
    - 训练唯一原始样本：短内部缺口线性插值 + 训练中位数兜底；
    - OOB/时间/扩展窗口验证样本：只使用训练中位数，不使用未来观测。
    """
    train = train_raw.copy(deep=True)
    valid = valid_raw.copy(deep=True)

    unique_train = (
        train.drop_duplicates(subset=["_row_id"], keep="first")
        .sort_values(["_row_id"], kind="mergesort")
        .copy()
    )

    info: Dict[str, Any] = {
        "train_medians": {},
        "training_missing_before": {},
        "training_linear_filled": {},
        "training_median_filled": {},
        "validation_missing_filled_by_training_median": {},
    }

    for col in IMPUTE_COLUMNS:
        original = convert_to_numeric_series(unique_train[col]).astype(float)
        median_value = original.median(skipna=True)
        if pd.isna(median_value):
            raise ValueError(f"训练数据变量 {col} 全部缺失，无法插补。")

        interpolated = interpolate_short_internal_gaps(
            original,
            unique_train["样本序号"],
            MAX_INTERPOLATION_GAP,
        )
        linear_filled = int(original.isna().sum() - interpolated.isna().sum())
        median_filled = int(interpolated.isna().sum())
        unique_train[col] = interpolated.fillna(float(median_value))

        mapping = unique_train.set_index("_row_id")[col]
        train[col] = train["_row_id"].map(mapping)

        valid_series = convert_to_numeric_series(valid[col]).astype(float)
        valid_missing = int(valid_series.isna().sum())
        valid[col] = valid_series.fillna(float(median_value))

        info["train_medians"][col] = float(median_value)
        info["training_missing_before"][col] = int(original.isna().sum())
        info["training_linear_filled"][col] = linear_filled
        info["training_median_filled"][col] = median_filled
        info["validation_missing_filled_by_training_median"][col] = valid_missing

    remaining_train = {c: int(train[c].isna().sum()) for c in IMPUTE_COLUMNS if train[c].isna().any()}
    remaining_valid = {c: int(valid[c].isna().sum()) for c in IMPUTE_COLUMNS if valid[c].isna().any()}
    if remaining_train or remaining_valid:
        raise ValueError(
            "插补后仍存在缺失："
            + json.dumps({"train": remaining_train, "valid": remaining_valid}, ensure_ascii=False)
        )
    return train, valid, info


def construct_mechanism_features(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy(deep=True)
    operation_missing = {c: int(out[c].isna().sum()) for c in OPERATION_COLUMNS if out[c].isna().any()}
    if operation_missing:
        raise ValueError("运行工况存在缺失：" + json.dumps(operation_missing, ensure_ascii=False))

    aerobic_fraction = 1.0 - out["Anaerobic_Volume_Fraction"] - out["Anoxic_Volume_Fraction"]
    if (aerobic_fraction <= 0).any():
        raise ValueError("构造机理特征时发现好氧区容积分数<=0。")

    out["Anaerobic_HRT"] = out["HRT"] * out["Anaerobic_Volume_Fraction"]
    out["Anoxic_HRT"] = out["HRT"] * out["Anoxic_Volume_Fraction"]
    out["Aerobic_HRT"] = out["HRT"] * aerobic_fraction
    out["C_N_ratio"] = out["Inf.COD"] / (out["Inf.TN"] + EPS)
    out["N_P_ratio"] = out["Inf.TN"] / (out["Inf.TP"] + EPS)
    out["COD_volumetric_load"] = out["Inf.COD"] / (out["HRT"] + EPS)
    out["TN_volumetric_load"] = out["Inf.TN"] / (out["HRT"] + EPS)
    out["NH4_N_volumetric_load"] = out["Inf.NH4+-N"] / (out["HRT"] + EPS)
    out["TP_volumetric_load"] = out["Inf.TP"] / (out["HRT"] + EPS)
    return out.replace([np.inf, -np.inf], np.nan)


def prepare_split(
    train_raw: pd.DataFrame,
    valid_raw: pd.DataFrame,
    target_name: str,
    feature_group: str,
) -> Dict[str, Any]:
    cfg = TARGET_CONFIGS[target_name]
    features = feature_list_for_target(feature_group, cfg["history"])

    train_imp, valid_imp, impute_info = fit_train_imputer_and_apply(train_raw, valid_raw)

    if feature_group in {"F1", "F3"}:
        train_feat = construct_mechanism_features(train_imp)
        valid_feat = construct_mechanism_features(valid_imp)
    else:
        train_feat = train_imp.copy()
        valid_feat = valid_imp.copy()

    problems: Dict[str, Dict[str, int]] = {}
    for col in features:
        if col not in train_feat.columns or col not in valid_feat.columns:
            problems[col] = {"train": -1, "valid": -1}
            continue
        n_train = int(train_feat[col].isna().sum())
        n_valid = int(valid_feat[col].isna().sum())
        if n_train or n_valid:
            problems[col] = {"train": n_train, "valid": n_valid}
    if problems:
        raise ValueError(
            f"{target_name}-{feature_group} 特征构造后仍有缺失或字段不存在："
            + json.dumps(problems, ensure_ascii=False)
        )

    X_train_raw = train_feat[features].astype(float)
    X_valid_raw = valid_feat[features].astype(float)

    scaler = StandardScaler()
    X_train = scaler.fit_transform(X_train_raw)
    X_valid = scaler.transform(X_valid_raw)

    return {
        "X_train": X_train,
        "X_valid": X_valid,
        "y_train": train_feat[cfg["target_col"]].astype(float).to_numpy(),
        "y_valid": valid_feat[cfg["target_col"]].astype(float).to_numpy(),
        "impute_info": impute_info,
        "feature_names": features,
    }


def prepare_time_and_expanding_cache(
    df: pd.DataFrame,
    target_name: str,
    feature_group: str,
) -> Tuple[Dict[str, Any], List[Dict[str, Any]]]:
    split_idx = int(math.floor(len(df) * TIME_TRAIN_RATIO))
    if split_idx < 10 or len(df) - split_idx < 5:
        raise ValueError(f"{target_name}-{feature_group} 时间验证样本数不足。")

    time_prepared = prepare_split(
        df.iloc[:split_idx].copy(),
        df.iloc[split_idx:].copy(),
        target_name,
        feature_group,
    )

    initial_train = int(math.floor(len(df) * EXPANDING_INITIAL_TRAIN_RATIO))
    windows: List[Dict[str, Any]] = []
    for window_id, start_idx in enumerate(
        range(initial_train, len(df), EXPANDING_STEP_SIZE),
        start=1,
    ):
        end_idx = min(start_idx + EXPANDING_VALID_WINDOW_SIZE, len(df))
        if end_idx - start_idx < 2:
            continue
        train_raw = df.iloc[:start_idx].copy()
        valid_raw = df.iloc[start_idx:end_idx].copy()
        prepared = prepare_split(train_raw, valid_raw, target_name, feature_group)
        prepared["窗口编号"] = window_id
        prepared["验证样本数"] = len(valid_raw)
        windows.append(prepared)

    if not windows:
        raise ValueError(f"{target_name}-{feature_group} 未生成有效扩展窗口。")
    return time_prepared, windows


# =============================================================================
# 4. 六类模型、指标与随机稳定性
# =============================================================================

def build_model(model_name: str, params: Mapping[str, Any], random_state: int):
    p = deepcopy(dict(params))

    if model_name == "ElasticNet":
        return ElasticNet(
            alpha=p["alpha"],
            l1_ratio=p["l1_ratio"],
            max_iter=50000,
            random_state=random_state,
        )

    if model_name == "SVR":
        return SVR(
            kernel=p["kernel"],
            C=p["C"],
            epsilon=p["epsilon"],
            gamma=p["gamma"],
        )

    if model_name == "GPR":
        if p["kernel_type"] == "RBF":
            kernel = ConstantKernel(1.0, constant_value_bounds="fixed") * RBF(
                length_scale=p["length_scale"],
                length_scale_bounds="fixed",
            )
        else:
            kernel = ConstantKernel(1.0, constant_value_bounds="fixed") * Matern(
                length_scale=p["length_scale"],
                length_scale_bounds="fixed",
                nu=1.5,
            )
        return GaussianProcessRegressor(
            kernel=kernel,
            alpha=p["alpha"],
            optimizer=None,
            normalize_y=True,
            random_state=random_state,
        )

    if model_name == "RandomForest":
        return RandomForestRegressor(
            n_estimators=int(p["n_estimators"]),
            max_depth=p["max_depth"],
            min_samples_leaf=int(p["min_samples_leaf"]),
            min_samples_split=int(p["min_samples_split"]),
            max_features=p["max_features"],
            random_state=random_state,
            n_jobs=N_JOBS_TREE,
        )

    if model_name == "XGBoost":
        if not XGBOOST_AVAILABLE:
            raise ImportError("未安装 xgboost，请执行：pip install xgboost")
        return XGBRegressor(
            objective="reg:squarederror",
            n_estimators=int(p["n_estimators"]),
            max_depth=int(p["max_depth"]),
            learning_rate=float(p["learning_rate"]),
            subsample=float(p["subsample"]),
            colsample_bytree=float(p["colsample_bytree"]),
            random_state=random_state,
            n_jobs=N_JOBS_TREE,
            reg_lambda=1.0,
            verbosity=0,
        )

    if model_name == "MLP":
        hidden = p["hidden_layer_sizes"]
        if isinstance(hidden, list):
            hidden = tuple(hidden)
        return MLPRegressor(
            hidden_layer_sizes=hidden,
            activation=p["activation"],
            alpha=float(p["alpha"]),
            learning_rate_init=float(p["learning_rate_init"]),
            batch_size=int(p.get("batch_size", 32)),
            early_stopping=bool(p.get("early_stopping", True)),
            validation_fraction=0.15,
            max_iter=int(p.get("max_iter", 1000)),
            n_iter_no_change=int(p.get("n_iter_no_change", 30)),
            random_state=random_state,
        )

    raise ValueError(f"未知模型：{model_name}")


def fit_model_safely(model: Any, X_train: np.ndarray, y_train: np.ndarray) -> str:
    with warnings.catch_warnings(record=True) as caught:
        warnings.simplefilter("always")
        warnings.filterwarnings("always", category=ConvergenceWarning)
        model.fit(X_train, y_train)
    if not caught:
        return ""
    return " | ".join(str(item.message)[:300] for item in caught[:5])


def calc_metrics(y_true: Iterable[float], y_pred: Iterable[float]) -> Dict[str, float]:
    true = np.asarray(list(y_true), dtype=float)
    pred = np.asarray(list(y_pred), dtype=float)
    mask = np.isfinite(true) & np.isfinite(pred)
    true = true[mask]
    pred = pred[mask]
    n = len(true)

    if n == 0:
        return {"n": 0, "R2": np.nan, "MSE": np.nan, "RMSE": np.nan, "MAE": np.nan}

    mse = float(mean_squared_error(true, pred))
    rmse = float(math.sqrt(mse))
    mae = float(mean_absolute_error(true, pred))
    r2 = float(r2_score(true, pred)) if n >= 2 and np.nanstd(true) > 0 else np.nan
    return {"n": int(n), "R2": r2, "MSE": mse, "RMSE": rmse, "MAE": mae}


def evaluate_stochastic_candidates(
    target_name: str,
    feature_group: str,
    candidate_rows: pd.DataFrame,
    target_df: pd.DataFrame,
) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """对少量随机近优候选使用相同 5 个随机种子重新执行 Time + Expanding。"""
    if candidate_rows.empty:
        return pd.DataFrame(), pd.DataFrame()

    time_prepared, windows = prepare_time_and_expanding_cache(
        target_df,
        target_name,
        feature_group,
    )

    detail_rows: List[Dict[str, Any]] = []

    for _, candidate in candidate_rows.iterrows():
        model_name = str(candidate["模型"])
        candidate_id = str(candidate["候选参数ID"])
        params = json.loads(str(candidate["参数JSON"]))

        for repeat_id, seed in enumerate(STABILITY_SEEDS, start=1):
            start = time.time()
            warning_messages: List[str] = []
            try:
                model = build_model(model_name, params, int(seed))
                warning_text = fit_model_safely(
                    model,
                    time_prepared["X_train"],
                    time_prepared["y_train"],
                )
                if warning_text:
                    warning_messages.append("Time:" + warning_text)
                time_pred = model.predict(time_prepared["X_valid"])
                time_metrics = calc_metrics(time_prepared["y_valid"], time_pred)

                all_true: List[float] = []
                all_pred: List[float] = []
                for window in windows:
                    model_w = build_model(model_name, params, int(seed))
                    warning_w = fit_model_safely(
                        model_w,
                        window["X_train"],
                        window["y_train"],
                    )
                    if warning_w and len(warning_messages) < 6:
                        warning_messages.append(f"Window{window['窗口编号']}:" + warning_w)
                    pred_w = model_w.predict(window["X_valid"])
                    all_true.extend(window["y_valid"].tolist())
                    all_pred.extend(pred_w.tolist())

                exp_metrics = calc_metrics(all_true, all_pred)

                detail_rows.append({
                    "目标指标": target_name,
                    "特征组": feature_group,
                    "模型": model_name,
                    "候选参数ID": candidate_id,
                    "参数JSON": candidate["参数JSON"],
                    "稳定性重复": repeat_id,
                    "随机种子": int(seed),
                    "Time_R2": time_metrics["R2"],
                    "Time_RMSE": time_metrics["RMSE"],
                    "Time_MAE": time_metrics["MAE"],
                    "Expanding_R2": exp_metrics["R2"],
                    "Expanding_RMSE": exp_metrics["RMSE"],
                    "Expanding_MAE": exp_metrics["MAE"],
                    "警告数": len(warning_messages),
                    "警告信息": " || ".join(warning_messages),
                    "是否成功": "是",
                    "错误信息": "",
                    "耗时秒": time.time() - start,
                })
            except Exception as exc:
                detail_rows.append({
                    "目标指标": target_name,
                    "特征组": feature_group,
                    "模型": model_name,
                    "候选参数ID": candidate_id,
                    "参数JSON": candidate["参数JSON"],
                    "稳定性重复": repeat_id,
                    "随机种子": int(seed),
                    "是否成功": "否",
                    "错误信息": repr(exc) + "\n" + traceback.format_exc(limit=3),
                    "耗时秒": time.time() - start,
                })

    detail = pd.DataFrame(detail_rows)
    success = detail.loc[detail["是否成功"] == "是"].copy()
    if success.empty:
        return detail, pd.DataFrame()

    summary_rows: List[Dict[str, Any]] = []
    for (model_name, candidate_id), group in success.groupby(["模型", "候选参数ID"], sort=False):
        row: Dict[str, Any] = {
            "目标指标": target_name,
            "特征组": feature_group,
            "模型": model_name,
            "候选参数ID": candidate_id,
            "参数JSON": group["参数JSON"].iloc[0],
            "成功重复次数": int(group["稳定性重复"].nunique()),
            "计划重复次数": STABILITY_REPEATS,
            "随机种子列表": ",".join(str(int(x)) for x in sorted(group["随机种子"].unique())),
            "警告重复次数": int((pd.to_numeric(group["警告数"], errors="coerce").fillna(0) > 0).sum()),
        }
        for metric in [
            "Time_R2", "Time_RMSE", "Time_MAE",
            "Expanding_R2", "Expanding_RMSE", "Expanding_MAE",
        ]:
            values = pd.to_numeric(group[metric], errors="coerce").dropna()
            row[f"{metric}均值"] = values.mean() if len(values) else np.nan
            row[f"{metric}标准差"] = values.std(ddof=1) if len(values) > 1 else 0.0
            row[f"{metric}最小值"] = values.min() if len(values) else np.nan
            row[f"{metric}最大值"] = values.max() if len(values) else np.nan
        summary_rows.append(row)

    summary = pd.DataFrame(summary_rows)
    return detail, summary


# =============================================================================
# 5. 读取已有正式结果、统一候选评价与组内代表
# =============================================================================

def load_formal_candidates_and_baseline(
    target_name: str,
    feature_group: str,
) -> Tuple[pd.DataFrame, pd.DataFrame, Path]:
    path = result_file_for(target_name, feature_group)
    formal = pd.read_excel(path, sheet_name="正式候选综合信息", engine="openpyxl")
    baseline = pd.read_excel(path, sheet_name="朴素动态基线", engine="openpyxl")

    missing_cols = [c for c in REQUIRED_FORMAL_COLUMNS if c not in formal.columns]
    if missing_cols:
        raise ValueError(f"{path} 正式候选综合信息缺少字段：{missing_cols}")

    formal = formal.copy()
    numeric_cols = [
        "特征数", "R2_632plus均值", "RMSE_632plus均值", "MAE_632plus均值",
        "Time_R2", "Time_RMSE", "Time_MAE",
        "Expanding_R2", "Expanding_RMSE", "Expanding_MAE",
        "Train_Time_R2差距",
    ]
    for col in numeric_cols:
        if col in formal.columns:
            formal[col] = pd.to_numeric(formal[col], errors="coerce")

    formal = formal.loc[
        formal["模型"].isin(MODEL_ORDER)
        & formal["Expanding_RMSE"].notna()
        & formal["Time_RMSE"].notna()
        & formal["RMSE_632plus均值"].notna()
    ].copy()

    if formal.empty:
        raise RuntimeError(f"{target_name}-{feature_group} 没有可用于统一选择的正式候选。")

    # 防止同一文件中同一候选ID重复，避免最终合并时产生一对多。
    dup = formal.duplicated(subset=["目标指标", "特征组", "模型", "候选参数ID"], keep=False)
    if dup.any():
        preview = formal.loc[dup, ["目标指标", "特征组", "模型", "候选参数ID"]].head(20)
        raise RuntimeError(
            f"{target_name}-{feature_group} 正式候选存在重复键，停止终选：\n"
            + preview.to_string(index=False)
        )

    return formal, baseline, path


def mark_near_optimal_within_algorithm(formal: pd.DataFrame) -> pd.DataFrame:
    """
    仅用于生成“目标×F组×算法”的24个组内代表配置，供消融与算法比较使用。
    最终目标模型不再受24代表配置限制。
    """
    rows: List[pd.DataFrame] = []
    for model_name in MODEL_ORDER:
        group = formal.loc[formal["模型"] == model_name].copy()
        if group.empty:
            raise RuntimeError(f"正式候选中缺少算法：{model_name}")

        best_exp = float(group["Expanding_RMSE"].min())
        threshold = best_exp * (1.0 + NEAR_OPTIMAL_REL_TOL)
        group["算法内最低Expanding_RMSE"] = best_exp
        group["算法内5%近优上限"] = threshold
        group["算法内Expanding相对最佳差异百分比"] = (
            (group["Expanding_RMSE"] - best_exp) / best_exp * 100.0
            if best_exp > 0 else np.nan
        )
        group["是否算法内5%近优"] = np.where(
            group["Expanding_RMSE"] <= threshold + 1e-15,
            "是",
            "否",
        )
        rows.append(group)
    return pd.concat(rows, ignore_index=True)


def deterministic_candidate_to_unified(row: pd.Series) -> Dict[str, Any]:
    return {
        "目标指标": row["目标指标"],
        "特征组": row["特征组"],
        "特征数": row["特征数"],
        "模型": row["模型"],
        "候选参数ID": row["候选参数ID"],
        "参数JSON": row["参数JSON"],
        "Bootstrap_R2": row["R2_632plus均值"],
        "Bootstrap_RMSE": row["RMSE_632plus均值"],
        "Bootstrap_MAE": row["MAE_632plus均值"],
        "原正式Time_R2": row["Time_R2"],
        "原正式Time_RMSE": row["Time_RMSE"],
        "原正式Expanding_R2": row["Expanding_R2"],
        "原正式Expanding_RMSE": row["Expanding_RMSE"],
        "Eval_Time_R2": row["Time_R2"],
        "Eval_Time_RMSE": row["Time_RMSE"],
        "Eval_Time_MAE": row["Time_MAE"],
        "Eval_Expanding_R2": row["Expanding_R2"],
        "Eval_Expanding_RMSE": row["Expanding_RMSE"],
        "Eval_Expanding_MAE": row["Expanding_MAE"],
        "Time_RMSE稳定性标准差": 0.0,
        "Expanding_RMSE稳定性标准差": 0.0,
        "稳定性重复次数": 0,
        "稳定性随机种子": "",
        "稳定性警告重复次数": 0,
        "评价指标来源": "确定性模型：原正式Time/Expanding",
    }


def stochastic_candidate_to_unified(
    summary_row: pd.Series,
    original_row: pd.Series,
) -> Dict[str, Any]:
    return {
        "目标指标": original_row["目标指标"],
        "特征组": original_row["特征组"],
        "特征数": original_row["特征数"],
        "模型": original_row["模型"],
        "候选参数ID": original_row["候选参数ID"],
        "参数JSON": original_row["参数JSON"],
        "Bootstrap_R2": original_row["R2_632plus均值"],
        "Bootstrap_RMSE": original_row["RMSE_632plus均值"],
        "Bootstrap_MAE": original_row["MAE_632plus均值"],
        "原正式Time_R2": original_row["Time_R2"],
        "原正式Time_RMSE": original_row["Time_RMSE"],
        "原正式Expanding_R2": original_row["Expanding_R2"],
        "原正式Expanding_RMSE": original_row["Expanding_RMSE"],
        # 随机模型最终比较统一使用5随机种子均值，而不是原来某一个随机种子。
        "Eval_Time_R2": summary_row["Time_R2均值"],
        "Eval_Time_RMSE": summary_row["Time_RMSE均值"],
        "Eval_Time_MAE": summary_row["Time_MAE均值"],
        "Eval_Expanding_R2": summary_row["Expanding_R2均值"],
        "Eval_Expanding_RMSE": summary_row["Expanding_RMSE均值"],
        "Eval_Expanding_MAE": summary_row["Expanding_MAE均值"],
        "Time_RMSE稳定性标准差": summary_row["Time_RMSE标准差"],
        "Expanding_RMSE稳定性标准差": summary_row["Expanding_RMSE标准差"],
        "稳定性重复次数": summary_row["成功重复次数"],
        "稳定性随机种子": summary_row["随机种子列表"],
        "稳定性警告重复次数": summary_row["警告重复次数"],
        "评价指标来源": "随机模型：同5随机种子Time/Expanding均值",
    }


def build_all_evaluated_candidates(
    formal_all: pd.DataFrame,
    stability_summary: pd.DataFrame,
) -> pd.DataFrame:
    """
    构造最终目标级直选的完整候选池。

    关键原则：
    - 每一条进入“正式候选综合信息”的候选都必须进入统一候选池；
    - 确定性模型直接使用原正式Time/Expanding；
    - 随机模型必须存在完整5随机种子稳定性结果，再使用均值替代单随机种子结果；
    - 不允许24代表配置成为最终候选的硬门槛。
    """
    rows: List[Dict[str, Any]] = []

    deterministic = formal_all.loc[formal_all["模型"].isin(DETERMINISTIC_MODELS)].copy()
    for _, row in deterministic.iterrows():
        rows.append(deterministic_candidate_to_unified(row))

    stochastic_formal = formal_all.loc[formal_all["模型"].isin(STOCHASTIC_MODELS)].copy()

    # 每一个正式随机候选都必须完成稳定性复核。
    expected_keys = set(
        tuple(x)
        for x in stochastic_formal[
            ["目标指标", "特征组", "模型", "候选参数ID"]
        ].astype(str).itertuples(index=False, name=None)
    )
    actual_keys = set(
        tuple(x)
        for x in stability_summary[
            ["目标指标", "特征组", "模型", "候选参数ID"]
        ].astype(str).itertuples(index=False, name=None)
    )
    missing_keys = sorted(expected_keys - actual_keys)
    extra_keys = sorted(actual_keys - expected_keys)
    if missing_keys or extra_keys:
        raise RuntimeError(
            "随机稳定性汇总与正式随机候选无法一一对应。\n"
            f"缺失稳定性候选数={len(missing_keys)}，额外稳定性候选数={len(extra_keys)}。\n"
            f"缺失前10={missing_keys[:10]}\n额外前10={extra_keys[:10]}"
        )

    for _, summary_row in stability_summary.iterrows():
        if int(summary_row["成功重复次数"]) != STABILITY_REPEATS:
            raise RuntimeError(
                f"{summary_row['目标指标']}-{summary_row['特征组']}-"
                f"{summary_row['模型']}-{summary_row['候选参数ID']} "
                f"稳定性成功重复数不是{STABILITY_REPEATS}。"
            )

        match = stochastic_formal.loc[
            (stochastic_formal["目标指标"].astype(str) == str(summary_row["目标指标"]))
            & (stochastic_formal["特征组"].astype(str) == str(summary_row["特征组"]))
            & (stochastic_formal["模型"].astype(str) == str(summary_row["模型"]))
            & (stochastic_formal["候选参数ID"].astype(str) == str(summary_row["候选参数ID"]))
        ]
        if len(match) != 1:
            raise RuntimeError(
                f"随机候选原正式记录匹配异常："
                f"{summary_row['目标指标']}-{summary_row['特征组']}-"
                f"{summary_row['模型']}-{summary_row['候选参数ID']}，匹配行数={len(match)}"
            )
        rows.append(stochastic_candidate_to_unified(summary_row, match.iloc[0]))

    evaluated = pd.DataFrame(rows)

    # 完整性检查：统一候选池必须和原正式候选一行不多、一行不少。
    if len(evaluated) != len(formal_all):
        raise RuntimeError(
            f"统一评价候选总数应等于正式候选总数{len(formal_all)}，当前为{len(evaluated)}。"
        )

    key_cols = ["目标指标", "特征组", "模型", "候选参数ID"]
    if evaluated.duplicated(subset=key_cols).any():
        preview = evaluated.loc[
            evaluated.duplicated(subset=key_cols, keep=False),
            key_cols,
        ].head(20)
        raise RuntimeError("统一评价候选出现重复键：\n" + preview.to_string(index=False))

    required_eval = [
        "Bootstrap_R2", "Bootstrap_RMSE",
        "Eval_Time_R2", "Eval_Time_RMSE", "Eval_Time_MAE",
        "Eval_Expanding_R2", "Eval_Expanding_RMSE", "Eval_Expanding_MAE",
    ]
    missing_eval = {
        c: int(pd.to_numeric(evaluated[c], errors="coerce").isna().sum())
        for c in required_eval
        if pd.to_numeric(evaluated[c], errors="coerce").isna().any()
    }
    if missing_eval:
        raise RuntimeError("统一评价候选存在缺失评价指标：" + json.dumps(missing_eval, ensure_ascii=False))

    return evaluated


def _selection_sort_columns() -> Tuple[List[str], List[bool]]:
    """
    全项目统一的排序规则。
    该顺序同时用于四个污染物的目标级最终选择，不针对单个目标改变。
    """
    cols = [
        "Eval_Time_RMSE",
        "Time_RMSE稳定性标准差",
        "Bootstrap_RMSE",
        "Eval_Time_MAE",
        "Eval_Expanding_MAE",
        "Expanding_RMSE稳定性标准差",
        "特征数",
        "候选参数ID",
    ]
    return cols, [True] * len(cols)


def select_group_representatives(
    all_evaluated_candidates: pd.DataFrame,
) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """
    生成每个“目标×F组×算法”的代表配置，共4×4×6=96行。

    这些代表仅用于：
    - F0-F3消融结果展示；
    - 六算法横向比较；
    - 结果表压缩展示。

    它们不再作为目标最终模型的硬性候选入口。
    """
    annotated_parts: List[pd.DataFrame] = []
    rep_rows: List[pd.Series] = []
    sort_cols, ascending = _selection_sort_columns()

    for target_name in TARGET_ORDER:
        for fg in FEATURE_GROUP_ORDER:
            for model_name in MODEL_ORDER:
                group = all_evaluated_candidates.loc[
                    (all_evaluated_candidates["目标指标"] == target_name)
                    & (all_evaluated_candidates["特征组"] == fg)
                    & (all_evaluated_candidates["模型"] == model_name)
                ].copy()
                if group.empty:
                    raise RuntimeError(f"{target_name}-{fg}-{model_name} 统一评价候选为空。")

                best_exp = float(group["Eval_Expanding_RMSE"].min())
                threshold = best_exp * (1.0 + NEAR_OPTIMAL_REL_TOL)
                group["组内最低Eval_Expanding_RMSE"] = best_exp
                group["组内5%近优上限"] = threshold
                group["组内Expanding相对最佳差异百分比"] = (
                    (group["Eval_Expanding_RMSE"] - best_exp) / best_exp * 100.0
                    if best_exp > 0 else np.nan
                )
                group["是否组内5%近优"] = np.where(
                    group["Eval_Expanding_RMSE"] <= threshold + 1e-15,
                    "是",
                    "否",
                )

                pool = group.loc[group["是否组内5%近优"] == "是"].copy()
                pool = pool.sort_values(sort_cols, ascending=ascending, na_position="last")
                best = pool.iloc[0].copy()
                best["组内代表选择规则"] = (
                    f"组内Eval_Expanding_RMSE最低值的{NEAR_OPTIMAL_REL_TOL*100:.1f}%以内近优；"
                    "近优后按统一Time优先规则选择。仅用于组内比较，不限制目标最终候选池。"
                )
                best["组内5%近优配置数"] = len(pool)
                rep_rows.append(best)

                # 审计排序：先近优候选，再其余候选按Expanding排序。
                near_sorted = pool.copy()
                nonnear = group.loc[group["是否组内5%近优"] != "是"].copy().sort_values(
                    ["Eval_Expanding_RMSE", "Eval_Time_RMSE", "Bootstrap_RMSE", "候选参数ID"],
                    ascending=[True, True, True, True],
                    na_position="last",
                )
                ordered = pd.concat([near_sorted, nonnear], ignore_index=True)
                ordered["组内统一排序"] = np.arange(1, len(ordered) + 1)
                annotated_parts.append(ordered)

    reps = pd.DataFrame(rep_rows).reset_index(drop=True)
    if len(reps) != 96:
        raise RuntimeError(f"组内代表配置应为96行（4目标×4F组×6算法），当前为{len(reps)}。")

    return pd.concat(annotated_parts, ignore_index=True), reps


# =============================================================================
# 6. 四个目标使用完全相同规则直接终选
# =============================================================================

def select_final_ml_per_target(
    all_evaluated_candidates: pd.DataFrame,
) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """
    四个污染物使用完全相同的目标级最终选择规则。

    注意：
    最终模型直接从该目标“全部正式复核候选”的统一评价表中选择，
    不先压缩到24个组内代表，因此彻底消除两层5%筛选导致的候选丢失。
    """
    annotated_parts: List[pd.DataFrame] = []
    near_parts: List[pd.DataFrame] = []
    final_rows: List[pd.Series] = []
    sort_cols, ascending = _selection_sort_columns()

    for target_name in TARGET_ORDER:
        group = all_evaluated_candidates.loc[
            all_evaluated_candidates["目标指标"] == target_name
        ].copy()
        if group.empty:
            raise RuntimeError(f"{target_name} 没有统一评价候选。")

        best_exp = float(group["Eval_Expanding_RMSE"].min())
        threshold = best_exp * (1.0 + NEAR_OPTIMAL_REL_TOL)

        group["目标内最低Eval_Expanding_RMSE"] = best_exp
        group["目标内5%近优上限"] = threshold
        group["目标内Expanding相对最佳差异百分比"] = (
            (group["Eval_Expanding_RMSE"] - best_exp) / best_exp * 100.0
            if best_exp > 0 else np.nan
        )
        group["是否目标内5%近优"] = np.where(
            group["Eval_Expanding_RMSE"] <= threshold + 1e-15,
            "是",
            "否",
        )

        pool = group.loc[group["是否目标内5%近优"] == "是"].copy()
        if pool.empty:
            raise RuntimeError(f"{target_name} 目标级5%近优候选为空。")

        pool = pool.sort_values(sort_cols, ascending=ascending, na_position="last")
        best = pool.iloc[0].copy()
        best["最终目标级选择规则"] = (
            f"四个目标统一：全部正式候选直接比较；"
            f"Eval_Expanding_RMSE在目标最低值{NEAR_OPTIMAL_REL_TOL*100:.1f}%以内视为近优；"
            "近优中依次比较Eval_Time_RMSE、Time RMSE稳定性SD、Bootstrap RMSE、"
            "Time MAE、Expanding MAE、Expanding RMSE稳定性SD、特征数。"
        )
        best["目标内5%近优配置数"] = len(pool)
        final_rows.append(best)

        pool_out = pool.copy()
        pool_out["目标级近优内部排序"] = np.arange(1, len(pool_out) + 1)
        near_parts.append(pool_out)

        # 全候选审计排序：目标级5%近优候选排在最前；其余按Expanding距离排序。
        nonnear = group.loc[group["是否目标内5%近优"] != "是"].copy().sort_values(
            ["Eval_Expanding_RMSE", "Eval_Time_RMSE", "Bootstrap_RMSE", "特征数", "候选参数ID"],
            ascending=[True, True, True, True, True],
            na_position="last",
        )
        ordered = pd.concat([pool, nonnear], ignore_index=True)
        ordered["目标内统一排序"] = np.arange(1, len(ordered) + 1)
        annotated_parts.append(ordered)

    final_df = pd.DataFrame(final_rows).reset_index(drop=True)
    if len(final_df) != len(TARGET_ORDER):
        raise RuntimeError(f"最终机器学习模型应为4个，当前为{len(final_df)}。")

    return (
        pd.concat(annotated_parts, ignore_index=True),
        pd.concat(near_parts, ignore_index=True),
        final_df,
    )


# =============================================================================
# 7. 基线读取与最终结论
# =============================================================================

def collect_baselines() -> pd.DataFrame:
    rows: List[pd.DataFrame] = []
    for target_name in TARGET_ORDER:
        # F0-F3样本一致，基线理论上完全相同；读取F0并同时核对其他F组。
        reference: pd.DataFrame | None = None
        for fg in FEATURE_GROUP_ORDER:
            _, baseline, path = load_formal_candidates_and_baseline(target_name, fg)
            baseline = baseline.copy()
            baseline["来源特征组"] = fg
            baseline["来源文件"] = str(path)
            if reference is None:
                reference = baseline.drop(columns=["来源特征组", "来源文件"]).copy()
            else:
                # 数值级核对，不因浮点显示微差而误报。
                key_cols = ["验证方式", "基线"]
                ref2 = reference[key_cols + ["R2", "RMSE", "MAE"]].copy().sort_values(key_cols).reset_index(drop=True)
                cur2 = baseline[key_cols + ["R2", "RMSE", "MAE"]].copy().sort_values(key_cols).reset_index(drop=True)
                if len(ref2) != len(cur2):
                    raise RuntimeError(f"{target_name} 不同F组的朴素基线行数不一致。")
                for col in ["R2", "RMSE", "MAE"]:
                    a = pd.to_numeric(ref2[col], errors="coerce").to_numpy(dtype=float)
                    b = pd.to_numeric(cur2[col], errors="coerce").to_numpy(dtype=float)
                    if not np.allclose(a, b, rtol=1e-10, atol=1e-12, equal_nan=True):
                        raise RuntimeError(f"{target_name} 的 {fg} 基线与F0不一致，停止终选。")
        assert reference is not None
        reference["目标指标"] = target_name
        rows.append(reference)
    return pd.concat(rows, ignore_index=True)


def build_final_baseline_comparison(
    final_ml: pd.DataFrame,
    baselines: pd.DataFrame,
) -> Tuple[pd.DataFrame, pd.DataFrame]:
    compare_rows: List[Dict[str, Any]] = []
    conclusion_rows: List[Dict[str, Any]] = []

    for _, model in final_ml.iterrows():
        target_name = model["目标指标"]
        b = baselines.loc[baselines["目标指标"] == target_name].copy()

        time_b = b.loc[b["验证方式"].astype(str).str.contains("时间")].copy()
        exp_b = b.loc[b["验证方式"].astype(str).str.contains("扩展")].copy()
        if time_b.empty or exp_b.empty:
            raise RuntimeError(f"{target_name} 缺少时间或扩展窗口朴素基线。")

        time_b["RMSE"] = pd.to_numeric(time_b["RMSE"], errors="coerce")
        exp_b["RMSE"] = pd.to_numeric(exp_b["RMSE"], errors="coerce")
        best_time_baseline = time_b.sort_values("RMSE").iloc[0]
        best_exp_baseline = exp_b.sort_values("RMSE").iloc[0]

        for _, base in b.iterrows():
            validation_type = str(base["验证方式"])
            if "时间" in validation_type:
                model_rmse = safe_float(model["Eval_Time_RMSE"])
                model_r2 = safe_float(model["Eval_Time_R2"])
            elif "扩展" in validation_type:
                model_rmse = safe_float(model["Eval_Expanding_RMSE"])
                model_r2 = safe_float(model["Eval_Expanding_R2"])
            else:
                continue
            baseline_rmse = safe_float(base["RMSE"])
            compare_rows.append({
                "目标指标": target_name,
                "最终ML特征组": model["特征组"],
                "最终ML模型": model["模型"],
                "最终ML候选参数ID": model["候选参数ID"],
                "验证方式": validation_type,
                "基线": base["基线"],
                "模型R2": model_r2,
                "模型RMSE": model_rmse,
                "基线R2": safe_float(base["R2"]),
                "基线RMSE": baseline_rmse,
                "模型相对基线RMSE降低率百分比": pct_improvement(model_rmse, baseline_rmse),
                "模型是否优于该基线": "是" if np.isfinite(model_rmse) and np.isfinite(baseline_rmse) and model_rmse < baseline_rmse else "否",
            })

        beats_best_time = safe_float(model["Eval_Time_RMSE"]) < safe_float(best_time_baseline["RMSE"])
        beats_best_exp = safe_float(model["Eval_Expanding_RMSE"]) < safe_float(best_exp_baseline["RMSE"])
        stable_win = bool(beats_best_time and beats_best_exp)

        if stable_win:
            formal_prediction = f"采用 {model['特征组']}-{model['模型']}（{model['候选参数ID']}）作为最终预测模型。"
        else:
            formal_prediction = (
                f"最佳机器学习模型为 {model['特征组']}-{model['模型']}（{model['候选参数ID']}），"
                f"但未在 Time 与 Expanding 两种验证中同时超过最佳历史基线；"
                f"预测性能对照中同时保留 Time 最佳基线“{best_time_baseline['基线']}”与 "
                f"Expanding 最佳基线“{best_exp_baseline['基线']}”。"
            )

        conclusion_rows.append({
            "目标指标": target_name,
            "最佳机器学习特征组": model["特征组"],
            "最佳机器学习模型": model["模型"],
            "最佳机器学习候选参数ID": model["候选参数ID"],
            "参数JSON": model["参数JSON"],
            "Bootstrap_R2": model["Bootstrap_R2"],
            "Bootstrap_RMSE": model["Bootstrap_RMSE"],
            "Time_R2": model["Eval_Time_R2"],
            "Time_RMSE": model["Eval_Time_RMSE"],
            "Expanding_R2": model["Eval_Expanding_R2"],
            "Expanding_RMSE": model["Eval_Expanding_RMSE"],
            "Time_RMSE稳定性标准差": model.get("Time_RMSE稳定性标准差", np.nan),
            "Expanding_RMSE稳定性标准差": model.get("Expanding_RMSE稳定性标准差", np.nan),
            "Time最佳历史基线": best_time_baseline["基线"],
            "Time最佳基线R2": best_time_baseline["R2"],
            "Time最佳基线RMSE": best_time_baseline["RMSE"],
            "Expanding最佳历史基线": best_exp_baseline["基线"],
            "Expanding最佳基线R2": best_exp_baseline["R2"],
            "Expanding最佳基线RMSE": best_exp_baseline["RMSE"],
            "ML是否在两种时间验证中均优于最佳历史基线": "是" if stable_win else "否",
            "最终预测结论": formal_prediction,
        })

    return pd.DataFrame(compare_rows), pd.DataFrame(conclusion_rows)


# =============================================================================
# 8. 输出格式与完整TXT说明
# =============================================================================

def sanitize_dataframe_for_excel(df: pd.DataFrame) -> pd.DataFrame:
    if df is None:
        return pd.DataFrame()
    out = df.copy()
    for col in out.columns:
        if out[col].dtype == "object":
            out[col] = out[col].map(
                lambda x: json.dumps(x, ensure_ascii=False, default=str)
                if isinstance(x, (list, tuple, dict, set)) else x
            )
    return out


def format_excel_workbook(writer: pd.ExcelWriter) -> None:
    workbook = writer.book
    for sheet_name in workbook.sheetnames:
        worksheet = workbook[sheet_name]
        if worksheet.max_row >= 1:
            worksheet.freeze_panes = "A2"
            worksheet.auto_filter.ref = worksheet.dimensions
            for cell in worksheet[1]:
                cell.font = Font(bold=True)
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        worksheet.sheet_view.showGridLines = False
        for row in worksheet.iter_rows(min_row=2):
            for cell in row:
                cell.alignment = Alignment(vertical="center", wrap_text=False)
                if isinstance(cell.value, datetime):
                    cell.number_format = "yyyy-mm-dd"
                elif isinstance(cell.value, float):
                    cell.number_format = "0.000000"
        for col_idx in range(1, worksheet.max_column + 1):
            letter = get_column_letter(col_idx)
            max_length = 0
            for cell in worksheet[letter][: min(worksheet.max_row, 400)]:
                value = "" if cell.value is None else str(cell.value)
                max_length = max(max_length, len(value))
            worksheet.column_dimensions[letter].width = min(max(max_length + 2, 10), 52)


def safe_write_excel(path: Path, tables: Mapping[str, pd.DataFrame]) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)

    def _write(target: Path) -> None:
        with pd.ExcelWriter(target, engine="openpyxl", datetime_format="yyyy-mm-dd") as writer:
            for sheet_name, table in tables.items():
                sanitize_dataframe_for_excel(table).to_excel(
                    writer,
                    sheet_name=str(sheet_name)[:31],
                    index=False,
                )
            format_excel_workbook(writer)

    try:
        _write(path)
        return path
    except PermissionError:
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        alt = path.with_name(f"{path.stem}_{stamp}{path.suffix}")
        _write(alt)
        return alt


def fmt(x: Any, digits: int = 6) -> str:
    v = safe_float(x)
    return "NA" if not np.isfinite(v) else f"{v:.{digits}f}"


def write_complete_text_report(
    output_path: Path,
    start_time: datetime,
    input_check: pd.DataFrame,
    formal_all: pd.DataFrame,
    stability_pool: pd.DataFrame,
    stability_summary: pd.DataFrame,
    all_evaluated_candidates: pd.DataFrame,
    representatives_24: pd.DataFrame,
    target_ranked: pd.DataFrame,
    target_near: pd.DataFrame,
    final_ml: pd.DataFrame,
    baselines: pd.DataFrame,
    baseline_compare: pd.DataFrame,
    final_conclusions: pd.DataFrame,
) -> Path:
    lines: List[str] = []
    add = lines.append

    add("A²/O F0-F3 最终模型统一终选与随机稳定性复核——完整说明")
    add("=" * 108)
    add("")
    add("一、运行与统一选择设置")
    add(f"1. 开始时间：{start_time.strftime('%Y-%m-%d %H:%M:%S')}")
    add(f"2. 完成时间：{now_str()}")
    add(f"3. 原F0-F3结果根目录：{RESULT_ROOT}")
    add(f"4. 动态可用数据：{DYNAMIC_INPUT_FILE}")
    add(f"5. 输出目录：{OUTPUT_DIR}")
    add(f"6. 目标级Expanding RMSE近优阈值：最佳值 + {NEAR_OPTIMAL_REL_TOL*100:.1f}%")
    add(f"7. 随机模型稳定性重复：{STABILITY_REPEATS}次；所有随机候选使用同一组随机种子：{','.join(str(x) for x in STABILITY_SEEDS)}")
    add("8. 不重新运行1000次Bootstrap；Bootstrap .632+指标直接读取已有正式候选结果。")
    add("9. RandomForest、XGBoost、MLP：所有进入正式复核的候选参数都重新执行同5随机种子Time + Expanding。")
    add("10. TP、TN、NH4+-N、COD四个目标严格使用同一套最终模型选择规则，不对单个目标改变阈值、权重或排序顺序。")
    add("")

    add("二、本版最终修正")
    add("1. 完整指标统一输出：Bootstrap .632+ R²/RMSE、Time R²/RMSE、Expanding R²/RMSE。")
    add("2. 随机模型不再由某一个随机种子的表现决定；最终Time/Expanding使用5随机种子均值并记录标准差。")
    add("3. 取消“两层5%筛选”的最终模型门槛：")
    add("   24个“F组×算法”组内代表仍保留，用于消融与算法比较；")
    add("   但每个污染物的最终模型直接从该目标全部正式复核候选的统一评价表中选择。")
    add("4. 四个目标统一终选顺序：")
    add(f"   (a) 先保留Eval_Expanding_RMSE不超过该目标最低值{NEAR_OPTIMAL_REL_TOL*100:.1f}%的候选；")
    add("   (b) 再依次比较Eval_Time_RMSE、Time RMSE稳定性SD、Bootstrap RMSE、")
    add("       Time MAE、Expanding MAE、Expanding RMSE稳定性SD、特征数。")
    add("")

    add("三、输入与完整性核查")
    add(f"1. 16个F0-F3结果文件均通过结构核查：{len(input_check)}个。")
    add(f"2. 合并正式候选总行数：{len(formal_all)}。")
    add(f"3. 正式随机模型候选总数：{len(stability_pool)}；本版全部进行稳定性复核。")
    add(f"4. 随机稳定性汇总行数：{len(stability_summary)}。")
    add(f"5. 统一评价候选总数：{len(all_evaluated_candidates)}；应与正式候选总数完全一致。")
    add(f"6. 24个组内代表总数：{len(representatives_24)}（4目标×4F组×6算法=96行）；仅用于比较，不限制最终终选。")
    add("")

    add("四、四个目标最终机器学习模型")
    for target_name in TARGET_ORDER:
        row = final_ml.loc[final_ml["目标指标"] == target_name].iloc[0]
        conclusion = final_conclusions.loc[final_conclusions["目标指标"] == target_name].iloc[0]
        near = target_near.loc[target_near["目标指标"] == target_name].copy()

        add("")
        add(f"【{target_name}】")
        add(f"1. 最佳机器学习配置：{row['特征组']}-{row['模型']} | {row['候选参数ID']}")
        add(f"2. 参数：{row['参数JSON']}")
        add(f"3. 该目标全部正式候选数：{len(all_evaluated_candidates.loc[all_evaluated_candidates['目标指标'] == target_name])}")
        add(f"4. 目标级5%近优配置数：{len(near)}")
        add("5. 完整泛化指标：")
        add(f"   Bootstrap .632+：R²={fmt(row['Bootstrap_R2'])}，RMSE={fmt(row['Bootstrap_RMSE'])}")
        add(f"   Time：R²={fmt(row['Eval_Time_R2'])}，RMSE={fmt(row['Eval_Time_RMSE'])}")
        add(f"   Expanding：R²={fmt(row['Eval_Expanding_R2'])}，RMSE={fmt(row['Eval_Expanding_RMSE'])}")
        if str(row["模型"]) in STOCHASTIC_MODELS:
            add(
                "6. 随机稳定性："
                f"Time RMSE SD={fmt(row.get('Time_RMSE稳定性标准差'))}；"
                f"Expanding RMSE SD={fmt(row.get('Expanding_RMSE稳定性标准差'))}；"
                f"重复次数={int(safe_float(row.get('稳定性重复次数')))}；"
                f"警告重复={int(safe_float(row.get('稳定性警告重复次数')))}。"
            )
        else:
            add("6. 随机稳定性：该算法为确定性模型，相同数据与参数下结果确定，稳定性SD记为0。")
        add("7. 与历史基线：")
        add(
            f"   Time最佳基线={conclusion['Time最佳历史基线']}，"
            f"R²={fmt(conclusion['Time最佳基线R2'])}，RMSE={fmt(conclusion['Time最佳基线RMSE'])}。"
        )
        add(
            f"   Expanding最佳基线={conclusion['Expanding最佳历史基线']}，"
            f"R²={fmt(conclusion['Expanding最佳基线R2'])}，RMSE={fmt(conclusion['Expanding最佳基线RMSE'])}。"
        )
        add(
            f"   ML是否在两种时间验证中均优于最佳历史基线："
            f"{conclusion['ML是否在两种时间验证中均优于最佳历史基线']}。"
        )
        add(f"8. 最终预测结论：{conclusion['最终预测结论']}")
        add(f"9. 最终选择规则：{row['最终目标级选择规则']}")

        add("10. 该目标目标级5%近优候选（已按最终规则排序）：")
        for _, t in near.sort_values("目标级近优内部排序").iterrows():
            stable_text = (
                f", TimeSD={fmt(t['Time_RMSE稳定性标准差'])}, ExpSD={fmt(t['Expanding_RMSE稳定性标准差'])}"
                if str(t["模型"]) in STOCHASTIC_MODELS else ""
            )
            add(
                f"   #{int(t['目标级近优内部排序'])} {t['特征组']}-{t['模型']} | {t['候选参数ID']} | "
                f"Bootstrap R²={fmt(t['Bootstrap_R2'])}, RMSE={fmt(t['Bootstrap_RMSE'])}; "
                f"Time R²={fmt(t['Eval_Time_R2'])}, RMSE={fmt(t['Eval_Time_RMSE'])}; "
                f"Expanding R²={fmt(t['Eval_Expanding_R2'])}, RMSE={fmt(t['Eval_Expanding_RMSE'])}"
                f"{stable_text}"
            )

        top = target_ranked.loc[target_ranked["目标指标"] == target_name].sort_values("目标内统一排序").head(10)
        add("11. 该目标全部正式候选统一排序前10（便于核查）：")
        for _, t in top.iterrows():
            add(
                f"   #{int(t['目标内统一排序'])} {t['特征组']}-{t['模型']} | {t['候选参数ID']} | "
                f"Bootstrap R²={fmt(t['Bootstrap_R2'])}, RMSE={fmt(t['Bootstrap_RMSE'])}; "
                f"Time R²={fmt(t['Eval_Time_R2'])}, RMSE={fmt(t['Eval_Time_RMSE'])}; "
                f"Expanding R²={fmt(t['Eval_Expanding_R2'])}, RMSE={fmt(t['Eval_Expanding_RMSE'])}; "
                f"目标级5%近优={t['是否目标内5%近优']}"
            )

    add("")
    add("五、24个组内代表的用途说明")
    add("1. 每个目标仍保留4个F组×6算法=24个组内代表，用于F0-F3消融和六算法结果展示。")
    add("2. 组内代表也使用“组内Expanding 5%近优 → 统一Time优先”的规则。")
    add("3. 24个代表不再作为最终模型的硬性入口，因此不会再因组内替换而丢失目标级有竞争力的参数。")
    add("")

    add("六、随机模型稳定性复核摘要")
    if stability_summary.empty:
        add("- 未生成随机稳定性结果。")
    else:
        for target_name in TARGET_ORDER:
            ss = stability_summary.loc[stability_summary["目标指标"] == target_name].copy()
            warning_candidates = int((pd.to_numeric(ss["警告重复次数"], errors="coerce").fillna(0) > 0).sum())
            add(
                f"- {target_name}：稳定性复核候选{len(ss)}组；"
                f"存在至少1次警告的候选{warning_candidates}组；"
                f"所有候选均计划{STABILITY_REPEATS}个相同随机种子。"
            )
        add("  每个随机候选的5次明细、均值、标准差、警告信息均完整保存在Excel“随机稳定性明细/汇总”工作表。")

    add("")
    add("七、结果解释边界")
    add("1. Bootstrap .632+、Time、Expanding是三种不同验证证据，不应把Bootstrap R²直接当作未来时间泛化R²。")
    add("2. 随机模型的Time/Expanding最终比较使用同5随机种子均值；Bootstrap仍使用原1000次正式复核结果。")
    add("3. 最终模型选择对TP、TN、NH4+-N、COD完全使用同一套规则。")
    add("4. 如果某目标的最佳ML未在Time和Expanding中同时超过最佳历史基线，不应强行称其为性能最好的预测器。")
    add("5. 即使历史基线更强，最佳ML仍可用于SHAP/PDP和运行响应探索，但论文中必须明确预测性能边界。")
    add("6. F0-F3消融可使用24个统一组内代表；四个目标的最终预测模型结论则以“全部正式候选目标级直选”结果为准。")

    output_path.parent.mkdir(parents=True, exist_ok=True)
    try:
        output_path.write_text("\n".join(lines), encoding="utf-8-sig")
        return output_path
    except PermissionError:
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        alt = output_path.with_name(f"{output_path.stem}_{stamp}{output_path.suffix}")
        alt.write_text("\n".join(lines), encoding="utf-8-sig")
        return alt


# =============================================================================
# 9. 主程序
# =============================================================================

def main() -> None:
    start_time = datetime.now()
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    log("开始F0-F3最终模型统一终选与随机稳定性复核（目标级直选修正版）")
    log(f"结果根目录：{RESULT_ROOT}")
    log(f"动态数据：{DYNAMIC_INPUT_FILE}")
    log(f"输出目录：{OUTPUT_DIR}")

    # -------------------------------------------------------------------------
    # 步骤1：核查输入
    # -------------------------------------------------------------------------
    log("步骤1/8：核查16个结果文件和工作表结构")
    input_check = validate_environment()

    # -------------------------------------------------------------------------
    # 步骤2：读取所有正式候选
    # -------------------------------------------------------------------------
    log("步骤2/8：读取16个正式候选结果")
    all_formal_parts: List[pd.DataFrame] = []
    original_algorithm_near_parts: List[pd.DataFrame] = []

    for target_name in TARGET_ORDER:
        for fg in FEATURE_GROUP_ORDER:
            formal, _, source_file = load_formal_candidates_and_baseline(target_name, fg)
            formal = formal.copy()
            formal["来源文件"] = str(source_file)
            all_formal_parts.append(formal)

            # 只作为原正式单次结果的审计信息，不参与目标最终终选。
            marked = mark_near_optimal_within_algorithm(formal)
            original_algorithm_near_parts.append(
                marked.loc[marked["是否算法内5%近优"] == "是"].copy()
            )

    formal_all = pd.concat(all_formal_parts, ignore_index=True)
    original_algorithm_near = pd.concat(original_algorithm_near_parts, ignore_index=True)

    key_cols = ["目标指标", "特征组", "模型", "候选参数ID"]
    if formal_all.duplicated(subset=key_cols).any():
        preview = formal_all.loc[
            formal_all.duplicated(subset=key_cols, keep=False),
            key_cols,
        ].head(20)
        raise RuntimeError("合并正式候选存在重复键：\n" + preview.to_string(index=False))

    # -------------------------------------------------------------------------
    # 步骤3：所有正式随机候选统一稳定性复核
    # -------------------------------------------------------------------------
    stability_pool = formal_all.loc[formal_all["模型"].isin(STOCHASTIC_MODELS)].copy()
    log(
        f"步骤3/8：对所有正式随机候选补跑同{STABILITY_REPEATS}随机种子Time+Expanding；"
        f"候选总数={len(stability_pool)}"
    )

    stability_detail_parts: List[pd.DataFrame] = []
    stability_summary_parts: List[pd.DataFrame] = []
    target_tables: Dict[str, pd.DataFrame] = {}

    for target_name in TARGET_ORDER:
        if target_name not in target_tables:
            log(f"读取动态数据：{target_name}")
            target_tables[target_name] = read_target_table(target_name)
        target_df = target_tables[target_name]

        for fg in FEATURE_GROUP_ORDER:
            pool = stability_pool.loc[
                (stability_pool["目标指标"] == target_name)
                & (stability_pool["特征组"] == fg)
            ].copy()
            if pool.empty:
                raise RuntimeError(f"{target_name}-{fg} 没有正式随机模型候选。")

            model_counts = pool["模型"].value_counts().to_dict()
            missing_models = [m for m in STOCHASTIC_MODELS if m not in model_counts]
            if missing_models:
                raise RuntimeError(f"{target_name}-{fg} 缺少正式随机算法候选：{missing_models}")

            log(
                f"稳定性复核 {target_name}-{fg}："
                f"{len(pool)}组正式随机候选 × {STABILITY_REPEATS}种子；"
                f"RF={model_counts.get('RandomForest',0)}, "
                f"XGB={model_counts.get('XGBoost',0)}, "
                f"MLP={model_counts.get('MLP',0)}"
            )
            detail, summary = evaluate_stochastic_candidates(
                target_name,
                fg,
                pool,
                target_df,
            )
            stability_detail_parts.append(detail)
            stability_summary_parts.append(summary)

    stability_detail = pd.concat(stability_detail_parts, ignore_index=True)
    stability_summary = pd.concat(stability_summary_parts, ignore_index=True)

    # 任何复核失败都停止，绝不静默回退。
    failed = stability_detail.loc[stability_detail["是否成功"] != "是"].copy()
    if not failed.empty:
        error_preview = failed[
            ["目标指标", "特征组", "模型", "候选参数ID", "稳定性重复", "错误信息"]
        ].head(20)
        raise RuntimeError(
            "随机稳定性复核存在失败，已停止最终模型选定。请查看以下前20条：\n"
            + error_preview.to_string(index=False)
        )

    # -------------------------------------------------------------------------
    # 步骤4：建立所有正式候选的统一评价表
    # -------------------------------------------------------------------------
    log("步骤4/8：建立全部正式候选统一评价表")
    all_evaluated_candidates = build_all_evaluated_candidates(
        formal_all,
        stability_summary,
    )
    log(
        f"统一评价候选总数={len(all_evaluated_candidates)}；"
        f"与正式候选总数={len(formal_all)}一致"
    )

    # -------------------------------------------------------------------------
    # 步骤5：24组内代表（只供比较） + 目标级直接终选
    # -------------------------------------------------------------------------
    log("步骤5/8：生成24组内代表，并从全部正式候选直接选定四目标最终ML")
    group_ranked, representatives_24 = select_group_representatives(
        all_evaluated_candidates
    )
    target_ranked, target_near, final_ml = select_final_ml_per_target(
        all_evaluated_candidates
    )

    # 再做关键一致性检查：最终模型必须真的来自目标级5%池。
    for _, row in final_ml.iterrows():
        match = target_near.loc[
            (target_near["目标指标"] == row["目标指标"])
            & (target_near["特征组"] == row["特征组"])
            & (target_near["模型"] == row["模型"])
            & (target_near["候选参数ID"].astype(str) == str(row["候选参数ID"]))
        ]
        if len(match) != 1:
            raise RuntimeError(
                f"{row['目标指标']} 最终模型未唯一出现在目标级5%近优池中，停止输出。"
            )

    # -------------------------------------------------------------------------
    # 步骤6：基线比较
    # -------------------------------------------------------------------------
    log("步骤6/8：核对lag1/roll3基线并形成最终预测结论")
    baselines = collect_baselines()
    baseline_compare, final_conclusions = build_final_baseline_comparison(
        final_ml,
        baselines,
    )

    final_ml_out = final_ml.merge(
        final_conclusions[[
            "目标指标",
            "ML是否在两种时间验证中均优于最佳历史基线",
            "最终预测结论",
        ]],
        on="目标指标",
        how="left",
    )

    # -------------------------------------------------------------------------
    # 步骤7：输出Excel
    # -------------------------------------------------------------------------
    log("步骤7/8：输出完整Excel")
    settings = pd.DataFrame([
        {"项目": "运行时间", "内容": now_str()},
        {"项目": "结果根目录", "内容": str(RESULT_ROOT)},
        {"项目": "动态可用数据", "内容": str(DYNAMIC_INPUT_FILE)},
        {"项目": "输出目录", "内容": str(OUTPUT_DIR)},
        {"项目": "目标级Expanding近优阈值", "内容": f"最低值+{NEAR_OPTIMAL_REL_TOL*100:.1f}%"},
        {"项目": "四目标终选规则", "内容": "TP/TN/NH4+-N/COD完全一致，不设置目标特异阈值或权重"},
        {"项目": "最终候选入口", "内容": "每目标全部正式复核候选直接进入统一评价；24组内代表不限制最终候选"},
        {"项目": "随机模型稳定性范围", "内容": "所有正式RandomForest/XGBoost/MLP候选"},
        {"项目": "随机模型稳定性重复", "内容": STABILITY_REPEATS},
        {"项目": "统一随机种子", "内容": ",".join(str(x) for x in STABILITY_SEEDS)},
        {"项目": "Bootstrap处理", "内容": "不重跑；读取原1000次正式Bootstrap .632+ R2/RMSE"},
        {"项目": "确定性模型Time/Expanding", "内容": "使用原正式验证结果"},
        {"项目": "随机模型Time/Expanding", "内容": "使用同5随机种子复核均值，并记录标准差"},
        {"项目": "最终统一排序", "内容": "目标Expanding最低值5%近优 → Time RMSE → Time SD → Bootstrap RMSE → Time MAE → Expanding MAE → Expanding SD → 特征数"},
        {"项目": "完整指标输出", "内容": "Bootstrap .632+ R2/RMSE；Time R2/RMSE；Expanding R2/RMSE"},
    ])

    excel_tables = {
        "运行设置": settings,
        "输入文件核查": input_check,
        "全部正式候选原始": formal_all,
        "原正式算法内5%近优": original_algorithm_near,
        "随机稳定性候选": stability_pool,
        "随机稳定性明细": stability_detail,
        "随机稳定性汇总": stability_summary,
        "全部统一评价候选": all_evaluated_candidates,
        "组内候选统一排序": group_ranked,
        "24个组内代表": representatives_24,
        "目标级5%终选候选": target_near,
        "目标级全部候选排序": target_ranked,
        "四指标最终ML": final_ml_out,
        "朴素动态基线": baselines,
        "最终ML基线对比": baseline_compare,
        "最终预测结论": final_conclusions,
    }
    excel_path = safe_write_excel(OUTPUT_EXCEL, excel_tables)

    # -------------------------------------------------------------------------
    # 步骤8：输出TXT
    # -------------------------------------------------------------------------
    log("步骤8/8：输出完整TXT模型选定说明")
    text_path = write_complete_text_report(
        OUTPUT_TEXT,
        start_time,
        input_check,
        formal_all,
        stability_pool,
        stability_summary,
        all_evaluated_candidates,
        representatives_24,
        target_ranked,
        target_near,
        final_ml_out,
        baselines,
        baseline_compare,
        final_conclusions,
    )

    log("=" * 96)
    log("最终模型统一终选完成")
    log(f"Excel：{excel_path}")
    log(f"TXT：{text_path}")
    log("四个目标最终结果（四目标使用完全相同选择规则）：")
    for _, row in final_conclusions.iterrows():
        log(
            f"{row['目标指标']}：{row['最佳机器学习特征组']}-{row['最佳机器学习模型']} "
            f"({row['最佳机器学习候选参数ID']}) | "
            f"Bootstrap R2={fmt(row['Bootstrap_R2'])}, RMSE={fmt(row['Bootstrap_RMSE'])} | "
            f"Time R2={fmt(row['Time_R2'])}, RMSE={fmt(row['Time_RMSE'])} | "
            f"Expanding R2={fmt(row['Expanding_R2'])}, RMSE={fmt(row['Expanding_RMSE'])} | "
            f"稳定优于最佳历史基线={row['ML是否在两种时间验证中均优于最佳历史基线']}"
        )


if __name__ == "__main__":
    main()