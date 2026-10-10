# -*- coding: utf-8 -*-
r"""
A²/O 四指标 DO×HRT×进水分配比例联合正向预测（固定粗网格十一文件版）
=======================================================================

定位
----
本脚本只负责“正向预测”：建立完整运行情景库并记录四项一级A工程参考判定，
不在本阶段删除情景，也不直接推荐运行方案。

核心问题：
    在一个真实历史状态下，固定当前进水、近期出水状态和反应器结构，
    如果改变 Aerobic_DO、HRT 和 Anaerobic_Feed_Fraction，
    下一步 TP、TN、NH4+-N 和 COD 的冻结模型预测值分别是多少？

本版正式逻辑
------------
1. 使用已经冻结的四个最终 F2 模型，不重新训练、不重新调参：
   - TP       : F2 + SVR        / SVR_032
   - TN       : F2 + SVR        / SVR_018
   - NH4+-N   : F2 + ElasticNet / ElasticNet_025
   - COD      : F2 + ElasticNet / ElasticNet_040

2. 使用四个目标共同有效的全部真实历史状态。按当前确认数据应为 631 个。

3. 每个历史状态固定：
   - 四项进水水质；
   - TP/TN/NH4/COD 各自 lag1、lag2、roll3；
   - Anaerobic_Volume_Fraction；
   - Anoxic_Volume_Fraction。

4. 每个历史状态只改变三个运行调节量：
   - Aerobic_DO；
   - HRT；
   - Anaerobic_Feed_Fraction。

5. 正式参数库只采用统一的第一轮全范围粗扫描，不做局部细化：
   - DO 默认步长 0.10 mg/L；
   - HRT 默认步长 0.50 h；
   - Anaerobic_Feed_Fraction 默认步长 0.05；
   - inclusive_grid() 强制包含搜索上下界，因此即使最后一个常规步长点不能恰好落在上界，也会补入上界。
   在当前确认数据范围下，DO=0.35~3.00 共 28 个水平，HRT=8~12 共 9 个水平，
   Feed=0.30~1.00 共 15 个水平，每个历史状态形成 28×9×15=3780 组统一粗网格。

6. 631 个共同历史状态对应的正式粗网格记录数为：
       631 × 3780 = 2,385,180 条。
   为保留当前真实工况对照，如果某一历史状态的当前 DO/HRT/Feed 不在统一粗网格上，
   仍额外补入该当前工况，并标记“网格来源=当前实际工况补充”。按当前数据实际额外补充 180 条时，
   完整正向明细为 2,385,360 条。后续反向筛选只使用“第一轮粗扫描”记录，
   不把这些不规则补充记录纳入统一参数组合统计。

7. 正向阶段不删除任何情景，也不输出运行参数推荐；但为与后续反向筛选完全衔接，
   每条记录直接附加四项一级A工程参考判定：
   - TP <= 0.5 mg/L；
   - TN <= 15 mg/L；
   - NH4+-N <= 5 mg/L；
   - COD <= 50 mg/L；
   并记录“四项一级A参考同时满足”。这些阈值只作为一步条件预测的工程参考筛选标签，
   不等同于法规意义上的日均达标结论。后续反向程序仍应仅使用“第一轮粗扫描”记录，
   排除负预测/非有限预测后，对四项条件执行 AND 联合筛选。

8. 历史5NN距离、历史接近程度等旧版辅助量不再计算，也不写入数百万行明细。
   本版只保留历史状态ID、历史真实完整工况标记等必要追踪信息，避免历史接近程度影响后续固定阈值筛选。

9. 模型预测准确性参考仍直接读取冻结模型包中保存的样本外验证结果：
   Bootstrap .632+、Time、Expanding 的 R²、RMSE、MAE。
   需要单值误差参考时，使用 Time RMSE 与 Expanding RMSE 中较大的有效值，
   仅作为结果解释和审计信息，不参与正向筛选。

10. HRT 改变后同步计算：
        Q(L/h) = 57.6 / HRT
        Q(L/d) = 57.6 * 24 / HRT
    Q 只用于运行换算与结果报告，不进入任何模型输入。

11. 输出目录保持不变。为降低单个Excel体积，完整正向预测结果按历史状态ID连续拆成十一个结果Excel，
    不再输出旧版单一主Excel，也不输出“结果检查包”：
    1) 四指标_DO-HRT-进水分配联合正向预测_第1部分.xlsx
    2) 四指标_DO-HRT-进水分配联合正向预测_第2部分.xlsx
    3) 四指标_DO-HRT-进水分配联合正向预测_第3部分.xlsx
    4) 四指标_DO-HRT-进水分配联合正向预测_第4部分.xlsx
    5) 四指标_DO-HRT-进水分配联合正向预测_第5部分.xlsx
    6) 四指标_DO-HRT-进水分配联合正向预测_第6部分.xlsx
    7) 四指标_DO-HRT-进水分配联合正向预测_第7部分.xlsx
    8) 四指标_DO-HRT-进水分配联合正向预测_第8部分.xlsx
    9) 四指标_DO-HRT-进水分配联合正向预测_第9部分.xlsx
    10) 四指标_DO-HRT-进水分配联合正向预测_第10部分.xlsx
    11) 四指标_DO-HRT-进水分配联合正向预测_第11部分.xlsx
    12) 四指标_DO-HRT-进水分配联合正向预测_说明.txt

    十一个Excel按历史状态数量尽量均匀、按ID连续拆分。默认631个背景时自动分为58+58+58+58+57+57+57+57+57+57+57：
    - 第1部分：B0001~B0058；
    - 第2部分：B0059~B0116；
    - 第3部分：B0117~B0174；
    - 第4部分：B0175~B0232；
    - 第5部分：B0233~B0289；
    - 第6部分：B0290~B0346；
    - 第7部分：B0347~B0403；
    - 第8部分：B0404~B0460；
    - 第9部分：B0461~B0517；
    - 第10部分：B0518~B0574；
    - 第11部分：B0575~B0631。
    每个历史背景的全部3780条统一粗网格记录及其当前工况补充记录只会写入其中一个文件，绝不跨文件拆分。
    为避免重复，公共说明/模型核查/全部历史状态/全部背景汇总只写入第1部分；第2~11部分只保存各自未重复的明细副表。
    每个结果文件的明细默认按该部分预计总行数均分为约3个Sheet；若显式设置A2O_FORWARD_DETAIL_ROWS_PER_SHEET为更小值，则以该上限为准。

解释边界
--------
- 本脚本输出的是“固定当前进水和历史状态，只改变当前运行设定后的单步条件预测”。
- 不是长期稳态模拟，不代表工况改变后系统最终一定达到该浓度。
- 正向预测本身不输出“最优工况”。后续反向阶段只对统一粗网格按四项一级A工程参考条件进行联合筛选。
- 对 NH4+-N，如果最终模型对运行变量没有有效响应，不能据此声称实际 NH4 不受运行参数影响。

默认输入
--------
动态数据：
${PROJECT_ROOT}/outputs/preprocessing/dynamic/动态可用数据与插补比较结果.xlsx

最终模型目录：
${PROJECT_ROOT}/outputs/final_selection/moxingbaocun

默认输出目录：
${PROJECT_ROOT}/outputs/screening/forward

依赖
----
pip install pandas numpy openpyxl xlsxwriter scikit-learn joblib

可选环境变量
------------
A2O_FORWARD_INPUT_FILE
A2O_FORWARD_MODEL_DIR
A2O_FORWARD_OUTPUT_DIR
A2O_FORWARD_STRICT=1

A2O_FORWARD_DO_MIN / A2O_FORWARD_DO_MAX
A2O_FORWARD_HRT_MIN / A2O_FORWARD_HRT_MAX
A2O_FORWARD_FEED_MIN / A2O_FORWARD_FEED_MAX

A2O_FORWARD_DO_STEP=0.10
A2O_FORWARD_HRT_STEP=0.50
A2O_FORWARD_FEED_STEP=0.05

A2O_FORWARD_DETAIL_ROWS_PER_SHEET=900000  # 明细单Sheet硬上限；默认另按每个结果文件约3个明细Sheet自动拆分
"""

from __future__ import annotations

import hashlib
import json
import math
import os
import re
import time
import traceback
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, Iterable, List, Mapping, Optional, Sequence, Tuple

import joblib
import numpy as np
import pandas as pd
import xlsxwriter
from sklearn.metrics import mean_absolute_error, mean_squared_error, r2_score


# =============================================================================
# 1. 路径、常量与冻结配置
# =============================================================================

PROJECT_ROOT = Path(__file__).resolve().parents[1]

INPUT_FILE = Path(
    os.environ.get(
        "A2O_FORWARD_INPUT_FILE",
        str(PROJECT_ROOT / "outputs/preprocessing/dynamic/动态可用数据与插补比较结果.xlsx"),
    )
)

MODEL_DIR = Path(
    os.environ.get(
        "A2O_FORWARD_MODEL_DIR",
        str(PROJECT_ROOT / "outputs/final_selection/moxingbaocun"),
    )
)

OUTPUT_DIR = Path(
    os.environ.get(
        "A2O_FORWARD_OUTPUT_DIR",
        str(PROJECT_ROOT / "outputs/screening/forward"),
    )
)

OUTPUT_EXCEL_PART1 = OUTPUT_DIR / "四指标_DO-HRT-进水分配联合正向预测_第1部分.xlsx"
OUTPUT_EXCEL_PART2 = OUTPUT_DIR / "四指标_DO-HRT-进水分配联合正向预测_第2部分.xlsx"
OUTPUT_EXCEL_PART3 = OUTPUT_DIR / "四指标_DO-HRT-进水分配联合正向预测_第3部分.xlsx"
OUTPUT_EXCEL_PART4 = OUTPUT_DIR / "四指标_DO-HRT-进水分配联合正向预测_第4部分.xlsx"
OUTPUT_EXCEL_PART5 = OUTPUT_DIR / "四指标_DO-HRT-进水分配联合正向预测_第5部分.xlsx"
OUTPUT_EXCEL_PART6 = OUTPUT_DIR / "四指标_DO-HRT-进水分配联合正向预测_第6部分.xlsx"
OUTPUT_EXCEL_PART7 = OUTPUT_DIR / "四指标_DO-HRT-进水分配联合正向预测_第7部分.xlsx"
OUTPUT_EXCEL_PART8 = OUTPUT_DIR / "四指标_DO-HRT-进水分配联合正向预测_第8部分.xlsx"
OUTPUT_EXCEL_PART9 = OUTPUT_DIR / "四指标_DO-HRT-进水分配联合正向预测_第9部分.xlsx"
OUTPUT_EXCEL_PART10 = OUTPUT_DIR / "四指标_DO-HRT-进水分配联合正向预测_第10部分.xlsx"
OUTPUT_EXCEL_PART11 = OUTPUT_DIR / "四指标_DO-HRT-进水分配联合正向预测_第11部分.xlsx"
OUTPUT_EXCEL_PARTS = [
    OUTPUT_EXCEL_PART1,
    OUTPUT_EXCEL_PART2,
    OUTPUT_EXCEL_PART3,
    OUTPUT_EXCEL_PART4,
    OUTPUT_EXCEL_PART5,
    OUTPUT_EXCEL_PART6,
    OUTPUT_EXCEL_PART7,
    OUTPUT_EXCEL_PART8,
    OUTPUT_EXCEL_PART9,
    OUTPUT_EXCEL_PART10,
    OUTPUT_EXCEL_PART11,
]
OUTPUT_TEXT = OUTPUT_DIR / "四指标_DO-HRT-进水分配联合正向预测_说明.txt"

# 旧版输出名仅用于启动时清理，当前版本不再生成这些文件。
LEGACY_OUTPUT_EXCEL = OUTPUT_DIR / "四指标_DO-HRT-进水分配联合正向预测.xlsx"
LEGACY_OUTPUT_CHECK_EXCEL = OUTPUT_DIR / "四指标_DO-HRT-进水分配联合正向预测_结果检查包.xlsx"

SCRIPT_VERSION = "A2O_FORWARD_DOHRTFEED_v3.12_A1_eleven_workbooks_unique_public_sheets_split_detail"
BUILD_FINGERPRINT = "FORWARD_ELEVEN_WORKBOOKS_PUBLIC_SHEETS_ONLY_ONCE_DETAIL_SPLIT_3_20260922"
STRICT_MODE = os.environ.get("A2O_FORWARD_STRICT", "1").strip() != "0"

RANDOM_SEED = 20260620
MAX_INTERPOLATION_GAP = 7
ROUNDTRIP_METRIC_TOL = 1e-9
MEDIAN_TOL = 1e-10
RAW_CONSISTENCY_TOL = 1e-10
NH4_ZERO_COEF_TOL = 1e-12

# 反应器固定有效容积
REACTOR_VOLUME_L = 57.6

# 第一轮粗扫描步长
DO_STEP = float(os.environ.get("A2O_FORWARD_DO_STEP", "0.10"))
HRT_STEP = float(os.environ.get("A2O_FORWARD_HRT_STEP", "0.50"))
FEED_STEP = float(os.environ.get("A2O_FORWARD_FEED_STEP", "0.05"))

# 正向阶段只生成统一粗网格；后续反向阶段直接筛选该粗网格，不再做局部细化。
FORMAL_INVERSE_GRID_SOURCE = "第一轮粗扫描"

# 水质浓度预测的物理检查容差。低于 -1e-12 视为负预测；原始预测值仍保留。
PREDICTION_NEGATIVE_TOL = 1e-12

# Excel 单工作簿内的明细副表行数硬上限。
# 默认仍保留900000作为安全上限；实际主程序会按每个结果文件预计总行数自动均分为约3个明细Sheet，
# 从而避免单个明细Sheet过大。若用户显式把A2O_FORWARD_DETAIL_ROWS_PER_SHEET设得更小，则优先采用更小上限。
DETAIL_ROWS_PER_SHEET = int(
    os.environ.get("A2O_FORWARD_DETAIL_ROWS_PER_SHEET", "900000")
)
DETAIL_TARGET_SHEETS_PER_WORKBOOK = 3
EXCEL_HARD_MAX_ROWS = 1_048_576
if DETAIL_ROWS_PER_SHEET <= 0 or DETAIL_ROWS_PER_SHEET >= EXCEL_HARD_MAX_ROWS - 1:
    raise ValueError(
        "A2O_FORWARD_DETAIL_ROWS_PER_SHEET 必须大于0且小于 Excel 单表最大数据行数。"
    )

if min(DO_STEP, HRT_STEP, FEED_STEP) <= 0:
    raise ValueError("DO、HRT、进水分配比例的粗扫描步长必须为正数。")

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

ALL_OPERATION_FEATURES = [
    "Aerobic_DO",
    "Anaerobic_Feed_Fraction",
    "HRT",
    "Anaerobic_Volume_Fraction",
    "Anoxic_Volume_Fraction",
]

CONTROL_FEATURES = ["Aerobic_DO", "HRT", "Anaerobic_Feed_Fraction"]
STRUCTURE_FEATURES = ["Anaerobic_Volume_Fraction", "Anoxic_Volume_Fraction"]

PREDICTION_COLUMN = {
    "TP": "TP_预测",
    "TN": "TN_预测",
    "NH4+-N": "NH4_预测",
    "COD": "COD_预测",
}

# 四项一级A工程参考限值。
# 仅用于正向结果附加判定字段和后续反向筛选衔接；
# 本脚本输出为一步条件预测，不将这些标签解释为法规意义上的日均达标结论。
A1_REFERENCE_LIMITS: Dict[str, float] = {
    "TP": 0.5,
    "TN": 15.0,
    "NH4+-N": 5.0,
    "COD": 50.0,
}

A1_REFERENCE_FLAG_COLUMN: Dict[str, str] = {
    "TP": "TP_一级A参考满足",
    "TN": "TN_一级A参考满足",
    "NH4+-N": "NH4_一级A参考满足",
    "COD": "COD_一级A参考满足",
}
A1_ALL_FLAG_COLUMN = "四项一级A参考同时满足"
A1_SATISFY_TOL = 1e-12

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
        result = float(value)
        return result if np.isfinite(result) else np.nan
    except Exception:
        return np.nan


def params_equal(left: Mapping[str, Any], right: Mapping[str, Any]) -> bool:
    if set(left.keys()) != set(right.keys()):
        return False
    for key in left:
        lv, rv = left[key], right[key]
        if isinstance(lv, (int, float, np.integer, np.floating)) and isinstance(
            rv, (int, float, np.integer, np.floating)
        ):
            if not math.isclose(float(lv), float(rv), rel_tol=1e-12, abs_tol=1e-12):
                return False
        elif str(lv) != str(rv):
            return False
    return True


def calc_metrics(y_true: Sequence[float], y_pred: Sequence[float]) -> Dict[str, float]:
    y_true_arr = np.asarray(y_true, dtype=float)
    y_pred_arr = np.asarray(y_pred, dtype=float)
    return {
        "R2": float(r2_score(y_true_arr, y_pred_arr)),
        "RMSE": float(math.sqrt(mean_squared_error(y_true_arr, y_pred_arr))),
        "MAE": float(mean_absolute_error(y_true_arr, y_pred_arr)),
    }


def sanitize_excel_value(value: Any) -> Any:
    if isinstance(value, str):
        return re.sub(r"[\x00-\x08\x0B\x0C\x0E-\x1F]", "", value)
    return value


def sanitize_frame(frame: pd.DataFrame) -> pd.DataFrame:
    output = frame.copy()
    for column in output.select_dtypes(include=["object", "string"]).columns:
        output[column] = output[column].map(sanitize_excel_value)
    return output


def atomic_write_text(path: Path, text: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(path.suffix + ".part")
    tmp.write_text(text, encoding="utf-8-sig")
    os.replace(tmp, path)




def cleanup_previous_forward_outputs(output_dir: Path) -> None:
    """删除本脚本及旧版正向脚本的同名前缀输出，避免新旧结果混在一起。"""
    patterns = [
        OUTPUT_EXCEL_PART1.name,
        OUTPUT_EXCEL_PART1.stem + ".part.xlsx",
        OUTPUT_EXCEL_PART2.name,
        OUTPUT_EXCEL_PART2.stem + ".part.xlsx",
        OUTPUT_EXCEL_PART3.name,
        OUTPUT_EXCEL_PART3.stem + ".part.xlsx",
        OUTPUT_EXCEL_PART4.name,
        OUTPUT_EXCEL_PART4.stem + ".part.xlsx",
        OUTPUT_EXCEL_PART5.name,
        OUTPUT_EXCEL_PART5.stem + ".part.xlsx",
        OUTPUT_EXCEL_PART6.name,
        OUTPUT_EXCEL_PART6.stem + ".part.xlsx",
        OUTPUT_EXCEL_PART7.name,
        OUTPUT_EXCEL_PART7.stem + ".part.xlsx",
        OUTPUT_EXCEL_PART8.name,
        OUTPUT_EXCEL_PART8.stem + ".part.xlsx",
        OUTPUT_EXCEL_PART9.name,
        OUTPUT_EXCEL_PART9.stem + ".part.xlsx",
        OUTPUT_EXCEL_PART10.name,
        OUTPUT_EXCEL_PART10.stem + ".part.xlsx",
        OUTPUT_EXCEL_PART11.name,
        OUTPUT_EXCEL_PART11.stem + ".part.xlsx",
        LEGACY_OUTPUT_EXCEL.name,
        LEGACY_OUTPUT_EXCEL.stem + ".part.xlsx",
        LEGACY_OUTPUT_CHECK_EXCEL.name,
        LEGACY_OUTPUT_CHECK_EXCEL.stem + ".part.xlsx",
        "四指标_DO-HRT-进水分配联合正向预测_汇总.xlsx",
        "四指标_DO-HRT-进水分配联合正向预测_汇总.part.xlsx",
        "四指标_DO-HRT-进水分配联合正向预测_明细_*.xlsx",
        "四指标_DO-HRT-进水分配联合正向预测_明细_*.part.xlsx",
        OUTPUT_TEXT.name,
        OUTPUT_TEXT.name + ".part",
    ]
    for pattern in patterns:
        for path in output_dir.glob(pattern):
            try:
                path.unlink()
            except FileNotFoundError:
                pass


def numeric_series_equal(left: pd.Series, right: pd.Series, tol: float) -> bool:
    a = pd.to_numeric(left, errors="coerce").to_numpy(dtype=float)
    b = pd.to_numeric(right, errors="coerce").to_numpy(dtype=float)
    both_nan = np.isnan(a) & np.isnan(b)
    both_finite = np.isfinite(a) & np.isfinite(b)
    okay = both_nan | (both_finite & (np.abs(a - b) <= tol))
    return bool(np.all(okay))


def rounded_key(values: Sequence[float], digits: int = 8) -> Tuple[float, ...]:
    return tuple(round(float(v), digits) for v in values)


def full_operation_key(values: Mapping[str, Any]) -> Tuple[float, ...]:
    return rounded_key([float(values[c]) for c in ALL_OPERATION_FEATURES])


def get_optional_float_env(name: str) -> Optional[float]:
    raw = os.environ.get(name)
    if raw is None or str(raw).strip() == "":
        return None
    value = float(raw)
    if not np.isfinite(value):
        raise ValueError(f"环境变量 {name} 必须是有限数值。")
    return value


def inclusive_grid(low: float, high: float, step: float, digits: int = 10) -> np.ndarray:
    """生成包含上下界的稳定网格，避免浮点累积误差。"""
    low = float(low)
    high = float(high)
    step = float(step)
    if not (np.isfinite(low) and np.isfinite(high) and np.isfinite(step)):
        raise ValueError("网格边界/步长必须为有限数。")
    if step <= 0 or high < low:
        raise ValueError(f"网格设置错误：low={low}, high={high}, step={step}")
    if math.isclose(low, high, rel_tol=0, abs_tol=1e-15):
        return np.asarray([round(low, digits)], dtype=float)

    n = int(math.floor((high - low) / step + 1e-12))
    values = [low + i * step for i in range(n + 1)]
    if values[-1] < high - 1e-10:
        values.append(high)
    else:
        values[-1] = high if abs(values[-1] - high) <= 1e-10 else values[-1]
    values = sorted(set(round(float(v), digits) for v in values if low - 1e-12 <= v <= high + 1e-12))
    if not math.isclose(values[0], low, abs_tol=1e-10):
        values.insert(0, round(low, digits))
    if not math.isclose(values[-1], high, abs_tol=1e-10):
        values.append(round(high, digits))
    return np.asarray(values, dtype=float)


# =============================================================================
# 3. 读取冻结模型、动态表并严格重建训练输入
# =============================================================================

def interpolate_short_internal_gaps(
    series: pd.Series,
    sample_ids: pd.Series,
    max_gap: int = MAX_INTERPOLATION_GAP,
) -> pd.Series:
    """与最终模型固化脚本保持一致：短内部缺口线性插值，其余中位数兜底。"""
    output = convert_numeric(series).astype(float).copy()
    values = output.to_numpy(dtype=float)
    coordinates = convert_numeric(sample_ids).to_numpy(dtype=float)

    index = 0
    while index < len(values):
        if np.isfinite(values[index]):
            index += 1
            continue
        start = index
        while index < len(values) and not np.isfinite(values[index]):
            index += 1
        end = index - 1
        left = start - 1
        right = index

        if (
            left >= 0
            and right < len(values)
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
                for i in range(start, end + 1):
                    fraction = (coordinates[i] - coordinates[left]) / coordinate_span
                    values[i] = left_value + fraction * (right_value - left_value)

    return pd.Series(values, index=output.index, name=output.name)


def load_model_package(target_name: str) -> Tuple[Path, Dict[str, Any]]:
    path = MODEL_DIR / TARGET_CONFIG[target_name]["joblib"]
    if not path.exists():
        raise FileNotFoundError(f"未找到 {target_name} 最终模型：{path}")

    package = joblib.load(path)
    if not isinstance(package, dict):
        raise TypeError(f"{target_name} joblib 不是字典型最终模型包。")

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
    expected_features = RAW_FEATURES + TARGET_CONFIG[target_name]["history"]
    checks = [
        str(package["target_name"]) == target_name,
        str(package["selected_feature_group"]) == expected["feature_group"],
        str(package["algorithm"]) == expected["algorithm"],
        str(package["candidate_id"]) == expected["candidate_id"],
        params_equal(package["parameters"], expected["parameters"]),
        list(package["feature_names"]) == expected_features,
        len(package["feature_names"]) == 12,
        "Q" not in package["feature_names"],
    ]
    if STRICT_MODE and not all(checks):
        raise RuntimeError(
            f"{target_name} 模型包与锁定配置不一致："
            f"F={package['selected_feature_group']}，算法={package['algorithm']}，"
            f"ID={package['candidate_id']}，参数={package['parameters']}，"
            f"特征={package['feature_names']}"
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
    data.columns = [clean_column_name(c) for c in data.columns]

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
    if data[cfg["target_col"]].isna().any():
        raise ValueError(f"{target_name} 目标变量仍存在缺失。")
    if data[cfg["history"]].isna().any().any():
        raise ValueError(f"{target_name} lag1/lag2/roll3 仍存在缺失。")
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
        if col not in saved_medians:
            raise KeyError(f"{target_name} 模型包 preprocessing.training_medians 缺少 {col}。")
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
                "当前原始中位数": current_median,
                "中位数绝对差": median_diff,
            }
        )

    feature_names = list(package["feature_names"])
    if data[feature_names].isna().any().any():
        bad = data[feature_names].columns[data[feature_names].isna().any()].tolist()
        raise RuntimeError(f"{target_name} 重建后仍有缺失输入：{bad}")

    expected_ids = [int(x) for x in package["training_data"].get("sample_ids", [])]
    current_ids = data["样本序号"].astype(int).tolist()
    if STRICT_MODE and expected_ids and expected_ids != current_ids:
        raise RuntimeError(f"{target_name} 动态表样本ID与模型包训练样本ID不一致。")

    saved_dynamic_hash = str(package["training_data"].get("dynamic_input_file_sha256", ""))
    current_dynamic_hash = sha256_file(INPUT_FILE)
    if STRICT_MODE and saved_dynamic_hash and saved_dynamic_hash != current_dynamic_hash:
        raise RuntimeError(f"{target_name} 模型包记录的动态数据哈希与当前输入文件不一致。")

    return data, pd.DataFrame(audit_rows)


def validate_full_fit(
    target_name: str,
    data: pd.DataFrame,
    package: Mapping[str, Any],
) -> Tuple[np.ndarray, Dict[str, float], Dict[str, float]]:
    feature_names = list(package["feature_names"])
    x = data[feature_names].to_numpy(dtype=float)
    y = data[TARGET_CONFIG[target_name]["target_col"]].to_numpy(dtype=float)
    pred = np.asarray(package["model"].predict(package["scaler"].transform(x)), dtype=float)
    metrics = calc_metrics(y, pred)
    saved = package.get("full_fit_metrics", {})
    diffs = {
        key: abs(float(metrics[key]) - float(saved[key]))
        for key in ["R2", "RMSE", "MAE"]
        if key in saved and np.isfinite(metrics[key]) and np.isfinite(safe_float(saved[key]))
    }
    if STRICT_MODE and diffs and max(diffs.values()) > ROUNDTRIP_METRIC_TOL:
        raise RuntimeError(f"{target_name} 全量拟合指标复现失败：{diffs}")
    return pred, metrics, diffs



# =============================================================================
# 4. 631个共同历史状态与历史真实运行组合
# =============================================================================

def build_common_history(raw_by_target: Mapping[str, pd.DataFrame]) -> Tuple[pd.DataFrame, pd.DataFrame]:
    common_ids: Optional[set[int]] = None
    for target in TARGET_ORDER:
        ids = set(raw_by_target[target]["样本序号"].astype(int).tolist())
        common_ids = ids if common_ids is None else common_ids & ids
    if not common_ids:
        raise RuntimeError("四个动态表没有共同有效样本。")

    ordered_ids = sorted(common_ids)
    indexed = {
        target: raw_by_target[target].set_index("样本序号", drop=False).loc[ordered_ids].copy()
        for target in TARGET_ORDER
    }
    base = indexed["TP"]
    audit_rows: List[Dict[str, Any]] = []

    for target in TARGET_ORDER[1:]:
        other = indexed[target]
        date_match = bool(
            np.all(pd.to_datetime(base["Date"]).to_numpy() == pd.to_datetime(other["Date"]).to_numpy())
        )
        audit_rows.append(
            {
                "核查项目": f"TP与{target}共同样本日期一致",
                "核查结果": "通过" if date_match else "失败",
                "共同样本数": len(ordered_ids),
            }
        )
        if STRICT_MODE and not date_match:
            raise RuntimeError(f"TP 与 {target} 的共同样本日期不一致。")

        for col in RAW_FEATURES:
            match = numeric_series_equal(base[col], other[col], RAW_CONSISTENCY_TOL)
            audit_rows.append(
                {
                    "核查项目": f"TP与{target}共同样本 {col} 一致",
                    "核查结果": "通过" if match else "失败",
                    "共同样本数": len(ordered_ids),
                }
            )
            if STRICT_MODE and not match:
                raise RuntimeError(f"TP 与 {target} 的共同样本 {col} 不一致。")

    common = base[["样本序号", "Date"] + RAW_FEATURES].copy().reset_index(drop=True)
    for target in TARGET_ORDER:
        cfg = TARGET_CONFIG[target]
        frame = indexed[target].reset_index(drop=True)
        common[cfg["target_col"]] = frame[cfg["target_col"]].to_numpy()
        for col in cfg["history"]:
            common[col] = frame[col].to_numpy()

    common = common.sort_values(["样本序号", "Date"], kind="mergesort").reset_index(drop=True)
    common.insert(0, "历史状态ID", [f"B{i:04d}" for i in range(1, len(common) + 1)])

    audit_rows.append(
        {
            "核查项目": "四指标共同有效历史状态数量",
            "核查结果": "通过",
            "共同样本数": len(common),
        }
    )
    return common, pd.DataFrame(audit_rows)


def build_historical_operation_table(common_history: pd.DataFrame) -> Tuple[pd.DataFrame, Dict[Tuple[float, ...], str]]:
    source = common_history.dropna(subset=ALL_OPERATION_FEATURES).copy()
    if source.empty:
        raise RuntimeError("共同历史状态中没有完整运行参数。")

    rows: List[Dict[str, Any]] = []
    grouped: Dict[Tuple[float, ...], List[int]] = {}
    for idx, row in source.iterrows():
        key = full_operation_key(row)
        grouped.setdefault(key, []).append(idx)

    for i, (key, idxs) in enumerate(sorted(grouped.items()), start=1):
        group = source.loc[idxs]
        values = dict(zip(ALL_OPERATION_FEATURES, key))
        rows.append(
            {
                "历史运行组合ID": f"OP{i:02d}",
                **values,
                "历史出现次数": len(group),
                "首次出现日期": pd.to_datetime(group["Date"]).min(),
                "最近出现日期": pd.to_datetime(group["Date"]).max(),
            }
        )

    table = pd.DataFrame(rows)
    lookup = {
        rounded_key([row[c] for c in ALL_OPERATION_FEATURES]): str(row["历史运行组合ID"])
        for _, row in table.iterrows()
    }
    return table, lookup


# =============================================================================
# 5. 运行参数范围、粗网格与当前工况补充
# =============================================================================

def resolve_search_bounds(common_history: pd.DataFrame) -> Dict[str, Tuple[float, float]]:
    mapping = {
        "Aerobic_DO": ("A2O_FORWARD_DO_MIN", "A2O_FORWARD_DO_MAX"),
        "HRT": ("A2O_FORWARD_HRT_MIN", "A2O_FORWARD_HRT_MAX"),
        "Anaerobic_Feed_Fraction": ("A2O_FORWARD_FEED_MIN", "A2O_FORWARD_FEED_MAX"),
    }
    bounds: Dict[str, Tuple[float, float]] = {}
    for col, (min_env, max_env) in mapping.items():
        history = pd.to_numeric(common_history[col], errors="coerce")
        hist_min = float(history.min())
        hist_max = float(history.max())
        low = get_optional_float_env(min_env)
        high = get_optional_float_env(max_env)
        low = hist_min if low is None else low
        high = hist_max if high is None else high
        if high < low:
            raise ValueError(f"{col} 搜索上限小于下限：{low}~{high}")
        bounds[col] = (float(low), float(high))
    return bounds


def build_coarse_grid(bounds: Mapping[str, Tuple[float, float]]) -> pd.DataFrame:
    do_values = inclusive_grid(*bounds["Aerobic_DO"], DO_STEP)
    hrt_values = inclusive_grid(*bounds["HRT"], HRT_STEP)
    feed_values = inclusive_grid(*bounds["Anaerobic_Feed_Fraction"], FEED_STEP)

    mi = pd.MultiIndex.from_product(
        [do_values, hrt_values, feed_values],
        names=["候选_DO", "候选_HRT", "候选_Feed"],
    )
    grid = mi.to_frame(index=False)
    grid["网格来源"] = "第一轮粗扫描"
    return grid


def add_current_operation_if_needed(grid: pd.DataFrame, background: Mapping[str, Any]) -> pd.DataFrame:
    current = rounded_key(
        [background["Aerobic_DO"], background["HRT"], background["Anaerobic_Feed_Fraction"]]
    )
    keys = set(
        rounded_key(vals)
        for vals in grid[["候选_DO", "候选_HRT", "候选_Feed"]].to_numpy(dtype=float)
    )
    out = grid.copy()
    if current not in keys:
        out = pd.concat(
            [
                out,
                pd.DataFrame(
                    [
                        {
                            "候选_DO": float(background["Aerobic_DO"]),
                            "候选_HRT": float(background["HRT"]),
                            "候选_Feed": float(background["Anaerobic_Feed_Fraction"]),
                            "网格来源": "当前实际工况补充",
                        }
                    ]
                ),
            ],
            ignore_index=True,
        )
    return out.drop_duplicates(["候选_DO", "候选_HRT", "候选_Feed"], keep="first").reset_index(drop=True)



# =============================================================================
# 6. 单个历史状态的情景构造与批量预测
# =============================================================================

def build_target_scenario_features(
    scenarios: pd.DataFrame,
    target_name: str,
    processed_target_data: pd.DataFrame,
) -> pd.DataFrame:
    cfg = TARGET_CONFIG[target_name]
    lookup = processed_target_data.set_index(processed_target_data["样本序号"].astype(int), drop=False)
    sample_ids = scenarios["样本序号"].astype(int)

    out = pd.DataFrame(index=scenarios.index)
    out["Aerobic_DO"] = pd.to_numeric(scenarios["候选_DO"], errors="coerce").to_numpy(dtype=float)
    out["Anaerobic_Feed_Fraction"] = pd.to_numeric(scenarios["候选_Feed"], errors="coerce").to_numpy(dtype=float)
    out["HRT"] = pd.to_numeric(scenarios["候选_HRT"], errors="coerce").to_numpy(dtype=float)
    out["Anaerobic_Volume_Fraction"] = pd.to_numeric(
        scenarios["Anaerobic_Volume_Fraction"], errors="coerce"
    ).to_numpy(dtype=float)
    out["Anoxic_Volume_Fraction"] = pd.to_numeric(
        scenarios["Anoxic_Volume_Fraction"], errors="coerce"
    ).to_numpy(dtype=float)

    for feature in INFLUENT_FEATURES:
        out[feature] = sample_ids.map(lookup[feature]).to_numpy(dtype=float)
    for feature in cfg["history"]:
        out[feature] = sample_ids.map(lookup[feature]).to_numpy(dtype=float)
    return out


def build_scenarios_for_background(
    background: Mapping[str, Any],
    grid: pd.DataFrame,
) -> pd.DataFrame:
    out = grid.copy()
    out.insert(0, "历史状态ID", str(background["历史状态ID"]))
    out.insert(1, "样本序号", int(background["样本序号"]))
    out.insert(2, "Date", pd.to_datetime(background["Date"]))

    out["当前_DO"] = float(background["Aerobic_DO"])
    out["当前_HRT"] = float(background["HRT"])
    out["当前_Feed"] = float(background["Anaerobic_Feed_Fraction"])
    out["当前_Q_L_h"] = REACTOR_VOLUME_L / float(background["HRT"])
    out["候选_Q_L_h"] = REACTOR_VOLUME_L / pd.to_numeric(out["候选_HRT"], errors="coerce")
    out["候选_Q_L_d"] = REACTOR_VOLUME_L * 24.0 / pd.to_numeric(out["候选_HRT"], errors="coerce")

    out["Anaerobic_Volume_Fraction"] = float(background["Anaerobic_Volume_Fraction"])
    out["Anoxic_Volume_Fraction"] = float(background["Anoxic_Volume_Fraction"])

    current_mask = (
        np.isclose(out["候选_DO"].to_numpy(dtype=float), float(background["Aerobic_DO"]), atol=1e-10)
        & np.isclose(out["候选_HRT"].to_numpy(dtype=float), float(background["HRT"]), atol=1e-10)
        & np.isclose(
            out["候选_Feed"].to_numpy(dtype=float),
            float(background["Anaerobic_Feed_Fraction"]),
            atol=1e-10,
        )
    )
    out["是否当前实际工况"] = np.where(current_mask, "是", "否")

    physical = (
        pd.to_numeric(out["候选_DO"], errors="coerce").ge(0)
        & pd.to_numeric(out["候选_HRT"], errors="coerce").gt(0)
        & pd.to_numeric(out["候选_Feed"], errors="coerce").between(0, 1, inclusive="both")
        & (float(background["Anaerobic_Volume_Fraction"]) >= 0)
        & (float(background["Anoxic_Volume_Fraction"]) >= 0)
        & (
            float(background["Anaerobic_Volume_Fraction"])
            + float(background["Anoxic_Volume_Fraction"])
            < 1.0 - 1e-12
        )
    )
    out["物理条件是否合理"] = np.where(physical.fillna(False), "是", "否")
    return out.loc[out["物理条件是否合理"] == "是"].reset_index(drop=True)


def add_prediction_physical_check(result: pd.DataFrame) -> pd.DataFrame:
    """
    对四项水质预测做最基本的物理/数值检查。

    规则：
    - 任一预测非有限：标记为“存在非有限预测”；
    - 否则任一预测低于 -PREDICTION_NEGATIVE_TOL：标记为“存在负预测”；
    - 其余标记为“通过”。

    原始模型预测值不做截断、不改写，便于审计。后续反向正式筛选应只使用
    “预测物理检查=通过”的第一轮粗扫描记录。
    """
    out = result.copy()
    negative_names: List[Tuple[str, np.ndarray]] = []
    nonfinite_names: List[Tuple[str, np.ndarray]] = []
    negative_flags: List[np.ndarray] = []
    nonfinite_flags: List[np.ndarray] = []

    for target in TARGET_ORDER:
        col = PREDICTION_COLUMN[target]
        values = pd.to_numeric(out[col], errors="coerce").to_numpy(dtype=float)
        finite = np.isfinite(values)
        negative = finite & (values < -PREDICTION_NEGATIVE_TOL)
        nonfinite = ~finite
        negative_flags.append(negative)
        nonfinite_flags.append(nonfinite)
        negative_names.append((target, negative))
        nonfinite_names.append((target, nonfinite))

    any_negative = (
        np.logical_or.reduce(negative_flags)
        if negative_flags
        else np.zeros(len(out), dtype=bool)
    )
    any_nonfinite = (
        np.logical_or.reduce(nonfinite_flags)
        if nonfinite_flags
        else np.zeros(len(out), dtype=bool)
    )

    status = np.full(len(out), "通过", dtype=object)
    status[any_negative] = "存在负预测"
    # 非有限预测优先级更高，避免数值异常被后续阈值比较误判。
    status[any_nonfinite] = "存在非有限预测"
    out["预测物理检查"] = status

    negative_labels: List[str] = []
    nonfinite_labels: List[str] = []
    for i in range(len(out)):
        negative_labels.append(
            "、".join(target for target, flag in negative_names if bool(flag[i]))
        )
        nonfinite_labels.append(
            "、".join(target for target, flag in nonfinite_names if bool(flag[i]))
        )
    out["负预测指标"] = negative_labels
    out["非有限预测指标"] = nonfinite_labels
    return out


def add_a1_reference_flags(result: pd.DataFrame) -> pd.DataFrame:
    """
    为每条正向预测记录附加四项一级A工程参考判定。

    单项判定彼此独立。某一指标记为“是”仅要求该指标自身：
    1) 预测值为有限值；
    2) 预测值不低于允许的数值容差（即不存在该指标自身的负预测）；
    3) 预测值 <= 对应一级A工程参考限值。

    因此，如果某条记录仅 TP 出现负预测，而 TN/NH4/COD 自身预测正常且低于
    各自参考限值，则 TP 单项记“否”，其余三个单项仍可记“是”。

    “四项一级A参考同时满足”则更严格：必须整条记录预测物理检查=通过，
    并且四个单项判定全部为“是”。

    正向阶段只记录标签，不依据这些标签删除记录或生成推荐方案。
    """
    out = result.copy()
    physical_ok = out["预测物理检查"].astype(str).eq("通过").to_numpy(dtype=bool)
    all_single_meet = np.ones(len(out), dtype=bool)

    for target in TARGET_ORDER:
        pred_col = PREDICTION_COLUMN[target]
        flag_col = A1_REFERENCE_FLAG_COLUMN[target]
        limit = float(A1_REFERENCE_LIMITS[target])
        values = pd.to_numeric(out[pred_col], errors="coerce").to_numpy(dtype=float)

        # 单项标签只由该指标自身决定，不受其他指标是否异常影响。
        meet = (
            np.isfinite(values)
            & (values >= -PREDICTION_NEGATIVE_TOL)
            & (values <= limit + A1_SATISFY_TOL)
        )
        out[flag_col] = np.where(meet, "是", "否")
        all_single_meet &= meet

    # 联合标签仍要求整条记录四项预测均通过物理/数值检查。
    all_meet = physical_ok & all_single_meet
    out[A1_ALL_FLAG_COLUMN] = np.where(all_meet, "是", "否")
    return out


def predict_scenarios(
    scenarios: pd.DataFrame,
    processed_by_target: Mapping[str, pd.DataFrame],
    packages: Mapping[str, Mapping[str, Any]],
    historical_operation_lookup: Mapping[Tuple[float, ...], str],
) -> pd.DataFrame:
    """对同一历史状态下的 DO×HRT×Feed 候选批量调用四个冻结模型。"""
    result = scenarios.copy()

    for target in TARGET_ORDER:
        package = packages[target]
        feature_names = list(package["feature_names"])
        features = build_target_scenario_features(result, target, processed_by_target[target])
        if features[feature_names].isna().any().any():
            bad = features[feature_names].columns[features[feature_names].isna().any()].tolist()
            raise RuntimeError(f"{target} 情景输入仍有缺失：{bad}")

        x = features[feature_names].to_numpy(dtype=float)
        pred = np.asarray(package["model"].predict(package["scaler"].transform(x)), dtype=float)
        result[PREDICTION_COLUMN[target]] = pred

    # 标记候选完整工况是否在历史中真实出现过（含当前固定结构）。
    # 该字段只用于追踪/审计，不作为一级A反向筛选的硬门槛。
    hist_ids: List[str] = []
    for _, row in result.iterrows():
        key = rounded_key(
            [
                row["候选_DO"],
                row["候选_Feed"],
                row["候选_HRT"],
                row["Anaerobic_Volume_Fraction"],
                row["Anoxic_Volume_Fraction"],
            ]
        )
        hist_ids.append(str(historical_operation_lookup.get(key, "")))
    result["历史真实运行组合ID"] = hist_ids
    result["是否历史真实完整工况"] = np.where(result["历史真实运行组合ID"].ne(""), "是", "否")

    # 物理/数值检查只标记，不修改原始模型预测值，也不删除候选。
    result = add_prediction_physical_check(result)

    # 附加四项一级A工程参考判定；仍不在正向阶段删除任何情景。
    result = add_a1_reference_flags(result)
    return result


# =============================================================================
# 7. 流式汇总与五结果Excel分块写出
# =============================================================================

@dataclass
class NumericAccumulator:
    count: int = 0
    value_sum: float = 0.0
    value_sum_sq: float = 0.0
    minimum: float = np.inf
    maximum: float = -np.inf

    def update(self, values: Sequence[float]) -> None:
        arr = np.asarray(values, dtype=float)
        arr = arr[np.isfinite(arr)]
        if len(arr) == 0:
            return
        self.count += int(len(arr))
        self.value_sum += float(np.sum(arr))
        self.value_sum_sq += float(np.sum(arr * arr))
        self.minimum = min(self.minimum, float(np.min(arr)))
        self.maximum = max(self.maximum, float(np.max(arr)))

    def to_row(self, target: str) -> Dict[str, Any]:
        if self.count == 0:
            return {"目标指标": target, "情景数": 0}
        mean = self.value_sum / self.count
        variance = max(self.value_sum_sq / self.count - mean * mean, 0.0)
        return {
            "目标指标": target,
            "情景数": self.count,
            "预测最小值": self.minimum,
            "预测均值": mean,
            "预测标准差": math.sqrt(variance),
            "预测最大值": self.maximum,
        }


@dataclass
class GlobalSummaryAccumulator:
    prediction_stats: Dict[str, NumericAccumulator] = field(
        default_factory=lambda: {t: NumericAccumulator() for t in TARGET_ORDER}
    )
    total_rows: int = 0
    coarse_rows: int = 0
    current_supplement_rows: int = 0
    actual_rows: int = 0
    historical_operation_rows: int = 0
    negative_prediction_rows: int = 0
    nonfinite_prediction_rows: int = 0
    invalid_prediction_rows: int = 0
    negative_prediction_counts: Dict[str, int] = field(
        default_factory=lambda: {t: 0 for t in TARGET_ORDER}
    )
    nonfinite_prediction_counts: Dict[str, int] = field(
        default_factory=lambda: {t: 0 for t in TARGET_ORDER}
    )
    a1_meet_rows: Dict[str, int] = field(
        default_factory=lambda: {t: 0 for t in TARGET_ORDER}
    )
    a1_meet_coarse_rows: Dict[str, int] = field(
        default_factory=lambda: {t: 0 for t in TARGET_ORDER}
    )
    all_a1_meet_rows: int = 0
    all_a1_meet_coarse_rows: int = 0

    def update(self, frame: pd.DataFrame) -> None:
        self.total_rows += len(frame)
        source = frame["网格来源"].astype(str)
        coarse_mask = source.eq(FORMAL_INVERSE_GRID_SOURCE).to_numpy(dtype=bool)
        self.coarse_rows += int(np.sum(coarse_mask))
        self.current_supplement_rows += int(source.eq("当前实际工况补充").sum())
        self.actual_rows += int(frame["是否当前实际工况"].eq("是").sum())

        any_negative = np.zeros(len(frame), dtype=bool)
        any_nonfinite = np.zeros(len(frame), dtype=bool)
        for target in TARGET_ORDER:
            pred_col = PREDICTION_COLUMN[target]
            values = pd.to_numeric(frame[pred_col], errors="coerce").to_numpy(dtype=float)
            finite = np.isfinite(values)
            negative = finite & (values < -PREDICTION_NEGATIVE_TOL)
            nonfinite = ~finite
            any_negative |= negative
            any_nonfinite |= nonfinite
            self.negative_prediction_counts[target] += int(np.sum(negative))
            self.nonfinite_prediction_counts[target] += int(np.sum(nonfinite))
            self.prediction_stats[target].update(values)

            flag_col = A1_REFERENCE_FLAG_COLUMN[target]
            meet = frame[flag_col].astype(str).eq("是").to_numpy(dtype=bool)
            self.a1_meet_rows[target] += int(np.sum(meet))
            self.a1_meet_coarse_rows[target] += int(np.sum(meet & coarse_mask))

        self.negative_prediction_rows += int(np.sum(any_negative))
        self.nonfinite_prediction_rows += int(np.sum(any_nonfinite))
        self.invalid_prediction_rows += int(np.sum(any_negative | any_nonfinite))

        all_meet = frame[A1_ALL_FLAG_COLUMN].astype(str).eq("是").to_numpy(dtype=bool)
        self.all_a1_meet_rows += int(np.sum(all_meet))
        self.all_a1_meet_coarse_rows += int(np.sum(all_meet & coarse_mask))

        self.historical_operation_rows += int(frame["是否历史真实完整工况"].eq("是").sum())




class SingleWorkbookExcelWriter:
    """负责写出一个结果工作簿；当前主程序会实例化十一次，把631个背景连续拆成十一个文件。"""

    def __init__(self, final_path: Path, rows_per_sheet: int, part_no: int) -> None:
        self.final_path = final_path
        self.part_no = int(part_no)
        self.tmp_path = final_path.with_name(final_path.stem + ".part.xlsx")
        if self.tmp_path.exists():
            self.tmp_path.unlink()
        self.rows_per_sheet = int(rows_per_sheet)
        self.workbook = xlsxwriter.Workbook(
            str(self.tmp_path),
            {
                "constant_memory": True,
                "strings_to_urls": False,
                "nan_inf_to_errors": True,
            },
        )
        self.header_fmt = self.workbook.add_format(
            {"bold": True, "bg_color": "#D9EAF7", "border": 1, "align": "center", "valign": "vcenter"}
        )
        self.section_fmt = self.workbook.add_format(
            {"bold": True, "bg_color": "#B4C6E7", "border": 1, "font_size": 11}
        )
        self.date_fmt = self.workbook.add_format({"num_format": "yyyy-mm-dd"})
        self.num_fmt = self.workbook.add_format({"num_format": "0.000000"})
        self.int_fmt = self.workbook.add_format({"num_format": "0"})
        self.detail_sheet_index = 0
        self.detail_ws = None
        self.detail_row = 0
        self.detail_data_rows = 0
        self.detail_columns: Optional[List[str]] = None
        self.detail_index_rows: List[Dict[str, Any]] = []
        self.sheet_start_background: Optional[str] = None
        self.sheet_end_background: Optional[str] = None
        self.closed = False

    @staticmethod
    def _excel_value(value: Any) -> Any:
        if value is None:
            return None
        if isinstance(value, pd.Timestamp):
            if pd.isna(value):
                return None
            return value.to_pydatetime()
        if isinstance(value, np.datetime64):
            ts = pd.Timestamp(value)
            return None if pd.isna(ts) else ts.to_pydatetime()
        if isinstance(value, (np.integer,)):
            return int(value)
        if isinstance(value, (np.floating,)):
            value = float(value)
        if isinstance(value, float) and not np.isfinite(value):
            return None
        if isinstance(value, str):
            return sanitize_excel_value(value)
        return value

    def _set_detail_widths(self, ws, columns: Sequence[str]) -> None:
        for i, col in enumerate(columns):
            name = str(col)
            width = min(max(len(name) + 2, 10), 24)
            if name in {"历史状态ID", "Date", "网格来源"}:
                width = max(width, 16)
            if "预测" in name:
                width = max(width, 12)
            ws.set_column(i, i, width)

    def _open_detail_sheet(self, columns: Sequence[str]) -> None:
        self.detail_sheet_index += 1
        sheet_name = f"明细_第{self.part_no}部分_{self.detail_sheet_index:02d}"
        self.detail_ws = self.workbook.add_worksheet(sheet_name)
        self.detail_columns = list(columns)
        self.detail_row = 1
        self.detail_data_rows = 0
        self.sheet_start_background = None
        self.sheet_end_background = None
        for c, col in enumerate(self.detail_columns):
            self.detail_ws.write(0, c, col, self.header_fmt)
        self.detail_ws.freeze_panes(1, 0)
        self._set_detail_widths(self.detail_ws, self.detail_columns)

    def _finish_detail_sheet(self) -> None:
        if self.detail_ws is None or self.detail_columns is None:
            return
        if self.detail_data_rows > 0:
            self.detail_ws.autofilter(
                0, 0, self.detail_data_rows, len(self.detail_columns) - 1
            )
        self.detail_index_rows.append(
            {
                "明细副表序号": self.detail_sheet_index,
                "副表名": self.detail_ws.get_name(),
                "数据行数": self.detail_data_rows,
                "起始历史状态ID": self.sheet_start_background,
                "结束历史状态ID": self.sheet_end_background,
            }
        )
        self.detail_ws = None
        self.detail_columns = None
        self.detail_row = 0
        self.detail_data_rows = 0

    def write_detail(self, frame: pd.DataFrame) -> None:
        if frame.empty:
            return
        clean = sanitize_frame(frame)
        columns = list(clean.columns)
        start = 0
        while start < len(clean):
            if self.detail_ws is None:
                self._open_detail_sheet(columns)
            elif self.detail_columns != columns:
                raise RuntimeError("不同批次正向明细列顺序不一致。")

            remaining = self.rows_per_sheet - self.detail_data_rows
            if remaining <= 0:
                self._finish_detail_sheet()
                continue

            part = clean.iloc[start : start + remaining]
            for values in part.itertuples(index=False, name=None):
                row_values = [self._excel_value(v) for v in values]
                self.detail_ws.write_row(self.detail_row, 0, row_values)
                self.detail_row += 1
            self.detail_data_rows += len(part)

            first_bg = str(part["历史状态ID"].iloc[0])
            last_bg = str(part["历史状态ID"].iloc[-1])
            if self.sheet_start_background is None:
                self.sheet_start_background = first_bg
            self.sheet_end_background = last_bg
            start += len(part)

            if self.detail_data_rows >= self.rows_per_sheet:
                self._finish_detail_sheet()

    def finish_detail(self) -> pd.DataFrame:
        self._finish_detail_sheet()
        return pd.DataFrame(self.detail_index_rows)

    def _write_dataframe(self, ws, start_row: int, frame: pd.DataFrame, title: Optional[str] = None) -> int:
        clean = sanitize_frame(frame if frame is not None else pd.DataFrame())
        row = int(start_row)
        if title:
            ws.write(row, 0, title, self.section_fmt)
            row += 1
        if clean.empty and len(clean.columns) == 0:
            clean = pd.DataFrame([{"说明": "无记录"}])
        columns = list(clean.columns)
        for c, col in enumerate(columns):
            ws.write(row, c, col, self.header_fmt)
        header_row = row
        row += 1
        for values in clean.itertuples(index=False, name=None):
            for c, value in enumerate(values):
                excel_value = self._excel_value(value)
                if hasattr(excel_value, "year") and hasattr(excel_value, "month") and hasattr(excel_value, "day"):
                    ws.write_datetime(row, c, excel_value, self.date_fmt)
                elif isinstance(excel_value, float):
                    ws.write_number(row, c, excel_value, self.num_fmt)
                elif isinstance(excel_value, int) and not isinstance(excel_value, bool):
                    ws.write_number(row, c, excel_value, self.int_fmt)
                else:
                    ws.write(row, c, excel_value)
            row += 1
        # 同一工作表只能设置一个自动筛选。仅对单表型副表启用；
        # 多段合并副表不重复设置筛选，避免Excel范围冲突。
        if columns and len(clean) > 0 and title is None:
            ws.autofilter(header_row, 0, row - 1, len(columns) - 1)
        for c, col in enumerate(columns):
            sample = clean[col].head(300).astype(str) if col in clean.columns else pd.Series(dtype=str)
            if len(sample):
                width = int(sample.map(len).quantile(0.95)) + 2
            else:
                width = 10
            width = min(max(width, len(str(col)) + 2, 10), 36)
            ws.set_column(c, c, width)
        return row + 2

    def write_section_sheet(self, sheet_name: str, sections: Sequence[Tuple[str, pd.DataFrame]]) -> None:
        ws = self.workbook.add_worksheet(sheet_name[:31])
        ws.freeze_panes(1, 0)
        row = 0
        for title, frame in sections:
            row = self._write_dataframe(ws, row, frame, title=title)

    def write_table_sheet(self, sheet_name: str, frame: pd.DataFrame) -> None:
        ws = self.workbook.add_worksheet(sheet_name[:31])
        ws.freeze_panes(1, 0)
        self._write_dataframe(ws, 0, frame, title=None)

    def close(self) -> None:
        if self.closed:
            return
        self._finish_detail_sheet()
        self.workbook.close()
        os.replace(self.tmp_path, self.final_path)
        self.closed = True

    def abort(self) -> None:
        if self.closed:
            return
        try:
            self.workbook.close()
        except Exception:
            pass
        try:
            if self.tmp_path.exists():
                self.tmp_path.unlink()
        except Exception:
            pass
        self.closed = True


# =============================================================================
# 8. 汇总表、误差参考和 TXT 报告
# =============================================================================

def extract_validation_metric_table(
    packages: Mapping[str, Mapping[str, Any]],
) -> Tuple[pd.DataFrame, Dict[str, float]]:
    rows: List[Dict[str, Any]] = []
    main_rmse: Dict[str, float] = {}

    for target in TARGET_ORDER:
        source = packages[target].get("source_validation_metrics", {})
        boot = source.get("bootstrap_632plus", {}) or {}
        time_metrics = source.get("time", {}) or {}
        expanding = source.get("expanding", {}) or {}

        exp_rmse = safe_float(expanding.get("RMSE"))
        time_rmse = safe_float(time_metrics.get("RMSE"))
        valid_time_rmse = [
            value for value in (time_rmse, exp_rmse)
            if np.isfinite(value) and value > 0
        ]
        # 单值误差参考采用两种时间验证RMSE中较大的有效值，避免只取较小值显得过于乐观。
        main_rmse[target] = max(valid_time_rmse) if valid_time_rmse else 1.0

        rows.append(
            {
                "目标指标": target,
                "Bootstrap_R2": safe_float(boot.get("R2")),
                "Bootstrap_RMSE": safe_float(boot.get("RMSE")),
                "Bootstrap_MAE": safe_float(boot.get("MAE")),
                "Time_R2": safe_float(time_metrics.get("R2")),
                "Time_RMSE": time_rmse,
                "Time_MAE": safe_float(time_metrics.get("MAE")),
                "Expanding_R2": safe_float(expanding.get("R2")),
                "Expanding_RMSE": exp_rmse,
                "Expanding_MAE": safe_float(expanding.get("MAE")),
                "本脚本主要误差参考": "Time RMSE 与 Expanding RMSE 中较大的有效值",
                "主要误差参考_RMSE": main_rmse[target],
                "两种时间验证均优于最佳历史基线": (
                    "是"
                    if bool(source.get("beats_best_historical_baseline_in_both_time_validations", False))
                    else "否"
                ),
                "最终选择说明": str(source.get("final_selection_conclusion", "")),
                "说明": "均来自模型包保存的既有样本外验证结果，本脚本不重新验证或训练。",
            }
        )
    return pd.DataFrame(rows), main_rmse


def build_model_check_table(
    model_paths: Mapping[str, Path],
    packages: Mapping[str, Mapping[str, Any]],
    processed_by_target: Mapping[str, pd.DataFrame],
    full_metrics: Mapping[str, Mapping[str, float]],
    metric_diffs: Mapping[str, Mapping[str, float]],
) -> pd.DataFrame:
    rows: List[Dict[str, Any]] = []
    for target in TARGET_ORDER:
        package = packages[target]
        diffs = metric_diffs[target]
        rows.append(
            {
                "目标指标": target,
                "特征组": package["selected_feature_group"],
                "算法": package["algorithm"],
                "候选参数ID": package["candidate_id"],
                "参数JSON": json.dumps(package["parameters"], ensure_ascii=False, sort_keys=True),
                "特征数": len(package["feature_names"]),
                "训练样本数": len(processed_by_target[target]),
                "模型文件": str(model_paths[target]),
                "模型文件SHA256": sha256_file(model_paths[target]),
                "动态数据SHA256": sha256_file(INPUT_FILE),
                "全量拟合R2_仅核查": full_metrics[target]["R2"],
                "全量拟合RMSE_仅核查": full_metrics[target]["RMSE"],
                "全量拟合MAE_仅核查": full_metrics[target]["MAE"],
                "全量拟合复现最大指标差": max(diffs.values()) if diffs else 0.0,
                "核查状态": "通过",
            }
        )
    return pd.DataFrame(rows)


def build_grid_config_table(
    bounds: Mapping[str, Tuple[float, float]],
    coarse_grid: pd.DataFrame,
) -> pd.DataFrame:
    do_values = inclusive_grid(*bounds["Aerobic_DO"], DO_STEP)
    hrt_values = inclusive_grid(*bounds["HRT"], HRT_STEP)
    feed_values = inclusive_grid(*bounds["Anaerobic_Feed_Fraction"], FEED_STEP)

    n_do, n_hrt, n_feed = len(do_values), len(hrt_values), len(feed_values)
    expected = n_do * n_hrt * n_feed
    if expected != len(coarse_grid):
        raise RuntimeError(
            f"粗网格组合数核查失败：{n_do}×{n_hrt}×{n_feed}={expected}，实际={len(coarse_grid)}。"
        )

    rows = [
        {
            "变量": "Aerobic_DO",
            "搜索下限": bounds["Aerobic_DO"][0],
            "搜索上限": bounds["Aerobic_DO"][1],
            "粗扫描步长": DO_STEP,
            "候选水平数": n_do,
            "生成规则": "按步长生成并强制包含上下界",
            "备注": "运行调节变量",
        },
        {
            "变量": "HRT",
            "搜索下限": bounds["HRT"][0],
            "搜索上限": bounds["HRT"][1],
            "粗扫描步长": HRT_STEP,
            "候选水平数": n_hrt,
            "生成规则": "按步长生成并强制包含上下界",
            "备注": "改变后同步换算Q=57.6/HRT；Q不进入模型",
        },
        {
            "变量": "Anaerobic_Feed_Fraction",
            "搜索下限": bounds["Anaerobic_Feed_Fraction"][0],
            "搜索上限": bounds["Anaerobic_Feed_Fraction"][1],
            "粗扫描步长": FEED_STEP,
            "候选水平数": n_feed,
            "生成规则": "按步长生成并强制包含上下界",
            "备注": "运行调节变量",
        },
    ]
    table = pd.DataFrame(rows)
    table["粗扫描组合数计算"] = f"{n_do}×{n_hrt}×{n_feed}={expected}"
    table["粗扫描全局组合数"] = len(coarse_grid)
    table["正向阶段局部细化"] = "关闭"
    table["后续反向筛选正式来源"] = FORMAL_INVERSE_GRID_SOURCE
    return table


def build_a1_reference_summary(global_acc: GlobalSummaryAccumulator) -> pd.DataFrame:
    rows: List[Dict[str, Any]] = []
    for target in TARGET_ORDER:
        rows.append(
            {
                "参考项目": target,
                "一级A工程参考限值_mg_L": A1_REFERENCE_LIMITS[target],
                "全部正向记录满足数": global_acc.a1_meet_rows[target],
                "全部正向记录满足比例%": (
                    100.0 * global_acc.a1_meet_rows[target] / max(global_acc.total_rows, 1)
                ),
                "第一轮粗扫描满足数": global_acc.a1_meet_coarse_rows[target],
                "第一轮粗扫描满足比例%": (
                    100.0 * global_acc.a1_meet_coarse_rows[target] / max(global_acc.coarse_rows, 1)
                ),
                "说明": "仅为一步条件预测的工程参考判定，不等同于法规意义上的日均达标。",
            }
        )

    rows.append(
        {
            "参考项目": "四项同时满足",
            "一级A工程参考限值_mg_L": np.nan,
            "全部正向记录满足数": global_acc.all_a1_meet_rows,
            "全部正向记录满足比例%": (
                100.0 * global_acc.all_a1_meet_rows / max(global_acc.total_rows, 1)
            ),
            "第一轮粗扫描满足数": global_acc.all_a1_meet_coarse_rows,
            "第一轮粗扫描满足比例%": (
                100.0 * global_acc.all_a1_meet_coarse_rows / max(global_acc.coarse_rows, 1)
            ),
            "说明": "TP<=0.5、TN<=15、NH4+-N<=5、COD<=50，且预测物理检查=通过。",
        }
    )
    return pd.DataFrame(rows)


def background_summary_row(frame: pd.DataFrame) -> Dict[str, Any]:
    source = frame["网格来源"].astype(str)
    coarse_mask = source.eq(FORMAL_INVERSE_GRID_SOURCE)
    row: Dict[str, Any] = {
        "历史状态ID": frame["历史状态ID"].iloc[0],
        "样本序号": int(frame["样本序号"].iloc[0]),
        "Date": pd.to_datetime(frame["Date"].iloc[0]),
        "当前_DO": float(frame["当前_DO"].iloc[0]),
        "当前_HRT": float(frame["当前_HRT"].iloc[0]),
        "当前_Feed": float(frame["当前_Feed"].iloc[0]),
        "候选总数": len(frame),
        "规则粗扫描候选数": int(coarse_mask.sum()),
        "当前工况额外补充候选数": int(source.eq("当前实际工况补充").sum()),
        "预测物理检查未通过候选数": int(frame["预测物理检查"].astype(str).ne("通过").sum()),
        "存在负预测候选数": int(frame["负预测指标"].astype(str).ne("").sum()),
        "存在非有限预测候选数": int(frame["非有限预测指标"].astype(str).ne("").sum()),
        "历史真实完整工况候选数": int(frame["是否历史真实完整工况"].eq("是").sum()),
        "四项一级A参考同时满足候选数": int(frame[A1_ALL_FLAG_COLUMN].eq("是").sum()),
        "规则粗扫描四项一级A参考同时满足数": int(
            (coarse_mask & frame[A1_ALL_FLAG_COLUMN].eq("是")).sum()
        ),
    }
    for target in TARGET_ORDER:
        col = PREDICTION_COLUMN[target]
        prefix = "NH4" if target == "NH4+-N" else target
        values = pd.to_numeric(frame[col], errors="coerce")
        row[f"{prefix}_预测最小"] = float(values.min())
        row[f"{prefix}_预测中位数"] = float(values.median())
        row[f"{prefix}_预测最大"] = float(values.max())
        flag_col = A1_REFERENCE_FLAG_COLUMN[target]
        row[f"{prefix}_一级A参考满足候选数"] = int(frame[flag_col].eq("是").sum())
        row[f"规则粗扫描_{prefix}_一级A参考满足数"] = int(
            (coarse_mask & frame[flag_col].eq("是")).sum()
        )
    return row


def split_contiguous_history_parts(frame: pd.DataFrame, n_parts: int = 11) -> List[pd.DataFrame]:
    """按原顺序把历史状态尽量均匀地连续拆分；同一历史状态永远只属于一个部分。"""
    if n_parts <= 0:
        raise ValueError("n_parts 必须为正整数。")
    total = len(frame)
    base, remainder = divmod(total, n_parts)
    sizes = [base + (1 if i < remainder else 0) for i in range(n_parts)]
    parts: List[pd.DataFrame] = []
    start_idx = 0
    for size in sizes:
        stop_idx = start_idx + size
        parts.append(frame.iloc[start_idx:stop_idx].copy())
        start_idx = stop_idx
    if start_idx != total:
        raise RuntimeError("历史状态十一部分拆分失败：拆分后总数与原始总数不一致。")
    return parts


def build_text_report(
    elapsed_minutes: float,
    common_history: pd.DataFrame,
    historical_ops: pd.DataFrame,
    bounds: Mapping[str, Tuple[float, float]],
    coarse_grid: pd.DataFrame,
    validation_metrics: pd.DataFrame,
    model_check: pd.DataFrame,
    background_summary: pd.DataFrame,
    detail_index: pd.DataFrame,
    global_acc: GlobalSummaryAccumulator,
    warning_table: pd.DataFrame,
) -> str:
    lines: List[str] = []
    do_values = inclusive_grid(*bounds["Aerobic_DO"], DO_STEP)
    hrt_values = inclusive_grid(*bounds["HRT"], HRT_STEP)
    feed_values = inclusive_grid(*bounds["Anaerobic_Feed_Fraction"], FEED_STEP)
    n_do, n_hrt, n_feed = len(do_values), len(hrt_values), len(feed_values)
    per_background = len(coarse_grid)
    formal_coarse_total = len(common_history) * per_background

    history_parts = split_contiguous_history_parts(common_history, 11)
    part_stats: List[Dict[str, Any]] = []
    for part_no, history_part in enumerate(history_parts, start=1):
        ids = set(history_part["历史状态ID"].astype(str))
        summary_part = background_summary[background_summary["历史状态ID"].astype(str).isin(ids)]
        rows = int(pd.to_numeric(summary_part.get("候选总数", pd.Series(dtype=float)), errors="coerce").fillna(0).sum())
        coarse = int(pd.to_numeric(summary_part.get("规则粗扫描候选数", pd.Series(dtype=float)), errors="coerce").fillna(0).sum())
        supp = int(pd.to_numeric(summary_part.get("当前工况额外补充候选数", pd.Series(dtype=float)), errors="coerce").fillna(0).sum())
        part_stats.append(
            {
                "part_no": part_no,
                "history": history_part,
                "rows": rows,
                "coarse": coarse,
                "supp": supp,
            }
        )

    lines.append("A²/O 四指标 DO×HRT×进水分配比例联合正向预测说明")
    lines.append("=" * 100)
    lines.append("")

    lines.append("一、任务与文件")
    lines.append(f"1. 完成时间：{now_str()}")
    lines.append(f"2. 脚本版本：{SCRIPT_VERSION}")
    lines.append(f"3. 版本标识：{BUILD_FINGERPRINT}")
    lines.append(f"4. 总运行耗时：{elapsed_minutes:.2f} 分钟")
    lines.append(f"5. 动态数据：{INPUT_FILE}")
    lines.append(f"6. 最终模型目录：{MODEL_DIR}")
    lines.append(f"7. 输出目录：{OUTPUT_DIR}")
    lines.append(f"8. 正向预测结果Excel-第1部分：{OUTPUT_EXCEL_PART1}")
    lines.append(f"9. 正向预测结果Excel-第2部分：{OUTPUT_EXCEL_PART2}")
    lines.append(f"10. 正向预测结果Excel-第3部分：{OUTPUT_EXCEL_PART3}")
    lines.append(f"11. 正向预测结果Excel-第4部分：{OUTPUT_EXCEL_PART4}")
    lines.append(f"12. 正向预测结果Excel-第5部分：{OUTPUT_EXCEL_PART5}")
    lines.append(f"13. 正向预测结果Excel-第6部分：{OUTPUT_EXCEL_PART6}")
    lines.append(f"14. 正向预测结果Excel-第7部分：{OUTPUT_EXCEL_PART7}")
    lines.append(f"15. 正向预测结果Excel-第8部分：{OUTPUT_EXCEL_PART8}")
    lines.append(f"16. 正向预测结果Excel-第9部分：{OUTPUT_EXCEL_PART9}")
    lines.append(f"17. 正向预测结果Excel-第10部分：{OUTPUT_EXCEL_PART10}")
    lines.append(f"18. 正向预测结果Excel-第11部分：{OUTPUT_EXCEL_PART11}")
    lines.append(f"19. 十一个Excel合计明细副表数：{len(detail_index)}")
    lines.append(f"20. 说明TXT：{OUTPUT_TEXT}")
    lines.append("21. 当前版本不再生成旧版单一主Excel，也不生成‘四指标_DO-HRT-进水分配联合正向预测_结果检查包.xlsx’。")
    lines.append("22. 本脚本不重新训练模型、不重新调参，也不执行反向筛选。")
    lines.append("")

    lines.append("二、正向预测固定逻辑")
    lines.append(f"1. 四指标共同有效真实历史状态共 {len(common_history)} 条，按当前确认数据应为631条。")
    lines.append("2. 每个历史状态固定四项进水、各目标lag1/lag2/roll3以及两个反应区体积分数。")
    lines.append("3. 仅联合改变 Aerobic_DO、HRT 和 Anaerobic_Feed_Fraction。")
    lines.append("4. 四个冻结F2模型分别输出TP、TN、NH4+-N和COD的一步条件预测。")
    lines.append("5. 正向阶段不删除任何情景、不输出运行参数推荐；每条记录附加四项一级A工程参考判定字段。")
    lines.append("6. 一级A工程参考条件固定为TP<=0.5、TN<=15、NH4+-N<=5、COD<=50 mg/L；四项同时满足采用AND逻辑。")
    lines.append("7. 这些判定只是一步条件预测的工程参考标签，不等同于法规意义上的日均达标结论。")
    lines.append("8. 旧版历史5NN距离/历史接近程度不再计算，也不进入数百万行预测明细。")
    lines.append("")

    lines.append("三、三参数网格如何生成")
    lines.append(
        f"1. DO范围={bounds['Aerobic_DO'][0]:g}~{bounds['Aerobic_DO'][1]:g} mg/L，步长={DO_STEP:g} mg/L，"
        f"共 {n_do} 个候选水平。inclusive_grid会强制包含上界。"
    )
    lines.append(
        f"2. HRT范围={bounds['HRT'][0]:g}~{bounds['HRT'][1]:g} h，步长={HRT_STEP:g} h，共 {n_hrt} 个候选水平。"
    )
    lines.append(
        f"3. 厌氧区进水分配比例范围={bounds['Anaerobic_Feed_Fraction'][0]:g}~{bounds['Anaerobic_Feed_Fraction'][1]:g}，"
        f"步长={FEED_STEP:g}，共 {n_feed} 个候选水平。"
    )
    lines.append(f"4. 每个历史状态的统一粗网格组合数={n_do}×{n_hrt}×{n_feed}={per_background}组。")
    lines.append("5. 三个参数使用笛卡尔积联合枚举；每个组合都在同一历史状态下分别调用四个冻结模型。")
    lines.append("6. 正向阶段不做第二轮局部细化；后续反向阶段也直接筛选这一套统一粗网格。")
    lines.append("")

    lines.append("四、两百多万条正向记录如何得到")
    lines.append(f"1. 正式统一粗网格记录数={len(common_history)}×{per_background}={formal_coarse_total:,}条。")
    lines.append(
        f"2. 实际写出的规则粗扫描记录数={global_acc.coarse_rows:,}条；"
        f"与理论值之差={global_acc.coarse_rows-formal_coarse_total:,}。"
    )
    lines.append(
        f"3. 当前实际工况如果不在规则网格上，会额外补入作现状对照；本次额外补充={global_acc.current_supplement_rows:,}条。"
    )
    lines.append(
        f"4. 因此完整正向明细总行数={global_acc.coarse_rows:,}+{global_acc.current_supplement_rows:,}={global_acc.total_rows:,}条。"
    )
    lines.append("5. 后续一级A四指标反向筛选只读取‘网格来源=第一轮粗扫描’记录，不把当前实际工况补充行纳入统一组合统计。")
    lines.append("")

    lines.append("五、冻结模型与预测误差参考")
    for _, row in model_check.iterrows():
        lines.append(
            f"- {row['目标指标']}：{row['特征组']} + {row['算法']}（{row['候选参数ID']}），"
            f"训练样本={int(row['训练样本数'])}，模型核查={row['核查状态']}。"
        )
    for _, row in validation_metrics.iterrows():
        lines.append(
            f"- {row['目标指标']}：Time RMSE={row['Time_RMSE']:.6g}，Expanding RMSE={row['Expanding_RMSE']:.6g}，"
            f"主要RMSE参考={row['主要误差参考_RMSE']:.6g}。"
        )
    lines.append("上述误差只用于结果解释和审计，不参与正向预测的候选删除。")
    lines.append("")

    lines.append("六、物理检查与Q换算")
    lines.append("1. DO必须>=0，HRT必须>0，Feed必须位于0~1；固定厌氧/缺氧容积分数应满足两者之和<1。")
    lines.append("2. HRT改变后同步计算Q(L/h)=57.6/HRT和Q(L/d)=57.6×24/HRT；Q不进入模型。")
    lines.append("3. 四项水质预测若低于0，原始模型值保留并标记‘存在负预测’，不截断为0。")
    lines.append("4. 任一水质预测为NaN/Inf时标记‘存在非有限预测’；原始结果仍保留用于审计。")
    lines.append(f"5. 存在至少一项负预测的候选记录数={global_acc.negative_prediction_rows}。")
    lines.append(f"6. 存在至少一项非有限预测的候选记录数={global_acc.nonfinite_prediction_rows}。")
    lines.append(f"7. 预测物理/数值检查未通过记录数={global_acc.invalid_prediction_rows}。")
    for target in TARGET_ORDER:
        lines.append(
            f"   - {target}：负预测={global_acc.negative_prediction_counts[target]}；"
            f"非有限预测={global_acc.nonfinite_prediction_counts[target]}。"
        )
    lines.append("")

    lines.append("七、四项预测总体范围")
    for target in TARGET_ORDER:
        row = global_acc.prediction_stats[target].to_row(target)
        lines.append(
            f"- {target}：n={row.get('情景数', 0)}，最小={row.get('预测最小值', np.nan):.6g}，"
            f"均值={row.get('预测均值', np.nan):.6g}，最大={row.get('预测最大值', np.nan):.6g}。"
        )
    lines.append("")

    lines.append("八、一级A工程参考判定汇总")
    for target in TARGET_ORDER:
        limit = A1_REFERENCE_LIMITS[target]
        lines.append(
            f"- {target} <= {limit:g} mg/L：全部正向记录满足="
            f"{global_acc.a1_meet_rows[target]:,}/{global_acc.total_rows:,}；"
            f"第一轮粗扫描满足={global_acc.a1_meet_coarse_rows[target]:,}/{global_acc.coarse_rows:,}。"
        )
    lines.append(
        "- 四项同时满足（TP<=0.5、TN<=15、NH4+-N<=5、COD<=50，且预测物理检查=通过）："
        f"全部正向记录={global_acc.all_a1_meet_rows:,}/{global_acc.total_rows:,}；"
        f"第一轮粗扫描={global_acc.all_a1_meet_coarse_rows:,}/{global_acc.coarse_rows:,}。"
    )
    lines.append("单项标签按各指标自身的有限性、非负性和参考限值独立判定；四项联合标签另要求整条记录预测物理检查=通过。")
    lines.append("以上仅为工程参考筛选标签；正向阶段不据此删除记录或推荐工况。")
    lines.append("")

    lines.append("九、十一个结果Excel的拆分与后续反向筛选衔接")
    lines.append("1. 完整正向结果按历史状态ID连续拆为十一个Excel，进一步降低单个结果文件体积。")
    lines.append("2. 拆分采用尽量均匀原则；631个背景时为58+58+58+58+57+57+57+57+57+57+57。任一历史状态的全部预测记录只存在于其中一个文件，不跨文件拆分。")
    item_no = 3
    for stat in part_stats:
        history_part = stat["history"]
        if len(history_part):
            lines.append(
                f"{item_no}. 第{stat['part_no']}部分：{history_part['历史状态ID'].iloc[0]}~{history_part['历史状态ID'].iloc[-1]}，"
                f"共{len(history_part)}个历史状态；粗网格{stat['coarse']:,}条，当前工况补充{stat['supp']:,}条，合计{stat['rows']:,}条明细。"
            )
            item_no += 1
    lines.append(f"{item_no}. 为避免重复，01_说明与设置、02_模型与误差、03_历史状态、04_背景预测汇总只在第1部分保存一次；第2~11部分不再重复公共Sheet，只保留各自独立明细。")
    item_no += 1
    lines.append(f"{item_no}. 每个结果文件的明细按该部分预计总行数默认均分为约{DETAIL_TARGET_SHEETS_PER_WORKBOOK}个‘明细_’副表；同一历史状态仍只属于一个结果文件。")
    item_no += 1
    lines.append(f"{item_no}. 后续反向程序应同时读取十一个Excel中的所有‘明细_’副表，再合并为完整正向粗网格数据库；不重新运行四个模型。")
    item_no += 1
    lines.append(
        f"{item_no}. 后续仅保留第一轮粗扫描记录，并按‘预测物理检查=通过、"
        "TP<=0.5、TN<=15、NH4+-N<=5、COD<=50’四项AND筛选。"
    )
    item_no += 1
    lines.append(f"{item_no}. 正向明细已直接保存四个单项一级A参考判定以及‘四项一级A参考同时满足’字段；反向脚本仍建议按原始预测值独立复算一次并与标签交叉核查。")
    item_no += 1
    lines.append(f"{item_no}. 反向阶段不再使用调整档位、改变参数个数、默认推荐或调整最少方案等旧字段。")
    lines.append("")

    lines.append("十、解释边界")
    lines.append("1. 所有结果均为冻结模型的一步条件预测，不是长期稳态模拟。")
    lines.append("2. 正向阶段只建立运行参数与预测出水之间的对应关系，不宣称某个参数组合已经通过真实反应器验证。")
    lines.append("3. NH4+-N若对运行参数几乎无预测变化，只说明当前最终ML没有学到明显可调信号。")
    lines.append("")

    lines.append("十一、运行警告")
    lines.append(f"警告/异常记录数={len(warning_table)}。")
    if len(warning_table):
        for _, row in warning_table.head(20).iterrows():
            lines.append(f"- {row.get('位置', '')}：{row.get('内容', '')}")
        if len(warning_table) > 20:
            lines.append(f"- 其余{len(warning_table)-20}条请查看第1部分Excel‘01_说明与设置’中的警告部分。")

    return "\n".join(lines)


# =============================================================================
# 9. 主程序
# =============================================================================

def main() -> None:
    start_time = time.time()
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    cleanup_previous_forward_outputs(OUTPUT_DIR)
    warnings_rows: List[Dict[str, Any]] = []
    excel_writers: List[SingleWorkbookExcelWriter] = []

    log("A²/O 四指标 DO×HRT×进水分配联合正向预测开始。")
    log(f"脚本版本：{SCRIPT_VERSION}")
    log(f"动态数据：{INPUT_FILE}")
    log(f"模型目录：{MODEL_DIR}")
    log(f"输出目录：{OUTPUT_DIR}")
    log("本版建立统一粗网格正向预测库，并附加四项一级A工程参考判定；不删除情景、不运行NSGA-II、不做局部细化。")
    log("完整结果按历史状态连续拆分为十一个Excel；公共Sheet只在第1部分保存一次；同一历史状态不会跨文件。")

    try:
        model_paths: Dict[str, Path] = {}
        packages: Dict[str, Dict[str, Any]] = {}
        raw_by_target: Dict[str, pd.DataFrame] = {}
        processed_by_target: Dict[str, pd.DataFrame] = {}
        metrics_by_target: Dict[str, Dict[str, float]] = {}
        metric_diffs_by_target: Dict[str, Dict[str, float]] = {}
        imputation_tables: List[pd.DataFrame] = []

        for target in TARGET_ORDER:
            log(f"读取并核查 {target} 冻结模型。")
            model_path, package = load_model_package(target)
            raw = read_target_data(target)
            processed, imputation = rebuild_training_data(target, raw, package)
            _, metrics, diffs = validate_full_fit(target, processed, package)
            model_paths[target] = model_path
            packages[target] = package
            raw_by_target[target] = raw
            processed_by_target[target] = processed
            metrics_by_target[target] = metrics
            metric_diffs_by_target[target] = diffs
            imputation_tables.append(imputation)

        validation_metric_table, _main_rmse = extract_validation_metric_table(packages)
        model_check = build_model_check_table(
            model_paths,
            packages,
            processed_by_target,
            metrics_by_target,
            metric_diffs_by_target,
        )

        log("构建四指标共同历史状态。")
        common_history, common_audit = build_common_history(raw_by_target)
        if STRICT_MODE and len(common_history) != 631:
            raise RuntimeError(
                f"当前四指标共同历史状态数={len(common_history)}，与已确认的631不一致。"
                "若数据版本确实改变，请先核查后再关闭严格模式。"
            )

        if common_history[ALL_OPERATION_FEATURES].isna().any().any():
            bad_cols = common_history[ALL_OPERATION_FEATURES].columns[
                common_history[ALL_OPERATION_FEATURES].isna().any()
            ].tolist()
            raise RuntimeError(f"共同历史状态的运行/结构变量存在缺失：{bad_cols}")

        historical_ops, historical_op_lookup = build_historical_operation_table(common_history)
        log(f"共同历史状态={len(common_history)}；历史真实完整运行组合={len(historical_ops)}。")

        current_op_ids: List[str] = []
        for _, row in common_history.iterrows():
            current_op_ids.append(str(historical_op_lookup.get(full_operation_key(row), "")))
        common_history["当前历史运行组合ID"] = current_op_ids

        bounds = resolve_search_bounds(common_history)
        coarse_grid = build_coarse_grid(bounds)
        grid_config = build_grid_config_table(bounds, coarse_grid)
        do_values = inclusive_grid(*bounds["Aerobic_DO"], DO_STEP)
        hrt_values = inclusive_grid(*bounds["HRT"], HRT_STEP)
        feed_values = inclusive_grid(*bounds["Anaerobic_Feed_Fraction"], FEED_STEP)
        log(
            f"第一轮粗网格：DO水平={len(do_values)}，HRT水平={len(hrt_values)}，"
            f"Feed水平={len(feed_values)}，组合数={len(coarse_grid)}。"
        )

        grid_override_envs = [
            "A2O_FORWARD_DO_MIN", "A2O_FORWARD_DO_MAX", "A2O_FORWARD_HRT_MIN", "A2O_FORWARD_HRT_MAX",
            "A2O_FORWARD_FEED_MIN", "A2O_FORWARD_FEED_MAX", "A2O_FORWARD_DO_STEP",
            "A2O_FORWARD_HRT_STEP", "A2O_FORWARD_FEED_STEP",
        ]
        grid_overridden = any(os.environ.get(name) not in (None, "") for name in grid_override_envs)
        if STRICT_MODE and not grid_overridden:
            expected_shape = (28, 9, 15, 3780)
            actual_shape = (len(do_values), len(hrt_values), len(feed_values), len(coarse_grid))
            if actual_shape != expected_shape:
                raise RuntimeError(
                    f"默认粗网格与已确认流程不一致：实际DO/HRT/Feed/组合数={actual_shape}，"
                    f"预期={expected_shape}。请先核查共同历史状态范围和输入数据版本。"
                )

        nh4_package = packages["NH4+-N"]
        nh4_model = nh4_package["model"]
        nh4_zero_response = False
        if hasattr(nh4_model, "coef_"):
            coef = np.asarray(nh4_model.coef_, dtype=float).reshape(-1)
            features = list(nh4_package["feature_names"])
            op_idx = [i for i, f in enumerate(features) if f in ALL_OPERATION_FEATURES]
            if op_idx:
                nh4_zero_response = bool(np.all(np.abs(coef[op_idx]) <= NH4_ZERO_COEF_TOL))
        nh4_note = pd.DataFrame(
            [
                {
                    "项目": "NH4最终模型运行变量系数是否全部近似为0",
                    "结果": "是" if nh4_zero_response else "否/不适用",
                    "解释": "若为是，NH4正向预测仍输出，但不能用该模型声称调整DO/HRT/Feed会改善实际NH4。",
                }
            ]
        )

        history_parts = split_contiguous_history_parts(common_history, 11)
        if STRICT_MODE:
            part_sizes = [len(part) for part in history_parts]
            if len(common_history) == 631 and part_sizes != [58, 58, 58, 58, 57, 57, 57, 57, 57, 57, 57]:
                raise RuntimeError(
                    f"631个背景的十一文件拆分数量异常：实际={part_sizes}，"
                    "预期=[58, 58, 58, 58, 57, 57, 57, 57, 57, 57, 57]。"
                )

        part_background_ids: List[set[str]] = [
            set(part["历史状态ID"].astype(str)) for part in history_parts
        ]
        background_to_part: Dict[str, int] = {}
        for part_idx, ids in enumerate(part_background_ids):
            for bg_id in ids:
                if bg_id in background_to_part:
                    raise RuntimeError(f"历史状态 {bg_id} 被重复分配到多个结果文件。")
                background_to_part[bg_id] = part_idx
        if len(background_to_part) != len(common_history):
            raise RuntimeError("十一文件背景分配不完整。")

        for part_no, history_part in enumerate(history_parts, start=1):
            if history_part.empty:
                bg_range = "无"
            else:
                bg_range = f"{history_part['历史状态ID'].iloc[0]}~{history_part['历史状态ID'].iloc[-1]}"
            log(
                f"结果Excel第{part_no}部分：历史状态数={len(history_part)}，范围={bg_range}，"
                f"理论粗网格={len(history_part) * len(coarse_grid):,}条。"
            )

        detail_rows_per_part: List[int] = []
        for history_part in history_parts:
            # 每个历史状态最多在统一粗网格之外补1条“当前实际工况”。
            max_expected_rows = len(history_part) * (len(coarse_grid) + 1)
            auto_rows_per_sheet = max(
                1,
                int(math.ceil(max_expected_rows / DETAIL_TARGET_SHEETS_PER_WORKBOOK)),
            )
            # 保留原环境变量作为单Sheet硬上限；默认情况下自动拆为约3个Sheet。
            detail_rows_per_part.append(min(DETAIL_ROWS_PER_SHEET, auto_rows_per_sheet))

        excel_writers = [
            SingleWorkbookExcelWriter(path, detail_rows_per_part[i], i + 1)
            for i, path in enumerate(OUTPUT_EXCEL_PARTS)
        ]
        for part_no, rows_per_sheet in enumerate(detail_rows_per_part, start=1):
            log(
                f"第{part_no}部分明细单Sheet目标上限={rows_per_sheet:,}行，"
                f"预计拆为约{DETAIL_TARGET_SHEETS_PER_WORKBOOK}个明细Sheet。"
            )
        global_acc = GlobalSummaryAccumulator()
        background_summary_rows: List[Dict[str, Any]] = []

        total_backgrounds = len(common_history)
        for bg_i, (_, bg) in enumerate(common_history.iterrows(), start=1):
            bg_id = str(bg["历史状态ID"])
            if bg_i == 1 or bg_i % 20 == 0 or bg_i == total_backgrounds:
                log(f"正向预测进度：{bg_i}/{total_backgrounds}，当前={bg_id}。")

            bg_grid = add_current_operation_if_needed(coarse_grid, bg)
            coarse_scenarios = build_scenarios_for_background(bg, bg_grid)
            coarse_result = predict_scenarios(
                coarse_scenarios,
                processed_by_target,
                packages,
                historical_op_lookup,
            )

            combined = coarse_result.reset_index(drop=True)

            front_cols = [
                "历史状态ID", "样本序号", "Date", "网格来源",
                "当前_DO", "当前_HRT", "当前_Feed", "当前_Q_L_h",
                "候选_DO", "候选_HRT", "候选_Feed", "候选_Q_L_h", "候选_Q_L_d",
                "Anaerobic_Volume_Fraction", "Anoxic_Volume_Fraction",
                "是否当前实际工况", "是否历史真实完整工况", "历史真实运行组合ID",
                "预测物理检查", "负预测指标", "非有限预测指标",
                "TP_预测", "TN_预测", "NH4_预测", "COD_预测",
                "TP_一级A参考满足", "TN_一级A参考满足",
                "NH4_一级A参考满足", "COD_一级A参考满足",
                "四项一级A参考同时满足",
            ]
            remaining = [c for c in combined.columns if c not in front_cols]
            combined = combined[front_cols + remaining]

            global_acc.update(combined)
            background_summary_rows.append(background_summary_row(combined))

            detail_cols = [
                "历史状态ID", "样本序号", "Date", "网格来源",
                "当前_DO", "当前_HRT", "当前_Feed", "当前_Q_L_h",
                "候选_DO", "候选_HRT", "候选_Feed", "候选_Q_L_h", "候选_Q_L_d",
                "是否当前实际工况", "是否历史真实完整工况",
                "预测物理检查", "负预测指标", "非有限预测指标",
                "TP_预测", "TN_预测", "NH4_预测", "COD_预测",
                "TP_一级A参考满足", "TN_一级A参考满足",
                "NH4_一级A参考满足", "COD_一级A参考满足",
                "四项一级A参考同时满足",
            ]
            detail_frame = combined[detail_cols].copy()
            part_idx = background_to_part[bg_id]
            excel_writers[part_idx].write_detail(detail_frame)

        expected_coarse_total = len(common_history) * len(coarse_grid)
        if STRICT_MODE and global_acc.coarse_rows != expected_coarse_total:
            raise RuntimeError(
                f"统一粗网格总记录数核查失败：实际={global_acc.coarse_rows}，"
                f"理论={len(common_history)}×{len(coarse_grid)}={expected_coarse_total}。"
            )
        if STRICT_MODE and global_acc.total_rows != global_acc.coarse_rows + global_acc.current_supplement_rows:
            raise RuntimeError(
                f"正向明细总数核查失败：总数={global_acc.total_rows}，"
                f"粗网格+当前工况补充={global_acc.coarse_rows}+{global_acc.current_supplement_rows}。"
            )
        if STRICT_MODE:
            if global_acc.all_a1_meet_rows > global_acc.total_rows:
                raise RuntimeError("四项一级A参考同时满足记录数超过正向明细总数。")
            if global_acc.all_a1_meet_coarse_rows > global_acc.coarse_rows:
                raise RuntimeError("粗网格四项一级A参考同时满足记录数超过粗网格总数。")
            if global_acc.invalid_prediction_rows > global_acc.total_rows:
                raise RuntimeError("预测物理/数值检查未通过记录数超过正向明细总数。")

        detail_indexes: List[pd.DataFrame] = []
        for part_no, (writer, output_path) in enumerate(zip(excel_writers, OUTPUT_EXCEL_PARTS), start=1):
            detail_part = writer.finish_detail()
            if not detail_part.empty:
                detail_part.insert(0, "结果文件", output_path.name)
                detail_part.insert(1, "结果部分", f"第{part_no}部分")
            detail_indexes.append(detail_part)
        detail_index = pd.concat(detail_indexes, ignore_index=True) if detail_indexes else pd.DataFrame()

        background_summary = pd.DataFrame(background_summary_rows)
        imputation_audit = pd.concat(imputation_tables, ignore_index=True)

        prediction_global = pd.DataFrame(
            [global_acc.prediction_stats[t].to_row(t) for t in TARGET_ORDER]
        )

        physical_check_summary = pd.DataFrame(
            [
                {
                    "目标指标": target,
                    "负预测记录数": global_acc.negative_prediction_counts[target],
                    "非有限预测记录数": global_acc.nonfinite_prediction_counts[target],
                    "负预测占全部候选比例%": 100.0 * global_acc.negative_prediction_counts[target] / max(global_acc.total_rows, 1),
                    "非有限预测占全部候选比例%": 100.0 * global_acc.nonfinite_prediction_counts[target] / max(global_acc.total_rows, 1),
                    "处理原则": "保留原始预测并标记；后续正式反向筛选排除物理/数值检查未通过记录",
                }
                for target in TARGET_ORDER
            ]
        )

        a1_reference_summary = build_a1_reference_summary(global_acc)

        run_settings_rows: List[Dict[str, Any]] = [
            {"项目": "脚本版本", "值": SCRIPT_VERSION},
            {"项目": "版本标识", "值": BUILD_FINGERPRINT},
            {"项目": "动态数据", "值": str(INPUT_FILE)},
            {"项目": "动态数据SHA256", "值": sha256_file(INPUT_FILE)},
            {"项目": "最终模型目录", "值": str(MODEL_DIR)},
            {"项目": "输出目录", "值": str(OUTPUT_DIR)},
            {"项目": "正向预测结果Excel_第1部分", "值": str(OUTPUT_EXCEL_PART1)},
            {"项目": "正向预测结果Excel_第2部分", "值": str(OUTPUT_EXCEL_PART2)},
            {"项目": "正向预测结果Excel_第3部分", "值": str(OUTPUT_EXCEL_PART3)},
            {"项目": "正向预测结果Excel_第4部分", "值": str(OUTPUT_EXCEL_PART4)},
            {"项目": "正向预测结果Excel_第5部分", "值": str(OUTPUT_EXCEL_PART5)},
            {"项目": "正向预测结果Excel_第6部分", "值": str(OUTPUT_EXCEL_PART6)},
            {"项目": "正向预测结果Excel_第7部分", "值": str(OUTPUT_EXCEL_PART7)},
            {"项目": "正向预测结果Excel_第8部分", "值": str(OUTPUT_EXCEL_PART8)},
            {"项目": "正向预测结果Excel_第9部分", "值": str(OUTPUT_EXCEL_PART9)},
            {"项目": "正向预测结果Excel_第10部分", "值": str(OUTPUT_EXCEL_PART10)},
            {"项目": "正向预测结果Excel_第11部分", "值": str(OUTPUT_EXCEL_PART11)},
            {"项目": "共同历史状态数", "值": len(common_history)},
            {"项目": "历史真实完整运行组合数", "值": len(historical_ops)},
            {"项目": "第一轮粗网格组合数/每个历史状态", "值": len(coarse_grid)},
            {"项目": "DO候选水平数", "值": len(inclusive_grid(*bounds["Aerobic_DO"], DO_STEP))},
            {"项目": "HRT候选水平数", "值": len(inclusive_grid(*bounds["HRT"], HRT_STEP))},
            {"项目": "Feed候选水平数", "值": len(inclusive_grid(*bounds["Anaerobic_Feed_Fraction"], FEED_STEP))},
            {"项目": "统一粗网格理论总记录数", "值": len(common_history) * len(coarse_grid)},
            {"项目": "正向阶段局部细化", "值": "关闭；后续反向阶段直接筛选统一粗网格"},
            {"项目": "正向预测明细总行数", "值": global_acc.total_rows},
            {"项目": "规则粗扫描记录数", "值": global_acc.coarse_rows},
            {"项目": "当前实际工况额外补充记录数", "值": global_acc.current_supplement_rows},
            {"项目": "存在负预测候选数", "值": global_acc.negative_prediction_rows},
            {"项目": "存在非有限预测候选数", "值": global_acc.nonfinite_prediction_rows},
            {"项目": "预测物理/数值检查未通过候选数", "值": global_acc.invalid_prediction_rows},
            {"项目": "四项一级A参考同时满足记录数_全部正向", "值": global_acc.all_a1_meet_rows},
            {"项目": "四项一级A参考同时满足记录数_仅粗网格", "值": global_acc.all_a1_meet_coarse_rows},
            {"项目": "十一个Excel合计明细副表数", "值": len(detail_index)},
            {"项目": "公共Sheet保存位置", "值": "仅第1部分Excel；第2~11部分不重复"},
            {"项目": "每个结果文件目标明细Sheet数", "值": DETAIL_TARGET_SHEETS_PER_WORKBOOK},
        ]
        for part_no, history_part in enumerate(history_parts, start=1):
            bg_range = "无" if history_part.empty else f"{history_part['历史状态ID'].iloc[0]}~{history_part['历史状态ID'].iloc[-1]}"
            run_settings_rows.extend(
                [
                    {"项目": f"第{part_no}部分历史状态数", "值": len(history_part)},
                    {"项目": f"第{part_no}部分历史状态范围", "值": bg_range},
                    {"项目": f"第{part_no}部分理论粗网格记录数", "值": len(history_part) * len(coarse_grid)},
                ]
            )
        run_settings_rows.extend(
            [
                {"项目": "正向阶段是否按参考限值删除情景", "值": "否；只记录一级A工程参考判定标签"},
                {"项目": "是否运行NSGA-II", "值": "否"},
                {"项目": "主要RMSE参考规则", "值": "Time RMSE与Expanding RMSE中较大的有效值；只作误差量级提示"},
                {"项目": "后续反向筛选正式来源", "值": "仅使用网格来源=第一轮粗扫描"},
                {"项目": "一级A工程参考_TP_mg_L", "值": A1_REFERENCE_LIMITS["TP"]},
                {"项目": "一级A工程参考_TN_mg_L", "值": A1_REFERENCE_LIMITS["TN"]},
                {"项目": "一级A工程参考_NH4_N_mg_L", "值": A1_REFERENCE_LIMITS["NH4+-N"]},
                {"项目": "一级A工程参考_COD_mg_L", "值": A1_REFERENCE_LIMITS["COD"]},
                {"项目": "四项联合参考逻辑", "值": "TP<=0.5 AND TN<=15 AND NH4+-N<=5 AND COD<=50，且预测物理检查=通过"},
                {"项目": "参考限值解释边界", "值": "仅作为一步条件预测的工程筛选参考，不直接等同法规意义上的日均达标"},
                {"项目": "负预测/非有限预测处理", "值": "保留原始值并标记；单项一级A标签仅按该指标自身判定，异常指标自身记否；四项联合标签要求预测物理检查=通过；后续正式反向筛选排除异常记录"},
                {"项目": "预测类型", "值": "固定当前状态、改变当前运行设定的一步条件预测"},
            ]
        )
        run_settings = pd.DataFrame(run_settings_rows)

        warning_table = pd.DataFrame(warnings_rows, columns=["级别", "位置", "内容"])

        elapsed_minutes = (time.time() - start_time) / 60.0
        text = build_text_report(
            elapsed_minutes,
            common_history,
            historical_ops,
            bounds,
            coarse_grid,
            validation_metric_table,
            model_check,
            background_summary,
            detail_index,
            global_acc,
            warning_table,
        )

        # ------------------------------------------------------------------
        # Excel公共Sheet只写一次：统一放在第1部分。
        # 第2~11部分不重复说明/模型/历史状态/背景汇总，仅保留各自独立明细。
        # 明细本身已在流式预测阶段分别写入十一个writer。
        # ------------------------------------------------------------------
        file_split_index_rows: List[Dict[str, Any]] = []
        for part_no, history_part in enumerate(history_parts, start=1):
            bg_range = (
                "无"
                if history_part.empty
                else f"{history_part['历史状态ID'].iloc[0]}~{history_part['历史状态ID'].iloc[-1]}"
            )
            detail_part = detail_indexes[part_no - 1]
            file_split_index_rows.append(
                {
                    "结果部分": f"第{part_no}部分",
                    "结果文件": OUTPUT_EXCEL_PARTS[part_no - 1].name,
                    "历史状态数": len(history_part),
                    "历史状态范围": bg_range,
                    "理论粗网格记录数": len(history_part) * len(coarse_grid),
                    "明细副表数": len(detail_part),
                    "明细单Sheet目标上限行数": detail_rows_per_part[part_no - 1],
                    "公共Sheet": "是（集中保存）" if part_no == 1 else "否（不重复）",
                }
            )
        file_split_index = pd.DataFrame(file_split_index_rows)

        first_writer = excel_writers[0]
        first_writer.write_section_sheet(
            "01_说明与设置",
            [
                ("运行设置", run_settings),
                ("十一个结果文件拆分索引", file_split_index),
                ("参数范围与网格", grid_config),
                ("历史真实运行组合", historical_ops),
                ("十一个文件全部明细副表索引", detail_index),
                ("警告与说明", warning_table),
            ],
        )
        first_writer.write_section_sheet(
            "02_模型与误差",
            [
                ("模型核查", model_check),
                ("预测误差参考", validation_metric_table),
                ("共同样本核查", common_audit),
                ("插补重建核查", imputation_audit),
                ("NH4专项说明", nh4_note),
            ],
        )
        first_writer.write_table_sheet("03_历史状态", common_history)
        first_writer.write_section_sheet(
            "04_背景预测汇总",
            [
                ("全部历史背景预测汇总", background_summary),
                ("十一个部分合计四指标预测总体统计", prediction_global),
                ("十一个部分合计一级A工程参考判定汇总", a1_reference_summary),
                ("十一个部分合计预测物理检查汇总", physical_check_summary),
            ],
        )

        for writer in excel_writers:
            writer.close()

        excel_writers = []
        atomic_write_text(OUTPUT_TEXT, text)

        log("正向预测完成。")
        for part_no, path in enumerate(OUTPUT_EXCEL_PARTS, start=1):
            log(f"正向预测结果Excel-第{part_no}部分：{path}")
        log(f"十一个Excel合计明细副表：{len(detail_index)}个；公共Sheet仅保存在第1部分。")
        log(f"说明TXT：{OUTPUT_TEXT}")
        log(f"总明细行数：{global_acc.total_rows}")
        log(
            "第一轮粗扫描四项一级A参考同时满足："
            f"{global_acc.all_a1_meet_coarse_rows}/{global_acc.coarse_rows}"
        )
        log(f"总耗时：{elapsed_minutes:.2f} 分钟")

    except Exception as exc:
        for writer in excel_writers:
            try:
                writer.abort()
            except Exception:
                pass

        failure_lines = [
            "A²/O 四指标联合正向预测运行失败",
            "=" * 80,
            f"时间：{now_str()}",
            f"脚本版本：{SCRIPT_VERSION}",
            f"动态数据：{INPUT_FILE}",
            f"模型目录：{MODEL_DIR}",
            f"输出目录：{OUTPUT_DIR}",
        ]
        for part_no, path in enumerate(OUTPUT_EXCEL_PARTS, start=1):
            failure_lines.append(f"预期Excel-第{part_no}部分：{path}")
        failure_lines.extend(
            [
                "",
                f"错误类型：{type(exc).__name__}",
                f"错误信息：{exc}",
                "",
                "完整错误堆栈：",
                traceback.format_exc(),
                "",
                "说明：本脚本遇到错误后不会自动修改模型、数据或参数。请先根据错误信息核查输入文件和模型版本。",
            ]
        )
        failure_text = "\n".join(failure_lines)
        try:
            atomic_write_text(OUTPUT_TEXT, failure_text)
        except Exception:
            pass
        log(f"运行失败：{type(exc).__name__}: {exc}")
        raise


if __name__ == "__main__":
    main()
