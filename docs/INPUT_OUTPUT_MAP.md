# Input and output workflow

All paths are relative to the repository root. Private `data/` and generated `outputs/` are excluded from public version control by `.gitignore`.

| Step | Script | Main input | Main output directory and dependency |
|---|---|---|---|
| 01 | `01_outlier_preprocessing.py` | `data/raw/数据总表.xlsx` | `outputs/preprocessing/outlier/` → `异常值处理结果.xlsx` |
| 02 | `02_dynamic_samples_imputation.py` | `outputs/preprocessing/outlier/异常值处理结果.xlsx` | `outputs/preprocessing/dynamic/` → `动态可用数据与插补比较结果.xlsx` |
| 03 | `03_model_comparison_F0.py` | Dynamic workbook from 02 | `outputs/model_comparison/F0/` → four target-specific comparison workbooks |
| 04 | `04_model_comparison_F1.py` | Dynamic workbook from 02 | `outputs/model_comparison/F1/` → four target-specific comparison workbooks |
| 05 | `05_model_comparison_F2.py` | Dynamic workbook from 02 | `outputs/model_comparison/F2/` → four target-specific comparison workbooks |
| 06 | `06_model_comparison_F3.py` | Dynamic workbook from 02 | `outputs/model_comparison/F3/` → four target-specific comparison workbooks |
| 07 | `07_final_model_selection.py` | F0–F3 comparison workbooks and dynamic workbook | `outputs/final_selection/` → `F0-F3_统一终选与随机稳定性复核结果.xlsx` |
| **08** | **`08_final_model_freezing.py`** | **Output from 07 plus dynamic workbook from 02** | **`outputs/final_selection/moxingbaocun/` → four `.joblib`, `四指标_最终模型固化汇总.xlsx`, and TXT report** |
| 09 | `09_shap_permutation_importance.py` | Four frozen `.joblib`, freezing summary and dynamic workbook | `outputs/interpretation/shap_pi/` → `四指标_SHAP与PermutationImportance综合结果.xlsx` |
| 10 | `10_pdp_ice_1d.py` | Frozen models, dynamic workbook, SHAP/PI output | `outputs/response/one_dimensional/` → `四指标_一维PDP_ICE响应分析结果.xlsx` |
| 11 | `11_pdp_2d.py` | Frozen models, dynamic workbook, SHAP/PI and 1D output | `outputs/response/two_dimensional/` → `四指标_二维PDP响应分析结果.xlsx` |
| 12 | `12_joint_forward_prediction.py` | Frozen models and dynamic workbook | `outputs/screening/forward/` → 11 `四指标_DO-HRT-进水分配联合正向预测_第N部分.xlsx` files |
| 13 | `13_inverse_parameter_screening.py` | 11 Excel outputs from 12 | `outputs/screening/inverse/` → `四指标_目标出水反向优化结果.xlsx` |
| 14 | `14_low_do_supplementary_screening.py` | Inverse-screening workbook from 13 | `outputs/screening/inverse/low_do/` → `四指标_DO不高于2mg_L补充筛选结果.xlsx` |

## Static interface compatibility checks

- Step 07's exact final-selection workbook filename matches step 08's input default.
- Step 08's `.joblib` filenames, model-package dictionary fields, and `模型总览` Excel sheet match step 09's required inputs.
- Step 08 writes to the default `moxingbaocun/` directory expected by stages 09–12.
- The F2 model configurations locked by step 08 are those described in the manuscript and expected by later interpretation scripts.
- Step 12 output naming is intended to match stage 13's 11 input workbooks; stage 14 consumes the result of stage 13.

**Important limits:** These are static source-level checks; model fitting, output generation and numeric reproduction were not run because original experimental data were not supplied. If final plot-generation programs are needed for figure-level reproduction, those scripts were not provided.
