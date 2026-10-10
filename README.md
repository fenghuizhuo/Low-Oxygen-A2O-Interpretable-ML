# Low-oxygen A2/O process: interpretable machine learning

## Associated manuscript

**Prediction of Multiple Effluent Indicators and Operational Parameter Screening for an A2/O Process under Low Oxygen Conditions Using Interpretable Machine Learning**

This repository contains **14 Python scripts** corresponding to the manuscript's analytical workflow: anomaly handling, construction of dynamic prediction samples and offline missing-value method comparison, six-algorithm comparisons across F0–F3, unified final-model selection, full-history final-model freezing, model interpretation (SHAP and permutation importance), one- and two-dimensional PDP/ICE, and model-based forward prediction and operational-parameter screening.

All 14 programs are the authors' supplied scripts, converted from `.txt` to `.py`, given stage-ordered English filenames, and adapted to repository-relative **default file locations**. Numerical methods and model configuration constants were not deliberately changed. See `CHANGELOG.md` for one corrected packaging-path mistake in the earlier archive.

**Scope:** Model-generated response and screening scenarios should not be interpreted as experimentally demonstrated operating interventions. Consult the manuscript for analytical assumptions, limitations and interpretation.

## Data availability and reproducibility status

The underlying experimental wastewater-treatment data are **not publicly deposited here**. They may be requested from the corresponding authors upon reasonable request, subject to the authors' data-sharing arrangements.

The `data/` and `outputs/` directories are excluded from this public code repository. Consequently, these scripts cannot produce the study's numeric outputs **without the requisite nonpublic data**. No dataset or pretrained `.joblib` model package is included in this archive. We performed syntax, code-preservation, default-path, and output-dependency consistency checks during packaging, **not** an independent full-data reproduction of manuscript results.

The exact working Python/package versions were not provided; `requirements.txt` lists the top-level Python packages inferred from script imports, not a verified fully pinned environment. Plotting scripts that generated the final manuscript figures were not among the materials provided.

## Repository layout

- `scripts/`: 14 Python analysis scripts in intended sequence; see `CODE_INVENTORY.md`.
- `docs/INPUT_OUTPUT_MAP.md`: stage-by-stage paths and prerequisite files.
- `docs/RELEASE_CHECKLIST.md`: checks to complete before making the repository public.
- `docs/packaging_manifest.json`: per-script provenance, original and packaged SHA-256 and default-path edits.
- `CHANGELOG.md`: archive preparation changes, including verification scope.
- `requirements.txt`: unpinned observed package dependencies.
- `.gitignore`: excludes raw/private data, computed workbooks, joblib artifacts, checkpoints and local configuration.

Scripts resolve their root folder as `Path(__file__).resolve().parents[1]`. Run from the **unpacked project directory** while retaining the `scripts/` folder structure. With access to the necessary private workbooks, default input/output paths are portable across computers.

## Run order (after obtaining the necessary data)

1. Install the packages in a suitable Python environment:

   ```bash
   python -m pip install -r requirements.txt
   ```

2. Put the nonpublic input workbook in your **local working copy** at `data/raw/数据总表.xlsx`. This `data/` folder is ignored by Git. Do not commit it to GitHub.

3. From the repository root, run each stage in this order:

   ```bash
   python scripts/01_outlier_preprocessing.py
   python scripts/02_dynamic_samples_imputation.py
   python scripts/03_model_comparison_F0.py
   python scripts/04_model_comparison_F1.py
   python scripts/05_model_comparison_F2.py
   python scripts/06_model_comparison_F3.py
   python scripts/07_final_model_selection.py
   python scripts/08_final_model_freezing.py
   python scripts/09_shap_permutation_importance.py
   python scripts/10_pdp_ice_1d.py
   python scripts/11_pdp_2d.py
   python scripts/12_joint_forward_prediction.py
   python scripts/13_inverse_parameter_screening.py
   python scripts/14_low_do_supplementary_screening.py
   ```

**Sequence detail:** Stages 03–06 produce four sets of comparison workbooks. Stage 07 reads these workbooks and reports the final selection. **Stage 08 (provided separately after the first archive was prepared)** uses the final selection workbook and dynamic data to fit/freeze the four already-selected F2 models (`TP: SVR_032`, `TN: SVR_018`, `NH4+-N: ElasticNet_025`, `COD: ElasticNet_040`). It creates four `.joblib` packages plus `四指标_最终模型固化汇总.xlsx`. These outputs are required by stage 09 (SHAP/PI) and later PDP/screening programs. Stage 12 generates 11 forward-prediction Excel files which stage 13 reads; stage 14 uses stage 13's stratified screening results.

These computations can be intensive: bootstrap/model validation, permutation-based explanation and millions of predicted scenarios can require significant CPU time, RAM and disk space. `MODEL_SMOKE_TEST` and other original optional environment variables were not changed during packaging.

Input workbooks rely on specific original Chinese sheet and column names; no raw data or dummy outputs were created or added. Where a script supports input/output path overrides by command line or environment variables, that original interface is retained. For actual stage filenames and outputs, refer to `docs/INPUT_OUTPUT_MAP.md`.

## Citing and reusing the code

Please cite the associated manuscript when referring to these programs. No software license was provided with the original submission. A public GitHub repository can be viewed, but it does not grant an explicit right to reuse the code absent a selected license. The authors should confirm institutional requirements and copyright permissions before choosing a license.
