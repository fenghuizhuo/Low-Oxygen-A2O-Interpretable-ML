# Change log – GitHub publication packaging

## v2: complete 14-script pipeline

- Added the exact author-provided final model-training and artifact-freezing script as `scripts/08_final_model_freezing.py`. It is inserted between the original final selection (07) and SHAP/PI (now 09).
- Renamed previous steps 08–13 to 09–14. Their Python executable bytes were unchanged.
- Fixed a **previous packaging-only mistake**: first script now defaults to `data/raw/数据总表.xlsx` instead of mistakenly passing `data/raw` as a file. No numerical preprocessing logic changed.
- Changed the new script's three original Windows drive path literals to repository-relative paths, adding `PROJECT_ROOT = Path(__file__).resolve().parents[1]`.
- Updated README, requirements comments, pipeline map, provenance manifest and release checklist to reflect all 14 steps.
- No raw Excel datasets, generated artifacts, fitted `.joblib` binaries or synthetic outputs are included.

## Code preservation and validation
- Original 13 source files: previously packaged executable programs preserved byte-for-byte (except the single first-step default-input path repair).
- New eighth source: reversing the three default path substitutions and added project-root definition reproduces the supplied original executable source segment exactly.
- All 14 scripts parse and compile. No execution on original data or re-creation of manuscript numerical outputs was performed.
- Final-model freezing output filenames/sheet names and model-package keys were checked against the next SHAP script.
- No original package versions or private datasets were supplied.

## Model identity retained from the supplied freezing script
- TP: F2-SVR, candidate SVR_032, C=10, epsilon=0.05, gamma=0.01, RBF.
- TN: F2-SVR, candidate SVR_018, C=1, epsilon=0.10, gamma=0.03, RBF.
- NH4+-N: F2-ElasticNet, candidate ElasticNet_025, alpha=0.01, l1_ratio=0.9.
- COD: F2-ElasticNet, candidate ElasticNet_040, alpha=0.30, l1_ratio=0.9.

## Updated default paths in new freezing script

| Old default path | Portable default |
|---|---|
| `D:\tools\pythonxiangmu\moxingduibi\jianmo\zuizhongmoxingqueding\F0-F3_统一终选与随机稳定性复核结果.xlsx` | `outputs/final_selection/F0-F3_统一终选与随机稳定性复核结果.xlsx` |
| `D:\tools\pythonxiangmu\moxingduibi\shujuyuchuli\queshizhichabu\动态可用数据与插补比较结果.xlsx` | `outputs/preprocessing/dynamic/动态可用数据与插补比较结果.xlsx` |
| `D:\tools\pythonxiangmu\moxingduibi\jianmo\zuizhongmoxingqueding\moxingbaocun` | `outputs/final_selection/moxingbaocun` |
