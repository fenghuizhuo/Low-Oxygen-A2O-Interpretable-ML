# Code inventory and provenance

Fourteen author-supplied programs were named in analysis order. The `source` columns show the uploaded source texts; sha256 hashes identify the exact input and packaged files. The original untouched source archive should remain private.

| Stage | Script filename | Original attachment | Function | Default paths made portable |
|---|---|---|---|---:|
| 01 | `01_outlier_preprocessing.py` | `粘贴的文本 (1).txt` | 数据清洗：固定边界与3×IQR | 2 |
| 02 | `02_dynamic_samples_imputation.py` | `粘贴的文本 (2).txt` | 四目标动态样本和插补方法比较 | 2 |
| 03 | `03_model_comparison_F0.py` | `粘贴的文本 (3).txt` | F0: 原始特征、六算法比较 | 2 |
| 04 | `04_model_comparison_F1.py` | `粘贴的文本 (1)(1).txt` | F1: 原始特征+机理衍生特征 | 2 |
| 05 | `05_model_comparison_F2.py` | `粘贴的文本 (2)(1).txt` | F2: 原始特征+近期出水 | 2 |
| 06 | `06_model_comparison_F3.py` | `粘贴的文本 (3)(1).txt` | F3: 全部特征 | 2 |
| 07 | `07_final_model_selection.py` | `粘贴的文本 (4).txt` | 统一终选+随机稳定性复核 | 3 |
| 08 | `08_final_model_freezing.py` | `粘贴的文本 (1)(2).txt` | 四目标最终F2模型全量拟合、joblib固化与Excel/TXT记录 | 3 |
| 09 | `09_shap_permutation_importance.py` | `粘贴的文本 (5).txt` | 最终模型SHAP+PI解释 | 3 |
| 10 | `10_pdp_ice_1d.py` | `粘贴的文本 (6).txt` | 一维PDP/ICE分析 | 4 |
| 11 | `11_pdp_2d.py` | `粘贴的文本 (7).txt` | 二维PDP响应分析 | 5 |
| 12 | `12_joint_forward_prediction.py` | `粘贴的文本 (8).txt` | 联合参数正向预测与11份明细输出 | 3 |
| 13 | `13_inverse_parameter_screening.py` | `粘贴的文本 (9).txt` | 四目标反向筛选和代表方案 | 2 |
| 14 | `14_low_do_supplementary_screening.py` | `粘贴的文本 (10).txt` | 进水TP分层低DO筛选 | 2 |

## File hashes

| Script | Original SHA-256 | Packaged SHA-256 |
|---|---|---|
| `01_outlier_preprocessing.py` | `a558d6fb83c73c3a999b8a9cf8fbf885545f2a01fe03054911030c89d680f0c2` | `690c68dc14076244a847ddb71925343a66e72f1167d23806aee940688618ce00` |
| `02_dynamic_samples_imputation.py` | `f4f5b4b05ced8dc6d9140c7337da7d5d49bd98fabbd51adc13366e561e0289b6` | `6116c6b7da0bd4709643118a69e9388a6a2782ac879e38a4c2bcfc3b3d547c02` |
| `03_model_comparison_F0.py` | `cf8ccadaa7a6c231e6ccfe576e8b601ac5c92643b9b42c8e85aaaae5fb3ae97d` | `f4923b8c8024e87d840ad0a29f761261eb6fab0e219fc5bb7f011fcfce9960cc` |
| `04_model_comparison_F1.py` | `18822dc65f451f71a71964ee5cd6b80f42744f5c922a86ee4625f494d6950f71` | `40d8ce91b3d46f383c5bd22e97355ee16153ab0ce9e2b9ffee06425bbf3839e8` |
| `05_model_comparison_F2.py` | `4c54ac7fa55f5812534aff40b321c82e2c2d714c6e5262c31094fd94266da2e7` | `0a711e577aee7fd0da4a893601e04b52179ba4d87b24b084f0ae99250afdc8d0` |
| `06_model_comparison_F3.py` | `50c085b1a4d1e8c75c1519a3d01c6016bc5b7a3091022f8fba4d3324c65db25c` | `93d0ee1bd9bddb90e245b5e7e20c0f2fb79ac7b50b7d4350b8e1f0ef845027f0` |
| `07_final_model_selection.py` | `dc1a3fe5d5bfef9498974e019f4127e8a41337d657231bcac033207abf75e71a` | `3f1906ab7f0cfd455ca4ba1e6142b7bbf38a20087ebccea4c7b1edaaa95799a8` |
| `08_final_model_freezing.py` | `2999f79c437f4fdde9db1a9a3eb4ed4eb731fda3214fe44221b3fb2d6c16a03f` | `2b6b2a9d1151627c052572b1f7e202765d11e9fcecd192c58c1702f5c76b33a5` |
| `09_shap_permutation_importance.py` | `e3881a72c771fba0a8d06a3bdccd215c61b56f99092c24271c72e0acf50fb59b` | `230602bfafdbb69b72f30aef62153e717ef63173974ba8022977e278e850cb7d` |
| `10_pdp_ice_1d.py` | `b25f7e8bd6057033c8060cf2206b90efc67c734f71d8fee6336e2c76f9ae7db3` | `244282a544aa6d8375c94ea5fe793bc4af6304c2ab1e012603536ecd819a2145` |
| `11_pdp_2d.py` | `90e3b24e4293d062881d5bc81654717fcd7366c62f812e57fbdfd965ed29c339` | `58339d4bd411392b6d370abe2c10d82cafa50b231a410c7fc7405e2bd846cb6e` |
| `12_joint_forward_prediction.py` | `ff88617f451401c935b6279787a8e804a55834d9878ede6cc2a7a46a3d42724e` | `8395b17f4852f5451840da23321191180d191ded28c18474b4173b9edb194263` |
| `13_inverse_parameter_screening.py` | `c16760d839ecd97a218bb6115987f285e8124f692924856e1735d5ff01c9a670` | `fd0bb0ed31e161ecca911de03e3a9cc56e0a123ce19b0f69a4c53fe92b0235a5` |
| `14_low_do_supplementary_screening.py` | `df290c75154214b9435b495998968f25ee8736a2d4b0db9ff12c88af9550d615` | `a3bebc7daa88b10e40fa2d6ffcbb00c3629a6efe517b26dd426bd244987e5fc9` |

## Additional packaging correction

The first archive accidentally set the default input of `01_outlier_preprocessing.py` to `data/raw/` rather than the intended workbook `data/raw/数据总表.xlsx`. This revision corrects that path only; the original input file and preprocessing rules were not changed.

The newly supplied full-model freezing script is now numbered 08, with downstream programs shifted to 09–14. This was a missing workflow *component*, not a new analysis method invented during packaging.
