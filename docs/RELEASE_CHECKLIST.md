# Checklist before making the repository public

- [x] 14 supplied Python programs staged, ordered, and renamed.
- [x] Original contents preserved except repository-relative default path replacements and project-root definitions; first-step workbook default corrected from earlier packaging.
- [x] Python AST and bytecode compilation checked on all 14 scripts.
- [x] Final model-freezing step now present between selection (07) and interpretation (09).
- [x] `data/` and `outputs/` ignored by Git; no private Excel or `.joblib` files packaged.
- [ ] Confirm source versions are exactly the ones used for the submitted manuscript.
- [ ] Run end-to-end on the authors' privately held source data; check study result tables and all dependent files.
- [ ] Confirm exact Python and third-party package versions and document a known-working environment.
- [ ] Review every script and repository history for source data, passwords, tokens or other private information.
- [ ] Check compatibility of workbook sheet names and filenames on authors' actual computer.
- [ ] Consider separately archiving figure-generation scripts if full figure reproducibility is claimed (not supplied).
- [ ] Confirm copyright and institutional approval; choose an appropriate license if broader software reuse is intended.
- [ ] Once all reviews pass, switch GitHub repository from Private to Public.
- [ ] Publish a versioned release or tag and link that exact version from the manuscript.

**No raw experimental data were available to independently execute the full 14-step pipeline. The checks above do not establish reproduced research results.**
