# AGENTS.md

## Cursor Cloud specific instructions

### What this project is
A single-notebook data science project (`Heart_Attack_Possibility.ipynb`) that trains a
Logistic Regression model to predict heart-attack risk from the classic 303-row heart
disease dataset. There is no web app or long-running service — the "application" is the
notebook. Python dependencies are listed in `requirements.txt`.

### Dataset location (non-obvious)
The notebook was authored in Google Colab and hard-codes the dataset path
`/content/heart.csv`. The dataset is committed at the repo root as `heart.csv`, but the
notebook will NOT edit that path, so a copy must exist at `/content/heart.csv` before
running. Create it with:

```bash
sudo mkdir -p /content && sudo cp heart.csv /content/heart.csv
```

(`/content` is root-owned; passwordless `sudo` is available.) If `heart.csv` is ever
missing from the repo root, it is the Kaggle "Heart Attack Analysis & Prediction" /
UCI Cleveland dataset (303 rows, 14 columns) and can be re-fetched from
`https://raw.githubusercontent.com/kb22/Heart-Disease-Prediction/master/dataset.csv`.

### Running the notebook
`jupyter` and `nbconvert` are installed under `~/.local/bin`, which is not on `PATH` by
default — either add it or call the binaries directly. Run end-to-end headlessly with:

```bash
export PATH="$HOME/.local/bin:$PATH"
jupyter nbconvert --to notebook --execute --output /tmp/executed.ipynb Heart_Attack_Possibility.ipynb
```

Executing the notebook writes several PNG plots (e.g. `roc_curve.png`,
`feature_importance.png`, `confusion_matrix.png`) into the repo root; these are generated
artifacts and are gitignored. Expected results: accuracy ≈ 0.87 and ROC AUC ≈ 0.90.

### Interactive development
Launch Jupyter with `jupyter notebook` or `jupyter lab` (from `~/.local/bin`).

### Lint / test
There is no test suite or linter configured. "Testing" this project means executing the
notebook end-to-end (above) and confirming it runs without errors and reproduces the
expected metrics.
