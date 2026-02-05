# BenfordAnalysis

This repo runs a first-digit Benford analysis on `je_samples.xlsx`, generates a
markdown report, and outputs two SVG visuals.

## How to run locally

```bash
python benford_analysis.py
```

Outputs are written to the `output/` folder:

- `output/benford_report.md`
- `output/first_digit_observed_vs_expected.svg`
- `output/first_digit_deviation.svg`

## GitHub Actions

The **Benford Analysis** workflow runs the script on every push or manual
dispatch, then uploads the `output/` folder as the `benford-analysis-output`
artifact so you can download the report and visuals from the Actions tab.
