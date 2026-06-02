# One-shot GitHub deployment

Open a terminal in this folder, then paste:

```bash
# 1. Create the repo on GitHub (manual — 30 sec, browser)
#    https://github.com/new
#    Name:        FPA-Commercial-Variance-Engine
#    Visibility:  Public
#    Do NOT add README, .gitignore, or licence (we ship our own)

# 2. From inside FPA_Commercial_Engine_UK/, run:
git init
git add .
git commit -m "feat: init FP&A operational forecast engine"
git branch -M main
git remote add origin https://github.com/<your-username>/FPA-Commercial-Variance-Engine.git
git push -u origin main
```

That's it — the README renders as the project landing page on GitHub.

## CV bullet (suggested)

> **FP&A Commercial Variance Engine** — github.com/&lt;you&gt;/FPA-Commercial-Variance-Engine
> Built an end-to-end FP&A engine for a UK multi-site leisure operator: FAST-standard rolling forecast,
> EBITDA waterfall and price-vs-volume bridge over a deliberately dirty 6.8k-row Xero export, plus an
> ARIMA / Holt-Winters demand model in R and Python (MAPE 10.95% on a 28-day holdout, beating
> seasonal-naive by 19%).
