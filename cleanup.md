# Cleanup Guide - Fleet Electrification Analyzer v3.0.1

## Goal
Clean up codebase for GitHub upload while maintaining full functionality.

---

## Phase 1: Basic Organization (5 minutes)

### Step 1: Create folder structure
```bash
mkdir -p docs data/samples
```

### Step 2: Move documentation files
```bash
mv COMMERCIAL_VEHICLE_DEPLOYMENT_CHECKLIST.md docs/
mv commercial_vehicle_deploymentguide.md docs/
mv fuellyfix.md docs/
mv prompt docs/development_prompt.md
```

### Step 3: Move sample data files
```bash
mv sand_city_VINs*.csv data/samples/ 2>/dev/null || true
mv test_vins.csv data/samples/ 2>/dev/null || true
```

---

## Phase 2: Version Update (1 minute)

### Step 4: Update version number
- Open `settings.py`
- Find line: `APP_VERSION = "3.0.0"`
- Change to: `APP_VERSION = "3.0.1"`

---

## Phase 3: GitHub Preparation (2 minutes)

### Step 5: Create .gitignore
Create file `.gitignore` with:
# Python
__pycache__/
*.pyc
*.pyo
*.pyd
.Python

# OS
.DS_Store
Thumbs.db

# App specific
data/exports/
*.log


### Step 6: Test the app
```bash
python app.py --help
```

---

## Phase 4: Final Structure
After cleanup, your folder should look like:
VD/
├── app.py
├── settings.py
├── requirements.txt
├── README.md
├── .gitignore
├── data/
│ ├── models.py
│ ├── processor.py
│ ├── providers.py
│ ├── config.json
│ └── samples/
│ ├── sand_city_VINs.csv
│ └── test_vins.csv
├── analysis/
│ ├── calculations.py
│ ├── charts.py
│ └── reports.py
├── ui/
│ ├── main_window.py
│ ├── process_panel.py
│ ├── results_panel.py
│ └── analysis_panel.py
├── docs/
│ ├── COMMERCIAL_VEHICLE_DEPLOYMENT_CHECKLIST.md
│ ├── commercial_vehicle_deploymentguide.md
│ ├── fuellyfix.md
│ └── development_prompt.md
└── commercial_vehicle_scraper.py


---

## Rollback Instructions
If something goes wrong:
```bash
# Move files back to root
mv docs/* . 2>/dev/null || true
mv data/samples/* . 2>/dev/null || true
rmdir docs data/samples 2>/dev/null || true
```

---

## Validation
After cleanup, verify:
- [ ] App still runs: `python3 app.py --help`
- [ ] Version shows 3.0.1 in app
- [ ] All files are in correct locations
- [ ] No broken imports or missing files