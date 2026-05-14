# GitHub Push/Pull Workflow

Current Conrod report tool folder:

```text
C:\Users\TechnoStar\Python\Conrod\Post_Processing\Report\
├── __init__.py
├── output.py
├── report.py
├── template.py
├── pixel_report_diff.py
├── requirements.txt
├── README_pixel_report_diff.md
├── GITHUB_PUSH_WORKFLOW.md
└── runtime.txt
```

Do not push generated output files unless you intentionally want to keep test evidence:

```text
page_001_overlay.png
page_001_mask.png
summary.json
object_summary.json
__pycache__/
```

## 1. Push From Local Machine

Run from the Conrod repository root. If this folder is inside the Git repo, the expected root is usually:

```powershell
cd "C:\Users\TechnoStar\Python\Conrod"
```

Check changed files:

```powershell
git status --short
```

Stage the report tool files:

```powershell
git add Post_Processing/Report/pixel_report_diff.py
git add Post_Processing/Report/requirements.txt
git add Post_Processing/Report/README_pixel_report_diff.md
git add Post_Processing/Report/GITHUB_PUSH_WORKFLOW.md
git add Post_Processing/Report/runtime.txt
```

If you also want to push existing module files:

```powershell
git add Post_Processing/Report/__init__.py
git add Post_Processing/Report/output.py
git add Post_Processing/Report/report.py
git add Post_Processing/Report/template.py
```

Commit:

```powershell
git commit -m "Add PPT report comparison tool"
```

Push:

```powershell
git push
```

If the branch has not been pushed before:

```powershell
git push -u origin HEAD
```

## 2. Pull On Remote Machine

If the repository is not cloned yet:

```powershell
git clone <github-repo-url>
cd "<repo-folder>"
```

If the repository already exists:

```powershell
cd "<repo-folder>"
git pull
```

Go to the current tool folder:

```powershell
cd "Post_Processing\Report"
```

## 3. Prepare Python 3.9 Environment

Confirm Python version:

```powershell
python --version
```

Expected:

```text
Python 3.9.x
```

Install dependencies:

```powershell
python -m pip install -r requirements.txt
```

The machine also needs Microsoft PowerPoint installed because slide image export uses `pywin32`.

## 4. Run Report Comparison

Recommended command:

```powershell
python pixel_report_diff.py expected.pptx actual.pptx --mode both --threshold 3
```

Default output folder:

```text
C:/Users/TechnoStar/Python/Conrod/Post_Processing/Report
```

To choose another output folder:

```powershell
python pixel_report_diff.py expected.pptx actual.pptx -o diff_output --mode both --threshold 3
```

## 5. Read Results

Main outputs:

```text
page_001_overlay.png
page_001_mask.png
summary.json
object_summary.json
```

Exit codes:

```text
0 = pass
1 = differences found
2 = runtime/input error
```
