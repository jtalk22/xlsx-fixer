<div align="center">
  <h1><code>xlsx-fixer</code></h1>
  <p><strong>A zero-dependency delivery-assurance layer for Python data pipelines.</strong></p>
  <p>Stop delivering "corrupted" Excel files to your clients and executives.</p>
  
  <a href="https://pypi.org/project/xlsx-fixer/"><img src="https://img.shields.io/pypi/v/xlsx-fixer.svg?color=blue" alt="PyPI Version"></a>
  <a href="https://pypi.org/project/xlsx-fixer/"><img src="https://img.shields.io/pypi/pyversions/xlsx-fixer.svg" alt="Python Versions"></a>
  <a href="LICENSE"><img src="https://img.shields.io/badge/License-MIT-blue.svg" alt="License"></a>
</div>

<br/>

## 🚨 The 17-Year-Old Bug

If you generate `.xlsx` files with Python's `openpyxl` library and open them on a Mac, your stakeholders will see this terrifying warning:

> ⚠️ **"We found a problem with some content in 'report.xlsx'. Do you want us to try to recover as much as we can?"**

For a solo developer, this is annoying. For a data engineering team delivering automated reporting to C-suite executives or high-value clients, **this completely destroys the professional credibility of the automation pipeline.**

### Why does this happen?
`openpyxl` hardcodes inline strings (`t="inlineStr"`) for every text cell. Mac Excel's strict OOXML parser rejects this. The openpyxl maintainer has [declined to fix this](https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1804) for 17+ years, calling it a permanent design decision.

### How we fix it
`xlsx-fixer` intercepts the generated file, safely rebuilds the XML to use a proper shared string table (`xl/sharedStrings.xml`), drops inconsistent calculation states, and strips illegal control characters. One function call. Zero dependencies.

---

## 🛠️ The Open Source Tier (Free)
This repository contains the core `fix()` engine. It is perfect for ad-hoc scripts and individual files.

```bash
pip install xlsx-fixer
```

### Usage (Python)

```python
from openpyxl import Workbook
from xlsx_fixer import fix

wb = Workbook()
ws = wb.active
ws["A1"] = "Hello, Mac Excel!"

# 1. Save normally
wb.save("report.xlsx")

# 2. Fix it before sending to the client
fix("report.xlsx") 
```

### Usage (CLI)

```bash
# Fix a single file in-place
xlsx-fixer fix report.xlsx

# Check a file for corruption patterns without modifying
xlsx-fixer check report.xlsx
```

---

## 🚀 The Enterprise Tier (xlsx-fixer Pro)
For data engineering teams and CI/CD pipelines, we offer **xlsx-fixer Pro**. 

If you are running Airflow, GitHub Actions, or processing hundreds of files, the Open Source version is too slow and requires too much manual code modification. **xlsx-fixer Pro** is a pre-compiled, un-hackable binary (Linux/Mac/Windows) designed for enterprise scale.

| Feature | `xlsx-fixer` (Free) | `xlsx-fixer Pro` |
| :--- | :---: | :---: |
| Single File Fix (`fix()`) | ✅ | ✅ |
| **Drop-in Patch** (`patch_openpyxl()`) | ❌ | ✅ |
| **Multithreaded Batch CLI** (`batch`) | ❌ | ✅ |
| **Big Data Safe Mode** (OOM prevention) | ❌ | ✅ |
| **CI Audit Reporting** (`--report json`) | ❌ | ✅ |
| Deployment Format | Source Code | Compiled Native Binary |

### ⚡ The Drop-In Patch
Don't rewrite 100 legacy scripts. Just call `patch_openpyxl()` once at the top of your application, and *every* `wb.save()` call is automatically fixed in-place before it hits the disk.

### 💼 Commercial Licenses

| License | Target Audience | Link |
| :--- | :--- | :--- |
| **Pipeline License** ($299/yr) | Solo Automators / Boutique Agencies | [Purchase on Lemon Squeezy](https://revasser-llc.lemonsqueezy.com/checkout/buy/87b672ef-1143-4bd5-b593-b32c9df28f9f) |
| **Site License** ($899/yr) | Enterprise Data Teams (Uncapped) | [Purchase on Lemon Squeezy](https://revasser-llc.lemonsqueezy.com/checkout/buy/8c7977ac-2f87-455e-a979-56b6033ce947) |

*Upon purchase, you will immediately receive your License Key and a secure download link for the compiled CI binaries.*

---

## 🔬 What It Fixes Under The Hood

| # | Issue | Root Cause | Impact |
|---|-------|-----------|--------|
| 1 | **Inline strings** | `openpyxl` hardcodes `t="inlineStr"` | "We found a problem" dialog on every Mac open |
| 2 | **fullCalcOnLoad** | `openpyxl` sets `fullCalcOnLoad="1"` without generating `calcChain.xml` | Recovery dialog; formulas show 0 |
| 3 | **Stale calcChain.xml** | References cells that no longer contain formulas | "Removed Records: Formula" in repair log |
| 4 | **Control characters** | Illegal XML 1.0 chars (U+0000-U+0008, etc.) in cell values | ST_Xstring validation failures |

## License
The free `xlsx-fixer` engine is licensed under the MIT License. See [LICENSE](LICENSE) for details.
