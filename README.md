# xlsx-fixer

**Stop delivering "corrupted" Excel files to your clients and executives.**
A zero-dependency delivery-assurance layer for Python data pipelines.

If you generate `.xlsx` files with Python's `openpyxl` library and open them on a Mac, your stakeholders will see this terrifying warning:

> **"We found a problem with some content in 'file.xlsx'. Do you want us to try to recover as much as we can?"**

For a solo developer, this is annoying. For a data engineering team delivering automated reporting to C-suite executives or high-value clients, **this completely destroys the professional credibility of the automation pipeline.**

## The Root Cause
`openpyxl` hardcodes inline strings (`t="inlineStr"`) for every text cell. Mac Excel's strict OOXML parser rejects this. The openpyxl maintainer has [declined to fix this](https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1804) for 17+ years.

`xlsx-fixer` intercepts the generated file, safely rebuilds the XML to use a proper shared string table (`xl/sharedStrings.xml`), drops inconsistent calculation states, and strips illegal control characters. 

---

## The Open Source Tier (Free)
This repository contains the core `fix()` engine. It is perfect for ad-hoc scripts and individual files. **Zero dependencies** beyond the standard library.

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

If you are running Airflow, GitHub Actions, or processing hundreds of files, the Open Source version is too slow and requires too much manual code modification. 

**xlsx-fixer Pro** is a pre-compiled, un-hackable binary (Linux/Mac/Windows) designed for scale.

### Pro Features:
*   **The Drop-In Patch:** Don't rewrite 100 legacy scripts. Just call `xlsx_fixer_pro.patch_openpyxl()` once, and every `wb.save()` call across your entire application is automatically fixed.
*   **High-Performance Batch Processing:** A multithreaded CLI command (`batch`) that rips through directories of hundreds of `.xlsx` files using all available CPU cores.
*   **Big Data Safe Mode:** Disk-backed XML streaming prevents Out-of-Memory (OOM) crashes on 500MB+ Excel dumps.
*   **CI Audit Reporting:** Generates machine-readable `JSON` reports proving data integrity for your compliance logs.

### Commercial Licenses

| License | Target Audience | Features | Link |
| :--- | :--- | :--- | :--- |
| **Pipeline License** | Solo Automators / Agencies | Unlocks the `batch` command, drop-in patches, and CI execution for a single pipeline. | [Buy Pipeline License ($299/yr)](https://revasser-llc.lemonsqueezy.com/checkout/buy/87b672ef-1143-4bd5-b593-b32c9df28f9f) |
| **Site License** | Enterprise Data Teams | Uncapped deployments across all cloud runners, dev machines, and internal tools. | [Buy Site License ($899/yr)](https://revasser-llc.lemonsqueezy.com/checkout/buy/8c7977ac-2f87-455e-a979-56b6033ce947) |

*Upon purchase, you will immediately receive your License Key and a secure download link for the compiled CI binaries.*

---

## What It Fixes Under The Hood

| # | Issue | Impact |
|---|-------|--------|
| 1 | **Inline strings** | "We found a problem" dialog on every Mac open |
| 2 | **fullCalcOnLoad** | Recovery dialog; formulas show 0 |
| 3 | **Stale calcChain.xml** | "Removed Records: Formula" in repair log |
| 4 | **Control characters** | Illegal XML 1.0 chars violate OOXML ST_Xstring |

## License
The free `xlsx-fixer` engine is licensed under the MIT License. See [LICENSE](LICENSE) for details.
