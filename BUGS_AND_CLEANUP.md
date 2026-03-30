# system1803 — Bug Report & Cleanup Guide

## Summary

| Category | Count |
|---|---|
| Critical bugs | 4 |
| Minor bugs | 4 |
| Files to delete | 8+ |
| Security issues | 1 |

---

## 🔴 Critical Bugs

### 1. Report formatting destroyed by Word placeholder replacement
**File:** `app.py` — `replace_placeholders_in_tables()` and every `p.text = p.text.replace(k, v)` call

**What happens:** When you do `cell.text = cell.text.replace(...)` or `p.text = ...` in python-docx,
it wipes ALL formatting runs inside that paragraph/cell — including bold, font size, font name, colour.
The text becomes plain unstyled text regardless of what the original Word template looked like.

**Fix:** Replace the entire Word report system with the new `report_generator_excel.py`
module provided. Excel reports are built from scratch in Python with explicit styling —
bold is bold, sizes are exact, colours are exact. No templates needed.

---

### 2. `elif test_type ==` chain was silently broken (Enter Results page)
**File:** `app.py`, Enter Results section

**What happens:** The code structure was:
```python
if test_type == "Bioburden":
    ...

bioburden_report_path = find_report_path(...)  # <-- this if/else broke the chain
if bioburden_report_path:
    ...download...

elif test_type == "Sterility":   # ← attached to `if bioburden_report_path`, NOT to Bioburden!
```
The `elif test_type == "Sterility":` was chained to `if bioburden_report_path:`, not to
`if test_type == "Bioburden":`. This worked by accident (the bioburden path was None for
non-bioburden tests) but would break if a bioburden report file happened to exist.

**Fix:** Each test type is now a clean, independent `if / elif` block. Download buttons
appear immediately after generation (no file path tracking needed).

---

### 3. Environmental report download unreachable
**File:** `app.py`, Enter Results → Environmental section

**What happens:** The `env_report_path` check and download buttons were placed **inside**
the `if sample_ids_range:` block. If the range was empty (mismatched IDs), a previously
generated report could never be downloaded.

**Fix:** Report generation and download are now in the same button click handler.

---

### 4. `end_col` column-letter calculation breaks with >26 columns
**File:** `app.py`, `update_rows_targeted()`

```python
end_col = chr(64 + len(COLUMNS))  # works only for ≤26 columns
```
If columns ever exceed 26, `chr(90+)` produces wrong characters (e.g. `[`, `\`, `]`).

**Fix:** Use `openpyxl.utils.get_column_letter(len(COLUMNS))` which correctly handles
any number of columns (AA, AB, etc.).

---

## 🟡 Minor Bugs

### 5. `import matplotlib.pyplot as plt` imported twice
**File:** `app.py`, Dashboard section

`import matplotlib.pyplot as plt` and `from matplotlib.ticker import MaxNLocator`
appeared inside the `if not df_filtered_2.empty:` block — they were also imported
inside Chart 1. Imports inside conditional blocks are bad practice (re-imports every render).

**Fix:** Both imports moved to the top of the file (done).

### 6. Bare `except:` clauses swallow all errors silently
**File:** `app.py`

- `_read_list_sheet()` used a bare `except:` — any connection error, API quota error,
  or bug returned an empty list with no feedback to the user.
- `generate_sample_id_range()` used a bare `except:` — silently returned `[]` on
  malformed Sample IDs.

**Fix:** Both now use `except Exception as e:` with proper `st.warning()` / `st.error()` output.

### 7. `add_custom_value()` checked all four list conditions with `if` instead of `elif`
**File:** `app.py`, `add_custom_value()`

If `list_name == "SampleTypes"`, the code would still evaluate the remaining three
`if` conditions unnecessarily.

**Fix:** Changed to `if / elif / elif / elif` chain.

### 8. Endotoxin report allowed generation with empty result field
**File:** `app.py`, Enter Results → Endotoxin

No validation that `endotoxin_result` was non-empty before generating the report,
resulting in a report with "Not specified" or blank in the result column.

**Fix:** Added `if not endotoxin_result: st.error(...)` guard before generation.

---

## 🔒 Security Issue

### Secrets committed to the repository
**Files:** `.streamlit/secrets.toml`, `.streamlit/secrets.toml.save`, `nano.14500.save`

Your **GCP private key** is committed in plain text in all three of these files.
Anyone with access to the repo can access your Google Sheet.

**Immediate actions:**
1. **Rotate the key** — Go to Google Cloud Console → IAM → Service Accounts →
   `system1803-bot@system1803.iam.gserviceaccount.com` → Keys → Delete the key with ID
   `de739aa00db0ef95d286dab14f5565ce0fe98666` → Create a new key → Download it.
2. **Add `.gitignore`** — The provided `.gitignore` file prevents secrets from being
   committed in future.
3. On Streamlit Cloud, set the secret via the Secrets dashboard (not in a file).

---

## 🗑️ Files to Delete

These files serve no purpose in a deployed application and some are harmful:

| File | Reason to Delete |
|---|---|
| `nano.14500.save` | Editor crash-recovery file. Contains embedded shell commands and partial secrets. |
| `check_report.py` | Debug/diagnostic script. Not used by the app. |
| `diagnose_sheet.py` | Debug/diagnostic script. Not used by the app. |
| `DATA_LOSS_ANALYSIS.md` | Internal debugging notes. Should not be in the repo. |
| `.streamlit/secrets.toml.save` | Duplicate backup of the secrets file — also contains the private key. |
| `file_backups/*.xlsx` | 15+ auto-generated backup files. They grow unboundedly and don't belong in git. Keep them on local disk only. |
| `__MACOSX/` (entire folder) | macOS metadata — invisible on Mac, noise for everyone else. |
| `microbiologymicrobiology_store.py` | Has a double-name typo in the filename and is never imported or used in `app.py`. If this is a future feature, rename it to `microbiology_store.py` and integrate it properly. |
| `BioburdenReport1803.docx` | Word templates no longer needed (replaced by Excel). |
| `SterilityReport.docx` | Same — no longer needed. |
| `EndotoxinReport.docx` | Same — no longer needed. |
| `EnvironmentReport.docx` | Same — no longer needed. |

### Files to KEEP
| File | Why |
|---|---|
| `app.py` | Main application (use the fixed version) |
| `report_generator_excel.py` | New Excel report generator (provided) |
| `requirements.txt` | Use the cleaned-up version provided |
| `.gitignore` | Use the version provided |
| `header.png` | Used in the Streamlit UI |
| `logo.jpeg` | Used in Excel reports |
| `packages.txt` | Streamlit Cloud system packages |
| `Database1803.xlsx` | Only keep the master — delete all `file_backups/*.xlsx` from git |
| `microbiology_inventory.xlsx` | Keep |
| `microbiology_transactions.xlsx` | Keep |
| `needed_materials.xlsx` | Keep |
| `custom_lists.xlsx` | Keep |
| `media_preparation.xlsx` | Keep |

---

## Why Excel is Better Than Word for Your Reports

| Issue | Word (python-docx) | Excel (openpyxl) |
|---|---|---|
| Bold/font preserved | ❌ `p.text =` destroys all formatting | ✅ Every cell styled explicitly in code |
| Font size control | ❌ Inherited from template, not reliable | ✅ Set per-cell: `Font(size=12, bold=True)` |
| Colours/borders | ❌ Template-dependent | ✅ `PatternFill`, `Border` — pixel-perfect |
| Logo embedding | ❌ Complex, fragile | ✅ `XLImage(path)` — simple |
| No template needed | ❌ Must maintain 4 .docx templates | ✅ Zero external files |
| PDF export | ❌ Requires LibreOffice on server | ✅ Excel → PDF via "Save As" on any computer |
| Arabic text | ❌ RTL in Word is unreliable via python-docx | ✅ Excel handles Arabic natively |
| Professional look | ❌ Depends on template design | ✅ Consistent, reproducible, always correct |

---

## Cleaned-up `requirements.txt`

```
streamlit
pandas
openpyxl
matplotlib
gspread
google-auth
google-auth-oauthlib
google-auth-httplib2
```

**Removed:**
- `python-docx` — no longer needed (Word reports replaced by Excel)
- `plotly` — imported in requirements but never used in `app.py`
- `oauth2client` — deprecated; `google-auth` is the modern replacement
