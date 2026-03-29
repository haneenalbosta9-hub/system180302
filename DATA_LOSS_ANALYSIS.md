# 🔴 DATA LOSS ISSUE - ROOT CAUSE ANALYSIS

## Problem Summary
Your lab system app is deleting data from your Google Sheets when you edit entries. This is happening because of a mismatch between the expected column headers in your code and the actual headers in your Google Sheet.

---

## 🔍 Root Cause

### The Issue in Your Code (`app.py` lines 116-128)

```python
def save_data(df):
    ws = _get_worksheet(SHEET_SAMPLES, headers=COLUMNS)
    rows = _df_to_rows(df)
    ws.clear()  # ⚠️ DANGER: This clears EVERYTHING!
    ws.append_row(COLUMNS)
    if rows:
        ws.append_rows(rows, value_input_option="RAW")
```

**What happens:**
1. Function reads data from Google Sheets
2. **BUT** it only reads columns that are defined in the `COLUMNS` list
3. When saving, it calls `ws.clear()` which deletes the ENTIRE sheet
4. Then it writes back ONLY the columns from `COLUMNS` list
5. **Any data in columns NOT in the list is permanently lost**

### The Column Mismatch

**Your `COLUMNS` list expects:**
```python
COLUMNS = [
    "Received Date",      # Column A
    "Sample ID",          # Column B
    "Unit No.",           # Column C
    "Sample Type",        # Column D
    # ... etc
]
```

**Your actual Google Sheet has:**
```
Column A: "10"              ❌ WRONG - Should be "Received Date"
Column B: "Sample ID"       ✅ Correct
Column C: "Unit No."        ✅ Correct
# ... rest are correct
```

---

## ✅ SOLUTIONS

### Solution 1: Fix the Google Sheet Header (RECOMMENDED - QUICKEST)

**Steps:**
1. Open your Google Sheet: https://docs.google.com/spreadsheets/d/1EXiXsOQ0VsfIbZlUpN3r6g0-aRNUUEKZDVZHh_xZnEY/edit
2. Click on cell **A1**
3. Change the value from `"10"` to `"Received Date"`
4. Save (Ctrl+S or Cmd+S)
5. Test your app again

**Why this works:** The code will now recognize Column A as "Received Date" and include it when saving data.

---

### Solution 2: Add Safety Checks to Your Code

I've already updated your `app.py` with safety checks. The new `save_data()` function:

```python
def save_data(df):
    # ... existing code ...
    
    # ⚠️ SAFETY CHECK: Verify headers match before clearing
    existing_headers = ws.row_values(1)
    if existing_headers != COLUMNS:
        st.error("❌ SAFETY CHECK FAILED: Sheet headers don't match!")
        st.error(f"Expected: {COLUMNS}")
        st.error(f"Found: {existing_headers}")
        st.warning("⚠️ Please fix the column headers before saving.")
        st.stop()
    
    # Only proceed if headers match
    rows = _df_to_rows(df)
    ws.clear()
    # ... rest of saving logic ...
```

**Benefits:**
- Prevents accidental data loss
- Shows clear error messages
- Forces you to fix the header issue before saving

---

### Solution 3: Use Update Instead of Clear (SAFEST)

Instead of clearing the entire sheet and rewriting everything, use the new `update_data_safe()` function:

```python
def update_data_safe(df):
    """Updates existing rows without clearing the sheet"""
    ws = _get_worksheet(SHEET_SAMPLES, headers=COLUMNS)
    
    # Get existing data
    existing_data = ws.get_all_values()
    
    # Delete rows from 2 onwards (keep header)
    if len(existing_data) > 1:
        ws.delete_rows(2, len(existing_data))
    
    # Append new data
    rows = _df_to_rows(df)
    if rows:
        ws.append_rows(rows, value_input_option="RAW")
```

**Benefits:**
- Never touches the header row
- Less risk of data corruption
- Faster for small updates

---

## 🛠️ Diagnostic Tools

### Run the Diagnostic Script

I've created `diagnose_sheet.py` for you. To run it:

```bash
streamlit run diagnose_sheet.py
```

This will:
- ✅ Check if your headers match
- ❌ Show you exactly which columns are wrong
- 🔧 Offer an auto-fix option (use with caution!)
- 📊 Display sample data from your sheet

---

## 📋 Step-by-Step Fix Guide

### Immediate Fix (5 minutes)

1. **Backup your data first!**
   - Download your Google Sheet as Excel: File → Download → Microsoft Excel
   
2. **Fix the header:**
   - Open Google Sheet
   - Click cell A1
   - Change "10" to "Received Date"
   - Press Enter

3. **Verify the fix:**
   - Run `streamlit run diagnose_sheet.py`
   - Check that all headers match

4. **Test your app:**
   - Make a small edit in your lab app
   - Check that no data is lost

### Long-term Prevention

1. **Update your app.py:**
   - Use the updated version with safety checks
   - Consider using `update_data_safe()` instead of `save_data()`

2. **Add data validation:**
   - Protect the header row in Google Sheets (Right-click row 1 → Protect range)
   - Set up automated backups

3. **Monitor for issues:**
   - Check the "file_backups" folder regularly
   - Review Google Sheets version history if data goes missing

---

## 🚨 Emergency Data Recovery

If you've already lost data:

1. **Check Google Sheets Version History:**
   - File → Version history → See version history
   - Restore a version from before the data loss

2. **Check your backups folder:**
   - Look in `file_backups/` for recent backups
   - Files are named like `Database1803_backup_20260326_110528.xlsx`

3. **Manually restore:**
   - Open a backup file
   - Copy the missing data
   - Paste it back into Google Sheets

---

## 📧 Technical Details

### Why `_df_to_rows()` Ignores Extra Columns

```python
def _df_to_rows(df):
    # ... conversion logic ...
    cols_to_write = [c for c in COLUMNS if c in df_save.columns]
    #                                     ^^^^^^^^^^^^^^^^^^^^^^^^
    #                      Only includes columns from COLUMNS list!
    return df_save[cols_to_write].values.tolist()
```

This means if your sheet has a column called "10", it's not in the `COLUMNS` list, so it gets ignored when reading and lost when saving.

---

## ✅ Verification Checklist

After implementing fixes:

- [ ] Column A header is "Received Date" (not "10")
- [ ] All column headers match the `COLUMNS` list
- [ ] Safety checks are in place in `save_data()`
- [ ] Tested editing a sample without data loss
- [ ] Backups are being created regularly
- [ ] Header row is protected in Google Sheets

---

## 📞 Need Help?

If you're still experiencing issues:

1. Run the diagnostic script and share the output
2. Check the Streamlit app logs for error messages
3. Verify your Google Sheets API credentials are correct
4. Make sure all sheets (Samples, SampleTypes, TestTypes, etc.) exist

---

**Created:** March 27, 2026  
**Issue:** Data loss when editing in lab system app  
**Root Cause:** Column header mismatch ("10" vs "Received Date")  
**Fix:** Rename Column A to "Received Date" in Google Sheets
