# AI Review Analyzer

A Python tool that reads customer reviews from an Excel spreadsheet, automatically sends them to ChatGPT for analysis, and writes the summarized results back into the same spreadsheet — no manual reading or summarizing required.

> **Goal:** Help researchers understand large volumes of customer feedback quickly, specifically to build a value proposition for a digital business card networking platform.

---

## Table of Contents
- [What You Need Before Running This](#what-you-need-before-running-this)
- [How to Set It Up](#how-to-set-it-up)
- [How to Run It](#how-to-run-it)
- [What Happens to Your Excel File](#what-happens-to-your-excel-file)
- [Important Notes & Warnings](#important-notes--warnings)
- [Troubleshooting](#troubleshooting)

---

## What You Need Before Running This

### 1. Python Installed
You need Python installed on your computer.
Download it from: [https://www.python.org/downloads/](https://www.python.org/downloads/)

### 2. Required Libraries
Open your terminal (or Command Prompt on Windows) and run:

```bash
pip install openai
pip install openpyxl
```

### 3. An OpenAI API Key
This tool connects to ChatGPT through OpenAI's service. You need an account and an API key from:
[https://platform.openai.com/account/api-keys](https://platform.openai.com/account/api-keys)

> ⚠️ **Note:** Using the API costs money based on usage. Check OpenAI's pricing page before running this on a large spreadsheet.

### 4. Your Excel File (.xlsx format)
Your spreadsheet must be set up like this:
- Reviews must be in **Column B**
- Reviews must start from **Row 4**
- The sheet tab name inside the file must match what is set in the script (see setup below)

---

## How to Set It Up

**Step 1** — Open the script file in any text editor (e.g. Notepad).

**Step 2** — Find this line near the top and replace with your actual OpenAI API key:
```python
op.api_key = "Enter_Key_Here"
```
Example:
```python
op.api_key = "sk-abc123yourkeyhere"
```

**Step 3** — Find this line and replace with the full path to your Excel file:
```python
file_path = "Enter_Filepath_Here"
```
Example (Windows):
```python
file_path = "C:/Users/YourName/Desktop/reviews.xlsx"
```
Example (Mac):
```python
file_path = "/Users/YourName/Desktop/reviews.xlsx"
```

**Step 4** — Find this line and replace `Sheet_Name` with the actual name of your Excel tab:
```python
sheet_r = workbook_r['Sheet_Name']
```
Example:
```python
sheet_r = workbook_r['Facebook Reviews']
```

---

## How to Run It

1. Open your terminal (Mac) or Command Prompt (Windows)
2. Navigate to the folder where the script is saved
3. Run the following command:

```bash
python your_script_name.py
```

4. Wait — the script will process every review. This may take several minutes depending on how many reviews you have
5. When you see **"Done!"** in the terminal, it is finished

---

## What Happens to Your Excel File

After the script runs, your Excel file will be updated with:

| Location | Content |
|----------|---------|
| **Column B** | Your original reviews (unchanged) |
| **Column C** | AI bullet-point summary for each individual review |
| **Cell D4** | Final overall analysis of all reviews combined |

---

## Important Notes & Warnings

> ⚠️ **Back up your Excel file** before running this script. The script saves directly over your original file.

> ⚠️ **Deprecated AI model:** This script uses `text-davinci-003`, which has been shut down by OpenAI. A developer will need to update it to use a newer model such as `gpt-3.5-turbo` or `gpt-4`.

> 💰 **Cost warning:** Every review sent to ChatGPT costs a small amount via the OpenAI API. Running this on hundreds or thousands of reviews can add up. Check your OpenAI usage dashboard regularly.

> ℹ️ The script processes reviews in **batches of 25** to keep the analysis manageable for the AI. This is normal and intentional.

---

## Troubleshooting

**`ModuleNotFoundError` when running the script**
> You forgot to install the libraries. Run: `pip install openai openpyxl`

**`Invalid API Key` error**
> Double-check that you copied your OpenAI API key correctly with no extra spaces.

**`KeyError: 'Sheet_Name'`**
> The sheet name in the script doesn't match the tab name in your Excel file. Open the Excel file, check the tab name at the bottom, and update the script to match exactly — it is case-sensitive.

**The script finishes but Column C is empty**
> Make sure your reviews start at Row 4 and are in Column B, and that the sheet name is correct.
