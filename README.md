# Quince Logistics Calculator

A simple web tool to compute **Logistics Cost Ocean** and **Logistics Cost Air** for any list of products in an Excel file.

**Live URL:** https://akshayg9511.github.io/quince-logistics-calculator/

---

## How to use

1. Open the URL above.
2. Drop your Excel file into the page (or click "browse").
3. Wait a second or two.
4. Click **Download Excel**. The file you downloaded is your original file with two new columns appended at the right:
   - `Logistics Cost Ocean ($/unit)`
   - `Logistics Cost Air ($/unit)`
5. Open the downloaded file, select those two columns, and paste them into the matching rows of your main Bid Inputs sheet (columns AT and AU).

Row order is preserved exactly. Whatever order your input rows are in, the output is in the same order.

---

## What your Excel file must contain

The file must have these **five column headers** spelled exactly (case doesn't matter, but spelling and spacing do):

| Header | What goes in it |
|---|---|
| `Length (in)` | Length in inches (number) |
| `Width (in)`  | Width in inches (number) |
| `Height (in)` | Height in inches (number) |
| `Weight (g)`  | Weight in grams (number) |
| `COO` | Country of origin |

Any other columns you have (Style #, Size, FOB, etc.) are kept in the output untouched.

### COO column accepts three formats

- ISO code: `IN`, `CN`, `VN`
- Full name: `India`, `China`, `Viet Nam`
- Combined: `India | IN`, `China | CN`

If a row has an unknown country, the tool tells you which row and what value, and won't produce a file until you fix it.

---

## When it won't accept your file

If something is wrong, the page shows a clear red message saying exactly what to fix. Common cases:

- One of the 5 required columns is missing → it tells you which one.
- A row has a blank or non-numeric Length/Width/Height/Weight → it tells you which row.
- A row has a COO it doesn't recognize → it tells you which row and what value.

Fix it in your Excel and drop the file again.

---

## Notes

- The tool runs entirely in your browser. Nothing is uploaded to a server.
- Works on Chrome, Safari, Firefox.
- For 100 rows it finishes in well under a second. For 1,000 rows in about a second.

If something looks off in the output, message **@Akshay** with the file and the row in question.
