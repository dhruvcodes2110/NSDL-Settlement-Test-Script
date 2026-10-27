# NSDL Fixed-Width File Comparison Tool

A standalone Node.js utility to compare two fixed-width `.txt` files character by character. It does not depend on the generator and uses only Node.js built-in modules.

## Run

Keep `compare.js` in the project folder and open PowerShell there. The generated files are in the `output` subfolder. Replace the expected filename with the reference file you want to compare:

```powershell
node compare.js --expected "sysGen\0000003.txt" --actual "output\pledge-invocation-0000003.txt"
```

To choose a report folder:

```powershell
node compare.js --expected "manual-reference.txt" --actual "output\pledge-invocation-0000101.txt" --outdir reports
```

The tool creates:
- A `.txt` summary report
- A `.csv` detailed discrepancy report that opens in Excel

## What it checks

- Number of records in each file
- Record-by-record lengths
- Every character at every position, including spaces
- Missing records
- Differences in trailing characters / record lengths

Line endings (Windows CRLF vs Unix LF) are normalized. Meaningful spaces within each record are **not** trimmed.

## Important limitation

This version reports the exact 1-based character position of each mismatch but labels it `Unmapped field`. To report the NSDL field name (for example, ISIN, BO ID, quantity), field-position layouts need to be configured separately for each file type and specification. The comparator does not yet validate whether either value is semantically valid under the NSDL specification.

## Exit codes

- `0`: files match
- `1`: discrepancies found
- `2`: usage or file error
