# warden-report-gen
Excel Report generator for instance tables for reporting based on Horangi warden output

## Usage
```console
foo@bar ~ $ python3 reportparse.py --excel_file <filename>.xlsx --raw <raw check spreadsheet name> --grouped <grouped checks spreadsheet name>
```

Any rules that are determined to not be issues, should be set as "Not an Issue" (case insensitive)

The grouped checks sheet should include the following headers:
* Issue Title
* Affected Module: This field corresponds to the Rule Title in the raw sheet
* Severity
* Score

## TBD
* Output excel formatting
* Cleanup to remove unused code
