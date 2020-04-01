# Combine XLS Name Columns

Combines names in separate columns (John, Jane, Doe) and appends a new column to the spreadsheet containing "John and Jane Doe".

## Usage

1. Edit the input file and header columns to match your spreadsheet, and the output file to whatever you want:
    ```python
    input_file = "input.xlsx"
    output_file = "output.xlsx"

    xls_headers = [["first name",False,-1],["spouse/partner name",False,-1],["last name",False,-1]]
    ```

2. Run the script:
    ```bash
    python3 xls-combine-name-cols.py
    ```
