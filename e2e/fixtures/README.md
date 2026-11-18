# Fixtures

Made-up workbooks for the browser scenarios. Each folder is one fixture: `seedData('basic')` copies `basic/` into `.e2e-data`, which the test server serves as the `E2E` folder. Never copy real spreadsheets here.

The `basic` workbooks were written with the app's own CLI, so they are real files of the current format:

```bash
cd e2e/fixtures/basic
just cli new budget.ods
just cli set budget.ods --in-place A1=Month B1=Rent C1=Groceries D1=Total A2=January B2=1200 C2=310 "D2==B2+C2" A3=February B3=1200 C3=285 "D3==B3+C3"
just cli new empty.ods
just cli new "Café ☕ notes.ods"
just cli set "Café ☕ notes.ods" --in-place "A1=Crème brûlée" "A2=日本語のテキスト"
```

`expenses.xlsx` is `budget.ods` converted by LibreOffice, so the `.xlsx` import opens a file Excel-format tools really write:

```bash
soffice --headless --convert-to xlsx budget.ods   # then rename budget.xlsx to expenses.xlsx
```
