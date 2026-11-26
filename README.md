# Cellestra

A spreadsheet editor for OpenDocument spreadsheets (`.ods`). You can use it in your browser, on your NAS from any browser on your network, or as a desktop app, all from the same code, and LibreOffice files survive the round trip.

![Cellestra in dark mode with the Home ribbon open: a small monthly budget with bold column headings Month, Rent, Groceries, Transit and Total, amounts in dollars for January to April, and an Average row whose selected total cell shows the formula =AVERAGE(E2:E5) in the formula bar, with the grid zoomed to 160%](docs/images/budget.webp)

- **Works with LibreOffice files.** It reads and writes `.ods`, keeps anything it doesn't understand intact, imports and exports CSV and TSV, and opens Excel `.xlsx` files as new `.ods` workbooks.
- **Formulas like LibreOffice Calc.** 47 functions with the same syntax, automatic recalculation and circular-reference detection.
- **Every command from the keyboard.** `⌘K` opens the command menu and `⌘⇧P` switches sheets.

## Getting started

Download the desktop app for macOS, Windows or Linux from the [Releases page](https://github.com/gridlocdev/cellestra/releases/latest), use it in your browser at [gridloc.dev](https://gridloc.dev/apps/cellestra/), or run it on your NAS with Docker:

```bash
docker build -t cellestra . && docker run -d -p 8787:8787 -e CELLESTRA_TOKEN=<a-long-secret> -v /volume1/Documents:/data cellestra
```

See [Getting started](docs/getting-started.md) for which file to pick, the first-launch warnings, and running it from source.

## Learn more

- [Features](docs/features.md) · [Controls](docs/controls.md)
- [All documentation](docs/README.md) · [Working on the code](docs/development.md) · [Glossary](GLOSSARY.md)

## License

[MIT](LICENSE)
