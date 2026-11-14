# Cellestra

A spreadsheet editor for OpenDocument spreadsheets (`.ods`). You can use it in your browser, on your NAS from any browser on your network, or as a desktop app, all from the same code, and LibreOffice files survive the round trip.

- **Works with LibreOffice files.** It reads and writes `.ods`, keeps anything it doesn't understand intact, and imports and exports CSV and TSV.
- **Formulas like LibreOffice Calc.** 47 functions with the same syntax, automatic recalculation and circular-reference detection.
- **Every command from the keyboard.** `⌘K` opens the command menu and `⌘⇧P` switches sheets.

## Getting started

There are no prebuilt downloads yet. Run the desktop app from source with Node 22.18 or later and [just](https://github.com/casey/just):

```bash
just setup && just run-desktop
```

See [Getting started](docs/getting-started.md) for running it in the browser, on a NAS with Docker, and installing the desktop app.

## Learn more

- [Features](docs/features.md) · [Controls](docs/controls.md)
- [All documentation](docs/README.md) · [Working on the code](docs/development.md)

## License

[MIT](LICENSE)
