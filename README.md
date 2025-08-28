# Mini Excel Editor

A lightweight, in-browser spreadsheet editor that supports CSV and XLSX files. The current implementation is a single-page app written in vanilla JavaScript and HTML.

## Features
- Open or save data as CSV or XLSX
- Add and remove rows or columns
- Evaluate basic formulas (e.g. `=SUM(A1:B2)`)
- Built-in debugging log and self-test routine

## Getting Started
Open `index.html` in any modern browser. No build step or server is required, but you can run a simple static server for convenience:

```
python3 -m http.server 8000
```

Then navigate to `http://localhost:8000/index.html`.

## Roadmap
The long-term goal is to rebuild this prototype using a modern stack (React, TypeScript, Vite, etc.) with richer spreadsheet features. See [development.md](development.md) for the detailed plan.

