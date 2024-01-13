# r-xlsx

This is a fork of `xlsx`, which has spreadsheet data parser only, and it's smaller.

## Difference

Compared with official `xlsx`.

### APIs

| Module            | API                          | Support |
| ----------------- | ---------------------------- | ------- |
| Reading Files     | XLSX.read                    | ✔       |
| Reading Files     | XLSX.readFile                | ✘       |
| Writing Files     | XLSX.write                   | ✘       |
| Writing Files     | XLSX.writeFile               | ✘       |
| Writing Files     | XLSX.writeXLSX               | ✘       |
| Writing Files     | XLSX.writeFileXLSX           | ✘       |
| Utility Functions | XLSX.utils.sheet_to_json     | ✔       |
| Utility Functions | XLSX.utils.json_to_sheet     | ✘       |
| Utility Functions | XLSX.utils.table_to_sheet    | ✘       |
| Utility Functions | XLSX.utils.sheet_to_html     | ✔       |
| Utility Functions | XLSX.utils.sheet_to_csv      | ✔       |
| Utility Functions | XLSX.utils.sheet_to_txt      | ✔       |
| Utility Functions | XLSX.utils.sheet_to_formulae | ✔       |
| Utility Functions | XLSX.stream                  | ✘       |

### Dist Size

| Dist             | Origin Size | Lite Size                               |
| ---------------- | ----------- | --------------------------------------- |
| xlsx.full.min    | 944 KB      | 845 KB (10% smaller, 306 KB after gzip) |
| xlsx.core.min.js | 500 KB      | 403 KB (20% smaller, 131 KB after gzip) |
| xlsx.mini.min.js | 276 KB      | 219 KB (20% smaller, 60 KB after gzip)  |

## Install

```
npm install r-xlsx
```
