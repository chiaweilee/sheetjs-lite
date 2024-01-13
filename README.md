# r-xlsx

A fork of `xlsx`. Spreadsheet data parser only, but smaller.

## Difference

Compared with official `xlsx`.

### APIs

| Module            | API                          | Support |
| ----------------- | ---------------------------- | ------- |
| Reading Files     | XLSX.read                    | ✔      |
| Reading Files     | XLSX.readFile                | ✘       |
| Writing Files     | XLSX.write                   | ✘       |
| Writing Files     | XLSX.writeFile               | ✘       |
| Writing Files     | XLSX.writeXLSX               | ✘       |
| Writing Files     | XLSX.writeFileXLSX           | ✘       |
| Utility Functions | XLSX.utils.sheet_to_json     | ✔      |
| Utility Functions | XLSX.utils.json_to_sheet     | ✘       |
| Utility Functions | XLSX.utils.table_to_sheet    | ✘       |
| Utility Functions | XLSX.utils.sheet_to_html     | ✔      |
| Utility Functions | XLSX.utils.sheet_to_csv      | ✔      |
| Utility Functions | XLSX.utils.sheet_to_txt      | ✔      |
| Utility Functions | XLSX.utils.sheet_to_formulae | ✔      |
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

## Usage

```ts
import XLSX from "r-xlsx";

const workbook = XLSX.read(fileData);
const { SheetNames, Sheets } = workbook;
const content = XLSX.utils.sheet_to_json(Sheets[SheetNames[0]]);
```

## Use Web Worker

*require webpack5*

```ts
const worker = new Worker(new URL("./myworker.ts", import.meta.url));
worker.postMessage(fileData);
```

```ts
// myworker.ts
import XLSX from "r-xlsx";

self.onmessage = ({ data }) => {
  const result: any[] = [];
  const { SheetNames, Sheets } = XLSX.read(data);
  SheetNames.forEach((name: string) => {
    const content = XLSX.utils.sheet_to_json(Sheets[name]);
    result.push(content);
  });
  self.postMessage(result);
};
```
