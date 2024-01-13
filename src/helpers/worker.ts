import XLSX from "../../dist/xlsx.full.min.js";

self.onmessage = ({ data }) => {
  const result: any[] = [];
  const { SheetNames, Sheets } = XLSX.read(data);
  SheetNames.forEach((name: string) => {
    const content = XLSX.utils.sheet_to_json(Sheets[name]);
    result.push(content);
  });
  self.postMessage(result);
};
