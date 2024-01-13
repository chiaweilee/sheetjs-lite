import { useState } from "react";

function DemoPage() {
  const [loading, $loading] = useState(false);
  const [result, $result] = useState([]);

  const onChange = (e: any) => {
    const file = e.target.files?.[0];
    if (file) {
      $loading(true);
      const reader = new FileReader();
      reader.onload = (e: any) => {
        const worker = new Worker(
          new URL("../helpers/worker.ts", import.meta.url)
        );
        worker.postMessage(e.target.result);
        worker.onmessage = ({ data }) => {
          result.push(data);
          $result(result.slice(0));
        };
        $loading(false);
      };
      reader.readAsArrayBuffer(file);
    }
    e.target.value = null;
  };

  return (
    <div style={{ padding: "8px" }}>
      <input type="file" onChange={onChange} disabled={loading} />
      {loading && <div>loading...</div>}
      <div>result: {JSON.stringify(result)}</div>
    </div>
  );
}

export default DemoPage;
