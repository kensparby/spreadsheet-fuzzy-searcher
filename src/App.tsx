import React, { useState } from "react";
import * as XLSX from "xlsx";
import Fuse from "fuse.js";

export default function KnowledgeBaseApp() {
  const [data, setData] = useState<{ [key: string]: any }[]>([]);       // parsed rows
  const [query, setQuery] = useState("");     // search input
  const [results, setResults] = useState<{ [key: string]: any }[]>([]); // search results

  // Handle Excel upload
  const handleFileUpload = (e: any) => {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      const binaryStr = evt.target?.result;
      const workbook = XLSX.read(binaryStr, { type: "binary" });

      // Assume first sheet
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];
      const jsonData = XLSX.utils.sheet_to_json(sheet) as { [key: string]: any }[];

      setData(jsonData);
      setResults(jsonData); // show everything by default
    };
    reader.readAsBinaryString(file);
  };

  // Handle search
  const handleSearch = (e: any) => {
    const value = e.target.value;
    setQuery(value);

    if (!value) {
      setResults(data); // reset
      return;
    }

    const fuse = new Fuse(data, {
      keys: Object.keys(data[0] || {}), // search all columns
      threshold: 0.3,
    });

    const found = fuse.search(value).map(res => res.item);
    setResults(found);
  };

  return (
    <div className="p-4 max-w-3xl mx-auto">
      <h1 className="text-2xl font-bold mb-4">Knowledge Base Search</h1>

      {/* File upload */}
      <input type="file" accept=".xlsx,.xls,.csv" onChange={handleFileUpload} className="mb-4" />

      {/* Search box */}
      {data.length > 0 && (
        <input
          type="text"
          value={query}
          onChange={handleSearch}
          placeholder="Search issues..."
          className="border p-2 w-full mb-4"
        />
      )}

      {/* Results table */}
      {results.length > 0 && (
        <table className="w-full border-collapse border">
          <thead>
            <tr>
              {Object.keys(results[0]).map((key) => (
                <th key={key} className="border px-2 py-1">{key}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {results.map((row, i) => (
              <tr key={i}>
                {Object.values(row).map((val, j) => (
                  <td key={j} className="border px-2 py-1">{val}</td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
      )}
    </div>
  );
}
