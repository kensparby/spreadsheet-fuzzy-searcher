import { useState, useRef, type JSX } from "react";
import * as XLSX from "xlsx";
import Fuse, { type FuseResultMatch } from "fuse.js";

import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import {
  Table,
  TableBody,
  TableCell,
  TableHead,
  TableHeader,
  TableRow,
} from "@/components/ui/table";

type Row = Record<string, unknown>;
type SearchResult = { item: Row; matches?: FuseResultMatch[]; score?: number };

export default function KnowledgeBaseApp() {
  const [data, setData] = useState<Row[]>([]);
  const [query, setQuery] = useState("");
  const [results, setResults] = useState<SearchResult[]>([]);

  const fileInputRef = useRef<HTMLInputElement>(null);
  
  // Excel upload
  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
  const file = e.target.files?.[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = (evt) => {
    const buffer = evt.target?.result as ArrayBuffer;
    try {
      const wb = XLSX.read(buffer, { type: "array" });
      const sheetName = wb.SheetNames[0];
      const sheet = wb.Sheets[sheetName];
      const rows = XLSX.utils.sheet_to_json<Row>(sheet, {
        defval: "",  // keep empty cells as empty strings
        raw: false,  // let XLSX parse dates/numbers nicely
      });

      // remove the first property from each row
      const trimmedRows = rows.map((row) => {
        return Object.fromEntries(Object.entries(row).slice(1));
      })

      setData(trimmedRows);
      setResults(rows.map((item) => ({ item, matches: [] })));
    } catch (err) {
      console.error("Failed to parse workbook:", err);
    }
  };
  reader.readAsArrayBuffer(file);                              // ✅ use ArrayBuffer

  // allow re-uploading same file later
  e.currentTarget.value = "";
};

  // Highlight helper
  const highlightText = (
    text: string,
    matches: FuseResultMatch[] | undefined,
    key: string
  ) => {
    if (!matches?.length) return text;

    const m = matches.find((mm) => mm.key === key);
    if (!m?.indices?.length) return text;

    let last = 0;
    const parts: (string | JSX.Element)[] = [];
    for (let i = 0; i < m.indices.length; i++) {
      const [start, end] = m.indices[i]!;
      parts.push(text.slice(last, start));
      parts.push(<mark key={i}>{text.slice(start, end + 1)}</mark>);
      last = end + 1;
    }
    parts.push(text.slice(last));
    return parts;
  };

  // AND-search across space-separated terms with highlighting
  const handleSearch = (e: React.ChangeEvent<HTMLInputElement>) => {
    const value = e.target.value;
    setQuery(value);

    if (!value || data.length === 0) {
      setResults(data.map((item) => ({ item, matches: [] })));
      return;
    }

    const columns = Object.keys(data[0] ?? {});
    const fuse = new Fuse<Row>(data, {
      keys: columns,
      threshold: 0.2, // tweak fuzziness here
      includeMatches: true,
      includeScore: true,
      ignoreLocation: true,
      findAllMatches: true,
      minMatchCharLength: 2,
    });

    const terms = value.split(/\s+/).filter(Boolean);
    if (terms.length === 0) {
      setResults(data.map((item) => ({ item, matches: [] })));
      return;
    }

    // Search each term independently
    const perTerm = terms.map((t) => fuse.search(t));

    // Build maps termIndex -> Map(itemKey -> result)
    const termMaps = perTerm.map((list) => {
      const m = new Map<string, SearchResult>();
      for (const r of list) {
        const key = JSON.stringify(r.item); // simple stable key
        m.set(key, { item: r.item, matches: r.matches as FuseResultMatch[] | undefined, score: r.score ?? 0 });
      }
      return m;
    });

    // Intersect keys to enforce AND logic
    const commonKeys = [...termMaps[0].keys()].filter((k) =>
      termMaps.every((m) => m.has(k))
    );

    // Merge matches + average score
    const merged: SearchResult[] = commonKeys.map((k) => {
      const rs = termMaps.map((m) => m.get(k)!);
      const item = rs[0].item;
      const matches = rs.flatMap((r) => r.matches ?? []);
      const avgScore =
        rs.reduce((acc, r) => acc + (r.score ?? 0), 0) / rs.length;
      return { item, matches, score: avgScore };
    });

    // Best (lowest) score first
    merged.sort((a, b) => (a.score ?? 0) - (b.score ?? 0));
    setResults(merged);
  };

  const columns = data.length ? Object.keys(data[0]!) : [];

  return (
    <div className="p-6 max-w-[90vw] mx-auto space-y-6">
    <h1 className="text-3xl font-bold">Knowledge Base Search</h1>

      {/* File Upload */}
      <div className="flex items-center gap-2">
      <input
        ref={fileInputRef}
        type="file"
        accept=".xlsx,.xls,.csv"
        onChange={handleFileUpload}
        className="hidden"
      />
      <Button
        type="button"
        onClick={() => fileInputRef.current?.click()}
      >
        Upload Excel
      </Button>
    </div>

      {/* Search */}
      {data.length > 0 && (
        <Input
          type="text"
          value={query}
          onChange={handleSearch}
          placeholder="Search (space-separated, AND logic)…"
        />
      )}

      {/* Results */}
      {results.length > 0 && (
        <div className="rounded-lg border shadow">
          <Table>
            <TableHeader>
              <TableRow>
                {columns.map((key) => (
                  <TableHead key={key}>{key}</TableHead>
                ))}
              </TableRow>
            </TableHeader>
            <TableBody>
              {results.map((res, i) => (
                <TableRow key={i}>
                  {columns.map((key, j) => {
                    const raw = res.item[key];
                    const text =
                      typeof raw === "string"
                        ? raw
                        : raw == null
                        ? ""
                        : String(raw);
                    return (
                      <TableCell key={j} className="whitespace-normal break-words max-w-xs">
                        {highlightText(text, res.matches, key)}
                      </TableCell>
                    );
                  })}
                </TableRow>
              ))}
            </TableBody>
          </Table>
        </div>
      )}
    </div>
  );
}
