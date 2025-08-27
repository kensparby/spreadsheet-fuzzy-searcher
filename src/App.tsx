import { useEffect, useMemo, useRef, useState, type JSX } from "react";
import * as XLSX from "xlsx";
import Fuse, { type FuseResultMatch } from "fuse.js";

import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Slider } from "@/components/ui/slider";
import { Card } from "@/components/ui/card";
import {
  Table,
  TableBody,
  TableCell,
  TableHead,
  TableHeader,
  TableRow,
} from "@/components/ui/table";
import {
  Dialog,
  DialogContent,
  DialogDescription,
  DialogHeader,
  DialogTitle,
  DialogTrigger,
} from "@/components/ui/dialog";

type Row = Record<string, unknown>;
type SearchResult = { item: Row; matches?: FuseResultMatch[]; score?: number };

export default function KnowledgeBaseApp() {
  const [data, setData] = useState<Row[]>([]);
  const [query, setQuery] = useState("");
  const [fuzz, setFuzz] = useState(0.2);
  const [results, setResults] = useState<SearchResult[]>([]);
  const [modalOpen, setModalOpen] = useState(false);
  const [modalTitle, setModalTitle] = useState<string>("");
  const [modalText, setModalText] = useState<string>("");

  const fileInputRef = useRef<HTMLInputElement>(null);

  // ---------- handle modal ------------
  const openCell = (col: string, value: unknown) => {
    const text = typeof value === "string" ? value : value == null ? "" : String(value);
    setModalTitle(col);
    setModalText(text);
    setModalOpen(true);
  };


  // ---------- Upload & parse ----------
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
          defval: "",
          raw: false,
        });

        // Drop the first column
        const trimmedRows = rows.map((row) =>
          Object.fromEntries(Object.entries(row).slice(1))
        );

        setData(trimmedRows);
        setResults(trimmedRows.map((item) => ({ item, matches: [] }))); // ✅ use trimmedRows
      } catch (err) {
        console.error("Failed to parse workbook:", err);
      }
    };
    reader.readAsArrayBuffer(file);
    e.currentTarget.value = ""; // allow re-uploading same file
  };

  // ---------- Highlight helper ----------
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

  // ---------- Columns & Fuse index ----------
  const columns = useMemo(() => (data.length ? Object.keys(data[0]!) : []), [data]);

  const fuse = useMemo(() => {
    if (!data.length) return null;
    return new Fuse<Row>(data, {
      keys: columns,
      threshold: fuzz,          // <- slider-controlled
      includeMatches: true,
      includeScore: true,
      ignoreLocation: true,
      findAllMatches: true,
      minMatchCharLength: 2,
      ignoreDiacritics: true,
      useExtendedSearch: true,
    });
  }, [data, columns, fuzz]);

  // ---------- Search logic (AND across terms) ----------
  const runSearch = (q: string) => {
    if (!q || !data.length || !fuse) {
      setResults(data.map((item) => ({ item, matches: [] })));
      return;
    }

    const terms = q.split(/\s+/).filter(Boolean);
    if (terms.length === 0) {
      setResults(data.map((item) => ({ item, matches: [] })));
      return;
    }

    const perTerm = terms.map((t) => fuse.search(t));

    const termMaps = perTerm.map((list) => {
      const m = new Map<string, SearchResult>();
      for (const r of list) {
        const key = JSON.stringify(r.item); // simple key
        m.set(key, { item: r.item, matches: r.matches as FuseResultMatch[], score: r.score ?? 0 });
      }
      return m;
    });

    // AND = intersection of keys
    const commonKeys = [...termMaps[0].keys()].filter((k) =>
      termMaps.every((m) => m.has(k))
    );

    const merged: SearchResult[] = commonKeys.map((k) => {
      const rs = termMaps.map((m) => m.get(k)!);
      const item = rs[0].item;
      const matches = rs.flatMap((r) => r.matches ?? []);
      const avgScore =
        rs.reduce((acc, r) => acc + (r.score ?? 0), 0) / rs.length;
      return { item, matches, score: avgScore };
    });

    merged.sort((a, b) => (a.score ?? 0) - (b.score ?? 0));
    setResults(merged);
  };

  // Recompute whenever data/query/fuzz change
  useEffect(() => {
    runSearch(query);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [data, query, fuzz, fuse]);

  // ---------- UI ----------
  return (
    <div className="p-6 max-w-[1000px] mx-auto space-y-6">
      <h1 className="text-3xl font-bold tracking-tight">Knowledge Base Search</h1>

      <Card className="p-4 space-y-4">
        <div className="flex flex-wrap items-center gap-3">
          <input
            ref={fileInputRef}
            type="file"
            accept=".xlsx,.xls,.csv"
            onChange={handleFileUpload}
            className="hidden"
          />
          <Button type="button" onClick={() => fileInputRef.current?.click()}>
            Upload Excel
          </Button>

          <div className="flex items-center gap-3 w-full sm:w-auto sm:min-w-[320px]">
            <Label htmlFor="fuzz" className="whitespace-nowrap">
              Fuzziness
            </Label>
            <div className="flex-1">
              <Slider
                id="fuzz"
                value={[fuzz]}
                step={0.01}
                min={0}
                max={1}
                onValueChange={([v]) => setFuzz(Number(v))}
              />
            </div>
            <div className="w-12 text-right tabular-nums">{fuzz.toFixed(2)}</div>
          </div>
        </div>

        {data.length > 0 && (
          <Input
            type="text"
            value={query}
            onChange={(e) => setQuery(e.target.value)} // effect does the searching
            placeholder="Search (space-separated, AND logic)…"
          />
        )}
      </Card>

      {results.length > 0 && (
        <div className="rounded-lg border shadow-sm overflow-x-auto">
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
                      typeof raw === "string" ? raw : raw == null ? "" : String(raw);
                    return (
                      <TableCell
                        key={j}
                        onClick={() => openCell(key, raw)}
                        onKeyDown={(ev) => (ev.key === "Enter" || ev.key === " ") && openCell(key, raw)}
                        role="button"
                        tabIndex={0}
                        className="whitespace-normal break-words max-w-sm align-top cursor-pointer hover:bg-muted/50"
                        title="Click to expand"
                      >
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
      <Dialog open={modalOpen} onOpenChange={setModalOpen}>
        <DialogContent className="sm:max-w-2xl">
          <DialogHeader>
            <DialogTitle>{modalTitle}</DialogTitle>
          </DialogHeader>

          <div className="max-h-[70vh] overflow-auto whitespace-pre-wrap break-words leading-relaxed text-sm">
            {modalText}
          </div>

          <div className="flex justify-end gap-2 pt-2">
            <Button
              type="button"
              variant="secondary"
              onClick={() => navigator.clipboard?.writeText(modalText)}
            >
              Copy
            </Button>
            <Button type="button" onClick={() => setModalOpen(false)}>Close</Button>
          </div>
        </DialogContent>
      </Dialog>

    </div>
  );
}
