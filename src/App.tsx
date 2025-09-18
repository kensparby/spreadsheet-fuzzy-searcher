import { useEffect, useMemo, useRef, useState, type JSX } from "react";
import * as XLSX from "xlsx";
import Fuse, { type FuseResultMatch } from "fuse.js";
import { expandMerges, hasContent, getNonEmptyRowSet, getSavedSheet, saveSheet, getHyperlinkFromCell, getSavedColumns, saveColumns } from "./lib/utils";

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
  DialogHeader,
  DialogTitle,
} from "@/components/ui/dialog";
import {
  Tooltip,
  TooltipContent,
  TooltipTrigger,
} from "@/components/ui/tooltip";
import {
  Select,
  SelectTrigger,
  SelectValue,
  SelectContent,
  SelectItem,
} from "@/components/ui/select";


type Row = Record<string, unknown>;
type SearchResult = { item: Row; matches?: FuseResultMatch[]; score?: number };

export default function KnowledgeBaseApp() {
  const [data, setData] = useState<Row[]>([]);
  const [query, setQuery] = useState("");
  const [fuzz, setFuzz] = useState(() => {
    const storedData = localStorage.getItem("fuzz");
    return storedData ? Number(storedData) : 0.2;
  });
  const [results, setResults] = useState<SearchResult[]>([]);
  const [modalOpen, setModalOpen] = useState(false);
  const [modalTitle, setModalTitle] = useState<string>("");
  const [modalText, setModalText] = useState<string>("");
  const [workbook, setWorkbook] = useState<XLSX.WorkBook | null>(null);
  const [fileName, setFileName] = useState<string>("");
  const [sheetNames, setSheetNames] = useState<string[]>([]);
  const [selectedSheet, setSelectedSheet] = useState<string>("");
  const [linkMap, setLinkMap] = useState<Map<number, Map<String, string>>>(new Map());
  const [colVisibility, setColVisibility] = useState<Record<string, boolean>>({});

  const fileInputRef = useRef<HTMLInputElement>(null);

  /**
   * Opens a modal with the given cell's content, for ease of reading 
   * and click to copy functionality.
   * @param col the column name to use as the title
   * @param value the value to show in the modal
   */
  function openCell(col: string, value: unknown) {
    const text = typeof value === "string" ? value : value == null ? "" : String(value);
    setModalTitle(col);
    setModalText(text);
    setModalOpen(true);
  }


  /**
   * Returns an array of visible column names based on the data and search results.
   * A column is considered visible if it contains content in any of the search results.
   * 
   * @returns {string[]} An array of visible column names.
   */
  const visibleColumns = useMemo(() => {
    if (!results.length) return [];
    const all = data.length ? Object.keys(data[0]!) : [];
    return all
      .filter((k) => k !== "__rowNum__")
      .filter((k) => results.some((r) => hasContent((r.item as Record<string, unknown>)[k])))
      .filter((k) => colVisibility[k] !== false);
  }, [data, results, colVisibility]);


  function parseSheetByName(wb: XLSX.WorkBook, sheetName: string) {
    const sheet = wb.Sheets[sheetName];
    if (!sheet) return;

    // --- COLLECT EXCEL HYPERLINKS (BEFORE sheet_to_json) ---
    const ref = sheet["!ref"] as string | undefined;
    if (!ref) return;
    const range = XLSX.utils.decode_range(ref);
    const headerRow = range.s.r;

    // Build header names from the first row
    const headers: string[] = [];
    for (let c = range.s.c; c <= range.e.c; c++) {
      const addr = XLSX.utils.encode_cell({ r: headerRow, c });
      const cell = (sheet as any)[addr];
      headers.push(cell?.v != null ? String(cell.v) : `__col${c}__`);
    }

    // rowNumber(1-based) -> Map<header -> href>
    const lm = new Map<number, Map<string, string>>();
    for (let r = headerRow + 1; r <= range.e.r; r++) {
      for (let c = range.s.c; c <= range.e.c; c++) {
        const addr = XLSX.utils.encode_cell({ r, c });
        const cell = (sheet as any)[addr];
        const link = getHyperlinkFromCell(cell);
        if (!link) continue;

        const header = headers[c - range.s.c];
        const rowNum1 = r + 1; // Excel is 1-based
        if (!lm.has(rowNum1)) lm.set(rowNum1, new Map());
        lm.get(rowNum1)!.set(header, link);
      }
    }
    setLinkMap(lm);

    // compute empties BEFORE merges
    const keepRows = getNonEmptyRowSet(sheet);
    expandMerges(sheet);

    type RowWithNum = Row & { __rowNum__?: number };
    const rows = XLSX.utils.sheet_to_json<RowWithNum>(sheet, {
      defval: "",
      raw: false,
    });

    // drop spacer rows
    const filteredRows = rows.filter(r => {
      const rowNum = (r.__rowNum__ ?? -1) + 1;
      return keepRows.has(rowNum);
    });

    // drop first column but keep original __rowNum__ (0-based) for hyperlink lookup
    const finalRows = filteredRows.map((row) => {
      const { __rowNum__, ...rest } = row;
      const obj = { ...rest } as Row;
      (obj as any).__rowNum__ = __rowNum__;
      return obj as Row;
    });

    setData(finalRows);
    setResults(finalRows.map((item) => ({ item, matches: [] })));

    const cols = finalRows.length ? Object.keys(finalRows[0] as Record<string, unknown>) : [];
    const defaults = Object.fromEntries(cols.map(k => [k, k !== "__rowNum__"]));
    const saved = getSavedColumns(fileName, sheetName) ?? {};
    setColVisibility({ ...defaults, ...saved });
  }

  /**
   * Handles file upload events, by reading the file into a worksheet and
   * populating the data state with the contents of the sheet. The first column
   * is dropped, and only rows with at least one non-empty cell are kept.
   * @param e The file input change event.
   */
  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      const buffer = evt.target?.result as ArrayBuffer;
      try {
        const wb = XLSX.read(buffer, { type: "array" });
        const names = wb.SheetNames || [];
        setWorkbook(wb);
        setSheetNames(names);
        const fname = file.name || "";
        setFileName(fname);

        // pick saved sheet or default to first
        const preferred = getSavedSheet(fname);
        const chosen = preferred && names.includes(preferred) ? preferred : names[0] || "";
        setSelectedSheet(chosen);

        if (chosen) parseSheetByName(wb, chosen);
      } catch (err) {
        console.error("Klarte ikke lese fil:", err);
      }
    };
    reader.readAsArrayBuffer(file);
    e.currentTarget.value = ""; // allow re-uploading same file
  };


  function handleSheetChange(name: string) {
    setSelectedSheet(name);
    if (workbook && name) {
      parseSheetByName(workbook, name);
      if (fileName) saveSheet(fileName, name);
    }
  }


  /**
   * Highlights the given text by wrapping the matched parts in <mark> elements.
   * @param text The text to highlight
   * @param matches The Fuse.js match data
   * @param key The key to highlight in the match data
   * @returns The highlighted text as a React node
   */
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

  // Columns & Fuse index
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

  /**
   * Runs a full-text search on the data using the Fuse.js library.
   * The search query is split on whitespace and each term is searched
   * individually. The results are then merged using an intersection of
   * keys with the average score of the individual searches.
   * @param q The search query
   */
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

    // Create a map for each term
    const termMaps = perTerm.map((list) => {
      const m = new Map<string, SearchResult>();
      for (const r of list) {
        const key = JSON.stringify(r.item);
        m.set(key, { item: r.item, matches: r.matches as FuseResultMatch[], score: r.score ?? 0 });
      }
      return m;
    });

    // Find keys that are present in all maps
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
  }, [data, query, fuzz, fuse]);

  // ctrl+k to focus on input text field
  document.addEventListener('keydown', function (e) {
    const textField = document.querySelector('input#textField') as HTMLElement;
    if (e.ctrlKey && e.key === 'k') {
      e.preventDefault();
      textField?.focus();
    }
  })


  /**
   * If the given string does not start with a protocol, prefix it with "https://".
   * @param s a string that may or may not start with a protocol
   * @returns a string with a protocol if one was not present
   */
  const normalizeHref = (s: string) =>
    s.startsWith("http") ? s : s.startsWith("www.") ? `https://${s}` : s;


  const urlOnlyGlobal = /(https?:\/\/[^\s]+|www\.[^\s]+)/gi;              // no mailto here
  const urlOnlySingle = /^(https?:\/\/[^\s]+|www\.[^\s]+)$/i;             // no mailto here
  const emailGlobal =
    /(mailto:)?([A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,})(?![^<]*>)/gi;      // plain emails or mailto:

  // For the modal: turn http(s)/www and emails into anchors
  const linkifyInModal = (s: string) => {
    // one pass that matches either url or email
    const combined = new RegExp(
      `${urlOnlyGlobal.source}|${emailGlobal.source}`,
      "gi"
    );

    const parts: (string | JSX.Element)[] = [];
    let last = 0;

    s.replace(combined, (match: string, ...args: any[]) => {
      const index = args[args.length - 2] as number; // match index
      if (index > last) parts.push(s.slice(last, index));

      const m = match.trim();

      // URL?
      if (m.startsWith("http") || m.startsWith("www.")) {
        parts.push(
          <a
            key={`u-${index}`}
            href={normalizeHref(m)}
            title={normalizeHref(m)}
            target="_blank"
            rel="noopener noreferrer"
            className="underline text-primary"
          >
            {m}
          </a>
        );
      } else {
        // Email (with or without mailto:)
        const withoutPrefix = m.replace(/^mailto:/i, "");
        parts.push(
          <a
            key={`e-${index}`}
            href={`mailto:${withoutPrefix}`}
            className="underline text-primary"
          >
            {withoutPrefix}
          </a>
        );
      }

      last = index + m.length;
      return m;
    });

    if (last < s.length) parts.push(s.slice(last));
    return parts;
  };


  /**
   * Returns true if the given string is a single URL (i.e. only contains one URL)
   * @param s The string to check.
   * @returns True if the string is a single URL.
   */
  const isSingleLink = (s: string) => urlOnlySingle.test(s.trim());

  return (
    <div className="p-6 max-w-[2600px] mx-auto space-y-6">
      <h1 className="text-3xl font-bold tracking-tight">Søk i Excel-ark</h1>

      <Card className="p-3 space-y-1">
        <div className="flex flex-wrap items-center gap-3">
          <input
            ref={fileInputRef}
            type="file"
            accept=".xlsx,.xls,.csv"
            onChange={handleFileUpload}
            className="hidden"
          />
          <Button type="button" onClick={() => fileInputRef.current?.click()}>
            Last inn Excel-fil
          </Button>

          <div className="flex items-center gap-3 w-full sm:w-auto sm:min-w-[320px]">
            {sheetNames.length > 1 && (
              <>
                <Label className="whitespace-nowrap py-2 pl-8">
                  Ark
                </Label>
                <div>
                  <Select value={selectedSheet} onValueChange={handleSheetChange}>
                    <SelectTrigger aria-label="Velg ark">
                      <SelectValue placeholder="Velg ark" />
                    </SelectTrigger>
                    <SelectContent>
                      {sheetNames.map((n) => (
                        <SelectItem key={n} value={n}>{n}</SelectItem>
                      ))}
                    </SelectContent>
                  </Select>
                </div>
              </>
            )}

            <Label htmlFor="fuzz" className="whitespace-nowrap py-2 pl-4 border-l-2">
              Nøyaktighet
            </Label>
            <div className="flex-1 w-[120px]">
              <Slider
                id="fuzz"
                value={[-fuzz + 1]}
                step={0.1}
                min={0}
                max={1}
                onValueChange={([v]) => setFuzz(() => {
                  localStorage.setItem('fuzz', String(Math.round((1 - v) * 10) / 10));
                  return Number(Math.round((1 - v) * 10) / 10);
                })}
              />
            </div>
            <div className="w-[3ch] text-sm tabular-nums">{(1 - fuzz).toFixed(1)}</div>
          </div>

          {/* Column visibility toggles */}
          {data.length > 0 && (
            <div className="flex flex-wrap gap-3 items-center pt-1 pl-4 border-l-2">
              {(
                // base set: columns that currently have any content (same as visibleColumns base but before user filter)
                (data.length ? Object.keys(data[0]!) : [])
                  .filter((k) => k !== "__rowNum__")
                  .filter((k) => results.some((r) => hasContent((r.item as Record<string, unknown>)[k])))
              ).map((col) => {
                const label = col;
                return (
                  <label key={col} className="inline-flex items-center gap-2 px-2 py-1 rounded border border-border">
                    <input
                      type="checkbox"
                      className="h-4 w-4 accent-primary"
                      checked={colVisibility[col] !== false}
                      onChange={(e) => {
                        const next = { ...colVisibility, [col]: e.target.checked };
                        setColVisibility(next);
                        if (fileName && selectedSheet) saveColumns(fileName, selectedSheet, next);
                      }}
                    />
                    <span className="max-w-[10ch] truncate text-sm" title={label}>
                      {label}
                    </span>
                  </label>
                );
              })}
            </div>
          )}
        </div>


        <Tooltip>
          <TooltipTrigger>
            {data.length > 0 && (
              <Input
                id="textField"
                type="text"
                value={query}
                onChange={(e) => setQuery(e.target.value)} // effect does the actual searching
                // placeholder="Search (space-separated, AND logic)…"
                placeholder="Søk… (separer søkeord med mellomrom. Trykk Ctrl+K for å fokusere på dette feltet)"
              />
            )}
          </TooltipTrigger>
          <TooltipContent>
            <Label>Tips: Bruk en apostrof foran et ord for å kun søke etter ordet nøyaktig slik du skriver det, eksempel: <kbd>&apos;avgifter</kbd></Label>
          </TooltipContent>
        </Tooltip>
      </Card>

      {results.length > 0 && (
        <div className="rounded-lg border shadow-sm overflow-x-auto">
          <Table>
            <TableHeader>
              <TableRow>
                {visibleColumns.map((key) => (
                  <TableHead key={key}>{key}</TableHead>
                ))}
              </TableRow>
            </TableHeader>
            <TableBody>
              {results.map((res, i) => (
                <TableRow key={i}>
                  {visibleColumns.map((key, j) => {
                    const raw = res.item[key];
                    const text = typeof raw === "string" ? raw : raw == null ? "" : String(raw);

                    // 1) Excel-native hyperlink on this cell? (use it directly, no modal)
                    const rowNum1Based = ((res.item as any).__rowNum__ as number | undefined) != null
                      ? ((res.item as any).__rowNum__ as number) + 1
                      : undefined;
                    const hrefFromXlsx =
                      rowNum1Based ? linkMap.get(rowNum1Based)?.get(key) : undefined;

                    if (hrefFromXlsx) {
                      const href = hrefFromXlsx.startsWith("mailto:")
                        ? hrefFromXlsx
                        : normalizeHref(hrefFromXlsx);
                      return (
                        <TableCell key={j} className="whitespace-normal break-words max-w-sm align-top">
                          <a
                            href={href}
                            target={href.startsWith("mailto:") ? undefined : "_blank"}
                            rel={href.startsWith("mailto:") ? undefined : "noopener noreferrer"}
                            className="underline text-primary"
                            onClick={(e) => e.stopPropagation()} // don't open modal
                            title={href}
                          >
                            {text /* keep any display text from the cell */}
                          </a>
                        </TableCell>
                      );
                    }

                    // 2) Fallback: if the displayed text is itself a single URL → link it directly
                    if (isSingleLink(text)) {
                      return (
                        <TableCell key={j} className="whitespace-normal break-words max-w-sm align-top">
                          <a
                            href={normalizeHref(text.trim())}
                            target="_blank"
                            rel="noopener noreferrer"
                            className="underline text-primary"
                            onClick={(e) => e.stopPropagation()}
                          >
                            {text}
                          </a>
                        </TableCell>
                      );
                    }

                    // Otherwise keep existing behavior of opening modal
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

      <div className="pt-8 text-center text-sm text-muted-foreground">
        <Button
          asChild
          variant="link"
          className="text-muted-foreground hover:text-primary"
        >
          <a
            href="https://github.com/kensparby/spreadsheet-fuzzy-searcher"
            target="_blank"
            rel="noopener noreferrer"
          >
            Se kildekoden på Github
          </a>
        </Button>
      </div>

      <Dialog open={modalOpen} onOpenChange={setModalOpen}>
        <DialogContent className="sm:max-w-2xl">
          <DialogHeader>
            <DialogTitle>{modalTitle}</DialogTitle>
          </DialogHeader>

          <div className="max-h-[70vh] overflow-auto whitespace-pre-wrap break-words leading-relaxed text-sm">
            {linkifyInModal(modalText)}
          </div>

          <div className="flex justify-end gap-2 pt-2">
            <Button
              type="button"
              variant="secondary"
              onClick={() => navigator.clipboard?.writeText(modalText)}
              className="hover:bg-slate-200 hover:text-slate-900 active:bg-slate-300 active:text-slate-950 dark:hover:bg-slate-800 dark:hover:text-slate-50 dark:active:bg-slate-700 dark:active:text-slate-50"
            >
              Kopier
            </Button>
            <Button type="button" onClick={() => setModalOpen(false)}>Close</Button>
          </div>
        </DialogContent>
      </Dialog>
    </div>

  );
}
