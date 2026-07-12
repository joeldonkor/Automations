"use client";

import { useId, useState } from "react";
import { parseRosterWorkbook, type ParsedGroup } from "@/lib/parseRoster";
import { parseAttendanceWorkbook } from "@/lib/parseAttendance";
import { computeGroupTotals } from "@/lib/computeReport";
import { buildWorkbook } from "@/lib/buildWorkbook";
import { REPORT_ROWS, type GroupTotals } from "@/lib/types";

const DEFAULT_IRREGULAR_KEYWORDS = "BN,irregular,DF,DA,moved";

function labelFromFileName(name: string): string {
  return name.replace(/\.xlsx$/i, "");
}

function checkFileTotals(group: ParsedGroup): string[] {
  if (!group.fileTotals) return [];
  const warnings: string[] = [];
  const reportedAll = group.rows.filter((r) => r.reported).length;
  const bibleStudiesAll = group.rows.reduce((sum, r) => sum + r.bibleStudies, 0);
  const hoursAll = group.rows.reduce((sum, r) => sum + (r.hours ?? 0), 0);

  const { reported, bibleStudies, hours } = group.fileTotals;
  if (reported !== null && reported !== reportedAll) {
    warnings.push(
      `${group.label}: file's Total row shows ${reported} reported, but parsed rows total ${reportedAll} — check for missed or misread rows.`
    );
  }
  if (bibleStudies !== null && bibleStudies !== bibleStudiesAll) {
    warnings.push(
      `${group.label}: file's Total row shows ${bibleStudies} Bible studies, but parsed rows total ${bibleStudiesAll}.`
    );
  }
  if (hours !== null && hours !== hoursAll) {
    warnings.push(
      `${group.label}: file's Total row shows ${hours} hours, but parsed rows total ${hoursAll}.`
    );
  }
  return warnings;
}

function dedupeLabels(labels: string[]): string[] {
  const seen = new Map<string, number>();
  return labels.map((label) => {
    const count = seen.get(label) ?? 0;
    seen.set(label, count + 1);
    return count === 0 ? label : `${label} (${count + 1})`;
  });
}

function FileIcon() {
  return (
    <svg viewBox="0 0 24 24" fill="none" className="size-4 shrink-0 text-emerald-600 dark:text-emerald-500">
      <path
        d="M6 2.75h7.5L18 7.25V19a2.25 2.25 0 0 1-2.25 2.25h-9.5A2.25 2.25 0 0 1 4 19V5A2.25 2.25 0 0 1 6.25 2.75Z"
        stroke="currentColor"
        strokeWidth="1.5"
        strokeLinejoin="round"
      />
      <path d="M13 2.75V7a1 1 0 0 0 1 1h4" stroke="currentColor" strokeWidth="1.5" strokeLinejoin="round" />
    </svg>
  );
}

function Spinner() {
  return (
    <svg viewBox="0 0 24 24" fill="none" className="size-4 animate-spin">
      <circle cx="12" cy="12" r="9" stroke="currentColor" strokeWidth="3" className="opacity-25" />
      <path d="M21 12a9 9 0 0 0-9-9" stroke="currentColor" strokeWidth="3" strokeLinecap="round" />
    </svg>
  );
}

export default function Home() {
  const monthId = useId();
  const keywordsId = useId();
  const attendanceId = useId();

  const [monthLabel, setMonthLabel] = useState("");
  const [files, setFiles] = useState<File[]>([]);
  const [irregularKeywords, setIrregularKeywords] = useState(DEFAULT_IRREGULAR_KEYWORDS);

  const [includeAttendance, setIncludeAttendance] = useState(false);
  const [weekend, setWeekend] = useState(["", "", "", ""]);
  const [midweek, setMidweek] = useState(["", "", "", ""]);
  const [attendanceFile, setAttendanceFile] = useState<File | null>(null);

  const [status, setStatus] = useState("");
  const [statusTone, setStatusTone] = useState<"idle" | "success" | "warning">("idle");
  const [isGenerating, setIsGenerating] = useState(false);
  const [isDownloading, setIsDownloading] = useState(false);
  const [warnings, setWarnings] = useState<string[]>([]);
  const [groupLabelsFinal, setGroupLabelsFinal] = useState<string[]>([]);
  const [tableValues, setTableValues] = useState<number[][] | null>(null);
  const [remarks, setRemarks] = useState<string[]>(REPORT_ROWS.map(() => ""));

  function addFiles(selected: FileList | null) {
    if (!selected) return;
    const incoming = Array.from(selected);
    setFiles((prev) => {
      const existingNames = new Set(prev.map((f) => f.name));
      return [...prev, ...incoming.filter((f) => !existingNames.has(f.name))];
    });
  }

  function removeFile(name: string) {
    setFiles((prev) => prev.filter((f) => f.name !== name));
  }

  async function handleGenerate() {
    if (files.length === 0) {
      setStatus("Upload at least one group roster file before generating.");
      setStatusTone("warning");
      return;
    }
    setIsGenerating(true);
    setStatus("Parsing files...");
    setStatusTone("idle");
    setWarnings([]);

    const keywords = irregularKeywords.split(",").map((k) => k.trim()).filter(Boolean);
    const allWarnings: string[] = [];
    const totalsPerGroup: GroupTotals[] = [];
    const rawLabels: string[] = [];

    for (const file of files) {
      const { groups, warnings: w } = await parseRosterWorkbook(file, keywords);
      allWarnings.push(...w);
      for (const group of groups) {
        rawLabels.push(group.label);
        totalsPerGroup.push(computeGroupTotals(group.rows));
        allWarnings.push(...checkFileTotals(group));
      }
    }

    if (attendanceFile) {
      const parsedAttendance = await parseAttendanceWorkbook(attendanceFile);
      allWarnings.push(...parsedAttendance.warnings);
      if (parsedAttendance.weekend.length > 0) {
        setWeekend(parsedAttendance.weekend.map((w) => String(w.attendance)));
      }
      if (parsedAttendance.midweek.length > 0) {
        setMidweek(parsedAttendance.midweek.map((w) => String(w.attendance)));
      }
      if (parsedAttendance.weekend.length > 0 || parsedAttendance.midweek.length > 0) {
        setIncludeAttendance(true);
      }
      if (parsedAttendance.monthLabel && !monthLabel) {
        setMonthLabel(parsedAttendance.monthLabel);
      }
    }

    const labels = dedupeLabels(rawLabels);
    const values = REPORT_ROWS.map((rowDef) => totalsPerGroup.map((t) => t[rowDef.key]));

    setGroupLabelsFinal(labels);
    setTableValues(values);
    setRemarks(REPORT_ROWS.map(() => ""));
    setWarnings(allWarnings);
    setStatus(allWarnings.length ? "Done, with warnings — check below." : "Done. Review the preview below.");
    setStatusTone(allWarnings.length ? "warning" : "success");
    setIsGenerating(false);
  }

  function updateCell(rowIndex: number, groupIndex: number, value: string) {
    if (!tableValues) return;
    const n = Number(value);
    const next = tableValues.map((row) => [...row]);
    next[rowIndex][groupIndex] = Number.isFinite(n) ? n : 0;
    setTableValues(next);
  }

  function updateRemark(rowIndex: number, value: string) {
    setRemarks((r) => r.map((x, i) => (i === rowIndex ? value : x)));
  }

  function rowTotal(rowIndex: number): number {
    if (!tableValues) return 0;
    return tableValues[rowIndex].reduce((a, b) => a + (b || 0), 0);
  }

  const weekendValid = weekend.map((v) => (v === "" ? 0 : Number(v)));
  const midweekValid = midweek.map((v) => (v === "" ? 0 : Number(v)));
  const weekendWeeks = weekend.filter((v) => v !== "").length;
  const midweekWeeks = midweek.filter((v) => v !== "").length;
  const totalWeekend = weekendValid.reduce((a, b) => a + b, 0);
  const totalMidweek = midweekValid.reduce((a, b) => a + b, 0);
  const avgWeekend = weekendWeeks ? Math.round(totalWeekend / weekendWeeks) : 0;
  const avgMidweek = midweekWeeks ? Math.round(totalMidweek / midweekWeeks) : 0;
  const hasAttendance = includeAttendance && (weekendWeeks > 0 || midweekWeeks > 0);

  async function handleDownload() {
    if (!tableValues) return;
    setIsDownloading(true);
    setStatus("Building workbook...");
    setStatusTone("idle");
    const buffer = await buildWorkbook({
      monthLabel: monthLabel || "MONTH",
      groupLabels: groupLabelsFinal,
      tableValues,
      remarks,
      attendance: hasAttendance ? { weekend: weekendValid, midweek: midweekValid } : null,
    });
    const blob = new Blob([buffer], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `Monthly_Report_${(monthLabel || "report").replace(/\s+/g, "_")}.xlsx`;
    a.click();
    URL.revokeObjectURL(url);
    setStatus("Downloaded.");
    setStatusTone("success");
    setIsDownloading(false);
  }

  const statusColor =
    statusTone === "success"
      ? "text-emerald-600 dark:text-emerald-500"
      : statusTone === "warning"
        ? "text-amber-600 dark:text-amber-500"
        : "text-zinc-500";

  return (
    <div className="flex-1 bg-[radial-gradient(ellipse_80%_60%_at_50%_-20%,rgba(8,145,178,0.12),transparent)]">
      <div className="mx-auto max-w-4xl px-5 py-10 flex flex-col gap-6">
        <header className="flex flex-col gap-3">
          <div className="flex items-center gap-3">
            <div className="flex size-10 shrink-0 items-center justify-center rounded-xl bg-cyan-600 text-white shadow-sm">
              <svg viewBox="0 0 24 24" fill="none" className="size-5">
                <path
                  d="M4 19V6a1 1 0 0 1 1-1h9l6 6v8a1 1 0 0 1-1 1H5a1 1 0 0 1-1-1Z"
                  stroke="currentColor"
                  strokeWidth="1.6"
                  strokeLinejoin="round"
                />
                <path d="M14 5v5a1 1 0 0 0 1 1h5M8 13h8M8 16.5h5" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" />
              </svg>
            </div>
            <div>
              <h1 className="text-2xl font-semibold tracking-tight">Monthly Congregational Report Generator</h1>
              <p className="text-zinc-500 text-sm mt-0.5">
                Upload each group&apos;s field service roster and get the finished monthly report.
              </p>
            </div>
          </div>
          <span className="inline-flex w-fit items-center gap-1.5 rounded-full border border-emerald-300/70 bg-emerald-50 px-3 py-1 text-xs font-medium text-emerald-700 dark:border-emerald-800 dark:bg-emerald-950/40 dark:text-emerald-400">
            <svg viewBox="0 0 24 24" fill="none" className="size-3.5">
              <path
                d="M12 3 5 6v5c0 4.5 3 7.5 7 9 4-1.5 7-4.5 7-9V6l-7-3Z"
                stroke="currentColor"
                strokeWidth="1.6"
                strokeLinejoin="round"
              />
            </svg>
            Runs entirely in your browser — files are never uploaded anywhere
          </span>
        </header>

        <section className="rounded-xl border border-zinc-200 dark:border-zinc-800 bg-white/60 dark:bg-zinc-900/40 p-5 flex flex-col gap-3 shadow-sm">
          <span className="text-xs font-semibold uppercase tracking-wide text-cyan-700 dark:text-cyan-500">Step 1 · Report details</span>
          <label htmlFor={monthId} className="flex flex-col gap-1 text-sm max-w-xs">
            <span className="font-medium">Report month</span>
            <input
              id={monthId}
              className="border border-zinc-300 dark:border-zinc-700 bg-transparent rounded-md px-3 py-2 transition-colors focus:border-cyan-600 placeholder:text-zinc-400 dark:placeholder:text-zinc-500"
              value={monthLabel}
              onChange={(e) => setMonthLabel(e.target.value)}
              placeholder="APRIL 2026"
            />
          </label>
        </section>

        <section className="rounded-xl border border-zinc-200 dark:border-zinc-800 bg-white/60 dark:bg-zinc-900/40 p-5 flex flex-col gap-3 shadow-sm">
          <span className="text-xs font-semibold uppercase tracking-wide text-cyan-700 dark:text-cyan-500">Step 2 · Group rosters</span>
          <p className="text-zinc-500 text-sm">
            Select all of your group files at once (one per group). Each group becomes a column in the
            report, named after its file.
          </p>

          <label className="flex flex-col items-center justify-center gap-2 rounded-lg border-2 border-dashed border-zinc-300 dark:border-zinc-700 px-4 py-8 text-center cursor-pointer transition-colors hover:border-cyan-500 hover:bg-cyan-50/50 dark:hover:bg-cyan-950/20">
            <svg viewBox="0 0 24 24" fill="none" className="size-6 text-zinc-400">
              <path d="M12 16V4m0 0 4 4m-4-4-4 4" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round" />
              <path d="M4 16v3a1 1 0 0 0 1 1h14a1 1 0 0 0 1-1v-3" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round" />
            </svg>
            <span className="text-sm">
              <span className="font-medium text-cyan-700 dark:text-cyan-500">Click to browse</span>{" "}
              <span className="text-zinc-500">for .xlsx roster files</span>
            </span>
            <input
              type="file"
              accept=".xlsx"
              multiple
              onChange={(e) => {
                addFiles(e.target.files);
                e.target.value = "";
              }}
              className="sr-only"
            />
          </label>

          {files.length > 0 && (
            <ul className="flex flex-col gap-1 mt-1">
              {files.map((f) => (
                <li
                  key={f.name}
                  className="flex items-center justify-between gap-3 text-sm rounded-md border border-zinc-200 dark:border-zinc-800 px-3 py-2"
                >
                  <span className="flex items-center gap-2 min-w-0">
                    <FileIcon />
                    <span className="truncate">{labelFromFileName(f.name)}</span>
                  </span>
                  <button
                    type="button"
                    onClick={() => removeFile(f.name)}
                    className="shrink-0 text-xs px-2 py-0.5 rounded-md border border-rose-300 text-rose-500 transition-colors hover:bg-rose-50 dark:border-rose-800 dark:hover:bg-rose-950/40"
                  >
                    Remove
                  </button>
                </li>
              ))}
            </ul>
          )}
        </section>

        <section className="rounded-xl border border-zinc-200 dark:border-zinc-800 bg-white/60 dark:bg-zinc-900/40 p-5 flex flex-col gap-3 shadow-sm">
          <span className="text-xs font-semibold uppercase tracking-wide text-cyan-700 dark:text-cyan-500">
            Step 3 · Meeting attendance (optional)
          </span>

          <label htmlFor={attendanceId} className="flex items-center gap-2 text-sm">
            <input
              id={attendanceId}
              type="checkbox"
              checked={includeAttendance}
              onChange={(e) => setIncludeAttendance(e.target.checked)}
              className="size-4 accent-cyan-600"
            />
            <span>Include meeting attendance table</span>
          </label>

          {includeAttendance && (
            <>
              <p className="text-zinc-500 text-sm">
                Upload a meeting attendance workbook (with &quot;Week&quot;/&quot;Attendance&quot; columns and
                &quot;Midweek meeting&quot;/&quot;Weekend meeting&quot; sections) to auto-fill the weekly figures
                below, or type them in by hand.
              </p>

              {attendanceFile ? (
                <div className="flex items-center justify-between gap-3 text-sm rounded-md border border-zinc-200 dark:border-zinc-800 px-3 py-2 max-w-md">
                  <span className="flex items-center gap-2 min-w-0">
                    <FileIcon />
                    <span className="truncate">{labelFromFileName(attendanceFile.name)}</span>
                  </span>
                  <button
                    type="button"
                    onClick={() => setAttendanceFile(null)}
                    className="shrink-0 text-xs px-2 py-0.5 rounded-md border border-rose-300 text-rose-500 transition-colors hover:bg-rose-50 dark:border-rose-800 dark:hover:bg-rose-950/40"
                  >
                    Remove
                  </button>
                </div>
              ) : (
                <label className="flex items-center justify-center gap-2 rounded-lg border-2 border-dashed border-zinc-300 dark:border-zinc-700 px-4 py-4 text-center cursor-pointer transition-colors hover:border-cyan-500 hover:bg-cyan-50/50 dark:hover:bg-cyan-950/20 max-w-md">
                  <span className="text-sm">
                    <span className="font-medium text-cyan-700 dark:text-cyan-500">Click to browse</span>{" "}
                    <span className="text-zinc-500">for an attendance .xlsx file</span>
                  </span>
                  <input
                    type="file"
                    accept=".xlsx"
                    onChange={(e) => {
                      const f = e.target.files?.[0] ?? null;
                      setAttendanceFile(f);
                      e.target.value = "";
                    }}
                    className="sr-only"
                  />
                </label>
              )}

              <div className="grid grid-cols-2 gap-6 max-w-md">
                <div className="flex flex-col gap-2">
                  <h3 className="text-sm font-medium">Weekend</h3>
                  {weekend.map((v, i) => (
                    <input
                      key={i}
                      type="number"
                      aria-label={`Weekend attendance week ${i + 1}`}
                      className="border border-zinc-300 dark:border-zinc-700 bg-transparent rounded-md px-2 py-1 text-sm transition-colors focus:border-cyan-600 placeholder:text-zinc-400 dark:placeholder:text-zinc-500"
                      placeholder={`Week ${i + 1}`}
                      value={v}
                      onChange={(e) => setWeekend((arr) => arr.map((x, j) => (j === i ? e.target.value : x)))}
                    />
                  ))}
                </div>
                <div className="flex flex-col gap-2">
                  <h3 className="text-sm font-medium">Midweek</h3>
                  {midweek.map((v, i) => (
                    <input
                      key={i}
                      type="number"
                      aria-label={`Midweek attendance week ${i + 1}`}
                      className="border border-zinc-300 dark:border-zinc-700 bg-transparent rounded-md px-2 py-1 text-sm transition-colors focus:border-cyan-600 placeholder:text-zinc-400 dark:placeholder:text-zinc-500"
                      placeholder={`Week ${i + 1}`}
                      value={v}
                      onChange={(e) => setMidweek((arr) => arr.map((x, j) => (j === i ? e.target.value : x)))}
                    />
                  ))}
                </div>
              </div>
            </>
          )}
        </section>

        <details className="group rounded-xl border border-zinc-200 dark:border-zinc-800 bg-white/60 dark:bg-zinc-900/40 p-5 shadow-sm">
          <summary className="font-medium cursor-pointer list-none flex items-center gap-2">
            <svg viewBox="0 0 24 24" fill="none" className="size-4 text-zinc-400 transition-transform group-open:rotate-90">
              <path d="M9 5l7 7-7 7" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round" />
            </svg>
            Advanced settings
          </summary>
          <div className="mt-3 flex flex-col gap-3 pl-6">
            <p className="text-zinc-500 text-sm">
              A person counts as an <b>Auxiliary Pioneer</b> if the Aux. column has a mark; a{" "}
              <b>Regular Pioneer</b> if Hours is filled in and Aux. is not marked; otherwise a regular
              Publisher. Adjust the irregular keyword list below if your notes use different codes.
            </p>
            <label htmlFor={keywordsId} className="flex flex-col gap-1 text-sm max-w-md">
              <span className="text-zinc-500">Irregular keyword(s) in Notes (comma separated)</span>
              <input
                id={keywordsId}
                className="border border-zinc-300 dark:border-zinc-700 bg-transparent rounded-md px-3 py-2 transition-colors focus:border-cyan-600"
                value={irregularKeywords}
                onChange={(e) => setIrregularKeywords(e.target.value)}
              />
            </label>
          </div>
        </details>

        <section className="flex flex-wrap items-center gap-4">
          <button
            type="button"
            onClick={handleGenerate}
            disabled={files.length === 0 || isGenerating}
            className="inline-flex items-center gap-2 px-5 py-2.5 rounded-lg bg-cyan-600 text-white font-medium transition-colors hover:bg-cyan-700 disabled:cursor-not-allowed disabled:bg-zinc-300 disabled:text-zinc-500 dark:disabled:bg-zinc-800 dark:disabled:text-zinc-600"
          >
            {isGenerating && <Spinner />}
            {isGenerating ? "Generating…" : "Generate report"}
          </button>
          <span className={`text-sm ${statusColor}`} aria-live="polite">
            {status}
          </span>
        </section>

        {warnings.length > 0 && (
          <section className="rounded-xl border border-amber-300 bg-amber-50 dark:bg-amber-950/30 p-4 text-sm text-amber-700 dark:text-amber-300">
            <ul className="list-disc pl-5 space-y-0.5">
              {warnings.map((w, i) => (
                <li key={i}>{w}</li>
              ))}
            </ul>
          </section>
        )}

        {tableValues && (
          <section className="rounded-xl border border-zinc-200 dark:border-zinc-800 bg-white/60 dark:bg-zinc-900/40 p-5 flex flex-col gap-4 shadow-sm">
            <div>
              <h2 className="font-medium">Preview &amp; edit before download</h2>
              <p className="text-zinc-500 text-sm">Every number below is editable — correct anything before downloading.</p>
            </div>

            <div className="overflow-x-auto rounded-lg border border-zinc-200 dark:border-zinc-800">
              <table className="min-w-full border-collapse text-sm">
                <caption className="text-left font-semibold text-cyan-700 dark:text-cyan-500 px-3 py-2 bg-cyan-50/60 dark:bg-cyan-950/20">
                  MONTHLY CONGREGATIONAL REPORT — {monthLabel || "MONTH"}
                </caption>
                <thead>
                  <tr className="bg-zinc-50 dark:bg-zinc-900">
                    <th className="border-b border-zinc-200 dark:border-zinc-800 px-3 py-2 text-left text-cyan-700 dark:text-cyan-500">GROUPS</th>
                    {groupLabelsFinal.map((label) => (
                      <th key={label} className="border-b border-zinc-200 dark:border-zinc-800 px-3 py-2 text-cyan-700 dark:text-cyan-500">
                        {label}
                      </th>
                    ))}
                    <th className="border-b border-zinc-200 dark:border-zinc-800 px-3 py-2 text-rose-500">TOTALS</th>
                    <th className="border-b border-zinc-200 dark:border-zinc-800 px-3 py-2 text-left text-cyan-700 dark:text-cyan-500">REMARKS</th>
                  </tr>
                </thead>
                <tbody>
                  {REPORT_ROWS.map((rowDef, rowIndex) => (
                    <tr key={rowDef.key} className="odd:bg-transparent even:bg-zinc-50/60 dark:even:bg-zinc-900/40">
                      <td className="border-b border-zinc-200 dark:border-zinc-800 px-3 py-2 font-semibold text-rose-500 whitespace-nowrap">
                        {rowDef.label}
                      </td>
                      {groupLabelsFinal.map((_, groupIndex) => (
                        <td key={groupIndex} className="border-b border-zinc-200 dark:border-zinc-800 p-0">
                          <input
                            type="number"
                            aria-label={`${rowDef.label} — ${groupLabelsFinal[groupIndex]}`}
                            className="w-full text-center py-2 bg-transparent text-emerald-600 font-semibold transition-colors focus:bg-cyan-50/60 dark:focus:bg-cyan-950/30"
                            value={tableValues[rowIndex][groupIndex]}
                            onChange={(e) => updateCell(rowIndex, groupIndex, e.target.value)}
                          />
                        </td>
                      ))}
                      <td className="border-b border-zinc-200 dark:border-zinc-800 px-3 py-2 text-center font-bold text-rose-500">
                        {rowTotal(rowIndex)}
                      </td>
                      <td className="border-b border-zinc-200 dark:border-zinc-800 p-0">
                        <input
                          type="text"
                          aria-label={`Remarks — ${rowDef.label}`}
                          className="w-full text-left px-2 py-2 bg-transparent transition-colors focus:bg-cyan-50/60 dark:focus:bg-cyan-950/30"
                          value={remarks[rowIndex]}
                          onChange={(e) => updateRemark(rowIndex, e.target.value)}
                        />
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>

            {hasAttendance && (
              <div className="overflow-x-auto rounded-lg border border-zinc-200 dark:border-zinc-800">
                <table className="border-collapse text-sm w-full">
                  <tbody>
                    <tr>
                      <td rowSpan={2} className="border-b border-zinc-200 dark:border-zinc-800 px-3 py-2 font-semibold text-cyan-700 dark:text-cyan-500">
                        MEETING ATTENDANCE
                      </td>
                      <td className="border-b border-zinc-200 dark:border-zinc-800 px-3 py-2">TOTAL</td>
                      <td className="border-b border-zinc-200 dark:border-zinc-800 px-3 py-2 text-center">{totalWeekend}</td>
                      <td className="border-b border-zinc-200 dark:border-zinc-800 px-3 py-2 text-center">÷</td>
                      <td className="border-b border-zinc-200 dark:border-zinc-800 px-3 py-2 text-center">{weekendWeeks}</td>
                      <td className="border-b border-zinc-200 dark:border-zinc-800 px-3 py-2 text-center">=</td>
                      <td className="border-b border-zinc-200 dark:border-zinc-800 px-3 py-2">AVERAGE</td>
                      <td className="border-b border-zinc-200 dark:border-zinc-800 px-3 py-2 text-center font-bold text-rose-500">{avgWeekend}</td>
                      <td className="border-b border-zinc-200 dark:border-zinc-800 px-3 py-2">WEEKEND</td>
                    </tr>
                    <tr>
                      <td className="px-3 py-2"></td>
                      <td className="px-3 py-2 text-center">{totalMidweek}</td>
                      <td className="px-3 py-2 text-center">÷</td>
                      <td className="px-3 py-2 text-center">{midweekWeeks}</td>
                      <td className="px-3 py-2 text-center">=</td>
                      <td className="px-3 py-2">AVERAGE</td>
                      <td className="px-3 py-2 text-center font-bold text-rose-500">{avgMidweek}</td>
                      <td className="px-3 py-2">MIDWEEK</td>
                    </tr>
                  </tbody>
                </table>
              </div>
            )}

            <button
              type="button"
              onClick={handleDownload}
              disabled={isDownloading}
              className="self-start inline-flex items-center gap-2 px-5 py-2.5 rounded-lg bg-cyan-600 text-white font-medium transition-colors hover:bg-cyan-700 disabled:cursor-not-allowed disabled:bg-zinc-300 disabled:text-zinc-500 dark:disabled:bg-zinc-800 dark:disabled:text-zinc-600"
            >
              {isDownloading && <Spinner />}
              {isDownloading ? "Building…" : "Download Excel (.xlsx)"}
            </button>
          </section>
        )}

        <footer className="text-xs text-zinc-400 text-center mt-4">
          Nothing you upload leaves this browser tab — parsing and report generation happen entirely on your device.
        </footer>
      </div>
    </div>
  );
}
