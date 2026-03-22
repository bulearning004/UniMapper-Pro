"use client";

import { useState } from "react";
import { Upload, FileDown, AlertTriangle, LayoutDashboard, Settings, Bell, Search } from "lucide-react";
import * as XLSX from "xlsx";
import { Document, Packer, Paragraph, Table, TableCell, TableRow, WidthType, TextRun, AlignmentType, BorderStyle, PageOrientation } from "docx";
import { saveAs } from "file-saver";

// --- Types ---
type SlotType = 1 | 2 | 3 | 4 | "Review Needed";

interface ParsedRow {
  courseCode: string;
  section: string;
  day: string;
  time: string;
  room: string;
  slot: SlotType;
}

// --- Logic: Robust Time Matching ---
function mapToFixedSlot(inputTime: string): SlotType {
  if (!inputTime) return "Review Needed";
  const cleanTime = inputTime.replace(/:/g, ".").replace(/\s/g, "");
  if (cleanTime.includes("8.40") || cleanTime.includes("11.00")) return 1;
  if (cleanTime.includes("12.00") || cleanTime.includes("14.20")) return 2;
  if (cleanTime.includes("14.30") || cleanTime.includes("16.50")) return 3;
  if (cleanTime.includes("17.00") || cleanTime.includes("19.20")) return 4;
  return "Review Needed";
}

const DAYS = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"];
const SLOTS = [
  { id: 1, label: "8.40-11.00" },
  { id: 2, label: "12.00-14.20" },
  { id: 3, label: "14.30-16.50" },
  { id: 4, label: "17.00-19.20" },
];

export default function ScheduleMapper() {
  const [mappedData, setMappedData] = useState<ParsedRow[]>([]);
  const [scheduleName, setScheduleName] = useState("");

  // --- File Upload & Parsing ---
  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      const bstr = evt.target?.result;
      const wb = XLSX.read(bstr, { type: "binary" });
      const wsname = wb.SheetNames[0];
      const ws = wb.Sheets[wsname];
      const data = XLSX.utils.sheet_to_json<any>(ws, { header: 1 });

      const parsed: ParsedRow[] = [];
      let currentCourseCode = "Unknown";
      let roomColIdx = -1;
      let sectionColIdx = -1;

      for (let i = 0; i < data.length; i++) {
        const row = data[i];
        if (!row || row.length === 0) continue;

        const isHeaderRow = row.some((c: any) => typeof c === 'string' && c.includes("Section")) && 
                            row.some((c: any) => typeof c === 'string' && c.includes("Time"));
        if (isHeaderRow) {
          roomColIdx = row.findIndex((c: any) => typeof c === 'string' && c.includes("Room"));
          sectionColIdx = row.findIndex((c: any) => typeof c === 'string' && c.includes("Section"));
          continue;
        }

        if (typeof row[0] === "string" && row[0].trim() !== "" && !row.includes("Time") && !row.includes("Section")) {
          const hasTime = row.some((c: any) => typeof c === 'string' && /\d{1,2}[.:]\d{2}/.test(c));
          if (!hasTime && row.length <= 5) {
            currentCourseCode = row[0].trim();
          }
        }

        let section = sectionColIdx !== -1 && row[sectionColIdx] ? String(row[sectionColIdx]).trim() : String(row[0] || "").trim();
        let day = "";
        let time = "";
        let room = roomColIdx !== -1 && row[roomColIdx] ? String(row[roomColIdx]).trim() : "";

        for (let j = 0; j < row.length; j++) {
          const cell = String(row[j] || "").trim();
          if (DAYS.includes(cell)) day = cell;
          if (cell.includes(".") || cell.includes("-") || cell.includes(":")) {
            if (/\d{1,2}[.:]\d{2}/.test(cell)) {
              time = cell;
            }
          }
          if (roomColIdx === -1 && !room && j > 0 && j !== sectionColIdx && !DAYS.includes(cell) && !/\d{1,2}[.:]\d{2}/.test(cell)) {
            if (cell.length >= 3 && cell.length <= 10 && !cell.includes("Section") && !cell.includes("Time")) {
              room = cell;
            }
          }
        }

        if (section && day && time && section !== "Section" && day !== "Day" && time !== "Time") {
          parsed.push({
            courseCode: currentCourseCode,
            section,
            day,
            time,
            room,
            slot: mapToFixedSlot(time),
          });
        }
      }

      setMappedData(parsed);
      setScheduleName(file.name.split('.').slice(0, -1).join('.'));
    };
    reader.readAsBinaryString(file);
  };

  // --- Word Export ---
  const handleExportWord = async (dataToExport: ParsedRow[] = mappedData, name: string = scheduleName) => {
    if (dataToExport.length === 0) return;

    const tableRows = [];

    tableRows.push(
      new TableRow({
        children: [
          new TableCell({
            width: { size: 15, type: WidthType.PERCENTAGE },
            children: [new Paragraph({ children: [new TextRun({ text: "Day / Time", bold: true })], alignment: AlignmentType.CENTER })],
          }),
          ...SLOTS.map(
            (slot) =>
              new TableCell({
                width: { size: 21.25, type: WidthType.PERCENTAGE },
                children: [new Paragraph({ children: [new TextRun({ text: slot.label, bold: true })], alignment: AlignmentType.CENTER })],
              })
          ),
        ],
      })
    );

    DAYS.forEach((day) => {
      const rowCells = [
        new TableCell({
          width: { size: 15, type: WidthType.PERCENTAGE },
          children: [new Paragraph({ children: [new TextRun({ text: day, bold: true })], alignment: AlignmentType.CENTER })],
        }),
      ];

      SLOTS.forEach((slot) => {
        const classesInSlot = dataToExport.filter((d) => d.day === day && d.slot === slot.id);
        
        const cellParagraphs = classesInSlot.map((c) => {
          return new Paragraph({
            children: [
              new TextRun({ text: c.courseCode, size: 20, bold: true }),
              new TextRun({ text: `\nSect: ${c.section}   ${c.room || ""}`, size: 20, break: 1 }),
            ],
            alignment: AlignmentType.CENTER,
          });
        });

        rowCells.push(
          new TableCell({
            width: { size: 21.25, type: WidthType.PERCENTAGE },
            children: cellParagraphs.length > 0 ? cellParagraphs : [new Paragraph("")],
          })
        );
      });

      tableRows.push(new TableRow({ children: rowCells }));
    });

    const table = new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: tableRows,
      borders: {
        top: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
        bottom: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
        left: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
        right: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
        insideHorizontal: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
        insideVertical: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
      },
    });

    const doc = new Document({
      sections: [
        {
          properties: {
            page: {
              size: {
                orientation: PageOrientation.LANDSCAPE,
              },
            },
          },
          children: [
            new Paragraph({
              children: [new TextRun({ text: name || "Schedule", bold: true, size: 28 })],
              alignment: AlignmentType.CENTER,
              spacing: { after: 400 },
            }),
            table,
          ],
        },
      ],
    });

    const blob = await Packer.toBlob(doc);
    saveAs(blob, `${name || "Schedule"}.docx`);
  };

  return (
    <div className="min-h-screen bg-gradient-to-br from-[#d4e6e6] via-[#e2e8f0] to-[#d1d5eb] p-4 md:p-8 font-sans flex items-center justify-center">
      <div className="w-full max-w-[1400px] bg-white/40 backdrop-blur-2xl rounded-[2.5rem] shadow-[0_20px_50px_rgba(0,0,0,0.05)] border border-white/50 flex flex-col md:flex-row overflow-hidden min-h-[800px]">
        
        {/* Sidebar */}
        <div className="w-full md:w-64 p-8 flex flex-col gap-8 border-b md:border-b-0 md:border-r border-white/30 shrink-0 bg-white/20">
          <div className="text-2xl font-extrabold text-[#1e3a3a] tracking-tight">UniMapper Pro</div>
          <nav className="flex flex-col gap-2 flex-1">
            <button className="flex items-center gap-4 px-4 py-3.5 bg-gradient-to-r from-[#0d7a7a] to-[#149b9b] text-white rounded-2xl shadow-lg shadow-teal-500/30 font-semibold">
              <LayoutDashboard className="w-5 h-5" /> Dashboard
            </button>
          </nav>
          <div className="mt-auto hidden md:block">
            <button className="flex items-center gap-4 px-4 py-3.5 text-[#5a7c7c] hover:bg-white/40 rounded-2xl font-semibold transition-colors w-full">
              <Settings className="w-5 h-5" /> Settings
            </button>
          </div>
        </div>

        {/* Main Content */}
        <div className="flex-1 p-6 md:p-10 flex flex-col gap-8 overflow-y-auto custom-scrollbar">
          
          {/* Header */}
          <div className="flex flex-col md:flex-row justify-between items-start md:items-end gap-4">
            <div>
              <p className="text-[#0d7a7a] font-semibold mb-1 flex items-center gap-2">
                Welcome <span className="text-xl">👋</span>
              </p>
              <h1 className="text-4xl font-extrabold text-[#1e3a3a] tracking-tight">Dashboard</h1>
            </div>
            <div className="flex items-center gap-5 self-end md:self-auto">
              <button className="p-2.5 text-[#5a7c7c] hover:bg-white/40 rounded-full transition-colors hidden sm:block">
                <Search className="w-5 h-5" />
              </button>
              <button className="p-2.5 text-[#5a7c7c] hover:bg-white/40 rounded-full transition-colors relative hidden sm:block">
                <Bell className="w-5 h-5" />
                <span className="absolute top-2 right-2.5 w-2 h-2 bg-red-500 rounded-full border-2 border-[#dbe2eb]"></span>
              </button>
            </div>
          </div>

          {/* Top Grid */}
          <div className="grid grid-cols-1 gap-6">
            {/* Upload Card */}
            <div className="bg-white/60 backdrop-blur-xl border border-white/60 rounded-[2rem] p-6 shadow-[0_8px_30px_rgb(0,0,0,0.04)] flex flex-col min-h-[220px]">
              <div className="flex justify-between items-center mb-4">
                <h3 className="text-[#5a7c7c] font-semibold">Upload Schedule</h3>
              </div>
              <label className="flex-1 flex flex-col items-center justify-center border-2 border-[#0d7a7a]/20 border-dashed rounded-2xl cursor-pointer bg-white/40 hover:bg-white/60 transition-all group">
                <div className="flex flex-col items-center justify-center p-6 text-center">
                  <div className="w-12 h-12 bg-gradient-to-br from-[#e0f0f0] to-[#c5dada] rounded-full flex items-center justify-center mb-3 text-[#0d7a7a] group-hover:scale-110 transition-transform shadow-sm">
                    <Upload className="w-5 h-5" />
                  </div>
                  <p className="text-sm text-[#1e3a3a] font-bold mb-1">Click to upload</p>
                  <p className="text-xs text-[#5a7c7c] font-medium">or drag Excel file here</p>
                </div>
                <input type="file" className="hidden" accept=".xlsx, .xls" onChange={handleFileUpload} />
              </label>
            </div>
          </div>

          {/* Table Card */}
          <div className="bg-white/60 backdrop-blur-xl border border-white/60 rounded-[2rem] p-6 shadow-[0_8px_30px_rgb(0,0,0,0.04)] flex-1 flex flex-col">
            <div className="flex flex-col sm:flex-row items-start sm:items-center justify-between mb-6 gap-4">
              <h2 className="text-xl font-extrabold text-[#1e3a3a]">
                {scheduleName ? scheduleName : "Mapped Schedule"}
              </h2>
              <div className="flex gap-3 w-full sm:w-auto">
                <button 
                  onClick={() => handleExportWord()}
                  disabled={mappedData.length === 0}
                  className="flex-1 sm:flex-none flex items-center justify-center gap-2 px-5 py-2.5 text-sm font-bold text-white bg-gradient-to-r from-[#0d7a7a] to-[#149b9b] rounded-xl hover:shadow-lg hover:shadow-teal-500/30 transition-all disabled:opacity-50 disabled:shadow-none"
                >
                  <FileDown className="w-4 h-4" /> Export Word
                </button>
              </div>
            </div>

            {mappedData.length === 0 ? (
              <div className="flex-1 flex flex-col items-center justify-center text-[#5a7c7c] min-h-[200px] bg-white/30 rounded-2xl border border-white/40 border-dashed">
                <p className="font-medium">Upload an Excel file to see the mapped schedule.</p>
              </div>
            ) : (
              <div className="overflow-x-auto custom-scrollbar pb-4">
                <table className="w-full text-sm text-left border-collapse min-w-[800px]">
                  <thead>
                    <tr className="text-[#5a7c7c] border-b-2 border-[#5a7c7c]/20">
                      <th className="pb-4 font-semibold uppercase tracking-wider text-xs w-[12%] pl-2">Day</th>
                      {SLOTS.map(slot => (
                        <th key={slot.id} className="pb-4 font-semibold uppercase tracking-wider text-xs w-[22%]">
                          {slot.label}
                        </th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {DAYS.map((day, rowIdx) => (
                      <tr key={day} className="border-b-2 border-[#5a7c7c]/30 last:border-0">
                        <td className="py-5 pl-2 font-extrabold text-[#1e3a3a] text-base">{day}</td>
                        {SLOTS.map(slot => {
                          const classes = mappedData.filter(d => d.day === day && d.slot === slot.id);
                          return (
                            <td key={slot.id} className="py-4 pr-4 align-top">
                              {classes.length > 0 ? (
                                <div className="space-y-3">
                                  {classes.map((c, idx) => (
                                    <div key={idx} className="bg-white/80 p-3.5 rounded-2xl shadow-sm border border-white/60 flex flex-col gap-2 hover:shadow-md transition-shadow relative overflow-hidden group">
                                      <div className="absolute left-0 top-0 bottom-0 w-1 bg-gradient-to-b from-[#0d7a7a] to-[#149b9b]"></div>
                                      <div className="flex items-start gap-2 pl-1">
                                        <p className="font-extrabold text-[#1e3a3a] text-xs leading-tight">{c.courseCode}</p>
                                      </div>
                                      <div className="flex justify-between items-center pl-1 mt-1">
                                        <p className="text-[#5a7c7c] text-[11px] font-bold bg-[#f0f5f5] px-2 py-1 rounded-lg">Sect: {c.section}</p>
                                        {c.room && (
                                          <p className="text-[#0d7a7a] text-[11px] font-extrabold bg-[#e0f0f0] px-2 py-1 rounded-lg whitespace-nowrap">
                                            {c.room}
                                          </p>
                                        )}
                                      </div>
                                    </div>
                                  ))}
                                </div>
                              ) : null}
                            </td>
                          );
                        })}
                      </tr>
                    ))}
                  </tbody>
                </table>

                {/* Review Needed Section */}
                {mappedData.filter(d => d.slot === "Review Needed").length > 0 && (
                  <div className="mt-8 bg-white/60 rounded-2xl p-6 border border-white/60">
                    <h3 className="text-lg font-extrabold text-amber-600 flex items-center gap-2 mb-4">
                      <AlertTriangle className="w-5 h-5" /> Review Needed
                    </h3>
                    <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
                      {mappedData.filter(d => d.slot === "Review Needed").map((c, idx) => (
                        <div key={idx} className="bg-amber-50/80 p-4 rounded-2xl border border-amber-200/50 shadow-sm relative overflow-hidden">
                          <div className="absolute left-0 top-0 bottom-0 w-1 bg-amber-400"></div>
                          <p className="font-extrabold text-[#1e3a3a] leading-tight mb-3 pl-1">{c.courseCode}</p>
                          <div className="grid grid-cols-2 gap-y-2 gap-x-4 pl-1">
                            <p className="text-[#5a7c7c] text-xs font-semibold">Sect: <span className="text-[#1e3a3a]">{c.section}</span></p>
                            <p className="text-[#5a7c7c] text-xs font-semibold">Room: <span className="text-[#1e3a3a]">{c.room || "-"}</span></p>
                            <p className="text-[#5a7c7c] text-xs font-semibold">Day: <span className="text-[#1e3a3a]">{c.day}</span></p>
                            <p className="text-[#5a7c7c] text-xs font-semibold">Time: <span className="text-amber-600">{c.time}</span></p>
                          </div>
                        </div>
                      ))}
                    </div>
                  </div>
                )}
              </div>
            )}
          </div>
        </div>
      </div>
    </div>
  );
}
