import { useState, useRef, useCallback } from 'react';
import { Button } from '@/components/ui/button';
import { Card, CardContent, CardHeader, CardTitle } from '@/components/ui/card';
import { Alert, AlertDescription } from '@/components/ui/alert';
import { Progress } from '@/components/ui/progress';
import { 
  Upload, 
  Play, 
  Download, 
  RefreshCw, 
  FileSpreadsheet, 
  CheckCircle, 
  XCircle,
  AlertCircle,
  Trash2
} from 'lucide-react';
import * as XLSX from 'xlsx';
import './App.css';

// Definisi item ekstraksi berdasarkan Ekstrak_FDM_V5.py
interface ExtractItem {
  label: string;
  sheet: string;
  addr?: string;
  keyword?: string;
  mode: string;
}

const itemsDefinitions: ExtractItem[] = [
  { label: "KPP", sheet: "Sheet Home", addr: "H5", mode: "Static" },
  { label: "Sektor", sheet: "Sheet Home", addr: "D8", mode: "Static" },
  { label: "NAMA WAJIB PAJAK", sheet: "Sheet Home", addr: "H12", mode: "Static" },
  { label: "NOMOR OBJEK PAJAK", sheet: "Sheet Home", addr: "H14", mode: "Static" },
  { label: "KELURAHAN", sheet: "Sheet Home", addr: "H20", mode: "Static" },
  { label: "KECAMATAN", sheet: "Sheet Home", addr: "H22", mode: "Static" },
  { label: "KABUPATEN/KOTA", sheet: "Sheet Home", addr: "H24", mode: "Static" },
  { label: "PROVINSI", sheet: "Sheet Home", addr: "H26", mode: "Static" },
  { label: "LUAS BUMI", sheet: "Sheet Home", mode: "Formula_LuasBumi" },
  { label: "Areal Produktif", sheet: "Sheet Home", addr: "J73", mode: "Static" },
  { label: "Areal Belum Diolah", sheet: "Sheet Home", addr: "J75", mode: "Static" },
  { label: "Areal Sudah Diolah Belum Ditanami", sheet: "Sheet Home", addr: "J76", mode: "Static" },
  { label: "Areal Pembibitan", sheet: "Sheet Home", addr: "J77", mode: "Static" },
  { label: "Areal Tidak Produktif", sheet: "Sheet Home", addr: "J78", mode: "Static" },
  { label: "Areal Pengaman", sheet: "Sheet Home", addr: "J79", mode: "Static" },
  { label: "Areal Emplasemen", sheet: "Sheet Home", addr: "J80", mode: "Static" },
  { label: "Areal Produktif (Copy)", sheet: "Sheet Home", mode: "Formula_CopyProduktif" },
  { label: "NJOP/M Areal Belum Produktif", sheet: "C.1", addr: "BK23", mode: "Static" },
  { label: "NJOP Bumi Berupa Tanah (Rp)", sheet: "Sheet Home", mode: "Formula_NJOPTanah" },
  { label: "NJOP Bumi Berupa Pengembangan Tanah (Rp)", sheet: "C.2", keyword: "Pengembangan Tanah", mode: "Dynamic_Col_G" },
  { label: "NJOP Bumi Berupa Pengembangan Tanah (Rp) (Kenaikan BIT 10.3%)", sheet: "N/A", mode: "Formula_BIT" },
  { label: "NJOP Bumi Areal Produktif (Rp)", sheet: "N/A", mode: "Formula_NJOP_Total" },
  { label: "Luas Bumi Areal Produktif (m²)", sheet: "N/A", mode: "Formula_Luas_Ref" },
  { label: "NJOP Bumi Per M2 Areal Produktif (Rp/m2)", sheet: "N/A", mode: "Formula_NJOP_PerM2" },
  { label: "NJOP BUMI (Rp) AREA PRODUKTIF pada A. DATA BUMI", sheet: "N/A", mode: "Formula_Final_Calc" },
  { label: "NJOP BUMI (Rp) AREA PRODUKTIF pada A. DATA BUMI (Proyeksi NDT Naik 46%)", sheet: "N/A", mode: "Formula_Proyeksi" },
  { label: "NJOP BUMI (Rp) AREAL BELUM PRODUKTIF pada A. DATA BUMI", sheet: "FDM Kebun ABC", addr: "E19", mode: "Static" },
  { label: "NJOP BUMI (Rp) AREAL BELUM PRODUKTIF pada A. DATA BUMI (Proyeksi NDT Naik 46%)", sheet: "N/A", mode: "Formula_Proyeksi_BelumProd" },
  { label: "Areal Tidak Produktif (Copy)", sheet: "N/A", mode: "Formula_CopyTidakProduktif" },
  { label: "NJOP/M Areal Tidak Produktif", sheet: "C.1", addr: "BK64", mode: "Static" },
  { label: "NJOP BUMI (Rp) AREAL TIDAK PRODUKTIF pada A. DATA BUMI", sheet: "N/A", mode: "Formula_Calc_TidakProd" },
  { label: "NJOP BUMI (Rp) AREAL TIDAK PRODUKTIF pada A. DATA BUMI (Proyeksi NDT Naik 46%)", sheet: "N/A", mode: "Formula_Proyeksi_TidakProd" },
  { label: "Areal Pengaman (Copy)", sheet: "N/A", mode: "Formula_CopyPengaman" },
  { label: "NJOP/M Areal Pengaman", sheet: "D", addr: "L23", mode: "Static" },
  { label: "NJOP BUMI (Rp) AREAL PENGAMAN pada A. DATA BUMI", sheet: "N/A", mode: "Formula_Calc_Pengaman" },
  { label: "NJOP BUMI (Rp) AREAL PENGAMAN pada A. DATA BUMI (Proyeksi NDT Naik 46%)", sheet: "N/A", mode: "Formula_Proyeksi_Pengaman" },
  { label: "Areal Emplasemen (Copy)", sheet: "N/A", mode: "Formula_CopyEmplasemen" },
  { label: "NJOP/M Areal Emplasemen", sheet: "C.1", addr: "BK105", mode: "Static" },
  { label: "NJOP BUMI (Rp) AREAL EMPLASEMEN pada A. DATA BUMI", sheet: "N/A", mode: "Formula_Calc_Emplasemen" },
  { label: "NJOP BUMI (Rp) AREAL EMPLASEMEN pada A. DATA BUMI (Proyeksi NDT Naik 46%)", sheet: "N/A", mode: "Formula_Proyeksi_Emplasemen" },
  { label: "JUMLAH Luas (m2) pada A. DATA BUMI", sheet: "N/A", mode: "Formula_Total_Luas_Ref" },
  { label: "JUMLAH NJOP BUMI (Rp) pada A. DATA BUMI", sheet: "N/A", mode: "Formula_Total_NJOP_Sum" },
  { label: "NJOP BUMI (Rp) NJOP Bumi Per Meter Persegi pada A. DATA BUMI", sheet: "FDM Kebun ABC", addr: "E24", mode: "Static" },
  { label: "Jumlah LUAS pada B. DATA BANGUNAN", sheet: "FDM Kebun ABC", mode: "Dynamic_FDM_Bangunan_Luas" },
  { label: "Jumlah NJOP BANGUNAN pada B. DATA BANGUNAN", sheet: "N/A", mode: "Formula_Calc_Bangunan" },
  { label: "NJOP BANGUNAN PER METER PERSEGI*) pada B. DATA BANGUNAN", sheet: "FDM Kebun ABC", mode: "Dynamic_FDM_Bangunan_PerM2" },
  { label: "TOTAL NJOP (TANAH + BANGUNAN) 2025", sheet: "N/A", mode: "Formula_Grand_Total" },
  { label: "SPPT 2025", sheet: "N/A", mode: "Formula_SPPT_2025" },
  { label: "SIMULASI TOTAL NJOP (TANAH + BANGUNAN) 2026 (Hanya Kenaikan BIT 10,3% + NDT Tetap)", sheet: "N/A", mode: "Formula_Simulasi_NJOP_2026" },
  { label: "SIMULASI SPPT 2026 (Hanya Kenaikan BIT 10,3% + NDT Tetap)", sheet: "N/A", mode: "Formula_Simulasi_SPPT_2026" },
  { label: "Kenaikan", sheet: "N/A", mode: "Formula_Kenaikan" },
  { label: "Persentase", sheet: "N/A", mode: "Formula_Persentase" },
  { label: "SIMULASI TOTAL NJOP (TANAH + BANGUNAN) 2026 (Kenaikan BIT 10,3% + NDT 46%)", sheet: "N/A", mode: "Formula_Simulasi_Total_2026_NDT46" },
  { label: "SIMULASI SPPT 2026 (Kenaikan BIT 10,3% + NDT 46%)", sheet: "N/A", mode: "Formula_Simulasi_SPPT_2026_NDT46" },
];

interface UploadedFile {
  id: string;
  file: File;
  name: string;
  status: 'pending' | 'processing' | 'completed' | 'error';
  errorMessage?: string;
  extractedData?: Record<string, number | string>;
}

// Tipe untuk cell value - bisa number, string, atau rumus Excel
interface FormulaCell {
  f: string;
}

type CellValue = number | string | null | FormulaCell;

interface ExtractionResult {
  headers: string[];
  rows: CellValue[][];
}

// Fungsi untuk mengkonversi indeks kolom ke huruf kolom Excel (0 = A, 1 = B, dst)
const getColumnLetter = (index: number): string => {
  let result = '';
  let temp = index;
  while (temp >= 0) {
    result = String.fromCharCode(65 + (temp % 26)) + result;
    temp = Math.floor(temp / 26) - 1;
  }
  return result;
};

export default function App() {
  const [uploadedFiles, setUploadedFiles] = useState<UploadedFile[]>([]);
  const [isExtracting, setIsExtracting] = useState(false);
  const [extractionProgress, setExtractionProgress] = useState(0);
  const [extractionResult, setExtractionResult] = useState<ExtractionResult | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [successMessage, setSuccessMessage] = useState<string | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);

  // Fungsi untuk mendapatkan sheet dengan pencarian pintar
  const getSheetSmart = (wb: XLSX.WorkBook, nameHint: string): XLSX.WorkSheet | null => {
    if (nameHint === "N/A") return null;
    const nameHintLower = nameHint.toLowerCase();
    const sheetMap: Record<string, string> = {};
    wb.SheetNames.forEach((name: string) => {
      sheetMap[name.toLowerCase()] = name;
    });
    
    if (wb.SheetNames.includes(nameHint)) return wb.Sheets[nameHint];
    if (sheetMap[nameHintLower]) return wb.Sheets[sheetMap[nameHintLower]];
    
    for (const existingSheet of Object.keys(sheetMap)) {
      if (nameHintLower.includes("c.1") && existingSheet.includes("c.1")) return wb.Sheets[sheetMap[existingSheet]];
      if (nameHintLower.includes("c.2") && existingSheet.includes("c.2")) return wb.Sheets[sheetMap[existingSheet]];
      if (nameHintLower.includes("home") && existingSheet.includes("home")) return wb.Sheets[sheetMap[existingSheet]];
      if (nameHintLower.includes("fdm") && existingSheet.includes("fdm")) return wb.Sheets[sheetMap[existingSheet]];
      if ((nameHintLower === "d" || nameHintLower === "sheet d") && (existingSheet === "d" || existingSheet === "sheet d")) {
        return wb.Sheets[sheetMap[existingSheet]];
      }
    }
    return null;
  };

  // Fungsi untuk mendapatkan nilai sel
  const getCellValue = (ws: XLSX.WorkSheet, addr: string): number | string | null => {
    const cell = ws[addr];
    if (!cell) return null;
    return cell.v !== undefined ? cell.v : null;
  };

  // Fungsi untuk mencari anchor row di FDM Kebun ABC
  const findFDMAnchorRow = (ws: XLSX.WorkSheet): number | null => {
    const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
    for (let row = 20; row <= Math.min(150, range.e.r); row++) {
      for (let col = 0; col <= 4; col++) {
        const cellAddr = XLSX.utils.encode_cell({ r: row, c: col });
        const cell = ws[cellAddr];
        if (cell && cell.v && typeof cell.v === 'string') {
          if (cell.v.toUpperCase().includes("NJOP BANGUNAN PER METER PERSEGI")) {
            return row;
          }
        }
      }
    }
    return null;
  };

  // Fungsi untuk pencarian dinamis kolom G
  const findDynamicColG = (ws: XLSX.WorkSheet, keyword: string): number | string | null => {
    const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
    const keywordLower = keyword.toLowerCase();
    for (let row = 0; row <= Math.min(150, range.e.r); row++) {
      for (let col = 0; col <= 4; col++) {
        const cellAddr = XLSX.utils.encode_cell({ r: row, c: col });
        const cell = ws[cellAddr];
        if (cell && cell.v && typeof cell.v === 'string') {
          const cellText = cell.v.toLowerCase().replace(/\s+/g, ' ');
          if (cellText.includes(keywordLower)) {
            const colGAddr = XLSX.utils.encode_cell({ r: row, c: 6 });
            const colGCell = ws[colGAddr];
            return colGCell && colGCell.v !== undefined ? colGCell.v : null;
          }
        }
      }
    }
    return "TIDAK DITEMUKAN";
  };

  // Handle file upload
  const handleFileUpload = useCallback((event: React.ChangeEvent<HTMLInputElement>) => {
    const files = event.target.files;
    if (!files) return;

    const excelFiles = Array.from(files).filter(file => 
      file.name.endsWith('.xlsm') || file.name.endsWith('.xlsx') || file.name.endsWith('.xls')
    );

    if (excelFiles.length === 0) {
      setError('Harap upload file Excel (.xlsm, .xlsx, atau .xls)');
      return;
    }

    if (excelFiles.length > 50) {
      setError('Maksimal 50 file dapat diupload sekaligus');
      return;
    }

    const newFiles: UploadedFile[] = excelFiles.map((file, index) => ({
      id: `file-${Date.now()}-${index}`,
      file,
      name: file.name,
      status: 'pending' as const,
    }));

    setUploadedFiles(prev => [...prev, ...newFiles]);
    setError(null);
    setSuccessMessage(`${excelFiles.length} file berhasil ditambahkan`);
    setTimeout(() => setSuccessMessage(null), 3000);
  }, []);

  // Handle drag and drop
  const handleDrop = useCallback((event: React.DragEvent) => {
    event.preventDefault();
    const files = event.dataTransfer.files;
    if (!files) return;

    const excelFiles = Array.from(files).filter(file => 
      file.name.endsWith('.xlsm') || file.name.endsWith('.xlsx') || file.name.endsWith('.xls')
    );

    if (excelFiles.length === 0) {
      setError('Harap upload file Excel (.xlsm, .xlsx, atau .xls)');
      return;
    }

    if (uploadedFiles.length + excelFiles.length > 50) {
      setError('Total file tidak boleh lebih dari 50');
      return;
    }

    const newFiles: UploadedFile[] = excelFiles.map((file, index) => ({
      id: `file-${Date.now()}-${index}`,
      file,
      name: file.name,
      status: 'pending' as const,
    }));

    setUploadedFiles(prev => [...prev, ...newFiles]);
    setError(null);
    setSuccessMessage(`${excelFiles.length} file berhasil ditambahkan`);
    setTimeout(() => setSuccessMessage(null), 3000);
  }, [uploadedFiles.length]);

  const handleDragOver = useCallback((event: React.DragEvent) => {
    event.preventDefault();
  }, []);

  // Remove file
  const removeFile = useCallback((id: string) => {
    setUploadedFiles(prev => prev.filter(f => f.id !== id));
  }, []);

  // Clear all files
  const clearAllFiles = useCallback(() => {
    setUploadedFiles([]);
    setExtractionResult(null);
    setError(null);
  }, []);

  // Ekstraksi data dari satu file
  const extractFileData = async (file: File): Promise<Record<string, number | string>> => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const data = new Uint8Array(e.target?.result as ArrayBuffer);
          const workbook = XLSX.read(data, { type: 'array' });
          
          const result: Record<string, number | string> = {};
          
          // Pre-scan FDM Kebun ABC untuk anchor row
          const fdmSheet = getSheetSmart(workbook, "FDM Kebun ABC");
          let fdmAnchorRow: number | null = null;
          if (fdmSheet) {
            fdmAnchorRow = findFDMAnchorRow(fdmSheet);
          }

          // Ekstraksi data untuk setiap item
          for (const item of itemsDefinitions) {
            const mode = item.mode;
            
            if (mode === "Formula") {
              result[item.label] = "";
              continue;
            }

            const ws = getSheetSmart(workbook, item.sheet);
            
            if (!ws) {
              result[item.label] = "Sheet Not Found";
            } else {
              if (mode === "Static" && item.addr) {
                result[item.label] = getCellValue(ws, item.addr) ?? "Error";
              } else if (mode === "Dynamic_Col_G" && item.keyword) {
                result[item.label] = findDynamicColG(ws, item.keyword) ?? "TIDAK DITEMUKAN";
              } else if (mode.startsWith("Dynamic_FDM_Bangunan")) {
                if (fdmAnchorRow !== null && fdmSheet) {
                  if (mode === "Dynamic_FDM_Bangunan_PerM2") {
                    result[item.label] = getCellValue(fdmSheet, XLSX.utils.encode_cell({ r: fdmAnchorRow, c: 4 })) ?? "Anchor Not Found";
                  } else if (mode === "Dynamic_FDM_Bangunan_Luas") {
                    result[item.label] = getCellValue(fdmSheet, XLSX.utils.encode_cell({ r: fdmAnchorRow - 1, c: 3 })) ?? "Anchor Not Found";
                  }
                } else {
                  result[item.label] = "Anchor Not Found";
                }
              } else {
                result[item.label] = "";
              }
            }

            // Bersihkan KELURAHAN dari nomor
            if (item.label === "KELURAHAN" && result[item.label] && typeof result[item.label] === "string") {
              const val = result[item.label] as string;
              if (val.includes("#")) {
                result[item.label] = val.replace(/#\s*\d+.*$/, '').trim();
              }
            }
          }

          resolve(result);
        } catch (err) {
          reject(err);
        }
      };
      reader.onerror = reject;
      reader.readAsArrayBuffer(file);
    });
  };

  // Hitung rumus-rumus - mengembalikan array dengan rumus Excel, bukan nilai hasil
  const calculateFormulas = (rowCount: number): Record<number, Record<string, string>> => {
    // Mapping kolom berdasarkan definisi itemsDefinitions
    // Indeks: 0=NO, 1=KPP, 2=Sektor, dst
    const colMap: Record<string, number> = {};
    const headers = ["NO", ...itemsDefinitions.map(item => item.label)];
    headers.forEach((h, i) => { colMap[h] = i; });

    const formulasByRow: Record<number, Record<string, string>> = {};

    for (let rowIdx = 0; rowIdx < rowCount; rowIdx++) {
      const excelRow = rowIdx + 2; // Baris Excel dimulai dari 2 (setelah header)
      const rowFormulas: Record<string, string> = {};

      // Helper untuk mendapatkan huruf kolom
      const col = (label: string): string => getColumnLetter(colMap[label]);

      // 1. LUAS BUMI (kolom J) = SUM(K:Q) - Areal Produktif sampai Areal Emplasemen
      // K = Areal Produktif, L = Areal Belum Diolah, M = Areal Sudah Diolah Belum Ditanami
      // N = Areal Pembibitan, O = Areal Tidak Produktif, P = Areal Pengaman, Q = Areal Emplasemen
      rowFormulas["LUAS BUMI"] = `=SUM(${col("Areal Produktif")}${excelRow}:${col("Areal Emplasemen")}${excelRow})`;

      // 2. Areal Produktif (Copy) (kolom R) = K (Areal Produktif)
      rowFormulas["Areal Produktif (Copy)"] = `=${col("Areal Produktif")}${excelRow}`;

      // 3. NJOP Bumi Berupa Tanah (Rp) (kolom T) = K * S (Areal Produktif * NJOP/M Areal Belum Produktif)
      rowFormulas["NJOP Bumi Berupa Tanah (Rp)"] = `=${col("Areal Produktif")}${excelRow}*${col("NJOP/M Areal Belum Produktif")}${excelRow}`;

      // 4. NJOP Bumi Areal Produktif (Rp) (kolom W) = T + U (NJOP Tanah + NJOP Pengembangan)
      rowFormulas["NJOP Bumi Areal Produktif (Rp)"] = `=${col("NJOP Bumi Berupa Tanah (Rp)")}${excelRow}+${col("NJOP Bumi Berupa Pengembangan Tanah (Rp)")}${excelRow}`;

      // 5. Luas Bumi Areal Produktif (m²) (kolom X) = K (Areal Produktif)
      rowFormulas["Luas Bumi Areal Produktif (m²)"] = `=${col("Areal Produktif")}${excelRow}`;

      // 6. NJOP Bumi Per M2 Areal Produktif (Rp/m2) (kolom Y) = W / X
      rowFormulas["NJOP Bumi Per M2 Areal Produktif (Rp/m2)"] = `=${col("NJOP Bumi Areal Produktif (Rp)")}${excelRow}/${col("Luas Bumi Areal Produktif (m²)")}${excelRow}`;

      // 7. NJOP BUMI (Rp) AREA PRODUKTIF pada A. DATA BUMI (kolom Z) = X * Y
      rowFormulas["NJOP BUMI (Rp) AREA PRODUKTIF pada A. DATA BUMI"] = `=${col("Luas Bumi Areal Produktif (m²)")}${excelRow}*${col("NJOP Bumi Per M2 Areal Produktif (Rp/m2)")}${excelRow}`;

      // 8. Areal Tidak Produktif (Copy) (kolom AD) = O (Areal Tidak Produktif)
      rowFormulas["Areal Tidak Produktif (Copy)"] = `=${col("Areal Tidak Produktif")}${excelRow}`;

      // 9. NJOP BUMI (Rp) AREAL TIDAK PRODUKTIF pada A. DATA BUMI (kolom AF) = AD * AE
      rowFormulas["NJOP BUMI (Rp) AREAL TIDAK PRODUKTIF pada A. DATA BUMI"] = `=${col("Areal Tidak Produktif (Copy)")}${excelRow}*${col("NJOP/M Areal Tidak Produktif")}${excelRow}`;

      // 10. Areal Pengaman (Copy) (kolom AH) = P (Areal Pengaman)
      rowFormulas["Areal Pengaman (Copy)"] = `=${col("Areal Pengaman")}${excelRow}`;

      // Rumus tambahan yang diperlukan

      // NJOP Bumi Berupa Pengembangan Tanah (Rp) (Kenaikan BIT 10.3%) = U * 1.103
      rowFormulas["NJOP Bumi Berupa Pengembangan Tanah (Rp) (Kenaikan BIT 10.3%)"] = `=${col("NJOP Bumi Berupa Pengembangan Tanah (Rp)")}${excelRow}*1.103`;

      // NJOP BUMI (Rp) AREA PRODUKTIF pada A. DATA BUMI (Proyeksi NDT Naik 46%)
      rowFormulas["NJOP BUMI (Rp) AREA PRODUKTIF pada A. DATA BUMI (Proyeksi NDT Naik 46%)"] = `=ROUND((${col("NJOP Bumi Berupa Tanah (Rp)")}${excelRow}+${col("NJOP Bumi Berupa Pengembangan Tanah (Rp) (Kenaikan BIT 10.3%)")}${excelRow})/${col("Areal Produktif")}${excelRow},0)*${col("Luas Bumi Areal Produktif (m²)")}${excelRow}`;

      // NJOP BUMI (Rp) AREAL BELUM PRODUKTIF pada A. DATA BUMI (Proyeksi NDT Naik 46%) = AB * 1.46
      rowFormulas["NJOP BUMI (Rp) AREAL BELUM PRODUKTIF pada A. DATA BUMI (Proyeksi NDT Naik 46%)"] = `=${col("NJOP BUMI (Rp) AREAL BELUM PRODUKTIF pada A. DATA BUMI")}${excelRow}*1.46`;

      // NJOP BUMI (Rp) AREAL TIDAK PRODUKTIF pada A. DATA BUMI (Proyeksi NDT Naik 46%) = AF * 1.46
      rowFormulas["NJOP BUMI (Rp) AREAL TIDAK PRODUKTIF pada A. DATA BUMI (Proyeksi NDT Naik 46%)"] = `=${col("NJOP BUMI (Rp) AREAL TIDAK PRODUKTIF pada A. DATA BUMI")}${excelRow}*1.46`;

      // NJOP BUMI (Rp) AREAL PENGAMAN pada A. DATA BUMI = AH * AI
      rowFormulas["NJOP BUMI (Rp) AREAL PENGAMAN pada A. DATA BUMI"] = `=${col("Areal Pengaman (Copy)")}${excelRow}*${col("NJOP/M Areal Pengaman")}${excelRow}`;

      // NJOP BUMI (Rp) AREAL PENGAMAN pada A. DATA BUMI (Proyeksi NDT Naik 46%) = AJ * 1.46
      rowFormulas["NJOP BUMI (Rp) AREAL PENGAMAN pada A. DATA BUMI (Proyeksi NDT Naik 46%)"] = `=${col("NJOP BUMI (Rp) AREAL PENGAMAN pada A. DATA BUMI")}${excelRow}*1.46`;

      // Areal Emplasemen (Copy) = Q
      rowFormulas["Areal Emplasemen (Copy)"] = `=${col("Areal Emplasemen")}${excelRow}`;

      // NJOP BUMI (Rp) AREAL EMPLASEMEN pada A. DATA BUMI = AL * AM
      rowFormulas["NJOP BUMI (Rp) AREAL EMPLASEMEN pada A. DATA BUMI"] = `=${col("Areal Emplasemen (Copy)")}${excelRow}*${col("NJOP/M Areal Emplasemen")}${excelRow}`;

      // NJOP BUMI (Rp) AREAL EMPLASEMEN pada A. DATA BUMI (Proyeksi NDT Naik 46%) = AN * 1.46
      rowFormulas["NJOP BUMI (Rp) AREAL EMPLASEMEN pada A. DATA BUMI (Proyeksi NDT Naik 46%)"] = `=${col("NJOP BUMI (Rp) AREAL EMPLASEMEN pada A. DATA BUMI")}${excelRow}*1.46`;

      // JUMLAH Luas (m2) pada A. DATA BUMI = J (LUAS BUMI)
      rowFormulas["JUMLAH Luas (m2) pada A. DATA BUMI"] = `=${col("LUAS BUMI")}${excelRow}`;

      // JUMLAH NJOP BUMI (Rp) pada A. DATA BUMI = SUM(Z, AB, AF, AJ, AN)
      rowFormulas["JUMLAH NJOP BUMI (Rp) pada A. DATA BUMI"] = `=${col("NJOP BUMI (Rp) AREA PRODUKTIF pada A. DATA BUMI")}${excelRow}+${col("NJOP BUMI (Rp) AREAL BELUM PRODUKTIF pada A. DATA BUMI")}${excelRow}+${col("NJOP BUMI (Rp) AREAL TIDAK PRODUKTIF pada A. DATA BUMI")}${excelRow}+${col("NJOP BUMI (Rp) AREAL PENGAMAN pada A. DATA BUMI")}${excelRow}+${col("NJOP BUMI (Rp) AREAL EMPLASEMEN pada A. DATA BUMI")}${excelRow}`;

      // Jumlah NJOP BANGUNAN pada B. DATA BANGUNAN = AR * AS
      rowFormulas["Jumlah NJOP BANGUNAN pada B. DATA BANGUNAN"] = `=${col("Jumlah LUAS pada B. DATA BANGUNAN")}${excelRow}*${col("NJOP BANGUNAN PER METER PERSEGI*) pada B. DATA BANGUNAN")}${excelRow}`;

      // TOTAL NJOP (TANAH + BANGUNAN) 2025 = AT + AU
      rowFormulas["TOTAL NJOP (TANAH + BANGUNAN) 2025"] = `=${col("JUMLAH NJOP BUMI (Rp) pada A. DATA BUMI")}${excelRow}+${col("Jumlah NJOP BANGUNAN pada B. DATA BANGUNAN")}${excelRow}`;

      // SPPT 2025 = ((AV - 12000000) * 40%) * 0.5%
      rowFormulas["SPPT 2025"] = `=(${col("TOTAL NJOP (TANAH + BANGUNAN) 2025")}${excelRow}-12000000)*0.4*0.005`;

      // SIMULASI TOTAL NJOP (TANAH + BANGUNAN) 2026 (Hanya Kenaikan BIT 10,3% + NDT Tetap)
      rowFormulas["SIMULASI TOTAL NJOP (TANAH + BANGUNAN) 2026 (Hanya Kenaikan BIT 10,3% + NDT Tetap)"] = `=ROUND((${col("NJOP Bumi Berupa Tanah (Rp)")}${excelRow}+${col("NJOP Bumi Berupa Pengembangan Tanah (Rp) (Kenaikan BIT 10.3%)")}${excelRow})/${col("Areal Produktif")}${excelRow},0)*${col("Luas Bumi Areal Produktif (m²)")}${excelRow}+${col("NJOP BUMI (Rp) AREAL BELUM PRODUKTIF pada A. DATA BUMI")}${excelRow}+${col("NJOP BUMI (Rp) AREAL TIDAK PRODUKTIF pada A. DATA BUMI")}${excelRow}+${col("NJOP BUMI (Rp) AREAL PENGAMAN pada A. DATA BUMI")}${excelRow}+${col("NJOP BUMI (Rp) AREAL EMPLASEMEN pada A. DATA BUMI")}${excelRow}+${col("Jumlah NJOP BANGUNAN pada B. DATA BANGUNAN")}${excelRow}`;

      // SIMULASI SPPT 2026 (Hanya Kenaikan BIT 10,3% + NDT Tetap)
      rowFormulas["SIMULASI SPPT 2026 (Hanya Kenaikan BIT 10,3% + NDT Tetap)"] = `=(${col("SIMULASI TOTAL NJOP (TANAH + BANGUNAN) 2026 (Hanya Kenaikan BIT 10,3% + NDT Tetap)")}${excelRow}-12000000)*0.4*0.005`;

      // Kenaikan = AY - AW
      rowFormulas["Kenaikan"] = `=${col("SIMULASI SPPT 2026 (Hanya Kenaikan BIT 10,3% + NDT Tetap)")}${excelRow}-${col("SPPT 2025")}${excelRow}`;

      // Persentase = AZ / AW
      rowFormulas["Persentase"] = `=${col("Kenaikan")}${excelRow}/${col("SPPT 2025")}${excelRow}`;

      // SIMULASI TOTAL NJOP (TANAH + BANGUNAN) 2026 (Kenaikan BIT 10,3% + NDT 46%)
      rowFormulas["SIMULASI TOTAL NJOP (TANAH + BANGUNAN) 2026 (Kenaikan BIT 10,3% + NDT 46%)"] = `=${col("NJOP BUMI (Rp) AREA PRODUKTIF pada A. DATA BUMI (Proyeksi NDT Naik 46%)")}${excelRow}+${col("NJOP BUMI (Rp) AREAL BELUM PRODUKTIF pada A. DATA BUMI (Proyeksi NDT Naik 46%)")}${excelRow}+${col("NJOP BUMI (Rp) AREAL TIDAK PRODUKTIF pada A. DATA BUMI (Proyeksi NDT Naik 46%)")}${excelRow}+${col("NJOP BUMI (Rp) AREAL PENGAMAN pada A. DATA BUMI (Proyeksi NDT Naik 46%)")}${excelRow}+${col("NJOP BUMI (Rp) AREAL EMPLASEMEN pada A. DATA BUMI (Proyeksi NDT Naik 46%)")}${excelRow}+${col("Jumlah NJOP BANGUNAN pada B. DATA BANGUNAN")}${excelRow}`;

      // SIMULASI SPPT 2026 (Kenaikan BIT 10,3% + NDT 46%)
      rowFormulas["SIMULASI SPPT 2026 (Kenaikan BIT 10,3% + NDT 46%)"] = `=(${col("SIMULASI TOTAL NJOP (TANAH + BANGUNAN) 2026 (Kenaikan BIT 10,3% + NDT 46%)")}${excelRow}-12000000)*0.4*0.005`;

      formulasByRow[rowIdx] = rowFormulas;
    }

    return formulasByRow;
  };

  // Ekstraksi semua file
  const handleExtract = async () => {
    if (uploadedFiles.length === 0) {
      setError('Tidak ada file untuk diekstrak. Silakan upload file terlebih dahulu.');
      return;
    }

    setIsExtracting(true);
    setExtractionProgress(0);
    setError(null);

    try {
      const allRowsData: CellValue[][] = [];
      
      for (let i = 0; i < uploadedFiles.length; i++) {
        const fileData = uploadedFiles[i];
        
        setUploadedFiles(prev => prev.map(f => 
          f.id === fileData.id ? { ...f, status: 'processing' } : f
        ));

        try {
          const extractedData = await extractFileData(fileData.file);
          // Data statis dari file
          const row: CellValue[] = [i + 1, ...itemsDefinitions.map(item => extractedData[item.label] ?? null)];
          allRowsData.push(row);

          setUploadedFiles(prev => prev.map(f => 
            f.id === fileData.id ? { ...f, status: 'completed', extractedData } : f
          ));
        } catch (err) {
          setUploadedFiles(prev => prev.map(f => 
            f.id === fileData.id ? { ...f, status: 'error', errorMessage: 'Gagal mengekstrak' } : f
          ));
        }

        setExtractionProgress(Math.round(((i + 1) / uploadedFiles.length) * 100));
      }

      // Generate formulas untuk setiap baris
      const formulasByRow = calculateFormulas(allRowsData.length);

      // Gabungkan data statis dengan formulas
      const finalRows = allRowsData.map((row, rowIdx) => {
        const rowFormulas = formulasByRow[rowIdx];
        const newRow = [...row];
        
        // Ganti nilai dengan formulas untuk kolom yang sesuai
        itemsDefinitions.forEach((item, colIdx) => {
          if (rowFormulas[item.label]) {
            newRow[colIdx + 1] = { f: rowFormulas[item.label].replace(/^=/, '') }; // Hapus = di awal karena xlsx sudah handle
          }
        });
        
        return newRow;
      });

      const headers = ["NO", ...itemsDefinitions.map(item => item.label)];
      setExtractionResult({ headers, rows: finalRows });
      setSuccessMessage('Ekstraksi berhasil diselesaikan!');
      setTimeout(() => setSuccessMessage(null), 3000);
    } catch (err) {
      setError('Terjadi kesalahan saat ekstraksi: ' + (err as Error).message);
    } finally {
      setIsExtracting(false);
    }
  };

  // Download hasil ekstraksi
  const handleDownload = () => {
    if (!extractionResult) {
      setError('Tidak ada hasil ekstraksi untuk didownload');
      return;
    }

    try {
      // Buat workbook baru
      const wb = XLSX.utils.book_new();
      
      // Sheet 1: Hasil
      const wsData = [extractionResult.headers, ...extractionResult.rows.map(row => 
        row.map(cell => {
          if (cell && typeof cell === 'object' && 'f' in cell) {
            return cell; // Formula cell
          }
          return cell;
        })
      )];
      const ws1 = XLSX.utils.aoa_to_sheet(wsData);
      
      // Dapatkan mapping kolom untuk header formulas
      const headers = extractionResult.headers;
      const colMap: Record<string, string> = {};
      headers.forEach((h, i) => { 
        colMap[h] = getColumnLetter(i);
      });

      // Header formulas yang dinamis - mengacu ke sheet '2. Kesimpulan'
      const headerFormulas: Record<string, { f: string }> = {
        'V1': { f: '="NJOP Bumi Berupa Pengembangan Tanah (Rp) (Kenaikan BIT "&\'2. Kesimpulan\'!$E$2*100&"%)"' },
        'AA1': { f: '="NJOP BUMI (Rp) AREA PRODUKTIF pada A. DATA BUMI (Proyeksi NDT Naik "&\'2. Kesimpulan\'!$E$14*100&"%)"' },
        'AC1': { f: '="NJOP BUMI (Rp) AREAL BELUM PRODUKTIF pada A. DATA BUMI (Proyeksi NDT Naik "&\'2. Kesimpulan\'!$E$14*100&"%)"' },
        'AG1': { f: '="NJOP BUMI (Rp) AREAL TIDAK PRODUKTIF pada A. DATA BUMI (Proyeksi NDT Naik "&\'2. Kesimpulan\'!$E$14*100&"%)"' },
        'AK1': { f: '="NJOP BUMI (Rp) AREAL PENGAMAN pada A. DATA BUMI (Proyeksi NDT Naik "&\'2. Kesimpulan\'!$E$14*100&"%)"' },
        'AO1': { f: '="NJOP BUMI (Rp) AREAL EMPLASEMEN pada A. DATA BUMI (Proyeksi NDT Naik "&\'2. Kesimpulan\'!$E$14*100&"%)"' },
        'AX1': { f: '="SIMULASI TOTAL NJOP (TANAH + BANGUNAN) 2026 (Hanya Kenaikan BIT "&\'2. Kesimpulan\'!$E$2*100&"% dan NDT Tetap)"' },
        'AY1': { f: '="SIMULASI SPPT 2026 (Hanya Kenaikan BIT "&\'2. Kesimpulan\'!$E$2*100&"% dan NDT Tetap)"' },
        'BB1': { f: '="SIMULASI TOTAL NJOP (TANAH + BANGUNAN) 2026 (Kenaikan BIT "&\'2. Kesimpulan\'!$E$2*100&"% + NDT "&\'2. Kesimpulan\'!$E$14*100&"%)"' },
        'BC1': { f: '="SIMULASI SPPT 2026 (Kenaikan BIT "&\'2. Kesimpulan\'!$E$2*100&"% + NDT "&\'2. Kesimpulan\'!$E$14*100&"%)"' },
      };

      // Apply header formulas
      Object.entries(headerFormulas).forEach(([cellAddr, formulaObj]) => {
        ws1[cellAddr] = formulaObj;
      });

      // Format kolom J (indeks 9) sampai BC (indeks 54) dengan format number Comma Style tanpa desimal
      const range = XLSX.utils.decode_range(ws1['!ref'] || 'A1');
      
      // Kolom J = indeks 9, kolom BC = indeks 54
      for (let col = 9; col <= 54; col++) {
        const colLetter = getColumnLetter(col);
        // Apply format untuk setiap baris data (mulai dari baris 2, indeks 1)
        for (let row = 1; row <= range.e.r; row++) {
          const cellAddr = `${colLetter}${row + 1}`;
          if (ws1[cellAddr]) {
            // Jika cell adalah formula atau nilai number, apply format
            if (ws1[cellAddr].f || typeof ws1[cellAddr].v === 'number') {
              ws1[cellAddr].z = '#,##0';
            }
          }
        }
      }
      
      XLSX.utils.book_append_sheet(wb, ws1, '1. Hasil');

      // Sheet 2: Kesimpulan - DENGAN ISIAN A2, B2, C2 YANG SUDAH DITAMBAHKAN
      const kesimpulanData = [
        // Baris 1: Header utama
        ['Poin', { f: '="Keterangan (BIT + "&\'2. Kesimpulan\'!$E$2*100&"% dan NDT Tetap)"' }, 'Nilai', 'Keterangan', 'Skenario Kenaikan BIT'],
        // Baris 2: Isian A2, B2, C2 sudah ditambahkan
        ['Simulasi Penerimaan PBB 2026', { f: '="Perkebunan (BIT + "&\'2. Kesimpulan\'!$E$2*100&"% dan NDT Tetap)"' }, { f: "=SUMIF('1. Hasil'!C2:C10000,\"Perkebunan\",'1. Hasil'!AY2:AY10000)" }, '', 0.103],
        // Baris 3-7: Data Simulasi (DINAIKKAN dari baris 2-6)
        ['Simulasi Penerimaan PBB 2026', 'Minerba', { f: "=SUMIF('1. Hasil'!C2:C10000,\"Minerba\",'1. Hasil'!AY2:AY10000)" }, '', ''],
        ['Simulasi Penerimaan PBB 2026', 'Perhutanan (HTI)', { f: "=SUMIF('1. Hasil'!C2:C10000,\"Perhutanan (HTI)\",'1. Hasil'!AY2:AY10000)" }, '', ''],
        ['Simulasi Penerimaan PBB 2026', 'Perhutanan (Hutan Alam)', { f: "=SUMIF('1. Hasil'!C2:C10000,\"Perhutanan (Hutan Alam)\",'1. Hasil'!AY2:AY10000)" }, '', ''],
        ['Simulasi Penerimaan PBB 2026', 'Sektor Lainnya', { f: "=SUMIF('1. Hasil'!C2:C10000,\"Sektor Lainnya\",'1. Hasil'!AY2:AY10000)" }, '', ''],
        // Baris 8: Collection Rate 100% (DINAIKKAN dari baris 7)
        ['Simulasi Penerimaan PBB 2026 (Collection Rate 100%)', { f: '=(COUNT(\'1. Hasil\'!A2:A10000))&" NOP"' }, { f: '=SUM(C2:C6)' }, '', ''],
        // Baris 9: Target (DINAIKKAN dari baris 8)
        ['Target Penerimaan PBB 2026', '', 110289165592, '', ''],
        // Baris 10: Selisih (DINAIKKAN dari baris 9)
        ['Selisih antara Simulasi (Collection Rate 100%) & Target', '', { f: '=C8-C9' }, { f: '=IF(C10>0,"Tercapai","Tidak Tercapai")' }, ''],
        // Baris kosong
        ['', '', '', '', ''],
        // Baris 12: Collection Rate 95% (DINAIKKAN dari baris 11)
        ['Simulasi Penerimaan PBB 2026 (Collection Rate 95%)', 0.95, { f: '=C8*B12' }, '', ''],
        // Baris 13: Selisih 95% (DINAIKKAN dari baris 12)
        ['Selisih antara Simulasi (Collection Rate 95%) Target', '', { f: '=C12-C9' }, { f: '=IF(C13>0,"Tercapai","Tidak Tercapai")' }, ''],
        // Baris kosong
        ['', '', '', '', ''],
        // Baris 15: Header kedua - NDT + 46% (DINAIKKAN dari baris 14)
        ['Poin', { f: '="Keterangan (BIT + "&\'2. Kesimpulan\'!$E$2*100&"% dan NDT + "&\'2. Kesimpulan\'!$E$15*100&"%)"' }, 'Nilai', 'Keterangan', 'Skenario Kenaikan NDT'],
        // Baris 16: Skenario Kenaikan NDT (DINAIKKAN dari baris 15)
        ['', '', '', '', 0.46],
        // Baris 17-21: Data kedua - mengacu ke baris 2-6 (BUKAN 3-7)
        ['=A2', '=B2', { f: "=SUMIF('1. Hasil'!C2:C10000,\"Perkebunan\",'1. Hasil'!BC2:BC10000)" }, '', ''],
        ['=A3', '=B3', { f: "=SUMIF('1. Hasil'!C2:C10000,\"Minerba\",'1. Hasil'!BC2:BC10000)" }, '', ''],
        ['=A4', '=B4', { f: "=SUMIF('1. Hasil'!C2:C10000,\"Perhutanan (HTI)\",'1. Hasil'!BC2:BC10000)" }, '', ''],
        ['=A5', '=B5', { f: "=SUMIF('1. Hasil'!C2:C10000,\"Perhutanan (Hutan Alam)\",'1. Hasil'!BC2:BC10000)" }, '', ''],
        ['=A6', '=B6', { f: "=SUMIF('1. Hasil'!C2:C10000,\"Sektor Lainnya\",'1. Hasil'!BC2:BC10000)" }, '', ''],
        // Baris 22: Summary kedua (DINAIKKAN dari baris 21)
        ['=A7', '=B7', { f: '=SUM(C17:C21)' }, '', ''],
        // Baris 23: Target kedua (DINAIKKAN dari baris 22)
        ['=A8', '', '=C9', '', ''],
        // Baris 24: Selisih kedua (DINAIKKAN dari baris 23)
        ['=A9', '', { f: '=C22-C23' }, { f: '=IF(C24>0,"Tercapai","Tidak Tercapai")' }, ''],
        // Baris kosong
        ['', '', '', '', ''],
        // Baris 26: Collection Rate 95% kedua (DINAIKKAN dari baris 25)
        ['Simulasi Penerimaan PBB 2026 (Collection Rate 95%)', 0.95, { f: '=C22*B26' }, '', ''],
        // Baris 27: Selisih 95% kedua (DINAIKKAN dari baris 26)
        ['Selisih antara Simulasi (Collection Rate 95%) Target', '', { f: '=C26-C23' }, { f: '=IF(C27>0,"Tercapai","Tidak Tercapai")' }, ''],
      ];
      
      const ws2 = XLSX.utils.aoa_to_sheet(kesimpulanData);
      
      // Format persentase untuk cell E2, B12, B26, E16
      if (ws2['E2']) ws2['E2'].z = '0.00%';
      if (ws2['B12']) ws2['B12'].z = '0%';
      if (ws2['B26']) ws2['B26'].z = '0%';
      if (ws2['E16']) ws2['E16'].z = '0%';
      
      // Format number untuk kolom C
      for (let row = 2; row <= 27; row++) {
        const cellAddr = `C${row}`;
        if (ws2[cellAddr] && (ws2[cellAddr].f || typeof ws2[cellAddr].v === 'number')) {
          ws2[cellAddr].z = '#,##0';
        }
      }
      
      XLSX.utils.book_append_sheet(wb, ws2, '2. Kesimpulan');

      // Download file
      const timestamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19);
      XLSX.writeFile(wb, `Hasil_Ekstraksi_FDM_${timestamp}.xlsx`);
      
      setSuccessMessage('File berhasil didownload!');
      setTimeout(() => setSuccessMessage(null), 3000);
    } catch (err) {
      setError('Gagal mendownload file: ' + (err as Error).message);
    }
  };

  // Reset untuk ekstraksi baru
  const handleNewExtraction = () => {
    clearAllFiles();
    setExtractionResult(null);
    setExtractionProgress(0);
    if (fileInputRef.current) {
      fileInputRef.current.value = '';
    }
  };

  return (
    <div className="min-h-screen bg-gradient-to-br from-slate-50 to-slate-100 p-4 md:p-8">
      <div className="max-w-6xl mx-auto">
        {/* Header */}
        <div className="text-center mb-8">
          <h1 className="text-3xl md:text-4xl font-bold text-slate-800 mb-2">
            Ekstraktor FDM
          </h1>
          <p className="text-slate-600">
            Ekstraksi data dari file FDM (Formulir Data Maklumat) dengan mudah dan cepat
          </p>
        </div>

        {/* Alert Messages */}
        {error && (
          <Alert variant="destructive" className="mb-6">
            <AlertCircle className="h-4 w-4" />
            <AlertDescription>{error}</AlertDescription>
          </Alert>
        )}
        
        {successMessage && (
          <Alert className="mb-6 bg-green-50 border-green-200 text-green-800">
            <CheckCircle className="h-4 w-4 text-green-600" />
            <AlertDescription>{successMessage}</AlertDescription>
          </Alert>
        )}

        {/* Main Card */}
        <Card className="shadow-xl border-0">
          <CardHeader className="bg-gradient-to-r from-blue-600 to-blue-700 text-white rounded-t-lg">
            <CardTitle className="text-xl font-semibold flex items-center gap-2">
              <FileSpreadsheet className="h-6 w-6" />
              Panel Ekstraksi FDM
            </CardTitle>
          </CardHeader>
          
          <CardContent className="p-6">
            {/* Tombol Utama */}
            <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-4 mb-8">
              <Button
                onClick={() => fileInputRef.current?.click()}
                className="h-16 flex flex-col items-center justify-center gap-2 bg-blue-600 hover:bg-blue-700"
                disabled={isExtracting}
              >
                <Upload className="h-6 w-6" />
                <span className="font-semibold">Upload FDM</span>
                <span className="text-xs opacity-80">Max 50 file</span>
              </Button>

              <Button
                onClick={handleExtract}
                disabled={isExtracting || uploadedFiles.length === 0}
                className="h-16 flex flex-col items-center justify-center gap-2 bg-green-600 hover:bg-green-700 disabled:bg-slate-300"
              >
                <Play className="h-6 w-6" />
                <span className="font-semibold">Ekstrak Sekarang</span>
                <span className="text-xs opacity-80">
                  {uploadedFiles.length > 0 ? `${uploadedFiles.length} file` : 'Belum ada file'}
                </span>
              </Button>

              <Button
                onClick={handleDownload}
                disabled={!extractionResult || isExtracting}
                className="h-16 flex flex-col items-center justify-center gap-2 bg-purple-600 hover:bg-purple-700 disabled:bg-slate-300"
              >
                <Download className="h-6 w-6" />
                <span className="font-semibold">Download Ulang</span>
                <span className="text-xs opacity-80">Hasil Ekstraksi</span>
              </Button>

              <Button
                onClick={handleNewExtraction}
                disabled={isExtracting}
                className="h-16 flex flex-col items-center justify-center gap-2 bg-orange-600 hover:bg-orange-700 disabled:bg-slate-300"
              >
                <RefreshCw className="h-6 w-6" />
                <span className="font-semibold">Ekstraksi FDM Lain</span>
                <span className="text-xs opacity-80">Mulai Baru</span>
              </Button>
            </div>

            <input
              ref={fileInputRef}
              type="file"
              multiple
              accept=".xlsm,.xlsx,.xls"
              onChange={handleFileUpload}
              className="hidden"
            />

            {/* Drop Zone */}
            {uploadedFiles.length === 0 && !isExtracting && (
              <div
                onDrop={handleDrop}
                onDragOver={handleDragOver}
                className="border-3 border-dashed border-blue-300 rounded-xl p-12 text-center bg-blue-50 hover:bg-blue-100 transition-colors cursor-pointer"
                onClick={() => fileInputRef.current?.click()}
              >
                <Upload className="h-16 w-16 mx-auto text-blue-400 mb-4" />
                <h3 className="text-lg font-semibold text-slate-700 mb-2">
                  Drag & Drop File FDM di sini
                </h3>
                <p className="text-slate-500 mb-4">
                  atau klik untuk memilih file
                </p>
                <p className="text-sm text-slate-400">
                  Format yang didukung: .xlsm, .xlsx, .xls (Max 50 file)
                </p>
              </div>
            )}

            {/* Progress Bar */}
            {isExtracting && (
              <div className="mb-6">
                <div className="flex justify-between mb-2">
                  <span className="text-sm font-medium text-slate-700">Sedang mengekstrak...</span>
                  <span className="text-sm font-medium text-slate-700">{extractionProgress}%</span>
                </div>
                <Progress value={extractionProgress} className="h-3" />
              </div>
            )}

            {/* File List */}
            {uploadedFiles.length > 0 && (
              <div className="mt-6">
                <div className="flex justify-between items-center mb-4">
                  <h3 className="text-lg font-semibold text-slate-700">
                    Daftar File ({uploadedFiles.length})
                  </h3>
                  <Button
                    variant="outline"
                    size="sm"
                    onClick={clearAllFiles}
                    disabled={isExtracting}
                    className="text-red-600 border-red-300 hover:bg-red-50"
                  >
                    <Trash2 className="h-4 w-4 mr-1" />
                    Hapus Semua
                  </Button>
                </div>

                <div className="max-h-96 overflow-y-auto border rounded-lg">
                  <table className="w-full">
                    <thead className="bg-slate-100 sticky top-0">
                      <tr>
                        <th className="px-4 py-3 text-left text-sm font-semibold text-slate-700">No</th>
                        <th className="px-4 py-3 text-left text-sm font-semibold text-slate-700">Nama File</th>
                        <th className="px-4 py-3 text-center text-sm font-semibold text-slate-700">Status</th>
                        <th className="px-4 py-3 text-center text-sm font-semibold text-slate-700">Aksi</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y">
                      {uploadedFiles.map((file, index) => (
                        <tr key={file.id} className="hover:bg-slate-50">
                          <td className="px-4 py-3 text-sm text-slate-600">{index + 1}</td>
                          <td className="px-4 py-3 text-sm text-slate-800">
                            <div className="flex items-center gap-2">
                              <FileSpreadsheet className="h-4 w-4 text-green-600" />
                              <span className="truncate max-w-xs" title={file.name}>
                                {file.name}
                              </span>
                            </div>
                          </td>
                          <td className="px-4 py-3 text-center">
                            {file.status === 'pending' && (
                              <span className="inline-flex items-center px-2 py-1 rounded-full text-xs font-medium bg-slate-100 text-slate-600">
                                Menunggu
                              </span>
                            )}
                            {file.status === 'processing' && (
                              <span className="inline-flex items-center px-2 py-1 rounded-full text-xs font-medium bg-blue-100 text-blue-600">
                                <RefreshCw className="h-3 w-3 mr-1 animate-spin" />
                                Proses
                              </span>
                            )}
                            {file.status === 'completed' && (
                              <span className="inline-flex items-center px-2 py-1 rounded-full text-xs font-medium bg-green-100 text-green-600">
                                <CheckCircle className="h-3 w-3 mr-1" />
                                Selesai
                              </span>
                            )}
                            {file.status === 'error' && (
                              <span className="inline-flex items-center px-2 py-1 rounded-full text-xs font-medium bg-red-100 text-red-600">
                                <XCircle className="h-3 w-3 mr-1" />
                                Error
                              </span>
                            )}
                          </td>
                          <td className="px-4 py-3 text-center">
                            <Button
                              variant="ghost"
                              size="sm"
                              onClick={() => removeFile(file.id)}
                              disabled={isExtracting}
                              className="text-red-600 hover:bg-red-50"
                            >
                              <Trash2 className="h-4 w-4" />
                            </Button>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            )}

            {/* Hasil Ekstraksi Summary */}
            {extractionResult && (
              <div className="mt-8 p-6 bg-green-50 border border-green-200 rounded-xl">
                <h3 className="text-lg font-semibold text-green-800 mb-4 flex items-center gap-2">
                  <CheckCircle className="h-5 w-5" />
                  Ringkasan Hasil Ekstraksi
                </h3>
                <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
                  <div className="bg-white p-4 rounded-lg shadow-sm">
                    <p className="text-sm text-slate-500">Total File</p>
                    <p className="text-2xl font-bold text-slate-800">{extractionResult.rows.length}</p>
                  </div>
                  <div className="bg-white p-4 rounded-lg shadow-sm">
                    <p className="text-sm text-slate-500">Total Kolom</p>
                    <p className="text-2xl font-bold text-slate-800">{extractionResult.headers.length}</p>
                  </div>
                  <div className="bg-white p-4 rounded-lg shadow-sm">
                    <p className="text-sm text-slate-500">Status</p>
                    <p className="text-lg font-bold text-green-600">Selesai</p>
                  </div>
                  <div className="bg-white p-4 rounded-lg shadow-sm">
                    <p className="text-sm text-slate-500">Output</p>
                    <p className="text-lg font-bold text-blue-600">Excel</p>
                  </div>
                </div>
              </div>
            )}
          </CardContent>
        </Card>

        {/* Footer */}
        <div className="mt-8 text-center text-sm text-slate-500">
          <p>
            <strong>Privasi Terjamin:</strong> File yang diupload tidak disimpan di server.
            Semua proses dilakukan di browser Anda.
          </p>
          <p className="mt-2">
            Ekstraktor FDM V5 &copy; {new Date().getFullYear()}
          </p>
        </div>
      </div>
    </div>
  );
}
