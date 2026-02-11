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
  AlertCircle,
  Trash2,
  X
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
  
  // Cari di seluruh baris sheet (tanpa batasan 150 baris)
  for (let row = 0; row <= range.e.r; row++) {
    // Cari di kolom A-E (0-4) seperti sebelumnya
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
      rowFormulas["NJOP Bumi Per M2 Areal Produktif (Rp/m2)"] = `=IF(${col("Luas Bumi Areal Produktif (m²)")}${excelRow}=0,0,${col("NJOP Bumi Areal Produktif (Rp)")}${excelRow}/${col("Luas Bumi Areal Produktif (m²)")}${excelRow})`;

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
      rowFormulas["NJOP Bumi Berupa Pengembangan Tanah (Rp) (Kenaikan BIT 10.3%)"] = `=${col("NJOP Bumi Berupa Pengembangan Tanah (Rp)")}${excelRow}*(1+'2. Kesimpulan'!$E$2)`;

      // NJOP BUMI (Rp) AREA PRODUKTIF pada A. DATA BUMI (Proyeksi NDT Naik 46%)
      rowFormulas["NJOP BUMI (Rp) AREA PRODUKTIF pada A. DATA BUMI (Proyeksi NDT Naik 46%)"] = `=IF(${col("Areal Produktif")}${excelRow}=0,0,ROUND((${col("NJOP Bumi Berupa Tanah (Rp)")}${excelRow}+${col("NJOP Bumi Berupa Pengembangan Tanah (Rp) (Kenaikan BIT 10.3%)")}${excelRow})/${col("Areal Produktif")}${excelRow},0)*${col("Luas Bumi Areal Produktif (m²)")}${excelRow})`;

      // NJOP BUMI (Rp) AREAL BELUM PRODUKTIF pada A. DATA BUMI (Proyeksi NDT Naik 46%) = AB * 1.46
      rowFormulas["NJOP BUMI (Rp) AREAL BELUM PRODUKTIF pada A. DATA BUMI (Proyeksi NDT Naik 46%)"] = `=${col("NJOP BUMI (Rp) AREAL BELUM PRODUKTIF pada A. DATA BUMI")}${excelRow}*(1+'2. Kesimpulan'!$E$14)`;

      // NJOP BUMI (Rp) AREAL TIDAK PRODUKTIF pada A. DATA BUMI (Proyeksi NDT Naik 46%) = AF * 1.46
      rowFormulas["NJOP BUMI (Rp) AREAL TIDAK PRODUKTIF pada A. DATA BUMI (Proyeksi NDT Naik 46%)"] = `=${col("NJOP BUMI (Rp) AREAL TIDAK PRODUKTIF pada A. DATA BUMI")}${excelRow}*(1+'2. Kesimpulan'!$E$14)`;

      // NJOP BUMI (Rp) AREAL PENGAMAN pada A. DATA BUMI = AH * AI
      rowFormulas["NJOP BUMI (Rp) AREAL PENGAMAN pada A. DATA BUMI"] = `=${col("Areal Pengaman (Copy)")}${excelRow}*${col("NJOP/M Areal Pengaman")}${excelRow}`;

      // NJOP BUMI (Rp) AREAL PENGAMAN pada A. DATA BUMI (Proyeksi NDT Naik 46%) = AJ * 1.46
      rowFormulas["NJOP BUMI (Rp) AREAL PENGAMAN pada A. DATA BUMI (Proyeksi NDT Naik 46%)"] = `=${col("NJOP BUMI (Rp) AREAL PENGAMAN pada A. DATA BUMI")}${excelRow}*(1+'2. Kesimpulan'!$E$14)`;

      // Areal Emplasemen (Copy) = Q
      rowFormulas["Areal Emplasemen (Copy)"] = `=${col("Areal Emplasemen")}${excelRow}`;

      // NJOP BUMI (Rp) AREAL EMPLASEMEN pada A. DATA BUMI = AL * AM
      rowFormulas["NJOP BUMI (Rp) AREAL EMPLASEMEN pada A. DATA BUMI"] = `=${col("Areal Emplasemen (Copy)")}${excelRow}*${col("NJOP/M Areal Emplasemen")}${excelRow}`;

      // NJOP BUMI (Rp) AREAL EMPLASEMEN pada A. DATA BUMI (Proyeksi NDT Naik 46%) = AN * 1.46
      rowFormulas["NJOP BUMI (Rp) AREAL EMPLASEMEN pada A. DATA BUMI (Proyeksi NDT Naik 46%)"] = `=${col("NJOP BUMI (Rp) AREAL EMPLASEMEN pada A. DATA BUMI")}${excelRow}*(1+'2. Kesimpulan'!$E$14)`;

      // JUMLAH Luas (m2) pada A. DATA BUMI = J (LUAS BUMI)
      rowFormulas["JUMLAH Luas (m2) pada A. DATA BUMI"] = `=${col("LUAS BUMI")}${excelRow}`;

      // JUMLAH NJOP BUMI (Rp) pada A. DATA BUMI = SUM(Z, AB, AF, AJ, AN)
      rowFormulas["JUMLAH NJOP BUMI (Rp) pada A. DATA BUMI"] = `=${col("NJOP BUMI (Rp) AREA PRODUKTIF pada A. DATA BUMI")}${excelRow}+${col("NJOP BUMI (Rp) AREAL BELUM PRODUKTIF pada A. DATA BUMI")}${excelRow}+${col("NJOP BUMI (Rp) AREAL TIDAK PRODUKTIF pada A. DATA BUMI")}${excelRow}+${col("NJOP BUMI (Rp) AREAL PENGAMAN pada A. DATA BUMI")}${excelRow}+${col("NJOP BUMI (Rp) AREAL EMPLASEMEN pada A. DATA BUMI")}${excelRow}`;

      // Jumlah NJOP BANGUNAN pada B. DATA BANGUNAN = AR * AS
      rowFormulas["Jumlah NJOP BANGUNAN pada B. DATA BANGUNAN"] = `=${col("Jumlah LUAS pada B. DATA BANGUNAN")}${excelRow}*${col("NJOP BANGUNAN PER METER PERSEGI*) pada B. DATA BANGUNAN")}${excelRow}`;

      // TOTAL NJOP (TANAH + BANGUNAN) 2025 = AT + AU
      rowFormulas["TOTAL NJOP (TANAH + BANGUNAN) 2025"] = `=${col("JUMLAH NJOP BUMI (Rp) pada A. DATA BUMI")}${excelRow}+${col("Jumlah NJOP BANGUNAN pada B. DATA BANGUNAN")}${excelRow}`;

      // SPPT 2025 = ((AV - 12000000) * 40%) * 0.5%
      rowFormulas["SPPT 2025"] = `=(${col("TOTAL NJOP (TANAH + BANGUNAN) 2025")}${excelRow}-12000000)*40%*0.5%`;

      // SIMULASI TOTAL NJOP (TANAH + BANGUNAN) 2026 (Hanya Kenaikan BIT 10,3% + NDT Tetap)
      rowFormulas["SIMULASI TOTAL NJOP (TANAH + BANGUNAN) 2026 (Hanya Kenaikan BIT 10,3% + NDT Tetap)"] = `=IF(${col("Areal Produktif")}${excelRow}=0,0,ROUND((${col("NJOP Bumi Berupa Tanah (Rp)")}${excelRow}+${col("NJOP Bumi Berupa Pengembangan Tanah (Rp) (Kenaikan BIT 10.3%)")}${excelRow})/${col("Areal Produktif")}${excelRow},0)*${col("Luas Bumi Areal Produktif (m²)")}${excelRow})+${col("NJOP BUMI (Rp) AREAL BELUM PRODUKTIF pada A. DATA BUMI")}${excelRow}+${col("NJOP BUMI (Rp) AREAL TIDAK PRODUKTIF pada A. DATA BUMI")}${excelRow}+${col("NJOP BUMI (Rp) AREAL PENGAMAN pada A. DATA BUMI")}${excelRow}+${col("NJOP BUMI (Rp) AREAL EMPLASEMEN pada A. DATA BUMI")}${excelRow}+${col("Jumlah NJOP BANGUNAN pada B. DATA BANGUNAN")}${excelRow}`;

      // SIMULASI SPPT 2026 (Hanya Kenaikan BIT 10,3% + NDT Tetap)
      rowFormulas["SIMULASI SPPT 2026 (Hanya Kenaikan BIT 10,3% + NDT Tetap)"] = `=(${col("SIMULASI TOTAL NJOP (TANAH + BANGUNAN) 2026 (Hanya Kenaikan BIT 10,3% + NDT Tetap)")}${excelRow}-12000000)*40%*0.5%`;

      // Kenaikan = AY - AW
      rowFormulas["Kenaikan"] = `=${col("SIMULASI SPPT 2026 (Hanya Kenaikan BIT 10,3% + NDT Tetap)")}${excelRow}-${col("SPPT 2025")}${excelRow}`;

      // Persentase = AZ / AW
      rowFormulas["Persentase"] = `=IF(${col("SPPT 2025")}${excelRow}=0,0,${col("Kenaikan")}${excelRow}/${col("SPPT 2025")}${excelRow})`;

      // SIMULASI TOTAL NJOP (TANAH + BANGUNAN) 2026 (Kenaikan BIT 10,3% + NDT 46%)
      rowFormulas["SIMULASI TOTAL NJOP (TANAH + BANGUNAN) 2026 (Kenaikan BIT 10,3% + NDT 46%)"] = `=${col("NJOP BUMI (Rp) AREA PRODUKTIF pada A. DATA BUMI (Proyeksi NDT Naik 46%)")}${excelRow}+${col("NJOP BUMI (Rp) AREAL BELUM PRODUKTIF pada A. DATA BUMI (Proyeksi NDT Naik 46%)")}${excelRow}+${col("NJOP BUMI (Rp) AREAL TIDAK PRODUKTIF pada A. DATA BUMI (Proyeksi NDT Naik 46%)")}${excelRow}+${col("NJOP BUMI (Rp) AREAL PENGAMAN pada A. DATA BUMI (Proyeksi NDT Naik 46%)")}${excelRow}+${col("NJOP BUMI (Rp) AREAL EMPLASEMEN pada A. DATA BUMI (Proyeksi NDT Naik 46%)")}${excelRow}+${col("Jumlah NJOP BANGUNAN pada B. DATA BANGUNAN")}${excelRow}`;

      // SIMULASI SPPT 2026 (Kenaikan BIT 10,3% + NDT 46%)
      rowFormulas["SIMULASI SPPT 2026 (Kenaikan BIT 10,3% + NDT 46%)"] = `=(${col("SIMULASI TOTAL NJOP (TANAH + BANGUNAN) 2026 (Kenaikan BIT 10,3% + NDT 46%)")}${excelRow}-12000000)*40%*0.5%`;

      formulasByRow[rowIdx] = rowFormulas;
    }

    return formulasByRow;
  };

  // Ekstraksi semua file
  const handleExtract = async () => {
    if (uploadedFiles.length === 0) {
      setError('Silakan Upload File FDM Terlebih Dulu');
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
      setSuccessMessage(`${allRowsData.length} dari ${uploadedFiles.length} file berhasil diekstrak`);
      setTimeout(() => setSuccessMessage(null), 3000);
    } catch (err) {
      setError('Terjadi kesalahan saat memproses file');
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
      
      // Set column widths for better readability
      ws1['!cols'] = headers.map(() => ({ wch: 25 }));
      
      XLSX.utils.book_append_sheet(wb, ws1, '1. Hasil');

      // Sheet 2: Kesimpulan - FIXED (removed empty row 2, shifted all data up by 1 row)
      const kesimpulanData: (string | number | { f: string; t?: 'n' | 's' })[][] = [
        // Row 1: Headers
        ['Poin', { f: '="Keterangan (BIT + "&E2*100&"% dan NDT Tetap)"', t: 's' as const }, 'Nilai', 'Keterangan', 'Skenario Kenaikan BIT'],
        // Row 2: Data starts immediately (no empty row)
        ['Simulasi Penerimaan PBB 2026', 'Perkebunan', { f: "SUMIF('1. Hasil'!C2:C10000,\"Perkebunan\",'1. Hasil'!AY2:AY10000)", t: 'n' as const }, '', 0.103],
        ['Simulasi Penerimaan PBB 2026', 'Minerba', { f: "SUMIF('1. Hasil'!C2:C10000,\"Minerba\",'1. Hasil'!AY2:AY10000)", t: 'n' as const }, '', ''],
        ['Simulasi Penerimaan PBB 2026', 'Perhutanan (HTI)', { f: "SUMIF('1. Hasil'!C2:C10000,\"Perhutanan (HTI)\",'1. Hasil'!AY2:AY10000)", t: 'n' as const }, '', ''],
        ['Simulasi Penerimaan PBB 2026', 'Perhutanan (Hutan Alam)', { f: "SUMIF('1. Hasil'!C2:C10000,\"Perhutanan (Hutan Alam)\",'1. Hasil'!AY2:AY10000)", t: 'n' as const }, '', ''],
        ['Simulasi Penerimaan PBB 2026', 'Sektor Lainnya', { f: "SUMIF('1. Hasil'!C2:C10000,\"Sektor Lainnya\",'1. Hasil'!AY2:AY10000)", t: 'n' as const }, '', ''],
        ['Simulasi Penerimaan PBB 2026 (Collection Rate 100%)', { f: "=(COUNT('1. Hasil'!A2:A10000))&\" NOP\"", t: 's' as const }, { f: 'SUM(C2:C6)', t: 'n' as const }, '', ''],
        ['Target Penerimaan PBB 2026', '', 110289165592, '', ''],
        ['Selisih antara Simulasi (Collection Rate 100%) & Target', '', { f: 'C7-C8', t: 'n' as const }, { f: 'IF(C9>0,"Tercapai","Tidak Tercapai")', t: 's' as const }, ''],
        [{ f: '="Simulasi Penerimaan PBB 2026 (Collection Rate "&B10*100&"%)"', t: 's' as const }, '95%', { f: 'C7*B10', t: 'n' as const }, '', ''],
        [{ f: '="Selisih antara Simulasi (Collection Rate "&B10*100&"%)"&" Target"', t: 's' as const }, '', { f: 'C10-C8', t: 'n' as const }, { f: 'IF(C11>0,"Tercapai","Tidak Tercapai")', t: 's' as const }, ''],
        // Empty row 12
        ['', '', '', '', ''],
        // Row 13: Second section header
        ['Poin', { f: '="Keterangan (BIT + "&E2*100&"% dan NDT + "&E14*100&"%)"', t: 's' as const }, 'Nilai', 'Keterangan', 'Skenario Kenaikan NDT'],
        // Row 14: Second section data starts
        ['Simulasi Penerimaan PBB 2026', 'Perkebunan', { f: "SUMIF('1. Hasil'!C2:C10000,\"Perkebunan\",'1. Hasil'!BC2:BC10000)", t: 'n' as const }, '', 0.46],
        ['Simulasi Penerimaan PBB 2026', 'Minerba', { f: "SUMIF('1. Hasil'!C2:C10000,\"Minerba\",'1. Hasil'!BC2:BC10000)", t: 'n' as const }, '', ''],
        ['Simulasi Penerimaan PBB 2026', 'Perhutanan (HTI)', { f: "SUMIF('1. Hasil'!C2:C10000,\"Perhutanan (HTI)\",'1. Hasil'!BC2:BC10000)", t: 'n' as const }, '', ''],
        ['Simulasi Penerimaan PBB 2026', 'Perhutanan (Hutan Alam)', { f: "SUMIF('1. Hasil'!C2:C10000,\"Perhutanan (Hutan Alam)\",'1. Hasil'!BC2:BC10000)", t: 'n' as const }, '', ''],
        ['Simulasi Penerimaan PBB 2026', 'Sektor Lainnya', { f: "SUMIF('1. Hasil'!C2:C10000,\"Sektor Lainnya\",'1. Hasil'!BC2:BC10000)", t: 'n' as const }, '', ''],
        ['Simulasi Penerimaan PBB 2026 (Collection Rate 100%)', { f: "=(COUNT('1. Hasil'!A2:A10000))&\" NOP\"", t: 's' as const }, { f: 'SUM(C14:C18)', t: 'n' as const }, '', ''],
        ['Target Penerimaan PBB 2026', '', { f: 'C8', t: 'n' as const }, '', ''],
        ['Selisih antara Simulasi (Collection Rate 100%) & Target', '', { f: 'C19-C20', t: 'n' as const }, { f: 'IF(C21>0,"Tercapai","Tidak Tercapai")', t: 's' as const }, ''],
        [{ f: '="Simulasi Penerimaan PBB 2026 (Collection Rate "&B22*100&"%)"', t: 's' as const }, '95%', { f: 'C19*B22', t: 'n' as const }, '', ''],
        [{ f: '="Selisih antara Simulasi (Collection Rate "&B22*100&"%)"&" Target"', t: 's' as const }, '', { f: 'C22-C20', t: 'n' as const }, { f: 'IF(C23>0,"Tercapai","Tidak Tercapai")', t: 's' as const }, ''],
      ];
      
      const ws2 = XLSX.utils.aoa_to_sheet(kesimpulanData);
      
      // Apply number formats to Sheet 2
      // Format E2 and E14 as percentage (10.3% and 46%)
      ws2['E2'] = { v: 0.103, t: 'n', z: '0.0%' };
      ws2['E14'] = { v: 0.46, t: 'n', z: '0.0%' };
      
      // Format C2:C11 and C14:C23 as Comma Style (#,##0)
      for (let row = 2; row <= 11; row++) {
        if (ws2[`C${row}`] && ws2[`C${row}`].t === 'n') {
          ws2[`C${row}`].z = '#,##0';
        }
      }
      for (let row = 14; row <= 23; row++) {
        if (ws2[`C${row}`] && ws2[`C${row}`].t === 'n') {
          ws2[`C${row}`].z = '#,##0';
        }
      }
      
      // Format B10 and B22 as percentage display (95%)
      ws2['B10'] = { v: 0.95, t: 'n', z: '0%' };
      ws2['B22'] = { v: 0.95, t: 'n', z: '0%' };
      
      ws2['!cols'] = [
        { wch: 60 },
        { wch: 30 },
        { wch: 25 },
        { wch: 20 },
        { wch: 20 }
      ];
      
      XLSX.utils.book_append_sheet(wb, ws2, '2. Kesimpulan');

      // Download file
      const timestamp = new Date().toISOString().slice(0, 19).replace(/:/g, '-');
      XLSX.writeFile(wb, `Hasil_Ekstraksi_FDM_${timestamp}.xlsx`);
      
      setSuccessMessage('File berhasil didownload!');
      setTimeout(() => setSuccessMessage(null), 3000);
    } catch (err) {
      setError('Gagal mendownload file: ' + (err as Error).message);
    }
  };

  return (
    <>
      {/* Font Imports */}
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Playfair+Display:wght@400;500;600;700&family=Inter:wght@300;400;500;600;700&family=Source+Code+Pro:wght@400;500;600&display=swap');
        
        :root {
          --font-heading: 'Playfair Display', serif;
          --font-body: 'Inter', sans-serif;
          --font-mono: 'Source Code Pro', monospace;
        }
        
        /* Headings - MongoDB Value Serif (Playfair Display) */
        h1, h2, h3, h4, h5, h6,
        .font-heading,
        [class*="CardTitle"] {
          font-family: var(--font-heading) !important;
        }
        
        /* Body text - Euclid Circular A (Inter) */
        body, p, span, div, button, input, label,
        .font-body {
          font-family: var(--font-body) !important;
        }
        
        /* Monospace - Source Code Pro */
        code, pre, .font-mono, .mono {
          font-family: var(--font-mono) !important;
        }
        
        /* Specific overrides for UI components */
        .text-3xl, .text-4xl, .font-bold {
          font-family: var(--font-heading) !important;
        }
        
        button, .button, [role="button"] {
          font-family: var(--font-body) !important;
        }
      `}</style>
      
      <div className="min-h-screen bg-gradient-to-br from-slate-50 to-slate-100 p-4 md:p-8 font-body">
        <div className="max-w-4xl mx-auto space-y-6">
          {/* Header */}
          <div className="text-center space-y-2">
            <h1 className="text-3xl md:text-4xl font-bold text-slate-900 font-heading">
              Ekstraktor FDM V5
            </h1>
            <p className="text-slate-600 font-body">
              Website ini Tidak Menyimpan File Apapun yang Di-upload dan Diekstrak
            </p>
          </div>

          {/* Alert Messages */}
          {error && (
            <Alert variant="destructive">
              <AlertCircle className="w-4 h-4" />
              <AlertDescription className="font-body">{error}</AlertDescription>
            </Alert>
          )}
          
          {successMessage && (
            <Alert className="bg-green-50 border-green-200">
              <CheckCircle className="w-4 h-4 text-green-600" />
              <AlertDescription className="text-green-800 font-body">{successMessage}</AlertDescription>
            </Alert>
          )}

          {/* Main Card */}
          <Card className="shadow-lg">
            <CardHeader>
              <CardTitle className="flex items-center gap-2 font-heading">
                <FileSpreadsheet className="w-5 h-5 text-blue-600" />
                Upload File FDM Versi V5
              </CardTitle>
            </CardHeader>
            
            <CardContent className="space-y-4">
              {/* Upload Area */}
              <div 
                className="border-2 border-dashed border-slate-300 rounded-lg p-8 text-center hover:border-blue-500 transition-colors"
                onDrop={handleDrop}
                onDragOver={handleDragOver}
              >
                <input
                  ref={fileInputRef}
                  type="file"
                  multiple
                  accept=".xlsm,.xlsx,.xls"
                  onChange={handleFileUpload}
                  className="hidden"
                  id="file-upload"
                />
                <label htmlFor="file-upload" className="cursor-pointer flex flex-col items-center gap-3">
                  <div className="w-16 h-16 bg-blue-100 rounded-full flex items-center justify-center">
                    <Upload className="w-8 h-8 text-blue-600" />
                  </div>
                  <div>
                    <p className="font-medium text-slate-900 font-body">
                      Klik Di Sini untuk Upload File
                    </p>
                    <p className="text-sm text-slate-500 font-body">
                      Pastikan File berupa FDM V5 dalam Format xlsm, xlsx, atau xls
                    </p>
                  </div>
                  <p className="text-xs text-slate-400 font-body">
                    Maksimal 50 file
                  </p>
                </label>
              </div>

              {/* File List */}
              {uploadedFiles.length > 0 && (
                <div className="space-y-3">
                  <div className="flex items-center justify-between">
                    <h3 className="font-medium text-slate-900 font-heading">
                      File yang dipilih ({uploadedFiles.length})
                    </h3>
                    <Button
                      variant="ghost"
                      size="sm"
                      onClick={clearAllFiles}
                      className="text-red-600 hover:text-red-700 font-body"
                    >
                      <Trash2 className="w-4 h-4 mr-1" />
                      Hapus Semua
                    </Button>
                  </div>

                  <div className="max-h-64 overflow-y-auto space-y-2">
                    {uploadedFiles.map((fileItem) => (
                      <div 
                        key={fileItem.id} 
                        className="flex items-center justify-between p-3 bg-slate-50 rounded-lg border"
                      >
                        <div className="flex items-center gap-3 min-w-0">
                          <FileSpreadsheet className="w-5 h-5 text-green-600 flex-shrink-0" />
                          <span className="truncate text-sm font-body">{fileItem.name}</span>
                        </div>
                        <div className="flex items-center gap-2 flex-shrink-0">
                          {fileItem.status === 'pending' && (
                            <span className="text-xs text-slate-500 font-body">Menunggu untuk Diekstrak</span>
                          )}
                          {fileItem.status === 'processing' && (
                            <RefreshCw className="w-4 h-4 text-blue-600 animate-spin" />
                          )}
                          {fileItem.status === 'completed' && (
                            <CheckCircle className="w-4 h-4 text-green-600" />
                          )}
                          {fileItem.status === 'error' && (
                            <AlertCircle className="w-4 h-4 text-red-600" />
                          )}
                          <button 
                            onClick={() => removeFile(fileItem.id)}
                            className="p-1 hover:bg-slate-200 rounded"
                          >
                            <X className="w-4 h-4 text-slate-500" />
                          </button>
                        </div>
                      </div>
                    ))}
                  </div>
                </div>
              )}

              {/* Progress Bar */}
              {isExtracting && (
                <div className="space-y-2">
                  <div className="flex justify-between text-sm font-body">
                    <span>Memproses file...</span>
                    <span>{Math.round(extractionProgress)}%</span>
                  </div>
                  <Progress value={extractionProgress} className="h-2" />
                </div>
              )}

              {/* Action Buttons */}
              <div className="flex gap-3">
                <Button
                  onClick={handleExtract}
                  disabled={isExtracting || uploadedFiles.length === 0}
                  className="flex-1 bg-blue-600 hover:bg-blue-700 font-body"
                >
                  <Play className="w-4 h-4 mr-2" />
                  {isExtracting ? 'Memproses...' : 'Ekstrak Sekarang'}
                </Button>
                
                {extractionResult && (
                  <Button
                    onClick={handleDownload}
                    variant="outline"
                    className="flex-1 border-green-600 text-green-600 hover:bg-green-50 font-body"
                  >
                    <Download className="w-4 h-4 mr-2" />
                    Download Hasil
                  </Button>
                )}
              </div>
            </CardContent>
          </Card>

          {/* Footer */}
          <div className="text-center text-sm text-slate-500 font-body">
            <p>Saran/Masukan: 0822-9411-6001 (Dedek)</p>
            <p className="mt-1">Update: {new Date().toLocaleDateString('id-ID', { day: 'numeric', month: 'long', year: 'numeric' })}</p>
          </div>
        </div>
      </div>
    </>
  );
}
