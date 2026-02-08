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

interface ExtractionResult {
  headers: string[];
  rows: (number | string | null)[][];
}

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
    wb.SheetNames.forEach(name => {
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

  // Hitung rumus-rumus
  const calculateFormulas = (rows: (number | string | null)[][]): (number | string | null)[][] => {
    const headers = ["NO", ...itemsDefinitions.map(item => item.label)];
    const colMap: Record<string, number> = {};
    headers.forEach((h, i) => { colMap[h] = i; });

    return rows.map((row) => {
      const newRow = [...row];

      // Helper untuk mendapatkan nilai kolom
      const getVal = (label: string): number => {
        const val = newRow[colMap[label]];
        return typeof val === 'number' ? val : 0;
      };

      // LUAS BUMI = SUM(Areal Produktif, Areal Belum Diolah, Areal Sudah Diolah Belum Ditanami, Areal Pembibitan, Areal Tidak Produktif, Areal Pengaman, Areal Emplasemen)
      const arealCols = ["Areal Produktif", "Areal Belum Diolah", "Areal Sudah Diolah Belum Ditanami", "Areal Pembibitan", "Areal Tidak Produktif", "Areal Pengaman", "Areal Emplasemen"];
      newRow[colMap["LUAS BUMI"]] = arealCols.reduce((sum, col) => sum + getVal(col), 0);

      // Areal Produktif (Copy)
      newRow[colMap["Areal Produktif (Copy)"]] = getVal("Areal Produktif");

      // NJOP Bumi Berupa Tanah (Rp) = Areal Produktif * NJOP/M Areal Belum Produktif
      newRow[colMap["NJOP Bumi Berupa Tanah (Rp)"]] = getVal("Areal Produktif") * getVal("NJOP/M Areal Belum Produktif");

      // NJOP Bumi Berupa Pengembangan Tanah (Rp) (Kenaikan BIT 10.3%) = NJOP Bumi Berupa Pengembangan Tanah (Rp) * 1.103
      const njopPengembangan = getVal("NJOP Bumi Berupa Pengembangan Tanah (Rp)");
      newRow[colMap["NJOP Bumi Berupa Pengembangan Tanah (Rp) (Kenaikan BIT 10.3%)"]] = njopPengembangan * 1.103;

      // NJOP Bumi Areal Produktif (Rp) = NJOP Bumi Berupa Tanah (Rp) + NJOP Bumi Berupa Pengembangan Tanah (Rp)
      newRow[colMap["NJOP Bumi Areal Produktif (Rp)"]] = getVal("NJOP Bumi Berupa Tanah (Rp)") + getVal("NJOP Bumi Berupa Pengembangan Tanah (Rp)");

      // Luas Bumi Areal Produktif (m²) = Areal Produktif
      newRow[colMap["Luas Bumi Areal Produktif (m²)"]] = getVal("Areal Produktif");

      // NJOP Bumi Per M2 Areal Produktif (Rp/m2) = NJOP Bumi Areal Produktif (Rp) / Luas Bumi Areal Produktif (m²)
      const luasProd = getVal("Luas Bumi Areal Produktif (m²)");
      const njopTotalProd = getVal("NJOP Bumi Areal Produktif (Rp)");
      newRow[colMap["NJOP Bumi Per M2 Areal Produktif (Rp/m2)"]] = luasProd > 0 ? Math.round(njopTotalProd / luasProd) : 0;

      // NJOP BUMI (Rp) AREA PRODUKTIF pada A. DATA BUMI = Luas Bumi Areal Produktif (m²) * NJOP Bumi Per M2 Areal Produktif (Rp/m2)
      newRow[colMap["NJOP BUMI (Rp) AREA PRODUKTIF pada A. DATA BUMI"]] = getVal("Luas Bumi Areal Produktif (m²)") * getVal("NJOP Bumi Per M2 Areal Produktif (Rp/m2)");

      // NJOP BUMI (Rp) AREA PRODUKTIF pada A. DATA BUMI (Proyeksi NDT Naik 46%) = ROUND((NJOP Bumi Berupa Tanah (Rp) + NJOP Bumi Berupa Pengembangan Tanah (Rp) (Kenaikan BIT 10.3%)) / Areal Produktif) * Luas Bumi Areal Produktif (m²)
      const tanah = getVal("NJOP Bumi Berupa Tanah (Rp)");
      const bitNaik = getVal("NJOP Bumi Berupa Pengembangan Tanah (Rp) (Kenaikan BIT 10.3%)");
      const arealProd = getVal("Areal Produktif");
      newRow[colMap["NJOP BUMI (Rp) AREA PRODUKTIF pada A. DATA BUMI (Proyeksi NDT Naik 46%)"]] = arealProd > 0 ? Math.round((tanah + bitNaik) / arealProd) * getVal("Luas Bumi Areal Produktif (m²)") : 0;

      // NJOP BUMI (Rp) AREAL BELUM PRODUKTIF pada A. DATA BUMI (Proyeksi NDT Naik 46%) = NJOP BUMI (Rp) AREAL BELUM PRODUKTIF pada A. DATA BUMI * 1.46
      newRow[colMap["NJOP BUMI (Rp) AREAL BELUM PRODUKTIF pada A. DATA BUMI (Proyeksi NDT Naik 46%)"]] = getVal("NJOP BUMI (Rp) AREAL BELUM PRODUKTIF pada A. DATA BUMI") * 1.46;

      // Areal Tidak Produktif (Copy)
      newRow[colMap["Areal Tidak Produktif (Copy)"]] = getVal("Areal Tidak Produktif");

      // NJOP BUMI (Rp) AREAL TIDAK PRODUKTIF pada A. DATA BUMI = Areal Tidak Produktif (Copy) * NJOP/M Areal Tidak Produktif
      newRow[colMap["NJOP BUMI (Rp) AREAL TIDAK PRODUKTIF pada A. DATA BUMI"]] = getVal("Areal Tidak Produktif (Copy)") * getVal("NJOP/M Areal Tidak Produktif");

      // NJOP BUMI (Rp) AREAL TIDAK PRODUKTIF pada A. DATA BUMI (Proyeksi NDT Naik 46%) = NJOP BUMI (Rp) AREAL TIDAK PRODUKTIF pada A. DATA BUMI * 1.46
      newRow[colMap["NJOP BUMI (Rp) AREAL TIDAK PRODUKTIF pada A. DATA BUMI (Proyeksi NDT Naik 46%)"]] = getVal("NJOP BUMI (Rp) AREAL TIDAK PRODUKTIF pada A. DATA BUMI") * 1.46;

      // Areal Pengaman (Copy)
      newRow[colMap["Areal Pengaman (Copy)"]] = getVal("Areal Pengaman");

      // NJOP BUMI (Rp) AREAL PENGAMAN pada A. DATA BUMI = Areal Pengaman (Copy) * NJOP/M Areal Pengaman
      newRow[colMap["NJOP BUMI (Rp) AREAL PENGAMAN pada A. DATA BUMI"]] = getVal("Areal Pengaman (Copy)") * getVal("NJOP/M Areal Pengaman");

      // NJOP BUMI (Rp) AREAL PENGAMAN pada A. DATA BUMI (Proyeksi NDT Naik 46%) = NJOP BUMI (Rp) AREAL PENGAMAN pada A. DATA BUMI * 1.46
      newRow[colMap["NJOP BUMI (Rp) AREAL PENGAMAN pada A. DATA BUMI (Proyeksi NDT Naik 46%)"]] = getVal("NJOP BUMI (Rp) AREAL PENGAMAN pada A. DATA BUMI") * 1.46;

      // Areal Emplasemen (Copy)
      newRow[colMap["Areal Emplasemen (Copy)"]] = getVal("Areal Emplasemen");

      // NJOP BUMI (Rp) AREAL EMPLASEMEN pada A. DATA BUMI = Areal Emplasemen (Copy) * NJOP/M Areal Emplasemen
      newRow[colMap["NJOP BUMI (Rp) AREAL EMPLASEMEN pada A. DATA BUMI"]] = getVal("Areal Emplasemen (Copy)") * getVal("NJOP/M Areal Emplasemen");

      // NJOP BUMI (Rp) AREAL EMPLASEMEN pada A. DATA BUMI (Proyeksi NDT Naik 46%) = NJOP BUMI (Rp) AREAL EMPLASEMEN pada A. DATA BUMI * 1.46
      newRow[colMap["NJOP BUMI (Rp) AREAL EMPLASEMEN pada A. DATA BUMI (Proyeksi NDT Naik 46%)"]] = getVal("NJOP BUMI (Rp) AREAL EMPLASEMEN pada A. DATA BUMI") * 1.46;

      // JUMLAH Luas (m2) pada A. DATA BUMI = LUAS BUMI
      newRow[colMap["JUMLAH Luas (m2) pada A. DATA BUMI"]] = getVal("LUAS BUMI");

      // JUMLAH NJOP BUMI (Rp) pada A. DATA BUMI = SUM(NJOP BUMI components)
      const njopComponents = ["NJOP BUMI (Rp) AREA PRODUKTIF pada A. DATA BUMI", "NJOP BUMI (Rp) AREAL BELUM PRODUKTIF pada A. DATA BUMI", "NJOP BUMI (Rp) AREAL TIDAK PRODUKTIF pada A. DATA BUMI", "NJOP BUMI (Rp) AREAL PENGAMAN pada A. DATA BUMI", "NJOP BUMI (Rp) AREAL EMPLASEMEN pada A. DATA BUMI"];
      newRow[colMap["JUMLAH NJOP BUMI (Rp) pada A. DATA BUMI"]] = njopComponents.reduce((sum, col) => sum + getVal(col), 0);

      // Jumlah NJOP BANGUNAN pada B. DATA BANGUNAN = Jumlah LUAS pada B. DATA BANGUNAN * NJOP BANGUNAN PER METER PERSEGI*) pada B. DATA BANGUNAN
      newRow[colMap["Jumlah NJOP BANGUNAN pada B. DATA BANGUNAN"]] = getVal("Jumlah LUAS pada B. DATA BANGUNAN") * getVal("NJOP BANGUNAN PER METER PERSEGI*) pada B. DATA BANGUNAN");

      // TOTAL NJOP (TANAH + BANGUNAN) 2025 = JUMLAH NJOP BUMI (Rp) pada A. DATA BUMI + Jumlah NJOP BANGUNAN pada B. DATA BANGUNAN
      newRow[colMap["TOTAL NJOP (TANAH + BANGUNAN) 2025"]] = getVal("JUMLAH NJOP BUMI (Rp) pada A. DATA BUMI") + getVal("Jumlah NJOP BANGUNAN pada B. DATA BANGUNAN");

      // SPPT 2025 = ((TOTAL NJOP (TANAH + BANGUNAN) 2025 - 12000000) * 40%) * 0.5%
      const totalNJOP25 = getVal("TOTAL NJOP (TANAH + BANGUNAN) 2025");
      newRow[colMap["SPPT 2025"]] = totalNJOP25 > 12000000 ? ((totalNJOP25 - 12000000) * 0.4) * 0.005 : 0;

      // SIMULASI TOTAL NJOP (TANAH + BANGUNAN) 2026 (Hanya Kenaikan BIT 10,3% + NDT Tetap)
      const T = getVal("NJOP Bumi Berupa Tanah (Rp)");
      const V = getVal("NJOP Bumi Berupa Pengembangan Tanah (Rp) (Kenaikan BIT 10.3%)");
      const R = getVal("Areal Produktif");
      const X = getVal("Luas Bumi Areal Produktif (m²)");
      const AB = getVal("NJOP BUMI (Rp) AREAL BELUM PRODUKTIF pada A. DATA BUMI");
      const AF = getVal("NJOP BUMI (Rp) AREAL TIDAK PRODUKTIF pada A. DATA BUMI");
      const AJ = getVal("NJOP BUMI (Rp) AREAL PENGAMAN pada A. DATA BUMI");
      const AN = getVal("NJOP BUMI (Rp) AREAL EMPLASEMEN pada A. DATA BUMI");
      const AT = getVal("Jumlah NJOP BANGUNAN pada B. DATA BANGUNAN");
      newRow[colMap["SIMULASI TOTAL NJOP (TANAH + BANGUNAN) 2026 (Hanya Kenaikan BIT 10,3% + NDT Tetap)"]] = R > 0 ? (Math.round((T + V) / R) * X + AB + AF + AJ + AN) + AT : AT;

      // SIMULASI SPPT 2026 (Hanya Kenaikan BIT 10,3% + NDT Tetap)
      const simNJOP26 = getVal("SIMULASI TOTAL NJOP (TANAH + BANGUNAN) 2026 (Hanya Kenaikan BIT 10,3% + NDT Tetap)");
      newRow[colMap["SIMULASI SPPT 2026 (Hanya Kenaikan BIT 10,3% + NDT Tetap)"]] = simNJOP26 > 12000000 ? ((simNJOP26 - 12000000) * 0.4) * 0.005 : 0;

      // Kenaikan = SIMULASI SPPT 2026 (Hanya Kenaikan BIT 10,3% + NDT Tetap) - SPPT 2025
      const simSPPT26 = getVal("SIMULASI SPPT 2026 (Hanya Kenaikan BIT 10,3% + NDT Tetap)");
      const sppt25 = getVal("SPPT 2025");
      newRow[colMap["Kenaikan"]] = simSPPT26 - sppt25;

      // Persentase = Kenaikan / SPPT 2025
      newRow[colMap["Persentase"]] = sppt25 > 0 ? (simSPPT26 - sppt25) / sppt25 : 0;

      // SIMULASI TOTAL NJOP (TANAH + BANGUNAN) 2026 (Kenaikan BIT 10,3% + NDT 46%)
      const AA = getVal("NJOP BUMI (Rp) AREA PRODUKTIF pada A. DATA BUMI (Proyeksi NDT Naik 46%)");
      const AC = getVal("NJOP BUMI (Rp) AREAL BELUM PRODUKTIF pada A. DATA BUMI (Proyeksi NDT Naik 46%)");
      const AG = getVal("NJOP BUMI (Rp) AREAL TIDAK PRODUKTIF pada A. DATA BUMI (Proyeksi NDT Naik 46%)");
      const AK = getVal("NJOP BUMI (Rp) AREAL PENGAMAN pada A. DATA BUMI (Proyeksi NDT Naik 46%)");
      const AO = getVal("NJOP BUMI (Rp) AREAL EMPLASEMEN pada A. DATA BUMI (Proyeksi NDT Naik 46%)");
      newRow[colMap["SIMULASI TOTAL NJOP (TANAH + BANGUNAN) 2026 (Kenaikan BIT 10,3% + NDT 46%)"]] = AA + AC + AG + AK + AO + AT;

      // SIMULASI SPPT 2026 (Kenaikan BIT 10,3% + NDT 46%)
      const simNJOP26NDT = getVal("SIMULASI TOTAL NJOP (TANAH + BANGUNAN) 2026 (Kenaikan BIT 10,3% + NDT 46%)");
      newRow[colMap["SIMULASI SPPT 2026 (Kenaikan BIT 10,3% + NDT 46%)"]] = simNJOP26NDT > 12000000 ? ((simNJOP26NDT - 12000000) * 0.4) * 0.005 : 0;

      return newRow;
    });
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
      const allRowsData: (number | string | null)[][] = [];
      
      for (let i = 0; i < uploadedFiles.length; i++) {
        const fileData = uploadedFiles[i];
        
        setUploadedFiles(prev => prev.map(f => 
          f.id === fileData.id ? { ...f, status: 'processing' } : f
        ));

        try {
          const extractedData = await extractFileData(fileData.file);
          const row = [i + 1, ...itemsDefinitions.map(item => extractedData[item.label] ?? null)];
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

      // Hitung rumus-rumus
      const calculatedRows = calculateFormulas(allRowsData);

      const headers = ["NO", ...itemsDefinitions.map(item => item.label)];
      setExtractionResult({ headers, rows: calculatedRows });
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
      const wsData = [extractionResult.headers, ...extractionResult.rows];
      const ws1 = XLSX.utils.aoa_to_sheet(wsData);
      
      // Format kolom J sampai BC sebagai number format
      const range = XLSX.utils.decode_range(ws1['!ref'] || 'A1');
      for (let col = 9; col <= 54; col++) {
        for (let row = 1; row <= range.e.r; row++) {
          const cellAddr = XLSX.utils.encode_cell({ r: row, c: col });
          if (ws1[cellAddr] && typeof ws1[cellAddr].v === 'number') {
            ws1[cellAddr].z = '#,##0';
          }
        }
      }
      
      XLSX.utils.book_append_sheet(wb, ws1, '1. Hasil');

      // Sheet 2: Kesimpulan
      const kesimpulanData = [
        ['', '', '', '', 'Skenario Kenaikan BIT'],
        ['', '', '', '', 0.103],
        ['Poin', 'Keterangan (BIT + 10.3% dan NDT Tetap)', 'Nilai', 'Keterangan', ''],
        ['Simulasi Penerimaan PBB 2026', 'Perkebunan', { f: "=SUMIF('1. Hasil'!C2:C10000,\"Perkebunan\",'1. Hasil'!AY2:AY10000)" }, '', ''],
        ['Simulasi Penerimaan PBB 2026', 'Minerba', { f: "=SUMIF('1. Hasil'!C2:C10000,\"Minerba\",'1. Hasil'!AY2:AY10000)" }, '', ''],
        ['Simulasi Penerimaan PBB 2026', 'Perhutanan (HTI)', { f: "=SUMIF('1. Hasil'!C2:C10000,\"Perhutanan (HTI)\",'1. Hasil'!AY2:AY10000)" }, '', ''],
        ['Simulasi Penerimaan PBB 2026', 'Perhutanan (Hutan Alam)', { f: "=SUMIF('1. Hasil'!C2:C10000,\"Perhutanan (Hutan Alam)\",'1. Hasil'!AY2:AY10000)" }, '', ''],
        ['Simulasi Penerimaan PBB 2026', 'Sektor Lainnya', { f: "=SUMIF('1. Hasil'!C2:C10000,\"Sektor Lainnya\",'1. Hasil'!AY2:AY10000)" }, '', ''],
        ['Simulasi Penerimaan PBB 2026 (Collection Rate 100%)', { f: '=(COUNT(\'1. Hasil\'!A2:A10000))&" NOP"' }, { f: '=SUM(D4:D8)' }, '', ''],
        ['Target Penerimaan PBB 2026', '', 110289165592, '', ''],
        ['Selisih antara Simulasi (Collection Rate 100%) & Target', '', { f: '=D9-D10' }, { f: '=IF(D11>0,"Tercapai","Tidak Tercapai")' }, ''],
        ['', '', '', '', ''],
        ['Simulasi Penerimaan PBB 2026 (Collection Rate 95%)', 0.95, { f: '=D9*B13' }, '', ''],
        ['Selisih antara Simulasi (Collection Rate 95%) Target', '', { f: '=D13-D10' }, { f: '=IF(D14>0,"Tercapai","Tidak Tercapai")' }, ''],
        ['', '', '', '', ''],
        ['Poin', 'Keterangan (BIT + 10.3% dan NDT + 46%)', 'Nilai', 'Keterangan', 'Skenario Kenaikan NDT'],
        ['=A4', '=B4', { f: "=SUMIF('1. Hasil'!C2:C10000,\"Perkebunan\",'1. Hasil'!BC2:BC10000)" }, '', 0.46],
        ['=A5', '=B5', { f: "=SUMIF('1. Hasil'!C2:C10000,\"Minerba\",'1. Hasil'!BC2:BC10000)" }, '', ''],
        ['=A6', '=B6', { f: "=SUMIF('1. Hasil'!C2:C10000,\"Perhutanan (HTI)\",'1. Hasil'!BC2:BC10000)" }, '', ''],
        ['=A7', '=B7', { f: "=SUMIF('1. Hasil'!C2:C10000,\"Perhutanan (Hutan Alam)\",'1. Hasil'!BC2:BC10000)" }, '', ''],
        ['=A8', '=B8', { f: "=SUMIF('1. Hasil'!C2:C10000,\"Sektor Lainnya\",'1. Hasil'!BC2:BC10000)" }, '', ''],
        ['=A9', '=B9', { f: '=SUM(D17:D21)' }, '', ''],
        ['=A10', '', '=D10', '', ''],
        ['=A11', '', { f: '=D22-D23' }, { f: '=IF(D24>0,"Tercapai","Tidak Tercapai")' }, ''],
        ['', '', '', '', ''],
        ['Simulasi Penerimaan PBB 2026 (Collection Rate 95%)', 0.95, { f: '=D22*B26' }, '', ''],
        ['Selisih antara Simulasi (Collection Rate 95%) Target', '', { f: '=D26-D23' }, { f: '=IF(D27>0,"Tercapai","Tidak Tercapai")' }, ''],
      ];
      
      const ws2 = XLSX.utils.aoa_to_sheet(kesimpulanData);
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
