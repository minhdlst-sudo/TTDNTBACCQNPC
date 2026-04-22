/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useEffect, useMemo } from "react";
import { motion, AnimatePresence } from "motion/react";
import { 
  Plus, 
  Table as TableIcon, 
  RefreshCw, 
  Filter, 
  Calendar as CalendarIcon, 
  ChevronDown,
  AlertCircle,
  CheckCircle2,
  Database,
  Search,
  Check,
  Download,
  Info,
  HelpCircle
} from "lucide-react";
import * as XLSX from "xlsx";
import { format } from "date-fns";
import { vi } from "date-fns/locale";
import {
  Dialog,
  DialogContent,
  DialogDescription,
  DialogHeader,
  DialogTitle,
  DialogTrigger,
} from "@/components/ui/dialog";
import {
  Command,
  CommandEmpty,
  CommandGroup,
  CommandInput,
  CommandItem,
  CommandList,
} from "@/components/ui/command";
import { Textarea } from "@/components/ui/textarea";

import { Button } from "@/components/ui/button";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from "@/components/ui/table";
import { Popover, PopoverContent, PopoverTrigger } from "@/components/ui/popover";
import { Calendar } from "@/components/ui/calendar";
import { Toaster } from "@/components/ui/sonner";
import { toast } from "sonner";
import { cn } from "@/lib/utils";
import { 
  BarChart, 
  Bar, 
  XAxis, 
  YAxis, 
  CartesianGrid, 
  Tooltip, 
  Legend, 
  ResponsiveContainer, 
  Cell,
  PieChart,
  Pie
} from "recharts";

const STATUS_COLORS: Record<string, string> = {
  "Dự kiến đạt kế hoạch cao": "#10b981",
  "Dự kiến đạt kế hoạch": "#3b82f6",
  "Khả năng đạt kế hoạch ở mức trung bình": "#f59e0b",
  "Khả năng đạt kế hoạch ở mức thấp": "#f97316",
  "Khả năng không đạt kế hoạch": "#ef4444",
  "Thiếu dữ liệu để đánh giá": "#94a3b8"
};

const normalizeString = (s: string) => {
  if (!s) return "";
  return String(s)
    .toLowerCase()
    .replace(/đ/g, "d")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ")
    .trim();
};

type SheetData = string[][];

const CustomTooltip = ({ active, payload, label }: any) => {
  if (active && payload && payload.length) {
    return (
      <div className="bg-white p-2 sm:p-3 border border-slate-200 shadow-xl rounded-lg text-[10px] sm:text-[12px] min-w-[160px] sm:min-w-[220px] z-[100]">
        <p className="font-bold text-slate-800 mb-1 sm:mb-2 border-b border-slate-100 pb-1">{label}</p>
        <div className="space-y-1 sm:space-y-1.5">
          {payload.map((entry: any, index: number) => (
            <div key={index} className="flex justify-between items-center gap-2 sm:gap-4">
              <div className="flex items-center gap-1.5 sm:gap-2">
                <div className="w-2 h-2 sm:w-2.5 sm:h-2.5 rounded-full shrink-0" style={{ backgroundColor: entry.fill }} />
                <span className="text-slate-600 font-medium whitespace-nowrap">{entry.name}</span>
              </div>
              <span className="font-bold" style={{ color: entry.fill }}>{entry.value}</span>
            </div>
          ))}
        </div>
      </div>
    );
  }
  return null;
};

export default function App() {
  const [dataSheet, setDataSheet] = useState<SheetData>([]);
  const [capNhatSheet, setCapNhatSheet] = useState<SheetData>([]);
  const [thuVienSheet, setThuVienSheet] = useState<SheetData>([]);
  const [tongHopSheet, setTongHopSheet] = useState<SheetData>([]);
  const [activeTab, setActiveTab] = useState<"cap-nhat" | "tong-hop" | "thong-ke">("cap-nhat");
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);

  // Form state
  const [dienLuc, setDienLuc] = useState("all");
  const [tenTram, setTenTram] = useState("");
  const [ngayThucHien, setNgayThucHien] = useState<Date | undefined>(new Date());
  const [phanLoai, setPhanLoai] = useState("");
  const [giaiPhap, setGiaiPhap] = useState("");
  const [vuongMac, setVuongMac] = useState("");
  const [deXuat, setDeXuat] = useState("");
  const [submitting, setSubmitting] = useState(false);

  // Filter state
  const [filterTenTram, setFilterTenTram] = useState("all");
  const [filterDataTenTram, setFilterDataTenTram] = useState("all");
  const [filterDonVi, setFilterDonVi] = useState("all");

  const [dateRange, setDateRange] = useState<{
    from: Date | undefined;
    to: Date | undefined;
  }>({
    from: undefined,
    to: undefined,
  });

  // Searchable Select state
  const [openTenTram, setOpenTenTram] = useState(false);
  const [openFilterCapNhat, setOpenFilterCapNhat] = useState(false);
  const [openFilterData, setOpenFilterData] = useState(false);
  const [openFilterDonVi, setOpenFilterDonVi] = useState(false);
  const [openDienLuc, setOpenDienLuc] = useState(false);
  const [selectedDienLucThongKe, setSelectedDienLucThongKe] = useState("all");
  const [openDienLucThongKe, setOpenDienLucThongKe] = useState(false);

  const exportToExcel = (header: string[], data: SheetData, fileName: string) => {
    if (data.length === 0) {
      toast.error("Không có dữ liệu để xuất");
      return;
    }
    try {
      const exportData = [header, ...data];
      const worksheet = XLSX.utils.aoa_to_sheet(exportData);
      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, worksheet, "Sheet1");
      XLSX.writeFile(workbook, `${fileName}_${format(new Date(), "yyyyMMdd_HHmm")}.xlsx`);
      toast.success(`Đã xuất file ${fileName}`);
    } catch (err) {
      console.error("Export error:", err);
      toast.error("Lỗi khi xuất file Excel");
    }
  };

  const exportStatisticsToExcel = () => {
    const statuses = Object.keys(STATUS_COLORS);
    const header = [
      "Điện lực",
      "Tổng số trạm",
      "Đã thực hiện",
      "Tỷ lệ thực hiện (%)",
      ...statuses.flatMap(s => [s, `Tỷ lệ ${s} (%)`]),
      "Điểm hiệu quả"
    ];

    const dataRows = statisticsData
      .filter(s => selectedDienLucThongKe === "all" || s.company === selectedDienLucThongKe || s.company === "TỔNG CỘNG")
      .map(stat => {
        const row = [
          stat.company,
          String(stat.totalStations),
          String(stat.implementedCount),
          stat.totalStations > 0 ? (stat.implementedCount / stat.totalStations * 100).toFixed(1) : "0",
          ...statuses.flatMap(status => [
            String(stat.counts[status as keyof typeof stat.counts]),
            stat.percentages[status].toFixed(1)
          ]),
          (stat.score || 0).toFixed(2)
        ];
        return row;
      });

    exportToExcel(header, dataRows, "Thong_ke_TTDN");
  };

  const exportPlanByUnitToExcel = () => {
    const header = [
      "Điện lực",
      "SCTX (KH)", "SCTX (TH)",
      "SCL (KH)", "SCL (TH)",
      "ĐTXD (KH)", "ĐTXD (TH)"
    ];

    const dataRows = planComparison.byUnit.map(u => [
      u.unit,
      String(u.SCTX.plan), String(u.SCTX.actual),
      String(u.SCL.plan), String(u.SCL.actual),
      String(u.DTXD.plan), String(u.DTXD.actual)
    ]);

    // Add total row
    dataRows.push([
      "TỔNG CỘNG",
      String(planComparison.total.SCTX.plan), String(planComparison.total.SCTX.actual),
      String(planComparison.total.SCL.plan), String(planComparison.total.SCL.actual),
      String(planComparison.total.DTXD.plan), String(planComparison.total.DTXD.actual)
    ]);

    exportToExcel(header, dataRows, "Tong_hop_Ke_hoach_Thuc_hien");
  };

  const [lastSync, setLastSync] = useState<Date | null>(null);

  const fetchData = async () => {
    setLoading(true);
    setError(null);
    try {
      // Add timestamp to bypass browser cache
      const ts = Date.now();
      const [dataRes, capNhatRes, thuVienRes, tongHopRes] = await Promise.all([
        fetch(`/api/sheets/data?t=${ts}`),
        fetch(`/api/sheets/cap-nhat?t=${ts}`),
        fetch(`/api/sheets/thu-vien?t=${ts}`),
        fetch(`/api/sheets/tong-hop?t=${ts}`)
      ]);

      if (!dataRes.ok || !capNhatRes.ok || !thuVienRes.ok || !tongHopRes.ok) {
        const dataErr = await dataRes.json();
        throw new Error(dataErr.error || "Failed to fetch data from sheets");
      }

      const data = await dataRes.json();
      const capNhat = await capNhatRes.json();
      const thuVien = await thuVienRes.json();
      const tongHop = await tongHopRes.json();

      console.log("Data fetched:", { 
        dataRows: data.length, 
        capNhatRows: capNhat.length, 
        thuVienRows: thuVien.length,
        tongHopRows: tongHop.length
      });

      setDataSheet(data);
      setCapNhatSheet(capNhat);
      setThuVienSheet(thuVien);
      setTongHopSheet(tongHop);
      setLastSync(new Date());
      
      if (ts) {
        toast.success("Dữ liệu đã được cập nhật mới nhất");
      }
    } catch (err: any) {
      console.error(err);
      setError(err.message);
      toast.error("Lỗi kết nối Google Sheets", {
        description: err.message
      });
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    fetchData();
  }, []);

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (
      dienLuc === "all" || 
      !tenTram || 
      !ngayThucHien || 
      !phanLoai || 
      !giaiPhap.trim() || 
      !vuongMac.trim() || 
      !deXuat.trim()
    ) {
      toast.warning("Vui lòng điền đầy đủ tất cả các trường thông tin");
      return;
    }

    setSubmitting(true);
    try {
      const res = await fetch("/api/sheets/cap-nhat", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          dienLuc,
          tenTram,
          ngayCapNhat: format(new Date(), "yyyy-MM-dd HH:mm:ss"),
          ngayThucHien: format(ngayThucHien, "yyyy-MM-dd"),
          phanLoai,
          giaiPhap,
          vuongMac,
          deXuat
        })
      });

      if (!res.ok) {
        const err = await res.json();
        throw new Error(err.error || "Failed to submit data");
      }

      toast.success("Cập nhật dữ liệu thành công!");
      setTenTram("");
      setPhanLoai("");
      setGiaiPhap("");
      setVuongMac("");
      setDeXuat("");
      fetchData();
    } catch (err: any) {
      toast.error("Lỗi khi gửi dữ liệu", {
        description: err.message
      });
    } finally {
      setSubmitting(false);
    }
  };

  // Extract station names from 'data' sheet (assuming column 'tên trạm' is at index 0 or has a header)
  const dienLucOptions = useMemo(() => {
    if (thuVienSheet.length === 0) return [];
    
    // Find header row (search first 5 rows)
    let headerIndex = -1;
    let dienLucIndex = -1;
    
    for (let i = 0; i < Math.min(thuVienSheet.length, 5); i++) {
      const row = thuVienSheet[i];
      if (!row || !Array.isArray(row)) continue;
      const normalizedRow = row.map(h => normalizeString(String(h || "")));
      
      // Look for "dien luc", "don vi", or if the first column looks like a power company list
      const idx = normalizedRow.findIndex(h => h.includes("dien luc") || h.includes("don vi"));
      
      if (idx !== -1) {
        headerIndex = i;
        dienLucIndex = idx;
        break;
      }
      
      // Fallback: If no header found but first row has data, assume first column is Điện lực if it contains common names
      if (i === 0 && normalizedRow[0] && (normalizedRow[0].includes("tp") || normalizedRow[0].includes("quan ngai"))) {
        headerIndex = -1; // Data starts from row 0
        dienLucIndex = 0;
        break;
      }
    }

    if (dienLucIndex === -1) {
      console.warn("Could not find 'Điện lực' column in 'thu vien' sheet. Headers found:", 
        thuVienSheet.slice(0, 5).map(row => row.join(", ")));
      return [];
    }

    console.log("Found 'Điện lực' at index:", dienLucIndex, "in row:", headerIndex);

    return Array.from(
      new Set(
        thuVienSheet
          .slice(headerIndex + 1)
          .map(row => row[dienLucIndex])
          .filter(Boolean)
          .map(s => s.trim())
      )
    ).sort();
  }, [thuVienSheet]);

  const stationNames = useMemo(() => {
    const findInSheet = (sheet: SheetData) => {
      if (sheet.length === 0) return null;
      
      let headerIndex = -1;
      let indexTenTram = -1;
      let indexDienLuc = -1;

      // Search first 5 rows for header
      for (let i = 0; i < Math.min(sheet.length, 5); i++) {
        const normalizedRow = sheet[i].map(h => normalizeString(h));
        const idxTram = normalizedRow.findIndex(h => h.includes("ten tram") || h === "tram" || h.includes("ten tba"));
        const idxDL = normalizedRow.findIndex(h => h.includes("dien luc") || h.includes("don vi"));
        
        if (idxTram !== -1) {
          headerIndex = i;
          indexTenTram = idxTram;
          indexDienLuc = idxDL;
          break;
        }
      }

      return { headerIndex, indexTenTram, indexDienLuc, sheet };
    };

    const thuVienInfo = findInSheet(thuVienSheet);
    const dataInfo = findInSheet(dataSheet);

    let info = null;
    if (thuVienInfo && thuVienInfo.indexTenTram !== -1 && thuVienInfo.indexDienLuc !== -1) {
      info = thuVienInfo;
    } else if (dataInfo && dataInfo.indexTenTram !== -1 && dataInfo.indexDienLuc !== -1) {
      info = dataInfo;
    } else if (dataInfo && dataInfo.indexTenTram !== -1) {
      info = dataInfo;
    } else if (thuVienInfo && thuVienInfo.indexTenTram !== -1) {
      info = thuVienInfo;
    }

    if (!info || info.indexTenTram === -1) return [];

    let result = info.sheet.slice(info.headerIndex + 1);
    if (dienLuc !== "all" && info.indexDienLuc !== -1) {
      const normalizedDienLuc = normalizeString(dienLuc);
      result = result.filter(row => {
        const val = row[info.indexDienLuc];
        return val && normalizeString(val) === normalizedDienLuc;
      });
    }

    return Array.from(
      new Set(
        result
          .map(row => String(row[info.indexTenTram] || ""))
          .filter(Boolean)
          .map(s => s.trim())
      )
    ).sort() as string[];
  }, [thuVienSheet, dataSheet, dienLuc]);

  const [selectedStationTongHop, setSelectedStationTongHop] = useState("");
  const [selectedDienLucTongHop, setSelectedDienLucTongHop] = useState("all");
  const [openSearchTongHop, setOpenSearchTongHop] = useState(false);
  const [openDienLucTongHop, setOpenDienLucTongHop] = useState(false);

  const donViOptionsTongHop = useMemo(() => {
    if (tongHopSheet.length < 2) return [];
    
    // Find header row (search first 5 rows)
    let headerIndex = -1;
    let indexDienLuc = -1;

    for (let i = 0; i < Math.min(tongHopSheet.length, 5); i++) {
      const normalizedRow = tongHopSheet[i].map(h => normalizeString(String(h || "")));
      const idx = normalizedRow.findIndex(h => h.includes("dien luc") || h.includes("don vi"));
      if (idx !== -1) {
        headerIndex = i;
        indexDienLuc = idx;
        break;
      }
    }
    
    if (indexDienLuc === -1) return [];
    
    const result = tongHopSheet.slice(headerIndex + 1);
    return Array.from(new Set(result.map(row => String(row[indexDienLuc] || "")).filter(Boolean))).sort() as string[];
  }, [tongHopSheet]);

  const stationNamesTongHop = useMemo(() => {
    if (tongHopSheet.length < 2) return [];
    
    // Find header row (search first 5 rows)
    let headerIndex = -1;
    let indexTenTram = -1;
    let indexDienLuc = -1;

    for (let i = 0; i < Math.min(tongHopSheet.length, 5); i++) {
      const normalizedRow = tongHopSheet[i].map(h => normalizeString(String(h || "")));
      const idxTram = normalizedRow.findIndex(h => h.includes("ten tram") || h === "tram" || h.includes("ten tba"));
      const idxDonVi = normalizedRow.findIndex(h => h.includes("dien luc") || h.includes("don vi"));
      
      if (idxTram !== -1) {
        headerIndex = i;
        indexTenTram = idxTram;
        indexDienLuc = idxDonVi;
        break;
      }
    }

    if (indexTenTram === -1) return [];

    let result = tongHopSheet.slice(headerIndex + 1);
    if (selectedDienLucTongHop !== "all" && indexDienLuc !== -1) {
      const normalizedDL = normalizeString(selectedDienLucTongHop);
      result = result.filter(row => {
        const val = row[indexDienLuc];
        return val && normalizeString(val) === normalizedDL;
      });
    }

    return Array.from(
      new Set(
        result
          .map(row => String(row[indexTenTram] || ""))
          .filter(Boolean)
          .map(s => s.trim())
      )
    ).sort() as string[];
  }, [tongHopSheet, selectedDienLucTongHop]);

  const filteredTongHopData = useMemo(() => {
    if (tongHopSheet.length < 2) return [];
    
    let headerIndex = -1;
    let indexDienLuc = -1;

    for (let i = 0; i < Math.min(tongHopSheet.length, 5); i++) {
      const normalizedRow = tongHopSheet[i].map(h => normalizeString(String(h || "")));
      const idx = normalizedRow.findIndex(h => h.includes("dien luc") || h.includes("don vi"));
      if (idx !== -1) {
        headerIndex = i;
        indexDienLuc = idx;
        break;
      }
    }
    
    let result = tongHopSheet.slice(headerIndex + 1);
    if (selectedDienLucTongHop !== "all" && indexDienLuc !== -1) {
      const normalizedDL = normalizeString(selectedDienLucTongHop);
      result = result.filter(row => {
        const val = row[indexDienLuc];
        return val && normalizeString(val) === normalizedDL;
      });
    }
    return result;
  }, [tongHopSheet, selectedDienLucTongHop]);

  const tongHopDetail = useMemo(() => {
    if (!selectedStationTongHop || tongHopSheet.length < 2) return null;
    
    const header = tongHopSheet[0];
    const normHeader = header.map(h => normalizeString(String(h || "")));
    const idxTram = normHeader.findIndex(h => h.includes("ten tram") || h === "tram" || h.includes("ten tba"));
    
    if (idxTram === -1) return null;
    
    const row = tongHopSheet.slice(1).find(r => r[idxTram] === selectedStationTongHop);
    if (!row) return null;

    const getAssessment = (val: any) => {
      const num = parseFloat(String(val).replace(/,/g, '.'));
      if (isNaN(num)) return "Thiếu dữ liệu để đánh giá";
      if (num >= 100) return "Dự kiến đạt kế hoạch cao";
      if (num >= 95) return "Dự kiến đạt kế hoạch";
      if (num >= 90) return "Khả năng đạt kế hoạch ở mức trung bình";
      if (num >= 70) return "Khả năng đạt kế hoạch ở mức thấp";
      return "Khả năng không đạt kế hoạch";
    };

    // Columns: A=0, B=1, ..., V=21, W=22
    // Group 1: B-H (1-7)
    // Group 2: I-N (8-13)
    // Group 3: W (22)
    // Group 4: O-V (14-21)
    
    const getGroup = (indices: number[]) => {
      return indices.map(i => ({
        label: header[i] || `Cột ${String.fromCharCode(65 + i)}`,
        value: row[i] || ""
      }));
    };

    const assessment = getAssessment(row[21]); // Column V is index 21

    return {
      group1: getGroup([1, 2, 3, 4, 5, 6, 7]),
      group2: getGroup([8, 9, 10, 11, 12, 13]),
      group3: getGroup([22]),
      group4: getGroup([14, 15, 16, 17, 18, 19, 20, 21]),
      assessment
    };
  }, [tongHopSheet, selectedStationTongHop]);

  const statisticsData = useMemo(() => {
    if (tongHopSheet.length < 2) return [];
    
    // Find header indices
    let headerIndex = -1;
    let idxTram = -1;
    let idxDL = -1;
    let idxDanhGia = 21; // Default index for assessment

    for (let i = 0; i < Math.min(tongHopSheet.length, 5); i++) {
      const row = tongHopSheet[i];
      if (!row || !Array.isArray(row)) continue;
      const normHeader = row.map(h => normalizeString(String(h || "")));
      const t = normHeader.findIndex(h => h.includes("ten tram") || h === "tram" || h.includes("ten tba"));
      const dl = normHeader.findIndex(h => h.includes("dien luc") || h.includes("don vi"));
      if (t !== -1) {
        headerIndex = i;
        idxTram = t;
        idxDL = dl;
        break;
      }
    }

    if (idxTram === -1 || idxDL === -1) return [];

    const rows = tongHopSheet.slice(headerIndex + 1);
    const companies = Array.from(new Set(rows.map(r => r[idxDL]).filter(Boolean).map(s => s.trim()))).sort();
    
    const getAssessment = (val: any) => {
      const num = parseFloat(String(val).replace(/,/g, '.'));
      if (isNaN(num)) return "Thiếu dữ liệu để đánh giá";
      if (num >= 100) return "Dự kiến đạt kế hoạch cao";
      if (num >= 95) return "Dự kiến đạt kế hoạch";
      if (num >= 90) return "Khả năng đạt kế hoạch ở mức trung bình";
      if (num >= 70) return "Khả năng đạt kế hoạch ở mức thấp";
      return "Khả năng không đạt kế hoạch";
    };

    // Prepare stations from capNhatSheet for checking "implemented"
    const implementedStations = new Set();
    if (capNhatSheet.length > 1) {
      const cnHeader = capNhatSheet[0].map(h => normalizeString(String(h || "")));
      const cnIdxTram = cnHeader.indexOf(normalizeString("tên trạm"));
      if (cnIdxTram !== -1) {
        capNhatSheet.slice(1).forEach(r => {
          if (r[cnIdxTram]) implementedStations.add(r[cnIdxTram].trim());
        });
      }
    }

    const stats = companies.map(company => {
      const companyRows = rows.filter(r => r[idxDL] && r[idxDL].trim() === company);
      const totalStations = companyRows.length;
      
      const implementedCount = companyRows.filter(r => r[idxTram] && implementedStations.has(r[idxTram].trim())).length;
      
      const counts = {
        "Dự kiến đạt kế hoạch cao": 0,
        "Dự kiến đạt kế hoạch": 0,
        "Khả năng đạt kế hoạch ở mức trung bình": 0,
        "Khả năng đạt kế hoạch ở mức thấp": 0,
        "Khả năng không đạt kế hoạch": 0,
        "Thiếu dữ liệu để đánh giá": 0
      };

      companyRows.forEach(r => {
        const status = getAssessment(r[idxDanhGia]);
        if (counts.hasOwnProperty(status)) {
          counts[status as keyof typeof counts]++;
        }
      });

      const implementedRate = totalStations > 0 ? (implementedCount / totalStations) * 100 : 0;
      const percentages = Object.fromEntries(
        Object.entries(counts).map(([k, v]) => [k, totalStations > 0 ? (v / totalStations) * 100 : 0])
      );

      const score = (implementedRate * 2) +
                    ((percentages["Dự kiến đạt kế hoạch cao"] || 0) * 2) +
                    ((percentages["Dự kiến đạt kế hoạch"] || 0) * 1.5) +
                    ((percentages["Khả năng đạt kế hoạch ở mức trung bình"] || 0) * 1) +
                    ((percentages["Khả năng đạt kế hoạch ở mức thấp"] || 0) * 0.5) -
                    ((percentages["Khả năng không đạt kế hoạch"] || 0) * 2) -
                    ((percentages["Thiếu dữ liệu để đánh giá"] || 0) * 1);

      return {
        company,
        totalStations,
        implementedCount,
        counts,
        percentages,
        score
      };
    });

    // Company Total
    const totals = {
      "Dự kiến đạt kế hoạch cao": stats.reduce((acc, curr) => acc + curr.counts["Dự kiến đạt kế hoạch cao"], 0),
      "Dự kiến đạt kế hoạch": stats.reduce((acc, curr) => acc + curr.counts["Dự kiến đạt kế hoạch"], 0),
      "Khả năng đạt kế hoạch ở mức trung bình": stats.reduce((acc, curr) => acc + curr.counts["Khả năng đạt kế hoạch ở mức trung bình"], 0),
      "Khả năng đạt kế hoạch ở mức thấp": stats.reduce((acc, curr) => acc + curr.counts["Khả năng đạt kế hoạch ở mức thấp"], 0),
      "Khả năng không đạt kế hoạch": stats.reduce((acc, curr) => acc + curr.counts["Khả năng không đạt kế hoạch"], 0),
      "Thiếu dữ liệu để đánh giá": stats.reduce((acc, curr) => acc + curr.counts["Thiếu dữ liệu để đánh giá"], 0)
    };

    const totalStationsAll = stats.reduce((acc, curr) => acc + curr.totalStations, 0);
    const totalImplementedAll = stats.reduce((acc, curr) => acc + curr.implementedCount, 0);
    const totalImplementedRate = totalStationsAll > 0 ? (totalImplementedAll / totalStationsAll) * 100 : 0;

    const totalPercentages = Object.fromEntries(
      Object.entries(totals).map(([k, v]) => [k, totalStationsAll > 0 ? (v / totalStationsAll) * 100 : 0])
    );

    const totalScore = (totalImplementedRate * 2) +
                       ((totalPercentages["Dự kiến đạt kế hoạch cao"] || 0) * 2) +
                       ((totalPercentages["Dự kiến đạt kế hoạch"] || 0) * 1.5) +
                       ((totalPercentages["Khả năng đạt kế hoạch ở mức trung bình"] || 0) * 1) +
                       ((totalPercentages["Khả năng đạt kế hoạch ở mức thấp"] || 0) * 0.5) -
                       ((totalPercentages["Khả năng không đạt kế hoạch"] || 0) * 2) -
                       ((totalPercentages["Thiếu dữ liệu để đánh giá"] || 0) * 1);

    const grandTotal = {
      company: "TỔNG CỘNG",
      totalStations: totalStationsAll,
      implementedCount: totalImplementedAll,
      counts: totals,
      percentages: totalPercentages,
      score: totalScore
    };

    return [...stats, grandTotal];
  }, [tongHopSheet, capNhatSheet]);

  const scoringData = useMemo(() => {
    return statisticsData
      .filter(s => s.company !== "TỔNG CỘNG")
      .sort((a, b) => (b.score || 0) - (a.score || 0));
  }, [statisticsData]);

  const planComparison = useMemo(() => {
    if (dataSheet.length < 2 || capNhatSheet.length < 2) {
      return {
        total: {
          SCTX: { plan: 0, actual: 0 },
          SCL: { plan: 0, actual: 0 },
          DTXD: { plan: 0, actual: 0 }
        },
        byUnit: []
      };
    }

    const dataHeader = dataSheet[0].map(h => normalizeString(String(h || "")));
    const cnHeader = capNhatSheet[0].map(h => normalizeString(String(h || "")));

    const findCol = (header: string[], keywords: string[]) => {
      return header.findIndex(h => keywords.some(k => h.includes(normalizeString(k))));
    };

    const idxDataTram = findCol(dataHeader, ["ten tram", "tba", "tram"]);
    const idxDataDL = findCol(dataHeader, ["dien luc", "don vi"]);
    const idxDataSCTX = findCol(dataHeader, ["ke hoach sctx", "kh sctx", "sctx"]);
    const idxDataSCL = findCol(dataHeader, ["ke hoach scl", "kh scl", "scl"]);
    const idxDataDTXD = findCol(dataHeader, ["ke hoach dtxd", "kh dtxd", "dtxd", "dau tu xay dung", "xdcb"]);

    const idxCnTram = findCol(cnHeader, ["ten tram", "tba", "tram"]);
    const idxCnDL = findCol(cnHeader, ["dien luc", "don vi"]);
    const idxCnPhanLoai = findCol(cnHeader, ["phan loai"]);
    const idxCnCongViec = findCol(cnHeader, ["cong viec da thuc hien", "giai phap da thuc hien", "giai phap", "noi dung"]);

    const units = new Set<string>();
    dataSheet.slice(1).forEach(row => {
      if (idxDataDL !== -1 && row[idxDataDL]) units.add(String(row[idxDataDL]).trim());
    });
    capNhatSheet.slice(1).forEach(row => {
      if (idxCnDL !== -1 && row[idxCnDL]) units.add(String(row[idxCnDL]).trim());
    });

    const unitList = Array.from(units).sort();
    
    // Initialize stats
    const unitStats: Record<string, any> = {};
    unitList.forEach(u => {
      unitStats[u] = {
        SCTX: { plan: new Set(), actual: new Set() },
        SCL: { plan: new Set(), actual: new Set() },
        DTXD: { plan: new Set(), actual: new Set() }
      };
    });

    const totalSets = {
      SCTX: { plan: new Set(), actual: new Set() },
      SCL: { plan: new Set(), actual: new Set() },
      DTXD: { plan: new Set(), actual: new Set() }
    };

    // Count in data sheet (Plan)
    dataSheet.slice(1).forEach(row => {
      const station = String(row[idxDataTram] || "").trim();
      const unit = String(row[idxDataDL] || "").trim();
      if (!station || idxDataTram === -1) return;
      
      if (idxDataSCTX !== -1 && row[idxDataSCTX] && String(row[idxDataSCTX]).trim() !== "") {
        totalSets.SCTX.plan.add(station);
        if (unitStats[unit]) unitStats[unit].SCTX.plan.add(station);
      }
      if (idxDataSCL !== -1 && row[idxDataSCL] && String(row[idxDataSCL]).trim() !== "") {
        totalSets.SCL.plan.add(station);
        if (unitStats[unit]) unitStats[unit].SCL.plan.add(station);
      }
      if (idxDataDTXD !== -1 && row[idxDataDTXD] && String(row[idxDataDTXD]).trim() !== "") {
        totalSets.DTXD.plan.add(station);
        if (unitStats[unit]) unitStats[unit].DTXD.plan.add(station);
      }
    });

    // Count in cap nhat sheet (Actual)
    capNhatSheet.slice(1).forEach(row => {
      const station = String(row[idxCnTram] || "").trim();
      const unit = String(row[idxCnDL] || "").trim();
      if (!station || idxCnTram === -1) return;

      const pl = normalizeString(String(row[idxCnPhanLoai] || ""));
      const hasWork = idxCnCongViec !== -1 && row[idxCnCongViec] && String(row[idxCnCongViec]).trim() !== "";
      
      if (hasWork) {
        if (pl.includes("sctx")) {
          totalSets.SCTX.actual.add(station);
          if (unitStats[unit]) unitStats[unit].SCTX.actual.add(station);
        } else if (pl.includes("scl")) {
          totalSets.SCL.actual.add(station);
          if (unitStats[unit]) unitStats[unit].SCL.actual.add(station);
        } else if (pl.includes("dtxd") || pl.includes("xdcb") || pl.includes("xaydung")) {
          totalSets.DTXD.actual.add(station);
          if (unitStats[unit]) unitStats[unit].DTXD.actual.add(station);
        }
      }
    });

    return {
      total: {
        SCTX: { plan: totalSets.SCTX.plan.size, actual: totalSets.SCTX.actual.size },
        SCL: { plan: totalSets.SCL.plan.size, actual: totalSets.SCL.actual.size },
        DTXD: { plan: totalSets.DTXD.plan.size, actual: totalSets.DTXD.actual.size }
      },
      byUnit: unitList.map(u => ({
        unit: u,
        SCTX: { plan: unitStats[u].SCTX.plan.size, actual: unitStats[u].SCTX.actual.size },
        SCL: { plan: unitStats[u].SCL.plan.size, actual: unitStats[u].SCL.actual.size },
        DTXD: { plan: unitStats[u].DTXD.plan.size, actual: unitStats[u].DTXD.actual.size }
      }))
    };
  }, [dataSheet, capNhatSheet]);

  const donViOptions = useMemo(() => {
    if (dataSheet.length < 2) return [];
    const header = dataSheet[0].map(h => h.toLowerCase());
    const index = header.indexOf("đơn vị");
    if (index === -1) return [];
    return Array.from(new Set(dataSheet.slice(1).map(row => row[index]).filter(Boolean))).sort();
  }, [dataSheet]);

  const filteredCapNhat = useMemo(() => {
    if (capNhatSheet.length < 2) return [];
    const header = capNhatSheet[0].map(h => normalizeString(String(h || "")));
    const indexTenTram = header.indexOf(normalizeString("tên trạm"));
    const indexDienLuc = header.indexOf(normalizeString("điện lực"));
    const indexNgayThucHien = header.indexOf(normalizeString("ngày thực hiện"));
    
    let result = capNhatSheet.slice(1).filter(row => {
      if (!row || row.length === 0) return false;
      const nonEmptyCount = row.filter(cell => cell && String(cell).trim() !== "").length;
      return nonEmptyCount > 1;
    });
    
    if (dienLuc !== "all" && indexDienLuc !== -1) {
      const normDL = normalizeString(dienLuc);
      result = result.filter(row => row[indexDienLuc] && normalizeString(String(row[indexDienLuc])) === normDL);
    }

    if (filterTenTram !== "all" && indexTenTram !== -1) {
      result = result.filter(row => row[indexTenTram] === filterTenTram);
    }

    // Date range filtering
    if (indexNgayThucHien !== -1 && (dateRange.from || dateRange.to)) {
      result = result.filter(row => {
        const dateStr = row[indexNgayThucHien];
        if (!dateStr) return false;
        const rowDate = new Date(dateStr);
        if (isNaN(rowDate.getTime())) return true; // Keep if invalid date to avoid losing data? Or hide it? Let's hide it.
        
        if (dateRange.from && rowDate < new Date(dateRange.from.setHours(0, 0, 0, 0))) return false;
        if (dateRange.to && rowDate > new Date(dateRange.to.setHours(23, 59, 59, 999))) return false;
        
        return true;
      });
    }
    
    return result;
  }, [capNhatSheet, filterTenTram, dienLuc, dateRange]);

  const filteredDataSheet = useMemo(() => {
    if (dataSheet.length < 2) return [];
    const header = dataSheet[0].map(h => normalizeString(String(h || "")));
    const indexTenTram = header.indexOf(normalizeString("tên trạm"));
    const indexDonVi = header.indexOf(normalizeString("đơn vị"));
    const indexDienLuc = header.indexOf(normalizeString("điện lực"));
    
    let result = dataSheet.slice(1).filter(row => {
      if (!row || row.length === 0) return false;
      const nonEmptyCount = row.filter(cell => cell && String(cell).trim() !== "").length;
      return nonEmptyCount > 1;
    });
    
    if (dienLuc !== "all") {
      const normDL = normalizeString(dienLuc);
      const idx = indexDonVi !== -1 ? indexDonVi : indexDienLuc;
      if (idx !== -1) {
        result = result.filter(row => row[idx] && normalizeString(String(row[idx])) === normDL);
      }
    }

    if (filterTenTram !== "all" && indexTenTram !== -1) {
      result = result.filter(row => row[indexTenTram] === filterTenTram);
    }

    if (filterDataTenTram !== "all" && indexTenTram !== -1) {
      result = result.filter(row => row[indexTenTram] === filterDataTenTram);
    }
    
    if (filterDonVi !== "all" && indexDonVi !== -1) {
      result = result.filter(row => row[indexDonVi] === filterDonVi);
    }
    
    return result;
  }, [dataSheet, filterDataTenTram, filterDonVi, dienLuc, filterTenTram]);

  const phanLoaiOptions = ["QLVH", "KD", "SCTX", "SCL", "ĐTXD"];

  if (error && error.includes("credentials missing")) {
    return (
      <div className="min-h-screen bg-slate-50 flex items-center justify-center p-4">
        <Card className="max-w-md w-full border-red-200 shadow-lg">
          <CardHeader className="text-center">
            <div className="mx-auto w-12 h-12 bg-red-100 rounded-full flex items-center justify-center mb-4">
              <AlertCircle className="text-red-600 w-6 h-6" />
            </div>
            <CardTitle className="text-red-900">Thiếu cấu hình Google Sheets</CardTitle>
            <CardDescription>
              Vui lòng thiết lập biến môi trường <code className="bg-slate-100 px-1 rounded">GOOGLE_SERVICE_ACCOUNT_EMAIL</code> và <code className="bg-slate-100 px-1 rounded">GOOGLE_PRIVATE_KEY</code> trong phần Secrets.
            </CardDescription>
          </CardHeader>
          <CardContent className="text-sm text-slate-600 space-y-4">
            <p>Để ứng dụng hoạt động, bạn cần:</p>
            <ol className="list-decimal list-inside space-y-2">
              <li>Tạo một Service Account trong Google Cloud Console.</li>
              <li>Tải file JSON và copy Email + Private Key.</li>
              <li>Chia sẻ quyền chỉnh sửa Google Sheet cho Email của Service Account.</li>
            </ol>
            <Button variant="outline" className="w-full mt-4" onClick={() => window.location.reload()}>
              Thử lại
            </Button>
          </CardContent>
        </Card>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-[#f0f2f5] text-foreground font-sans selection:bg-blue-100 flex flex-col">
      <Toaster position="top-right" richColors />
      
      {/* Top Bar */}
      <header className="bg-gradient-to-r from-[#1a73e8] to-[#0d47a1] border-b border-primary/20 flex flex-col md:h-[64px] md:flex-row md:items-center md:justify-between px-4 sm:px-6 py-3 md:py-0 flex-shrink-0 shadow-md z-20 gap-2 md:gap-4">
        <div className="flex items-center justify-between md:justify-start gap-3 overflow-hidden w-full md:w-auto">
          <div className="flex items-center gap-3 overflow-hidden">
            <div className="bg-white/10 p-1.5 rounded-lg backdrop-blur-sm shrink-0">
              <Database className="text-white w-5 h-5 md:w-6 md:h-6" />
            </div>
            <h1 className="font-bold text-[14px] sm:text-[16px] md:text-[20px] text-white tracking-tight uppercase drop-shadow-sm whitespace-nowrap overflow-hidden text-ellipsis">Quản lý TTĐN TBA công cộng - QNPC</h1>
          </div>
          
          <Button 
            variant="ghost" 
            size="sm" 
            onClick={fetchData} 
            disabled={loading}
            className="md:hidden text-white hover:bg-white/10 h-8 px-3 border border-white/20 rounded-full transition-all shrink-0"
          >
            <RefreshCw className={cn("w-3.5 h-3.5", loading && "animate-spin")} />
          </Button>
        </div>

        <div className="flex items-center justify-center md:flex-1">
          <div className="flex bg-white/15 p-1 rounded-lg w-full sm:w-auto max-w-[280px] sm:max-w-none">
            <button 
              onClick={() => setActiveTab("cap-nhat")}
              className={cn(
                "flex-1 sm:flex-none px-4 py-1.5 rounded-md text-[12px] md:text-[13px] font-semibold transition-all",
                activeTab === "cap-nhat" ? "bg-white text-[#1a73e8] shadow-sm" : "text-white/70 hover:text-white"
              )}
            >
              Cập nhật
            </button>
            <button 
              onClick={() => setActiveTab("tong-hop")}
              className={cn(
                "flex-1 sm:flex-none px-4 py-1.5 rounded-md text-[12px] md:text-[13px] font-semibold transition-all",
                activeTab === "tong-hop" ? "bg-white text-[#1a73e8] shadow-sm" : "text-white/70 hover:text-white"
              )}
            >
              Truy vấn
            </button>
            <button 
              onClick={() => setActiveTab("thong-ke")}
              className={cn(
                "flex-1 sm:flex-none px-4 py-1.5 rounded-md text-[12px] md:text-[13px] font-semibold transition-all",
                activeTab === "thong-ke" ? "bg-white text-[#1a73e8] shadow-sm" : "text-white/70 hover:text-white"
              )}
            >
              Thống kê
            </button>
          </div>
        </div>

        <div className="hidden md:flex items-center gap-4">
          <div className="flex flex-col items-end shrink-0">
            <div className="text-[12px] text-white/80 font-medium">Theo dõi: Đặng Xuân Duy - PKT</div>
            {lastSync && (
              <div className="text-[10px] text-white/60">Cập nhật lúc: {format(lastSync, "HH:mm:ss dd/MM")}</div>
            )}
          </div>
          <Button 
            variant="ghost" 
            size="sm" 
            onClick={fetchData} 
            disabled={loading}
            className="text-white hover:bg-white/10 h-9 px-4 border border-white/20 rounded-full transition-all"
          >
            <RefreshCw className={cn("w-4 h-4 mr-2", loading && "animate-spin")} />
            <span className="text-[13px] font-semibold">Làm mới</span>
          </Button>
        </div>
      </header>

      <main className="flex-1 flex flex-col md:flex-row p-2 sm:p-4 gap-4 overflow-hidden">
        {activeTab === "cap-nhat" ? (
          <>
            {/* Form Panel */}
            <section className="w-full md:w-[300px] flex-shrink-0">
          <motion.div
            initial={{ opacity: 0, x: -20 }}
            animate={{ opacity: 1, x: 0 }}
            transition={{ duration: 0.4 }}
            className="h-full"
          >
            <Card className="h-full shadow-sm border-border rounded-lg overflow-hidden flex flex-col">
              <CardHeader className="pb-4">
                <CardTitle className="text-[15px] font-semibold">Cập nhật thông tin</CardTitle>
                <CardDescription className="text-[12px]">Nhập thông tin cập nhật cho trạm</CardDescription>
              </CardHeader>
              <CardContent className="flex-1 overflow-y-auto">
                <form onSubmit={handleSubmit} className="space-y-4">
                  <div className="space-y-1">
                    <Label htmlFor="dien-luc" className="text-[12px] font-bold text-[#1a73e8]">Điện lực <span className="text-red-500">*</span></Label>
                    <Popover open={openDienLuc} onOpenChange={setOpenDienLuc}>
                      <PopoverTrigger 
                        render={
                          <Button
                            type="button"
                            variant="outline"
                            role="combobox"
                            aria-expanded={openDienLuc}
                            className="w-full h-10 justify-between bg-white border-border text-[13px] font-normal"
                          >
                            {dienLuc === "all" ? "Chọn Điện lực..." : dienLuc}
                            <Search className="ml-2 h-4 w-4 shrink-0 opacity-50" />
                          </Button>
                        }
                      />
                      <PopoverContent 
                        className="w-[var(--radix-popover-trigger-width)] p-0 bg-white shadow-xl border-border" 
                        align="start"
                      >
                        <Command>
                          <CommandInput placeholder="Tìm điện lực..." className="h-9 text-[13px]" autoFocus={false} />
                          <CommandList className="max-h-[300px]">
                            <CommandEmpty>Không tìm thấy.</CommandEmpty>
                            <CommandGroup>
                              <CommandItem
                                value="all"
                                onSelect={() => {
                                  setDienLuc("all");
                                  setTenTram("");
                                  setOpenDienLuc(false);
                                }}
                                className="text-[13px] cursor-pointer"
                              >
                                <Check className={cn("mr-2 h-4 w-4", dienLuc === "all" ? "opacity-100" : "opacity-0")} />
                                Tất cả điện lực
                              </CommandItem>
                              {dienLucOptions.map((name) => (
                                <CommandItem
                                  key={name}
                                  value={name}
                                  onSelect={() => {
                                    setDienLuc(name);
                                    setTenTram("");
                                    setOpenDienLuc(false);
                                  }}
                                  className="text-[13px] cursor-pointer"
                                >
                                  <Check className={cn("mr-2 h-4 w-4", dienLuc === name ? "opacity-100" : "opacity-0")} />
                                  {name}
                                </CommandItem>
                              ))}
                            </CommandGroup>
                          </CommandList>
                        </Command>
                      </PopoverContent>
                    </Popover>
                  </div>

                  <div className="space-y-1">
                    <Label htmlFor="ten-tram" className="text-[12px] font-bold text-[#1a73e8]">Tên trạm (từ Data) <span className="text-red-500">*</span></Label>
                    <Popover open={openTenTram} onOpenChange={setOpenTenTram}>
                      <PopoverTrigger 
                        render={
                          <Button
                            type="button"
                            variant="outline"
                            role="combobox"
                            aria-expanded={openTenTram}
                            className="w-full h-10 justify-between bg-white border-border text-[13px] font-normal"
                          >
                            {tenTram ? tenTram : "Tìm và chọn tên trạm..."}
                            <Search className="ml-2 h-4 w-4 shrink-0 opacity-50" />
                          </Button>
                        }
                      />
                      <PopoverContent 
                        className="w-[var(--radix-popover-trigger-width)] p-0 bg-white shadow-xl border-border" 
                        align="start"
                      >
                        <Command>
                          <CommandInput placeholder="Nhập tên trạm để tìm..." className="h-9 text-[13px]" autoFocus={false} />
                          <CommandList className="max-h-[300px]">
                            <CommandEmpty>Không tìm thấy trạm.</CommandEmpty>
                            <CommandGroup>
                              {stationNames.map((name) => (
                                <CommandItem
                                  key={name}
                                  value={name}
                                  onSelect={() => {
                                    setTenTram(name);
                                    setOpenTenTram(false);
                                  }}
                                  className="text-[13px] cursor-pointer"
                                >
                                  <Check
                                    className={cn(
                                      "mr-2 h-4 w-4",
                                      tenTram === name ? "opacity-100" : "opacity-0"
                                    )}
                                  />
                                  {name}
                                </CommandItem>
                              ))}
                            </CommandGroup>
                          </CommandList>
                        </Command>
                      </PopoverContent>
                    </Popover>
                  </div>

                  <div className="space-y-1">
                    <Label className="text-[12px] font-medium text-[#5f6368]">Ngày thực hiện <span className="text-red-500">*</span></Label>
                    <Popover>
                      <PopoverTrigger 
                        render={
                          <Button
                            variant="outline"
                            className={cn(
                              "w-full h-10 justify-start text-left font-normal bg-white border-border text-[13px] hover:bg-white",
                              !ngayThucHien && "text-muted-foreground"
                            )}
                          >
                            <CalendarIcon className="mr-2 h-4 w-4 text-[#5f6368]" />
                            {ngayThucHien ? format(ngayThucHien, "dd/MM/yyyy") : <span>Chọn ngày</span>}
                          </Button>
                        }
                      />
                      <PopoverContent className="w-auto p-0 bg-white" align="start">
                        <Calendar
                          mode="single"
                          selected={ngayThucHien}
                          onSelect={setNgayThucHien}
                          locale={vi}
                        />
                      </PopoverContent>
                    </Popover>
                  </div>

                  <div className="space-y-1">
                    <Label htmlFor="phan-loai" className="text-[12px] font-medium text-[#5f6368]">Phân loại giải pháp <span className="text-red-500">*</span></Label>
                    <Select value={phanLoai} onValueChange={setPhanLoai}>
                      <SelectTrigger id="phan-loai" className="h-9 text-[13px] bg-white border-border">
                        <SelectValue placeholder="Chọn phân loại" />
                      </SelectTrigger>
                      <SelectContent className="bg-white" align="start">
                        {phanLoaiOptions.map((opt) => (
                          <SelectItem key={opt} value={opt}>{opt}</SelectItem>
                        ))}
                      </SelectContent>
                    </Select>
                  </div>

                  <div className="space-y-1">
                    <Label htmlFor="giai-phap" className="text-[12px] font-bold text-[#1a73e8]">Giải pháp đã thực hiện <span className="text-red-500">*</span></Label>
                    <Textarea 
                      id="giai-phap" 
                      value={giaiPhap} 
                      onChange={(e) => setGiaiPhap(e.target.value)} 
                      className="min-h-[80px] text-[13px] bg-white border-border resize-none" 
                      placeholder="Nhập giải pháp chi tiết..."
                    />
                  </div>

                  <div className="space-y-1">
                    <Label htmlFor="vuong-mac" className="text-[12px] font-bold text-[#1a73e8]">Vướng mắc khó khăn <span className="text-red-500">*</span></Label>
                    <Textarea 
                      id="vuong-mac" 
                      value={vuongMac} 
                      onChange={(e) => setVuongMac(e.target.value)} 
                      className="min-h-[80px] text-[13px] bg-white border-border resize-none" 
                      placeholder="Nhập vướng mắc chi tiết..."
                    />
                  </div>

                  <div className="space-y-1">
                    <Label htmlFor="de-xuat" className="text-[12px] font-bold text-[#1a73e8]">Đề xuất <span className="text-red-500">*</span></Label>
                    <Textarea 
                      id="de-xuat" 
                      value={deXuat} 
                      onChange={(e) => setDeXuat(e.target.value)} 
                      className="min-h-[80px] text-[13px] bg-white border-border resize-none" 
                      placeholder="Nhập đề xuất chi tiết..."
                    />
                  </div>

                  <Button type="submit" className="w-full bg-primary hover:bg-primary/90 text-white font-semibold h-10 mt-2" disabled={submitting}>
                    {submitting ? (
                      <RefreshCw className="h-4 w-4 animate-spin" />
                    ) : (
                      "Gửi dữ liệu"
                    )}
                  </Button>
                  
                  <p className="text-[10px] text-[#9aa0a6] text-center mt-4">
                    Tự động lấy Ngày cập nhật: {format(new Date(), "yyyy-MM-dd HH:mm:ss")}
                  </p>
                </form>
              </CardContent>
            </Card>
          </motion.div>
        </section>

        {/* Data Panels */}
        <section className="flex-1 flex flex-col gap-4 overflow-hidden">
          {/* Cap Nhat Table Card */}
          <motion.div
            initial={{ opacity: 0, y: 20 }}
            animate={{ opacity: 1, y: 0 }}
            transition={{ duration: 0.4, delay: 0.1 }}
            className="flex-none h-[350px] min-h-0"
          >
            <Card className="h-full shadow-sm border-border rounded-lg overflow-hidden flex flex-col">
              <CardHeader className="flex flex-row items-center justify-between py-3 px-4 bg-[#fafafa] border-b border-border space-y-0">
                <CardTitle className="text-[14px] font-semibold text-[#5f6368]">
                  Dữ liệu mới cập nhật
                </CardTitle>
                <div className="flex items-center gap-3">
                  <div className="text-[11px] font-medium text-[#5f6368] bg-slate-100 px-2 py-1 rounded">
                    Tổng: {filteredCapNhat.length} dòng
                  </div>
                  <Button
                    variant="outline"
                    size="sm"
                    className="h-8 text-[12px] bg-white border-border text-emerald-600 hover:text-emerald-700 hover:bg-emerald-50"
                    onClick={() => exportToExcel(capNhatSheet[0] || [], filteredCapNhat, "Du_lieu_moi_cap_nhat")}
                  >
                    <Download className="w-3.5 h-3.5 mr-1.5" />
                    Xuất Excel
                  </Button>

                  <Popover>
                    <PopoverTrigger 
                      render={
                        <Button
                          variant="outline"
                          className={cn(
                            "h-8 text-[12px] bg-white border-border justify-start font-normal min-w-[200px]",
                            !dateRange.from && "text-[#5f6368]"
                          )}
                        >
                          <CalendarIcon className="mr-2 h-3 w-3" />
                          {dateRange.from ? (
                            dateRange.to ? (
                              <>
                                {format(dateRange.from, "dd/MM/yyyy")} - {format(dateRange.to, "dd/MM/yyyy")}
                              </>
                            ) : (
                              format(dateRange.from, "dd/MM/yyyy")
                            )
                          ) : (
                            <span>Lọc Ngày thực hiện...</span>
                          )}
                        </Button>
                      }
                    />
                    <PopoverContent className="w-auto p-0 bg-white shadow-xl border-border" align="start">
                      <Calendar
                        mode="range"
                        defaultMonth={dateRange.from}
                        selected={dateRange}
                        onSelect={(range: any) => setDateRange(range || { from: undefined, to: undefined })}
                        numberOfMonths={1}
                        locale={vi}
                      />
                      {(dateRange.from || dateRange.to) && (
                        <div className="p-2 border-t border-border flex justify-end">
                          <Button 
                            variant="ghost" 
                            size="sm" 
                            className="h-7 text-[11px] hover:bg-slate-100"
                            onClick={() => setDateRange({ from: undefined, to: undefined })}
                          >
                            Xóa lọc
                          </Button>
                        </div>
                      )}
                    </PopoverContent>
                  </Popover>

                  <Popover open={openFilterCapNhat} onOpenChange={setOpenFilterCapNhat}>
                    <PopoverTrigger 
                      render={
                        <Button
                          variant="outline"
                          role="combobox"
                          className="w-[180px] h-8 text-[12px] bg-white border-border justify-between"
                        >
                          {filterTenTram === "all" ? "Tất cả các trạm" : filterTenTram}
                          <Search className="ml-2 h-3 w-3 shrink-0 opacity-50" />
                        </Button>
                      }
                    />
                    <PopoverContent className="w-[var(--radix-popover-trigger-width)] p-0 bg-white shadow-xl border-border" align="start">
                      <Command>
                        <CommandInput placeholder="Tìm trạm..." className="h-8 text-[12px]" autoFocus={false} />
                        <CommandList>
                          <CommandEmpty>Không tìm thấy.</CommandEmpty>
                          <CommandGroup>
                            <CommandItem
                              value="all"
                              onSelect={() => {
                                setFilterTenTram("all");
                                setOpenFilterCapNhat(false);
                              }}
                              className="text-[12px] cursor-pointer"
                            >
                              <Check className={cn("mr-2 h-3 w-3", filterTenTram === "all" ? "opacity-100" : "opacity-0")} />
                              Tất cả các trạm
                            </CommandItem>
                            {stationNames.map((name) => (
                              <CommandItem
                                key={name}
                                value={name}
                                onSelect={(val) => {
                                  setFilterTenTram(val);
                                  setOpenFilterCapNhat(false);
                                }}
                                className="text-[12px] cursor-pointer"
                              >
                                <Check className={cn("mr-2 h-3 w-3", filterTenTram === name ? "opacity-100" : "opacity-0")} />
                                {name}
                              </CommandItem>
                            ))}
                          </CommandGroup>
                        </CommandList>
                      </Command>
                    </PopoverContent>
                  </Popover>
                </div>
              </CardHeader>
              <CardContent className="p-0 flex-1 overflow-auto relative">
                <table className="w-full caption-bottom text-[12px] border-separate border-spacing-0">
                  <TableHeader className="bg-[#f1f3f4] sticky top-0 z-20 shadow-sm">
                    <TableRow className="hover:bg-transparent border-b-2 border-border">
                      {capNhatSheet[0]?.map((header, i) => (
                        <TableHead key={i} className="h-10 font-bold text-[#3c4043] px-4 whitespace-nowrap bg-[#f1f3f4] border-b-2 border-border sticky top-0 z-20">{header}</TableHead>
                      ))}
                    </TableRow>
                  </TableHeader>
                  <TableBody>
                    {loading ? (
                      <TableRow>
                        <TableCell colSpan={capNhatSheet[0]?.length || 4} className="h-24 text-center">
                          <RefreshCw className="w-5 h-5 animate-spin mx-auto text-primary/30" />
                        </TableCell>
                      </TableRow>
                    ) : filteredCapNhat.length === 0 ? (
                      <TableRow>
                        <TableCell colSpan={capNhatSheet[0]?.length || 4} className="h-24 text-center text-slate-500">
                          Chưa có dữ liệu nào phù hợp.
                        </TableCell>
                      </TableRow>
                    ) : (
                      filteredCapNhat.map((row, i) => (
                        <TableRow key={i} className="hover:bg-blue-50/30 border-b border-border transition-colors">
                          {row.map((cell, j) => {
                            const headerName = capNhatSheet[0]?.[j]?.toLowerCase() || "";
                            const isLongText = headerName.includes("giải pháp") || 
                                             headerName.includes("vướng mắc") || 
                                             headerName.includes("đề xuất") ||
                                             headerName.includes("nguyên nhân") ||
                                             headerName.includes("kế hoạch");
                            
                            const isMediumText = headerName.includes("tên trạm") || 
                                               headerName.includes("đơn vị") ||
                                               headerName.includes("điện lực") ||
                                               headerName.includes("tiến độ");

                            return (
                              <TableCell key={j} className={cn(
                                "py-3 px-4 border-b border-border",
                                isLongText ? "min-w-[300px] max-w-[500px] whitespace-normal break-words leading-relaxed" : 
                                isMediumText ? "min-w-[150px] whitespace-normal" : "whitespace-nowrap"
                              )}>
                                {j === 3 ? (
                                  <span className={cn(
                                    "px-2 py-1 rounded-md text-[11px] font-bold shadow-sm inline-block",
                                    cell === "QLVH" ? "bg-blue-100 text-blue-700 border border-blue-200" : 
                                    cell === "KD" ? "bg-emerald-100 text-emerald-700 border border-emerald-200" : 
                                    "bg-slate-100 text-slate-700 border border-slate-200"
                                  )}>
                                    {cell}
                                  </span>
                                ) : cell}
                              </TableCell>
                            );
                          })}
                        </TableRow>
                      ))
                    )}
                  </TableBody>
                </table>
              </CardContent>
            </Card>
          </motion.div>

          {/* Data Table Card */}
          <motion.div
            initial={{ opacity: 0, y: 20 }}
            animate={{ opacity: 1, y: 0 }}
            transition={{ duration: 0.4, delay: 0.2 }}
            className="flex-none h-[480px] min-h-0"
          >
            <Card className="h-full shadow-sm border-border rounded-lg overflow-hidden flex flex-col">
              <CardHeader className="flex flex-row items-center justify-between py-3 px-4 bg-[#fafafa] border-b border-border space-y-0">
                <CardTitle className="text-[14px] font-bold text-[#3c4043]">
                  Chi tiết kế hoạch giảm TTĐN của Điện lực
                </CardTitle>
                <div className="flex items-center gap-3">
                  <div className="text-[11px] font-medium text-[#5f6368] bg-slate-100 px-2 py-1 rounded">
                    Tổng: {filteredDataSheet.length} dòng
                  </div>
                  <Button
                    variant="outline"
                    size="sm"
                    className="h-8 text-[12px] bg-white border-border text-emerald-600 hover:text-emerald-700 hover:bg-emerald-50"
                    onClick={() => exportToExcel(dataSheet[0] || [], filteredDataSheet, "Du_lieu_tong_hop")}
                  >
                    <Download className="w-3.5 h-3.5 mr-1.5" />
                    Xuất Excel
                  </Button>
                  
                  {/* Filter Ten Tram */}
                  <Popover open={openFilterData} onOpenChange={setOpenFilterData}>
                    <PopoverTrigger 
                      render={
                        <Button
                          type="button"
                          variant="outline"
                          role="combobox"
                          className="w-[180px] h-8 text-[12px] bg-white border-border justify-between"
                        >
                          {filterDataTenTram === "all" ? "Lọc Tên trạm" : filterDataTenTram}
                          <Search className="ml-2 h-3 w-3 shrink-0 opacity-50" />
                        </Button>
                      }
                    />
                    <PopoverContent 
                      className="w-[var(--radix-popover-trigger-width)] p-0 bg-white shadow-xl border-border" 
                      align="start"
                    >
                      <Command>
                        <CommandInput placeholder="Tìm trạm..." className="h-8 text-[12px]" autoFocus={false} />
                        <CommandList>
                          <CommandEmpty>Không tìm thấy.</CommandEmpty>
                          <CommandGroup>
                            <CommandItem
                              value="all"
                              onSelect={() => {
                                setFilterDataTenTram("all");
                                setOpenFilterData(false);
                              }}
                              className="text-[12px] cursor-pointer"
                            >
                              <Check className={cn("mr-2 h-3 w-3", filterDataTenTram === "all" ? "opacity-100" : "opacity-0")} />
                              Tất cả các trạm
                            </CommandItem>
                            {stationNames.map((name) => (
                              <CommandItem
                                key={name}
                                value={name}
                                onSelect={(val) => {
                                  setFilterDataTenTram(val);
                                  setOpenFilterData(false);
                                }}
                                className="text-[12px] cursor-pointer"
                              >
                                <Check className={cn("mr-2 h-3 w-3", filterDataTenTram === name ? "opacity-100" : "opacity-0")} />
                                {name}
                              </CommandItem>
                            ))}
                          </CommandGroup>
                        </CommandList>
                      </Command>
                    </PopoverContent>
                  </Popover>
                </div>
              </CardHeader>
              <CardContent className="p-0 flex-1 overflow-auto relative">
                <table className="w-full caption-bottom text-[12px] border-separate border-spacing-0">
                  <TableHeader className="bg-[#f1f3f4] sticky top-0 z-20 shadow-sm">
                    <TableRow className="hover:bg-transparent border-b-2 border-border">
                      {dataSheet[0]?.map((header, i) => (
                        <TableHead key={i} className="h-10 font-bold text-[#3c4043] px-4 whitespace-nowrap bg-[#f1f3f4] border-b-2 border-border sticky top-0 z-20">{header}</TableHead>
                      ))}
                    </TableRow>
                  </TableHeader>
                  <TableBody>
                    {loading ? (
                      <TableRow>
                        <TableCell colSpan={dataSheet[0]?.length || 1} className="h-24 text-center">
                          <RefreshCw className="w-5 h-5 animate-spin mx-auto text-primary/30" />
                        </TableCell>
                      </TableRow>
                    ) : (
                      filteredDataSheet.map((row, i) => (
                        <TableRow key={i} className="hover:bg-blue-50/30 border-b border-border transition-colors">
                          {row.map((cell, j) => {
                            const headerName = dataSheet[0]?.[j]?.toLowerCase() || "";
                            const isLongText = headerName.includes("giải pháp") || 
                                             headerName.includes("vướng mắc") || 
                                             headerName.includes("đề xuất") ||
                                             headerName.includes("nguyên nhân") ||
                                             headerName.includes("kế hoạch");
                            
                            const isMediumText = headerName.includes("tên trạm") || 
                                               headerName.includes("đơn vị") ||
                                               headerName.includes("điện lực") ||
                                               headerName.includes("tiến độ");

                            return (
                              <TableCell key={j} className={cn(
                                "py-3 px-4 border-b border-border",
                                isLongText ? "min-w-[300px] max-w-[500px] whitespace-normal break-words leading-relaxed" : 
                                isMediumText ? "min-w-[150px] whitespace-normal" : "whitespace-nowrap"
                              )}>
                                {cell}
                              </TableCell>
                            );
                          })}
                        </TableRow>
                      ))
                    )}
                  </TableBody>
                </table>
              </CardContent>
            </Card>
          </motion.div>
        </section>
      </>
    ) : activeTab === "tong-hop" ? (
      /* NEW: Tong Hop View */
      <section className="flex-1 flex flex-col gap-4 overflow-hidden w-full">
        {/* Comparison Summary Card */}
        <motion.div
          initial={{ opacity: 0, y: -10 }}
          animate={{ opacity: 1, y: 0 }}
          className="flex-none"
        >
          <Card className="shadow-sm border-border bg-white overflow-hidden">
            <CardHeader className="bg-slate-50 border-b border-border py-2 px-4">
              <CardTitle className="text-[14px] font-bold text-slate-700 flex items-center gap-2 uppercase tracking-wider">
                <TableIcon className="w-4 h-4 text-primary" />
                Tổng hợp so sánh kế hoạch và thực hiện
              </CardTitle>
            </CardHeader>
            <CardContent className="p-4">
              <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                {[
                  { label: "Sửa chữa thường xuyên (SCTX)", data: planComparison.total.SCTX, color: "blue" },
                  { label: "Sửa chữa lớn (SCL)", data: planComparison.total.SCL, color: "emerald" },
                  { label: "Đầu tư xây dựng (ĐTXD)", data: planComparison.total.DTXD, color: "orange" }
                ].map((item, idx) => (
                  <div key={idx} className={cn(
                    "p-3 rounded-lg border flex flex-col gap-2 transition-all hover:shadow-md",
                    item.color === "blue" ? "bg-blue-50/30 border-blue-100" :
                    item.color === "emerald" ? "bg-emerald-50/30 border-emerald-100" :
                    "bg-orange-50/30 border-orange-100"
                  )}>
                    <div className="flex items-center justify-between">
                      <span className={cn(
                        "text-[12px] font-bold uppercase",
                        item.color === "blue" ? "text-blue-700" :
                        item.color === "emerald" ? "text-emerald-700" :
                        "text-orange-700"
                      )}>{item.label}</span>
                    </div>
                    <div className="grid grid-cols-2 gap-2 mt-1">
                      <div className="flex flex-col">
                        <span className="text-[10px] text-slate-400 font-bold uppercase">Tổng kế hoạch</span>
                        <span className="text-[18px] font-bold text-slate-700">{item.data.plan} <span className="text-[10px] text-slate-400">trạm</span></span>
                      </div>
                      <div className="flex flex-col">
                        <span className="text-[10px] text-slate-400 font-bold uppercase">Đã thực hiện</span>
                        <div className="flex items-baseline gap-1">
                          <span className={cn(
                            "text-[18px] font-bold",
                            item.color === "blue" ? "text-blue-600" :
                            item.color === "emerald" ? "text-emerald-600" :
                            "text-orange-600"
                          )}>{item.data.actual}</span>
                          <span className="text-[10px] text-slate-400">trạm</span>
                        </div>
                      </div>
                    </div>
                    <div className="mt-2 w-full bg-slate-200 h-1.5 rounded-full overflow-hidden">
                      <div 
                        className={cn(
                          "h-full rounded-full transition-all duration-1000",
                          item.color === "blue" ? "bg-blue-500" :
                          item.color === "emerald" ? "bg-emerald-500" :
                          "bg-orange-500"
                        )}
                        style={{ width: `${item.data.plan > 0 ? Math.min(100, (item.data.actual / item.data.plan) * 100) : 0}%` }}
                      />
                    </div>
                    <div className="flex justify-between items-center text-[10px] font-bold text-slate-500">
                       <span>Tỷ lệ hoàn thành:</span>
                       <span>{item.data.plan > 0 ? ((item.data.actual / item.data.plan) * 100).toFixed(1) : "0"}%</span>
                    </div>
                  </div>
                ))}
              </div>
            </CardContent>
          </Card>
        </motion.div>

        {/* NEW: Per-Unit Comparison Table */}
        <motion.div
          initial={{ opacity: 0, y: 10 }}
          animate={{ opacity: 1, y: 0 }}
          className="flex-none"
        >
          <Card className="shadow-sm border-border overflow-hidden bg-white">
            <CardHeader className="py-2 px-4 bg-slate-50 border-b border-border">
              <div className="flex items-center justify-between gap-4">
                <div className="w-8 shrink-0" />
                <CardTitle className="text-[14px] font-bold text-slate-700 flex items-center justify-center gap-2 uppercase tracking-wider text-center">
                  <TableIcon className="w-4 h-4 text-emerald-600" />
                  Tổng hợp kế hoạch & thực hiện theo Điện lực
                </CardTitle>
                <Button 
                  variant="outline" 
                  size="sm" 
                  onClick={exportPlanByUnitToExcel}
                  className="h-8 px-2 text-[11px] font-bold text-emerald-600 border-emerald-200 hover:bg-emerald-50 hover:text-emerald-700 shrink-0"
                >
                  <Download className="w-3.5 h-3.5 mr-1" />
                  Xuất Excel
                </Button>
              </div>
            </CardHeader>
            <CardContent className="p-0 overflow-x-auto">
              <Table>
                <TableHeader>
                  <TableRow className="bg-slate-50/50">
                    <TableHead className="text-[11px] font-bold uppercase text-slate-500 py-3 px-4">Điện lực</TableHead>
                    <TableHead className="text-[11px] font-bold uppercase text-center text-blue-600 py-3 px-2 border-l" colSpan={2}>SCTX</TableHead>
                    <TableHead className="text-[11px] font-bold uppercase text-center text-emerald-600 py-3 px-2 border-l" colSpan={2}>SCL</TableHead>
                    <TableHead className="text-[11px] font-bold uppercase text-center text-orange-600 py-3 px-2 border-l" colSpan={2}>ĐTXD</TableHead>
                  </TableRow>
                  <TableRow className="bg-slate-50/30">
                    <TableHead className="py-2 px-4"></TableHead>
                    <TableHead className="text-[10px] text-center font-medium text-slate-400 border-l">KH</TableHead>
                    <TableHead className="text-[10px] text-center font-medium text-slate-400">TH</TableHead>
                    <TableHead className="text-[10px] text-center font-medium text-slate-400 border-l">KH</TableHead>
                    <TableHead className="text-[10px] text-center font-medium text-slate-400">TH</TableHead>
                    <TableHead className="text-[10px] text-center font-medium text-slate-400 border-l">KH</TableHead>
                    <TableHead className="text-[10px] text-center font-medium text-slate-400">TH</TableHead>
                  </TableRow>
                </TableHeader>
                <TableBody>
                  {planComparison.byUnit.map((u, i) => (
                    <TableRow key={i} className="hover:bg-slate-50 transition-colors">
                      <TableCell className="py-2 px-4 font-bold text-[12px] text-slate-700">{u.unit}</TableCell>
                      <TableCell className="py-2 px-2 text-center text-[12px] text-slate-600 border-l">{u.SCTX.plan}</TableCell>
                      <TableCell className="py-2 px-2 text-center text-[12px] font-bold text-blue-600">{u.SCTX.actual}</TableCell>
                      <TableCell className="py-2 px-2 text-center text-[12px] text-slate-600 border-l">{u.SCL.plan}</TableCell>
                      <TableCell className="py-2 px-2 text-center text-[12px] font-bold text-emerald-600">{u.SCL.actual}</TableCell>
                      <TableCell className="py-2 px-2 text-center text-[12px] text-slate-600 border-l">{u.DTXD.plan}</TableCell>
                      <TableCell className="py-2 px-2 text-center text-[12px] font-bold text-orange-600">{u.DTXD.actual}</TableCell>
                    </TableRow>
                  ))}
                  <TableRow className="bg-slate-100/50 font-bold">
                    <TableCell className="py-3 px-4 text-[12px] uppercase">TỔNG CỘNG</TableCell>
                    <TableCell className="py-3 px-2 text-center text-[13px] border-l">{planComparison.total.SCTX.plan}</TableCell>
                    <TableCell className="py-3 px-2 text-center text-[13px] text-blue-700">{planComparison.total.SCTX.actual}</TableCell>
                    <TableCell className="py-3 px-2 text-center text-[13px] border-l">{planComparison.total.SCL.plan}</TableCell>
                    <TableCell className="py-3 px-2 text-center text-[13px] text-emerald-700">{planComparison.total.SCL.actual}</TableCell>
                    <TableCell className="py-3 px-2 text-center text-[13px] border-l">{planComparison.total.DTXD.plan}</TableCell>
                    <TableCell className="py-3 px-2 text-center text-[13px] text-orange-700">{planComparison.total.DTXD.actual}</TableCell>
                  </TableRow>
                </TableBody>
              </Table>
            </CardContent>
          </Card>
        </motion.div>

        <motion.div
          initial={{ opacity: 0, y: 10 }}
          animate={{ opacity: 1, y: 0 }}
          className="flex-none"
        >
          <Card className="shadow-sm border-border">
            <CardHeader className="py-4">
              <CardTitle className="text-[16px] font-bold">Kết quả thực hiện từng trạm</CardTitle>
            </CardHeader>
            <CardContent className="space-y-4">
              <div className="flex flex-wrap items-center gap-4">
                <div className="flex items-center gap-2">
                  <Label className="whitespace-nowrap font-bold text-[#1a73e8]">Điện lực:</Label>
                  <Popover open={openDienLucTongHop} onOpenChange={setOpenDienLucTongHop}>
                    <PopoverTrigger 
                      render={
                        <Button
                          type="button"
                          variant="outline"
                          role="combobox"
                          aria-expanded={openDienLucTongHop}
                          className="w-[200px] h-10 justify-between bg-white border-border text-[13px]"
                        >
                          {selectedDienLucTongHop === "all" ? "Tất cả điện lực" : selectedDienLucTongHop}
                          <ChevronDown className="ml-2 h-4 w-4 shrink-0 opacity-50" />
                        </Button>
                      }
                    />
                    <PopoverContent 
                      className="w-[200px] p-0 bg-white shadow-xl z-[100]"
                      align="start"
                    >
                      <Command>
                        <CommandInput placeholder="Tìm điện lực..." className="h-9" autoFocus={false} />
                        <CommandList>
                          <CommandEmpty>Không tìm thấy.</CommandEmpty>
                          <CommandGroup>
                            <CommandItem
                              value="all"
                              onSelect={() => {
                                setSelectedDienLucTongHop("all");
                                setSelectedStationTongHop("");
                                setOpenDienLucTongHop(false);
                              }}
                              className="cursor-pointer"
                            >
                              <Check className={cn("mr-2 h-4 w-4", selectedDienLucTongHop === "all" ? "opacity-100" : "opacity-0")} />
                              Tất cả điện lực
                            </CommandItem>
                            {donViOptionsTongHop.map((name) => (
                              <CommandItem
                                key={name}
                                value={name}
                                onSelect={() => {
                                  setSelectedDienLucTongHop(name);
                                  setSelectedStationTongHop("");
                                  setOpenDienLucTongHop(false);
                                }}
                                className="cursor-pointer"
                              >
                                <Check className={cn("mr-2 h-4 w-4", selectedDienLucTongHop === name ? "opacity-100" : "opacity-0")} />
                                {name}
                              </CommandItem>
                            ))}
                          </CommandGroup>
                        </CommandList>
                      </Command>
                    </PopoverContent>
                  </Popover>
                </div>

                <div className="flex items-center gap-2">
                  <Label className="whitespace-nowrap font-bold text-[#1a73e8]">Chọn trạm:</Label>
                  <Popover open={openSearchTongHop} onOpenChange={setOpenSearchTongHop}>
                    <PopoverTrigger 
                      render={
                        <Button
                          type="button"
                          variant="outline"
                          role="combobox"
                          aria-expanded={openSearchTongHop}
                          className="w-[300px] h-10 justify-between bg-white border-border text-[13px]"
                        >
                          {selectedStationTongHop ? selectedStationTongHop : "Tìm tên trạm..."}
                          <Search className="ml-2 h-4 w-4 shrink-0 opacity-50" />
                        </Button>
                      }
                    />
                    <PopoverContent 
                      className="w-[300px] p-0 bg-white shadow-xl z-[100]"
                      align="start"
                    >
                      <Command>
                        <CommandInput placeholder="Gõ tên trạm..." className="h-9" autoFocus={false} />
                        <CommandList>
                          <CommandEmpty>Không tìm thấy.</CommandEmpty>
                          <CommandGroup>
                            {stationNamesTongHop.map((name) => (
                              <CommandItem
                                key={name}
                                value={name}
                                onSelect={() => {
                                  setSelectedStationTongHop(name);
                                  setOpenSearchTongHop(false);
                                }}
                                className="cursor-pointer"
                              >
                                <Check className={cn("mr-2 h-4 w-4", selectedStationTongHop === name ? "opacity-100" : "opacity-0")} />
                                {name}
                              </CommandItem>
                            ))}
                          </CommandGroup>
                        </CommandList>
                      </Command>
                    </PopoverContent>
                  </Popover>
                </div>
              </div>

              {tongHopDetail ? (
                <div className="space-y-6 animate-in fade-in slide-in-from-top-4 duration-500">
                  {/* Assessment Card */}
                  <div className={cn(
                    "p-4 rounded-lg border-2 flex items-center gap-4",
                    tongHopDetail.assessment.includes("cao") ? "bg-emerald-50 border-emerald-200 text-emerald-800" :
                    tongHopDetail.assessment.includes("đạt kế hoạch") ? "bg-blue-50 border-blue-200 text-blue-800" :
                    tongHopDetail.assessment.includes("trung bình") ? "bg-yellow-50 border-yellow-200 text-yellow-800" :
                    tongHopDetail.assessment.includes("thấp") ? "bg-orange-50 border-orange-200 text-orange-800" :
                    tongHopDetail.assessment.includes("Thiếu") ? "bg-gray-50 border-gray-200 text-gray-500" :
                    "bg-red-50 border-red-200 text-red-800"
                  )}>
                    <CheckCircle2 className="h-6 w-6 shrink-0" />
                    <div>
                      <h3 className="font-bold text-[14px] uppercase">Đánh giá khả năng đạt kế hoạch năm</h3>
                      <p className="font-semibold text-[16px]">{tongHopDetail.assessment}</p>
                    </div>
                  </div>

                  <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
                    {/* Group 1 */}
                    <div className="space-y-3">
                      <h4 className="font-bold text-[13px] text-slate-500 uppercase border-l-4 border-blue-500 pl-2">Thông tin chung</h4>
                      <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                        {tongHopDetail.group1.map((item, idx) => (
                          <div key={idx} className="flex flex-col p-2 bg-white rounded border border-slate-100 shadow-sm">
                            <span className="text-[11px] font-bold text-slate-400 mb-0.5">{item.label}</span>
                            <span className="text-[13px] font-medium text-slate-900">{item.value || "-"}</span>
                          </div>
                        ))}
                      </div>
                    </div>

                    {/* Group 2 */}
                    <div className="space-y-3">
                      <h4 className="font-bold text-[13px] text-slate-500 uppercase border-l-4 border-emerald-500 pl-2">Kế hoạch của Điện lực</h4>
                      <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                        {tongHopDetail.group2.map((item, idx) => (
                          <div key={idx} className="flex flex-col p-2 bg-white rounded border border-slate-100 shadow-sm">
                            <span className="text-[11px] font-bold text-slate-400 mb-0.5">{item.label}</span>
                            <span className="text-[13px] font-medium text-slate-900">{item.value || "-"}</span>
                          </div>
                        ))}
                      </div>
                    </div>

                    {/* Group 4 (Infrastructure/Plan) before Group 3 (Other) */}
                    <div className="space-y-3">
                      <h4 className="font-bold text-[13px] text-slate-500 uppercase border-l-4 border-orange-500 pl-2">Kết quả thực hiện TTĐN lũy kế và đánh giá</h4>
                      <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                        {tongHopDetail.group4.map((item, idx) => (
                          <div key={idx} className="flex flex-col p-2 bg-white rounded border border-slate-100 shadow-sm">
                            <span className="text-[11px] font-bold text-slate-400 mb-0.5">{item.label}</span>
                            <span className="text-[13px] font-medium text-slate-900 font-mono">{item.value || "-"}</span>
                          </div>
                        ))}
                      </div>
                    </div>

                    {/* Group 3 */}
                    <div className="space-y-3">
                      <h4 className="font-bold text-[13px] text-slate-500 uppercase border-l-4 border-purple-500 pl-2">Triển khai công tác</h4>
                      <div className="grid grid-cols-1 gap-3">
                        {tongHopDetail.group3.map((item, idx) => (
                          <div key={idx} className="flex flex-col p-2 bg-white rounded border border-slate-100 shadow-sm">
                            <span className="text-[11px] font-bold text-slate-400 mb-0.5">{item.label}</span>
                            <span className="text-[13px] font-medium text-slate-900">{item.value || "-"}</span>
                          </div>
                        ))}
                      </div>
                    </div>
                  </div>
                </div>
              ) : (
                <div className="text-center py-20 text-slate-400 italic bg-white border-2 border-dashed rounded-lg">
                  Vui lòng chọn Điện lực và Tên trạm để xem đánh giá chi tiết
                </div>
              )}
            </CardContent>
          </Card>
        </motion.div>

        <motion.div
          initial={{ opacity: 0, y: 10 }}
          animate={{ opacity: 1, y: 0 }}
          transition={{ delay: 0.1 }}
          className="flex-1 overflow-hidden"
        >
          <Card className="h-full shadow-sm border-border flex flex-col">
            <CardHeader className="flex flex-row items-center justify-between py-3 px-4 bg-[#fafafa] border-b space-y-0">
              <CardTitle className="text-[14px] font-bold text-[#3c4043]">Danh sách kết quả thực hiện</CardTitle>
              <Button
                variant="outline"
                size="sm"
                className="h-8 text-[12px] text-emerald-600 border-emerald-200"
                onClick={() => exportToExcel(tongHopSheet[0] || [], filteredTongHopData, "Sheet_Tong_Hop")}
              >
                <Download className="w-3.5 h-3.5 mr-1.5" />
                Xuất Excel
              </Button>
            </CardHeader>
            <CardContent className="p-0 flex-1 overflow-auto relative max-h-[600px]">
              <table className="w-full caption-bottom text-[12px] border-separate border-spacing-0">
                <TableHeader className="sticky top-0 z-20 shadow-sm">
                  <TableRow className="hover:bg-transparent bg-[#f1f3f4]">
                    {tongHopSheet[0]?.map((header, i) => (
                      <TableHead key={i} className="h-10 font-bold px-4 border-b border-border text-[#3c4043] whitespace-nowrap bg-[#f1f3f4] sticky top-0 z-20">{header}</TableHead>
                    ))}
                  </TableRow>
                </TableHeader>
                <TableBody>
                  {filteredTongHopData.map((row, i) => (
                    <TableRow key={i} className="hover:bg-blue-50/50 border-b border-border transition-colors">
                      {row.map((cell, j) => (
                        <TableCell key={j} className="px-4 py-2 border-b border-border text-[12px] whitespace-normal min-w-[120px] max-w-[400px] break-words">
                          {cell}
                        </TableCell>
                      ))}
                    </TableRow>
                  ))}
                </TableBody>
              </table>
            </CardContent>
          </Card>
        </motion.div>
      </section>
    ) : (
      <section className="flex-1 flex flex-col gap-4 sm:gap-6 p-2 sm:p-4 overflow-y-auto w-full">
        {/* Statistics Filters */}
        <motion.div
          initial={{ opacity: 0, y: -10 }}
          animate={{ opacity: 1, y: 0 }}
          className="flex flex-wrap items-center gap-2 sm:gap-4 bg-white p-3 sm:p-4 rounded-lg shadow-sm border border-border"
        >
          <div className="flex flex-col sm:flex-row sm:items-center gap-2 w-full sm:w-auto">
            <Label className="whitespace-nowrap font-bold text-[#1a73e8] text-[12px] sm:text-[14px]">Chọn đơn vị thống kê:</Label>
            <Popover open={openDienLucThongKe} onOpenChange={setOpenDienLucThongKe}>
              <PopoverTrigger 
                render={
                  <Button
                    type="button"
                    variant="outline"
                    role="combobox"
                    className="w-full sm:w-[250px] h-9 sm:h-10 justify-between bg-white border-border text-[12px] sm:text-[13px] font-semibold"
                  >
                    {selectedDienLucThongKe === "all" ? "Toàn Công ty" : selectedDienLucThongKe}
                    <ChevronDown className="ml-2 h-4 w-4 shrink-0 opacity-50" />
                  </Button>
                }
              />
              <PopoverContent 
                className="w-[250px] p-0 bg-white shadow-xl z-[100]"
                align="start"
              >
                <Command>
                  <CommandInput placeholder="Tìm đơn vị..." className="h-9" autoFocus={false} />
                  <CommandList>
                    <CommandEmpty>Không tìm thấy.</CommandEmpty>
                    <CommandGroup>
                      <CommandItem
                        value="all"
                        onSelect={() => {
                          setSelectedDienLucThongKe("all");
                          setOpenDienLucThongKe(false);
                        }}
                        className="cursor-pointer"
                      >
                        <Check className={cn("mr-2 h-4 w-4", selectedDienLucThongKe === "all" ? "opacity-100" : "opacity-0")} />
                        Toàn Công ty
                      </CommandItem>
                      {donViOptionsTongHop.map((name) => (
                        <CommandItem
                          key={name}
                          value={name}
                          onSelect={() => {
                            setSelectedDienLucThongKe(name);
                            setOpenDienLucThongKe(false);
                          }}
                          className="cursor-pointer"
                        >
                          <Check className={cn("mr-2 h-4 w-4", selectedDienLucThongKe === name ? "opacity-100" : "opacity-0")} />
                          {name}
                        </CommandItem>
                      ))}
                    </CommandGroup>
                  </CommandList>
                </Command>
              </PopoverContent>
            </Popover>
          </div>
        </motion.div>

        <motion.div
          initial={{ opacity: 0, scale: 0.98 }}
          animate={{ opacity: 1, scale: 1 }}
          transition={{ duration: 0.4 }}
          className="grid grid-cols-1 xl:grid-cols-3 gap-6"
        >
          {/* Table Summary */}
          <Card className="xl:col-span-3 shadow-sm border-border bg-white">
            <CardHeader className="bg-slate-50 border-b border-border py-2 sm:py-3 px-3 sm:px-4 flex flex-col sm:flex-row items-start sm:items-center justify-between space-y-2 sm:space-y-0">
              <CardTitle className="text-[14px] sm:text-[16px] font-bold text-slate-800 flex items-center gap-2">
                <TableIcon className="w-4 h-4 sm:w-5 h-5 text-[#1a73e8]" />
                Thống kê {selectedDienLucThongKe !== "all" ? `- ${selectedDienLucThongKe}` : ""}
              </CardTitle>
              <div className="flex flex-wrap items-center gap-1.5 sm:gap-2">
                <Dialog>
                  <DialogTrigger
                    render={
                      <Button
                        variant="outline"
                        size="sm"
                        className="h-7 sm:h-8 text-[11px] sm:text-[12px] bg-white border-border text-blue-600 hover:text-blue-700 hover:bg-blue-50 px-2"
                      >
                        <HelpCircle className="w-3 h-3 sm:w-3.5 h-3.5 mr-1 sm:mr-1.5" />
                        HD tính toán
                      </Button>
                    }
                  />
                  <DialogContent className="max-w-2xl max-h-[90vh] overflow-y-auto">
                  <DialogHeader>
                      <DialogTitle className="text-lg sm:text-xl font-bold flex items-center gap-2">
                        <Info className="w-5 h-5 sm:w-6 h-6 text-blue-600" />
                        Phương pháp tính & chấm điểm
                      </DialogTitle>
                      <DialogDescription className="text-xs sm:text-sm">
                        Xác định mức đánh giá và công thức chấm điểm hiệu quả thực hiện.
                      </DialogDescription>
                    </DialogHeader>

                    <div className="space-y-6 py-4">
                      <section className="space-y-3">
                        <h3 className="text-lg font-bold text-slate-800 border-l-4 border-blue-600 pl-3">
                          1. Xác định mức đánh giá khả năng thực hiện kế hoạch
                        </h3>
                        <p className="text-sm text-slate-600">
                          Dựa trên giá trị dữ liệu từ danh mục tổng hợp (TTĐN), hệ thống tự động phân loại trạm theo các ngưỡng sau:
                        </p>
                        <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
                          {[
                            { label: "Dự kiến đạt kế hoạch cao", condition: "≥ 100", color: "text-emerald-600" },
                            { label: "Dự kiến đạt kế hoạch", condition: "95 - 100", color: "text-blue-600" },
                            { label: "Khả năng đạt kế hoạch ở mức trung bình", condition: "90 - 95", color: "text-amber-600" },
                            { label: "Khả năng đạt kế hoạch ở mức thấp", condition: "70 - 90", color: "text-orange-600" },
                            { label: "Khả năng không đạt kế hoạch", condition: "< 70", color: "text-red-600" },
                            { label: "Thiếu dữ liệu để đánh giá", condition: "Không có số liệu", color: "text-slate-400" },
                          ].map((item, i) => (
                            <div key={i} className="flex justify-between items-center p-2 rounded bg-slate-50 border border-slate-100 italic">
                              <span className={cn("text-[13px] font-medium", item.color)}>{item.label}</span>
                              <span className="text-xs font-bold text-slate-500">{item.condition}</span>
                            </div>
                          ))}
                        </div>
                      </section>

                      <section className="space-y-3">
                        <h3 className="text-lg font-bold text-slate-800 border-l-4 border-emerald-600 pl-3">
                          2. Xác định điểm chấm các Điện lực (Điểm hiệu quả)
                        </h3>
                        <p className="text-sm text-slate-600">
                          Điểm hiệu quả của mỗi Điện lực được tính toán theo công thức trọng số dựa trên tỷ lệ (%) thực hiện và tỷ lệ các mức đánh giá:
                        </p>
                        <div className="bg-slate-900 text-slate-100 p-4 rounded-lg font-mono text-[13px] leading-relaxed shadow-inner overflow-x-auto">
                          <p className="text-emerald-400 font-bold mb-2">// Công thức chấm điểm:</p>
                          <p>Điểm = (Tỷ lệ % Đã thực hiện × 2)</p>
                          <p className="pl-4">+ (Tỷ lệ % Dự kiến đạt kế hoạch cao × 2)</p>
                          <p className="pl-4">+ (Tỷ lệ % Dự kiến đạt kế hoạch × 1.5)</p>
                          <p className="pl-4">+ (Tỷ lệ % Khả năng đạt KH mức TB × 1.0)</p>
                          <p className="pl-4">+ (Tỷ lệ % Khả năng đạt KH mức thấp × 0.5)</p>
                          <p className="pl-4 text-red-500">- (Tỷ lệ % Khả năng không đạt KH × 2.0)</p>
                          <p className="pl-4 text-red-500">- (Tỷ lệ % Thiếu dữ liệu đánh giá × 1.0)</p>
                        </div>
                        <div className="pt-2">
                          <p className="text-[11px] text-slate-400 italic">
                            * Ghi chú: Tỷ lệ (%) được lấy trên tổng số trạm của Điện lực đó.
                          </p>
                        </div>
                      </section>
                      
                      <section className="space-y-3">
                        <h3 className="text-md font-bold text-slate-800">Bảng xếp hạng & Phân loại</h3>
                        <div className="grid grid-cols-4 gap-2">
                          {[
                            { label: "Xuất sắc", range: "≥ 150", color: "bg-emerald-50 text-emerald-700 border-emerald-200" },
                            { label: "Tốt", range: "100 - 150", color: "bg-blue-50 text-blue-700 border-blue-200" },
                            { label: "Trung bình", range: "50 - 100", color: "bg-yellow-50 text-yellow-700 border-yellow-200" },
                            { label: "Cần cải thiện", range: "< 50", color: "bg-red-50 text-red-700 border-red-200" },
                          ].map((cat, i) => (
                            <div key={i} className={cn("text-center py-2 rounded-md border", cat.color)}>
                              <p className="text-[11px] font-bold">{cat.label}</p>
                              <p className="text-[10px]">{cat.range}</p>
                            </div>
                          ))}
                        </div>
                      </section>
                    </div>
                  </DialogContent>
                </Dialog>

                <Button
                  variant="outline"
                  size="sm"
                  className="h-8 text-[12px] bg-white border-border text-emerald-600 hover:text-emerald-700 hover:bg-emerald-50"
                  onClick={exportStatisticsToExcel}
                >
                  <Download className="w-3.5 h-3.5 mr-1.5" />
                  Xuất Excel
                </Button>
              </div>
            </CardHeader>
            <CardContent className="p-0 overflow-hidden flex flex-col h-[600px]">
              <div className="overflow-auto relative flex-1">
                <Table className="border-separate border-spacing-0">
                  <TableHeader className="sticky top-0 z-30 shadow-sm bg-slate-50">
                    <TableRow className="hover:bg-transparent">
                      <TableHead className="font-bold text-slate-700 min-w-[120px] sm:min-w-[150px] sticky left-0 top-0 bg-slate-50 z-40 border-b border-border text-[10px] sm:text-[11px]">Điện lực</TableHead>
                      <TableHead className="text-center font-bold text-slate-700 sticky top-0 bg-slate-50 z-30 border-b border-border whitespace-normal leading-tight h-12 uppercase text-[10px] sm:text-[11px]">Tổng số trạm</TableHead>
                      <TableHead className="text-center font-bold text-slate-700 sticky top-0 bg-slate-50 z-30 border-b border-border whitespace-normal leading-tight h-12 uppercase text-[10px] sm:text-[11px]">Đã thực hiện</TableHead>
                      {Object.keys(STATUS_COLORS).map(status => (
                        <TableHead key={status} className="text-center font-bold text-slate-700 min-w-[80px] sm:min-w-[100px] sticky top-0 bg-slate-50 z-30 border-b border-border whitespace-normal leading-tight h-12 uppercase text-[10px] sm:text-[11px]">
                          {status}
                        </TableHead>
                      ))}
                    </TableRow>
                  </TableHeader>
                  <TableBody>
                    {statisticsData
                      .filter(s => selectedDienLucThongKe === "all" || s.company === selectedDienLucThongKe || s.company === "TỔNG CỘNG")
                      .map((stat, idx) => (
                      <TableRow key={idx} className={cn(
                        "hover:bg-slate-50 transition-colors",
                        stat.company === "TỔNG CỘNG" ? "bg-slate-100 font-bold border-t-2 border-slate-300" : ""
                      )}>
                        <TableCell className="font-medium sticky left-0 bg-white group-hover:bg-slate-50 text-[11px] sm:text-[12px]">{stat.company}</TableCell>
                        <TableCell className="text-center font-semibold text-[11px] sm:text-[12px]">{stat.totalStations}</TableCell>
                        <TableCell className="text-center">
                          <div className="flex flex-col items-center">
                            <span className="font-semibold text-emerald-600 text-[11px] sm:text-[12px]">{stat.implementedCount}</span>
                            <span className="text-[9px] sm:text-[10px] text-slate-400">({stat.totalStations > 0 ? (stat.implementedCount / stat.totalStations * 100).toFixed(1) : 0}%)</span>
                          </div>
                        </TableCell>
                        {Object.keys(STATUS_COLORS).map(status => (
                          <TableCell key={status} className="text-center">
                            <div className="flex flex-col items-center">
                              <span className="font-semibold text-[11px] sm:text-[12px]" style={{ color: STATUS_COLORS[status] }}>
                                {stat.counts[status as keyof typeof stat.counts]}
                              </span>
                              <span className="text-[9px] sm:text-[10px] text-slate-400">
                                ({stat.percentages[status].toFixed(1)}%)
                              </span>
                            </div>
                          </TableCell>
                        ))}
                      </TableRow>
                    ))}
                  </TableBody>
                </Table>
              </div>
            </CardContent>
          </Card>

          {/* Charts section */}
          <Card className="xl:col-span-2 shadow-sm border-border bg-white flex flex-col h-[500px]">
            <CardHeader className="py-2.5 sm:py-4 border-b border-border">
              <CardTitle className="text-[14px] sm:text-[16px] font-bold">Biểu đồ kết quả theo Điện lực</CardTitle>
            </CardHeader>
            <CardContent className="pt-6 flex-1">
              <ResponsiveContainer width="100%" height="100%">
                <BarChart
                  data={statisticsData.filter(s => s.company !== "TỔNG CỘNG" && (selectedDienLucThongKe === "all" || s.company === selectedDienLucThongKe))}
                  margin={{ top: 20, right: 10, left: 0, bottom: 60 }}
                >
                  <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f1f5f9" />
                  <XAxis 
                    dataKey="company" 
                    angle={-45} 
                    textAnchor="end" 
                    interval={0} 
                    height={80}
                    tick={{ fontSize: 9, fill: '#64748b' }}
                  />
                  <YAxis tick={{ fontSize: 10, fill: '#64748b' }} width={30} />
                  <Tooltip 
                    content={<CustomTooltip />}
                    cursor={{ fill: '#f8fafc' }}
                  />
                  {Object.keys(STATUS_COLORS).map(status => (
                    <Bar 
                      key={status}
                      name={status}
                      dataKey={`counts.${status}`} 
                      stackId="a" 
                      fill={STATUS_COLORS[status]} 
                    />
                  ))}
                </BarChart>
              </ResponsiveContainer>
            </CardContent>
          </Card>

          <Card className="shadow-sm border-border bg-white flex flex-col h-[500px]">
            <CardHeader className="py-4 border-b border-border">
              <CardTitle className="text-[16px] font-bold">
                Cơ cấu đánh giá {selectedDienLucThongKe === "all" ? "toàn Công ty" : `- ${selectedDienLucThongKe}`}
              </CardTitle>
            </CardHeader>
            <CardContent className="pt-6 flex flex-col items-center justify-between flex-1">
              <div className="h-[280px] w-full">
                <ResponsiveContainer width="100%" height="100%">
                  <PieChart>
                    <Pie
                      data={Object.entries(
                        (selectedDienLucThongKe === "all" 
                          ? statisticsData.find(s => s.company === "TỔNG CỘNG") 
                          : statisticsData.find(s => s.company === selectedDienLucThongKe)
                        )?.counts || {}
                      ).map(([name, value]) => ({ name, value }))}
                      cx="50%"
                      cy="50%"
                      innerRadius={60}
                      outerRadius={90}
                      paddingAngle={5}
                      dataKey="value"
                      label={({ name, value }) => value > 0 ? `${value}` : ""}
                    >
                      {Object.keys(STATUS_COLORS).map((status, index) => (
                        <Cell key={`cell-${index}`} fill={STATUS_COLORS[status]} />
                      ))}
                    </Pie>
                    <Tooltip 
                      content={<CustomTooltip />}
                    />
                  </PieChart>
                </ResponsiveContainer>
              </div>
              <div className="grid grid-cols-2 gap-4 w-full pt-4 border-t border-slate-100">
                <div className="p-3 bg-blue-50 rounded-lg text-center">
                  <p className="text-[11px] text-blue-600 font-bold uppercase tracking-wider">Tổng số trạm</p>
                  <p className="text-[20px] font-bold text-blue-900 mt-1">
                    {(selectedDienLucThongKe === "all" 
                      ? statisticsData.find(s => s.company === "TỔNG CỘNG") 
                      : statisticsData.find(s => s.company === selectedDienLucThongKe)
                    )?.totalStations || 0}
                  </p>
                </div>
                <div className="p-3 bg-emerald-50 rounded-lg text-center">
                  <p className="text-[11px] text-emerald-600 font-bold uppercase tracking-wider">Đã thực hiện</p>
                  <p className="text-[20px] font-bold text-emerald-900 mt-1">
                    {(selectedDienLucThongKe === "all" 
                      ? statisticsData.find(s => s.company === "TỔNG CỘNG") 
                      : statisticsData.find(s => s.company === selectedDienLucThongKe)
                    )?.implementedCount || 0}
                  </p>
                </div>
              </div>
            </CardContent>
          </Card>

          {/* Scoring Table */}
          <Card className="xl:col-span-3 shadow-sm border-border bg-white mt-6">
            <CardHeader className="bg-slate-50 border-b border-border py-2 sm:py-3 px-3 sm:px-4 flex flex-col sm:flex-row items-start sm:items-center justify-between space-y-2 sm:space-y-0">
              <CardTitle className="text-[14px] sm:text-[16px] font-bold text-slate-800 flex items-center gap-2">
                <CheckCircle2 className="w-4 h-4 sm:w-5 h-5 text-emerald-600" />
                Bảng chấm điểm hiệu quả
              </CardTitle>
              <div className="text-[10px] sm:text-[11px] font-medium text-slate-500 italic">
                * Sắp xếp theo điểm số từ cao đến thấp
              </div>
            </CardHeader>
            <CardContent className="p-0 overflow-hidden">
              <div className="overflow-x-auto">
                <Table>
                  <TableHeader>
                    <TableRow className="bg-slate-100 hover:bg-slate-100">
                      <TableHead className="w-[60px] sm:w-[80px] text-center font-bold px-2 text-[11px] sm:text-[13px]">Thứ hạng</TableHead>
                      <TableHead className="font-bold px-2 text-[11px] sm:text-[13px]">Điện lực</TableHead>
                      <TableHead className="text-center font-bold px-2 text-[11px] sm:text-[13px]">Điểm số</TableHead>
                      <TableHead className="text-center font-bold hidden md:table-cell px-2 text-[11px] sm:text-[13px]">Trạng thái hiệu quả</TableHead>
                    </TableRow>
                  </TableHeader>
                  <TableBody>
                    {scoringData.map((stat, idx) => (
                      <TableRow key={idx} className="hover:bg-slate-50 transition-colors">
                        <TableCell className="text-center font-bold px-2 py-2.5 sm:py-4">
                          {idx === 0 ? <span className="text-base sm:text-xl">🥇</span> : 
                           idx === 1 ? <span className="text-base sm:text-xl">🥈</span> :
                           idx === 2 ? <span className="text-base sm:text-xl">🥉</span> : <span className="text-[11px] sm:text-[14px]">{idx + 1}</span>}
                        </TableCell>
                        <TableCell className="font-semibold text-slate-700 px-2 py-2.5 sm:py-4 text-[11px] sm:text-[14px]">{stat.company}</TableCell>
                        <TableCell className="text-center font-bold text-sm sm:text-lg text-blue-600 px-2 py-2.5 sm:py-4">
                          {stat.score.toFixed(1)}
                        </TableCell>
                        <TableCell className="text-center hidden md:table-cell px-2 py-3 sm:py-4">
                          <div className={cn(
                            "inline-flex items-center px-2.5 py-0.5 rounded-full text-xs font-medium border",
                            stat.score >= 150 ? "bg-emerald-50 text-emerald-700 border-emerald-200" :
                            stat.score >= 100 ? "bg-blue-50 text-blue-700 border-blue-200" :
                            stat.score >= 50 ? "bg-yellow-50 text-yellow-700 border-yellow-200" :
                            "bg-red-50 text-red-700 border-red-200"
                          )}>
                            {stat.score >= 150 ? "Xuất sắc" :
                             stat.score >= 100 ? "Tốt" :
                             stat.score >= 50 ? "Trung bình" : "Cần cải thiện"}
                          </div>
                        </TableCell>
                      </TableRow>
                    ))}
                  </TableBody>
                </Table>
              </div>
            </CardContent>
          </Card>
        </motion.div>
      </section>
    )}
  </main>
    </div>
  );
}
