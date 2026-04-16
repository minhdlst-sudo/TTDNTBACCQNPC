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
  Download
} from "lucide-react";
import * as XLSX from "xlsx";
import { format } from "date-fns";
import { vi } from "date-fns/locale";
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

export default function App() {
  const [dataSheet, setDataSheet] = useState<SheetData>([]);
  const [capNhatSheet, setCapNhatSheet] = useState<SheetData>([]);
  const [thuVienSheet, setThuVienSheet] = useState<SheetData>([]);
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

  // Searchable Select state
  const [openTenTram, setOpenTenTram] = useState(false);
  const [openFilterCapNhat, setOpenFilterCapNhat] = useState(false);
  const [openFilterData, setOpenFilterData] = useState(false);
  const [openFilterDonVi, setOpenFilterDonVi] = useState(false);
  const [openDienLuc, setOpenDienLuc] = useState(false);

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

  const [lastSync, setLastSync] = useState<Date | null>(null);

  const fetchData = async () => {
    setLoading(true);
    setError(null);
    try {
      // Add timestamp to bypass browser cache
      const ts = Date.now();
      const [dataRes, capNhatRes, thuVienRes] = await Promise.all([
        fetch(`/api/sheets/data?t=${ts}`),
        fetch(`/api/sheets/cap-nhat?t=${ts}`),
        fetch(`/api/sheets/thu-vien?t=${ts}`)
      ]);

      if (!dataRes.ok || !capNhatRes.ok || !thuVienRes.ok) {
        const dataErr = await dataRes.json();
        throw new Error(dataErr.error || "Failed to fetch data from sheets");
      }

      const data = await dataRes.json();
      const capNhat = await capNhatRes.json();
      const thuVien = await thuVienRes.json();

      console.log("Data fetched:", { 
        dataRows: data.length, 
        capNhatRows: capNhat.length, 
        thuVienRows: thuVien.length 
      });

      setDataSheet(data);
      setCapNhatSheet(capNhat);
      setThuVienSheet(thuVien);
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
          .map(row => row[info.indexTenTram])
          .filter(Boolean)
          .map(s => s.trim())
      )
    ).sort();
  }, [thuVienSheet, dataSheet, dienLuc]);

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
    
    return result;
  }, [capNhatSheet, filterTenTram, dienLuc]);

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
      <header className="bg-gradient-to-r from-[#1a73e8] to-[#0d47a1] border-b border-primary/20 h-[64px] flex items-center justify-between px-6 flex-shrink-0 shadow-md z-20">
        <div className="flex items-center gap-3">
          <div className="bg-white/10 p-2 rounded-lg backdrop-blur-sm">
            <Database className="text-white w-6 h-6" />
          </div>
          <h1 className="font-bold text-[20px] text-white tracking-tight uppercase drop-shadow-sm">Quản lý TTĐN TBA công cộng - QNPC</h1>
        </div>
        <div className="flex items-center gap-4">
          <div className="flex flex-col items-end hidden md:flex">
            <div className="text-[12px] text-white/80 font-medium">Chuyên viên theo dõi: Đặng Xuân Duy - Phòng kỹ thuật</div>
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

      <main className="flex-1 flex flex-col md:flex-row p-4 gap-4 overflow-hidden">
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
                      <PopoverContent className="w-[var(--radix-popover-trigger-width)] p-0 bg-white shadow-xl border-border" align="start">
                        <Command>
                          <CommandInput placeholder="Tìm điện lực..." className="h-9 text-[13px]" />
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
                      <PopoverContent className="w-[var(--radix-popover-trigger-width)] p-0 bg-white shadow-xl border-border" align="start">
                        <Command>
                          <CommandInput placeholder="Nhập tên trạm để tìm..." className="h-9 text-[13px]" />
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
                          initialFocus
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
                      <SelectContent className="bg-white">
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
                    <PopoverContent className="w-[var(--radix-popover-trigger-width)] p-0 bg-white shadow-xl border-border" align="end">
                      <Command>
                        <CommandInput placeholder="Tìm trạm..." className="h-8 text-[12px]" />
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
                  Dữ liệu tổng hợp
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
                  
                  {/* Filter Don Vi */}
                  <Popover open={openFilterDonVi} onOpenChange={setOpenFilterDonVi}>
                    <PopoverTrigger 
                      render={
                        <Button
                          variant="outline"
                          role="combobox"
                          className="w-[150px] h-8 text-[12px] bg-white border-border justify-between"
                        >
                          {filterDonVi === "all" ? "Lọc Đơn vị" : filterDonVi}
                          <ChevronDown className="ml-2 h-3 w-3 shrink-0 opacity-50" />
                        </Button>
                      }
                    />
                    <PopoverContent className="w-[var(--radix-popover-trigger-width)] p-0 bg-white shadow-xl border-border" align="end">
                      <Command>
                        <CommandInput placeholder="Tìm đơn vị..." className="h-8 text-[12px]" />
                        <CommandList>
                          <CommandEmpty>Không tìm thấy.</CommandEmpty>
                          <CommandGroup>
                            <CommandItem
                              value="all"
                              onSelect={() => {
                                setFilterDonVi("all");
                                setOpenFilterDonVi(false);
                              }}
                              className="text-[12px] cursor-pointer"
                            >
                              <Check className={cn("mr-2 h-3 w-3", filterDonVi === "all" ? "opacity-100" : "opacity-0")} />
                              Tất cả đơn vị
                            </CommandItem>
                            {donViOptions.map((name) => (
                              <CommandItem
                                key={name}
                                value={name}
                                onSelect={(val) => {
                                  setFilterDonVi(val);
                                  setOpenFilterDonVi(false);
                                }}
                                className="text-[12px] cursor-pointer"
                              >
                                <Check className={cn("mr-2 h-3 w-3", filterDonVi === name ? "opacity-100" : "opacity-0")} />
                                {name}
                              </CommandItem>
                            ))}
                          </CommandGroup>
                        </CommandList>
                      </Command>
                    </PopoverContent>
                  </Popover>

                  {/* Filter Ten Tram */}
                  <Popover open={openFilterData} onOpenChange={setOpenFilterData}>
                    <PopoverTrigger 
                      render={
                        <Button
                          variant="outline"
                          role="combobox"
                          className="w-[180px] h-8 text-[12px] bg-white border-border justify-between"
                        >
                          {filterDataTenTram === "all" ? "Lọc Tên trạm" : filterDataTenTram}
                          <Search className="ml-2 h-3 w-3 shrink-0 opacity-50" />
                        </Button>
                      }
                    />
                    <PopoverContent className="w-[var(--radix-popover-trigger-width)] p-0 bg-white shadow-xl border-border" align="end">
                      <Command>
                        <CommandInput placeholder="Tìm trạm..." className="h-8 text-[12px]" />
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
      </main>
    </div>
  );
}
