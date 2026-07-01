import express from "express";
import { google } from "googleapis";
import path from "path";
import { fileURLToPath } from "url";
import dotenv from "dotenv";

dotenv.config();

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
app.use(express.json());

// Middleware to allow iframe embedding
app.use((req, res, next) => {
  res.setHeader("Content-Security-Policy", "frame-ancestors *");
  res.setHeader("X-Frame-Options", "ALLOWALL");
  next();
});

// Google Sheets Auth Helper
const getSheetsClient = () => {
  const email = process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL;
  let privateKey = process.env.GOOGLE_PRIVATE_KEY;
  
  if (privateKey) {
    // Handle cases where Vercel might add extra quotes or escape characters
    privateKey = privateKey.replace(/^"|"$/g, '').replace(/\\n/g, "\n");
  }
  
  console.log("Diagnostic - Auth Check:", {
    hasEmail: !!email,
    hasKey: !!privateKey,
    keyStart: privateKey?.substring(0, 20),
    emailMatch: email === "minhdlst@gmail.com" ? "Warning: Using user email instead of service account email?" : "Check passed"
  });

  if (!email || !privateKey) {
    throw new Error("Google Service Account credentials missing. Please set GOOGLE_SERVICE_ACCOUNT_EMAIL and GOOGLE_PRIVATE_KEY (ensure newlines are handled).");
  }

  const auth = new google.auth.JWT({
    email,
    key: privateKey,
    scopes: ["https://www.googleapis.com/auth/spreadsheets"],
  });

  return google.sheets({ version: "v4", auth });
};

const SPREADSHEET_ID = process.env.SPREADSHEET_ID || "1OlyDLi9n4aS9pouStGP7Rsdix7M_8knhyKiogmfzXis";

// API Routes
app.get("/api/sheets/data", async (req, res) => {
  try {
    const sheets = getSheetsClient();
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "'data'!A:AK",
    });
    const values = response.data.values || [];
    console.log(`Fetched ${values.length} rows from 'data' sheet at ${new Date().toISOString()}`);
    res.json(values);
  } catch (error: any) {
    console.error("Error fetching 'data' sheet:", error);
    res.status(500).json({ error: error.message });
  }
});

app.get("/api/sheets/cap-nhat", async (req, res) => {
  try {
    const sheets = getSheetsClient();
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "'cap nhat'!A:AK",
    });
    res.json(response.data.values || []);
  } catch (error: any) {
    console.error("Error fetching 'cap nhat' sheet:", error);
    res.status(500).json({ error: error.message });
  }
});

app.get("/api/sheets/thu-vien", async (req, res) => {
  try {
    const sheets = getSheetsClient();
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "'thu vien'!A:AK",
    });
    const values = response.data.values || [];
    console.log(`Fetched ${values.length} rows from 'thu vien' sheet at ${new Date().toISOString()}`);
    res.json(values);
  } catch (error: any) {
    console.error("Error fetching 'thu vien' sheet:", error);
    res.status(500).json({ error: error.message });
  }
});

app.get("/api/sheets/tong-hop", async (req, res) => {
  try {
    const sheets = getSheetsClient();
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "'Tong hop'!A:AZ", // Range adjusted to cover potential many columns
    });
    const values = response.data.values || [];
    console.log(`Fetched ${values.length} rows from 'Tong hop' sheet at ${new Date().toISOString()}`);
    res.json(values);
  } catch (error: any) {
    console.error("Error fetching 'Tong hop' sheet:", error);
    res.status(500).json({ error: error.message });
  }
});

app.post("/api/sheets/cap-nhat", async (req, res) => {
  try {
    const sheets = getSheetsClient();
    const { dienLuc, tenTram, ngayCapNhat, ngayThucHien, phanLoai, giaiPhap, vuongMac, deXuat } = req.body;

    await sheets.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_ID,
      range: "'cap nhat'!A:I",
      valueInputOption: "USER_ENTERED",
      requestBody: {
        values: [[dienLuc, tenTram, ngayCapNhat, ngayThucHien, phanLoai, giaiPhap, vuongMac, deXuat]],
      },
    });

    res.json({ success: true });
  } catch (error: any) {
    console.error("Error appending to 'cap nhat' sheet:", error);
    res.status(500).json({ error: error.message });
  }
});

app.put("/api/sheets/cap-nhat", async (req, res) => {
  try {
    const sheets = getSheetsClient();
    const { rowIndex, dienLuc, tenTram, ngayCapNhat, ngayThucHien, phanLoai, giaiPhap, vuongMac, deXuat } = req.body;

    if (rowIndex === undefined || rowIndex < 1) {
      return res.status(400).json({ error: "Invalid row index" });
    }

    // rowIndex is 0-based index in the capNhatSheet array (where 0 is header)
    // Row 1 is header in Excel, so rowIndex + 1 is the 1-based row number
    const rowNumber = rowIndex + 1;
    const range = `'cap nhat'!A${rowNumber}:I${rowNumber}`;

    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: range,
      valueInputOption: "USER_ENTERED",
      requestBody: {
        values: [[dienLuc, tenTram, ngayCapNhat, ngayThucHien, phanLoai, giaiPhap, vuongMac, deXuat]],
      },
    });

    res.json({ success: true });
  } catch (error: any) {
    console.error("Error updating 'cap nhat' sheet:", error);
    res.status(500).json({ error: error.message });
  }
});

app.get("/api/sheets/luu", async (req, res) => {
  try {
    const sheets = getSheetsClient();
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "'Luu'!A:AZ",
    });
    const values = response.data.values || [];
    console.log(`Fetched ${values.length} rows from 'Luu' sheet at ${new Date().toISOString()}`);
    res.json(values);
  } catch (error: any) {
    console.error("Error fetching 'Luu' sheet:", error);
    res.status(500).json({ error: error.message });
  }
});

app.get("/api/sheets/lk-cda", async (req, res) => {
  try {
    const sheets = getSheetsClient();
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "'LK CDA'!A:G",
    });
    const values = response.data.values || [];
    console.log(`Fetched ${values.length} rows from 'LK CDA' sheet at ${new Date().toISOString()}`);
    res.json(values);
  } catch (error: any) {
    console.error("Error fetching 'LK CDA' sheet:", error);
    res.status(500).json({ error: error.message });
  }
});

app.get("/api/sheets/tba", async (req, res) => {
  try {
    const sheets = getSheetsClient();
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "'Ten tram'!A:AZ",
    });
    const values = response.data.values || [];
    console.log(`Fetched ${values.length} rows from 'Ten tram' sheet at ${new Date().toISOString()}`);
    res.json(values);
  } catch (error: any) {
    console.error("Error fetching 'Ten tram' sheet:", error);
    res.status(500).json({ error: error.message });
  }
});

app.get("/api/sheets/bien-dong", async (req, res) => {
  try {
    const sheets = getSheetsClient();
    let response;
    let actualSheetName = "bien dong";
    try {
      response = await sheets.spreadsheets.values.get({
        spreadsheetId: SPREADSHEET_ID,
        range: "'bien dong'!A:AZ",
      });
    } catch (err) {
      console.log("Failed fetching with lowercase 'bien dong', trying 'Bien dong'...");
      response = await sheets.spreadsheets.values.get({
        spreadsheetId: SPREADSHEET_ID,
        range: "'Bien dong'!A:AZ",
      });
      actualSheetName = "Bien dong";
    }
    const values = response.data.values || [];
    console.log(`Fetched ${values.length} rows from '${actualSheetName}' sheet at ${new Date().toISOString()}`);
    res.json(values);
  } catch (error: any) {
    console.error("Error fetching 'bien dong' sheet:", error);
    res.status(500).json({ error: error.message });
  }
});

app.put("/api/sheets/bien-dong", async (req, res) => {
  try {
    const sheets = getSheetsClient();
    const { dienLuc, tenTram, thang, nam, giaiTrinh } = req.body;

    if (!dienLuc || !tenTram || !thang || !nam) {
      return res.status(400).json({ error: "Missing required fields (dienLuc, tenTram, thang, nam)" });
    }

    let response;
    let actualSheetName = "bien dong";
    try {
      response = await sheets.spreadsheets.values.get({
        spreadsheetId: SPREADSHEET_ID,
        range: "'bien dong'!A:AZ",
      });
    } catch (err) {
      response = await sheets.spreadsheets.values.get({
        spreadsheetId: SPREADSHEET_ID,
        range: "'Bien dong'!A:AZ",
      });
      actualSheetName = "Bien dong";
    }

    const values = response.data.values || [];
    if (values.length === 0) {
      return res.status(404).json({ error: "Sheet is empty" });
    }

    const header = values[0];
    const idxDL = header.findIndex(h => {
      const nh = h ? h.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/đ/g, "d") : "";
      return nh.includes("don vi") || nh.includes("dien luc") || nh === "dl";
    });
    const idxMaTram = header.findIndex(h => {
      const nh = h ? h.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/đ/g, "d") : "";
      return nh.includes("ma tram");
    });
    const idxTenTram = header.findIndex(h => {
      const nh = h ? h.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/đ/g, "d") : "";
      return nh.includes("ten tram");
    });
    const idxThang = header.findIndex(h => {
      const nh = h ? h.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/đ/g, "d") : "";
      return nh === "thang" || nh === "thang lk";
    });
    const idxNam = header.findIndex(h => {
      const nh = h ? h.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/đ/g, "d") : "";
      return nh === "nam";
    });

    const colDL = idxDL !== -1 ? idxDL : 0;
    const colMaTram = idxMaTram !== -1 ? idxMaTram : 1;
    const colTenTram = idxTenTram !== -1 ? idxTenTram : 2;
    const colThang = idxThang !== -1 ? idxThang : 8;
    const colNam = idxNam !== -1 ? idxNam : 6;

    const normalizeText = (s: string) => {
      return String(s || "")
        .toLowerCase()
        .replace(/đ/g, "d")
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "")
        .replace(/\s+/g, " ")
        .trim();
    };

    const cleanCompany = (s: string) => {
      return normalizeText(s)
        .replace(/^dien luc\s+/, "")
        .replace(/^pc\s+/, "")
        .replace(/^p\s+/, "")
        .replace(/^cong ty\s+/, "")
        .trim();
    };

    const extractDigits = (s: string) => {
      return String(s || "").replace(/\D/g, "").trim();
    };

    const searchDL = cleanCompany(dienLuc);
    const searchTram = normalizeText(tenTram);
    const searchThangNum = extractDigits(thang);
    const searchNamNum = extractDigits(nam);

    let targetRowIndex = -1;
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      if (!row || row.length === 0) continue;

      const rDL = cleanCompany(row[colDL] || "");
      const rMa = normalizeText(row[colMaTram] || "");
      const rTen = normalizeText(row[colTenTram] || "");
      const rThang = extractDigits(row[colThang] || "");
      const rThangLK = idxThang !== -1 ? extractDigits(row[5] || "") : "";
      const rNam = extractDigits(row[colNam] || "");

      const isDLMatch = rDL === searchDL || rDL.includes(searchDL) || searchDL.includes(rDL);
      const isTramMatch = rMa === searchTram || rTen === searchTram || rMa.includes(searchTram) || rTen.includes(searchTram) || searchTram.includes(rMa) || searchTram.includes(rTen);
      const isThangMatch = rThang === searchThangNum || rThangLK === searchThangNum;
      const isNamMatch = rNam === searchNamNum;

      if (isDLMatch && isTramMatch && isThangMatch && isNamMatch) {
        targetRowIndex = i;
        break;
      }
    }

    if (targetRowIndex === -1) {
      return res.status(404).json({ error: `Không tìm thấy dòng khớp với đơn vị: ${dienLuc}, trạm: ${tenTram}, tháng: ${thang}, năm: ${nam}` });
    }

    const rowNumber = targetRowIndex + 1;
    const range = `'${actualSheetName}'!J${rowNumber}:K${rowNumber}`;

    const now = new Date();
    const timestamp = now.toLocaleString("vi-VN", {
      timeZone: "Asia/Ho_Chi_Minh",
      hour: "2-digit",
      minute: "2-digit",
      second: "2-digit",
      day: "2-digit",
      month: "2-digit",
      year: "numeric"
    });

    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: range,
      valueInputOption: "USER_ENTERED",
      requestBody: {
        values: [[giaiTrinh || "", timestamp]],
      },
    });

    res.json({ success: true, rowIndex: targetRowIndex });
  } catch (error: any) {
    console.error("Error updating 'bien dong' sheet:", error);
    res.status(500).json({ error: error.message });
  }
});

app.get("/api/sheets/tba-real", async (req, res) => {
  try {
    const sheets = getSheetsClient();
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "'TBA'!A:AZ",
    });
    const values = response.data.values || [];
    console.log(`Fetched ${values.length} rows from 'TBA' sheet at ${new Date().toISOString()}`);
    res.json(values);
  } catch (error: any) {
    console.error("Error fetching 'TBA' sheet:", error);
    res.status(500).json({ error: error.message });
  }
});

app.post("/api/sheets/tba-real", async (req, res) => {
  try {
    const sheets = getSheetsClient();
    const { dienLuc, tenTram, sdm, ngayDongDien } = req.body;

    await sheets.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_ID,
      range: "'TBA'!A:D",
      valueInputOption: "USER_ENTERED",
      requestBody: {
        values: [[dienLuc, tenTram, sdm, ngayDongDien]],
      },
    });

    res.json({ success: true });
  } catch (error: any) {
    console.error("Error appending to 'TBA' sheet:", error);
    res.status(500).json({ error: error.message });
  }
});

app.put("/api/sheets/tba-real", async (req, res) => {
  try {
    const sheets = getSheetsClient();
    const { rowIndex, dienLuc, tenTram, sdm, ngayDongDien } = req.body;

    if (rowIndex === undefined || rowIndex < 1) {
      return res.status(400).json({ error: "Invalid row index" });
    }

    const rowNumber = rowIndex + 1;
    const range = `'TBA'!A${rowNumber}:D${rowNumber}`;

    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: range,
      valueInputOption: "USER_ENTERED",
      requestBody: {
        values: [[dienLuc, tenTram, sdm, ngayDongDien]],
      },
    });

    res.json({ success: true });
  } catch (error: any) {
    console.error("Error updating 'TBA' sheet:", error);
    res.status(500).json({ error: error.message });
  }
});

app.post("/api/sheets/tba", async (req, res) => {
  try {
    const sheets = getSheetsClient();
    const { dienLuc, tenTram, sdm, ngayDongDien } = req.body;

    await sheets.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_ID,
      range: "'Ten tram'!A:D",
      valueInputOption: "USER_ENTERED",
      requestBody: {
        values: [[dienLuc, tenTram, sdm, ngayDongDien]],
      },
    });

    res.json({ success: true });
  } catch (error: any) {
    console.error("Error appending to 'Ten tram' sheet:", error);
    res.status(500).json({ error: error.message });
  }
});

app.put("/api/sheets/tba", async (req, res) => {
  try {
    const sheets = getSheetsClient();
    const { rowIndex, dienLuc, tenTram, sdm, ngayDongDien } = req.body;

    if (rowIndex === undefined || rowIndex < 1) {
      return res.status(400).json({ error: "Invalid row index" });
    }

    const rowNumber = rowIndex + 1;
    const range = `'Ten tram'!A${rowNumber}:D${rowNumber}`;

    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: range,
      valueInputOption: "USER_ENTERED",
      requestBody: {
        values: [[dienLuc, tenTram, sdm, ngayDongDien]],
      },
    });

    res.json({ success: true });
  } catch (error: any) {
    console.error("Error updating 'Ten tram' sheet:", error);
    res.status(500).json({ error: error.message });
  }
});

app.get("/api/sheets/mang-tai", async (req, res) => {
  try {
    const sheets = getSheetsClient();
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "'Mang tai'!A:AZ",
    });
    const values = response.data.values || [];
    console.log(`Fetched ${values.length} rows from 'Mang tai' sheet at ${new Date().toISOString()}`);
    res.json(values);
  } catch (error: any) {
    console.error("Error fetching 'Mang tai' sheet:", error);
    res.status(500).json({ error: error.message });
  }
});

app.get("/api/sheets/cham-diem", async (req, res) => {
  try {
    const sheets = getSheetsClient();
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "'Cham diem'!A:AZ",
    });
    const values = response.data.values || [];
    console.log(`Fetched ${values.length} rows from 'Cham diem' sheet at ${new Date().toISOString()}`);
    res.json(values);
  } catch (error: any) {
    console.error("Error fetching 'Cham diem' sheet:", error);
    res.status(500).json({ error: error.message });
  }
});

// Debug endpoint removed for production readiness

app.post("/api/sheets/mang-tai", async (req, res) => {
  try {
    const sheets = getSheetsClient();
    const { a, b, c, d, e, f, g, h, i } = req.body;

    await sheets.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_ID,
      range: "'Mang tai'!A:I",
      valueInputOption: "USER_ENTERED",
      requestBody: {
        values: [[a, b, c, d, e, f, g, h, i || ""]],
      },
    });

    res.json({ success: true });
  } catch (error: any) {
    console.error("Error appending to 'Mang tai' sheet:", error);
    res.status(500).json({ error: error.message });
  }
});

async function startServer() {
  const PORT = 3000;

  // Vite middleware for development
  if (process.env.NODE_ENV !== "production") {
    const { createServer } = await import("vite");
    const vite = await createServer({
      server: { middlewareMode: true },
      appType: "spa",
    });
    app.use(vite.middlewares);
  } else {
    const distPath = path.join(process.cwd(), "dist");
    app.use(express.static(distPath));
    app.get("*", (req, res) => {
      res.sendFile(path.join(distPath, "index.html"));
    });
  }

  app.listen(PORT, "0.0.0.0", () => {
    console.log(`Server running on http://localhost:${PORT}`);
  });
}

// Only start the server if not running on Vercel
if (process.env.VERCEL !== "1") {
  startServer().catch((err) => {
    console.error("Failed to start server:", err);
  });
}

export default app;
