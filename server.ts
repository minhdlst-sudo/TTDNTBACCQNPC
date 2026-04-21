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
