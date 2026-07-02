import express from "express";
import path from "path";
import { createServer as createViteServer } from "vite";
import cron from "node-cron";
import nodemailer from "nodemailer";
import { setupGoogleDriveAndSheets, appendGuestToSheet } from "./src/lib/google-sync.js";
import ExcelJS from "exceljs";

async function startServer() {
  const app = express();
  const PORT = 3000;

  app.use(express.json({ limit: '10mb' }));
  app.use(express.urlencoded({ extended: true, limit: '10mb' }));

  // In-memory data store for guests
  let guests: any[] = [];

  // Dummy mail setup. User should configure SMTP.
  // We'll log to console for preview purposes if keys are missing
  const transporter = nodemailer.createTransport({
    host: process.env.SMTP_HOST || "smtp.example.com",
    port: parseInt(process.env.SMTP_PORT || "587"),
    secure: false, // true for 465, false for other ports
    auth: {
      user: process.env.SMTP_USER || "user@example.com",
      pass: process.env.SMTP_PASS || "pass123",
    },
  });

  // Scheduled task: Runs every day at 17:00 (5 PM)
  cron.schedule("0 17 * * *", async () => {
    console.log("Running daily guest compilation job at 17:00...");
    const today = new Date().toISOString().split('T')[0];
    const todaysGuests = guests.filter(g => g.tanggalKunjungan === today);

    console.log(`Found ${todaysGuests.length} guests for today.`);
    
    let htmlContent = `<h2>Laporan Tamu Harian (${today})</h2>`;
    if (todaysGuests.length > 0) {
      htmlContent += `<table border="1" cellpadding="5" cellspacing="0">
        <thead>
          <tr>
            <th>No Kunjungan</th>
            <th>Nama Lengkap</th>
            <th>Instansi</th>
            <th>Keperluan</th>
            <th>Jam Datang</th>
          </tr>
        </thead>
        <tbody>
          ${todaysGuests.map(g => `
          <tr>
            <td>${g.noKunjungan}</td>
            <td>${g.namaLengkap}</td>
            <td>${g.instansi}</td>
            <td>${g.keperluan}</td>
            <td>${g.jamDatang}</td>
          </tr>
          `).join("")}
        </tbody>
      </table>`;
    } else {
      htmlContent += `<p>Tidak ada tamu hari ini.</p>`;
    }

    try {
      if (process.env.SMTP_USER && process.env.SMTP_PASS) {
        let info = await transporter.sendMail({
          from: '"Sistem Buku Tamu" <admin@bpmp-lampung.go.id>',
          to: "bpmplpg@gmail.com", // default user email
          subject: `Laporan Tamu Harian BPMP Lampung - ${today}`,
          html: htmlContent,
        });
        console.log("Daily report email sent: %s", info.messageId);
      } else {
        console.log("SMTP not configured. Skipping actual email send. Expected email body:");
        console.log(htmlContent);
      }
    } catch (error) {
      console.error("Error sending daily report:", error);
    }
  });

  // Scheduled task: Automatic cleanup every day at 18:00 (6 PM) (Default 6 Months)
  let autoCleanupMonths = 6;
  cron.schedule("0 18 * * *", async () => {
    console.log(`Running scheduled automatic cleanup for records older than ${autoCleanupMonths} months...`);
    const cutoffDate = new Date();
    cutoffDate.setMonth(cutoffDate.getMonth() - autoCleanupMonths);
    
    const originalCount = guests.length;
    guests = guests.filter((g) => new Date(g.tanggalKunjungan) >= cutoffDate);
    const deletedCount = originalCount - guests.length;
    console.log(`Scheduled cleanup completed. Deleted ${deletedCount} old records.`);
  });

  // API Routes
  app.get("/api/guests/export", async (req, res) => {
    try {
      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet("Data Tamu");

      worksheet.columns = [
        { header: 'No Kunjungan', key: 'noKunjungan', width: 20 },
        { header: 'Tanggal', key: 'tanggalKunjungan', width: 15 },
        { header: 'Jam', key: 'jamDatang', width: 10 },
        { header: 'Nama Lengkap', key: 'namaLengkap', width: 25 },
        { header: 'NIK', key: 'nik', width: 20 },
        { header: 'Instansi', key: 'instansi', width: 25 },
        { header: 'Jabatan', key: 'jabatan', width: 20 },
        { header: 'No HP', key: 'nomorHp', width: 20 },
        { header: 'Email', key: 'email', width: 25 },
        { header: 'Alamat', key: 'alamat', width: 30 },
        { header: 'Keperluan', key: 'keperluan', width: 25 },
        { header: 'Keperluan Lain', key: 'keperluanLainnya', width: 25 },
        { header: 'Tujuan Bertemu', key: 'tujuanBertemu', width: 25 },
        { header: 'Jumlah', key: 'jumlahPengunjung', width: 10 },
        { header: 'Link Foto', key: 'photoDataUrl', width: 40 },
        { header: 'Link Tanda Tangan', key: 'signatureDataUrl', width: 40 }
      ];

      guests.forEach((g) => {
        worksheet.addRow(g);
      });

      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      res.setHeader('Content-Disposition', 'attachment; filename=' + 'Data_Tamu_BPMP.xlsx');

      await workbook.xlsx.write(res);
      res.end();
    } catch (err) {
      res.status(500).json({ error: "Failed to generate Excel file" });
    }
  });

  app.post("/api/google-token", async (req, res) => {
    const { token } = req.body;
    if (token) {
      const success = await setupGoogleDriveAndSheets(token);
      if (success) {
        res.json({ message: "Google Drive and Sheets connected successfully." });
      } else {
        res.status(500).json({ error: "Failed to setup Google Drive/Sheets" });
      }
    } else {
      res.status(400).json({ error: "Token is missing" });
    }
  });

  app.get("/api/guests", (req, res) => {
    res.json(guests);
  });

  app.post("/api/guests", async (req, res) => {
    const now = new Date();
    const formatterDate = new Intl.DateTimeFormat('fr-CA', { timeZone: 'Asia/Jakarta', year: 'numeric', month: '2-digit', day: '2-digit' });
    const formatterTime = new Intl.DateTimeFormat('id-ID', { timeZone: 'Asia/Jakarta', hour: '2-digit', minute: '2-digit', hour12: false });
    
    // fr-CA might separate with hyphens: YYYY-MM-DD
    const tanggalKunjungan = formatterDate.format(now);
    const jamDatang = formatterTime.format(now).replace('.', ':');
    const noKunjunganDate = tanggalKunjungan.replace(/-/g, '');

    const newGuest = {
      id: guests.length + 1,
      noKunjungan: `BPMP-${noKunjunganDate}-${Math.floor(1000 + Math.random() * 9000)}`,
      ...req.body,
      tanggalKunjungan,
      jamDatang
    };
    guests.unshift(newGuest); // add to top
    
    // Sync to Google Sheets if connected (do this in background to avoid blocking response)
    appendGuestToSheet(newGuest).catch(console.error);

    res.json(newGuest);
  });

  app.post("/api/guests/cleanup", (req, res) => {
    const { months } = req.body;
    if (!months || isNaN(months)) {
      return res.status(400).json({ error: "Invalid months provided." });
    }

    const cutoffDate = new Date();
    cutoffDate.setMonth(cutoffDate.getMonth() - parseInt(months));
    
    const originalCount = guests.length;
    guests = guests.filter((g) => new Date(g.tanggalKunjungan) >= cutoffDate);
    const deletedCount = originalCount - guests.length;

    res.json({ message: `Successfully cleaned up ${deletedCount} records older than ${months} months.`, deletedCount });
  });

  app.post("/api/guests/autocleanup", (req, res) => {
    const { months } = req.body;
    if (months && !isNaN(months)) {
        autoCleanupMonths = parseInt(months);
        res.json({ message: `Auto cleanup threshold set to ${autoCleanupMonths} months.` });
    } else {
        res.json({ months: autoCleanupMonths });
    }
  });

  app.get("/api/health", (req, res) => {
    res.json({ status: "ok" });
  });

  // Vite middleware for development
  if (process.env.NODE_ENV !== "production") {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: "spa",
    });
    app.use(vite.middlewares);
  } else {
    const distPath = path.join(process.cwd(), 'dist');
    app.use(express.static(distPath));
    app.get('*', (req, res) => {
      res.sendFile(path.join(distPath, 'index.html'));
    });
  }

  app.listen(PORT, "0.0.0.0", () => {
    console.log(`Server running on http://localhost:${PORT}`);
  });
}

startServer();
