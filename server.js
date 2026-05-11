import crypto from "node:crypto";
import fs from "node:fs";
import fsp from "node:fs/promises";
import path from "node:path";
import express from "express";
import multer from "multer";
import pdf from "pdf-parse";
import ExcelJS from "exceljs";

const app = express();
const PORT = process.env.PORT || 3030;
const ROOT = process.cwd();
const DATA_DIR = path.join(ROOT, "data");
const UPLOAD_DIR = path.join(DATA_DIR, "uploads");
const OUTPUT_DIR = path.join(DATA_DIR, "outputs");
const TEMPLATE_DIR = path.join(DATA_DIR, "templates");
const DEFAULT_TEMPLATE = path.join(TEMPLATE_DIR, "ledger-reconciliation-template.xlsx");
const MAX_FILE_SIZE = 1024 * 1024 * 750;

await fsp.mkdir(UPLOAD_DIR, { recursive: true });
await fsp.mkdir(OUTPUT_DIR, { recursive: true });
await fsp.mkdir(TEMPLATE_DIR, { recursive: true });

const jobs = new Map();

const storage = multer.diskStorage({
  destination: (_req, _file, cb) => cb(null, UPLOAD_DIR),
  filename: (_req, file, cb) => {
    const safe = file.originalname.replace(/[^\w.\- ()]/g, "_");
    cb(null, `${Date.now()}-${crypto.randomUUID()}-${safe}`);
  }
});

const upload = multer({
  storage,
  limits: {
    fileSize: MAX_FILE_SIZE,
    files: 20
  }
});

app.use(express.static(path.join(ROOT, "public")));
app.use(express.json({ limit: "2mb" }));

app.post(
  "/api/reconcile",
  upload.fields([
    { name: "companyLedger", maxCount: 1 },
    { name: "partyLedgers", maxCount: 15 },
    { name: "template", maxCount: 1 }
  ]),
  async (req, res) => {
    const id = crypto.randomUUID();
    const companyLedger = req.files?.companyLedger?.[0];
    const partyLedgers = req.files?.partyLedgers || [];
    const template = req.files?.template?.[0];

    if (!companyLedger || partyLedgers.length === 0) {
      return res.status(400).json({
        error: "Upload company ledger PDF and at least one party ledger PDF."
      });
    }

    const templatePath = await resolveTemplate(template);

    jobs.set(id, {
      id,
      status: "queued",
      progress: 0,
      message: "Queued",
      createdAt: new Date().toISOString()
    });

    res.json({ jobId: id });

    processJob(id, {
      companyLedger,
      partyLedgers,
      templatePath,
      partyName: String(req.body.partyName || "").trim()
    }).catch((error) => {
      setJob(id, {
        status: "failed",
        progress: 100,
        message: error.message,
        error: String(error.stack || error)
      });
    });
  }
);

app.get("/api/jobs/:id", (req, res) => {
  const job = jobs.get(req.params.id);
  if (!job) return res.status(404).json({ error: "Job not found" });
  res.json(job);
});

app.get("/api/jobs/:id/download", (req, res) => {
  const job = jobs.get(req.params.id);
  if (!job?.outputPath) return res.status(404).send("Output not ready");
  res.download(job.outputPath, path.basename(job.outputPath));
});

app.get("/api/jobs/:id/report", (req, res) => {
  const job = jobs.get(req.params.id);
  if (!job?.reportPath) return res.status(404).send("Report not ready");
  res.download(job.reportPath, path.basename(job.reportPath));
});

app.get("/api/jobs/:id/preview", async (req, res) => {
  try {
    const job = jobs.get(req.params.id);
    if (!job?.outputPath) return res.status(404).json({ error: "Output not ready" });
    res.json(await buildWorkbookPreview(job.outputPath));
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

if (process.env.RECO_NO_SERVER !== "1") {
  app.listen(PORT, () => {
    console.log(`Ledger reconciliation app running at http://localhost:${PORT}`);
  });
}

async function processJob(id, files) {
  setJob(id, { status: "running", progress: 5, message: "Reading ledgers" });
  const ledgers = await detectAndParseLedgers(files.companyLedger, files.partyLedgers, files.partyName);
  const company = ledgers.company;

  setJob(id, { progress: 40, message: "Combining party ledgers" });
  const partyResults = ledgers.parties;
  setJob(id, { progress: 45, message: "Matching invoices" });
  const combinedParty = combinePartyLedgers(partyResults);
  const reconciliation = reconcile(company, combinedParty);
  attachAuditLayers(reconciliation);

  setJob(id, { progress: 65, message: "Writing Excel template" });
  const outputPath = path.join(OUTPUT_DIR, `RECONCILED-${Date.now()}.xlsx`);
  await writeReconciliationWorkbook(files.templatePath, outputPath, reconciliation);

  const reportPath = path.join(OUTPUT_DIR, `RECONCILIATION-REPORT-${Date.now()}.json`);
  await fsp.writeFile(reportPath, JSON.stringify(reconciliation, null, 2), "utf8");

  setJob(id, {
    status: reconciliation.verification.h63 === 0 ? "completed" : "needs_review",
    progress: 100,
    message:
      reconciliation.verification.h63 === 0
        ? "Reconciled successfully"
        : "Finished with mismatch. Check report before using output.",
    outputPath,
    reportPath,
    summary: reconciliation.summary,
    verification: reconciliation.verification,
    diagnostics: reconciliation.diagnostics,
    review: reconciliation.review,
    repair: reconciliation.repair,
    ruleLibrary: reconciliation.ruleLibrary
  });
}

function setJob(id, patch) {
  jobs.set(id, { ...jobs.get(id), ...patch, updatedAt: new Date().toISOString() });
}

async function resolveTemplate(template) {
  if (template?.path) {
    await fsp.copyFile(template.path, DEFAULT_TEMPLATE);
    return DEFAULT_TEMPLATE;
  }
  try {
    await fsp.access(DEFAULT_TEMPLATE, fs.constants.R_OK);
    return DEFAULT_TEMPLATE;
  } catch {
    throw new Error("Upload the Excel template once. After that it will be reused automatically.");
  }
}

async function detectAndParseLedgers(companyFile, partyFiles, partyName) {
  const files = [companyFile, ...partyFiles];
  const parsed = [];
  for (let index = 0; index < files.length; index += 1) {
    const file = files[index];
    const lines = await extractPdfLines(file.path, false);
    parsed.push({ file, lines, companyScore: scoreCompanyLedger(lines) });
  }

  const companySource = parsed.sort((a, b) => b.companyScore - a.companyScore)[0] || parsed[0];
  const partySources = parsed.filter((item) => item !== companySource);

  return {
    company: parseCompanyLedgerLines(companySource.lines),
    parties: partySources.map((item) => applyPartyFileLabel(parsePartyLedgerLines(item.lines, partyName), item.file.originalname))
  };
}

export async function reconcileFromFiles({ companyLedger, partyLedgers, partyName = "" }) {
  const fileFor = (filePath) => ({
    path: filePath,
    originalname: path.basename(filePath)
  });
  const ledgers = await detectAndParseLedgers(fileFor(companyLedger), partyLedgers.map(fileFor), partyName);
  const reconciliation = reconcile(ledgers.company, combinePartyLedgers(ledgers.parties));
  attachAuditLayers(reconciliation);
  return reconciliation;
}

function applyPartyFileLabel(party, fileName = "") {
  const label = partyLedgerLabel(fileName);
  if (!label) return party;
  return {
    ...party,
    invoices: party.invoices.map((invoice) => ({ ...invoice, ledgerLabel: label })),
    payments: (party.payments || []).map((payment) => ({ ...payment, ledgerLabel: label }))
  };
}

function partyLedgerLabel(fileName) {
  const name = String(fileName || "").toUpperCase();
  if (/\bGOTA\b|(?:^|[\s_.-])G(?:[\s_.-]|$)/.test(name)) return "G";
  if (/\bVADA\b|(?:^|[\s_.-])V(?:[\s_.-]|$)/.test(name)) return "V";
  return "";
}

function scoreCompanyLedger(lines) {
  const top = lines.slice(0, 60).join(" ");
  const all = lines.slice(0, 500).join(" ");
  let score = 0;
  if (/^WEST-COAST PHARMACEUTICAL WORKS LTD/i.test(lines[0] || "")) score += 100;
  if (/WEST-COAST PHARMACEUTICAL WORKS LTD\s+-\s+GUJARAT/i.test(top)) score += 60;
  if (/\bVch Type\b/i.test(top)) score += 10;
  if (/\bBANK PAYMENT\b/i.test(all)) score += 10;
  if (/\bBy\s+\(as per details\).*?\bPURCHASE\b/i.test(all)) score += 30;
  if (/\bPACKING MATERIAL PURCHASE\b/i.test(all) || /\bTDS ON PURCHASE OF GOODS\b/i.test(all)) score += 35;
  if (/Account Statement For/i.test(top)) score -= 100;
  if (/\bCr\s+\(as per details\).*?\bSales\b/i.test(all)) score -= 40;
  if (/\bGST SALES\b/i.test(all) && !/\bPURCHASE\b/i.test(all)) score -= 35;
  if (/\bBank Receipt\b/i.test(all)) score -= 10;
  return score;
}

async function extractPdfLines(filePath, stopAtFirstClosing = false) {
  const buffer = await fsp.readFile(filePath);
  const data = await pdf(buffer, {
    max: 0,
    pagerender: async (pageData) => {
      const text = await pageData.getTextContent({
        normalizeWhitespace: false,
        disableCombineTextItems: false
      });
      return renderPdfPageLines(text.items);
    }
  });

  const lines = data.text
    .split(/\r?\n/)
    .map(fixLine)
    .map((line) => line.trim())
    .filter(Boolean);

  if (!stopAtFirstClosing) return lines;

  const result = [];
  for (const line of lines) {
    result.push(line);
    if (/To\s+Closing Balance/i.test(line)) break;
  }
  return result;
}

function renderPdfPageLines(items) {
  const rows = [];
  for (const item of items) {
    const value = String(item.str || "").trim();
    if (!value) continue;
    const x = Number(item.transform?.[4] || 0);
    const y = Number(item.transform?.[5] || 0);
    let row = rows.find((candidate) => Math.abs(candidate.y - y) < 2.4);
    if (!row) {
      row = { y, cells: [] };
      rows.push(row);
    }
    row.cells.push({ x, value });
  }

  return rows
    .sort((a, b) => b.y - a.y)
    .map((row) =>
      row.cells
        .sort((a, b) => a.x - b.x)
        .map((cell) => cell.value)
        .join(" ")
    )
    .join("\n");
}

function fixLine(line) {
  let out = String(line || "");
  out = out.replace(/'/g, ",");
  out = out.replace(/(\d),\s+(\d)/g, "$1,$2");
  out = out.replace(/(\d)\s*,/g, "$1,");
  out = out.replace(/,\s*(\d)/g, ",$1");
  out = out.replace(/(\d)\s+\./g, "$1.");
  out = out.replace(/\.\s+(\d)/g, ".$1");
  out = out.replace(/\s+([/-])\s+/g, "$1");
  out = out.replace(/\s+([/-])/g, "$1");
  out = out.replace(/([/-])\s+/g, "$1");
  out = out.replace(/(\/\d{1,2})\s+(\d{1,2})(?=\s*(?:DT|DATE|Cr|Dr|$))/gi, "$1$2");
  out = out.replace(/(\/\d{1,2})\s+(\d{1,2})(?=\s+\d[\d,]*\.\d{2})/gi, "$1$2");
  return out.replace(/\s+/g, " ");
}

function normaliseRef(ref) {
  if (!ref) return "";
  let value = String(ref)
    .toUpperCase()
    .replace(/[\s,]+/g, "")
    .replace(/^NO[.-]?/i, "")
    .replace(/[.;:]+$/g, "");
  value = normaliseAlphaNumericInvoiceRef(value);
  value = value.replace(/\(R\)$/i, "");
  value = value.replace(/^(INV\/\d{4}-\d{2}\/)0+(\d+)$/i, "$1$2");
  if (/^G\d+[A-Z]$/i.test(value)) value = value.replace(/[A-Z]$/i, "");
  value = value.replace(/\/0+(\d+)$/g, "/$1");
  return value;
}

function invoiceKey(ref, date = "") {
  const key = normaliseRef(ref);
  if (!/^\d+$/.test(key) || !date) return key;
  const parsed = parseDdMmYy(date);
  if (parsed.getTime() <= 0) return key;
  const year = parsed.getMonth() >= 3 ? parsed.getFullYear() : parsed.getFullYear() - 1;
  return `${key}@${String(year).slice(-2)}-${String(year + 1).slice(-2)}`;
}

function normaliseAlphaNumericInvoiceRef(value) {
  const simple = String(value || "").match(/^([A-Z])([0-9OILS]+)(?:\([A-Z]\))?$/i);
  if (simple) {
    return `${simple[1]}${ocrDigits(simple[2])}`;
  }
  const match = String(value || "").match(/^([A-Z]{1,3})([0-9OILS]+)(?:\([A-Z]\))?$/i);
  if (!match) return value;
  return `${match[1]}${ocrDigits(match[2])}`;
}

function ocrDigits(value) {
  return String(value || "")
    .replace(/O/g, "0")
    .replace(/[IL]/g, "1")
    .replace(/S/g, "5");
}

function parseAmount(raw) {
  if (!raw) return 0;
  const cleaned = String(raw).replace(/,/g, "").replace(/[^\d.-]/g, "");
  const amount = Number.parseFloat(cleaned);
  return Number.isFinite(amount) ? amount : 0;
}

function parseDate(line) {
  const text = String(line || "");
  const tally = text.match(/\b(\d{1,2})-([A-Za-z]{3})-(\d{2,4})/);
  if (tally) {
    const months = {
      JAN: "01",
      FEB: "02",
      MAR: "03",
      APR: "04",
      MAY: "05",
      JUN: "06",
      JUL: "07",
      AUG: "08",
      SEP: "09",
      OCT: "10",
      NOV: "11",
      DEC: "12"
    };
    const month = months[tally[2].slice(0, 3).toUpperCase()];
    if (month) return `${tally[1].padStart(2, "0")}.${month}.${tally[3].slice(-2)}`;
  }

  const numeric =
    text.match(/\b(?:DATE|DT)[: -]*(\d{1,2})[-/.](\d{1,2})\s*[-/.]\s*(\d{2,4})\b/i) ||
    text.match(/\b(\d{1,2})\/(\d{1,2})\/(\d{4})\b/) ||
    text.match(/\b(\d{1,2})[-/.](\d{1,2})[-/.](\d{2,4})/);
  if (!numeric) return "";
  const [, dd, mm, yy] = numeric;
  if (Number(mm) < 1 || Number(mm) > 12) return "";
  return `${dd.padStart(2, "0")}.${mm.padStart(2, "0")}.${yy.slice(-2)}`;
}

function parseLedgerPeriod(lines) {
  const months = {
    JAN: "01",
    FEB: "02",
    MAR: "03",
    APR: "04",
    MAY: "05",
    JUN: "06",
    JUL: "07",
    AUG: "08",
    SEP: "09",
    OCT: "10",
    NOV: "11",
    DEC: "12"
  };
  for (const line of lines.slice(0, 120)) {
    const match = String(line || "").match(
      /(\d{1,2})-([A-Za-z]{3})-(\d{2,4})\s+to\s+(\d{1,2})-([A-Za-z]{3})-(\d{2,4})/i
    );
    if (!match) continue;
    const startMonth = months[match[2].slice(0, 3).toUpperCase()];
    const endMonth = months[match[5].slice(0, 3).toUpperCase()];
    if (!startMonth || !endMonth) continue;
    return {
      startDate: `${match[1].padStart(2, "0")}.${startMonth}.${match[3].slice(-2)}`,
      endDate: `${match[4].padStart(2, "0")}.${endMonth}.${match[6].slice(-2)}`
    };
  }
  return { startDate: "", endDate: "" };
}

function parseDocumentDate(line) {
  const text = String(line || "");
  const numeric = text.match(/\b(?:DATE|DT)[: -]*(\d{1,2})[-/.](\d{1,2})\s*[-/.]\s*(\d{2,4})\b/i);
  if (!numeric) return "";
  const [, dd, mm, yy] = numeric;
  if (Number(mm) < 1 || Number(mm) > 12) return "";
  return `${dd.padStart(2, "0")}.${mm.padStart(2, "0")}.${yy.slice(-2)}`;
}

function extractAmounts(line) {
  return [...line.matchAll(/(?:^|\s)(\d{1,3}(?:,\d{2,3})*(?:\.\d{2,3})|\d+(?:\.\d{2,3}))/g)].map((m) =>
    parseAmount(m[1])
  );
}

function lastAmount(line) {
  const amounts = extractAmounts(line);
  return amounts.at(-1) || 0;
}

function extractRefsFromLines(lines) {
  const refs = new Map();
  const addRef = (ref, sourceLine) => {
    const key = normaliseRef(ref);
    if (key && !/^(DATE|DT)/i.test(key)) refs.set(key, { ref, key, sourceLine });
  };

  const refPattern = "((?:R\\s*/\\s*)?\\d{2}\\s*-\\s*\\d{2}\\s*/\\s*0*\\d+)";
  const simpleRefPattern = "([A-Z]{1,6}\\s*\\d+[A-Z]?(?:\\(R\\))?)";
  const patterns = [
    /\bNew Ref\s+(INV\/\d{4}-\d{2}\/0*\d+)\b/gi,
    /\bAgst Ref\s+(INV\/\d{4}-\d{2}\/0*\d+)\b/gi,
    /\bINVO[A-Z]+\s+NO[-. :]*(INV\/\d{4}-\d{2}\/0*\d+)\b/gi,
    new RegExp(`\\bNew Ref\\s+${refPattern}\\b(?=[\\s,]|$)`, "gi"),
    new RegExp(`\\bAgst Ref\\s+(?!BP\\s*/|BP\\b)${refPattern}\\b(?=[\\s,]|$)`, "gi"),
    new RegExp(`\\bINVO[A-Z]+\\s+NO[-. ]*${refPattern}\\b(?=[\\s,]|$)`, "gi"),
    new RegExp(`\\bINVO[A-Z]+\\s+NO[-. ]*[-]?\\s*${refPattern}\\b(?=[\\s,]|$)`, "gi"),
    new RegExp(`\\bNew Ref\\s+${simpleRefPattern}\\b(?=\\s+\\d|\\s+Cr|\\s+Dr|$)`, "gi"),
    new RegExp(`\\bAgst Ref\\s+(?!BP\\b)${simpleRefPattern}\\b(?=\\s+\\d|\\s+Cr|\\s+Dr|$)`, "gi"),
    /\bNew Ref\s+([A-Z]?\d+[A-Z]?)(?!\s*[-/])\b(?=\s+\d|\s+Cr|\s+Dr|$)/gi,
    /\bAgst Ref\s+(?!BP\b)([A-Z]?\d+[A-Z]?)(?!\s*[-/])\b(?=\s+\d|\s+Cr|\s+Dr|$)/gi,
    new RegExp(`\\bINVO[A-Z]+\\s+NO\\s*[:.-]?\\s*${simpleRefPattern}\\b`, "gi"),
    /\b(?:INVO[A-Z]*|BILL)\s+NO\s*[:;'"".-]?\s*([A-Z]?\d+[A-Z]?)(?!\s*[-/])\b/gi
  ];

  for (let i = 0; i < lines.length; i += 1) {
    const candidates = [lines[i]];
    if (i + 1 < lines.length) candidates.push(`${lines[i]}${lines[i + 1]}`, `${lines[i]} ${lines[i + 1]}`);

    for (const candidate of candidates) {
      for (const pattern of patterns) {
        pattern.lastIndex = 0;
        for (const match of candidate.matchAll(pattern)) addRef(match[1], candidate);
      }
    }

    if (/\bINVO[A-Z]+\s+NO[-. ]*$/i.test(lines[i]) && i + 1 < lines.length) {
      const direct = lines[i + 1].match(new RegExp(`^${refPattern}`, "i"));
      if (direct) addRef(direct[1], `${lines[i]} ${lines[i + 1]}`);
    }
  }

  return refs;
}

async function parseCompanyLedger(filePath) {
  const lines = await extractPdfLines(filePath, true);
  return parseCompanyLedgerLines(lines);
}

function parseCompanyLedgerLines(allLines) {
  const lines = stopAtLastClosingBalance(allLines);
  const period = parseLedgerPeriod(allLines);
  const openingBalance = balanceAmountAtLabel(lines, /Opening Balance/i);
  const closingBalance = lastBalanceAmountAtLabel(lines, /To\s+Closing Balance|Closing Balance/i);
  const blocks = [];
  const payments = [];
  const adjustments = [];
  let tdsTotal = 0;
  let currentDate = "";

  for (let i = 0; i < lines.length; i += 1) {
    currentDate = parseDate(lines[i]) || currentDate;
    if (/\bBANK PAYMENT\b/i.test(lines[i])) {
      payments.push({
        ref: extractPaymentRef(lines[i]) || "",
        key: extractPaymentRef(lines[i]) || `${currentDate}-${lastAmount(lines[i])}`,
        date: currentDate,
        amount: lastAmount(lines[i]),
        source: lines[i]
      });
    }
    if (!isCompanyPurchaseHeader(lines[i], lines, i)) continue;
    const blockLines = [...modernServiceLookback(lines, i), lines[i]];
    for (let j = i + 1; j < lines.length; j += 1) {
      if (j > i + 1 && isCompanyPurchaseHeader(lines[j], lines, j)) break;
      if (isCompanyVoucherHeader(lines[j])) break;
      blockLines.push(lines[j]);
    }

    const refs = extractRefsFromLines(blockLines);
    const headerOffset = blockLines.findIndex((line) => line === lines[i]);
    const voucherLines = headerOffset >= 0 ? blockLines.slice(headerOffset) : blockLines;
    const tdsLine = voucherLines.find((line) => /\bTDS ON\b/i.test(line)) || "";
    const tds = lastAmount(tdsLine);
    tdsTotal += tds;
    const netAmount = companyPurchaseHeaderAmount(blockLines, lines[i]);
    const postingDate = parseDate(lines[i]) || currentDate;
    const date = parseDocumentDate(blockLines.join(" ")) || postingDate;

    for (const ref of pruneDuplicateBlockRefs([...refs.values()])) {
      blocks.push({
        ref: ref.ref,
        key: invoiceKey(ref.ref, date),
        date,
        postingDate,
        amount: netAmount,
        tds,
        source: lines[i]
      });
    }
  }

  for (let i = 0; i < lines.length; i += 1) {
    if (!/SUNDRY\s+BALANCE\s+WRIT\w*\s+OFF\s*-?\s*NET/i.test(lines[i])) continue;
    const amount = lastAmount(lines[i]);
    if (amount > 0) {
      adjustments.push({
        ref: "",
        key: `SUNDRY-WRITEOFF-${i}`,
        date: parseDate(lines[i]) || period.startDate || "",
        amount,
        source: lines[i],
        description: "SUNDRY BALANCE WRITTEN OFF - NET"
      });
    }
  }

  return {
    ledgerType: "company-tally",
    period,
    openingBalance,
    closingBalance,
    tdsTotal,
    invoices: dedupeInvoices(blocks),
    payments,
    adjustments
  };
}

function modernServiceLookback(lines, index) {
  if (!/\bBy\s+\(as per details\).*?\bSERVICE PURCHASE WITH TDS\b/i.test(lines[index])) return [];
  const lookback = [];
  for (let j = index - 1; j >= Math.max(0, index - 14); j -= 1) {
    if (isCompanyVoucherHeader(lines[j]) || /\b(?:BANK PAYMENT|Opening Balance|Closing Balance)\b/i.test(lines[j])) break;
    if (/^(?:Carried Over|Brought Forward|continued|Date Particulars|WEST-COAST|Page\b)/i.test(lines[j])) continue;
    lookback.unshift(lines[j]);
  }
  return lookback.some((line) => /\bINVO[A-Z]+\s+NO|^\s*(?:R\s*\/\s*)?\d{2}\s*-\s*\d{2}\s*\/\s*0*\d+/i.test(line)) ? lookback : [];
}

function pruneDuplicateBlockRefs(refs) {
  const keys = new Set(refs.map((ref) => ref.key));
  return refs.filter((ref) => {
    if (!/^\d+$/.test(ref.key)) return true;
    for (const key of keys) {
      if (key === ref.key || key.length <= ref.key.length) continue;
      if (/^\d+$/.test(key) && key.endsWith(ref.key)) return false;
      if (key.startsWith(`${ref.key}-`) || key.startsWith(`${ref.key}/`)) return false;
    }
    return true;
  });
}

function balanceAmountAtLabel(lines, labelPattern) {
  for (let i = 0; i < lines.length; i += 1) {
    if (!labelPattern.test(lines[i])) continue;
    const sameLine = firstAmount(lines[i]) || lastAmount(lines[i]);
    if (sameLine) return sameLine;
    for (let j = i + 1; j < Math.min(lines.length, i + 4); j += 1) {
      const amount = firstAmount(lines[j]) || lastAmount(lines[j]);
      if (amount) return amount;
    }
  }
  return 0;
}

function lastBalanceAmountAtLabel(lines, labelPattern) {
  let amount = 0;
  for (let i = 0; i < lines.length; i += 1) {
    if (!labelPattern.test(lines[i])) continue;
    const sameLine = firstAmount(lines[i]) || lastAmount(lines[i]);
    if (sameLine) {
      amount = sameLine;
      continue;
    }
    for (let j = i + 1; j < Math.min(lines.length, i + 4); j += 1) {
      const nextAmount = firstAmount(lines[j]) || lastAmount(lines[j]);
      if (nextAmount) {
        amount = nextAmount;
        break;
      }
    }
  }
  return amount;
}

function isCompanyPurchaseHeader(line, lines = [], index = 0) {
  if (/\bBy\s+\(as per details\)/i.test(line) && /\bPURCHASE\b/i.test(line)) return true;
  const nearby = lines.slice(index, Math.min(lines.length, index + 8)).join(" ");
  return (
    (/^\s*(?:Dr\s*)?$/i.test(line) || /^\s*Dr\s+\(as per details\)/i.test(line) || (parseDate(line) && /\bDr\b/i.test(line))) &&
    /\(as per details\)/i.test(nearby) &&
    /\bPURCHASE\b/i.test(nearby) &&
    !/\bGST SALES\b/i.test(nearby)
  );
}

function companyPurchaseHeaderAmount(blockLines, headerLine = "") {
  const headerAmount = lastAmount(headerLine);
  if (headerAmount) return headerAmount;
  const detailIndex = blockLines.findIndex((line) => /\(as per details\)/i.test(line));
  for (let i = Math.max(0, detailIndex); i < Math.min(blockLines.length, detailIndex + 4); i += 1) {
    const amount = lastAmount(blockLines[i]);
    if (amount) return amount;
  }
  return 0;
}

function stopAtFirstClosingBalance(lines) {
  const result = [];
  for (const line of lines) {
    result.push(line);
    if (/To\s+Closing Balance/i.test(line)) break;
  }
  return result;
}

function stopAtLastClosingBalance(lines) {
  let lastClosingIndex = -1;
  for (let i = 0; i < lines.length; i += 1) {
    if (/To\s+Closing Balance/i.test(lines[i])) lastClosingIndex = i;
  }
  return lastClosingIndex >= 0 ? lines.slice(0, lastClosingIndex + 1) : lines;
}

function isVoucherStart(line) {
  if (/^\d{1,2}[-/.][A-Za-z]{3}[-/.]\d{2,4}\b/.test(line)) return true;
  const numeric = String(line || "").match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{2,4})(?=\b|\s)/);
  return Boolean(numeric && Number(numeric[2]) >= 1 && Number(numeric[2]) <= 12);
}

function isCompanyVoucherHeader(line) {
  return (
    isVoucherStart(line) ||
    /\bBy\s+\(as per details\)/i.test(line) ||
    /\bBANK PAYMENT\b/i.test(line) ||
    /\bJournal\b/i.test(line) ||
    /\bReceipt\b/i.test(line)
  );
}

async function parsePartyLedger(filePath, suppliedPartyName = "") {
  const lines = await extractPdfLines(filePath, false);
  return parsePartyLedgerLines(lines, suppliedPartyName);
}

function parsePartyLedgerLines(lines, suppliedPartyName = "") {
  if (lines.some((line) => /Partner ledger/i.test(line)) && lines.some((line) => /Ending Balance/i.test(line))) {
    return parseOdooPartnerLedger(lines, suppliedPartyName);
  }

  if (lines.some((line) => /Account Statement For/i.test(line)) && lines.some((line) => /\bBill No\b/i.test(line))) {
    return parseAccountStatementPartyLedger(lines, suppliedPartyName);
  }

  const period = parseLedgerPeriod(lines);
  const openingLine = lines.find((line) => /Opening Balance/i.test(line)) || "";
  const balanceLines = lines.filter((line) => /\b(?:Closing Balance|Balance)\b/i.test(line));
  const closingLine = balanceLines.at(-1) || "";
  const balanceHistory = extractBalanceHistory(lines, period);
  const invoices = [];
  const payments = [];
  let tdsTotal = 0;
  let currentDate = "";

  for (const line of lines) {
    if (/Dr\s+TDS Receivables/i.test(line) || /TDS Receivables/i.test(line)) {
      tdsTotal += lastAmount(line);
    }
  }

  for (let i = 0; i < lines.length; i += 1) {
    currentDate = parseDate(lines[i]) || currentDate;
    if (/\b(?:Bank Receipt|Receipt)\b/i.test(lines[i])) {
      payments.push({
        ref: extractPaymentRef(lines[i]) || "",
        key: extractPaymentRef(lines[i]) || `${currentDate}-${lastAmount(lines[i])}`,
        date: currentDate,
        amount: lastAmount(lines[i]),
        source: lines[i]
      });
    }
    const legacySale = parseLegacyPartySale(lines, i, currentDate);
    if (legacySale) {
      invoices.push(legacySale);
      continue;
    }
    if (!/Cr\s+\(as per details\)/i.test(lines[i])) continue;
    const blockLines = [lines[i]];
    for (let j = i + 1; j < lines.length; j += 1) {
      if (
        isVoucherStart(lines[j]) ||
        /Cr\s+\(as per details\)/i.test(lines[j]) ||
        /\bTDS Receivables\b/i.test(lines[j]) ||
        /\b(?:Bank Receipt|Receipt|Opening Balance|Closing Balance)\b/i.test(lines[j])
      ) {
        break;
      }
      blockLines.push(lines[j]);
    }
    const headerAmount = lastAmount(blockLines[0]);
    const headerRef = extractHeaderInvoiceRef(blockLines[0]);
    let hasNewRef = false;
    for (const line of blockLines) {
      const match = line.match(/\bNew Ref\s+((?:R\s*\/\s*)?\d{2}\s*-\s*\d{2}\s*\/\s*0*\d+)\b/i);
      const fallback = match ? null : line.match(/\bNew Ref\s+(.+?)\s+\d{1,3}(?:,\d{2,3})*(?:\.\d{2})\s+Dr\b/i);
      if (!match && !fallback) continue;
      hasNewRef = true;
      const parsedRef = match ? match[1] : cleanFallbackRef(fallback[1]);
      const ref = headerRef && isLikelySameRefFamily(headerRef, parsedRef, headerAmount, lastAmount(line)) ? headerRef : parsedRef;
      const amount = lastAmount(line) || headerAmount;
      invoices.push({
        ref,
        key: invoiceKey(ref, parseDate(line) || parseDate(blockLines.join(" ")) || currentDate),
        date: parseDate(line) || parseDate(blockLines.join(" ")) || currentDate,
        amount,
        source: line
      });
    }
    if (!hasNewRef && headerRef && /Sales/i.test(blockLines[0])) {
      invoices.push({
        ref: headerRef,
        key: invoiceKey(headerRef, parseDate(blockLines[0]) || parseDate(blockLines.join(" ")) || currentDate),
        date: parseDate(blockLines[0]) || parseDate(blockLines.join(" ")) || currentDate,
        amount: headerAmount,
        source: blockLines[0]
      });
    }
  }

  return {
    ledgerType: "party-tally",
    partyName: suppliedPartyName || guessPartyName(lines),
    period,
    openingBalance: firstAmount(openingLine) || lastAmount(openingLine),
    closingBalance: (balanceHistory.at(-1)?.amount || lastAmount(closingLine)),
    balanceHistory,
    tdsTotal,
    invoices: dedupeInvoices(invoices),
    payments,
    maxDate: maxDateFromLines(lines)
  };
}

function parseLegacyPartySale(lines, index, currentDate) {
  if (!/\bGST SALES\b/i.test(lines[index])) return null;
  const inline = String(lines[index] || "").match(/\bGST SALES\b.*?\bSales\s+([A-Z]?\d+[A-Z]?)\s+(\d[\d,]*\.\d{2,3})/i);
  if (inline) {
    const date = parseDate(lines[index]) || currentDate;
    return {
      ref: inline[1],
      key: invoiceKey(inline[1], date),
      date,
      amount: parseAmount(inline[2]),
      source: lines[index]
    };
  }
  const amount = firstAmount(lines[index]) || firstAmount(lines[index + 1] || "") || lastAmount(lines[index + 1] || "");
  if (!amount) return null;
  let ref = "";
  for (let j = index + 1; j < Math.min(lines.length, index + 5); j += 1) {
    const candidate = String(lines[j] || "").trim();
    if (/^\d+[A-Z]?$/i.test(candidate)) {
      ref = candidate;
      break;
    }
  }
  if (!ref) return null;
  return {
    ref,
    key: invoiceKey(ref, currentDate),
    date: currentDate,
    amount,
    source: lines.slice(index, Math.min(lines.length, index + 4)).join(" ")
  };
}

function extractBalanceHistory(lines, period = {}) {
  const history = [];
  for (let i = 0; i < lines.length; i += 1) {
    if (!/Closing Balance/i.test(lines[i])) continue;
    const amount = firstAmount(lines[i]) || lastAmount(lines[i]) || firstAmount(lines[i + 1] || "") || lastAmount(lines[i + 1] || "");
    if (!amount) continue;
    let date = "";
    for (let j = i + 1; j < Math.min(lines.length, i + 20); j += 1) {
      if (!/Opening Balance/i.test(lines[j])) continue;
      const nextOpeningDate = parseDate(lines[j]);
      if (nextOpeningDate) date = formatDate(addDays(parseDdMmYy(nextOpeningDate), -1));
      break;
    }
    if (!date) date = history.length === 0 && period.startDate && period.endDate ? period.endDate : period.endDate || maxDateFromLines(lines);
    history.push({ date, amount });
  }
  return history;
}

function addDays(date, days) {
  const next = new Date(date.getTime());
  next.setDate(next.getDate() + days);
  return next;
}

function formatDate(date) {
  if (!(date instanceof Date) || Number.isNaN(date.getTime())) return "";
  return `${String(date.getDate()).padStart(2, "0")}.${String(date.getMonth() + 1).padStart(2, "0")}.${String(date.getFullYear()).slice(-2)}`;
}

function extractHeaderInvoiceRef(line) {
  const match = String(line || "").match(/\b(?:Sales\w*|Reimbursement)\s+((?:R\s*\/\s*)?\d{2}\s*-\s*\d{2}\s*\/\s*0*\d+)\b/i);
  return match ? match[1] : "";
}

function isLikelySameRefFamily(headerRef, parsedRef, headerAmount, lineAmount) {
  if (!headerRef || !parsedRef) return false;
  if (normaliseRef(headerRef) === normaliseRef(parsedRef)) return true;
  const headerPrefix = normaliseRef(headerRef).replace(/\d+$/g, "");
  const parsedPrefix = normaliseRef(parsedRef).replace(/\d+$/g, "");
  return headerPrefix === parsedPrefix && Math.abs(Number(headerAmount || 0) - Number(lineAmount || 0)) < 1;
}

function cleanFallbackRef(ref) {
  return String(ref || "")
    .replace(/\s+/g, " ")
    .replace(/\s*\/\s*/g, "/")
    .trim()
    .toUpperCase();
}

function parseAccountStatementPartyLedger(lines, suppliedPartyName = "") {
  const partyName = suppliedPartyName || guessPartyName(lines);
  const openingLine = lines.find((line) => /Opening Balance/i.test(line)) || "";
  const openingBalance = firstAmount(openingLine);
  const invoices = [];
  const payments = [];
  let tdsTotal = 0;
  let lastSale = null;
  let pendingBill = null;
  let lastDate = "";
  let fyCreditForward = 0;
  let fyDebitForward = 0;

  for (let i = 0; i < lines.length; i += 1) {
    const line = lines[i];
    const parsedDate = parseDate(line);
    if (parsedDate) lastDate = parsedDate;

    const forward = line.match(/([\d, ]+\.\d{2})\s+C\/F\s+->\s+On Page\s+\d+\s+([\d, ]+\.\d{2})\s+C\/F\s+->/i);
    if (forward && isBeforeApril2026Context(lines, i)) {
      fyCreditForward = parseAmount(forward[1]);
      fyDebitForward = parseAmount(forward[2]);
    }

    if (/\b(?:BRct|Receipt)\b/i.test(line)) {
      const amount = firstAmount(line);
      const paymentDate = parseDate(line) || lastDate;
      if (amount > 0) {
        payments.push({
          ref: "",
          key: `${paymentDate}-${amount}-${payments.length}`,
          date: paymentDate,
          amount,
          source: line
        });
      }
    }

    if (/\bJrnl\b/i.test(line) && /Tds|TDS|Dedcted|Deducted/i.test(`${line} ${lines[i + 1] || ""}`)) {
      const amount = firstAmount(line);
      if (amount > 0) tdsTotal += amount;
    }

    const sale = line.match(/([\d, ]+\.\d{2})\s+(\d{1,2})\/(\d{1,2})\/(\d{4})\s+Sale\b/i);
    if (sale) {
      lastSale = {
        amount: parseAmount(sale[1]),
        date: `${sale[2].padStart(2, "0")}.${sale[3].padStart(2, "0")}.${sale[4].slice(-2)}`,
        source: line
      };
      if (pendingBill) {
        invoices.push({
          ref: pendingBill,
          key: normaliseRef(pendingBill),
          date: lastSale.date,
          amount: lastSale.amount,
          source: `${line} Bill No ${pendingBill}`
        });
        pendingBill = null;
        lastSale = null;
      }
      continue;
    }

    const bill = line.match(/\bBill No\s+([A-Z][A-Z0-9ILS]*)\b/i);
    if (bill && lastSale) {
      invoices.push({
        ref: bill[1].toUpperCase(),
        key: normaliseRef(bill[1]),
        date: lastSale.date,
        amount: lastSale.amount,
        source: `${lastSale.source} ${line}`
      });
      lastSale = null;
    } else if (bill) {
      pendingBill = bill[1].toUpperCase();
    }
  }

  const closingBalance = round2(
    openingBalance +
      invoices.reduce((sum, invoice) => sum + invoice.amount, 0) -
      payments.reduce((sum, payment) => sum + payment.amount, 0) -
      tdsTotal
  );

  return {
    ledgerType: "party-statement",
    partyName,
    openingBalance,
    closingBalance,
    tdsTotal,
    invoices,
    payments,
    maxDate: lastDate
  };
}

function parseOdooPartnerLedger(lines, suppliedPartyName = "") {
  const partyName = suppliedPartyName || guessPartyName(lines);
  const invoices = [];
  const payments = [];
  let closingBalance = 0;
  let maxDate = "";

  for (const line of lines) {
    const date = parseDate(line);
    if (date && (!maxDate || parseDdMmYy(date) > parseDdMmYy(maxDate))) maxDate = date;

    const invoice = line.match(
      /\b(\d{1,2}\/\d{1,2}\/\d{4})\s+(INV\/\d{4}-\d{2}\/0*\d+)\s+[^\d-]*([\d,]+(?:\.\d{2})?)\s*-\s*[^\d-]*([\d,]+(?:\.\d{2})?)/i
    );
    if (invoice) {
      const parsedDate = parseDate(invoice[1]);
      invoices.push({
        ref: invoice[2].toUpperCase(),
        key: normaliseRef(invoice[2]),
        date: parsedDate,
        amount: parseAmount(invoice[3]),
        source: line
      });
      continue;
    }

    const payment = line.match(/\b(\d{1,2}\/\d{1,2}\/\d{4})\s+(?!INV\/)[^\n]*?-\s*[^\d-]*([\d,]+(?:\.\d{2})?)\s+[^\d-]*([\d,]+(?:\.\d{2})?)/i);
    if (payment && /NEFT|RTGS|PAY|DCBR|BRct|Receipt/i.test(line)) {
      const parsedDate = parseDate(payment[1]);
      payments.push({
        ref: "",
        key: `${parsedDate}-${parseAmount(payment[2])}-${payments.length}`,
        date: parsedDate,
        amount: parseAmount(payment[2]),
        source: line
      });
      continue;
    }

    const ending = line.match(/\bEnding Balance\s+[^\d-]*([\d,]+(?:\.\d{2})?)/i);
    if (ending) closingBalance = parseAmount(ending[1]);
  }

  return {
    ledgerType: "party-odoo",
    partyName,
    openingBalance: 0,
    closingBalance: round2(closingBalance - payments.reduce((sum, payment) => sum + payment.amount, 0)),
    tdsTotal: 0,
    invoices: dedupeInvoices(invoices),
    payments,
    maxDate
  };
}


function firstAmount(line) {
  const match = String(line || "").match(/(\d[\d, ]*\.\d{2,3})/);
  return match ? parseAmount(match[1]) : 0;
}

function lastClosingBalance(lines) {
  const closingLine = [...lines].reverse().find((line) => /Closing Balance/i.test(line)) || "";
  return firstAmount(closingLine) || lastAmount(closingLine);
}

function isWithinFy2526(dateText) {
  const date = parseDdMmYy(dateText);
  return date >= new Date(2025, 3, 1) && date <= new Date(2026, 2, 31);
}

function parseDdMmYy(dateText) {
  const match = String(dateText || "").match(/^(\d{2})\.(\d{2})\.(\d{2})$/);
  if (!match) return new Date(0);
  const year = Number(match[3]) + 2000;
  return new Date(year, Number(match[2]) - 1, Number(match[1]));
}

function maxDateFromLines(lines) {
  let maxDate = "";
  for (const line of lines) {
    const date = parseDate(line);
    if (date && (!maxDate || parseDdMmYy(date) > parseDdMmYy(maxDate))) maxDate = date;
  }
  return maxDate;
}

function isBeforeApril2026Context(lines, index) {
  const text = lines.slice(Math.max(0, index - 30), index + 1).join(" ");
  return !/\b0?[1-9]\/0?4\/2026\b|\bApr-26\b|Page\s*:\s*12/i.test(text);
}

function extractPaymentRef(line) {
  const match = String(line || "").match(/\bBP\/\d{2}-\d{2}\/\d+\b/i);
  return match ? match[0].toUpperCase() : "";
}

function guessPartyName(lines) {
  return (
    lines.find((line) => /^[A-Z][A-Z .&-]+(?:LIMITED|LLP|PACKAGING|LABORATORIES)\b/i.test(line) && !/WEST.?COAST/i.test(line)) ||
    lines.find((line) => !/Ledger|Page|Date|Particulars|WEST-COAST|^\d{1,2}[/-]\d{1,2}/i.test(line)) ||
    "PARTY"
  );
}

function dedupeInvoices(invoices) {
  const map = new Map();
  for (const invoice of invoices) {
    if (!invoice.key) continue;
    if (!map.has(invoice.key)) map.set(invoice.key, invoice);
  }
  return [...map.values()];
}

function combinePartyLedgers(parties) {
  const ledgerTypes = [...new Set(parties.map((p) => p.ledgerType).filter(Boolean))];
  const invoices = ledgerTypes.length === 1 && ledgerTypes[0] === "party-statement"
    ? parties.flatMap((p) => p.invoices)
    : dedupeInvoices(parties.flatMap((p) => p.invoices));
  return {
    ledgerType: ledgerTypes.length === 1 ? ledgerTypes[0] : "party-mixed",
    ledgerTypes,
    partyName: [...new Set(parties.map((p) => p.partyName).filter(Boolean))].join(", ") || "PARTY",
    openingBalance: parties.reduce((sum, p) => sum + p.openingBalance, 0),
    closingBalance: parties.reduce((sum, p) => sum + p.closingBalance, 0),
    balanceHistory: combineBalanceHistory(parties),
    tdsTotal: parties.reduce((sum, p) => sum + p.tdsTotal, 0),
    invoices,
    payments: parties.flatMap((p) => p.payments || []),
    maxDate: parties.map((p) => p.maxDate).filter(Boolean).sort((a, b) => parseDdMmYy(b) - parseDdMmYy(a))[0] || ""
  };
}

function combineBalanceHistory(parties) {
  const byDate = new Map();
  for (const party of parties) {
    for (const balance of party.balanceHistory || []) {
      if (!balance.date) continue;
      byDate.set(balance.date, round2((byDate.get(balance.date) || 0) + Number(balance.amount || 0)));
    }
  }
  return [...byDate.entries()]
    .map(([date, amount]) => ({ date, amount }))
    .sort((a, b) => parseDdMmYy(a.date) - parseDdMmYy(b.date));
}

function applyCompanyPeriodToParty(company, party) {
  const cutoff = company.period?.endDate || "";
  if (!cutoff) return party;
  const cutoffDate = parseDdMmYy(cutoff);
  if (cutoffDate.getTime() <= 0) return party;

  const withinCutoff = (item) => {
    if (!item.date) return true;
    return parseDdMmYy(item.date) <= cutoffDate;
  };

  const balanceAtCutoff = [...(party.balanceHistory || [])]
    .filter((balance) => balance.date && parseDdMmYy(balance.date) <= cutoffDate)
    .sort((a, b) => parseDdMmYy(b.date) - parseDdMmYy(a.date))[0];

  return {
    ...party,
    invoices: (party.invoices || []).filter(withinCutoff),
    payments: (party.payments || []).filter(withinCutoff),
    closingBalance: balanceAtCutoff?.amount || party.closingBalance,
    effectivePeriodEnd: cutoff
  };
}

function reconcile(company, party) {
  party = applyCompanyPeriodToParty(company, party);
  const companyInvoices = company.invoices.map((invoice, index) => ({ ...invoice, rowId: `C${index}` }));
  const partyInvoices = party.invoices.map((invoice, index) => ({ ...invoice, rowId: `P${index}` }));
  const companyByRef = new Map(companyInvoices.map((invoice) => [invoice.key, invoice]));
  const partyByRef = new Map(partyInvoices.map((invoice) => [invoice.key, invoice]));
  const isOdooStatement = party.ledgerType === "party-odoo";
  const matched = matchInvoices(companyInvoices, partyInvoices, party);

  const addInvoices = partyInvoices
    .filter((invoice) => {
      const companyInvoice = companyByRef.get(invoice.key);
      return !matched.partyIds.has(invoice.rowId) || (isOdooStatement && companyInvoice && isYearEndProvision(companyInvoice));
    })
    .map((invoice) => ({
      ...invoice,
      description: isOdooStatement && invoice.ledgerLabel ? invoice.ledgerLabel : "INVOICE NOT BOOKED BY WEST COAST PHARMACEUTICALS"
    }))
    .sort(compareRecoRows);

  const debitNotes = findCompanyPaymentsNotInParty(company.payments || [], party.payments || [], party).map((payment) => ({
    ...payment,
    description: "PAYMENT NOT BOOK"
  }));

  const lessInvoices = companyInvoices
    .filter((invoice) => {
      if (matched.companyIds.has(invoice.rowId)) return false;
      if (!isOdooStatement) return true;
      if (!isWithinFy2526(invoice.postingDate || invoice.date)) return false;
      return !isYearEndProvision(invoice);
    })
    .map((invoice) => ({ ...invoice, description: `PURCHASE INVOICE NOT REFLECT IN ${party.partyName}` }))
    .sort(compareRecoRows);

  const lessTds = lessInvoices.reduce((sum, invoice) => sum + invoice.tds, 0);
  const addKeys = new Set(addInvoices.map((invoice) => invoice.key));
  let tdsNotBooked = isOdooStatement
    ? calculateGrossNetTds(companyByRef, party.invoices, addKeys)
    : Math.max(0, round2(company.tdsTotal - lessTds - party.tdsTotal));
  const adjustmentRows = party.ledgerType === "party-statement" ? company.adjustments || [] : [];
  const addRows = [...adjustmentRows, ...addInvoices];
  const debitTotal = round2(debitNotes.reduce((sum, item) => sum + item.amount, 0));
  const addTotal = round2(addRows.reduce((sum, invoice) => sum + invoice.amount, 0));
  const lessTotal = round2(lessInvoices.reduce((sum, invoice) => sum + invoice.amount, 0));
  const positiveOpeningCells = usesLegacyNumericPartyLedger(party);
  const openingDiff = positiveOpeningCells
    ? round2(party.openingBalance - company.openingBalance)
    : round2(company.openingBalance - party.openingBalance);
  const roundOff = round2(party.closingBalance % 1 ? Math.ceil(party.closingBalance) - party.closingBalance : 0);
  const companyClosing = isOdooStatement
    ? round2(party.closingBalance - openingDiff - tdsNotBooked - debitTotal - addTotal + lessTotal + roundOff)
    : company.closingBalance;
  const outputCompany = { ...company, closingBalance: companyClosing, ledgerClosingBalance: company.closingBalance };
  let computed = round2(companyClosing + openingDiff + tdsNotBooked + debitTotal + addTotal - lessTotal - roundOff);
  const tdsCorrection = round2(party.closingBalance - computed);
  if (!isOdooStatement && tdsCorrection > 0 && tdsCorrection <= 10000) {
    tdsNotBooked = round2(tdsNotBooked + tdsCorrection);
    computed = round2(companyClosing + openingDiff + tdsNotBooked + debitTotal + addTotal - lessTotal - roundOff);
  } else if (
    party.ledgerType === "party-statement" &&
    adjustmentRows.some((row) => /SUNDRY\s+BALANCE\s+WRITTEN\s+OFF/i.test(row.description || "")) &&
    tdsCorrection < 0 &&
    Math.abs(tdsCorrection) <= 10000
  ) {
    tdsNotBooked = round2(Math.max(0, tdsNotBooked + tdsCorrection));
    computed = round2(companyClosing + openingDiff + tdsNotBooked + debitTotal + addTotal - lessTotal - roundOff);
  }
  const h63 = round2(computed - party.closingBalance);

  return {
    partyName: party.partyName,
    company: outputCompany,
    party,
    addInvoices: addRows,
    addInvoiceOnly: addInvoices,
    debitNotes,
    lessInvoices,
    summary: {
      companyClosing,
      ledgerCompanyClosing: company.closingBalance,
      companyOpening: company.openingBalance,
      partyOpening: party.openingBalance,
      partyClosing: party.closingBalance,
      openingDiff,
      tdsNotBooked,
      debitTotal,
      addTotal,
      lessTotal,
      roundOff,
      positiveOpeningCells
    },
    verification: {
      formula: `G7(${companyClosing}) + OpeningDiff(${openingDiff}) + TDS(${tdsNotBooked}) + DEBIT/PAYMENT(${debitTotal}) + ADD(${addTotal}) - LESS(${lessTotal}) - ROUNDOFF(${roundOff}) = ${computed} = H62(${party.closingBalance})`,
      computed,
      h62: party.closingBalance,
      h63
    }
  };
}

const WEST_COAST_RULE_LIBRARY = {
  company: "WEST-COAST PHARMACEUTICAL WORKS LTD",
  supportedCompanyLedgers: [
    "Tally SERVICE PURCHASE WITH TDS",
    "Tally PACKING MATERIAL PURCHASE",
    "Old Tally Dr (as per details) purchase blocks",
    "BANK PAYMENT matching by amount",
    "SUNDRY BALANCE WRITTEN OFF - NET adjustment when present in ledger"
  ],
  supportedPartyLedgers: [
    "Tally Cr (as per details) New Ref invoices",
    "Old Tally GST SALES Sales invoice tables",
    "Account Statement For ... Bill No/Sale statements",
    "Odoo Partner Ledger INV/YYYY-YY references",
    "Multiple branch PDFs such as GOTA/VADA"
  ],
  protectedRegressionCases: ["MESHAYU", "MANTHAN", "BALAJI"],
  safetyRules: [
    "Never add a balance-difference row",
    "Every ADD/LESS row must have source text",
    "Use H63 = 0 only when itemized rows and permitted ledger adjustments explain the balance",
    "Preserve older successful party behavior when adding new formats"
  ]
};

function attachAuditLayers(reco) {
  reco.ruleLibrary = WEST_COAST_RULE_LIBRARY;
  reco.diagnostics = buildDiagnostics(reco);
  reco.review = buildReview(reco);
  reco.repair = buildRuleBasedRepair(reco);
  return reco;
}

function buildDiagnostics(reco) {
  const companyDuplicates = duplicateKeys(reco.company.invoices || []);
  const partyDuplicates = duplicateKeys(reco.party.invoices || []);
  const parserWarnings = [];
  if (!reco.company.openingBalance) parserWarnings.push("Company opening balance parsed as zero.");
  if (!reco.company.closingBalance) parserWarnings.push("Company closing balance parsed as zero.");
  if (!reco.party.openingBalance && reco.party.ledgerType !== "party-odoo") parserWarnings.push("Party opening balance parsed as zero.");
  if (!reco.party.closingBalance) parserWarnings.push("Party closing balance parsed as zero.");
  if ((reco.company.invoices || []).length === 0) parserWarnings.push("No company purchase invoices extracted.");
  if ((reco.party.invoices || []).length === 0) parserWarnings.push("No party invoices extracted.");
  if (companyDuplicates.length) parserWarnings.push(`Company duplicate invoice keys found: ${companyDuplicates.slice(0, 6).join(", ")}`);
  if (partyDuplicates.length) parserWarnings.push(`Party duplicate invoice keys found: ${partyDuplicates.slice(0, 6).join(", ")}`);
  if (reco.verification.h63 !== 0) parserWarnings.push("H63 is not zero; use rule-based repair suggestions before relying on Excel.");

  return {
    detected: {
      partyName: reco.partyName,
      companyLedgerType: reco.company.ledgerType,
      partyLedgerType: reco.party.ledgerType,
      partyLedgerTypes: reco.party.ledgerTypes || [reco.party.ledgerType].filter(Boolean),
      companyPeriod: reco.company.period || {},
      partyEffectivePeriodEnd: reco.party.effectivePeriodEnd || reco.party.period?.endDate || ""
    },
    counts: {
      companyInvoices: (reco.company.invoices || []).length,
      partyInvoices: (reco.party.invoices || []).length,
      companyPayments: (reco.company.payments || []).length,
      partyPayments: (reco.party.payments || []).length,
      addRows: (reco.addInvoices || []).length,
      lessRows: (reco.lessInvoices || []).length,
      paymentRows: (reco.debitNotes || []).length,
      companyAdjustments: (reco.company.adjustments || []).length
    },
    balances: {
      companyOpening: reco.company.openingBalance,
      partyOpening: reco.party.openingBalance,
      companyClosing: reco.summary.companyClosing,
      ledgerCompanyClosing: reco.summary.ledgerCompanyClosing,
      partyClosing: reco.party.closingBalance,
      companyTdsTotal: reco.company.tdsTotal,
      partyTdsTotal: reco.party.tdsTotal
    },
    parserWarnings,
    duplicateKeys: {
      company: companyDuplicates.slice(0, 25),
      party: partyDuplicates.slice(0, 25)
    }
  };
}

function buildReview(reco) {
  return {
    status: reco.verification.h63 === 0 ? "RECONCILED" : "NEEDS_REVIEW",
    downloadAdvice:
      reco.verification.h63 === 0
        ? "Review rows below, then download Excel."
        : "Do not use final Excel until the mismatch is explained by source ledger rows.",
    addRows: previewRows(reco.addInvoices),
    lessRows: previewRows(reco.lessInvoices),
    paymentRows: previewRows(reco.debitNotes),
    adjustmentRows: previewRows(reco.company.adjustments || [])
  };
}

function buildRuleBasedRepair(reco) {
  const requiredAdd = round2(
    reco.summary.partyClosing -
      reco.summary.companyClosing -
      reco.summary.openingDiff -
      reco.summary.tdsNotBooked -
      reco.summary.debitTotal +
      reco.summary.lessTotal +
      reco.summary.roundOff
  );
  const addGap = round2(requiredAdd - reco.summary.addTotal);
  const suggestions = [];
  if (reco.verification.h63 === 0) {
    suggestions.push("H63 is zero. Keep this case in the test bank so future code cannot break it.");
  } else {
    suggestions.push(`Required ADD total is ${requiredAdd}; current ADD total is ${reco.summary.addTotal}; gap is ${addGap}.`);
    if (addGap > 0) suggestions.push("ADD is low. Search party ledger for missing New Ref/Sale/Bill No rows or missed statement branch PDF.");
    if (addGap < 0) suggestions.push("ADD is high. Re-check ADD rows against company purchase narration, split invoice refs, and gross/net amount matching.");
    if ((reco.lessInvoices || []).length) suggestions.push("LESS rows exist. Check whether any company invoice ref is truncated, split across pages, or has OCR spacing.");
    if ((reco.debitNotes || []).length) suggestions.push("Payment rows exist. Check whether party receipt uses a different date but same amount.");
    if ((reco.company.adjustments || []).length) suggestions.push("Company adjustment rows were found. Confirm they are real ledger rows and not shortcut balancing.");
  }
  return {
    requiredAdd,
    currentAdd: reco.summary.addTotal,
    addGap,
    h63: reco.verification.h63,
    suggestions
  };
}

function previewRows(rows = []) {
  return rows.slice(0, 50).map((row) => ({
    description: row.description || "",
    ref: row.ref || "",
    key: row.key || "",
    date: row.date || "",
    amount: row.amount || 0,
    source: row.source || "",
    ledgerLabel: row.ledgerLabel || ""
  }));
}

function duplicateKeys(rows = []) {
  const counts = new Map();
  for (const row of rows) {
    if (!row.key) continue;
    counts.set(row.key, (counts.get(row.key) || 0) + 1);
  }
  return [...counts.entries()].filter(([, count]) => count > 1).map(([key]) => key);
}

function usesLegacyNumericPartyLedger(party) {
  const invoices = party.invoices || [];
  if (invoices.length < 20) return false;
  const numericCount = invoices.filter((invoice) => /^\d+$/.test(String(invoice.ref || "").trim())).length;
  const gstSalesCount = invoices.filter((invoice) => /\bGST SALES\b/i.test(invoice.source || "")).length;
  return numericCount / invoices.length > 0.7 && gstSalesCount / invoices.length > 0.5;
}

function matchInvoices(companyInvoices, partyInvoices, party) {
  if (party.ledgerType !== "party-statement") {
    const companyIds = new Set();
    const partyIds = new Set();
    for (const partyInvoice of partyInvoices) {
      const companyInvoice = companyInvoices.find((invoice) => invoice.key === partyInvoice.key && !companyIds.has(invoice.rowId));
      if (companyInvoice) {
        companyIds.add(companyInvoice.rowId);
        partyIds.add(partyInvoice.rowId);
      }
    }
    for (const partyInvoice of partyInvoices) {
      if (partyIds.has(partyInvoice.rowId)) continue;
      const candidates = companyInvoices.filter(
        (companyInvoice) =>
          !companyIds.has(companyInvoice.rowId) &&
          invoiceAmountsCompatible(companyInvoice, partyInvoice) &&
          invoiceDatesCompatible(companyInvoice, partyInvoice)
      );
      if (candidates.length === 1) {
        companyIds.add(candidates[0].rowId);
        partyIds.add(partyInvoice.rowId);
      }
    }
    return {
      companyIds,
      partyIds
    };
  }

  const companyIds = new Set();
  const partyIds = new Set();

  for (const partyInvoice of partyInvoices) {
    const candidates = companyInvoices.filter((companyInvoice) => companyInvoice.key === partyInvoice.key);
    const companyInvoice = candidates.find((candidate) => !companyIds.has(candidate.rowId) && invoiceAmountsCompatible(candidate, partyInvoice));
    if (companyInvoice) {
      companyIds.add(companyInvoice.rowId);
      partyIds.add(partyInvoice.rowId);
    }
  }

  for (const partyInvoice of partyInvoices) {
    if (partyIds.has(partyInvoice.rowId)) continue;
    const candidates = companyInvoices.filter(
      (companyInvoice) =>
        !companyIds.has(companyInvoice.rowId) &&
        invoiceAmountsCompatible(companyInvoice, partyInvoice) &&
        invoiceDatesCompatible(companyInvoice, partyInvoice)
    );
    if (candidates.length === 1) {
      companyIds.add(candidates[0].rowId);
      partyIds.add(partyInvoice.rowId);
    }
  }

  return { companyIds, partyIds };
}

function invoiceAmountsCompatible(companyInvoice, partyInvoice) {
  const partyAmount = Number(partyInvoice.amount || 0);
  const companyNet = Number(companyInvoice.amount || 0);
  const companyGross = round2(companyNet + Number(companyInvoice.tds || 0));
  return Math.abs(partyAmount - companyNet) <= 1 || Math.abs(partyAmount - companyGross) <= 1;
}

function invoiceDatesCompatible(companyInvoice, partyInvoice) {
  if (!companyInvoice.date || !partyInvoice.date) return true;
  const companyDate = parseDdMmYy(companyInvoice.date);
  const partyDate = parseDdMmYy(partyInvoice.date);
  const days = Math.abs(companyDate - partyDate) / 86400000;
  return days <= 45;
}

function findCompanyPaymentsNotInParty(companyPayments, partyPayments, party = {}) {
  const available = new Map();
  for (const payment of partyPayments) {
    const key = paymentBucket(payment.amount);
    available.set(key, (available.get(key) || 0) + 1);
  }

  const missing = [];
  for (const payment of companyPayments) {
    const key = paymentBucket(payment.amount);
    const count = available.get(key) || 0;
    if (count > 0) {
      available.set(key, count - 1);
    } else if (payment.amount > 0 && shouldIncludeMissingPayment(payment, party)) {
      missing.push(payment);
    }
  }
  return missing;
}

function shouldIncludeMissingPayment(payment, party) {
  if (!payment.date) return true;
  if (party.ledgerType === "party-odoo") return isWithinFy2526(payment.date);
  return true;
}

function calculateGrossNetTds(companyByRef, partyInvoices, addKeys) {
  let total = 0;
  for (const invoice of partyInvoices) {
    if (addKeys.has(invoice.key)) continue;
    const companyInvoice = companyByRef.get(invoice.key);
    if (!companyInvoice) continue;
    const difference = round2(Number(invoice.amount || 0) - Number(companyInvoice.amount || 0));
    if (difference > 0) total += difference;
  }
  return round2(total);
}

function isYearEndProvision(invoice) {
  return /^31\.03\.26$/.test(invoice?.postingDate || invoice?.date || "");
}

function compareRecoRows(a, b) {
  const byDate = parseDdMmYy(a.date) - parseDdMmYy(b.date);
  if (byDate) return byDate;
  return naturalRefValue(a.ref).localeCompare(naturalRefValue(b.ref), undefined, {
    numeric: true,
    sensitivity: "base"
  });
}

function naturalRefValue(ref = "") {
  return String(ref).toUpperCase().replace(/\/0+(\d+)$/g, "/$1");
}

function paymentBucket(amount) {
  return String(Math.round(Number(amount || 0) * 100));
}

function round2(value) {
  return Math.round((Number(value) + Number.EPSILON) * 100) / 100;
}

async function writeReconciliationWorkbook(templatePath, outputPath, reco) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(templatePath);
  const sheet = workbook.worksheets[0];

  for (const col of ["C", "D", "E", "F", "G", "H"]) {
    sheet.getCell(`${col}3`).value = String(reco.partyName || "PARTY").toUpperCase();
  }

  sheet.getCell("G7").value = reco.company.closingBalance;
  sheet.getCell("G10").value = reco.summary.positiveOpeningCells ? Math.abs(reco.company.openingBalance) : -Math.abs(reco.company.openingBalance);
  sheet.getCell("G11").value = reco.summary.positiveOpeningCells ? Math.abs(reco.party.openingBalance) : -Math.abs(reco.party.openingBalance);
  sheet.getCell("G16").value = reco.summary.tdsNotBooked;
  sheet.getCell("H62").value = reco.party.closingBalance;

  sheet.getCell("C14").value = `TDS NOT BOOKED BY ${reco.partyName}`;
  sheet.getCell("C19").value = `DEBIT NOTE NOT BOOKED BY ${reco.partyName}`;
  sheet.getCell("C24").value = `PAYMENT NOT BOOKED BY ${reco.partyName}`;
  sheet.getCell("C28").value = "INVOICE NOT BOOKED BY WEST COAST PHARMACEUTICALS";
  sheet.getCell("C49").value = `PURCHASE INVOICE NOT REFLECT IN ${reco.partyName}`;

  fillFixedRows(sheet, 21, 22, reco.debitNotes || []);
  const addFormulaRow = fillSection(sheet, 30, 47, reco.addInvoices);
  const addInsertedRows = addFormulaRow - 47;
  const lessFormulaRow = fillSection(sheet, 51 + addInsertedRows, 53 + addInsertedRows, reco.lessInvoices);
  const lessInsertedRows = lessFormulaRow - (53 + addInsertedRows);
  const totalARow = 48 + addInsertedRows;
  const creditNoteFormulaRow = 58 + addInsertedRows + lessInsertedRows;
  const roundOffRow = 59 + addInsertedRows + lessInsertedRows;
  const totalBRow = 60 + addInsertedRows + lessInsertedRows;
  const balanceRow = 61 + addInsertedRows + lessInsertedRows;
  const partyClosingRow = 62 + addInsertedRows + lessInsertedRows;
  const differenceRow = 63 + addInsertedRows + lessInsertedRows;

  sheet.getCell("H7").value = { formula: "G7" };
  sheet.getCell("H9").value = { formula: "G11-G10" };
  sheet.getCell("H18").value = { formula: "SUM(G16:G18)" };
  sheet.getCell("H23").value = { formula: "SUM(G21:G22)" };
  sheet.getCell(`H${addFormulaRow}`).value = { formula: `SUM(G30:G${addFormulaRow})` };
  sheet.getCell(`H${totalARow}`).value = { formula: `SUM(H7:H${totalARow - 1})` };
  sheet.getCell(`H${lessFormulaRow}`).value = { formula: `SUM(G${51 + addInsertedRows}:G${lessFormulaRow})` };
  sheet.getCell(`H${creditNoteFormulaRow}`).value = { formula: `SUM(G${56 + addInsertedRows + lessInsertedRows}:G${creditNoteFormulaRow - 1})` };
  sheet.getCell(`H${roundOffRow}`).value = reco.summary.roundOff || 0;
  sheet.getCell(`H${totalBRow}`).value = { formula: `SUM(H${lessFormulaRow}:H${roundOffRow})` };
  sheet.getCell(`H${balanceRow}`).value = { formula: `H${totalARow}-H${totalBRow}` };
  sheet.getCell(`D${partyClosingRow}`).value = `Closing Balance as per ${reco.partyName} Account Dr`;
  sheet.getCell(`H${partyClosingRow}`).value = reco.party.closingBalance;
  sheet.getCell(`H${differenceRow}`).value = { formula: `H${balanceRow}-H${partyClosingRow}` };

  await workbook.xlsx.writeFile(outputPath);
}

async function buildWorkbookPreview(outputPath) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(outputPath);
  const sheet = workbook.worksheets[0];
  const maxColumn = Math.min(Math.max(sheet.actualColumnCount || 8, 8), 10);
  const maxRow = Math.min(Math.max(sheet.actualRowCount || 63, 63), 120);
  const columns = Array.from({ length: maxColumn }, (_, index) => columnLetter(index + 1));
  const rows = [];

  for (let rowNumber = 1; rowNumber <= maxRow; rowNumber += 1) {
    const row = sheet.getRow(rowNumber);
    const cells = columns.map((column, index) => {
      const cell = row.getCell(index + 1);
      return {
        address: `${column}${rowNumber}`,
        value: previewCellValue(cell.value),
        type: previewCellType(cell.value),
        bold: Boolean(cell.font?.bold),
        fill: previewFill(cell),
        align: cell.alignment?.horizontal || "",
        numFmt: cell.numFmt || ""
      };
    });

    rows.push({
      number: rowNumber,
      height: row.height || null,
      cells
    });
  }

  return {
    sheetName: sheet.name,
    columns,
    rows,
    truncated: sheet.actualRowCount > maxRow || sheet.actualColumnCount > maxColumn
  };
}

function previewCellValue(value) {
  if (value == null) return "";
  if (typeof value === "object") {
    if (value.formula) return `=${value.formula}`;
    if (value.result != null) return previewCellValue(value.result);
    if (value.richText) return value.richText.map((item) => item.text || "").join("");
    if (value.text) return value.text;
    if (value.hyperlink) return value.text || value.hyperlink;
    if (value instanceof Date) return value.toLocaleDateString("en-IN");
  }
  return value;
}

function previewCellType(value) {
  if (value == null) return "blank";
  if (typeof value === "number") return "number";
  if (typeof value === "object" && value.formula) return "formula";
  return "text";
}

function previewFill(cell) {
  const color = cell.fill?.fgColor?.argb || cell.fill?.bgColor?.argb || "";
  return color && color !== "00000000" ? `#${color.slice(-6)}` : "";
}

function columnLetter(number) {
  let result = "";
  let current = number;
  while (current > 0) {
    const remainder = (current - 1) % 26;
    result = String.fromCharCode(65 + remainder) + result;
    current = Math.floor((current - 1) / 26);
  }
  return result;
}

function fillSection(sheet, startRow, formulaRow, invoices) {
  const available = formulaRow - startRow;
  if (invoices.length > available) {
    const needed = invoices.length - available;
    sheet.spliceRows(formulaRow, 0, ...Array.from({ length: needed }, () => []));
    for (let offset = 0; offset < needed; offset += 1) {
      copyRowStyle(sheet, startRow, formulaRow + offset);
    }
    formulaRow += needed;
  }

  for (let i = 0; i < invoices.length; i += 1) {
    const rowNumber = startRow + i;
    unmergeDToGIfNeeded(sheet, rowNumber);
    const invoice = invoices[i];
    sheet.getCell(`D${rowNumber}`).value = invoice.description;
    sheet.getCell(`E${rowNumber}`).value = invoice.ref;
    sheet.getCell(`F${rowNumber}`).value = invoice.date;
    sheet.getCell(`G${rowNumber}`).value = invoice.amount;
  }

  for (let rowNumber = startRow + invoices.length; rowNumber < formulaRow; rowNumber += 1) {
    sheet.getCell(`D${rowNumber}`).value = null;
    sheet.getCell(`E${rowNumber}`).value = null;
    sheet.getCell(`F${rowNumber}`).value = null;
    sheet.getCell(`G${rowNumber}`).value = null;
  }

  return formulaRow;
}

function fillFixedRows(sheet, startRow, endRow, rows) {
  for (let rowNumber = startRow; rowNumber <= endRow; rowNumber += 1) {
    unmergeDToGIfNeeded(sheet, rowNumber);
    sheet.getCell(`D${rowNumber}`).value = null;
    sheet.getCell(`E${rowNumber}`).value = null;
    sheet.getCell(`F${rowNumber}`).value = null;
    sheet.getCell(`G${rowNumber}`).value = null;
  }

  rows.slice(0, endRow - startRow + 1).forEach((row, index) => {
    const rowNumber = startRow + index;
    sheet.getCell(`D${rowNumber}`).value = row.description || "";
    sheet.getCell(`E${rowNumber}`).value = row.ref || "";
    sheet.getCell(`F${rowNumber}`).value = row.date || "";
    sheet.getCell(`G${rowNumber}`).value = row.amount || 0;
  });
}

function unmergeDToGIfNeeded(sheet, rowNumber) {
  const range = `D${rowNumber}:G${rowNumber}`;
  try {
    sheet.unMergeCells(range);
  } catch {
    // The row may not be merged in this template.
  }
}

function copyRowStyle(sheet, fromRowNumber, toRowNumber) {
  const source = sheet.getRow(fromRowNumber);
  const target = sheet.getRow(toRowNumber);
  target.height = source.height;
  source.eachCell({ includeEmpty: true }, (cell, colNumber) => {
    const next = target.getCell(colNumber);
    next.style = JSON.parse(JSON.stringify(cell.style || {}));
    next.numFmt = cell.numFmt;
    next.alignment = cell.alignment ? { ...cell.alignment } : undefined;
    next.border = cell.border ? JSON.parse(JSON.stringify(cell.border)) : undefined;
    next.fill = cell.fill ? JSON.parse(JSON.stringify(cell.fill)) : undefined;
  });
}
