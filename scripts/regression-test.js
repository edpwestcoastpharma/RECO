import fs from "node:fs";
import path from "node:path";

process.env.RECO_NO_SERVER = "1";

const ROOT = process.cwd();
const MANIFEST = path.join(ROOT, "data", "test-bank", "manifest.json");

function resolveLocal(filePath) {
  return path.isAbsolute(filePath) ? filePath : path.join(ROOT, filePath);
}

function assertEqual(label, actual, expected) {
  if (actual !== expected) {
    throw new Error(`${label}: expected ${expected}, got ${actual}`);
  }
}

function assertClose(label, actual, expected) {
  const roundedActual = Math.round(Number(actual || 0) * 100) / 100;
  const roundedExpected = Math.round(Number(expected || 0) * 100) / 100;
  if (roundedActual !== roundedExpected) {
    throw new Error(`${label}: expected ${roundedExpected}, got ${roundedActual}`);
  }
}

async function main() {
  if (!fs.existsSync(MANIFEST)) {
    throw new Error(`Missing test bank manifest: ${MANIFEST}`);
  }

  const { reconcileFromFiles } = await import("../server.js");
  const manifest = JSON.parse(fs.readFileSync(MANIFEST, "utf8"));
  let passed = 0;

  for (const testCase of manifest.cases || []) {
    const companyLedger = resolveLocal(testCase.companyLedger);
    const partyLedgers = (testCase.partyLedgers || []).map(resolveLocal);
    const missing = [companyLedger, ...partyLedgers].filter((filePath) => !fs.existsSync(filePath));
    if (missing.length) {
      console.log(`SKIP ${testCase.name}: missing files`);
      missing.forEach((filePath) => console.log(`  ${filePath}`));
      continue;
    }

    const reco = await reconcileFromFiles({
      companyLedger,
      partyLedgers,
      partyName: testCase.partyName || ""
    });

    assertClose(`${testCase.name} H63`, reco.verification.h63, testCase.expect.h63);
    if (testCase.expect.addTotal !== undefined) assertClose(`${testCase.name} ADD`, reco.summary.addTotal, testCase.expect.addTotal);
    if (testCase.expect.lessTotal !== undefined) assertClose(`${testCase.name} LESS`, reco.summary.lessTotal, testCase.expect.lessTotal);
    if (testCase.expect.tdsNotBooked !== undefined) assertClose(`${testCase.name} TDS`, reco.summary.tdsNotBooked, testCase.expect.tdsNotBooked);
    if (testCase.expect.addRows !== undefined) assertEqual(`${testCase.name} ADD rows`, reco.addInvoices.length, testCase.expect.addRows);
    if (testCase.expect.lessRows !== undefined) assertEqual(`${testCase.name} LESS rows`, reco.lessInvoices.length, testCase.expect.lessRows);

    console.log(`PASS ${testCase.name}: H63=${reco.verification.h63}`);
    passed += 1;
  }

  if (!passed) throw new Error("No regression cases ran. Check data/test-bank/manifest.json paths.");
  console.log(`Regression test bank passed: ${passed} case(s).`);
}

main().catch((error) => {
  console.error(error.message);
  process.exit(1);
});
