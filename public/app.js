const form = document.querySelector("#reconcileForm");
const submitButton = document.querySelector("#submitButton");
const resultPanel = document.querySelector("#resultPanel");
const serverStatus = document.querySelector("#serverStatus");
const jobTitle = document.querySelector("#jobTitle");
const jobMessage = document.querySelector("#jobMessage");
const progressText = document.querySelector("#progressText");
const progressBar = document.querySelector("#progressBar");
const summary = document.querySelector("#summary");
const verificationText = document.querySelector("#verificationText");
const downloadActions = document.querySelector("#downloadActions");
const downloadLink = document.querySelector("#downloadLink");
const reportLink = document.querySelector("#reportLink");

const fields = {
  companyClosing: document.querySelector("#companyClosing"),
  openingDiff: document.querySelector("#openingDiff"),
  tdsNotBooked: document.querySelector("#tdsNotBooked"),
  addTotal: document.querySelector("#addTotal"),
  lessTotal: document.querySelector("#lessTotal"),
  partyClosing: document.querySelector("#partyClosing"),
  h63: document.querySelector("#h63")
};

form.addEventListener("submit", async (event) => {
  event.preventDefault();
  submitButton.disabled = true;
  serverStatus.textContent = "Uploading";
  resultPanel.hidden = false;
  summary.hidden = true;
  downloadActions.hidden = true;
  setProgress(2, "Uploading files", "Please keep this tab open.");

  try {
    const response = await fetch("/api/reconcile", {
      method: "POST",
      body: new FormData(form)
    });

    const payload = await response.json();
    if (!response.ok) throw new Error(payload.error || "Upload failed");
    serverStatus.textContent = "Running";
    pollJob(payload.jobId);
  } catch (error) {
    submitButton.disabled = false;
    serverStatus.textContent = "Error";
    setProgress(100, "Could not start reconciliation", error.message);
  }
});

async function pollJob(jobId) {
  try {
    const response = await fetch(`/api/jobs/${jobId}`);
    const job = await response.json();
    if (!response.ok) throw new Error(job.error || "Job not found");

    setProgress(job.progress || 0, titleForStatus(job.status), job.message || "");

    if (job.summary) renderSummary(job);

    if (["completed", "needs_review", "failed"].includes(job.status)) {
      submitButton.disabled = false;
      serverStatus.textContent = job.status === "completed" ? "Done" : job.status === "failed" ? "Error" : "Review";
      if (job.status !== "failed") {
        downloadLink.href = `/api/jobs/${jobId}/download`;
        reportLink.href = `/api/jobs/${jobId}/report`;
        downloadActions.hidden = false;
      }
      return;
    }

    window.setTimeout(() => pollJob(jobId), 1200);
  } catch (error) {
    submitButton.disabled = false;
    serverStatus.textContent = "Error";
    setProgress(100, "Processing stopped", error.message);
  }
}

function setProgress(progress, title, message) {
  const safeProgress = Math.max(0, Math.min(100, Number(progress) || 0));
  progressBar.style.width = `${safeProgress}%`;
  progressText.textContent = `${safeProgress}%`;
  jobTitle.textContent = title;
  jobMessage.textContent = message;
}

function renderSummary(job) {
  summary.hidden = false;
  fields.companyClosing.textContent = money(job.summary.companyClosing);
  fields.openingDiff.textContent = money(job.summary.openingDiff);
  fields.tdsNotBooked.textContent = money(job.summary.tdsNotBooked);
  fields.addTotal.textContent = money(job.summary.addTotal);
  fields.lessTotal.textContent = money(job.summary.lessTotal);
  fields.partyClosing.textContent = money(job.summary.partyClosing);
  fields.h63.textContent = money(job.verification?.h63 || 0);
  verificationText.textContent = job.verification?.formula
    ? `${job.verification.formula}. H63 = ${money(job.verification.h63)}`
    : "";
  verificationText.classList.toggle("is-warning", job.status === "needs_review");
  verificationText.classList.toggle("is-danger", job.status === "failed");
}

function titleForStatus(status) {
  if (status === "completed") return "Reconciled";
  if (status === "needs_review") return "Needs review";
  if (status === "failed") return "Failed";
  if (status === "queued") return "Queued";
  return "Processing";
}

function money(value) {
  return new Intl.NumberFormat("en-IN", {
    maximumFractionDigits: 2,
    minimumFractionDigits: 2
  }).format(Number(value) || 0);
}
