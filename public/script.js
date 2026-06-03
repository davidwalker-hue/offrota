const requestForm = document.querySelector("#requestForm");
const formTitle = document.querySelector("#formTitle");
const signedInText = document.querySelector("#signedInText");
const formMessage = document.querySelector("#formMessage");
const closedNotice = document.querySelector("#closedNotice");
const deadlineText = document.querySelector("#deadlineText");
const adminForm = document.querySelector("#adminForm");
const adminControls = document.querySelector("#adminControls");
const adminMessage = document.querySelector("#adminMessage");
const responseCount = document.querySelector("#responseCount");
const deadlineInput = document.querySelector("#deadlineInput");
const requestMonth = document.querySelector("#requestMonth");
const requestYear = document.querySelector("#requestYear");
const workbookSelect = document.querySelector("#workbookSelect");
const saveDeadline = document.querySelector("#saveDeadline");
const saveTitle = document.querySelector("#saveTitle");
const downloadWorkbook = document.querySelector("#downloadWorkbook");
const resetWorkbooks = document.querySelector("#resetWorkbooks");

let adminPassword = "";

function formatDeadline(value) {
  const date = new Date(value);
  return date.toLocaleString("en-GB", {
    weekday: "short",
    day: "2-digit",
    month: "short",
    year: "numeric",
    hour: "2-digit",
    minute: "2-digit"
  });
}

function setMessage(element, text, type = "") {
  element.textContent = text;
  element.className = `message ${type}`.trim();
}

function renderWorkbookOptions(workbooks = []) {
  workbookSelect.innerHTML = "";
  for (const workbook of workbooks) {
    const option = document.createElement("option");
    option.value = workbook.period;
    option.textContent = workbook.isCurrent ? `${workbook.label} (current)` : workbook.label;
    option.selected = workbook.isCurrent;
    workbookSelect.append(option);
  }
}

async function loadStatus() {
  const meResponse = await fetch("/api/me");
  if (meResponse.ok) {
    const me = await meResponse.json();
    signedInText.textContent = me.ssoEnabled && me.email ? `Signed in as ${me.email}` : "";
  }
  const response = await fetch("/api/status");
  const status = await response.json();
  formTitle.textContent = status.settings.title;
  document.title = status.settings.title;
  deadlineText.textContent = `Open until ${formatDeadline(status.settings.deadline)} (${status.settings.timezoneLabel}).`;
  closedNotice.classList.toggle("hidden", !status.isClosed);
  [...requestForm.elements].forEach(element => {
    if (element.tagName !== "BUTTON") element.disabled = status.isClosed;
  });
  requestForm.querySelector("button").disabled = status.isClosed;
}

requestForm.addEventListener("submit", async event => {
  event.preventDefault();
  setMessage(formMessage, "Submitting...");

  const formData = new FormData(requestForm);
  const payload = Object.fromEntries(formData.entries());

  const response = await fetch("/api/submit", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload)
  });
  const result = await response.json();

  if (!response.ok) {
    setMessage(formMessage, result.errors ? result.errors.join(" ") : result.error, "error");
    await loadStatus();
    return;
  }

  requestForm.reset();
  setMessage(formMessage, "Your request has been submitted.", "success");
});

adminForm.addEventListener("submit", async event => {
  event.preventDefault();
  adminPassword = new FormData(adminForm).get("password");
  setMessage(adminMessage, "");

  const response = await fetch("/api/admin", {
    headers: { "X-Admin-Password": adminPassword }
  });
  const result = await response.json();

  if (!response.ok) {
    setMessage(adminMessage, result.error, "error");
    return;
  }

  adminControls.classList.remove("hidden");
  deadlineInput.value = result.settings.deadline;
  requestMonth.value = String(result.settings.requestMonth);
  requestYear.value = String(result.settings.requestYear);
  renderWorkbookOptions(result.workbooks);
  responseCount.textContent = `${result.responseCount} response${result.responseCount === 1 ? "" : "s"} stored.`;
  setMessage(adminMessage, "Admin controls unlocked.", "success");
});

saveDeadline.addEventListener("click", async () => {
  setMessage(adminMessage, "Saving deadline...");
  const response = await fetch("/api/admin/deadline", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "X-Admin-Password": adminPassword
    },
    body: JSON.stringify({ deadline: deadlineInput.value })
  });
  const result = await response.json();

  if (!response.ok) {
    setMessage(adminMessage, result.error, "error");
    return;
  }

  setMessage(adminMessage, "Deadline updated.", "success");
  await loadStatus();
});

saveTitle.addEventListener("click", async () => {
  setMessage(adminMessage, "Saving title...");
  const response = await fetch("/api/admin/title", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "X-Admin-Password": adminPassword
    },
    body: JSON.stringify({
      requestMonth: Number(requestMonth.value),
      requestYear: Number(requestYear.value)
    })
  });
  const result = await response.json();

  if (!response.ok) {
    setMessage(adminMessage, result.error, "error");
    return;
  }

  requestMonth.value = String(result.settings.requestMonth);
  requestYear.value = String(result.settings.requestYear);
  renderWorkbookOptions(result.workbooks);
  responseCount.textContent = `${result.responseCount} response${result.responseCount === 1 ? "" : "s"} stored.`;
  setMessage(adminMessage, "Title updated.", "success");
  await loadStatus();
});

downloadWorkbook.addEventListener("click", () => {
  const period = encodeURIComponent(workbookSelect.value);
  fetch(`/api/admin/responses.xlsx?period=${period}`, {
    headers: { "X-Admin-Password": adminPassword }
  })
    .then(response => {
      if (!response.ok) throw new Error("Could not download the Excel file.");
      return response.blob();
    })
    .then(blob => {
      const link = document.createElement("a");
      link.href = URL.createObjectURL(blob);
      link.download = "off-rota-responses.xlsx";
      link.click();
      URL.revokeObjectURL(link.href);
    })
    .catch(error => setMessage(adminMessage, error.message, "error"));
});

resetWorkbooks.addEventListener("click", async () => {
  const confirmed = window.confirm("Are you sure you want to do this? This will permanently remove all saved response sheets and start again with a fresh empty workbook for the current month and year.");
  if (!confirmed) return;

  setMessage(adminMessage, "Resetting sheets...");
  const response = await fetch("/api/admin/workbooks/reset", {
    method: "POST",
    headers: { "X-Admin-Password": adminPassword }
  });
  const result = await response.json();

  if (!response.ok) {
    setMessage(adminMessage, result.error, "error");
    return;
  }

  renderWorkbookOptions(result.workbooks);
  responseCount.textContent = "0 responses stored.";
  setMessage(adminMessage, "All sheets and responses have been reset.", "success");
});

loadStatus();
