const sessionIdInput = document.getElementById("sessionId");
const sessionCodeInput = document.getElementById("sessionCode");
const gameSelect = document.getElementById("gameSelect");
const lmsGameBox = document.getElementById("lmsGameBox");
const imageFilesInput = document.getElementById("imageFiles");
const imageStatus = document.getElementById("imageStatus");
const fileNamePreview = document.getElementById("fileNamePreview");
const questionsInput = document.getElementById("questionsInput");
const formatBtn = document.getElementById("formatBtn");
const clearBtn = document.getElementById("clearBtn");
const exampleBtn = document.getElementById("exampleBtn");
const downloadBtn = document.getElementById("downloadBtn");
const backBtn = document.getElementById("backBtn");
const statusBox = document.getElementById("status");
const successMessage = document.getElementById("successMessage");
const excelHead = document.querySelector("#excelTable thead");
const excelBody = document.querySelector("#excelTable tbody");
const templateDescription = document.getElementById("templateDescription");
const formatterPage = document.getElementById("formatterPage");
const previewPage = document.getElementById("previewPage");

let configuration = "live";
let parsedQuestions = [];
let selectedImages = new Map();
let generatedBuffer = null;

const LIVE_HEADERS = [
  "Session ID","Session Code","Question Set","Question No","Question",
  "Answer Caption","Option A","Option B","Option C","Option D",
  "Correct Answer","Image","Answer Explaination","Vimeo URL"
];

const GAME_HEADERS = [
  "Session Configuration ID","Session Code","Question No","Question",
  "Answer Caption","Option A","Option B","Option C","Option D",
  "Correct Answer","Image","Answer Explaination","Vimeo URL",
  "Duration In Second","Difficulty Level Between 1 And 3",
  "Set ID Between 1 And 10","Category ID"
];

function sanitizeCode(value) {
  return value.trim()
    .replace(/^LIVEQZ-/i,"")
    .replace(/^TTTQZ-/i,"")
    .replace(/\.xlsx$/i,"")
    .replace(/[\\/:*?"<>|]/g,"-");
}

function getFileName() {
  const code = sanitizeCode(
    sessionCodeInput.value || "RO-TT-T2-S1-V1"
  );

  if (configuration === "live") {
    return `LIVEQZ-${code}.xlsx`;
  }

  if (configuration === "ttt") {
    return `TTTQZ-${code}.xlsx`;
  }

  return `${gameSelect.value || "GAMENAME"}-${code}.xlsx`;
}

function updateFileName() {
  fileNamePreview.textContent = getFileName();
}

sessionCodeInput.addEventListener("input", updateFileName);
gameSelect.addEventListener("change", updateFileName);

document.querySelectorAll(".config-card").forEach(card => {
  card.addEventListener("click", () => {
    document.querySelectorAll(".config-card").forEach(item => item.classList.remove("active"));
    card.classList.add("active");
    configuration = card.dataset.config;

    if (configuration === "lms") {
      lmsGameBox.classList.remove("hidden");
      templateDescription.textContent = "LMS Game — supplied standard 17-column template.";
    } else if (configuration === "ttt") {
      lmsGameBox.classList.add("hidden");
      templateDescription.textContent = "Tic Tac Toe — supplied standard 17-column template. Difficulty is filled only for a 27-question set.";
    } else {
      lmsGameBox.classList.add("hidden");
      templateDescription.textContent = "Live Session Quiz — supplied standard 14-column template.";
    }
    updateFileName();
  });
});

imageFilesInput.addEventListener("change", () => {
  selectedImages.clear();
  const files = Array.from(imageFilesInput.files || []);
  let pngCount = 0;

  files.forEach(file => {
    if (!file.name.toLowerCase().endsWith(".png")) return;
    const match = file.name.match(/^q\s*(\d+)\.png$/i);
    if (match) {
      selectedImages.set(Number(match[1]), file.name);
      pngCount++;
    }
  });

  imageStatus.textContent = files.length
    ? `${pngCount} question image${pngCount === 1 ? "" : "s"} selected.`
    : "No PNG images selected.";
});

function getImageName(questionNumber) {
  return selectedImages.get(questionNumber) || "";
}

function getDifficulty(questionNumber) {
  if (configuration !== "ttt") return "";
  if (parsedQuestions.length !== 27) return "";
  if (questionNumber >= 1 && questionNumber <= 9) return 1;
  if (questionNumber >= 10 && questionNumber <= 18) return 2;
  if (questionNumber >= 19 && questionNumber <= 27) return 3;
  return "";
}

function validate() {
  if (!sessionIdInput.value.trim()) throw new Error("Please enter the Session ID.");
  if (!sessionCodeInput.value.trim()) throw new Error("Please enter the Session Code.");
  if (configuration === "lms" && !gameSelect.value) throw new Error("Please select an LMS Game.");
  if (!parsedQuestions.length) throw new Error("Please format the questions first.");

  if (configuration === "ttt" && ![9,18,27].includes(parsedQuestions.length)) {
    throw new Error(`Tic Tac Toe requires 9, 18, or 27 questions. You currently have ${parsedQuestions.length}.`);
  }
}

function getRows() {
  const sessionId = sessionIdInput.value.trim();
  const sessionCode = sessionCodeInput.value.trim();

  if (configuration === "live") {
    return parsedQuestions.map(q => [
      sessionId, sessionCode, 1, q.number, q.question, "",
      q.A, q.B, q.C, q.D, q.correct, getImageName(q.number), "", ""
    ]);
  }

  return parsedQuestions.map(q => [
    sessionId, sessionCode, q.number, q.question, "",
    q.A, q.B, q.C, q.D, q.correct, getImageName(q.number),
    "", "", "", getDifficulty(q.number), "", ""
  ]);
}

function copyCellStyle(source, target) {
  target.font = { ...source.font };
  target.fill = { ...source.fill };
  target.border = { ...source.border };
  target.alignment = { ...source.alignment };
  target.protection = { ...source.protection };
  target.numFmt = source.numFmt;
}

function copyRowTemplate(ws, sourceRowNumber, targetRowNumber, columnCount) {
  const sourceRow = ws.getRow(sourceRowNumber);
  const targetRow = ws.getRow(targetRowNumber);

  targetRow.height = sourceRow.height;
  targetRow.hidden = sourceRow.hidden;
  targetRow.outlineLevel = sourceRow.outlineLevel;

  for (let c = 1; c <= columnCount; c++) {
    copyCellStyle(sourceRow.getCell(c), targetRow.getCell(c));
  }
}

async function buildWorkbook() {
  const templatePath =
    configuration === "live"
        ? "./templates/LIVEQZ-RO-TT-T2-S1-V1.xlsx"
        : configuration === "ttt"
        ? "./templates/TTTQZ-RO-TT-T2-S1-V1.xlsx"
        : "./templates/GAMENAME-RO-TT-T2-S1-V1.xlsx";

  const response = await fetch(templatePath);
  if (!response.ok) throw new Error(`Could not load the standard Excel template: ${templatePath}`);

  const buffer = await response.arrayBuffer();
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);

  const ws = configuration === "live"
    ? workbook.getWorksheet("Math_Session_Quiz")
    : workbook.getWorksheet("Session_Chapter_Quiz");

  if (!ws) throw new Error("The expected worksheet was not found in the standard template.");

  const rows = getRows();
  const columnCount = configuration === "live" ? 14 : 17;

  // The supplied template's row 2 is the data-row format.
  // Reuse that exact style for every generated question row.
  for (let r = 2; r < 2 + rows.length; r++) {
    if (r !== 2) copyRowTemplate(ws, 2, r, columnCount);

    const row = ws.getRow(r);
    const values = rows[r - 2];

    for (let c = 1; c <= columnCount; c++) {
      row.getCell(c).value = values[c - 1];
    }
  }

  // Remove any rows beyond the generated questions, while preserving the template header.
  const lastNeededRow = 1 + rows.length;
  while (ws.rowCount > lastNeededRow) {
    ws.spliceRows(lastNeededRow + 1, 1);
  }

  // Keep the original workbook/sheet names, widths, fonts, borders, fills and Help sheet.
  // Only the question data cells are changed.

  return workbook;
}

async function renderPreview() {
  const workbook = await buildWorkbook();
  const ws = configuration === "live"
    ? workbook.getWorksheet("Math_Session_Quiz")
    : workbook.getWorksheet("Session_Chapter_Quiz");

  const headers = [];
  for (let c = 1; c <= ws.columnCount; c++) headers.push(ws.getCell(1,c).value);

  excelHead.innerHTML = "";
  excelBody.innerHTML = "";

  const hr = document.createElement("tr");
  headers.forEach(h => {
    const th = document.createElement("th");
    th.textContent = h ?? "";
    hr.appendChild(th);
  });
  excelHead.appendChild(hr);

  for (let r = 2; r <= ws.rowCount; r++) {
    const tr = document.createElement("tr");
    for (let c = 1; c <= ws.columnCount; c++) {
      const td = document.createElement("td");
      const value = ws.getCell(r,c).value;
      td.textContent = value === null || value === undefined || value === "" ? "" : value;
      if (headers[c-1] === "Correct Answer") td.className = "correct";
      tr.appendChild(td);
    }
    excelBody.appendChild(tr);
  }
}

formatBtn.addEventListener("click", async () => {
  try {
    statusBox.textContent = "";
    statusBox.className = "";

    parsedQuestions = parseQuestions(questionsInput.value);
    validate();

    await renderPreview();

    generatedBuffer = null;
    successMessage.style.color = "#15803d";
    successMessage.textContent =
      `✓ Success! ${parsedQuestions.length} questions formatted successfully.`;

    formatterPage.classList.add("hidden");
    previewPage.classList.remove("hidden");

    document.querySelector(".preview-table-wrap").scrollTop = 0;
    document.querySelector(".preview-table-wrap").scrollLeft = 0;
  } catch (error) {
    statusBox.textContent = error.message;
    statusBox.className = "error";
  }
});

backBtn.addEventListener("click", () => {
  previewPage.classList.add("hidden");
  formatterPage.classList.remove("hidden");
  statusBox.textContent = "";
  statusBox.className = "";
});

clearBtn.addEventListener("click", () => {
  questionsInput.value = "";
  sessionIdInput.value = "";
  sessionCodeInput.value = "";
  gameSelect.value = "";
  configuration = "live";

  document.querySelectorAll(".config-card").forEach(card => card.classList.remove("active"));
  document.querySelector('.config-card[data-config="live"]').classList.add("active");

  lmsGameBox.classList.add("hidden");

  selectedImages.clear();
  imageFilesInput.value = "";
  imageStatus.textContent = "No PNG images selected.";

  parsedQuestions = [];
  generatedBuffer = null;
  successMessage.textContent = "";
  successMessage.style.color = "#15803d";
  statusBox.textContent = "";
  statusBox.className = "";

  excelHead.innerHTML = "";
  excelBody.innerHTML = "";

  previewPage.classList.add("hidden");
  formatterPage.classList.remove("hidden");

  updateFileName();
});

exampleBtn.addEventListener("click", () => {
  questionsInput.value = `Q1. How many pixels does the ELB's NeoPixel matrix have?

A. 12
B. 16
C. 25
D. 5

**Correct Answer:** **C**`;

  statusBox.textContent = "Example loaded.";
  statusBox.className = "success";
});

downloadBtn.addEventListener("click", async () => {
  try {
    validate();

    const workbook = await buildWorkbook();
    const output = await workbook.xlsx.writeBuffer();
    const blob = new Blob([output], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    });

    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = getFileName();
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);

    successMessage.style.color = "#15803d";
    successMessage.textContent = `✓ Downloaded: ${getFileName()}`;
  } catch (error) {
    successMessage.style.color = "#dc2626";
    successMessage.textContent = `✕ ${error.message}`;
  }
});

updateFileName();

function cleanMarkdown(text) {
  return text
    .replace(/\*\*/g, "")
    .replace(/__/g, "")
    .replace(/`/g, "")
    .replace(/[✅✓✔]/g, "")
    .replace(/^\s*[-•]\s*/, "")
    .trim();
}

function parseQuestions(text) {
  let source = text.replace(/\r/g, "").trim();

  if (!source) {
    throw new Error("Please paste your questions first.");
  }

  source = source.replace(/^\s*#{1,6}\s*/gm, "");

  const starts = [];
  const qStart = /^\s*(?:Q(?:uestion)?\s*)?(\d+)\s*[\.\):\-]\s*/gim;
  let m;

  while ((m = qStart.exec(source)) !== null) {
    starts.push({
      index: m.index,
      number: Number(m[1]),
      end: qStart.lastIndex
    });
  }

  if (!starts.length) {
    throw new Error("No questions found. Use Q1., Q2., etc.");
  }

  const questions = [];

  for (let i = 0; i < starts.length; i++) {
    const start = starts[i];
    const end = i + 1 < starts.length
      ? starts[i + 1].index
      : source.length;

    const block = source.slice(start.end, end).trim();

    const optionStarts = [];
    const optionRegex = /^\s*([ABCD])\s*[\.\):\-]\s*/gim;
    let om;

    while ((om = optionRegex.exec(block)) !== null) {
      optionStarts.push({
        letter: om[1].toUpperCase(),
        index: om.index,
        end: optionRegex.lastIndex
      });
    }

    if (optionStarts.length < 2) {
      throw new Error(
        `Question ${start.number}: At least two options A and B are required.`
      );
    }

    const questionText = cleanMarkdown(
      block.slice(0, optionStarts[0].index).trim()
    );

    /*
      IMPORTANT:
      Option D is the last option, so there is no next option marker.
      We must stop the option text before the answer line.

      Example:
      D. 5
      **Correct Answer:** **C**

      Option D must become only "5".
    */
    const answerStartPositions = [];

    const answerLabelRegex =
      /(?:^|\n)\s*(?:\*\*\s*)?(?:correct\s*answer|correct\s*option|answer|ans|correct)(?:\s*\*\*)?\s*(?:is|:|=|-)/im;

    const answerLabelMatch = block.match(answerLabelRegex);

    if (answerLabelMatch && answerLabelMatch.index !== undefined) {
      answerStartPositions.push(answerLabelMatch.index);
    }

    const symbolAnswerRegex = /(?:^|\n)\s*(?:✅|✓|✔)\s*\**\s*[ABCD]\b/i;
    const symbolAnswerMatch = block.match(symbolAnswerRegex);

    if (symbolAnswerMatch && symbolAnswerMatch.index !== undefined) {
      answerStartPositions.push(symbolAnswerMatch.index);
    }

    let optionsBlock = block;

    if (answerStartPositions.length > 0) {
      const firstAnswerPosition = Math.min(...answerStartPositions);
      optionsBlock = block.slice(0, firstAnswerPosition).trim();
    }

    // Re-find the option positions only inside the question/options area.
    const optionStartsForValues = [];
    const optionRegexForValues = /^\s*([ABCD])\s*[\.\):\-]\s*/gim;
    let optionValueMatch;

    while ((optionValueMatch = optionRegexForValues.exec(optionsBlock)) !== null) {
      optionStartsForValues.push({
        letter: optionValueMatch[1].toUpperCase(),
        index: optionValueMatch.index,
        end: optionRegexForValues.lastIndex
      });
    }

    const options = {
      A: "",
      B: "",
      C: "",
      D: ""
    };

    for (let j = 0; j < optionStartsForValues.length; j++) {
      const current = optionStartsForValues[j];
      const next = j + 1 < optionStartsForValues.length
        ? optionStartsForValues[j + 1].index
        : optionsBlock.length;

      const value = cleanMarkdown(
        optionsBlock.slice(current.end, next)
      );

      if (!options[current.letter]) {
        options[current.letter] = value;
      }
    }

    const answer = findAnswer(block);

    if (!answer) {
      throw new Error(
        `Question ${start.number}: Could not identify the correct answer.`
      );
    }

    questions.push({
      number: start.number,
      question: questionText,
      A: options.A,
      B: options.B,
      C: options.C,
      D: options.D,
      correct: answer
    });
  }

  return questions;
}

function findAnswer(block) {
  const patterns = [
    /(?:correct\s*answer|correct\s*option|answer|ans|correct)\s*(?:is|:|=|-)?\s*(?:option\s*)?\**\s*\(?([ABCD])\)?\s*[\.\)]?\s*\**/i,
    /(?:correct\s*answer|correct\s*option|answer|ans)\s*(?:is|:|=|-)?[\s\S]{0,25}?\b(?:option\s*)?([ABCD])\b/i,
    /(?:✅|✓|✔)\s*\**\s*\(?([ABCD])\)?/i
  ];

  for (const pattern of patterns) {
    const match = block.match(pattern);
    if (match) {
      return match[1].toUpperCase();
    }
  }

  const boldOption = block.match(/\*\*\s*([ABCD])\s*[\.\):\-]/i);
  if (boldOption) {
    return boldOption[1].toUpperCase();
  }

  const lines = block.split("\n");

  for (const line of lines) {
    const match = line.match(
      /^\s*\*\*\s*([ABCD])\s*[\.\):\-]\s*.*\*\*\s*$/i
    );

    if (match) {
      return match[1].toUpperCase();
    }
  }

  return null;
}

