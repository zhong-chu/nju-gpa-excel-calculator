"use strict";

const state = {
  courses: [],
  selected: new Set(),
  fileName: "",
  sheetName: ""
};

const elements = {
  uploadView: document.getElementById("upload-view"),
  calculatorView: document.getElementById("calculator-view"),
  fileInput: document.getElementById("file-input"),
  dropZone: document.getElementById("drop-zone"),
  uploadMessage: document.getElementById("upload-message"),
  chooseAgain: document.getElementById("choose-again"),
  fileName: document.getElementById("file-name"),
  sheetName: document.getElementById("sheet-name"),
  dataNote: document.getElementById("data-note"),
  selectedCount: document.getElementById("selected-count"),
  selectedCredits: document.getElementById("selected-credits"),
  gpaResult: document.getElementById("gpa-result"),
  typeFilters: document.getElementById("type-filters"),
  courseList: document.getElementById("course-list"),
  toggleAll: document.getElementById("toggle-all"),
  selectAll: document.getElementById("select-all"),
  selectNone: document.getElementById("select-none")
};

const PASS_FAIL = /^(通过|合格|不通过|不合格)(?:\s*[（(].*[）)])?$/;
const LEVEL_SCORES = { "优秀": 95, "良好": 85, "中等": 75, "及格": 65, "不及格": 0 };

function clean(value) {
  return String(value ?? "").replace(/\u00a0/g, " ").replace(/\s+/g, " ").trim();
}

function normalizeHeader(value) {
  return clean(value).replace(/[：:（）()【】\[\]\s]/g, "");
}

function parseNumber(value) {
  const match = clean(value).match(/-?\d+(?:\.\d+)?/);
  return match ? Number(match[0]) : null;
}

function parseGrade(displayValue) {
  const text = clean(displayValue);
  if (PASS_FAIL.test(text)) return null;
  const numeric = parseNumber(text);
  if (numeric !== null) return numeric;
  return Object.prototype.hasOwnProperty.call(LEVEL_SCORES, text) ? LEVEL_SCORES[text] : null;
}

function findColumn(headers, exactNames) {
  const normalized = headers.map(normalizeHeader);
  for (const name of exactNames) {
    const index = normalized.indexOf(name);
    if (index >= 0) return index;
  }
  return -1;
}

function detectHeader(rows) {
  const limit = Math.min(rows.length, 40);
  for (let rowIndex = 0; rowIndex < limit; rowIndex += 1) {
    const headers = rows[rowIndex].map(clean);
    const columns = {
      term: findColumn(headers, ["学年学期", "学期"]),
      name: findColumn(headers, ["课程名称", "课程名"]),
      type: findColumn(headers, ["课程性质", "课程类别", "课程属性", "课程类型"]),
      credit: findColumn(headers, ["学分", "课程学分"]),
      grade: findColumn(headers, ["总成绩", "总评成绩", "最终成绩", "成绩", "百分制成绩"])
    };
    if (columns.name >= 0 && columns.credit >= 0 && columns.grade >= 0) {
      return { rowIndex, columns };
    }
  }
  return null;
}

function parseSheet(sheet, sheetName) {
  const rows = XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    defval: "",
    raw: false,
    blankrows: false
  });
  const detected = detectHeader(rows);
  if (!detected) return { sheetName, courses: [] };

  const { columns, rowIndex } = detected;
  const courses = [];
  for (const row of rows.slice(rowIndex + 1)) {
    const name = clean(row[columns.name]);
    const credit = parseNumber(row[columns.credit]);
    const gradeDisplay = clean(row[columns.grade]);
    if (!name || credit === null || !gradeDisplay) continue;
    const grade = parseGrade(gradeDisplay);
    if (grade !== null && (grade < 0 || grade > 100)) continue;
    courses.push({
      term: columns.term >= 0 ? clean(row[columns.term]) : "",
      name,
      type: columns.type >= 0 ? clean(row[columns.type]) : "",
      credit,
      grade,
      gradeDisplay
    });
  }

  const unique = new Map();
  for (const course of courses) {
    const key = [course.term, course.name, course.type, course.credit, course.gradeDisplay].join("|");
    if (!unique.has(key)) unique.set(key, course);
  }
  return { sheetName, courses: Array.from(unique.values()) };
}

function readWorkbook(arrayBuffer) {
  const workbook = XLSX.read(arrayBuffer, { type: "array", cellDates: false });
  const candidates = workbook.SheetNames.map((sheetName) => parseSheet(workbook.Sheets[sheetName], sheetName));
  candidates.sort((a, b) => b.courses.length - a.courses.length);
  return candidates[0] || { sheetName: "", courses: [] };
}

async function loadFile(file) {
  elements.uploadMessage.textContent = "正在读取文件……";
  try {
    if (!window.XLSX) throw new Error("Excel 解析组件没有成功加载，请重新解压完整文件夹。");
    if (file.size > 25 * 1024 * 1024) throw new Error("文件超过 25 MB。请确认选择的是教务系统直接导出的成绩文件。");
    const result = readWorkbook(await file.arrayBuffer());
    if (result.courses.length === 0) {
      throw new Error("没有找到包含“课程名、学分、总成绩”的成绩表。请使用教务系统成绩页面直接导出的文件。");
    }
    state.fileName = file.name;
    state.sheetName = result.sheetName;
    state.courses = result.courses;
    state.selected = new Set(
      result.courses
        .map((course, index) => Number.isFinite(course.grade) && course.credit > 0 ? index : null)
        .filter((index) => index !== null)
    );
    renderCalculator();
  } catch (error) {
    elements.uploadMessage.textContent = error?.message || "读取失败，请换一个成绩导出文件重试。";
  }
}

function formatNumber(value) {
  return String(Number(Number(value).toFixed(4)));
}

function isSelectable(course) {
  return Number.isFinite(course.grade) && course.credit > 0;
}

function updateSummary() {
  const selectedCourses = Array.from(state.selected)
    .map((index) => state.courses[index])
    .filter((course) => course && isSelectable(course));
  const totalCredits = selectedCourses.reduce((sum, course) => sum + course.credit, 0);
  const weightedScore = selectedCourses.reduce((sum, course) => sum + course.grade * course.credit, 0);
  const selectableCount = state.courses.filter(isSelectable).length;

  elements.selectedCount.textContent = String(selectedCourses.length);
  elements.selectedCredits.textContent = formatNumber(totalCredits);
  elements.gpaResult.textContent = totalCredits > 0 ? (weightedScore / totalCredits / 20).toFixed(4) : "—";
  elements.toggleAll.checked = selectedCourses.length > 0 && selectedCourses.length === selectableCount;
  elements.toggleAll.indeterminate = selectedCourses.length > 0 && selectedCourses.length < selectableCount;
  updateFilterButtons();
}

function updateFilterButtons() {
  for (const button of elements.typeFilters.querySelectorAll("button[data-type]")) {
    const type = button.dataset.type;
    const indexes = state.courses
      .map((course, index) => isSelectable(course) && (course.type || "未分类") === type ? index : null)
      .filter((index) => index !== null);
    const selectedCount = indexes.filter((index) => state.selected.has(index)).length;
    button.classList.toggle("active", indexes.length > 0 && selectedCount === indexes.length);
    button.classList.toggle("partial", selectedCount > 0 && selectedCount < indexes.length);
    button.setAttribute("aria-pressed", String(indexes.length > 0 && selectedCount === indexes.length));
  }
}

function renderFilters() {
  elements.typeFilters.replaceChildren();
  const counts = new Map();
  for (const course of state.courses) {
    const type = course.type || "未分类";
    counts.set(type, (counts.get(type) || 0) + 1);
  }
  for (const [type, count] of counts) {
    const button = document.createElement("button");
    button.type = "button";
    button.dataset.type = type;
    button.textContent = `${type} ${count}`;
    button.title = `选中或取消全部“${type}”课程`;
    button.addEventListener("click", () => toggleType(type));
    elements.typeFilters.appendChild(button);
  }
}

function toggleType(type) {
  const indexes = state.courses
    .map((course, index) => isSelectable(course) && (course.type || "未分类") === type ? index : null)
    .filter((index) => index !== null);
  const shouldSelect = indexes.some((index) => !state.selected.has(index));
  for (const index of indexes) {
    if (shouldSelect) state.selected.add(index);
    else state.selected.delete(index);
    const checkbox = elements.courseList.querySelector(`input[data-index="${index}"]`);
    if (checkbox) checkbox.checked = shouldSelect;
  }
  updateSummary();
}

function setAll(selected) {
  state.selected.clear();
  state.courses.forEach((course, index) => {
    if (selected && isSelectable(course)) state.selected.add(index);
  });
  for (const checkbox of elements.courseList.querySelectorAll("input[type='checkbox']")) {
    checkbox.checked = selected && !checkbox.disabled;
  }
  updateSummary();
}

function renderCourses() {
  elements.courseList.replaceChildren();
  state.courses.forEach((course, index) => {
    const row = document.createElement("tr");
    if (!isSelectable(course)) row.className = "disabled";

    const checkCell = document.createElement("td");
    checkCell.className = "check-column";
    const checkbox = document.createElement("input");
    checkbox.type = "checkbox";
    checkbox.dataset.index = String(index);
    checkbox.checked = state.selected.has(index);
    checkbox.disabled = !isSelectable(course);
    checkbox.setAttribute("aria-label", `选择课程 ${course.name}`);
    checkbox.addEventListener("change", () => {
      if (checkbox.checked) state.selected.add(index);
      else state.selected.delete(index);
      updateSummary();
    });
    checkCell.appendChild(checkbox);

    const termCell = document.createElement("td");
    termCell.textContent = course.term || "—";
    const nameCell = document.createElement("td");
    nameCell.textContent = course.name;
    const typeCell = document.createElement("td");
    typeCell.textContent = course.type || "—";
    const creditCell = document.createElement("td");
    creditCell.className = "number-column";
    creditCell.textContent = formatNumber(course.credit);
    const gradeCell = document.createElement("td");
    gradeCell.className = "number-column";
    gradeCell.textContent = course.gradeDisplay;

    row.append(checkCell, termCell, nameCell, typeCell, creditCell, gradeCell);
    elements.courseList.appendChild(row);
  });
}

function renderCalculator() {
  elements.fileName.textContent = state.fileName;
  elements.sheetName.textContent = `工作表：${state.sheetName}`;
  const excluded = state.courses.filter((course) => !isSelectable(course)).length;
  elements.dataNote.textContent = `读取 ${state.courses.length} 门课程${excluded ? `，${excluded} 门非数值成绩不计入` : ""}`;
  elements.uploadView.hidden = true;
  elements.calculatorView.hidden = false;
  elements.chooseAgain.hidden = false;
  renderFilters();
  renderCourses();
  updateSummary();
}

function resetUpload() {
  elements.fileInput.value = "";
  elements.uploadMessage.textContent = "";
  elements.calculatorView.hidden = true;
  elements.uploadView.hidden = false;
  elements.chooseAgain.hidden = true;
}

elements.dropZone.addEventListener("click", () => elements.fileInput.click());
elements.fileInput.addEventListener("change", () => {
  const [file] = elements.fileInput.files;
  if (file) loadFile(file);
});
elements.dropZone.addEventListener("dragover", (event) => {
  event.preventDefault();
  elements.dropZone.classList.add("dragging");
});
elements.dropZone.addEventListener("dragleave", () => elements.dropZone.classList.remove("dragging"));
elements.dropZone.addEventListener("drop", (event) => {
  event.preventDefault();
  elements.dropZone.classList.remove("dragging");
  const [file] = event.dataTransfer.files;
  if (file) loadFile(file);
});
elements.chooseAgain.addEventListener("click", resetUpload);
elements.selectAll.addEventListener("click", () => setAll(true));
elements.selectNone.addEventListener("click", () => setAll(false));
elements.toggleAll.addEventListener("change", () => setAll(elements.toggleAll.checked));
