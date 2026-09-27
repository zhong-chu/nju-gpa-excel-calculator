"use strict";

const ALL_TERMS = "__all_terms__";

const state = {
  courses: [],
  selected: new Set(),
  simulatedGrades: new Map(),
  activeTerm: ALL_TERMS,
  simulationEnabled: false,
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
  simulatedGpa: document.getElementById("simulated-gpa"),
  gpaDelta: document.getElementById("gpa-delta"),
  termFilters: document.getElementById("term-filters"),
  termStats: document.getElementById("term-stats"),
  typeFilters: document.getElementById("type-filters"),
  courseList: document.getElementById("course-list"),
  toggleAll: document.getElementById("toggle-all"),
  selectAll: document.getElementById("select-all"),
  selectNone: document.getElementById("select-none"),
  simulationToggle: document.getElementById("simulation-toggle"),
  resetSimulation: document.getElementById("reset-simulation"),
  simulationMessage: document.getElementById("simulation-message")
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
    state.activeTerm = ALL_TERMS;
    state.simulationEnabled = false;
    state.simulatedGrades.clear();
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

function courseTerm(course) {
  return course.term || "未标明学期";
}

function courseInScope(course) {
  return state.activeTerm === ALL_TERMS || courseTerm(course) === state.activeTerm;
}

function getScopedIndexes(selectableOnly = false) {
  return state.courses
    .map((course, index) => courseInScope(course) && (!selectableOnly || isSelectable(course)) ? index : null)
    .filter((index) => index !== null);
}

function getEffectiveGrade(course, index) {
  return state.simulatedGrades.has(index) ? state.simulatedGrades.get(index) : course.grade;
}

function calculateStats(indexes, useSimulatedGrades = false) {
  const courses = indexes
    .map((index) => ({ course: state.courses[index], index }))
    .filter(({ course }) => course && isSelectable(course));
  const totalCredits = courses.reduce((sum, { course }) => sum + course.credit, 0);
  const weightedScore = courses.reduce((sum, { course, index }) => {
    const grade = useSimulatedGrades ? getEffectiveGrade(course, index) : course.grade;
    return sum + grade * course.credit;
  }, 0);
  return {
    courseCount: courses.length,
    totalCredits,
    gpa: totalCredits > 0 ? weightedScore / totalCredits / 20 : null
  };
}

function updateSummary() {
  const scopedIndexes = getScopedIndexes(false);
  const selectableIndexes = getScopedIndexes(true);
  const selectedIndexes = selectableIndexes.filter((index) => state.selected.has(index));
  const original = calculateStats(selectedIndexes, false);
  const simulated = calculateStats(selectedIndexes, true);
  const simulatedInScope = selectedIndexes.some((index) => state.simulatedGrades.has(index));

  elements.selectedCount.textContent = String(original.courseCount);
  elements.selectedCredits.textContent = formatNumber(original.totalCredits);
  elements.gpaResult.textContent = original.gpa === null ? "—" : original.gpa.toFixed(4);
  elements.simulatedGpa.textContent = simulated.gpa === null ? "—" : simulated.gpa.toFixed(4);
  elements.gpaDelta.className = "";
  if (original.gpa === null || simulated.gpa === null) {
    elements.gpaDelta.textContent = "";
  } else if (!simulatedInScope) {
    elements.gpaDelta.textContent = "尚未修改成绩";
  } else {
    const difference = simulated.gpa - original.gpa;
    elements.gpaDelta.textContent = `${difference >= 0 ? "+" : ""}${difference.toFixed(4)}`;
    elements.gpaDelta.className = difference > 0 ? "positive" : difference < 0 ? "negative" : "";
  }

  const scopeName = state.activeTerm === ALL_TERMS ? "全部学期" : state.activeTerm;
  const nonNumericCount = scopedIndexes.filter((index) => !isSelectable(state.courses[index])).length;
  elements.termStats.textContent = `${scopeName}：共 ${scopedIndexes.length} 门，可计分 ${selectableIndexes.length} 门${nonNumericCount ? `，非数值 ${nonNumericCount} 门` : ""}`;

  elements.toggleAll.checked = selectedIndexes.length > 0 && selectedIndexes.length === selectableIndexes.length;
  elements.toggleAll.indeterminate = selectedIndexes.length > 0 && selectedIndexes.length < selectableIndexes.length;
  elements.resetSimulation.hidden = state.simulatedGrades.size === 0;
  updateFilterButtons();
}

function updateFilterButtons() {
  for (const button of elements.typeFilters.querySelectorAll("button[data-type]")) {
    const type = button.dataset.type;
    const indexes = getScopedIndexes(true)
      .filter((index) => (state.courses[index].type || "未分类") === type);
    const selectedCount = indexes.filter((index) => state.selected.has(index)).length;
    button.classList.toggle("active", indexes.length > 0 && selectedCount === indexes.length);
    button.classList.toggle("partial", selectedCount > 0 && selectedCount < indexes.length);
    button.setAttribute("aria-pressed", String(indexes.length > 0 && selectedCount === indexes.length));
  }
}

function renderTermFilters() {
  elements.termFilters.replaceChildren();
  const counts = new Map();
  for (const course of state.courses) {
    const term = courseTerm(course);
    counts.set(term, (counts.get(term) || 0) + 1);
  }

  const options = [[ALL_TERMS, "全部学期", state.courses.length], ...Array.from(counts, ([term, count]) => [term, term, count])];
  for (const [value, label, count] of options) {
    const button = document.createElement("button");
    button.type = "button";
    button.dataset.term = value;
    button.textContent = `${label} ${count}`;
    button.classList.toggle("active", state.activeTerm === value);
    button.setAttribute("aria-pressed", String(state.activeTerm === value));
    button.addEventListener("click", () => {
      state.activeTerm = value;
      renderTermFilters();
      renderFilters();
      renderCourses();
      updateSummary();
    });
    elements.termFilters.appendChild(button);
  }
}

function renderFilters() {
  elements.typeFilters.replaceChildren();
  const counts = new Map();
  for (const index of getScopedIndexes(false)) {
    const course = state.courses[index];
    const type = course.type || "未分类";
    counts.set(type, (counts.get(type) || 0) + 1);
  }
  for (const [type, count] of counts) {
    const button = document.createElement("button");
    button.type = "button";
    button.dataset.type = type;
    button.textContent = `${type} ${count}`;
    button.title = `选中或取消当前学期范围内的“${type}”课程`;
    button.addEventListener("click", () => toggleType(type));
    elements.typeFilters.appendChild(button);
  }
}

function toggleType(type) {
  const indexes = getScopedIndexes(true)
    .filter((index) => (state.courses[index].type || "未分类") === type);
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
  for (const index of getScopedIndexes(true)) {
    if (selected) state.selected.add(index);
    else state.selected.delete(index);
  }
  for (const checkbox of elements.courseList.querySelectorAll("input[type='checkbox']")) {
    checkbox.checked = selected && !checkbox.disabled;
  }
  updateSummary();
}

function updateSimulationControls() {
  elements.simulationToggle.textContent = state.simulationEnabled ? "退出成绩模拟" : "开启成绩模拟";
  elements.simulationToggle.classList.toggle("active", state.simulationEnabled);
  elements.simulationToggle.setAttribute("aria-pressed", String(state.simulationEnabled));
  elements.simulationMessage.textContent = state.simulationEnabled
    ? "修改表格中的数值成绩即可试算；原始 Excel 不会被更改。"
    : "";
}

function renderGradeCell(course, index, row) {
  const gradeCell = document.createElement("td");
  gradeCell.className = "number-column grade-column";

  if (!state.simulationEnabled || !isSelectable(course)) {
    gradeCell.textContent = course.gradeDisplay;
    return gradeCell;
  }

  const wrapper = document.createElement("div");
  wrapper.className = "grade-input-wrap";
  const input = document.createElement("input");
  input.type = "number";
  input.min = "0";
  input.max = "100";
  input.step = "0.1";
  input.className = "grade-input";
  input.value = formatNumber(getEffectiveGrade(course, index));
  input.setAttribute("aria-label", `模拟 ${course.name} 的成绩`);

  const original = document.createElement("small");
  original.className = "grade-original";
  const refreshMarker = () => {
    const changed = state.simulatedGrades.has(index);
    row.classList.toggle("simulated", changed);
    input.classList.toggle("changed", changed);
    original.textContent = changed ? `原 ${course.gradeDisplay}` : "";
  };

  input.addEventListener("input", () => {
    const value = Number(input.value);
    if (input.value.trim() === "") {
      state.simulatedGrades.delete(index);
      elements.simulationMessage.textContent = "该课程暂按原始成绩计算；请输入新的模拟成绩。";
    } else if (!Number.isFinite(value) || value < 0 || value > 100) {
      elements.simulationMessage.textContent = "模拟成绩必须是 0 到 100 之间的数字。";
      return;
    } else if (value === course.grade) {
      state.simulatedGrades.delete(index);
      elements.simulationMessage.textContent = "该课程已恢复为原始成绩。";
    } else {
      state.simulatedGrades.set(index, value);
      elements.simulationMessage.textContent = `已模拟“${course.name}”：${course.gradeDisplay} → ${formatNumber(value)}。`;
    }
    refreshMarker();
    updateSummary();
  });
  input.addEventListener("blur", () => {
    const value = Number(input.value);
    if (input.value.trim() === "" || !Number.isFinite(value) || value < 0 || value > 100) {
      input.value = formatNumber(getEffectiveGrade(course, index));
    }
  });

  wrapper.append(input, original);
  gradeCell.appendChild(wrapper);
  refreshMarker();
  return gradeCell;
}

function renderCourses() {
  elements.courseList.replaceChildren();
  for (const index of getScopedIndexes(false)) {
    const course = state.courses[index];
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
    const gradeCell = renderGradeCell(course, index, row);

    row.append(checkCell, termCell, nameCell, typeCell, creditCell, gradeCell);
    elements.courseList.appendChild(row);
  }
}

function renderCalculator() {
  elements.fileName.textContent = state.fileName;
  elements.sheetName.textContent = `工作表：${state.sheetName}`;
  const excluded = state.courses.filter((course) => !isSelectable(course)).length;
  elements.dataNote.textContent = `读取 ${state.courses.length} 门课程${excluded ? `，${excluded} 门非数值成绩不计入` : ""}`;
  elements.uploadView.hidden = true;
  elements.calculatorView.hidden = false;
  elements.chooseAgain.hidden = false;
  renderTermFilters();
  renderFilters();
  renderCourses();
  updateSimulationControls();
  updateSummary();
}

function resetUpload() {
  elements.fileInput.value = "";
  elements.uploadMessage.textContent = "";
  elements.calculatorView.hidden = true;
  elements.uploadView.hidden = false;
  elements.chooseAgain.hidden = true;
  state.courses = [];
  state.selected.clear();
  state.simulatedGrades.clear();
  state.activeTerm = ALL_TERMS;
  state.simulationEnabled = false;
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
elements.simulationToggle.addEventListener("click", () => {
  state.simulationEnabled = !state.simulationEnabled;
  updateSimulationControls();
  renderCourses();
  updateSummary();
});
elements.resetSimulation.addEventListener("click", () => {
  state.simulatedGrades.clear();
  elements.simulationMessage.textContent = "已恢复全部原始成绩。";
  renderCourses();
  updateSummary();
});
