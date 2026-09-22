"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const projectDir = __dirname;
const XLSX = require(path.join(projectDir, "xlsx.full.min.js"));

function fakeElement() {
  return {
    addEventListener() {},
    append() {},
    appendChild() {},
    replaceChildren() {},
    querySelector() { return null; },
    querySelectorAll() { return []; },
    setAttribute() {},
    classList: { add() {}, remove() {}, toggle() {} },
    files: [],
    checked: false,
    disabled: false,
    hidden: false,
    indeterminate: false,
    textContent: "",
    value: ""
  };
}

const context = {
  XLSX,
  window: { XLSX },
  document: {
    getElementById() { return fakeElement(); },
    createElement() { return fakeElement(); }
  },
  console,
  Set,
  Map,
  Number,
  String,
  Object,
  Array,
  RegExp
};
vm.createContext(context);
const source = fs.readFileSync(path.join(projectDir, "app.js"), "utf8");
vm.runInContext(`${source}\nthis.__testApi = { readWorkbook, parseGrade, isSelectable };`, context);

const rows = [
  ["南京大学成绩单"],
  ["学年学期", "课程名称", "课程号", "学分", "课程性质", "总成绩"],
  ["2025-2026学年 第1学期", "问题求解（二）", "22020010B", 4, "平台", 67],
  ["2025-2026学年 第2学期", "学术思维与科研技能（二）", "24020300B", 1, "平台", "通过"],
  ["2025-2026学年 第1学期", "形势与政策", "00000080A", 0.25, "通修", 91],
  ["2025-2026学年 第1学期", "示例等级制课程", "TEST001", 2, "选修", "良好"]
];
const worksheet = XLSX.utils.aoa_to_sheet(rows);
const workbook = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(workbook, worksheet, "成绩查询");
const bytes = XLSX.write(workbook, { type: "array", bookType: "xlsx" });

const parsed = context.__testApi.readWorkbook(bytes);
assert.equal(parsed.sheetName, "成绩查询");
assert.equal(parsed.courses.length, 4);
assert.deepEqual(
  Array.from(parsed.courses, (course) => course.type),
  ["平台", "平台", "通修", "选修"]
);
assert.equal(parsed.courses[0].name, "问题求解（二）");
assert.equal(parsed.courses[0].grade, 67);
assert.equal(parsed.courses[1].gradeDisplay, "通过");
assert.equal(parsed.courses[1].grade, null);
assert.equal(context.__testApi.isSelectable(parsed.courses[1]), false);
assert.equal(parsed.courses[3].grade, 85);
assert.equal(context.__testApi.parseGrade("合格（百分制未提供）"), null);

const numeric = parsed.courses.filter((course) => context.__testApi.isSelectable(course));
const credits = numeric.reduce((sum, course) => sum + course.credit, 0);
const weighted = numeric.reduce((sum, course) => sum + course.credit * course.grade, 0);
assert.equal(credits, 6.25);
assert.equal((weighted / credits / 20).toFixed(4), "3.6860");

const html = fs.readFileSync(path.join(projectDir, "index.html"), "utf8");
for (const id of source.matchAll(/getElementById\("([^"]+)"\)/g)) {
  assert.match(html, new RegExp(`id=["']${id[1]}["']`), `index.html 缺少 #${id[1]}`);
}
assert.doesNotMatch(source, /\b(?:fetch|XMLHttpRequest|WebSocket|sendBeacon)\b/);
assert.doesNotMatch(source, /\b(?:localStorage|sessionStorage|document\.cookie)\b/);

console.log("全部测试通过：Excel 解析、课程性质、非数值成绩、GPA、页面结构和隐私检查。");
