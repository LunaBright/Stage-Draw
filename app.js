"use strict";

const $ = (id) => document.getElementById(id);
const state = { groups: [], activeId: null, serial: 0, editorDraft: null, toastTimer: null };
let activeRoll = null;
const OOXML = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const DOC_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships";

function nextId(prefix) { state.serial += 1; return prefix + state.serial; }
function groupNow() { return state.groups.find((g) => g.id === state.activeId) || null; }
function makeDraw() { return { representativeId: null, schoolNumbers: {}, programNumbers: {}, last: null }; }
function cloneDraw(draw) { return JSON.parse(JSON.stringify(draw)); }
function hasDraw(group) { return Boolean(group.draw.representativeId || Object.keys(group.draw.schoolNumbers).length || Object.keys(group.draw.programNumbers).length); }
function isComplete(group) {
  if (!group || !group.mode || !group.programs.length || !group.schools.length) return false;
  if (group.mode === "one") return Boolean(group.draw.representativeId) && Object.keys(group.draw.programNumbers).length === group.programs.length;
  return Object.keys(group.draw.schoolNumbers).length === group.schools.length && Object.keys(group.draw.programNumbers).length === group.programs.length;
}
function canStart(group) {
  return group.schools.length > 0 && group.programs.length > 0 &&
    group.schools.every((s) => s.name.trim() && group.programs.some((p) => p.schoolId === s.id)) &&
    group.programs.every((p) => p.title.trim() && group.schools.some((s) => s.id === p.schoolId));
}
function schoolName(group, id) { return group.schools.find((s) => s.id === id)?.name || "未知学校"; }
function currentPath(group) { return [group.rootName, ...group.directory].filter(Boolean).join(" / ") || "手动创建"; }
function categoryOf(group) { return group.directory[0] || "未分类"; }
function createNode(tag, className, textValue) {
  const node = document.createElement(tag);
  if (className) node.className = className;
  if (textValue !== undefined && textValue !== null) node.textContent = String(textValue);
  return node;
}
function add(parent, tag, className, textValue) { const child = createNode(tag, className, textValue); parent.appendChild(child); return child; }
function clear(node) { node.replaceChildren(); }
function notice(message) {
  const box = $("toast");
  box.textContent = message;
  box.classList.add("show");
  clearTimeout(state.toastTimer);
  state.toastTimer = setTimeout(() => box.classList.remove("show"), 4800);
}
function safeRandomIndex(length) {
  if (!Number.isInteger(length) || length < 1) throw new Error("没有可抽取的号码");
  const range = 4294967296;
  const limit = Math.floor(range / length) * length;
  const value = new Uint32Array(1);
  do { crypto.getRandomValues(value); } while (value[0] >= limit);
  return value[0] % length;
}
function shuffle(items) {
  const result = items.slice();
  for (let i = result.length - 1; i > 0; i--) {
    const j = safeRandomIndex(i + 1);
    [result[i], result[j]] = [result[j], result[i]];
  }
  return result;
}
function remainingNumbers(total, assignments) {
  const used = new Set(Object.values(assignments));
  return Array.from({ length: total }, (_, i) => i + 1).filter((n) => !used.has(n));
}
function saveHistory(group) { group.history.push(cloneDraw(group.draw)); }
function resetDraw(group) { group.draw = makeDraw(); group.history = []; group.pendingReveal = false; }
function parseXml(text, name) {
  const doc = new DOMParser().parseFromString(text, "application/xml");
  if (doc.getElementsByTagName("parsererror").length) throw new Error(name + " 的 XML 结构有误");
  return doc;
}
function nodes(parent, name) { return Array.from(parent.getElementsByTagNameNS("*", name)); }
function firstNode(parent, name) { return nodes(parent, name)[0] || null; }
function nodeText(parent, name) { return firstNode(parent, name)?.textContent || ""; }
function normalizedZipPath(path) {
  const parts = [];
  for (const piece of path.split("/")) {
    if (!piece || piece === ".") continue;
    if (piece === "..") parts.pop();
    else parts.push(piece);
  }
  return parts.join("/");
}
async function zipXml(zip, path) {
  const entry = zip.file(path);
  if (!entry) throw new Error("工作簿缺少 " + path);
  return parseXml(await entry.async("string"), path);
}
async function readXlsxRows(file) {
  const zip = await JSZip.loadAsync(await file.arrayBuffer());
  const workbook = await zipXml(zip, "xl/workbook.xml");
  const sheet = firstNode(workbook, "sheet");
  if (!sheet) throw new Error("工作簿没有工作表");
  const relId = sheet.getAttributeNS(DOC_REL, "id") || sheet.getAttribute("r:id");
  const relationships = await zipXml(zip, "xl/_rels/workbook.xml.rels");
  const relation = nodes(relationships, "Relationship").find((r) => r.getAttribute("Id") === relId);
  if (!relation) throw new Error("无法定位第一张工作表");
  const target = relation.getAttribute("Target") || "";
  const sheetPath = target.startsWith("/") ? normalizedZipPath(target) : normalizedZipPath("xl/" + target);
  const worksheet = await zipXml(zip, sheetPath);
  const shared = [];
  if (zip.file("xl/sharedStrings.xml")) {
    const strings = await zipXml(zip, "xl/sharedStrings.xml");
    for (const item of nodes(strings, "si")) shared.push(nodes(item, "t").map((n) => n.textContent).join(""));
  }
  const rows = [];
  for (const row of nodes(worksheet, "row")) {
    const values = ["", ""];
    let fallbackColumn = 0;
    for (const cell of nodes(row, "c")) {
      const address = cell.getAttribute("r") || "";
      const letters = (address.match(/^[A-Z]+/i) || [])[0];
      let column = fallbackColumn;
      if (letters) {
        column = 0;
        for (const letter of letters.toUpperCase()) column = column * 26 + letter.charCodeAt(0) - 64;
        column -= 1;
      }
      fallbackColumn = column + 1;
      if (column > 1) continue;
      const type = cell.getAttribute("t");
      let value = "";
      if (type === "inlineStr") value = nodes(cell, "t").map((n) => n.textContent).join("");
      else if (type === "s") value = shared[Number(nodeText(cell, "v"))] || "";
      else value = nodeText(cell, "v");
      values[column] = String(value).trim();
    }
    if (values.some(Boolean)) rows.push(values);
  }
  return rows;
}
function selectedFilePath(file) {
  const parts = (file.webkitRelativePath || file.name).replace(/\\/g, "/").split("/").filter(Boolean);
  const rootName = parts.length > 1 ? parts[0] : "";
  const directory = parts.length > 1 ? parts.slice(1, -1) : [];
  return { rootName, directory, name: parts[parts.length - 1] || file.name };
}
async function importFolder(fileList) {
  const files = Array.from(fileList).filter((f) => /\.xlsx$/i.test(f.name) && !f.name.startsWith("~$"));
  const warnings = [];
  const pairs = new Map();
  for (const file of files) {
    const info = selectedFilePath(file);
    const stem = info.name.replace(/\.xlsx$/i, "");
    const isSchool = stem.endsWith("-学校");
    const base = isSchool ? stem.slice(0, -3) : stem;
    const key = JSON.stringify([info.rootName, ...info.directory, base]);
    if (!pairs.has(key)) pairs.set(key, { ...info, base, programFile: null, schoolFile: null, key });
    const pair = pairs.get(key);
    const field = isSchool ? "schoolFile" : "programFile";
    if (pair[field]) warnings.push(info.name + "：出现重复文件，已使用第一份");
    else pair[field] = file;
  }
  let imported = 0;
  let skipped = 0;
  let schoolTotal = 0;
  let programTotal = 0;
  for (const pair of pairs.values()) {
    const label = [...pair.directory, pair.base].join(" / ");
    if (!pair.programFile || !pair.schoolFile) {
      warnings.push(label + "：缺少对应的" + (!pair.programFile ? "节目单" : "学校名单"));
      skipped += 1;
      continue;
    }
    if (state.groups.some((g) => g.sourceKey === pair.key)) {
      warnings.push(label + "：本页已导入，未覆盖现有数据或抽签结果");
      skipped += 1;
      continue;
    }
    try {
      const [schoolRows, programRows] = await Promise.all([readXlsxRows(pair.schoolFile), readXlsxRows(pair.programFile)]);
      const schools = [];
      const schoolIds = new Map();
      for (const row of schoolRows) {
        const name = (row[0] || "").trim();
        if (!name || name === "学校") continue;
        if (schoolIds.has(name)) continue;
        const school = { id: nextId("s"), name };
        schoolIds.set(name, school.id);
        schools.push(school);
      }
      const programs = [];
      for (const row of programRows) {
        const schoolNameText = (row[0] || "").trim();
        const title = (row[1] || "").trim();
        if (schoolNameText === "学校" && /^(节目|节目名称|作品|作品名称)$/.test(title)) continue;
        if (!schoolNameText && !title) continue;
        if (!schoolNameText || !title) { warnings.push(label + "：有一行学校或节目名称为空，已跳过"); continue; }
        if (!schoolIds.has(schoolNameText)) {
          const school = { id: nextId("s"), name: schoolNameText };
          schoolIds.set(schoolNameText, school.id);
          schools.push(school);
          warnings.push(label + "：" + schoolNameText + " 不在学校名单中，已从节目单补入");
        }
        programs.push({ id: nextId("p"), schoolId: schoolIds.get(schoolNameText), title });
      }
      if (!schools.length || !programs.length) {
        warnings.push(label + "：没有有效学校或节目，未导入");
        skipped += 1;
        continue;
      }
      const group = { id: nextId("g"), sourceKey: pair.key, rootName: pair.rootName, directory: pair.directory, name: pair.base, schools, programs, mode: null, draw: makeDraw(), history: [] };
      state.groups.push(group);
      schoolTotal += schools.length;
      programTotal += programs.length;
      imported += 1;
      if (!state.activeId) state.activeId = group.id;
    } catch (error) {
      warnings.push(label + "：读取失败（" + (error?.message || String(error)) + "）");
      skipped += 1;
    }
  }
  renderAll();
  showImportReport(imported, skipped, schoolTotal, programTotal, warnings);
}
function showImportReport(imported, skipped, schools, programs, warnings) {
  $("importTitle").textContent = imported ? "文件夹已读取" : "没有导入可用组别";
  $("importSummary").textContent = "新增 " + imported + " 个组别、" + schools + " 所学校、" + programs + " 条节目。" + (skipped ? "另有 " + skipped + " 个文件组未导入。" : "");
  clear($("importWarnings"));
  for (const warning of warnings) add($("importWarnings"), "p", "", warning);
  $("importDialog").showModal();
}
function orderedSchools(group) {
  return group.schools.slice().sort((a, b) => (group.draw.schoolNumbers[a.id] || 999999) - (group.draw.schoolNumbers[b.id] || 999999));
}
function nextProgram(group) {
  for (const school of orderedSchools(group)) {
    for (const program of group.programs) {
      if (program.schoolId === school.id && !group.draw.programNumbers[program.id]) return { school, program };
    }
  }
  return null;
}
function resultRows(group) {
  return group.programs.filter((p) => group.draw.programNumbers[p.id]).map((p) => ({
    number: group.draw.programNumbers[p.id], title: p.title, school: schoolName(group, p.schoolId)
  })).sort((a, b) => a.number - b.number);
}
function renderNavigation() {
  const select = $("categoryFilter");
  const selected = select.value;
  const categories = Array.from(new Set(state.groups.map(categoryOf))).sort((a, b) => a.localeCompare(b, "zh-CN"));
  clear(select);
  const all = createNode("option", "", "全部类别");
  all.value = "all";
  select.appendChild(all);
  for (const category of categories) {
    const option = createNode("option", "", category);
    option.value = category;
    select.appendChild(option);
  }
  select.value = categories.includes(selected) ? selected : "all";
  const search = $("searchInput").value.trim().toLocaleLowerCase();
  const visible = state.groups.filter((g) => {
    if (select.value !== "all" && categoryOf(g) !== select.value) return false;
    return !search || (currentPath(g) + " " + g.name).toLocaleLowerCase().includes(search);
  }).sort((a, b) => (categoryOf(a) + "/" + a.name).localeCompare(categoryOf(b) + "/" + b.name, "zh-CN"));
  const list = $("groupList");
  clear(list);
  if (!visible.length) {
    add(list, "p", "group-list-empty", state.groups.length ? "没有符合筛选条件的组别。" : "还没有导入组别。");
    return;
  }
  let previous = "";
  for (const group of visible) {
    const category = categoryOf(group);
    if (category !== previous) add(list, "div", "group-list-head", category);
    previous = category;
    const button = add(list, "button", "group-item" + (group.id === state.activeId ? " active" : ""));
    button.type = "button";
    if (isComplete(group)) add(button, "span", "done-dot", "✓");
    add(button, "strong", "", group.name);
    add(button, "small", "", currentPath(group));
    button.addEventListener("click", () => { cancelRoll(); state.activeId = group.id; renderAll(); });
  }
}
function stageFrame(panel, view) {
  clear(panel);
  panel.classList.toggle("is-rolling", Boolean(activeRoll && activeRoll.groupId === groupNow()?.id));
  add(panel, "div", "stage-kicker", view.step);
  const heading = add(panel, "h2", "stage-title");
  if (view.before) add(heading, "span", "", view.before);
  if (view.highlight) add(heading, "span", "stage-title-highlight", view.highlight);
  if (view.after) add(heading, "span", "", view.after);
  if (heading.textContent.length > 23) heading.classList.add("compact");
  add(panel, "p", "stage-copy", view.copy || " ");
  const previous = add(panel, "div", "stage-previous");
  add(previous, "span", "stage-previous-label", "上一签");
  if (view.previous) {
    add(previous, "strong", "stage-previous-name", view.previous.name);
    if (view.previous.number) add(previous, "b", "stage-previous-number", view.previous.number + " 号");
  } else add(previous, "span", "stage-previous-empty", "暂无结果");
  const focus = add(panel, "div", "stage-focus");
  const focusLabel = add(focus, "div", "stage-focus-label", view.focusLabel || "即将抽取");
  if (view.focusSchool) add(focusLabel, "strong", "stage-focus-school", view.focusSchool);
  const focusValue = add(focus, "div", "stage-focus-value", view.focusValue || "—");
  focusValue.id = "stagePreview";
  focusValue.setAttribute("aria-live", "off");
  const controls = add(panel, "div", "stage-controls");
  const selectSlot = add(controls, "div", "stage-select-slot");
  const button = add(controls, "button", "stage-action", view.button || "等待开始");
  button.id = "stageActionButton";
  button.type = "button";
  button.disabled = !view.onAction;
  if (view.onAction) button.addEventListener("click", view.onAction);
  const progress = add(panel, "div", "stage-progress-area");
  const bar = add(progress, "div", "stage-progress");
  const fill = add(bar, "span");
  fill.style.width = (view.total ? Math.round(view.done / view.total * 100) : 0) + "%";
  add(progress, "div", "stage-step", view.progress || ("已完成 " + (view.done || 0) + " / " + (view.total || 0)));
  if (activeRoll) updateRollingPreview();
}
function rollingAction(group, kind, targetId, candidates) {
  if (activeRoll) return finishRoll(group);
  beginRoll(group, kind, targetId, candidates);
}
function continueAfterReveal(group) {
  group.pendingReveal = null;
  renderAll();
}
function previousLabel(last) {
  if (!last) return null;
  if (last.kind === "representative") return { name: last.schoolName };
  if (last.kind === "schools") return { name: last.count + " 所学校顺序已确定" };
  if (last.kind === "program") return { name: last.programTitle, number: last.number };
  if (last.kind === "all") return { name: last.count + " 条节目已定序" };
  return null;
}
function renderReveal(group) {
  const last = group.draw.last;
  const prior = group.pendingReveal?.previous;
  if (!last) return false;
  const next = () => continueAfterReveal(group);
  if (last.kind === "representative") {
    stageFrame($("stageContent"), {
      step: "方式一 · 第一步结果", before: "抽中代表学校",
      previous: previousLabel(prior),
      copy: "请由这所学校的代表进行下一步操作。",
      focusLabel: "本次抽中学校", focusValue: last.schoolName,
      button: "进入第二步 · 抽节目顺序", onAction: next,
      total: 2, done: 1, progress: "代表学校已确定"
    });
  } else if (last.kind === "all") {
    stageFrame($("stageContent"), {
      step: "方式一 · 定序结果", before: "全组节目顺序已确定",
      previous: previousLabel(prior),
      copy: "本组全部节目已获得 1～" + group.programs.length + " 的唯一出场序号。",
      focusLabel: "本次抽签结果", focusValue: last.count + " 条节目已定序",
      button: "查看完整结果", onAction: next,
      total: 2, done: 2, progress: "抽签完成"
    });
  } else if (last.kind === "schools") {
    stageFrame($("stageContent"), {
      step: "方式二 · 第一轮结果", before: "全部学校顺序已确定",
      previous: previousLabel(prior),
      copy: "右侧显示本组学校的完整抽签顺序。确认后按此顺序逐校抽取节目序号。",
      focusLabel: "本次抽签结果",
      focusValue: last.count + " 所学校已排定顺序",
      button: "进入第二轮 · 抽节目", onAction: next,
      total: 2, done: 1,
      progress: "第一轮已完成"
    });
  } else if (last.kind === "program") {
    const done = Object.keys(group.draw.programNumbers).length;
    const allDone = done === group.programs.length;
    stageFrame($("stageContent"), {
      step: "方式二 · 第二轮结果", before: "节目序号已抽出",
      previous: previousLabel(prior),
      copy: allDone ? "本组所有节目已定序。点击查看完整结果。" : "请确认本次结果，再开始抽取下一条节目。",
      focusLabel: last.programTitle + " · " + last.schoolName,
      focusValue: last.number + " 号",
      button: allDone ? "查看本组结果" : "前往下一条节目", onAction: next,
      total: group.programs.length, done,
      progress: "节目序号已完成 " + done + " / " + group.programs.length
    });
  }
  return true;
}
function renderStage(group) {
  const panel = $("stageContent");
  const roll = activeRoll?.groupId === group.id ? activeRoll : null;
  const action = (kind, targetId, candidates, label) => ({
    button: roll ? "停止并确定结果" : label,
    onAction: () => rollingAction(group, kind, targetId, candidates)
  });
  if (!group.mode) {
    stageFrame(panel, { step: "准备阶段", before: "请选择一种抽签方式", copy: "先核对学校与节目名单，再选择上方的抽签方式。", focusLabel: "本组名单", focusValue: group.schools.length + " 所学校 · " + group.programs.length + " 条节目", progress: "选择方式后开始抽签" });
    return;
  }
  if (!canStart(group)) {
    stageFrame(panel, { step: "名单待完善", before: "请先补全学校与节目", copy: "每所学校至少需要一个节目，且每条节目都需要指定学校。", focusLabel: "当前状态", focusValue: "名单尚不完整", button: "编辑名单", onAction: () => openEditor(group), progress: "名单完善后即可抽签" });
    return;
  }
  if (group.pendingReveal && renderReveal(group)) return;
  if (group.mode === "one") {
    if (!group.draw.representativeId) {
      stageFrame(panel, { step: "方式一 · 第一步", before: "随机选出代表学校", copy: "工作人员开始滚动学校名单，由现场人员点击停止。", focusLabel: "待选学校", focusValue: "本组共 " + group.schools.length + " 所学校", ...action("representative", null, group.schools.map((s) => s.id), "开始抽取代表学校"), total: 2, done: 0, progress: "第一步 / 共两步" });
    } else if (!isComplete(group)) {
      stageFrame(panel, { step: "方式一 · 第二步", before: "由 ", highlight: schoolName(group, group.draw.representativeId), after: " 确定节目顺序", copy: "开始后全组节目持续随机排序，点击停止即一次确定 1～" + group.programs.length + " 号。", previous: { name: schoolName(group, group.draw.representativeId) }, focusLabel: "本组代表学校 · " + schoolName(group, group.draw.representativeId), focusValue: "点击开始，预览全组节目", ...action("all", null, group.programs.map((p) => p.id), "开始随机排序"), total: 2, done: 1, progress: "第二步 / 共两步" });
    } else {
      stageFrame(panel, { step: "方式一 · 已完成", before: "本组节目顺序已确定", copy: "请在下方核对结果并导出 Excel。", previous: previousLabel(group.draw.last), focusLabel: "已完成", focusValue: group.programs.length + " 条节目已定序", total: 2, done: 2, progress: "抽签完成" });
    }
    return;
  }
  const schoolDone = Object.keys(group.draw.schoolNumbers).length;
  const programDone = Object.keys(group.draw.programNumbers).length;
  if (schoolDone < group.schools.length) {
    stageFrame(panel, { step: "方式二 · 第一轮", before: "一次抽出全部学校顺序", copy: "由一位工作人员点击开始，右侧学校名单快速随机排列；点击停止后，一次确定所有学校的抽签顺序。", previous: previousLabel(group.draw.last), focusLabel: "全组学校顺序", focusValue: "等待抽取 " + group.schools.length + " 所学校", ...action("schools", null, group.schools.map((s) => s.id), "开始抽取学校顺序"), total: 2, done: 0, progress: "第一轮 / 共两轮" });
    return;
  }
  if (programDone < group.programs.length) {
    const current = nextProgram(group);
    if (!current) return;
    const previous = previousLabel(group.draw.last);
    stageFrame(panel, { step: "方式二 · 第二轮", before: "轮到 ", highlight: current.school.name, after: " 抽节目", copy: "按学校顺序，为当前节目从全组共用的剩余号码中抽取出场序号。", previous, focusLabel: "当前节目 · 第 " + group.draw.schoolNumbers[current.school.id] + " 位学校 · ", focusSchool: current.school.name, focusValue: current.program.title, ...action("program", current.program.id, remainingNumbers(group.programs.length, group.draw.programNumbers), "开始抽取节目序号"), total: group.programs.length, done: programDone });
    return;
  }
  stageFrame(panel, { step: "方式二 · 已完成", before: "本组节目顺序已确定", copy: "学校与节目两轮抽签全部完成。请在下方核对结果并导出 Excel。", previous: previousLabel(group.draw.last), focusLabel: "已完成", focusValue: group.programs.length + " 条节目已定序", total: group.programs.length, done: group.programs.length, progress: "抽签完成" });
}
function renderSide(group) {
  const side = $("drawSidePanel");
  clear(side);
  const schoolDone = Object.keys(group.draw.schoolNumbers).length;
  let current = group.mode === "two" && schoolDone === group.schools.length && !group.pendingReveal ? nextProgram(group) : null;
  if (!current && group.pendingReveal && group.draw.last?.kind === "program") {
    const program = group.programs.find((p) => p.id === group.draw.last.programId);
    const school = group.schools.find((s) => s.id === program?.schoolId);
    if (school && group.programs.some((p) => p.schoolId === school.id && !group.draw.programNumbers[p.id])) current = { school, program };
  }
  if (current) {
    add(side, "h3", "", current.school.name + " 的节目");
    const allPrograms = group.programs.filter((p) => p.schoolId === current.school.id);
    const done = allPrograms.filter((p) => group.draw.programNumbers[p.id]).length;
    add(side, "p", "", "该校共 " + allPrograms.length + " 条节目，已抽 " + done + " 条。按下列顺序逐个抽签。");
    const programList = add(side, "div", "info-list");
    for (const program of allPrograms) {
      const row = add(programList, "div", "info-row" + (program.id === current.program.id ? " current" : ""));
      add(row, "span", "", program.title);
      add(row, "b", "", group.draw.programNumbers[program.id] ? group.draw.programNumbers[program.id] + " 号" : program.id === current.program.id ? "当前" : "待抽");
    }
    const subtitle = add(side, "h3", "", "学校抽签顺序");
    subtitle.style.marginTop = "22px";
  } else {
    add(side, "h3", "", group.mode === "two" ? "学校抽签顺序" : "本组概况");
  }
  if (group.mode !== "two") {
    const stat1 = add(side, "div", "mini-stat"); add(stat1, "span", "", "学校"); add(stat1, "strong", "", group.schools.length + " 所");
    const stat2 = add(side, "div", "mini-stat"); add(stat2, "span", "", "节目"); add(stat2, "strong", "", group.programs.length + " 条");
    if (group.mode === "one") {
      const stat3 = add(side, "div", "mini-stat"); add(stat3, "span", "", "代表学校"); add(stat3, "strong", "", group.draw.representativeId ? schoolName(group, group.draw.representativeId) : "待抽取");
      add(side, "p", "", "节目序号在本组内唯一。导入的每一行节目作为一条抽签记录，单元格中的复合标题会完整保留。");
    } else {
      add(side, "p", "", "方式一由一所代表学校一次性抽出全部节目顺序。方式二一次排定全部学校顺序，再按顺序逐校抽节目序号。");
    }
    return;
  }
  const schoolList = add(side, "div", "info-list");
  const previewOrder = activeRoll?.groupId === group.id && activeRoll.kind === "schools" ? activeRoll.preview : null;
  const schools = previewOrder ? previewOrder.map((id) => group.schools.find((s) => s.id === id)).filter(Boolean) : orderedSchools(group);
  for (const [index, school] of schools.entries()) {
    const row = add(schoolList, "div", "info-row" + (current?.school.id === school.id ? " current" : ""));
    add(row, "span", "", school.name);
    add(row, "b", "", previewOrder ? index + 1 + " 位" : group.draw.schoolNumbers[school.id] ? group.draw.schoolNumbers[school.id] + " 位" : "待抽");
  }
  if (current) add(side, "p", "", "当前轮到第 " + group.draw.schoolNumbers[current.school.id] + " 位学校：" + current.school.name + "。");
}function renderResults(group) {
  const entries = resultRows(group);
  $("resultCount").textContent = entries.length + " / " + group.programs.length + " 已确定";
  const box = $("resultsContent");
  clear(box);
  if (!entries.length) { add(box, "div", "results-empty", "节目出场序号尚未产生。"); return; }
  const wrap = add(box, "div", "result-table-wrap");
  const table = add(wrap, "table", "result-table");
  const head = add(table, "thead");
  const header = add(head, "tr");
  for (const title of ["序号", "节目", "学校"]) add(header, "th", "", title);
  const body = add(table, "tbody");
  for (const entry of entries) {
    const row = add(body, "tr");
    add(row, "td", "", entry.number);
    add(row, "td", "", entry.title);
    add(row, "td", "", entry.school);
  }
}
function renderAll() {
  $("topStatus").textContent = state.groups.length ? "本页已有 " + state.groups.length + " 个组别" : "当前页面暂无数据";
  renderNavigation();
  const group = groupNow();
  $("emptyState").hidden = Boolean(group);
  $("groupView").hidden = !group;
  if (!group) return;
  $("groupPath").textContent = currentPath(group);
  $("groupTitle").textContent = group.name;
  $("groupCounts").textContent = group.schools.length + " 所学校 · " + group.programs.length + " 条节目";
  const status = $("groupStatus");
  status.textContent = isComplete(group) ? "已完成" : activeRoll?.groupId === group.id || hasDraw(group) ? "进行中" : "待开始";
  status.className = "status-badge" + (isComplete(group) ? " completed" : activeRoll?.groupId === group.id || hasDraw(group) ? " in-progress" : "");
  $("undoButton").disabled = Boolean(activeRoll) || !group.history.length;
  $("resetButton").disabled = Boolean(activeRoll) || !hasDraw(group);
  $("exportButton").disabled = Boolean(activeRoll) || !isComplete(group);
  $("editButton").disabled = Boolean(activeRoll);
  for (const button of document.querySelectorAll(".mode-button")) { button.classList.toggle("active", button.dataset.mode === group.mode); button.disabled = Boolean(activeRoll); }
  renderStage(group);
  renderSide(group);
  renderResults(group);
}
function chooseMode(group, mode) {
  if (activeRoll || group.mode === mode) return;
  if (hasDraw(group) && !window.confirm("切换抽签方式将清除当前组已抽出的号码。确认切换吗？")) return;
  resetDraw(group);
  group.mode = mode;
  renderAll();
}
function cancelRoll() {
  if (!activeRoll) return;
  clearInterval(activeRoll.timer);
  activeRoll = null;
}
function rollPreview(roll) {
  if (roll.kind === "all" || roll.kind === "schools") return shuffle(roll.candidates);
  return roll.candidates[safeRandomIndex(roll.candidates.length)];
}
function updateRollingPreview() {
  const roll = activeRoll;
  if (!roll || groupNow()?.id !== roll.groupId) return;
  const group = groupNow();
  const preview = $("stagePreview");
  const button = $("stageActionButton");
  if (!preview || !button) return;
  if (roll.kind === "representative") preview.textContent = schoolName(group, roll.preview);
  else if (roll.kind === "all") {
    const program = group.programs.find((p) => p.id === roll.preview[0]);
    preview.textContent = program?.title || "";
  } else if (roll.kind === "schools") {
    preview.textContent = "第 1 位 · " + schoolName(group, roll.preview[0]);
    renderSide(group);
  } else preview.textContent = roll.preview + " 号";
  preview.classList.add("is-flashing");
  button.disabled = false;
  button.textContent = "停止并确定结果";
}
function beginRoll(group, kind, targetId, candidates) {
  if (activeRoll || group.pendingReveal || groupNow()?.id !== group.id || !candidates.length) return;
  if (kind === "representative" && (group.mode !== "one" || group.draw.representativeId)) return;
  if (kind === "all" && (group.mode !== "one" || !group.draw.representativeId || isComplete(group))) return;
  if (kind === "schools" && (group.mode !== "two" || Object.keys(group.draw.schoolNumbers).length === group.schools.length)) return;
  if (kind === "program" && (group.mode !== "two" || nextProgram(group)?.program.id !== targetId)) return;
  activeRoll = { groupId: group.id, kind, targetId, candidates: candidates.slice(), preview: null, timer: null };
  activeRoll.preview = rollPreview(activeRoll);
  renderAll();
  activeRoll.timer = setInterval(() => {
    if (!activeRoll || groupNow()?.id !== group.id) return;
    activeRoll.preview = rollPreview(activeRoll);
    updateRollingPreview();
  }, kind === "all" ? 90 : kind === "schools" ? 70 : 60);
}
function finishRoll(group) {
  const roll = activeRoll;
  if (!roll || roll.groupId !== group.id) return;
  const result = roll.preview;
  cancelRoll();
  const previous = group.draw.last ? { ...group.draw.last } : null;
  saveHistory(group);
  if (roll.kind === "representative") {
    group.draw.representativeId = result;
    group.draw.last = { kind: "representative", schoolName: schoolName(group, result) };
  } else if (roll.kind === "all") {
    result.forEach((id, index) => { group.draw.programNumbers[id] = index + 1; });
    group.draw.last = { kind: "all", count: result.length };
  } else if (roll.kind === "schools") {
    result.forEach((id, index) => { group.draw.schoolNumbers[id] = index + 1; });
    group.draw.last = { kind: "schools", count: result.length };
  } else if (roll.kind === "program") {
    const program = group.programs.find((p) => p.id === roll.targetId);
    group.draw.programNumbers[roll.targetId] = result;
    group.draw.last = { kind: "program", schoolName: schoolName(group, program.schoolId), programTitle: program.title, programId: program.id, number: result };
  }
  group.pendingReveal = { previous };
  renderAll();
}
function undoLast(group) {
  if (activeRoll || !group.history.length) return;
  if (!window.confirm("确定撤销上一签吗？这签的结果将取消，可重新抽取。")) return;
  const undoneKind = group.draw.last?.kind;
  group.draw = group.history.pop();
  group.pendingReveal = false;
  renderAll();
  notice(undoneKind === "schools" ? "已撤销整轮学校顺序，可以重新一次抽取。" : "已撤销上一签，号码已返回可抽取号码池。");
}
function resetGroup(group) {
  if (activeRoll || !hasDraw(group)) return;
  if (!window.confirm("确定整组重新抽吗？本组已抽出的学校顺序和节目序号都将清空，名单会保留。")) return;
  resetDraw(group);
  renderAll();
  notice("本组抽签进度已清空，可以重新开始。");
}
function openEditor(group) {
  state.editorDraft = {
    groupId: group.id,
    directory: group.directory.slice(),
    name: group.name,
    schools: group.schools.map((s) => ({ ...s })),
    programs: group.programs.map((p) => ({ ...p }))
  };
  $("editPath").value = group.directory.join(" / ");
  $("editName").value = group.name;
  $("editorError").textContent = "";
  renderEditorRows();
  $("editorDialog").showModal();
}
function closeEditor() {
  $("editorDialog").close();
  state.editorDraft = null;
}
function editorError(message) { $("editorError").textContent = message; }
function renderEditorRows() {
  const draft = state.editorDraft;
  if (!draft) return;
  const schoolBox = $("schoolEditorRows");
  const programBox = $("programEditorRows");
  clear(schoolBox); clear(programBox);
  for (const school of draft.schools) {
    const row = add(schoolBox, "div", "edit-row");
    const input = add(row, "input", "field-control");
    input.value = school.name;
    input.placeholder = "学校名称";
    input.setAttribute("aria-label", "学校名称");
    input.addEventListener("input", () => {
      school.name = input.value;
      for (const option of programBox.querySelectorAll("option")) {
        if (option.value === school.id) option.textContent = school.name.trim() || "未命名学校";
      }
    });
    const remove = add(row, "button", "small-delete", "×");
    remove.type = "button";
    remove.setAttribute("aria-label", "删除学校 " + school.name);
    remove.addEventListener("click", () => {
      const count = draft.programs.filter((p) => p.schoolId === school.id).length;
      if (count && !window.confirm("删除这所学校会同时删除其 " + count + " 条节目。确认吗？")) return;
      draft.schools = draft.schools.filter((s) => s.id !== school.id);
      draft.programs = draft.programs.filter((p) => p.schoolId !== school.id);
      renderEditorRows();
    });
  }
  for (const program of draft.programs) {
    const row = add(programBox, "div", "edit-row");
    const schoolSelect = add(row, "select", "field-control");
    schoolSelect.setAttribute("aria-label", "节目所属学校");
    for (const school of draft.schools) {
      const option = add(schoolSelect, "option", "", school.name.trim() || "未命名学校");
      option.value = school.id;
    }
    schoolSelect.value = program.schoolId;
    schoolSelect.addEventListener("change", () => { program.schoolId = schoolSelect.value; });
    const input = add(row, "input", "field-control");
    input.value = program.title;
    input.placeholder = "节目名称";
    input.setAttribute("aria-label", "节目名称");
    input.addEventListener("input", () => { program.title = input.value; });
    const remove = add(row, "button", "small-delete", "×");
    remove.type = "button";
    remove.setAttribute("aria-label", "删除节目 " + program.title);
    remove.addEventListener("click", () => {
      draft.programs = draft.programs.filter((p) => p.id !== program.id);
      renderEditorRows();
    });
  }
}
function saveEditor() {
  const draft = state.editorDraft;
  if (!draft) return;
  const group = state.groups.find((g) => g.id === draft.groupId);
  if (!group) return;
  const name = $("editName").value.trim();
  const directory = $("editPath").value.split("/").map((part) => part.trim()).filter(Boolean);
  if (!name) { editorError("请填写组别名称。"); return; }
  const names = draft.schools.map((s) => s.name.trim());
  if (names.some((n) => !n)) { editorError("学校名称不能为空。"); return; }
  if (new Set(names).size !== names.length) { editorError("同一组内学校名称不能重复。"); return; }
  for (const program of draft.programs) {
    if (!program.title.trim()) { editorError("节目名称不能为空。"); return; }
    if (!draft.schools.some((s) => s.id === program.schoolId)) { editorError("每个节目必须关联一所学校。"); return; }
  }
  const duplicate = state.groups.some((g) => g.id !== group.id && g.rootName === group.rootName && g.name === name && JSON.stringify(g.directory) === JSON.stringify(directory));
  if (duplicate) { editorError("本页已有相同分类路径与名称的组别。"); return; }
  const changed = name !== group.name || JSON.stringify(directory) !== JSON.stringify(group.directory) || JSON.stringify(draft.schools) !== JSON.stringify(group.schools) || JSON.stringify(draft.programs) !== JSON.stringify(group.programs);
  if (changed && hasDraw(group) && !window.confirm("修改名单会清空当前组已抽出的号码。确认保存修改吗？")) return;
  group.name = name;
  group.directory = directory;
  group.schools = draft.schools.map((s) => ({ id: s.id, name: s.name.trim() }));
  group.programs = draft.programs.map((p) => ({ id: p.id, schoolId: p.schoolId, title: p.title.trim() }));
  if (changed) resetDraw(group);
  closeEditor();
  renderAll();
  notice("名单修改已保存在当前页面。");
}
function openNewGroup() {
  $("newGroupForm").reset();
  $("newGroupError").textContent = "";
  $("newGroupDialog").showModal();
}
function createManualGroup() {
  const name = $("newName").value.trim();
  const directory = $("newPath").value.split("/").map((x) => x.trim()).filter(Boolean);
  if (!name) { $("newGroupError").textContent = "请填写组别名称。"; return; }
  if (state.groups.some((g) => !g.rootName && g.name === name && JSON.stringify(g.directory) === JSON.stringify(directory))) {
    $("newGroupError").textContent = "本页已有同名组别。"; return;
  }
  const group = { id: nextId("g"), sourceKey: null, rootName: "", directory, name, schools: [], programs: [], mode: null, draw: makeDraw(), history: [] };
  state.groups.push(group);
  state.activeId = group.id;
  $("newGroupDialog").close();
  renderAll();
  openEditor(group);
}
function xmlEscape(value) {
  const clean = Array.from(String(value)).filter((character) => {
    const code = character.codePointAt(0);
    return code === 9 || code === 10 || code === 13 || code >= 32;
  }).join("");
  return clean.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&apos;");
}
function excelInlineCell(address, value) {
  return '<c r="' + address + '" t="inlineStr"><is><t xml:space="preserve">' + xmlEscape(value) + '</t></is></c>';
}
async function buildResultXlsx(group) {
  const entries = resultRows(group);
  const data = [["序号", "节目", "学校"], ...entries.map((r) => [r.number, r.title, r.school])];
  const rowXml = data.map((values, index) => {
    const rowNumber = index + 1;
    const first = index === 0 ? excelInlineCell("A" + rowNumber, values[0]) : '<c r="A' + rowNumber + '"><v>' + values[0] + '</v></c>';
    return '<row r="' + rowNumber + '">' + first + excelInlineCell("B" + rowNumber, values[1]) + excelInlineCell("C" + rowNumber, values[2]) + '</row>';
  }).join("");
  const header = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
  const zip = new JSZip();
  zip.file("[Content_Types].xml", header + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>');
  zip.file("_rels/.rels", header + '<Relationships xmlns="' + PKG_REL + '"><Relationship Id="rId1" Type="' + DOC_REL + '/officeDocument" Target="xl/workbook.xml"/></Relationships>');
  zip.file("xl/workbook.xml", header + '<workbook xmlns="' + OOXML + '" xmlns:r="' + DOC_REL + '"><sheets><sheet name="抽签结果" sheetId="1" r:id="rId1"/></sheets></workbook>');
  zip.file("xl/_rels/workbook.xml.rels", header + '<Relationships xmlns="' + PKG_REL + '"><Relationship Id="rId1" Type="' + DOC_REL + '/worksheet" Target="worksheets/sheet1.xml"/></Relationships>');
  zip.file("xl/worksheets/sheet1.xml", header + '<worksheet xmlns="' + OOXML + '"><cols><col min="1" max="1" width="10" customWidth="1"/><col min="2" max="2" width="44" customWidth="1"/><col min="3" max="3" width="30" customWidth="1"/></cols><sheetData>' + rowXml + '</sheetData></worksheet>');
  return zip.generateAsync({ type: "blob", mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", compression: "DEFLATE" });
}
async function exportCurrentGroup(group) {
  if (!isComplete(group)) return;
  const blob = await buildResultXlsx(group);
  const rawName = [...group.directory, group.name, "抽签结果"].filter(Boolean).join("_");
  const fileName = rawName.replaceAll("/", "_").replaceAll("\\", "_").replace(/[<>:"|?*]/g, "_").slice(0, 120) + ".xlsx";
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = fileName;
  document.body.appendChild(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(url), 30000);
  notice("已开始下载当前组的 Excel 结果。");
}
function bindEvents() {
  const chooseFolder = () => { const wasRolling = Boolean(activeRoll); cancelRoll(); if (wasRolling) renderAll(); $("folderInput").click(); };
  $("importButton").addEventListener("click", chooseFolder);
  $("emptyImportButton").addEventListener("click", chooseFolder);
  $("folderInput").addEventListener("change", async (event) => {
    const files = event.target.files;
    if (!files || !files.length) return;
    $("topStatus").textContent = "正在读取文件夹…";
    $("importButton").disabled = true;
    try { await importFolder(files); }
    catch (error) { notice("导入失败：" + (error?.message || String(error))); }
    finally { event.target.value = ""; $("importButton").disabled = false; renderAll(); }
  });
  $("newGroupButton").addEventListener("click", () => { const wasRolling = Boolean(activeRoll); cancelRoll(); if (wasRolling) renderAll(); openNewGroup(); });
  $("emptyNewGroupButton").addEventListener("click", () => { cancelRoll(); openNewGroup(); });
  $("categoryFilter").addEventListener("change", renderNavigation);
  $("searchInput").addEventListener("input", renderNavigation);
  for (const button of document.querySelectorAll(".mode-button")) {
    button.addEventListener("click", () => { const group = groupNow(); if (group) chooseMode(group, button.dataset.mode); });
  }
  $("editButton").addEventListener("click", () => { const group = groupNow(); if (group) openEditor(group); });
  $("undoButton").addEventListener("click", () => { const group = groupNow(); if (group) undoLast(group); });
  $("resetButton").addEventListener("click", () => { const group = groupNow(); if (group) resetGroup(group); });
  $("exportButton").addEventListener("click", async () => {
    const group = groupNow();
    if (!group) return;
    try { await exportCurrentGroup(group); }
    catch (error) { notice("导出失败：" + (error?.message || String(error))); }
  });
  $("presentationButton").addEventListener("click", async () => {
    try {
      if (document.fullscreenElement) await document.exitFullscreen();
      else await document.documentElement.requestFullscreen();
    } catch { notice("浏览器未能进入全屏，可使用 F11 键。"); }
  });
  document.addEventListener("fullscreenchange", () => {
    const full = Boolean(document.fullscreenElement);
    document.body.classList.toggle("is-fullscreen", full);
    $("presentationButton").textContent = full ? "退出全屏" : "全屏展示";
  });
  const form = $("newGroupForm");
  for (const cancel of form.querySelectorAll('[value="cancel"]')) {
    cancel.formNoValidate = true;
    cancel.addEventListener("click", (event) => { event.preventDefault(); $("newGroupDialog").close(); });
  }
  form.addEventListener("submit", (event) => { event.preventDefault(); createManualGroup(); });
  $("closeEditorButton").addEventListener("click", closeEditor);
  $("cancelEditorButton").addEventListener("click", closeEditor);
  $("saveEditorButton").addEventListener("click", saveEditor);
  $("addSchoolButton").addEventListener("click", () => {
    if (!state.editorDraft) return;
    state.editorDraft.schools.push({ id: nextId("s"), name: "" });
    renderEditorRows();
    $("schoolEditorRows").lastElementChild?.querySelector("input")?.focus();
  });
  $("addProgramButton").addEventListener("click", () => {
    const draft = state.editorDraft;
    if (!draft) return;
    if (!draft.schools.length) { editorError("请先添加一所学校。"); return; }
    editorError("");
    draft.programs.push({ id: nextId("p"), schoolId: draft.schools[0].id, title: "" });
    renderEditorRows();
    $("programEditorRows").lastElementChild?.querySelector("input")?.focus();
  });
  const closeImport = () => $("importDialog").close();
  $("closeImportButton").addEventListener("click", closeImport);
  $("importDoneButton").addEventListener("click", closeImport);
  window.addEventListener("beforeunload", (event) => {
    if (!state.groups.length && !state.editorDraft) return;
    event.preventDefault();
    event.returnValue = true;
  });
}
bindEvents();
renderAll();
