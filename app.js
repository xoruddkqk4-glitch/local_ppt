const $ = (selector) => document.querySelector(selector);
const stage = $("#presentationStage");
const layouts = window.PptLayouts || {};
const diagrams = window.PptDiagrams || {};
const designs = window.PptDesigns || {};

const MAX_HISTORY = 50;
const MIN_ITEMS = 1;
const MAX_ITEMS = 12;
const GRID_SIZE = 5;
const SNAP_DISTANCE_PX = 8;
const PROJECT_FORMAT = "local-ppt-json";
const PROJECT_VERSION = 1;
const PROJECT_PICKER_ID = "local-ppt-project";

let currentProjectFileHandle = null;
let currentProjectFileName = "local-ppt.txt";

const state = {
  design: "bauhaus",
  pages: [createCoverPage()],
  currentPageIndex: 0,
  selectedIds: new Set(),
  activeTextObjectId: null,
  guides: [],
  history: []
};

function createId(prefix = "object") {
  return `${prefix}-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
}

function createTextObject(role, text, x, y, w, h, extra = {}) {
  return { id: createId(role), type: "text", role, text, x, y, w, h, ...extra };
}

function createCoverPage() {
  return {
    id: createId("page"),
    type: "cover",
    template: "cover",
    objectCategory: null,
    variant: null,
    objects: [
      createTextObject("cover-title", "PRESENTATION\nTITLE", 12, 27, 76, 25),
      createTextObject("cover-subtitle", "더블클릭하여 부제목을 입력하세요", 20, 59, 60, 10)
    ]
  };
}

function createContentPage() {
  return {
    id: createId("page"),
    type: "content",
    template: null,
    objectCategory: "layout",
    variant: "cards",
    objects: []
  };
}

function currentPage() {
  return state.pages[state.currentPageIndex];
}

function snapshot() {
  state.history.push(JSON.stringify({
    design: state.design,
    pages: state.pages,
    currentPageIndex: state.currentPageIndex
  }));
  if (state.history.length > MAX_HISTORY) state.history.shift();
  updateUndoButton();
}

function undo() {
  const value = state.history.pop();
  if (!value) return;
  const restored = JSON.parse(value);
  state.design = restored.design;
  state.pages = restored.pages;
  state.currentPageIndex = restored.currentPageIndex;
  state.selectedIds.clear();
  state.guides = [];
  hideTextToolbar();
  render();
}

function updateUndoButton() {
  $("#undoButton").disabled = state.history.length === 0;
}

function buildTemplate(page, template, options = {}) {
  page.template = template;
  page.objects = page.objects.filter((object) => object.type === "image");

  if (template === "bullet") {
    page.objects.unshift(
      createTextObject("page-title", "핵심 메시지", 8, 10, 72, 18),
      createTextObject("bullet-item", "첫 번째 핵심 내용", 12, 0, 75, 9, { item: true }),
      createTextObject("bullet-item", "두 번째 핵심 내용", 12, 0, 75, 9, { item: true }),
      createTextObject("bullet-item", "세 번째 핵심 내용", 12, 0, 75, 9, { item: true })
    );
    layoutBulletItems(page);
  }

  if (template === "mindmap") {
    createDefaultMindmap(page);
  }

  if (template === "object") {
    page.objectCategory = options.category || page.objectCategory || "layout";
    page.variant = options.variant || (page.objectCategory === "layout" ? "cards" : "process");
    buildObjectTemplate(page, 3);
  }
}

function layoutBulletItems(page) {
  const items = page.objects.filter((object) => object.role === "bullet-item");
  if (!items.length) return;
  const top = 35;
  const bottom = 88;
  const gap = items.length > 10 ? .5 : items.length > 8 ? 1 : 2;
  const height = clamp(4, (bottom - top - gap * (items.length - 1)) / items.length, 10);
  items.forEach((item, index) => {
    item.x = 12;
    item.y = top + index * (height + gap);
    item.w = 75;
    item.h = height;
  });
}

function createDefaultMindmap(page) {
  const images = page.objects.filter((object) => object.type === "image");
  const root = createTextObject("mind-root", "CENTRAL IDEA", 42, 45, 16, 14, {
    item: false, node: true, root: true, mindLevel: 1
  });
  const level2 = ["CONTEXT", "METHOD"].map((text) => createMindNode(text, root.id, 2));
  const level3 = [
    createMindNode("BACKGROUND", level2[0].id, 3),
    createMindNode("CHALLENGE", level2[0].id, 3),
    createMindNode("PROCESS", level2[1].id, 3),
    createMindNode("RESULT", level2[1].id, 3)
  ];
  const level4 = [
    createMindNode("DETAIL 1", level3[0].id, 4),
    createMindNode("DETAIL 2", level3[2].id, 4)
  ];
  page.objects = [root, ...level2, ...level3, ...level4, ...images];
  layoutMindmapTree(page);
}

function createMindNode(text, parentId, mindLevel) {
  return createTextObject("mind-node", text, 0, 0, 12, 10, {
    item: true, node: true, parentId, mindLevel
  });
}

function layoutMindmapTree(page) {
  const root = page.objects.find((object) => object.root);
  if (!root) return;
  Object.assign(root, { x: 42, y: 45, w: 16, h: 14, mindLevel: 1, parentId: null, mindAngle: 0 });
  const nodes = page.objects.filter((object) => object.role === "mind-node");
  const childrenByParent = new Map();
  nodes.forEach((node) => {
    const children = childrenByParent.get(node.parentId) || [];
    children.push(node);
    childrenByParent.set(node.parentId, children);
  });

  const specs = {
    2: { radiusX: 23, radiusY: 19, w: 17, h: 14 },
    3: { radiusX: 34, radiusY: 29, w: 14, h: 11 },
    4: { radiusX: 44, radiusY: 38, w: 11, h: 9 }
  };
  const centerX = 50;
  const centerY = 52;

  const positionChildren = (parent) => {
    const children = childrenByParent.get(parent.id) || [];
    children.forEach((node, index) => {
      const level = Math.min(4, parent.mindLevel + 1);
      const spec = specs[level];
      const angle = parent.root
        ? (children.length === 1 ? 0 : Math.PI * 2 * index / children.length)
        : parent.mindAngle + (index - (children.length - 1) / 2) * (level === 3 ? .9 : .55);
      node.mindLevel = level;
      node.mindAngle = angle;
      node.x = centerX + Math.cos(angle) * spec.radiusX - spec.w / 2;
      node.y = centerY + Math.sin(angle) * spec.radiusY - spec.h / 2;
      node.w = spec.w;
      node.h = spec.h;
      positionChildren(node);
    });
  };
  positionChildren(root);
}

function buildObjectTemplate(page, itemCount) {
  const images = page.objects.filter((object) => object.type === "image");
  page.objects = [createTextObject("page-title", getVariantTitle(page), 7, 7, 52, 16, { textAlign: "left" }), ...images];
  const count = Math.max(MIN_ITEMS, Math.min(MAX_ITEMS, itemCount));

  if (page.objectCategory === "layout") {
    if (page.variant === "cards") addCardLayout(page, count);
    if (page.variant === "table") addTableLayout(page, count);
    if (page.variant === "compare") addCompareLayout(page, Math.max(2, count));
  } else {
    if (page.variant === "process") addProcessDiagram(page, count);
    if (page.variant === "timeline") addTimelineDiagram(page, count);
    if (page.variant === "pyramid") addPyramidDiagram(page, count);
    if (page.variant === "cycle") addCycleDiagram(page, count);
  }
}

function getVariantTitle(page) {
  const collection = page.objectCategory === "layout" ? layouts : diagrams;
  return collection[page.variant]?.name?.toUpperCase() || "OBJECT PAGE";
}

function addCardLayout(page, count) {
  const columns = count <= 4 ? 2 : 3;
  const rows = Math.ceil(count / columns);
  const w = columns === 2 ? 37 : 25;
  const h = Math.min(24, 55 / rows);
  for (let i = 0; i < count; i += 1) {
    const column = i % columns;
    const row = Math.floor(i / columns);
    page.objects.push(createTextObject("card", `카드 ${i + 1}`, 9 + column * (w + 6), 32 + row * (h + 5), w, h, { item: true }));
  }
}

function addTableLayout(page, count) {
  page.objects.push({
    id: createId("table"), type: "table", role: "table", x: 12, y: 29, w: 76, h: 54, item: true,
    cells: [
      ["항목", "현재", "목표"],
      ...Array.from({ length: count }, (_, index) => [`항목 ${index + 1}`, "내용", "내용"])
    ]
  });
}

function addCompareLayout(page, count) {
  const visibleCount = Math.max(2, Math.min(4, count));
  const w = 78 / visibleCount;
  for (let i = 0; i < visibleCount; i += 1) {
    page.objects.push(createTextObject("compare-card", i === 0 ? "BEFORE" : i === 1 ? "AFTER" : `OPTION ${i + 1}`, 8 + i * (w + 3), 32, w, 48, { item: true }));
  }
}

function addProcessDiagram(page, count) {
  const w = Math.min(18, 75 / count);
  for (let i = 0; i < count; i += 1) {
    page.objects.push(createTextObject("diagram-node", `단계 ${i + 1}`, 8 + i * (84 / count), 43, w, 20, { item: true, node: true, sequence: i }));
  }
}

function addTimelineDiagram(page, count) {
  for (let i = 0; i < count; i += 1) {
    const x = 7 + i * (86 / count);
    const y = i % 2 === 0 ? 34 : 57;
    page.objects.push(createTextObject("timeline-node", `시점 ${i + 1}`, x, y, Math.min(18, 75 / count), 16, { item: true, node: true, sequence: i }));
  }
}

function addPyramidDiagram(page, count) {
  const height = Math.min(13, 58 / count);
  for (let i = 0; i < count; i += 1) {
    const w = 28 + (i * 7);
    page.objects.push(createTextObject("pyramid-level", `단계 ${i + 1}`, 50 - w / 2, 28 + i * (height + 2), w, height, { item: true }));
  }
}

function addCycleDiagram(page, count) {
  const centerX = 50;
  const centerY = 55;
  const radiusX = 31;
  const radiusY = 27;
  for (let i = 0; i < count; i += 1) {
    const angle = (-Math.PI / 2) + (Math.PI * 2 * i / count);
    page.objects.push(createTextObject("cycle-node", `단계 ${i + 1}`, centerX + Math.cos(angle) * radiusX - 8, centerY + Math.sin(angle) * radiusY - 7, 16, 14, { item: true, node: true, sequence: i }));
  }
}

function getItemCount(page) {
  if (page.template === "bullet") return page.objects.filter((object) => object.role === "bullet-item").length;
  if (page.template === "mindmap") return page.objects.filter((object) => object.role === "mind-node").length;
  if (page.template === "object") {
    const table = page.objects.find((object) => object.type === "table");
    if (table) return table.cells.length - 1;
    return page.objects.filter((object) => object.item && object.type !== "image").length;
  }
  return 0;
}

function getSelectedActionObject(page) {
  if (state.selectedIds.size !== 1) return null;
  const selected = page.objects.find((object) => state.selectedIds.has(object.id));
  if (!selected) return null;
  if (page.template === "mindmap") return selected.root || selected.role === "mind-node" ? selected : null;
  return selected.item ? selected : null;
}

function canAddItem(page, selected) {
  if (!selected || getItemCount(page) >= MAX_ITEMS) return false;
  if (page.template === "mindmap") return selected.mindLevel < 4;
  return true;
}

function canRemoveItem(page, selected) {
  if (!selected) return false;
  if (page.template === "mindmap") return !selected.root;
  const minimum = page.template === "object" && page.objectCategory === "layout" && page.variant === "compare" ? 2 : MIN_ITEMS;
  return getItemCount(page) > minimum;
}

function addItem() {
  const page = currentPage();
  if (page.type !== "content" || !page.template) return;
  const count = getItemCount(page);
  const selected = getSelectedActionObject(page);
  if (!canAddItem(page, selected)) return;
  snapshot();

  if (page.template === "bullet") {
    const item = createTextObject("bullet-item", `새 핵심 내용 ${count + 1}`, 12, 0, 75, 9, { item: true });
    page.objects.splice(page.objects.indexOf(selected) + 1, 0, item);
    layoutBulletItems(page);
    state.selectedIds = new Set([item.id]);
  } else if (page.template === "mindmap") {
    const childCount = page.objects.filter((object) => object.parentId === selected.id).length;
    const item = createMindNode(`하위 항목 ${childCount + 1}`, selected.id, selected.mindLevel + 1);
    page.objects.push(item);
    layoutMindmapTree(page);
    state.selectedIds = new Set([item.id]);
  } else {
    buildObjectTemplate(page, count + 1);
    state.selectedIds.clear();
  }
  hideTextToolbar();
  render();
}

function removeItem() {
  const page = currentPage();
  const count = getItemCount(page);
  const selected = getSelectedActionObject(page);
  if (page.type !== "content" || !canRemoveItem(page, selected)) return;
  snapshot();

  if (page.template === "bullet") {
    page.objects = page.objects.filter((object) => object.id !== selected.id);
    layoutBulletItems(page);
    state.selectedIds.clear();
  } else if (page.template === "mindmap") {
    const deletedIds = new Set([selected.id]);
    let foundChild = true;
    while (foundChild) {
      foundChild = false;
      page.objects.forEach((object) => {
        if (object.parentId && deletedIds.has(object.parentId) && !deletedIds.has(object.id)) {
          deletedIds.add(object.id);
          foundChild = true;
        }
      });
    }
    page.objects = page.objects.filter((object) => !deletedIds.has(object.id));
    layoutMindmapTree(page);
    state.selectedIds = selected.parentId ? new Set([selected.parentId]) : new Set();
  } else {
    buildObjectTemplate(page, count - 1);
    state.selectedIds.clear();
  }
  hideTextToolbar();
  render();
}

function render() {
  if (state.activeTextObjectId && !currentPage().objects.some((object) => object.id === state.activeTextObjectId)) hideTextToolbar();
  document.body.dataset.design = state.design;
  renderControls();
  renderPages();
  renderStage();
  updateUndoButton();
}

function renderControls() {
  const page = currentPage();
  const isCover = page.type === "cover";
  $("#designSelect").value = state.design;
  $("#designSelect").disabled = !isCover;
  $("#designHint").textContent = isCover ? (designs[state.design] || "") : "디자인은 PAGE 01에서만 변경할 수 있습니다.";
  $("#templateSection").hidden = isCover;
  $("#itemSection").hidden = false;
  $("#templateSelect").value = page.template || "";
  $("#objectOptions").hidden = page.template !== "object";
  $("#layoutVariantSelect").value = page.objectCategory === "layout" ? page.variant : "";
  $("#diagramVariantSelect").value = page.objectCategory === "diagram" ? page.variant : "";
  const itemCount = getItemCount(page);
  const selected = getSelectedActionObject(page);
  $("#addItemButton").disabled = isCover || !page.template || !canAddItem(page, selected);
  $("#removeItemButton").disabled = isCover || !page.template || !canRemoveItem(page, selected);
  $("#itemHint").textContent = isCover
    ? "표지 텍스트를 더블클릭해 수정하세요. 이미지도 여러 개 추가할 수 있습니다."
    : page.template === "mindmap"
      ? selected ? `현재 ${selected.mindLevel}단계 선택 · +는 하위 항목 추가, −는 선택 가지 삭제` : `현재 항목 ${itemCount}개 · 추가하거나 삭제할 개체를 먼저 선택하세요.`
      : page.template ? `현재 항목 ${itemCount}개 · 추가하거나 삭제할 항목 개체를 먼저 선택하세요.` : "본문 템플릿을 먼저 선택하세요.";
}

function renderPages() {
  const list = $("#pageList");
  list.innerHTML = "";
  state.pages.forEach((page, index) => {
    const button = document.createElement("button");
    button.className = `page-item ${index === state.currentPageIndex ? "is-current" : ""} ${page.type === "cover" ? "cover" : ""}`;
    button.type = "button";
    button.innerHTML = `PAGE ${String(index + 1).padStart(2, "0")}${page.type === "content" ? '<span class="page-delete">×</span>' : ""}`;
    button.addEventListener("click", (event) => {
      if (event.target.classList.contains("page-delete")) {
        snapshot();
        state.pages.splice(index, 1);
        state.currentPageIndex = Math.min(state.currentPageIndex, state.pages.length - 1);
      } else {
        state.currentPageIndex = index;
      }
      state.selectedIds.clear();
      state.guides = [];
      hideTextToolbar();
      render();
    });
    list.append(button);
  });
}

function renderStage() {
  stage.innerHTML = "";
  const page = currentPage();
  if (page.type === "content" && !page.template) {
    stage.innerHTML = '<div class="empty-page"><strong>본문 템플릿을<br>선택하세요.</strong><p>왼쪽 패널에서 시작합니다.</p></div>';
    return;
  }

  renderConnections(page);
  renderAlignmentGuides();
  page.objects.forEach((object) => stage.append(createObjectElement(object)));
  requestAnimationFrame(fitAllText);
}

function renderAlignmentGuides() {
  state.guides.forEach((guide) => {
    const element = document.createElement("div");
    element.className = `alignment-guide ${guide.axis === "x" ? "vertical" : "horizontal"}`;
    if (guide.axis === "x") element.style.left = `${guide.position}%`;
    else element.style.top = `${guide.position}%`;
    stage.append(element);
  });
}

function createObjectElement(object) {
  const element = document.createElement("div");
  element.className = `canvas-object ${getObjectClass(object)} ${state.selectedIds.has(object.id) ? "is-selected" : ""}`;
  element.dataset.objectId = object.id;
  applyObjectBox(element, object);

  if (object.type === "image") {
    const image = new Image();
    image.className = "object-image";
    image.src = object.src;
    image.alt = object.name || "첨부 이미지";
    element.append(image);
  } else if (object.type === "table") {
    element.append(createTableElement(object));
  } else {
    const text = document.createElement("div");
    text.className = `canvas-text ${object.role}`;
    text.textContent = object.text;
    applyTextObjectStyle(text, object, element);
    element.append(text);
    element.addEventListener("dblclick", (event) => beginTextEdit(event, object, element, text));
  }

  const handle = document.createElement("span");
  handle.className = "resize-handle";
  handle.addEventListener("pointerdown", (event) => beginResize(event, object));
  element.append(handle);
  element.addEventListener("pointerdown", (event) => beginDrag(event, object));
  return element;
}

function getTextAlign(object) {
  if (object.textAlign) return object.textAlign;
  if (object.role === "bullet-item") return "left";
  return "center";
}

function applyTextObjectStyle(text, object, wrapper) {
  const align = getTextAlign(object);
  const justify = { left: "flex-start", center: "center", right: "flex-end" }[align];
  text.style.textAlign = align;
  text.style.justifyContent = justify;
  text.style.color = object.textColor || "";
  if (object.fontSize) {
    text.style.setProperty("--object-font-size", `${object.fontSize}px`);
    wrapper.dataset.manualFontSize = "true";
  } else {
    delete wrapper.dataset.manualFontSize;
  }
}

function getObjectClass(object) {
  const classes = [object.role || ""];
  if (object.node) classes.push("node");
  if (object.color) classes.push(object.color);
  if (object.mindLevel) classes.push(`mind-level-${object.mindLevel}`);
  return classes.join(" ");
}

function applyObjectBox(element, object) {
  element.style.left = `${object.x}%`;
  element.style.top = `${object.y}%`;
  element.style.width = `${object.w}%`;
  element.style.height = `${object.h}%`;
}

function createTableElement(object) {
  const table = document.createElement("table");
  table.className = "table-object";
  object.cells.forEach((row, rowIndex) => {
    const tr = document.createElement("tr");
    row.forEach((value, columnIndex) => {
      const cell = document.createElement(rowIndex === 0 ? "th" : "td");
      cell.textContent = value;
      cell.addEventListener("dblclick", (event) => beginCellEdit(event, object, rowIndex, columnIndex, cell));
      tr.append(cell);
    });
    table.append(tr);
  });
  return table;
}

function beginTextEdit(event, object, wrapper, text) {
  if (document.fullscreenElement) return;
  event.preventDefault();
  event.stopPropagation();
  snapshot();
  state.activeTextObjectId = object.id;
  state.selectedIds.clear();
  state.selectedIds.add(object.id);
  renderControls();
  showTextToolbar(object, text);
  wrapper.classList.add("is-editing");
  text.contentEditable = "true";
  text.focus();
  const range = document.createRange();
  range.selectNodeContents(text);
  const selection = window.getSelection();
  selection.removeAllRanges();
  selection.addRange(range);

  const finish = () => {
    object.text = text.innerText.trim() || "텍스트";
    text.contentEditable = "false";
    wrapper.classList.remove("is-editing");
    renderStage();
  };
  text.addEventListener("blur", finish, { once: true });
  text.addEventListener("keydown", (keyEvent) => {
    if (keyEvent.key === "Escape") text.blur();
    if (keyEvent.key === "Enter" && !keyEvent.shiftKey) {
      keyEvent.preventDefault();
      text.blur();
    }
  });
}

function showTextToolbar(object, textElement) {
  const toolbar = $("#textToolbar");
  toolbar.hidden = false;
  $("#textColorInput").value = normalizeColor(object.textColor || getComputedStyle(textElement).color);
  $("#textSizeInput").value = Math.round(object.fontSize || Number.parseFloat(getComputedStyle(textElement).fontSize) || 28);
  toolbar.querySelectorAll("[data-text-align]").forEach((button) => {
    button.classList.toggle("is-active", button.dataset.textAlign === getTextAlign(object));
  });
}

function hideTextToolbar() {
  state.activeTextObjectId = null;
  $("#textToolbar").hidden = true;
}

function normalizeColor(value) {
  if (value.startsWith("#")) return value;
  const channels = value.match(/\d+/g)?.slice(0, 3).map(Number);
  if (!channels || channels.length < 3) return "#151515";
  return `#${channels.map((channel) => channel.toString(16).padStart(2, "0")).join("")}`;
}

function updateActiveTextStyle(property, value) {
  const page = currentPage();
  const object = page.objects.find((item) => item.id === state.activeTextObjectId && item.type === "text");
  if (!object) return;
  snapshot();
  const targets = property === "fontSize" && object.mindLevel
    ? page.objects.filter((item) => item.type === "text" && item.mindLevel === object.mindLevel)
    : [object];
  targets.forEach((target) => { target[property] = value; });
  renderStage();
  const activeText = stage.querySelector(`[data-object-id="${object.id}"] .canvas-text`);
  if (activeText) showTextToolbar(object, activeText);
}

function beginCellEdit(event, object, rowIndex, columnIndex, cell) {
  if (document.fullscreenElement) return;
  event.preventDefault();
  event.stopPropagation();
  snapshot();
  cell.contentEditable = "true";
  cell.focus();
  const finish = () => {
    object.cells[rowIndex][columnIndex] = cell.innerText.trim() || "내용";
    cell.contentEditable = "false";
  };
  cell.addEventListener("blur", finish, { once: true });
  cell.addEventListener("keydown", (keyEvent) => {
    if (keyEvent.key === "Enter") {
      keyEvent.preventDefault();
      cell.blur();
    }
  });
}

function beginDrag(event, object) {
  if (event.button !== 0 || event.target.classList.contains("resize-handle") || event.target.isContentEditable || document.fullscreenElement) return;
  event.preventDefault();
  hideTextToolbar();

  if (event.ctrlKey || event.metaKey) {
    if (state.selectedIds.has(object.id)) state.selectedIds.delete(object.id);
    else state.selectedIds.add(object.id);
    updateSelectionClasses();
    renderControls();
    return;
  }

  if (!state.selectedIds.has(object.id)) {
    state.selectedIds.clear();
    state.selectedIds.add(object.id);
  }
  updateSelectionClasses();
  renderControls();
  const start = getStagePoint(event);
  const selected = currentPage().objects.filter((item) => state.selectedIds.has(item.id));
  const origins = selected.map((item) => ({ item, x: item.x, y: item.y }));
  const bounds = getObjectBounds(selected);
  const targets = getSnapTargets(state.selectedIds);
  const thresholdX = SNAP_DISTANCE_PX / stage.clientWidth * 100;
  const thresholdY = SNAP_DISTANCE_PX / stage.clientHeight * 100;
  let dragStarted = false;

  const move = (moveEvent) => {
    const point = getStagePoint(moveEvent);
    const distanceX = (point.x - start.x) / 100 * stage.clientWidth;
    const distanceY = (point.y - start.y) / 100 * stage.clientHeight;
    if (!dragStarted) {
      if (Math.hypot(distanceX, distanceY) < 3) return;
      snapshot();
      dragStarted = true;
    }
    const rawDx = clamp(-bounds.left, point.x - start.x, 100 - bounds.right);
    const rawDy = clamp(-bounds.top, point.y - start.y, 100 - bounds.bottom);
    const horizontalSnap = findSnap(
      [bounds.left + rawDx, bounds.centerX + rawDx, bounds.right + rawDx], targets.x, thresholdX
    );
    const verticalSnap = findSnap(
      [bounds.top + rawDy, bounds.centerY + rawDy, bounds.bottom + rawDy], targets.y, thresholdY
    );
    const dx = clamp(-bounds.left, rawDx + (horizontalSnap?.correction || 0), 100 - bounds.right);
    const dy = clamp(-bounds.top, rawDy + (verticalSnap?.correction || 0), 100 - bounds.bottom);
    state.guides = [
      ...(horizontalSnap ? [{ axis: "x", position: horizontalSnap.target }] : []),
      ...(verticalSnap ? [{ axis: "y", position: verticalSnap.target }] : [])
    ];
    origins.forEach(({ item, x, y }) => {
      item.x = x + dx;
      item.y = y + dy;
    });
    renderStage();
  };
  const end = () => {
    window.removeEventListener("pointermove", move);
    window.removeEventListener("pointerup", end);
    if (dragStarted) {
      state.guides = [];
      renderStage();
    }
  };
  window.addEventListener("pointermove", move);
  window.addEventListener("pointerup", end);
}

function updateSelectionClasses() {
  stage.querySelectorAll(".canvas-object").forEach((element) => {
    element.classList.toggle("is-selected", state.selectedIds.has(element.dataset.objectId));
  });
}

function beginResize(event, object) {
  if (document.fullscreenElement) return;
  event.preventDefault();
  event.stopPropagation();
  hideTextToolbar();
  snapshot();
  const start = getStagePoint(event);
  const startWidth = object.w;
  const startHeight = object.h;
  const targets = getSnapTargets(new Set([object.id]));
  const thresholdX = SNAP_DISTANCE_PX / stage.clientWidth * 100;
  const thresholdY = SNAP_DISTANCE_PX / stage.clientHeight * 100;

  const move = (moveEvent) => {
    const point = getStagePoint(moveEvent);
    const rawRight = object.x + clamp(5, startWidth + point.x - start.x, 100 - object.x);
    const rawBottom = object.y + clamp(4, startHeight + point.y - start.y, 100 - object.y);
    const horizontalSnap = findSnap([rawRight], targets.x, thresholdX);
    const verticalSnap = findSnap([rawBottom], targets.y, thresholdY);
    const right = clamp(object.x + 5, rawRight + (horizontalSnap?.correction || 0), 100);
    const bottom = clamp(object.y + 4, rawBottom + (verticalSnap?.correction || 0), 100);
    object.w = right - object.x;
    object.h = bottom - object.y;
    state.guides = [
      ...(horizontalSnap ? [{ axis: "x", position: horizontalSnap.target }] : []),
      ...(verticalSnap ? [{ axis: "y", position: verticalSnap.target }] : [])
    ];
    renderStage();
  };
  const end = () => {
    window.removeEventListener("pointermove", move);
    window.removeEventListener("pointerup", end);
    state.guides = [];
    renderStage();
  };
  window.addEventListener("pointermove", move);
  window.addEventListener("pointerup", end);
}

function getObjectBounds(objects) {
  const left = Math.min(...objects.map((item) => item.x));
  const top = Math.min(...objects.map((item) => item.y));
  const right = Math.max(...objects.map((item) => item.x + item.w));
  const bottom = Math.max(...objects.map((item) => item.y + item.h));
  return { left, top, right, bottom, centerX: (left + right) / 2, centerY: (top + bottom) / 2 };
}

function getSnapTargets(excludedIds) {
  const x = new Set();
  const y = new Set();
  for (let position = 0; position <= 100; position += GRID_SIZE) {
    x.add(position);
    y.add(position);
  }
  currentPage().objects.filter((item) => !excludedIds.has(item.id)).forEach((item) => {
    x.add(item.x);
    x.add(item.x + item.w / 2);
    x.add(item.x + item.w);
    y.add(item.y);
    y.add(item.y + item.h / 2);
    y.add(item.y + item.h);
  });
  return { x: [...x], y: [...y] };
}

function findSnap(movingPoints, targets, threshold) {
  let best = null;
  movingPoints.forEach((point) => {
    targets.forEach((target) => {
      const correction = target - point;
      if (Math.abs(correction) <= threshold && (!best || Math.abs(correction) < Math.abs(best.correction))) {
        best = { correction, target };
      }
    });
  });
  return best;
}

function renderConnections(page) {
  if (page.template === "mindmap") {
    const root = page.objects.find((object) => object.root);
    page.objects.filter((object) => object.role === "mind-node").forEach((node) => {
      const parent = page.objects.find((object) => object.id === node.parentId) || root;
      drawConnection(parent, node);
    });
  }
  if (page.template === "object" && page.objectCategory === "diagram" && ["process", "timeline", "cycle"].includes(page.variant)) {
    const nodes = page.objects.filter((object) => object.node).sort((a, b) => a.sequence - b.sequence);
    nodes.forEach((node, index) => {
      const next = nodes[index + 1] || (page.variant === "cycle" ? nodes[0] : null);
      if (next) drawConnection(node, next);
    });
  }
}

function drawConnection(from, to) {
  if (!from || !to) return;
  const stageWidth = stage.clientWidth || 1600;
  const stageHeight = stage.clientHeight || 900;
  const fromX = from.x + from.w / 2;
  const fromY = from.y + from.h / 2;
  const dx = (to.x + to.w / 2 - fromX) / 100 * stageWidth;
  const dy = (to.y + to.h / 2 - fromY) / 100 * stageHeight;
  const line = document.createElement("div");
  line.className = "connection";
  line.style.left = `${fromX}%`;
  line.style.top = `${fromY}%`;
  line.style.width = `${Math.hypot(dx, dy) / stageWidth * 100}%`;
  line.style.transform = `rotate(${Math.atan2(dy, dx) * 180 / Math.PI}deg)`;
  stage.append(line);
}

function fitAllText() {
  const mindTexts = new Set(stage.querySelectorAll(".mind-root .canvas-text, .mind-node .canvas-text"));
  stage.querySelectorAll(".canvas-text").forEach((text) => {
    if (mindTexts.has(text)) return;
    if (text.closest(".canvas-object")?.dataset.manualFontSize) return;
    const rect = text.getBoundingClientRect();
    const lines = (text.textContent || "").split("\n");
    const longest = Math.max(...lines.map((line) => line.length), 1);
    const byWidth = (rect.width - 18) / longest * 1.55;
    const byHeight = (rect.height - 12) / (lines.length * 1.15);
    const fontSize = clamp(11, Math.min(byWidth, byHeight), 112);
    text.style.setProperty("--object-font-size", `${fontSize}px`);
  });
  fitMindmapTextByLevel(mindTexts);
  stage.querySelectorAll(".table-object").forEach((table) => {
    const rect = table.getBoundingClientRect();
    const rows = Math.max(table.rows.length, 1);
    table.style.setProperty("--table-font-size", `${clamp(10, rect.height / rows * .28, 30)}px`);
  });
}

function fitMindmapTextByLevel(mindTexts) {
  if (!mindTexts.size) return;
  const levelRatios = { 1: 1.65, 2: 1.32, 3: 1.08, 4: .88 };
  const grouped = { 1: [], 2: [], 3: [], 4: [] };
  mindTexts.forEach((text) => {
    const wrapper = text.closest(".canvas-object");
    const levelClass = [...wrapper.classList].find((name) => name.startsWith("mind-level-"));
    const level = Number(levelClass?.replace("mind-level-", ""));
    if (grouped[level]) grouped[level].push(text);
  });

  const stageScale = stage.clientWidth / 1200;
  let baseSize = 32 * stageScale;
  Object.entries(grouped).forEach(([level, texts]) => {
    if (!texts.length) return;
    if (texts[0].closest(".canvas-object")?.dataset.manualFontSize) return;
    const available = Math.min(...texts.map(getTextFitSize));
    baseSize = Math.min(baseSize, available / levelRatios[level]);
  });
  baseSize = Math.max(10 * stageScale, baseSize);
  Object.entries(grouped).forEach(([level, texts]) => {
    if (texts[0]?.closest(".canvas-object")?.dataset.manualFontSize) return;
    const size = baseSize * levelRatios[level];
    texts.forEach((text) => text.style.setProperty("--object-font-size", `${size}px`));
  });
}

function getTextFitSize(text) {
  const rect = text.getBoundingClientRect();
  const lines = (text.textContent || "").split("\n");
  const longest = Math.max(...lines.map((line) => line.length), 1);
  return Math.max(1, Math.min((rect.width - 12) / longest * 1.55, (rect.height - 10) / (lines.length * 1.15)));
}

function getStagePoint(event) {
  const rect = stage.getBoundingClientRect();
  return { x: (event.clientX - rect.left) / rect.width * 100, y: (event.clientY - rect.top) / rect.height * 100 };
}

function clamp(min, value, max) {
  return Math.max(min, Math.min(max, value));
}

function populateVariantSelects() {
  $("#layoutVariantSelect").innerHTML = '<option value="">선택 안 함</option>' + Object.entries(layouts).map(([value, item]) => `<option value="${value}">${item.name}</option>`).join("");
  $("#diagramVariantSelect").innerHTML = '<option value="">선택 안 함</option>' + Object.entries(diagrams).map(([value, item]) => `<option value="${value}">${item.name}</option>`).join("");
}

function serializeProject() {
  return JSON.stringify({
    format: PROJECT_FORMAT,
    version: PROJECT_VERSION,
    savedAt: new Date().toISOString(),
    presentation: {
      design: state.design,
      pages: state.pages,
      currentPageIndex: state.currentPageIndex
    }
  }, null, 2);
}

function parseProject(text) {
  const parsed = JSON.parse(text);
  if (parsed?.format !== PROJECT_FORMAT || parsed?.version !== PROJECT_VERSION) {
    throw new Error("Local ppt 저장 파일 형식이 아닙니다.");
  }
  const presentation = parsed.presentation;
  if (!presentation || !Array.isArray(presentation.pages) || !presentation.pages.length) {
    throw new Error("페이지 데이터가 없습니다.");
  }
  if (presentation.pages[0]?.type !== "cover" || presentation.pages[0]?.template !== "cover") {
    throw new Error("첫 페이지는 표지 페이지여야 합니다.");
  }
  presentation.pages.forEach((page, pageIndex) => {
    if (!page || !Array.isArray(page.objects)) throw new Error(`PAGE ${pageIndex + 1}의 개체 데이터가 올바르지 않습니다.`);
    page.objects.forEach((object) => {
      if (!object?.id || !object.type || ![object.x, object.y, object.w, object.h].every(Number.isFinite)) {
        throw new Error(`PAGE ${pageIndex + 1}에 잘못된 개체가 있습니다.`);
      }
    });
  });
  return {
    design: designs[presentation.design] ? presentation.design : "bauhaus",
    pages: JSON.parse(JSON.stringify(presentation.pages)),
    currentPageIndex: clamp(0, Number(presentation.currentPageIndex) || 0, presentation.pages.length - 1)
  };
}

async function loadProjectFile(file, handle = null) {
  try {
    const loaded = parseProject(await file.text());
    state.design = loaded.design;
    state.pages = loaded.pages;
    state.currentPageIndex = loaded.currentPageIndex;
    state.history = [];
    state.selectedIds.clear();
    state.guides = [];
    hideTextToolbar();
    currentProjectFileHandle = handle;
    currentProjectFileName = file.name || "local-ppt.txt";
    document.title = `Local ppt — ${currentProjectFileName}`;
    render();
  } catch (error) {
    window.alert(`파일을 불러올 수 없습니다.\n${error.message}`);
  }
}

async function writeProjectToHandle(handle) {
  const writable = await handle.createWritable();
  await writable.write(serializeProject());
  await writable.close();
  currentProjectFileHandle = handle;
  currentProjectFileName = handle.name || currentProjectFileName;
  document.title = `Local ppt — ${currentProjectFileName}`;
}

function downloadProject(filename = currentProjectFileName) {
  const blob = new Blob([serializeProject()], { type: "text/plain;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = filename.endsWith(".txt") ? filename : `${filename}.txt`;
  document.body.append(anchor);
  anchor.click();
  anchor.remove();
  URL.revokeObjectURL(url);
  currentProjectFileName = anchor.download;
  document.title = `Local ppt — ${currentProjectFileName}`;
}

function suggestedProjectName() {
  const timestamp = new Date().toISOString().slice(0, 16).replace(/[T:]/g, "-");
  return `local-ppt-${timestamp}.txt`;
}

async function saveProjectAs() {
  try {
    if (window.showSaveFilePicker) {
      const handle = await window.showSaveFilePicker({
        id: PROJECT_PICKER_ID,
        suggestedName: suggestedProjectName(),
        types: [{ description: "Local ppt JSON 텍스트", accept: { "text/plain": [".txt"] } }]
      });
      await writeProjectToHandle(handle);
    } else {
      downloadProject(suggestedProjectName());
    }
  } catch (error) {
    if (error.name === "AbortError") return;
    if (["SecurityError", "NotSupportedError"].includes(error.name)) downloadProject(suggestedProjectName());
    else window.alert(`파일을 저장할 수 없습니다.\n${error.message}`);
  }
}

async function saveCurrentProject() {
  try {
    if (currentProjectFileHandle) await writeProjectToHandle(currentProjectFileHandle);
    else if (currentProjectFileName !== "local-ppt.txt") downloadProject(currentProjectFileName);
    else await saveProjectAs();
  } catch (error) {
    if (error.name !== "AbortError") window.alert(`파일을 저장할 수 없습니다.\n${error.message}`);
  }
}

$("#loadProjectButton").addEventListener("click", async () => {
  if (window.showOpenFilePicker) {
    try {
      const [handle] = await window.showOpenFilePicker({
        id: PROJECT_PICKER_ID,
        multiple: false,
        types: [{ description: "Local ppt JSON 텍스트", accept: { "text/plain": [".txt", ".json"] } }]
      });
      await loadProjectFile(await handle.getFile(), handle);
    } catch (error) {
      if (error.name === "AbortError") return;
      if (["SecurityError", "NotSupportedError"].includes(error.name)) {
        $("#projectFileInput").value = "";
        $("#projectFileInput").click();
      } else {
        window.alert(`파일을 불러올 수 없습니다.\n${error.message}`);
      }
    }
  } else {
    $("#projectFileInput").value = "";
    $("#projectFileInput").click();
  }
});

$("#projectFileInput").addEventListener("change", (event) => {
  const [file] = event.target.files;
  if (file) loadProjectFile(file);
});

$("#saveProjectButton").addEventListener("click", saveCurrentProject);
$("#saveAsProjectButton").addEventListener("click", saveProjectAs);

$("#designSelect").addEventListener("change", (event) => {
  if (state.currentPageIndex !== 0) return;
  snapshot();
  state.design = event.target.value;
  render();
});

$("#templateSelect").addEventListener("change", (event) => {
  const page = currentPage();
  if (page.type !== "content" || !event.target.value) return;
  snapshot();
  buildTemplate(page, event.target.value);
  state.selectedIds.clear();
  hideTextToolbar();
  render();
});

$("#layoutVariantSelect").addEventListener("change", (event) => {
  if (!event.target.value) return;
  const page = currentPage();
  snapshot();
  page.objectCategory = "layout";
  page.variant = event.target.value;
  buildObjectTemplate(page, page.variant === "compare" ? 2 : 3);
  hideTextToolbar();
  render();
});

$("#diagramVariantSelect").addEventListener("change", (event) => {
  if (!event.target.value) return;
  const page = currentPage();
  snapshot();
  page.objectCategory = "diagram";
  page.variant = event.target.value;
  buildObjectTemplate(page, 3);
  hideTextToolbar();
  render();
});

$("#addPageButton").addEventListener("click", () => {
  snapshot();
  state.pages.push(createContentPage());
  state.currentPageIndex = state.pages.length - 1;
  state.selectedIds.clear();
  hideTextToolbar();
  render();
});

$("#addItemButton").addEventListener("click", addItem);
$("#removeItemButton").addEventListener("click", removeItem);
$("#undoButton").addEventListener("click", undo);

$("#textColorInput").addEventListener("change", (event) => updateActiveTextStyle("textColor", event.target.value));
$("#textSizeInput").addEventListener("input", (event) => {
  const size = clamp(8, Number(event.target.value) || 8, 160);
  event.target.value = size;
  updateActiveTextStyle("fontSize", size);
});
document.querySelectorAll("[data-text-align]").forEach((button) => {
  button.addEventListener("click", () => updateActiveTextStyle("textAlign", button.dataset.textAlign));
});
$("#autoTextSizeButton").addEventListener("click", () => updateActiveTextStyle("fontSize", null));
$("#closeTextToolbarButton").addEventListener("click", hideTextToolbar);

$("#imageInput").addEventListener("change", (event) => {
  const files = [...event.target.files];
  if (!files.length) return;
  snapshot();
  const targetPage = currentPage();
  let completed = 0;
  files.forEach((file, index) => {
    const reader = new FileReader();
    reader.addEventListener("load", () => {
      targetPage.objects.push({
        id: createId("image"), type: "image", role: "image", src: String(reader.result), name: file.name,
        x: 62 + (index % 3) * 5, y: 55 + (index % 3) * 5, w: 25, h: 28
      });
      completed += 1;
      if (completed === files.length) {
        event.target.value = "";
        render();
      }
    });
    reader.readAsDataURL(file);
  });
});

$("#fullscreenButton").addEventListener("click", () => {
  if (document.fullscreenElement) document.exitFullscreen();
  else stage.requestFullscreen();
});

document.addEventListener("keydown", (event) => {
  const activeTag = document.activeElement?.tagName;
  if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "z") {
    event.preventDefault();
    undo();
    return;
  }
  if (event.key.toLowerCase() === "f" && !["INPUT", "TEXTAREA", "SELECT"].includes(activeTag) && !document.activeElement?.isContentEditable) {
    event.preventDefault();
    $("#fullscreenButton").click();
  }
  if (event.key === "Escape") {
    state.selectedIds.clear();
    state.guides = [];
    hideTextToolbar();
    renderStage();
  }
});

window.addEventListener("resize", () => requestAnimationFrame(fitAllText));
populateVariantSelects();
render();
