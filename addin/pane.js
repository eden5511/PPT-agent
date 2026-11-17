"use strict";
const el = id => document.getElementById(id);
let current = null;
let reading = false;
let busy = false;
function message(text, user = false) {
  const p = document.createElement("p");
  p.textContent = text;
  if (user) p.className = "user";
  el("log").append(p);
  p.scrollIntoView({block:"nearest"});
}
function check(ok, why) { if (!ok) throw new Error(why); }
function parse(text) {
  const page = "(?:这页|本页|第\\s*(\\d+)\\s*页)";
  const title = text.match(/^(?:把|将)?(?:这个|选中的)标题(?:改为|改成)(.+)$/u);
  if (title) return {type:"selected", value:title[1].trim()};
  const cell = text.match(new RegExp("^(?:把|将)?" + page + "的?第(\\d+)行第(\\d+)列改[成为](.+)$"));
  if (cell) return {type:"cell", page:cell[1], row:Number(cell[2])-1, col:Number(cell[3])-1, value:cell[4]};
  const named = text.match(new RegExp('^(?:把|将)?' + page + '[“"](.+?)[”"]这行第(\\d+)列改[成为](.+)$'));
  if (named) return {type:"cell", page:named[1], label:named[2], col:Number(named[3])-1, value:named[4]};
  const red = text.match(new RegExp("^(?:把|将)?" + page + "(?:的ppt)?的?行标题(?:全部)?标红$","i"));
  if (red) return {type:"rowHeaders", page:red[1]};
  const rearrange = text.match(new RegExp("^(区域.+)，填入" + page + "新表格$"));
  if (rearrange) return {type:"transform", transform:rearrange[1], page:rearrange[2]};
  const fill = text.match(new RegExp("^交叉表补空\\s*" + page + "$"));
  if (fill) return {type:"fill", page:fill[1]};
  throw new Error("可输入“把这个标题改成品牌考虑度”“把第5页第3行第2列改为待复核”或“交叉表补空 第17页”。");
}
function target() {
  el("target").textContent = "";
  if (!current) return;
  try {
    const action = parse(el("command").value.trim());
    const page = action.page || current.page;
    const detail = action.type === "selected" ? current.shapeName || "请选中一个文本框" :
      action.type === "cell" ? (action.label || ("第 " + (action.row+1) + " 行")) + " 第 " + (action.col+1) + " 列" : action.type === "rowHeaders" ? "第一列行标签" : "交叉表";
    el("target").textContent = "本次修改 · 第 " + page + " 页 · " + detail;
  } catch { /* Partial instructions have no established target yet. */ }
}
async function selection(ctx) {
  const pres = ctx.presentation;
  const slides = pres.slides.load("items/id");
  const selected = pres.getSelectedSlides().load("items/id");
  const shapes = pres.getSelectedShapes().load("items/id,items/name");
  await ctx.sync();
  check(selected.items.length === 1, "请选中一张幻灯片");
  const id = selected.items[0].id;
  const page = slides.items.findIndex(s => s.id === id) + 1;
  check(page > 0, "请重新选择幻灯片");
  const one = shapes.items.length === 1 ? shapes.items[0] : null;
  return {id, page, total:slides.items.length, count:shapes.items.length, shapeId:one?.id, shapeName:one?.name};
}
async function refresh() {
  if (reading || busy) return;
  reading = true;
  try {
    current = await PowerPoint.run(selection);
    const suffix = current.count > 1 ? "已选 " + current.count + " 个对象" : current.shapeName;
    el("selection").textContent = "Slide " + current.page + " of " + current.total + (suffix ? " · " + suffix : "");
  } catch (err) { current = null; el("selection").textContent = err.message; }
  finally { reading = false; target(); }
}
function readFile(file) {
  return new Promise((resolve,reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result.split(",")[1]);
    reader.onerror = () => reject(reader.error);
    reader.readAsDataURL(file);
  });
}
function readAliases() {
  const aliases = Object.create(null);
  for (const line of el("aliases").value.split(/\r?\n/).filter(x=>x.trim())) {
    const parts = line.split("=");
    check(parts.length === 2, "别名示例：Hair=洗发护发");
    aliases[parts[0].trim()] = parts[1].split(/[、,，]/).map(x=>x.trim());
  }
  return aliases;
}
async function getPlan(grid, transform) {
  const file = el("file").files[0];
  check(file, "请在“Excel 交叉表”中选择工作簿");
  const aliases = readAliases();
  const response = await fetch("/plan", {method:"POST", headers:{"Content-Type":"application/json"},
    body:JSON.stringify({file:await readFile(file), sheet:el("sheet").value.trim(), prefix:el("prefix").value.trim(), aliases, grid, transform})});
  if (response.status === 413) throw new Error("工作簿过大，请另存所需工作表后重试");
  const plan = await response.json();
  check(response.ok && !plan.error, plan.error || "读取 Excel 失败");
  return plan;
}
function chooseSource(choices) {
  message("多组数据都能完整匹配，请选择本次使用的来源。");
  const group = document.createElement("div");
  el("log").append(group);
  return new Promise(resolve => {
    for (const choice of [...choices, null]) {
      const button = document.createElement("button");
      button.type = "button";
      button.style.float = "none";
      button.textContent = choice ? choice.sheet + " · " + choice.prefix : "取消任务";
      button.onclick = () => { group.remove(); resolve(choice); };
      group.append(button);
    }
    group.scrollIntoView({block:"nearest"});
  });
}

async function postLocal(path, payload) {
  const response = await fetch(path,{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify(payload)});
  if (response.status === 413) throw new Error("工作簿过大，请另存所需工作表后重试");
  const result = await response.json();
  check(response.ok && !result.error,result.error || "本地服务请求失败");
  return result;
}

async function consult(ctx, text) {
  const file = el("file").files[0];
  check(file, "请先在 Excel 交叉表中选择 .xlsx 文件");
  const selected = await selection(ctx);
  const pages = [...new Set(Array.from(text.matchAll(/第\s*(\d+)\s*页/g), m=>Number(m[1])))];
  check(pages.length <= 1, "当前补表任务一次处理一页，请指定目标页");
  const page = pages[0] ?? selected.page;
  check(page >= 1 && page <= selected.total,"目标页超出范围");
  check(Office.context.requirements.isSetSupported("PowerPointApi","1.8"), "需要支持 PowerPointApi 1.8 的桌面版");
  const slide = ctx.presentation.slides.getItemAt(page-1);
  slide.load("id");
  const shapes = slide.shapes.load("items/id,items/type");
  await ctx.sync();
  let tables = shapes.items.filter(shape=>shape.type === PowerPoint.ShapeType.table);
  if (page === selected.page && selected.count) tables = tables.filter(shape=>shape.id === selected.shapeId);
  check(tables.length === 1, "请点选唯一的目标表格，或点击目标页缩略图后重试");
  const slideId = slide.id, shapeId = tables[0].id;
  const {grid} = await snapshot(ctx,tables[0].getTable());
  const payload = {file:await readFile(file), grid, aliases:readAliases()};
  message("已锁定第 " + page + " 页的表格，正在搜索工作簿并核对来源。");
  const result = await postLocal("/agent", {...payload,text,page});
  if (result.trace?.length) message(result.trace.join(" → "));
  if (result.notice) return result.notice;
  let source = result.source, plan = result.plan;
  if (result.choices) {
    source = await chooseSource(result.choices);
    if (!source) return "已取消，本次未写入。";
    plan = await postLocal("/plan",{...payload,sheet:source.sheet,prefix:source.prefix});
  }
  const table = ctx.presentation.slides.getItem(slideId).shapes.getItem(shapeId).getTable();
  const fresh = await snapshot(ctx, table);
  check(JSON.stringify(fresh.grid) === JSON.stringify(grid),"取数期间目标表格已变化，请重新发送");
  for (const e of plan.edits) {
    check(e.row > 0 && e.col > 0 && fresh.grid[e.row]?.[e.col]?.trim() === "" && typeof e.value === "string" && e.source,
      "计划包含无效位置或缺失来源");
  }
  for (const e of plan.edits) fresh.cells[e.row][e.col].text = e.value;
  await ctx.sync();
  const actual = await snapshot(ctx,table);
  const expected = grid.map(row=>row.slice());
  for (const e of plan.edits) expected[e.row][e.col] = e.value;
  check(JSON.stringify(actual.grid) === JSON.stringify(expected),"写后复核不一致，请检查目标表");
  const report = ["来源 " + source.sheet + " · " + source.prefix,
    "第 " + page + " 页已复核，补空 " + plan.edits.length + " 格，一致 " + plan.same + " 格，已有差异 " + plan.conflicts.length + " 格。"];
  for (const e of [...plan.edits,...plan.conflicts]) report.push(grid[e.row][0] + " / " + grid[0][e.col] + "：" +
    (e.old ? "保留 " + e.old + "，来源 " : "已填 ") + e.value + "（" + e.source + "）");
  return report.join("\n");
}

async function snapshot(ctx, table) {
  table.load("rowCount,columnCount");
  await ctx.sync();
  const cells = Array.from({length:table.rowCount}, (_,r) => Array.from({length:table.columnCount},
    (_,c) => table.getCellOrNullObject(r,c).load("text")));
  await ctx.sync();
  check(cells.every(row=>row.every(cell=>!cell.isNullObject)), "目标包含无法定位的单元格");
  return {cells, grid:cells.map(row=>row.map(cell=>cell.text))};
}
async function execute(ctx, action) {
  const selected = await selection(ctx);
  const pres = ctx.presentation;
  const page = action.page === undefined ? selected.page : Number(action.page);
  check(page >= 1 && page <= selected.total, "目标页超出范围");
  const slide = pres.slides.getItemAt(page-1);
  if (action.type === "selected") {
    check(selected.count === 1, "请点击要修改的一个标题框");
    const range = slide.shapes.getItem(selected.shapeId).textFrame.textRange;
    range.load("text");
    await ctx.sync();
    const before = range.text;
    range.text = action.value;
    await ctx.sync();
    range.load("text"); await ctx.sync();
    check(range.text === action.value, "写入后核对失败，请查看页面");
    return selected.shapeName + " · " + before + " → " + range.text;
  }
  check(Office.context.requirements.isSetSupported("PowerPointApi","1.8"), "当前 PowerPoint 版本不支持原生表格编辑（需要 PowerPointApi 1.8）");
  if (action.type === "transform") {
    const {matrix} = await getPlan(null, action.transform);
    const shape = slide.shapes.addTable(matrix.length, matrix[0].length, {values:matrix});
    await ctx.sync();
    const actual = await snapshot(ctx,shape.getTable());
    check(JSON.stringify(actual.grid) === JSON.stringify(matrix), "新表已插入，写入后核对失败");
    return "第 " + page + " 页已插入 " + matrix.length + " 行 × " + matrix[0].length + " 列的新表格";
  }
  const shapes = slide.shapes.load("items/id,items/type");
  await ctx.sync();
  let tables = shapes.items.filter(s=>s.type === PowerPoint.ShapeType.table);
  if (selected.page === page && selected.count) {
    check(selected.count === 1 && tables.some(t=>t.id === selected.shapeId), "请选中目标表格，或点击幻灯片缩略图取消对象选择");
    tables = tables.filter(t=>t.id === selected.shapeId);
  }
  check(tables.length === 1, "该页需要有一张可唯一确定的表格，多表时请点选目标表格");
  const table = tables[0].getTable();
  const {cells,grid} = await snapshot(ctx, table);
  if (action.type === "rowHeaders") {
    check(Office.context.requirements.isSetSupported("PowerPointApi","1.9"), "行标签改色需要 PowerPointApi 1.9");
    const labels = cells.slice(1).map(row=>row[0]).filter(cell=>cell.text.trim());
    for (const cell of labels) cell.font.color = "#FF0000";
    await ctx.sync();
    for (const cell of labels) cell.font.load("color");
    await ctx.sync();
    check(labels.every(cell=>cell.font.color.toUpperCase() === "#FF0000"), "颜色写后核对失败");
    return "第 " + page + " 页已标红 " + labels.length + " 个行标签";
  }
  let edits, report = "";
  if (action.type === "cell") {
    if (action.label) {
      const hits = grid.map((r,i)=>i > 0 && r[0].trim() === action.label ? i : -1).filter(i=>i>=0);
      check(hits.length === 1, "行标签缺失或重复");
      action.row = hits[0];
    }
    check(cells[action.row]?.[action.col], "行列编号超出表格范围");
    edits = [action];
  } else {
    const plan = await getPlan(grid);
    const fresh = await snapshot(ctx, table);
    check(JSON.stringify(fresh.grid) === JSON.stringify(grid), "取数期间表格已变化，请重新发送");
    edits = plan.edits;
    report = "一致 " + plan.same + " 格，已有差异 " + plan.conflicts.length + " 格。";
    for (const e of plan.conflicts) report += "\n" + grid[e.row][0] + " / " + grid[0][e.col] + "：PPT " + e.old + "，Excel " + e.value + "（" + e.source + "）";
  }
  for (const e of edits) cells[e.row][e.col].text = e.value;
  await ctx.sync();
  for (const e of edits) cells[e.row][e.col].load("text");
  await ctx.sync();
  check(edits.every(e=>cells[e.row][e.col].text === e.value), "写后核对失败，请检查目标表格");
  return "第 " + page + " 页已核对写入 " + edits.length + " 格。" + report;
}
Office.onReady(async info => {
  if (info.host !== Office.HostType.PowerPoint || !Office.context.requirements.isSetSupported("PowerPointApi","1.5")) {
    el("selection").textContent = "请在支持 PowerPointApi 1.5 的 PowerPoint 中打开"; return;
  }
  await refresh();
  Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, refresh);
  setInterval(refresh, 1500);
  el("command").addEventListener("input",target);
  el("send").disabled = false;
  document.querySelector("form").onsubmit = async event => {
    event.preventDefault();
    if (busy) return;
    busy = true; el("send").disabled = true;
    const text = el("command").value.trim(); message(text,true);
    try {
      if (el("agent-mode").checked) {
        message(await PowerPoint.run(ctx=>consult(ctx,text)));
        el("command").value = "";
        return;
      }
      if (/^DeepSeek\s+/i.test(text)) {
        const response = await fetch("/intent", {method:"POST", headers:{"Content-Type":"application/json"},
          body:JSON.stringify({text:text.replace(/^DeepSeek\s+/i,"")})});
        const result = await response.json();
        check(response.ok && result.command, result.error || "当前工具无法完整执行这项需求");
        parse(result.command);
        el("command").value = result.command;
        message("已生成指令，请核对输入框中的目标和内容。");
        return;
      }
      const action = parse(text);
      message(await PowerPoint.run(ctx=>execute(ctx,action)));
      el("command").value = "";
    } catch (err) { message(err.message || String(err)); }
    finally { busy = false; el("send").disabled = false; refresh(); }
  };
});
