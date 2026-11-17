const {test} = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const vm = require("node:vm");
const env = {Office:{onReady(){},context:{requirements:{isSetSupported:()=>true}}},PowerPoint:{ShapeType:{table:"Table"}}};
vm.createContext(env);
vm.runInContext(fs.readFileSync(__dirname + "/pane.js", "utf8"), env);
function loaded(object) { object.load = () => object; return object; }
function fixture() {
  const range = loaded({text:"原标题"});
  const shape = {id:"9", name:"Title 9", textFrame:{textRange:range}};
  const cells = [["", "Skin"],["Base","100"]].map(r=>r.map(text=>loaded({text})));
  const table = loaded({rowCount:2,columnCount:2,getCellOrNullObject:(r,c)=>cells[r][c]});
  const slides = Array.from({length:21}, (_,i)=>({id:String(i+1)}));
  slides[4].shapes = {getItem:id=>{assert.equal(id,"9"); return shape;}};
  slides[15].shapes = loaded({items:[{type:"Table",id:"3",getTable:()=>table}]});
  const ctx = {sync:async()=>{}, presentation:{
    slides:loaded({items:slides,getItemAt:i=>slides[i]}),
    getSelectedSlides:()=>loaded({items:[slides[4]]}),
    getSelectedShapes:()=>loaded({items:[shape]})
  }};
  return {range,cells,ctx};
}
test("selected shape is the sole text target", async()=>{
  const {ctx,range} = fixture();
  const selected = await env.selection(ctx);
  assert.deepEqual([selected.page,selected.total,selected.shapeName],[5,21,"Title 9"]);
  await env.execute(ctx,env.parse("把这个标题改成品牌考虑度"));
  assert.equal(range.text,"品牌考虑度");
});
test("explicit page and row label take precedence over current slide", async()=>{
  const {ctx,cells,range} = fixture();
  await env.execute(ctx,env.parse("把第16页“Base”这行第2列改为90"));
  assert.equal(cells[1][1].text,"90");
  assert.equal(range.text,"原标题");
  await assert.rejects(env.execute(ctx,env.parse("把第16页第0行第2列改为90")));
});

function consultFixture(changeValue) {
  const {ctx,cells,range} = fixture();
  loaded(ctx.presentation.slides.items[15]);
  ctx.presentation.slides.getItem = id => {
    assert.equal(id,"16");
    return ctx.presentation.slides.items[15];
  };
  const shapes = ctx.presentation.slides.items[15].shapes;
  shapes.getItem = id=>{assert.equal(id,"3");return shapes.items[0];};
  cells[1][1].text = "";
  env.FileReader = class {readAsDataURL(){this.result="data:application/octet-stream;base64,WA==";this.onload();}};
  env.document = {getElementById(id){if(id === "file")return {files:[{}]};
    if(id === "aliases")return {value:""}; return {append(){}};},
    createElement(){return {scrollIntoView(){}};}};
  env.fetch = async (_url, options)=>{
    const input = JSON.parse(options.body);
    assert.equal(input.page,16);
    ctx.presentation.getSelectedSlides = ()=>loaded({items:[ctx.presentation.slides.items[7]]});
    if(changeValue)cells[1][1].text="手动修改";
    return {ok:true,json:async()=>({source:{sheet:"本期",prefix:"Q"},plan:{same:0,conflicts:[],
      edits:[{row:1,col:1,old:"",value:"100",source:"本期!B2"}]}})};
  };
  return {ctx,cells,range};
}
test("consultation pins target and reports verified source after user changes slides", async()=>{
  const {ctx,cells,range} = consultFixture(false);
  assert.match(await env.consult(ctx,"用 Excel 补齐第16页"),/本期!B2/);
  assert.equal(cells[1][1].text,"100");
  assert.equal(range.text,"原标题");
});
test("concurrent manual edit survives and the outdated plan stops",async()=>{
  const {ctx,cells} = consultFixture(true);
  await assert.rejects(env.consult(ctx,"用 Excel 补齐第16页"),/已变化/);
  assert.equal(cells[1][1].text,"手动修改");
});
