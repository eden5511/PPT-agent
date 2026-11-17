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
