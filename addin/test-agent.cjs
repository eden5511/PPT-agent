const test = require("node:test");
const assert = require("node:assert/strict");
const {runAgent} = require("./agent");

const payload = {file:"local-xlsx",page:17,grid:[["","Hair"],["Base",""],["行为","confidential-value"]],text:"用 Excel 补齐第17页"};
const plan = {edits:[{row:1,col:1,value:"100",source:"本期!B2",old:""}],same:0,conflicts:[]};
function fixture(candidates) {
  const calls = [["inspect_target",{}],["find_sources",{}],["check_source",{source:1}],["finish_fill",{source:1}]];
  let round = 0;
  return runAgent(payload,async args => args.operation ? {candidates} : plan, async messages => {
    assert.doesNotMatch(JSON.stringify(messages),/confidential-value|local-xlsx/);
    if (round) assert.equal(messages.at(-1).role,"tool");
    const [name,args] = calls[round++];
    return {tool_calls:[{id:String(round),type:"function",function:{name,arguments:JSON.stringify(args)}}]};
  });
}
test("model receives tool results; local source alone supplies edits",async()=>{
  const result = await fixture([{sheet:"本期",prefix:"Q"}]);
  assert.deepEqual(result.plan,plan);
});
test("even when the model picks one, equally valid sources require clarification",async()=>{
  const candidates = [{sheet:"本期",prefix:"Q"},{sheet:"上期",prefix:"Q"}];
  const result = await fixture(candidates);
  assert.deepEqual(result.choices,candidates);
  assert.equal(result.plan,undefined);
});
