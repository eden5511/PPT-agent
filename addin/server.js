const https = require("node:https");
const fs = require("node:fs");
const path = require("node:path");
const { spawn } = require("node:child_process");
const certs = require("office-addin-dev-certs");
const origin = "https://localhost:3000";

function workbook(data) {
  return new Promise((resolve, reject) => {
    const worker = spawn(process.platform === "win32" ? "py" : "python3",
      ["-X", "utf8", path.join(__dirname, "workbook.py")], { windowsHide: true });
    let output = "";
    worker.stdout.on("data", chunk => output += chunk);
    worker.stderr.resume();
    worker.stdin.on("error", reject);
    worker.on("error", reject);
    worker.on("close", () => {
      try { resolve(JSON.parse(output)); }
      catch { reject(new Error("Excel 读取失败，请检查 Python 和 openpyxl 安装。")); }
    });
    worker.stdin.end(data);
  });
}

async function intent({text}) {
  if (!process.env.DEEPSEEK_API_KEY) return {error:"请先在 Windows 环境变量中配置 DEEPSEEK_API_KEY，然后重新启动加载项服务"};
  if (typeof text !== "string" || !text.trim() || text.length > 8000) return {error:"指令为空或过长"};
  const response = await fetch("https://api.deepseek.com/chat/completions", {
    method:"POST", signal:AbortSignal.timeout(60000),
    headers:{Authorization:"Bearer " + process.env.DEEPSEEK_API_KEY, "Content-Type":"application/json"},
    body:JSON.stringify({model:"deepseek-v4-pro",response_format:{type:"json_object"},messages:[
      {role:"system",content:'返回 JSON 对象 {"command":"..."}。只可生成一条完整表达原需求的指令。保留原页码、行列和新值。支持：把这个标题改成X，把第N页第R行第C列改为X，把第N页“标签”这行第C列改为X，把第N页的行标题全部标红，交叉表补空 第N页。交叉表的数据源由用户在输入区选择。仅明确当前页时使用这页。不支持、缺参数、需多条或语义有损时返回 {"command":null}。'},
      {role:"user",content:text}
    ]})
  });
  if (!response.ok) return {error:"DeepSeek 请求失败（HTTP " + response.status + "）"};
  const payload = await response.json();
  return JSON.parse(payload.choices[0].message.content);
}

async function serve(req, res) {
  try {
    if (req.headers.host !== "localhost:3000") { res.writeHead(403).end(); return; }
    if (req.method === "POST" && ["/plan", "/intent"].includes(req.url)) {
      if (req.headers.origin !== origin || req.headers["content-type"] !== "application/json") {
        res.writeHead(403).end(); return;
      }
      const chunks = [];
      let bytes = 0;
      for await (const chunk of req) {
        bytes += chunk.length;
        if (bytes > 48 * 1024 * 1024) { res.writeHead(413).end(); return; }
        chunks.push(chunk);
      }
      const body = Buffer.concat(chunks);
      const result = req.url === "/intent" ? await intent(JSON.parse(body)) : await workbook(body);
      res.writeHead(result.error ? 400 : 200, {"Content-Type":"application/json"}).end(JSON.stringify(result));
      return;
    }
    const types = {"/pane.html":"text/html", "/pane.js":"text/javascript", "/pane.css":"text/css",
      "/icon16.png":"image/png", "/icon32.png":"image/png", "/icon80.png":"image/png"};
    if (req.method !== "GET" || !Object.hasOwn(types, req.url)) { res.writeHead(404).end(); return; }
    res.writeHead(200, {"Content-Type": types[req.url], "Cache-Control":"no-store"});
    res.end(fs.readFileSync(path.join(__dirname, req.url.slice(1))));
  } catch (error) {
    if (!res.headersSent) res.writeHead(500, {"Content-Type":"application/json"});
    res.end(JSON.stringify({ error: error.message }));
  }
}
certs.getHttpsServerOptions().then(options => https.createServer(options, serve).listen(3000, "localhost"));
