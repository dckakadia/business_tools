// SpaTrack Pro — Server
import http from "http";
import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const PORT = 8081;
const DATA_FILE = path.join(__dirname, "data.json");

function readBody(req) {
  return new Promise((res,rej)=>{
    let b=""; req.on("data",c=>{b+=c;}); req.on("end",()=>{try{res(JSON.parse(b));}catch(e){rej(e);}});
  });
}

function readData() {
  try {
    return JSON.parse(fs.readFileSync(DATA_FILE,"utf8"));
  } catch(e) { return {items:[],products:[],purchaseOrders:[]}; }
}

function writeData(data) {
  try {
    const tmp = DATA_FILE+".tmp";
    fs.writeFileSync(tmp, JSON.stringify(data,null,2));
    fs.renameSync(tmp, DATA_FILE);
    return true;
  } catch(e) { return false; }
}

const mimeTypes = {
  '.html': 'text/html',
  '.js': 'text/javascript',
  '.css': 'text/css',
  '.json': 'application/json',
  '.png': 'image/png',
  '.jpg': 'image/jpg',
  '.svg': 'image/svg+xml'
};

const server = http.createServer(async (req, res) => {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type, Authorization");
  if (req.method==="OPTIONS") { res.writeHead(204); res.end(); return; }

  const url = req.url.split("?")[0];

  if (url.startsWith("/api/")) {
      res.setHeader("Content-Type", "application/json");
      if (req.method==="GET" && url==="/api/health") {
        res.writeHead(200); res.end(JSON.stringify({ok:true, uptime:process.uptime()}));
        return;
      }
      if (req.method==="GET" && url==="/api/data") {
        res.writeHead(200); res.end(JSON.stringify({ok:true, data:readData()}));
        return;
      }
      if (req.method==="POST" && url==="/api/data") {
        try {
          const payload = await readBody(req);
          if (!payload.items||!payload.products) throw new Error("Missing items or products");
          const ok = writeData({items:payload.items, products:payload.products, purchaseOrders:payload.purchaseOrders||[]});
          res.writeHead(ok?200:500);
          res.end(JSON.stringify(ok?{ok:true,savedAt:new Date().toISOString()}:{ok:false,error:"Write failed"}));
        } catch(e) { res.writeHead(400); res.end(JSON.stringify({ok:false,error:e.message})); }
        return;
      }
      res.writeHead(404); res.end(JSON.stringify({ok:false,error:"Not found"}));
      return;
  }

  // Serve static files from /dist
  let requestPath = url.replace(/^\/spa/, '') || '/index.html';
  if (requestPath === '/') requestPath = '/index.html';
  
  let filePath = path.join(__dirname, 'dist', requestPath);
  
  if (!fs.existsSync(filePath) || fs.statSync(filePath).isDirectory()) {
      filePath = path.join(__dirname, 'dist', 'index.html');
  }
  
  try {
      const extname = String(path.extname(filePath)).toLowerCase();
      const contentType = mimeTypes[extname] || 'application/octet-stream';
      const content = fs.readFileSync(filePath);
      res.writeHead(200, { 'Content-Type': contentType });
      res.end(content, 'utf-8');
  } catch (e) {
      res.writeHead(404);
      res.end('Not found');
  }
});

server.listen(PORT, "127.0.0.1", ()=>{
  console.log(`SpaTrack running → http://127.0.0.1:${PORT}`);
});

server.on("error", e=>{ console.error("Server error:", e.message); process.exit(1); });
