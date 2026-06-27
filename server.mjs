import http from "node:http";
import fs from "node:fs/promises";
import fssync from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import zlib from "node:zlib";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const PORT = Number(process.env.PORT || 3000);
const APP_VERSION = "quote-online-20260627-root-index-check";
const APP_PASSWORD = process.env.APP_PASSWORD || "quote123";
const ADMIN_PASSWORD = process.env.ADMIN_PASSWORD || APP_PASSWORD;
const DATA_DIR = process.env.DATA_DIR || path.join(__dirname, "data");
const PUBLIC_DIR = path.join(__dirname, "public");

const PRODUCTS_FILE = path.join(DATA_DIR, "products.json");
const SEED_PRODUCTS_FILE = path.join(DATA_DIR, "seed-products.json");
const QUOTES_FILE = path.join(DATA_DIR, "quotes.json");
const SEQUENCES_FILE = path.join(DATA_DIR, "sequences.json");
const SETTINGS_FILE = path.join(DATA_DIR, "settings.json");

let writeQueue = Promise.resolve();

async function ensureData() {
  await fs.mkdir(DATA_DIR, { recursive: true });
  if (!fssync.existsSync(PRODUCTS_FILE)) {
    const seed = fssync.existsSync(SEED_PRODUCTS_FILE)
      ? await fs.readFile(SEED_PRODUCTS_FILE, "utf8")
      : "[]";
    await fs.writeFile(PRODUCTS_FILE, seed);
  }
  for (const [file, fallback] of [
    [QUOTES_FILE, "[]"],
    [SEQUENCES_FILE, "{}"],
    [SETTINGS_FILE, "{}"],
  ]) {
    if (!fssync.existsSync(file)) await fs.writeFile(file, fallback);
  }
}

async function readJson(file, fallback) {
  try {
    return JSON.parse(await fs.readFile(file, "utf8"));
  } catch {
    return fallback;
  }
}

function queuedWrite(fn) {
  const next = writeQueue.then(fn, fn);
  writeQueue = next.catch(() => {});
  return next;
}

async function writeJson(file, value) {
  await fs.writeFile(file, JSON.stringify(value, null, 2));
}

function todayCompact() {
  const now = new Date();
  const parts = new Intl.DateTimeFormat("zh-CN", {
    timeZone: "Asia/Shanghai",
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
  }).formatToParts(now);
  const pick = (type) => parts.find((part) => part.type === type)?.value || "";
  return `${pick("year")}${pick("month")}${pick("day")}`;
}

function issueQuoteNo() {
  return queuedWrite(async () => {
    const today = todayCompact();
    const sequences = await readJson(SEQUENCES_FILE, {});
    const next = Number(sequences[today] || 0) + 1;
    sequences[today] = next;
    await writeJson(SEQUENCES_FILE, sequences);
    return { quoteNo: `PMS${today}${String(next).padStart(3, "0")}`, date: today };
  });
}

function send(res, status, payload, headers = {}) {
  const body = typeof payload === "string" ? payload : JSON.stringify(payload);
  res.writeHead(status, {
    "content-type": typeof payload === "string" ? "text/plain; charset=utf-8" : "application/json; charset=utf-8",
    "cache-control": "no-store",
    ...headers,
  });
  res.end(body);
}

function sendJson(res, status, payload) {
  send(res, status, payload);
}

async function readBody(req, limit = 15 * 1024 * 1024) {
  const chunks = [];
  let size = 0;
  for await (const chunk of req) {
    size += chunk.length;
    if (size > limit) throw new Error("上传文件太大");
    chunks.push(chunk);
  }
  return Buffer.concat(chunks);
}

async function readJsonBody(req) {
  const body = await readBody(req, 2 * 1024 * 1024);
  if (!body.length) return {};
  return JSON.parse(body.toString("utf8"));
}

function isAuthed(req) {
  if (!APP_PASSWORD) return true;
  return req.headers["x-access-code"] === APP_PASSWORD;
}

function isAdmin(req) {
  if (!ADMIN_PASSWORD) return true;
  return req.headers["x-admin-code"] === ADMIN_PASSWORD;
}

function requireAuth(req, res) {
  if (isAuthed(req)) return true;
  sendJson(res, 401, { error: "请输入正确的访问口令" });
  return false;
}

function requireAdmin(req, res) {
  if (isAdmin(req)) return true;
  sendJson(res, 403, { error: "请输入正确的管理员口令" });
  return false;
}

function normalizeCode(value) {
  return String(value || "").trim().toLowerCase();
}

async function getProducts() {
  return readJson(PRODUCTS_FILE, []);
}

function publicProduct(product) {
  return {
    category: product.category || "",
    name: product.name || product.description || "",
    description: product.description || product.name || "",
    brand: product.brand || "",
    code: product.code || "",
    unit: product.unit || "",
    dealerPrice: Number(product.dealerPrice || 0),
    temporaryDealerPrice: Number(product.temporaryDealerPrice || 0) || null,
    suggestedPrice: Number(product.suggestedPrice || 0),
    pack: product.pack || "",
    image: product.image || "",
    moq: product.moq || "",
    recommendedQty: product.recommendedQty || "",
  };
}

function xmlText(value) {
  return String(value || "")
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'");
}

function zipEntries(buffer) {
  const eocdSig = 0x06054b50;
  let eocd = -1;
  for (let i = buffer.length - 22; i >= Math.max(0, buffer.length - 66000); i--) {
    if (buffer.readUInt32LE(i) === eocdSig) {
      eocd = i;
      break;
    }
  }
  if (eocd < 0) throw new Error("无法读取 Excel 文件");

  const total = buffer.readUInt16LE(eocd + 10);
  const centralOffset = buffer.readUInt32LE(eocd + 16);
  const files = new Map();
  let pos = centralOffset;

  for (let i = 0; i < total; i++) {
    if (buffer.readUInt32LE(pos) !== 0x02014b50) break;
    const method = buffer.readUInt16LE(pos + 10);
    const compressedSize = buffer.readUInt32LE(pos + 20);
    const nameLen = buffer.readUInt16LE(pos + 28);
    const extraLen = buffer.readUInt16LE(pos + 30);
    const commentLen = buffer.readUInt16LE(pos + 32);
    const localOffset = buffer.readUInt32LE(pos + 42);
    const name = buffer.slice(pos + 46, pos + 46 + nameLen).toString("utf8");

    const localNameLen = buffer.readUInt16LE(localOffset + 26);
    const localExtraLen = buffer.readUInt16LE(localOffset + 28);
    const dataStart = localOffset + 30 + localNameLen + localExtraLen;
    const raw = buffer.slice(dataStart, dataStart + compressedSize);
    const data = method === 8 ? zlib.inflateRawSync(raw) : raw;
    files.set(name, data);

    pos += 46 + nameLen + extraLen + commentLen;
  }

  return files;
}

function colIndex(cellRef) {
  const letters = String(cellRef || "A").match(/[A-Z]+/)?.[0] || "A";
  let value = 0;
  for (const letter of letters) value = value * 26 + letter.charCodeAt(0) - 64;
  return value - 1;
}

function sharedStrings(files) {
  const xml = files.get("xl/sharedStrings.xml")?.toString("utf8") || "";
  const out = [];
  const matches = xml.matchAll(/<si[^>]*>([\s\S]*?)<\/si>/g);
  for (const match of matches) {
    const texts = [...match[1].matchAll(/<t[^>]*>([\s\S]*?)<\/t>/g)].map((m) => xmlText(m[1]));
    out.push(texts.join(""));
  }
  return out;
}

function cellValue(cellXml, attr, sst) {
  const type = attr.match(/\st="([^"]+)"/)?.[1] || "";
  if (type === "inlineStr") {
    const text = [...cellXml.matchAll(/<t[^>]*>([\s\S]*?)<\/t>/g)].map((m) => xmlText(m[1])).join("");
    return text.trim();
  }
  const raw = cellXml.match(/<v[^>]*>([\s\S]*?)<\/v>/)?.[1];
  if (raw == null) return null;
  if (type === "s") return (sst[Number(raw)] || "").trim();
  const number = Number(raw);
  return Number.isFinite(number) ? number : xmlText(raw).trim();
}

function parseXlsx(buffer) {
  const files = zipEntries(buffer);
  const sheetName = files.has("xl/worksheets/sheet1.xml")
    ? "xl/worksheets/sheet1.xml"
    : [...files.keys()].find((name) => name.startsWith("xl/worksheets/sheet"));
  if (!sheetName) throw new Error("Excel 里没有找到工作表");
  const sst = sharedStrings(files);
  const xml = files.get(sheetName).toString("utf8");
  const rows = [];

  for (const rowMatch of xml.matchAll(/<row[^>]*>([\s\S]*?)<\/row>/g)) {
    const values = [];
    for (const cellMatch of rowMatch[1].matchAll(/<c([^>]*)>([\s\S]*?)<\/c>/g)) {
      const attr = cellMatch[1];
      const ref = attr.match(/\sr="([^"]+)"/)?.[1] || "";
      values[colIndex(ref)] = cellValue(cellMatch[2], attr, sst);
    }
    rows.push(values.map((v) => (v == null ? "" : v)));
  }

  return rows;
}

function normalizeHeader(value) {
  return String(value || "").replace(/\s+/g, "").toLowerCase();
}

function pick(row, headerMap, names) {
  for (const name of names) {
    const index = headerMap.get(normalizeHeader(name));
    if (index != null) return row[index] ?? "";
  }
  return "";
}

function numberOrBlank(value) {
  const number = Number(String(value ?? "").replace(/[^\d.]/g, ""));
  return Number.isFinite(number) ? number : 0;
}

function productsFromRows(rows) {
  if (!rows.length) return [];
  const headers = rows[0];
  const headerMap = new Map(headers.map((header, index) => [normalizeHeader(header), index]));
  const products = [];

  for (const row of rows.slice(1)) {
    const code = pick(row, headerMap, ["货号", "产品货号", "Stock Code", "stockcode"]);
    if (!code) continue;
    const dealerPrice = numberOrBlank(pick(row, headerMap, ["经销商价格(单价)", "经销商价格", "经销商价", "内部价"]));
    const suggestedPrice = numberOrBlank(pick(row, headerMap, ["建议售价（单价）", "建议售价(单价)", "建议售价", "终端价格", "终端价"]));
    const temporaryDealerPrice = numberOrBlank(pick(row, headerMap, ["临时经销商价格（单价）", "临时经销商价格(单价)", "临时经销商价格"]));
    const name = pick(row, headerMap, ["产品名字", "产品名称", "名称"]);
    products.push({
      category: pick(row, headerMap, ["分类"]),
      name,
      description: name,
      brand: pick(row, headerMap, ["品牌"]),
      code: String(code).trim(),
      unit: pick(row, headerMap, ["单位"]),
      dealerPrice,
      temporaryDealerPrice: temporaryDealerPrice || null,
      suggestedPrice,
      pack: pick(row, headerMap, ["包装规格", "规格"]),
      image: pick(row, headerMap, ["图片"]),
      moq: pick(row, headerMap, ["最小起订量"]),
      recommendedQty: pick(row, headerMap, ["推荐订货量"]),
    });
  }

  return products;
}

function parseMultipart(buffer, contentType) {
  const boundary = contentType.match(/boundary=(?:"([^"]+)"|([^;]+))/)?.[1] || contentType.match(/boundary=(?:"([^"]+)"|([^;]+))/)?.[2];
  if (!boundary) throw new Error("没有找到上传边界");
  const delimiter = Buffer.from(`--${boundary}`);
  const parts = [];
  let start = buffer.indexOf(delimiter);
  while (start >= 0) {
    const next = buffer.indexOf(delimiter, start + delimiter.length);
    if (next < 0) break;
    let part = buffer.slice(start + delimiter.length, next);
    if (part.slice(0, 2).toString() === "\r\n") part = part.slice(2);
    if (part.slice(-2).toString() === "\r\n") part = part.slice(0, -2);
    const headerEnd = part.indexOf(Buffer.from("\r\n\r\n"));
    if (headerEnd > 0) {
      const headerText = part.slice(0, headerEnd).toString("latin1");
      const body = part.slice(headerEnd + 4);
      const name = headerText.match(/name="([^"]+)"/)?.[1] || "";
      const filename = headerText.match(/filename="([^"]*)"/)?.[1] || "";
      parts.push({ name, filename, body });
    }
    start = next;
  }
  return parts;
}

async function handleApi(req, res, url) {
  if (url.pathname === "/api/login" && req.method === "POST") {
    const body = await readJsonBody(req);
    if (body.password === APP_PASSWORD) return sendJson(res, 200, { ok: true });
    return sendJson(res, 401, { error: "访问口令不正确" });
  }

  if (url.pathname === "/api/health") {
    return sendJson(res, 200, {
      ok: true,
      version: APP_VERSION,
      files: {
        rootIndex: fssync.existsSync(path.join(__dirname, "index.html")),
        publicIndex: fssync.existsSync(path.join(PUBLIC_DIR, "index.html")),
        server: fssync.existsSync(path.join(__dirname, "server.mjs")),
      },
    });
  }
  if (!requireAuth(req, res)) return;

  if (url.pathname === "/api/quote-number" && req.method === "POST") {
    return sendJson(res, 200, await issueQuoteNo());
  }

  if (url.pathname === "/api/products/search" && req.method === "GET") {
    const q = normalizeCode(url.searchParams.get("q"));
    const products = await getProducts();
    const matches = products
      .filter((p) => normalizeCode(p.code).includes(q) || normalizeCode(p.name).includes(q) || normalizeCode(p.description).includes(q))
      .slice(0, 15)
      .map(publicProduct);
    return sendJson(res, 200, { products: matches });
  }

  if (url.pathname.startsWith("/api/products/") && req.method === "GET") {
    const code = decodeURIComponent(url.pathname.replace("/api/products/", ""));
    const products = await getProducts();
    const product = products.find((p) => normalizeCode(p.code) === normalizeCode(code));
    if (!product) return sendJson(res, 404, { error: "没有找到这个货号" });
    return sendJson(res, 200, { product: publicProduct(product) });
  }

  if (url.pathname === "/api/seal" && req.method === "GET") {
    const settings = await readJson(SETTINGS_FILE, {});
    return sendJson(res, 200, { sealImageData: settings.sealImageData || "" });
  }

  if (url.pathname === "/api/quotes" && req.method === "POST") {
    const body = await readJsonBody(req);
    return queuedWrite(async () => {
      const quotes = await readJson(QUOTES_FILE, []);
      const index = quotes.findIndex((q) => q.quoteNo === body.quoteNo);
      const record = { ...body, updatedAt: new Date().toISOString() };
      if (index >= 0) quotes[index] = record;
      else quotes.push(record);
      await writeJson(QUOTES_FILE, quotes);
      sendJson(res, 200, { ok: true });
    });
  }

  if (url.pathname === "/api/admin/products/upload" && req.method === "POST") {
    if (!requireAdmin(req, res)) return;
    const body = await readBody(req, 30 * 1024 * 1024);
    const file = parseMultipart(body, req.headers["content-type"] || "").find((part) => part.name === "file");
    if (!file?.body?.length) return sendJson(res, 400, { error: "请上传 Excel 文件" });
    const rows = parseXlsx(file.body);
    const products = productsFromRows(rows);
    if (!products.length) return sendJson(res, 400, { error: "没有识别到产品，请检查表头是否包含货号" });
    await writeJson(PRODUCTS_FILE, products);
    return sendJson(res, 200, { ok: true, count: products.length });
  }

  if (url.pathname === "/api/admin/seal" && req.method === "POST") {
    if (!requireAdmin(req, res)) return;
    const body = await readJsonBody(req);
    const settings = await readJson(SETTINGS_FILE, {});
    settings.sealImageData = String(body.sealImageData || "");
    await writeJson(SETTINGS_FILE, settings);
    return sendJson(res, 200, { ok: true });
  }

  sendJson(res, 404, { error: "接口不存在" });
}

async function serveStatic(req, res, url) {
  const requested = decodeURIComponent(url.pathname === "/" ? "/index.html" : url.pathname);
  const safePath = path.normalize(requested).replace(/^(\.\.[/\\])+/, "");

  const candidates = [path.join(PUBLIC_DIR, safePath)];
  if (safePath === "index.html") {
    candidates.push(path.join(__dirname, "index.html"));
  }

  for (const filePath of candidates) {
    const allowed =
      filePath.startsWith(PUBLIC_DIR) ||
      (safePath === "index.html" && filePath === path.join(__dirname, "index.html"));
    if (!allowed) return send(res, 403, "Forbidden");

    try {
      const data = await fs.readFile(filePath);
      const ext = path.extname(filePath).toLowerCase();
      const type = {
        ".html": "text/html; charset=utf-8",
        ".js": "application/javascript; charset=utf-8",
        ".css": "text/css; charset=utf-8",
        ".png": "image/png",
        ".jpg": "image/jpeg",
        ".jpeg": "image/jpeg",
        ".svg": "image/svg+xml",
      }[ext] || "application/octet-stream";
      res.writeHead(200, { "content-type": type });
      res.end(data);
      return;
    } catch {
      // Try the next safe candidate.
    }
  }

  send(res, 404, `首页文件没有找到。版本：${APP_VERSION}。请确认 GitHub 仓库根目录有 index.html，或有 public/index.html。`);
}

await ensureData();

http
  .createServer(async (req, res) => {
    try {
      const url = new URL(req.url || "/", `http://${req.headers.host || "localhost"}`);
      if (url.pathname.startsWith("/api/")) await handleApi(req, res, url);
      else await serveStatic(req, res, url);
    } catch (error) {
      sendJson(res, 500, { error: error.message || "服务器出错" });
    }
  })
  .listen(PORT, () => {
    console.log(`Quote app running on http://localhost:${PORT}`);
  });
