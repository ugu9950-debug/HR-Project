#!/usr/bin/env node
/**
 * 퇴직율 분석 대시보드 — 데이터 생성기
 * ─────────────────────────────────────────────────────────────────────────
 * 엑셀 워크북(.xlsx)의 RawData를 읽어 템플릿과 동일한 정의로 집계한 뒤,
 * turnover_dashboard_prototype.jsx 의 GENERATED-DATA 구간을 통째로 교체합니다.
 *
 *   node tools/build-dashboard-data.mjs [워크북.xlsx] [옵션]
 *
 * 옵션
 *   --기준일 2026-08-25   재직자 근속·당해연도 기말의 기준일 (기본: 워크북 기준설정 C4, 없으면 오늘)
 *   --연도 2022-2026      분석 5개년 (기본: 워크북 기준설정 7행, 없으면 기준일 기준 최근 5년)
 *   --본부최소 5          연평균 재직인원이 이보다 작은 본부는 [기타(미매핑)]로 합산 (기본 5)
 *   --교차표 12           본부 × 근속구간 교차표에 넣을 본부 수 (기본 12)
 *   --직급상세 20         직급 상세 표에 넣을 직급 수 (기본 20)
 *   --keep-reasons        퇴사사유가 비어 있어도 JSX의 기존 사유 블록을 그대로 둠 (데모용)
 *   --out <경로>          교체할 JSX 경로 (기본: ../turnover_dashboard_prototype.jsx)
 *   --dry-run             파일을 쓰지 않고 리포트만 출력
 *
 * 의존성 없음 — Node 18+ 표준 모듈만 사용합니다.
 */

import fs from "node:fs";
import path from "node:path";
import zlib from "node:zlib";
import { fileURLToPath } from "node:url";

const HERE = path.dirname(fileURLToPath(import.meta.url));

/* ═══════════════════════════════════════════════════════════════════
   0. 인자
   ═══════════════════════════════════════════════════════════════════ */
function parseArgs(argv) {
  const o = { flags: new Set(), opts: {}, positional: [] };
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i];
    if (!a.startsWith("--")) { o.positional.push(a); continue; }
    const key = a.slice(2);
    if (["keep-reasons", "dry-run"].includes(key)) { o.flags.add(key); continue; }
    o.opts[key] = argv[++i];
  }
  return o;
}
const ARGS = parseArgs(process.argv.slice(2));

const 워크북 = path.resolve(ARGS.positional[0] || path.join(HERE, "..", "퇴직율_분석_raw.xlsx"));
const 대상JSX = path.resolve(ARGS.opts.out || path.join(HERE, "..", "turnover_dashboard_prototype.jsx"));
const 본부최소 = Number(ARGS.opts["본부최소"] ?? 5);
const 교차표수 = Number(ARGS.opts["교차표"] ?? 12);
const 직급상세수 = Number(ARGS.opts["직급상세"] ?? 20);
const KEEP_REASONS = ARGS.flags.has("keep-reasons");
const DRY = ARGS.flags.has("dry-run");

const warn = (m) => console.warn("  ! " + m);
const die = (m) => { console.error("\n✗ " + m + "\n"); process.exit(1); };

/* ═══════════════════════════════════════════════════════════════════
   1. 최소 ZIP 리더 (xlsx = zip)
   ═══════════════════════════════════════════════════════════════════ */
function unzip(buf) {
  // End of Central Directory 를 뒤에서부터 찾는다
  let eocd = -1;
  for (let i = buf.length - 22; i >= 0 && i > buf.length - 22 - 65536; i--) {
    if (buf.readUInt32LE(i) === 0x06054b50) { eocd = i; break; }
  }
  if (eocd < 0) die("ZIP 구조를 찾을 수 없습니다. 올바른 .xlsx 파일인지 확인하세요: " + 워크북);

  let count = buf.readUInt16LE(eocd + 10);
  let ptr = buf.readUInt32LE(eocd + 16);

  // ZIP64 (엔트리가 많거나 파일이 큰 경우)
  if (count === 0xffff || ptr === 0xffffffff) {
    const loc = eocd - 20;
    if (loc >= 0 && buf.readUInt32LE(loc) === 0x07064b50) {
      const z64 = Number(buf.readBigUInt64LE(loc + 8));
      if (buf.readUInt32LE(z64) === 0x06064b50) {
        count = Number(buf.readBigUInt64LE(z64 + 32));
        ptr = Number(buf.readBigUInt64LE(z64 + 48));
      }
    }
  }

  const files = new Map();
  for (let i = 0; i < count; i++) {
    if (buf.readUInt32LE(ptr) !== 0x02014b50) break;
    const nameLen = buf.readUInt16LE(ptr + 28);
    const extraLen = buf.readUInt16LE(ptr + 30);
    const commentLen = buf.readUInt16LE(ptr + 32);
    const name = buf.toString("utf8", ptr + 46, ptr + 46 + nameLen);
    let localOff = buf.readUInt32LE(ptr + 42);

    if (localOff === 0xffffffff) { // ZIP64 extra field 에서 오프셋 회수
      let e = ptr + 46 + nameLen;
      const end = e + extraLen;
      while (e + 4 <= end) {
        const id = buf.readUInt16LE(e), sz = buf.readUInt16LE(e + 2);
        if (id === 0x0001) {
          let q = e + 4;
          if (buf.readUInt32LE(ptr + 24) === 0xffffffff) q += 8; // uncompressed
          if (buf.readUInt32LE(ptr + 20) === 0xffffffff) q += 8; // compressed
          localOff = Number(buf.readBigUInt64LE(q));
          break;
        }
        e += 4 + sz;
      }
    }

    // Local file header 에서 실제 데이터 시작점 계산
    const lNameLen = buf.readUInt16LE(localOff + 26);
    const lExtraLen = buf.readUInt16LE(localOff + 28);
    const dataStart = localOff + 30 + lNameLen + lExtraLen;
    const method = buf.readUInt16LE(ptr + 10);
    let compSize = buf.readUInt32LE(ptr + 20);
    if (compSize === 0xffffffff) {
      let e = ptr + 46 + nameLen; const end = e + extraLen;
      while (e + 4 <= end) {
        const id = buf.readUInt16LE(e), sz = buf.readUInt16LE(e + 2);
        if (id === 0x0001) {
          let q = e + 4;
          if (buf.readUInt32LE(ptr + 24) === 0xffffffff) q += 8;
          compSize = Number(buf.readBigUInt64LE(q));
          break;
        }
        e += 4 + sz;
      }
    }
    const raw = buf.subarray(dataStart, dataStart + compSize);
    files.set(name, method === 0 ? raw : zlib.inflateRawSync(raw));

    ptr += 46 + nameLen + extraLen + commentLen;
  }
  return files;
}

/* ═══════════════════════════════════════════════════════════════════
   2. 최소 XLSX 파서
   ═══════════════════════════════════════════════════════════════════ */
const unesc = (s) =>
  s.replace(/&#(\d+);/g, (_, d) => String.fromCharCode(+d))
   .replace(/&lt;/g, "<").replace(/&gt;/g, ">")
   .replace(/&quot;/g, '"').replace(/&apos;/g, "'")
   .replace(/&amp;/g, "&");

function readSharedStrings(files) {
  const xml = files.get("xl/sharedStrings.xml");
  if (!xml) return [];
  const s = xml.toString("utf8");
  const out = [];
  const si = /<si>([\s\S]*?)<\/si>/g;
  let m;
  while ((m = si.exec(s))) {
    let t = "";
    const tt = /<t[^>]*>([\s\S]*?)<\/t>/g;
    let n;
    while ((n = tt.exec(m[1]))) t += n[1];
    out.push(unesc(t));
  }
  return out;
}

/** 시트 이름 → 파일 경로 */
function sheetMap(files) {
  const wb = files.get("xl/workbook.xml")?.toString("utf8") || "";
  const rels = files.get("xl/_rels/workbook.xml.rels")?.toString("utf8") || "";
  const relTarget = new Map();
  for (const m of rels.matchAll(/<Relationship Id="([^"]+)"[^>]*Target="([^"]+)"/g)) {
    relTarget.set(m[1], m[2]);
  }
  const map = new Map();
  for (const m of wb.matchAll(/<sheet [^>]*name="([^"]+)"[^>]*r:id="([^"]+)"/g)) {
    let t = relTarget.get(m[2]);
    if (!t) continue;
    if (!t.startsWith("xl/")) t = "xl/" + t.replace(/^\/?/, "");
    map.set(unesc(m[1]), t);
  }
  return map;
}

const colToIdx = (c) => { let n = 0; for (const ch of c) n = n * 26 + (ch.charCodeAt(0) - 64); return n - 1; };

/** 시트를 rows[rowIndex][colIndex] = 값(문자열|숫자) 로 읽는다 */
function readSheet(files, sheetPath, strings) {
  const buf = files.get(sheetPath);
  if (!buf) return [];
  const xml = buf.toString("utf8");
  const rows = [];
  for (const rm of xml.matchAll(/<row[^>]*r="(\d+)"[^>]*>([\s\S]*?)<\/row>/g)) {
    const r = +rm[1] - 1;
    const cells = [];
    for (const cm of rm[2].matchAll(/<c r="([A-Z]+)\d+"([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)) {
      const ci = colToIdx(cm[1]);
      const attrs = cm[2] || "", inner = cm[3] || "";
      const t = /t="([^"]+)"/.exec(attrs)?.[1] || "n";
      const v = /<v>([\s\S]*?)<\/v>/.exec(inner)?.[1];
      const is = /<is>[\s\S]*?<t[^>]*>([\s\S]*?)<\/t>/.exec(inner)?.[1];
      let val;
      if (t === "s" && v != null) val = strings[+v];
      else if (t === "inlineStr" && is != null) val = unesc(is);
      else if (t === "str" && v != null) val = unesc(v);
      else if (v != null) val = Number(v);
      if (val !== undefined && val !== "") cells[ci] = val;
    }
    rows[r] = cells;
  }
  return rows;
}

/* ═══════════════════════════════════════════════════════════════════
   3. 날짜 유틸 (엑셀 serial 기준)
   ═══════════════════════════════════════════════════════════════════ */
const EPOCH = Date.UTC(1899, 11, 30);
const ymdToSerial = (y, m, d) => Math.round((Date.UTC(y, m - 1, d) - EPOCH) / 86400000);
const serialToYMD = (s) => {
  const d = new Date(EPOCH + s * 86400000);
  return [d.getUTCFullYear(), d.getUTCMonth() + 1, d.getUTCDate()];
};
const fmtDate = (s) => { const [y, m, d] = serialToYMD(s); return `${y}-${String(m).padStart(2, "0")}-${String(d).padStart(2, "0")}`; };

/** 셀 값을 엑셀 serial 로 정규화 (숫자 serial, "yyyy-mm-dd", "yyyy/mm/dd" 모두 허용) */
function toSerial(v) {
  if (v == null || v === "") return null;
  if (typeof v === "number") return Number.isFinite(v) && v > 0 ? Math.round(v) : null;
  const m = /^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})/.exec(String(v).trim());
  if (m) return ymdToSerial(+m[1], +m[2], +m[3]);
  return null;
}

/* ═══════════════════════════════════════════════════════════════════
   4. 기본값 — 워크북에 [기준설정]이 없을 때 쓰는 템플릿 표준값
   ═══════════════════════════════════════════════════════════════════ */
const 기본_근속구간 = [
  ["3개월 이내", 90], ["6개월 이내", 180], ["1년 이내", 365],
  ["1~3년", 1095], ["3~5년", 1825], ["5년 이상", 99999],
];
const 기본_직급그룹 = new Map(Object.entries({
  회장: "임원", 부회장: "임원", 사장: "임원", 부사장: "임원", 전무: "임원", 상무: "임원",
  상무보: "임원", 이사: "임원", 이사부장: "임원", 사외이사: "임원", 감사: "임원", 고문: "임원",
  부장: "부장급", 실장: "부장급", 공장장: "부장급",
  차장: "차장급", 과장: "과장급", 대리: "대리급",
  사원: "사원급", "사원 4을": "사원급", "사원 5급": "사원급",
  판매사원: "판매직", 인턴: "인턴",
}));
const 기본_사유구분 = new Map(); // 워크북에 목록이 없으면 사유 집계는 비활성

/* ═══════════════════════════════════════════════════════════════════
   5. 워크북 읽기
   ═══════════════════════════════════════════════════════════════════ */
if (!fs.existsSync(워크북)) die("워크북을 찾을 수 없습니다: " + 워크북);

console.log("\n■ 퇴직율 대시보드 데이터 생성");
console.log("  워크북 : " + 워크북);

const files = unzip(fs.readFileSync(워크북));
const strings = readSharedStrings(files);
const sheets = sheetMap(files);
console.log("  시트   : " + [...sheets.keys()].join(", "));

/* ---- RawData 찾기 ---- */
const rawName = [...sheets.keys()].find((n) => /rawdata/i.test(n.replace(/\s/g, "")));
if (!rawName) die("[RawData] 시트를 찾을 수 없습니다.");
const rawRows = readSheet(files, sheets.get(rawName), strings);

/* ---- 헤더 기반 컬럼 매핑 (raw / 템플릿 레이아웃 모두 대응) ---- */
const HEADER_ALIASES = {
  사번: ["사번", "사원번호"], 본부: ["본부", "소속본부"], 부서: ["부서", "팀"],
  직군: ["직군"], 직급: ["직급", "직위"], 입사일: ["입사일", "입사일자"],
  퇴사일: ["퇴사일", "퇴사일자", "퇴직일"], 퇴사사유: ["퇴사사유", "퇴직사유"],
  직급그룹: ["직급그룹"], 사유구분: ["사유구분"],
};
const headerRow = rawRows[0] || [];
const COL = {};
headerRow.forEach((h, i) => {
  const norm = String(h).replace(/\s|\(자동\)|\n|\r/g, "");
  for (const [key, aliases] of Object.entries(HEADER_ALIASES)) {
    if (COL[key] === undefined && aliases.some((a) => norm === a)) COL[key] = i;
  }
});
for (const need of ["본부", "직급", "입사일", "퇴사일"]) {
  if (COL[need] === undefined) die(`[${rawName}] 1행에서 [${need}] 열을 찾지 못했습니다. 헤더 이름을 확인하세요.`);
}

const 직원 = [];
for (let r = 1; r < rawRows.length; r++) {
  const row = rawRows[r];
  if (!row) continue;
  const 입사 = toSerial(row[COL.입사일]);
  if (입사 == null) continue; // 입사일 없는 행은 집계 불가
  직원.push({
    본부: String(row[COL.본부] ?? "").trim(),
    직급: String(row[COL.직급] ?? "").trim(),
    입사,
    퇴사: toSerial(row[COL.퇴사일]),
    사유: COL.퇴사사유 !== undefined ? String(row[COL.퇴사사유] ?? "").trim() : "",
  });
}
console.log(`  원본   : ${rawRows.length - 1}행 중 집계 대상 ${직원.length}행`);
if (직원.length === 0) die("집계할 데이터가 없습니다. [입사일]이 채워져 있는지 확인하세요.");

/* ---- 기준설정 읽기 ---- */
const cfgName = [...sheets.keys()].find((n) => /기준설정/.test(n));
const cfg = cfgName ? readSheet(files, sheets.get(cfgName), strings) : [];
const cell = (r, c) => cfg[r - 1]?.[colToIdx(c)];

/** 두 열을 위→아래로 훑어 [키, 값] 목록을 만든다 (빈 칸에서 멈추지 않고 끝까지 훑음) */
function readPairs(startRow, endRow, colA, colB) {
  const out = [];
  for (let r = startRow; r <= endRow; r++) {
    const a = cell(r, colA);
    if (a == null || String(a).trim() === "") continue;
    out.push([String(a).trim(), cell(r, colB)]);
  }
  return out;
}

/* 기준일 */
let 기준일 = toSerial(ARGS.opts["기준일"]);
if (기준일 == null && cfgName) 기준일 = toSerial(cell(4, "C"));
if (기준일 == null) {
  const now = new Date();
  기준일 = ymdToSerial(now.getFullYear(), now.getMonth() + 1, now.getDate());
  warn("기준일을 워크북에서 찾지 못해 오늘 날짜(" + fmtDate(기준일) + ")를 씁니다.");
}

/* 연도 */
let 연도 = [];
if (ARGS.opts["연도"]) {
  const m = /^(\d{4})\D+(\d{4})$/.exec(ARGS.opts["연도"]);
  if (!m) die("--연도 형식은 2022-2026 입니다.");
  for (let y = +m[1]; y <= +m[2]; y++) 연도.push(y);
} else if (cfgName) {
  for (const c of ["C", "D", "E", "F", "G", "H", "I", "J"]) {
    const v = cell(7, c);
    if (typeof v === "number" && v > 1900 && v < 2200) 연도.push(v);
  }
}
if (연도.length === 0) {
  const endY = serialToYMD(기준일)[0];
  연도 = [endY - 4, endY - 3, endY - 2, endY - 1, endY];
  warn("분석 연도를 워크북에서 찾지 못해 " + 연도[0] + "~" + 연도.at(-1) + " 을 씁니다.");
}

/* 근속구간 (기준설정 I·J) */
let 근속구간 = cfgName ? readPairs(15, 40, "I", "J").filter(([, v]) => typeof v === "number") : [];
if (근속구간.length === 0) { 근속구간 = 기본_근속구간; if (cfgName) warn("근속구간 목록(기준설정 I·J)이 비어 기본값을 씁니다."); }
근속구간.sort((a, b) => a[1] - b[1]);

/* 직급 → 직급그룹 (기준설정 D·E) */
let 직급그룹맵 = new Map();
if (cfgName) for (const [k, v] of readPairs(15, 60, "D", "E")) if (v) 직급그룹맵.set(k, String(v).trim());
if (직급그룹맵.size === 0) { 직급그룹맵 = 기본_직급그룹; if (cfgName) warn("직급그룹 매핑(기준설정 D·E)이 비어 기본값을 씁니다."); }

/* 직급그룹 표시 순서 (기준설정 G) */
let 직급그룹순서 = cfgName ? readPairs(15, 40, "G", "G").map(([k]) => k).filter((k) => !/^\d+$/.test(k)) : [];
if (직급그룹순서.length === 0) 직급그룹순서 = [...new Set([...직급그룹맵.values()])];

/* 퇴사사유 → 사유구분 (기준설정 L·M) */
let 사유구분맵 = new Map();
let 사유목록 = [];
if (cfgName) {
  for (const [k, v] of readPairs(15, 45, "L", "M")) {
    if (/^\d+$/.test(k)) continue; // 목록 여유칸에 들어간 숫자 잔재 무시
    사유목록.push(k);
    사유구분맵.set(k, v ? String(v).trim() : "기타");
  }
}
if (사유구분맵.size === 0) 사유구분맵 = 기본_사유구분;

console.log(`  기준일 : ${fmtDate(기준일)}`);
console.log(`  연도   : ${연도.join(", ")}`);

/* ═══════════════════════════════════════════════════════════════════
   6. 집계 — 템플릿 정의 그대로
   ═══════════════════════════════════════════════════════════════════ */
const 기간 = 연도.map((y) => {
  const s = ymdToSerial(y, 1, 1);
  const e = Math.min(ymdToSerial(y, 12, 31), 기준일);
  return { y, s, e, days: e - s + 1 };
});
const 진행중 = 기간.findIndex((p) => p.e < ymdToSerial(p.y, 12, 31));

const 재직at = (list, at) => list.filter((x) => x.입사 <= at && (x.퇴사 == null || x.퇴사 > at)).length;
const 퇴사in = (list, s, e) => list.filter((x) => x.퇴사 != null && x.퇴사 >= s && x.퇴사 <= e);
const 입사in = (list, s, e) => list.filter((x) => x.입사 >= s && x.입사 <= e).length;
const 평균재직 = (list, p) => (재직at(list, p.s) + 재직at(list, p.e)) / 2;

const 근속일 = (x) => (x.퇴사 != null ? x.퇴사 - x.입사 : 기준일 - x.입사);
const 구간of = (days) => (근속구간.find(([, cap]) => days <= cap) || 근속구간.at(-1))[0];
const 그룹of = (x) => 직급그룹맵.get(x.직급) || (x.직급 === "" ? "(미입력)" : "기타");

const r1 = (v) => Math.round(v * 10) / 10;
const r2 = (v) => Math.round(v * 100) / 100;

/* ---- 종합 ---- */
const 종합 = {
  기초: 기간.map((p) => 재직at(직원, p.s)),
  기말: 기간.map((p) => 재직at(직원, p.e)),
  입사: 기간.map((p) => 입사in(직원, p.s, p.e)),
  퇴사: 기간.map((p) => 퇴사in(직원, p.s, p.e).length),
  평균근속년: 기간.map((p) => {
    const L = 퇴사in(직원, p.s, p.e);
    return L.length ? r2(L.reduce((a, x) => a + (x.퇴사 - x.입사), 0) / L.length / 365) : 0;
  }),
};

/* ---- 차원별 집계 헬퍼 ---- */
function 차원(keyFn, keys) {
  return keys.map((k) => {
    const sub = 직원.filter((x) => keyFn(x) === k);
    return {
      name: k,
      avg: 기간.map((p) => r1(평균재직(sub, p))),
      cnt: 기간.map((p) => 퇴사in(sub, p.s, p.e).length),
    };
  });
}

/* ---- 본부: 이름이 정상이고 규모가 있는 것만 개별 행, 나머지는 합산 ---- */
const 잔여라벨 = "기타(미매핑)";
const 본부키 = [...new Set(직원.map((x) => x.본부))].filter(Boolean);
const 본부전체 = 차원((x) => x.본부, 본부키);
const 코드형 = 본부전체.filter((b) => /^\d+$/.test(b.name));
const 본부유지 = 본부전체
  .filter((b) => !/^\d+$/.test(b.name) && Math.max(...b.avg) >= 본부최소)
  .sort((a, b) => b.avg.at(-1) - a.avg.at(-1));
const 본부합산대상 = 본부전체.filter((b) => !본부유지.includes(b));

const 본부 = [...본부유지];
if (본부합산대상.length || 직원.some((x) => !x.본부)) {
  const 무본부 = 직원.filter((x) => !x.본부);
  const 잔여 = {
    name: 잔여라벨,
    avg: 기간.map((p, i) =>
      r1(Math.max(0, (재직at(직원, p.s) + 재직at(직원, p.e)) / 2 - 본부유지.reduce((a, b) => a + b.avg[i], 0)))),
    cnt: 기간.map((p, i) =>
      Math.max(0, 종합.퇴사[i] - 본부유지.reduce((a, b) => a + b.cnt[i], 0))),
  };
  if (잔여.cnt.some((v) => v > 0) || 잔여.avg.some((v) => v > 0)) 본부.push(잔여);
  if (무본부.length) warn(`본부가 비어 있는 행 ${무본부.length}건 → [${잔여라벨}]로 합산했습니다.`);
}
if (코드형.length) warn(`본부가 코드값인 항목 ${코드형.length}종 → [${잔여라벨}]로 합산했습니다. (예: ${코드형.slice(0, 3).map((b) => b.name).join(", ")})`);

/* ---- 근속구간 ---- */
const 근속 = 근속구간.map(([name]) => ({
  name,
  cnt: 기간.map((p) => 퇴사in(직원, p.s, p.e).filter((x) => 구간of(x.퇴사 - x.입사) === name).length),
}));
const 조기구간수 = Math.max(1, 근속구간.filter(([, cap]) => cap <= 365).length);

/* ---- 본부 × 근속구간 (5개년 누계) ---- */
const S0 = 기간[0].s, E9 = 기간.at(-1).e;
const 본부X근속 = [...본부]
  .map((b) => ({ name: b.name, 누계: b.cnt.reduce((a, v) => a + v, 0) }))
  .sort((a, b) => b.누계 - a.누계)
  .slice(0, 교차표수)
  .map(({ name }) => {
    const inScope = (x) =>
      name === 잔여라벨 ? !본부유지.some((b) => b.name === x.본부) : x.본부 === name;
    const L = 퇴사in(직원.filter(inScope), S0, E9);
    return { name, cnt: 근속구간.map(([bk]) => L.filter((x) => 구간of(x.퇴사 - x.입사) === bk).length) };
  });

/* ---- 직급그룹 / 직급 상세 ---- */
const 그룹키 = 직급그룹순서.filter((g) => 직원.some((x) => 그룹of(x) === g));
for (const g of [...new Set(직원.map(그룹of))]) if (!그룹키.includes(g)) 그룹키.push(g);
const 직급그룹 = 차원(그룹of, 그룹키).filter((g) => g.cnt.some((v) => v > 0) || g.avg.some((v) => v > 0));

const 직급키 = [...new Set(직원.map((x) => x.직급))].filter(Boolean);
const 직급상세 = 차원((x) => x.직급, 직급키)
  .filter((g) => g.cnt.some((v) => v > 0) || Math.max(...g.avg) >= 3)
  .sort((a, b) => b.cnt.reduce((x, y) => x + y, 0) - a.cnt.reduce((x, y) => x + y, 0))
  .slice(0, 직급상세수);

/* ---- 퇴사사유 ---- */
const 사유입력수 = 직원.filter((x) => x.퇴사 != null && x.사유).length;
const 사유_실측 = 사유입력수 > 0;
let 사유 = [];
if (사유_실측) {
  const keys = 사유목록.length ? 사유목록 : [...new Set(직원.map((x) => x.사유))].filter(Boolean);
  사유 = keys.map((k) => ({
    name: k,
    kind: 사유구분맵.get(k) || "기타",
    cnt: 기간.map((p) => 퇴사in(직원, p.s, p.e).filter((x) => x.사유 === k).length),
  })).filter((r) => r.cnt.some((v) => v > 0));
  const 미분류 = 퇴사in(직원, S0, E9).filter((x) => x.사유 && !사유구분맵.has(x.사유));
  if (미분류.length) warn(`기준설정 L열에 없는 퇴사사유 ${new Set(미분류.map((x) => x.사유)).size}종(${미분류.length}건)은 [기타]로 처리했습니다.`);
  const 미입력 = 퇴사in(직원, S0, E9).filter((x) => !x.사유).length;
  if (미입력) warn(`5개년 퇴사자 중 사유 미입력 ${미입력}건은 §4에서 제외됩니다.`);
} else {
  warn("[퇴사사유]가 한 건도 입력되어 있지 않습니다. §4와 자발/비자발 KPI는 값을 낼 수 없습니다.");
}

/* ═══════════════════════════════════════════════════════════════════
   7. 정합성 검증
   ═══════════════════════════════════════════════════════════════════ */
console.log("\n■ 정합성 검증");
let 실패 = 0;
const 확인 = (label, a, b) => {
  const ok = Math.abs(a - b) < 0.51;
  if (!ok) 실패++;
  console.log(`  ${ok ? "ok  " : "FAIL"} ${label.padEnd(28)} ${a} vs ${b}`);
};
기간.forEach((p, i) => {
  확인(`${p.y} 본부 퇴사자 합`, 본부.reduce((a, b) => a + b.cnt[i], 0), 종합.퇴사[i]);
  확인(`${p.y} 직급그룹 퇴사자 합`, 직급그룹.reduce((a, b) => a + b.cnt[i], 0), 종합.퇴사[i]);
  확인(`${p.y} 근속구간 퇴사자 합`, 근속.reduce((a, b) => a + b.cnt[i], 0), 종합.퇴사[i]);
  확인(`${p.y} 본부 평균재직 합`, 본부.reduce((a, b) => a + b.avg[i], 0), (종합.기초[i] + 종합.기말[i]) / 2);
  if (사유_실측) 확인(`${p.y} 사유 퇴사자 합`, 사유.reduce((a, b) => a + b.cnt[i], 0), 종합.퇴사[i]);
});

console.log("\n■ 요약");
기간.forEach((p, i) => {
  const 평균 = (종합.기초[i] + 종합.기말[i]) / 2;
  const rate = 평균 > 0 ? (종합.퇴사[i] / 평균) * 100 : null;
  const ann = i === 진행중 && rate != null ? (rate * 365) / p.days : null;
  console.log(
    `  ${p.y}  평균재직 ${String(r1(평균)).padStart(7)}명 · 입사 ${String(종합.입사[i]).padStart(4)} · 퇴사 ${String(종합.퇴사[i]).padStart(4)}` +
    ` · 퇴직율 ${rate == null ? "  —  " : rate.toFixed(1) + "%"}` +
    (ann != null ? `  (YTD ${p.days}일 · 연환산 ${ann.toFixed(1)}%)` : "")
  );
});

/* ═══════════════════════════════════════════════════════════════════
   8. JSX 데이터 블록 생성 & 교체
   ═══════════════════════════════════════════════════════════════════ */
const q = (s) => JSON.stringify(s);
const arr = (a) => "[" + a.join(", ") + "]";
const pad = (s, n) => s + " ".repeat(Math.max(0, n - [...s].reduce((w, ch) => w + (ch.charCodeAt(0) > 0x2e80 ? 2 : 1), 0)));

function rowsBlock(rows, withAvg = true) {
  const w = Math.max(...rows.map((r) => [...q(r.name)].reduce((a, c) => a + (c.charCodeAt(0) > 0x2e80 ? 2 : 1), 0)));
  return rows
    .map((r) => `  { name: ${pad(q(r.name) + ",", w + 1)} ${withAvg ? `avg: ${pad(arr(r.avg) + ",", 0)} ` : ""}cnt: ${arr(r.cnt)} },`)
    .join("\n");
}

const 사유Block = (!사유_실측 && KEEP_REASONS)
  ? null // 기존 블록 유지
  : 사유.length
    ? 사유.map((r) => `  { name: ${q(r.name)}, kind: ${q(r.kind)}, cnt: ${arr(r.cnt)} },`).join("\n")
    : "";

const 생성시각 = new Date().toISOString().slice(0, 19).replace("T", " ");

let block = `const 기준 = {
  기준일: ${q(fmtDate(기준일))},
  연도: ${arr(연도)},
  기간일수: ${arr(기간.map((p) => p.days))}, // 기말일 − 기초일 + 1
  진행중: ${진행중}, // 기준일까지만 집계된 연도 인덱스 (YTD, 없으면 -1)
};

/* [종합대시보드] 요약 표 */
const 종합 = {
  기초: ${arr(종합.기초)},
  기말: ${arr(종합.기말)},
  입사: ${arr(종합.입사)},
  퇴사: ${arr(종합.퇴사)},
  평균근속년: ${arr(종합.평균근속년)},
};

/* [①] 본부 × 연도 — avg = 평균 재직인원(명), cnt = 퇴사자 수(명) */
const 본부 = [
${rowsBlock(본부)}
];

/* [②] 근속구간 × 연도 퇴사자 수 — 순서는 기준설정 I열 순서 그대로 */
const 근속 = [
${rowsBlock(근속, false)}
];
const 조기구간수 = ${조기구간수}; // 앞 ${조기구간수}개 구간 = "1년 이내 조기퇴직"

/* 퇴직율 순위 차트에서 제외할 소표본 기준 (연평균 재직인원, 명).
   분모가 한 자릿수면 한두 명의 퇴사로 퇴직율이 100%를 넘어 순위가 무의미해집니다.
   제외된 조직·직급도 [표] 보기에서는 그대로 보입니다. */
const 최소분모 = 10;

/* [②-표3] 본부 × 근속구간 5개년 누계 퇴사자 수 */
const 본부X근속 = [
${rowsBlock(본부X근속, false)}
];

/* [③-표1] 직급그룹 × 연도 */
const 직급그룹 = [
${rowsBlock(직급그룹)}
];

/* [③-표4] 직급 상세 × 연도 */
const 직급상세 = [
${rowsBlock(직급상세)}
];

/* [④] 퇴사사유 × 연도 — kind 는 기준설정 M열(사유구분) */
`;

if (사유Block === null) {
  block += "__KEEP_REASONS__\n";
} else {
  block += `const 사유 = [\n${사유Block}\n];\n`;
}

block += `/* RawData [퇴사사유]가 실제로 입력되어 있으면 true. false면 §4와 자발/비자발 KPI는
   가상 분포이므로 화면에 경고를 띄웁니다. 생성기가 자동으로 갱신합니다. */
const 사유_실측 = ${사유_실측};
/* 집계 원본 정보 — 생성기가 채웁니다. */
const 출처 = {
  파일: ${q(path.basename(워크북))},
  생성시각: ${q(생성시각)},
  원본행수: ${직원.length},
  본부주석: ${q(`이름이 정상이고 연평균 재직인원 ${본부최소}명 이상인 본부만 개별 표시하고, 나머지는 [${잔여라벨}]로 합산합니다.`)},
};
`;

if (!fs.existsSync(대상JSX)) die("교체 대상 JSX를 찾을 수 없습니다: " + 대상JSX);
const src = fs.readFileSync(대상JSX, "utf8");
const START = "/* <<< GENERATED-DATA-START — 이 아래는 생성기가 덮어씁니다 >>> */";
const END = "/* <<< GENERATED-DATA-END >>> */";
const si = src.indexOf(START), ei = src.indexOf(END);
if (si < 0 || ei < 0) die("JSX에서 GENERATED-DATA 마커를 찾을 수 없습니다. 마커를 지우지 마세요.");

if (사유Block === null) {
  const old = src.slice(si, ei);
  const m = /const 사유 = \[[\s\S]*?\n\];\n/.exec(old);
  if (!m) die("--keep-reasons: 기존 사유 블록을 찾지 못했습니다.");
  block = block.replace("__KEEP_REASONS__\n", m[0]);
  warn("--keep-reasons: 기존 사유 블록을 그대로 두었습니다. §4는 여전히 가상 분포입니다.");
}

const out = src.slice(0, si) + START + "\n" + block + src.slice(ei);

if (DRY) {
  console.log("\n■ --dry-run — 파일을 쓰지 않았습니다. 생성될 블록 미리보기:\n");
  console.log(block.split("\n").slice(0, 30).join("\n") + "\n  ...");
} else {
  fs.writeFileSync(대상JSX, out);
  console.log("\n✓ 갱신: " + 대상JSX);
  console.log("  미리보기 HTML도 다시 만들려면: node tools/build-preview.mjs");
}

if (실패) { console.error(`\n✗ 정합성 검증 ${실패}건 실패 — 위 FAIL 항목을 확인하세요.\n`); process.exit(1); }
console.log("");
