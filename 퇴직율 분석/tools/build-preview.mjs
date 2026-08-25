#!/usr/bin/env node
/**
 * 퇴직율 분석 대시보드 — 미리보기 HTML 빌더
 * ─────────────────────────────────────────────────────────────────────────
 * turnover_dashboard_prototype.jsx 를 React·recharts까지 통째로 묶어
 * 브라우저에서 더블클릭만으로 열리는 단일 HTML 파일을 만듭니다.
 *
 *   npm install          (최초 1회)
 *   node tools/build-preview.mjs [--out 경로] [--no-minify]
 *
 * 데이터만 갱신할 때는 이 스크립트가 필요 없습니다.
 * build-dashboard-data.mjs 는 의존성 없이 동작합니다.
 */

import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(HERE, "..");
const SRC = path.join(ROOT, "turnover_dashboard_prototype.jsx");

const argv = process.argv.slice(2);
const outIdx = argv.indexOf("--out");
const OUT = path.resolve(outIdx >= 0 ? argv[outIdx + 1] : path.join(ROOT, "turnover_dashboard_preview.html"));
const MINIFY = !argv.includes("--no-minify");

let esbuild;
try {
  esbuild = await import("esbuild");
} catch {
  console.error(
    "\n✗ esbuild 를 찾을 수 없습니다.\n" +
    "  이 폴더에서 한 번만 실행해 주세요:  npm install\n" +
    "  (데이터 갱신만 하실 거면 build-dashboard-data.mjs 는 설치 없이 동작합니다.)\n"
  );
  process.exit(1);
}

if (!fs.existsSync(SRC)) {
  console.error("\n✗ 원본을 찾을 수 없습니다: " + SRC + "\n");
  process.exit(1);
}

console.log("\n■ 미리보기 HTML 빌드");
console.log("  원본 : " + SRC);

const result = await esbuild.build({
  stdin: {
    contents: `
      import React from "react";
      import { createRoot } from "react-dom/client";
      import Dash from ${JSON.stringify(SRC.split(path.sep).join("/"))};
      createRoot(document.getElementById("root")).render(React.createElement(Dash));
    `,
    resolveDir: ROOT,
    loader: "jsx",
    sourcefile: "preview-entry.jsx",
  },
  bundle: true,
  minify: MINIFY,
  format: "iife",
  platform: "browser",
  target: ["es2019"],
  loader: { ".jsx": "jsx" },
  define: { "process.env.NODE_ENV": '"production"' },
  write: false,
  logLevel: "warning",
});

const js = result.outputFiles[0].text.split("</script").join("<\\/script");

// 생성 정보를 JSX에서 읽어 페이지 하단 표시에 재사용
const src = fs.readFileSync(SRC, "utf8");
const 생성시각 = /생성시각:\s*"([^"]*)"/.exec(src)?.[1] || "";

const html = `<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>퇴직율 분석 대시보드 — 프로토타입 미리보기</title>
<style>
  html, body { margin: 0; padding: 0; background: #F4F5F7; }
  #root { min-height: 100vh; }
</style>
</head>
<body>
<!-- 자동 생성 파일입니다. 고치지 마세요.
     원본: turnover_dashboard_prototype.jsx${생성시각 ? "\n     데이터 생성: " + 생성시각 : ""}
     재생성: node tools/build-preview.mjs -->
<div id="root"></div>
<script>${js}</script>
</body>
</html>
`;

fs.writeFileSync(OUT, html);
console.log("  크기 : " + (Buffer.byteLength(html) / 1024).toFixed(0) + "KB");
console.log("\n✓ 생성: " + OUT);
console.log("  브라우저에서 파일을 열면 바로 확인할 수 있습니다.\n");
