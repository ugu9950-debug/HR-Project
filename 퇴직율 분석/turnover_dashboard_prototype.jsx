import React, { useMemo, useState } from "react";
import {
  LineChart, Line, BarChart, Bar, Cell,
  XAxis, YAxis, CartesianGrid, Tooltip, Legend, LabelList,
  ResponsiveContainer, ReferenceLine,
} from "recharts";
import {
  LogOut, Users, UserMinus, Clock, AlertTriangle, Building2, Info,
} from "lucide-react";

/* ═══════════════════════════════════════════════════════════════════════
   퇴직율 분석 대시보드 — 프로토타입
   기준 문서: 퇴직율_분석_템플릿_2022-2026.xlsx

   이 파일의 구조는 템플릿 워크북과 1:1로 대응합니다.
     [기준설정]              → 기준.연도 / 기준.기간일수
     [종합대시보드]          → KPI 6종 + §0 전사 추이
     [①_본부별_퇴직율]       → §1 (표1 퇴직율 / 표2 퇴사자 수 / 표3 평균 재직인원)
     [②_근속구간별_퇴직비중] → §2 (표2 연도별 비중 + 표3 본부 × 근속 교차표)
     [③_직급별_퇴직율]       → §3 (표1 직급그룹 / 표4 직급 상세)
     [④_퇴사사유별_퇴직율]   → §4 (표1·2 사유별 / 표4 자발·비자발 요약)

   계산 정의도 템플릿을 그대로 따릅니다.
     퇴직율        = 해당 연도 퇴사자 수 ÷ 평균 재직인원
     평균 재직인원 = (기초 재직인원 + 기말 재직인원) ÷ 2
     구성비        = 해당 연도 퇴사자 중 비율        (분모 = 퇴사자 수)
     사유별 퇴직율 = 사유별 퇴사자 ÷ 전사 평균 재직인원 (분모 = 재직인원)
     연환산 퇴직율 = 퇴직율 × 365 ÷ 기간 일수        (진행 중인 연도용)
     분모가 0이면 값을 내지 않고 "—" 로 표시합니다. (템플릿의 IF(...=0,"") 와 동일)
   ═══════════════════════════════════════════════════════════════════════ */

/* ---------------------------------------------------------------------
   1. 디자인 토큰
   시리즈 색은 색각이상(CVD) 분리도·대비 검증을 통과한 조합만 사용합니다.
     범주형 3슬롯 : #4353D9 / #C1443A / #C98A2C  (자발적 / 비자발적 / 기타)
     순서형 6단계 : 근속구간 램프 (단일 색상, 명도 단조 변화)
--------------------------------------------------------------------- */
const C = {
  bg: "#F4F5F7",
  card: "#FFFFFF",
  border: "#E3E6EC",
  grid: "#EDEFF3",
  text: "#191E2B",
  sub: "#6B7280",
  faint: "#9AA1AF",

  s1: "#4353D9", // 범주 1 · 기본 시리즈 / 자발적
  s2: "#C1443A", // 범주 2 · 퇴사 / 비자발적 / 평균 초과
  s3: "#C98A2C", // 범주 3 · 기타 (대비 2.93:1 → 반드시 직접 라벨 또는 표와 함께 노출)
  good: "#1E9C88", // 입사 등 반대 방향 지표
};

/* 근속구간 순서형 램프 — 근속이 짧을수록 진하게 (조기 이탈이 무겁게 읽히도록) */
const TENURE_RAMP = ["#252E7E", "#3441AC", "#4353D9", "#6975DC", "#8791E3", "#A5ADEA"];
/* 교차표 히트맵 램프 — 같은 색상 계열, 0에 가까울수록 배경으로 후퇴 */
const HEAT_RAMP = ["#F2F3FD", "#DFE2F9", "#C2C8F1", "#A5ADEA", "#6975DC", "#4353D9"];

const FONT_IMPORT =
  "@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');";

/* ---------------------------------------------------------------------
   2. 데이터 — 템플릿 워크북에서 생성한 집계값

   ⚠ 아래 GENERATED-DATA 구간은 생성기가 통째로 덮어씁니다. 직접 고치지 마세요.
      갱신:  node tools/build-dashboard-data.mjs 퇴직율_분석_raw.xlsx
      옵션:  --기준일 2026-08-25  --본부최소 5  --keep-reasons
--------------------------------------------------------------------- */
/* <<< GENERATED-DATA-START — 이 아래는 생성기가 덮어씁니다 >>> */
const 기준 = {
  기준일: "2026-08-25",
  연도: [2022, 2023, 2024, 2025, 2026],
  기간일수: [365, 365, 366, 365, 237], // 기말일 − 기초일 + 1
  진행중: 4, // 기준일까지만 집계된 연도 인덱스 (YTD, 없으면 -1)
};

/* [종합대시보드] 요약 표 */
const 종합 = {
  기초: [719, 752, 761, 784, 740],
  기말: [742, 758, 782, 735, 764],
  입사: [174, 174, 176, 161, 112],
  퇴사: [141, 158, 152, 208, 83],
  평균근속년: [5.23, 4.57, 3.64, 4.18, 4.77],
};

/* [①] 본부 × 연도 — avg = 평균 재직인원(명), cnt = 퇴사자 수(명) */
const 본부 = [
  { name: "소싱본부",       avg: [122.5, 141.5, 158.5, 155.5, 142], cnt: [11, 12, 31, 35, 10] },
  { name: "수출본부",       avg: [76.5, 82, 89.5, 92, 90.5], cnt: [0, 0, 0, 15, 7] },
  { name: "TT상품기획본부", avg: [51.5, 67.5, 83.5, 87, 84.5], cnt: [1, 2, 6, 26, 14] },
  { name: "TT운영본부",     avg: [59.5, 68, 69.5, 61, 62.5], cnt: [5, 10, 29, 44, 11] },
  { name: "E-Biz본부",      avg: [31.5, 40.5, 50, 54.5, 58.5], cnt: [2, 1, 3, 7, 8] },
  { name: "FS상품기획본부", avg: [26, 36.5, 51.5, 61, 58.5], cnt: [0, 0, 3, 7, 6] },
  { name: "경영지원본부",   avg: [57.5, 58, 53.5, 50, 49], cnt: [5, 9, 8, 10, 2] },
  { name: "경영기획실",     avg: [35, 37, 41, 44.5, 43], cnt: [12, 13, 0, 8, 7] },
  { name: "물류1팀",        avg: [34.5, 32, 34, 36, 35], cnt: [13, 2, 0, 2, 4] },
  { name: "FS운영본부",     avg: [12.5, 16.5, 21.5, 28.5, 33.5], cnt: [0, 0, 0, 10, 4] },
  { name: "TT마케팅본부",   avg: [14.5, 15.5, 18.5, 22, 22], cnt: [3, 1, 1, 12, 4] },
  { name: "유통사업본부",   avg: [15.5, 17, 18, 18, 18], cnt: [3, 2, 3, 1, 1] },
  { name: "물류2팀",        avg: [13, 17.5, 16, 12, 11], cnt: [0, 0, 8, 4, 0] },
  { name: "ESG지원팀",      avg: [0, 0, 0, 0, 8.5], cnt: [0, 0, 0, 0, 1] },
  { name: "FS마케팅본부",   avg: [0, 0, 0.5, 4.5, 8], cnt: [0, 0, 0, 0, 0] },
  { name: "인사본부",       avg: [4.5, 5.5, 7.5, 8.5, 7.5], cnt: [0, 0, 0, 9, 2] },
  { name: "회장비서실",     avg: [6, 7, 5.5, 2, 1.5], cnt: [0, 0, 6, 2, 1] },
  { name: "패션영업본부",   avg: [42.5, 27.5, 8, 0, 0], cnt: [19, 28, 18, 0, 0] },
  { name: "수출총괄",       avg: [45.5, 28.5, 12, 2, 0], cnt: [23, 21, 16, 4, 0] },
  { name: "R&D본부",        avg: [12.5, 8, 6.5, 3, 0], cnt: [7, 4, 1, 6, 0] },
  { name: "기타(미매핑)",   avg: [69.5, 49, 26.5, 17.5, 18.5], cnt: [37, 53, 19, 6, 1] },
];

/* [②] 근속구간 × 연도 퇴사자 수 — 순서는 기준설정 I열 순서 그대로 */
const 근속 = [
  { name: "3개월 이내", cnt: [8, 10, 13, 25, 7] },
  { name: "6개월 이내", cnt: [8, 10, 16, 15, 7] },
  { name: "1년 이내",   cnt: [12, 11, 12, 12, 2] },
  { name: "1~3년",      cnt: [42, 53, 54, 65, 28] },
  { name: "3~5년",      cnt: [22, 23, 18, 39, 16] },
  { name: "5년 이상",   cnt: [49, 51, 39, 52, 23] },
];
const 조기구간수 = 3; // 앞 3개 구간 = "1년 이내 조기퇴직"

/* 퇴직율 순위 차트에서 제외할 소표본 기준 (연평균 재직인원, 명).
   분모가 한 자릿수면 한두 명의 퇴사로 퇴직율이 100%를 넘어 순위가 무의미해집니다.
   제외된 조직·직급도 [표] 보기에서는 그대로 보입니다. */
const 최소분모 = 10;

/* [②-표3] 본부 × 근속구간 5개년 누계 퇴사자 수 */
const 본부X근속 = [
  { name: "기타(미매핑)",   cnt: [12, 11, 16, 48, 12, 17] },
  { name: "소싱본부",       cnt: [7, 4, 3, 37, 19, 29] },
  { name: "TT운영본부",     cnt: [9, 14, 6, 28, 17, 25] },
  { name: "패션영업본부",   cnt: [3, 1, 5, 22, 7, 27] },
  { name: "수출총괄",       cnt: [1, 0, 3, 18, 13, 29] },
  { name: "TT상품기획본부", cnt: [7, 11, 2, 16, 8, 5] },
  { name: "경영기획실",     cnt: [9, 4, 2, 14, 6, 5] },
  { name: "경영지원본부",   cnt: [3, 2, 0, 6, 7, 16] },
  { name: "수출본부",       cnt: [0, 1, 2, 5, 3, 11] },
  { name: "E-Biz본부",      cnt: [1, 0, 0, 12, 4, 4] },
  { name: "물류1팀",        cnt: [1, 1, 0, 5, 2, 12] },
  { name: "TT마케팅본부",   cnt: [2, 3, 3, 5, 5, 3] },
];

/* [③-표1] 직급그룹 × 연도 */
const 직급그룹 = [
  { name: "임원",   avg: [50.5, 47.5, 44, 43, 41.5], cnt: [4, 8, 4, 9, 6] },
  { name: "부장급", avg: [72, 66.5, 69.5, 67.5, 65.5], cnt: [14, 6, 7, 18, 5] },
  { name: "차장급", avg: [124, 116, 112, 105.5, 97.5], cnt: [14, 17, 10, 18, 10] },
  { name: "과장급", avg: [189, 195, 204.5, 201, 192], cnt: [21, 21, 20, 38, 18] },
  { name: "대리급", avg: [173.5, 196, 207.5, 199, 189], cnt: [31, 37, 45, 53, 20] },
  { name: "사원급", avg: [82, 109, 127.5, 143.5, 166], cnt: [32, 33, 41, 64, 24] },
  { name: "판매직", avg: [39.5, 25, 6.5, 0, 0], cnt: [25, 36, 20, 0, 0] },
  { name: "인턴",   avg: [0, 0, 0, 0, 0.5], cnt: [0, 0, 5, 8, 0] },
];

/* [③-표4] 직급 상세 × 연도 */
const 직급상세 = [
  { name: "대리",     avg: [173.5, 196, 207.5, 199, 189], cnt: [31, 37, 45, 53, 20] },
  { name: "사원",     avg: [65.5, 89, 110, 128, 149], cnt: [26, 29, 33, 60, 22] },
  { name: "과장",     avg: [189, 195, 204.5, 201, 192], cnt: [21, 21, 20, 38, 18] },
  { name: "판매사원", avg: [39.5, 25, 6.5, 0, 0], cnt: [25, 36, 20, 0, 0] },
  { name: "차장",     avg: [124, 116, 112, 105.5, 97.5], cnt: [14, 17, 10, 18, 10] },
  { name: "부장",     avg: [47, 44, 43.5, 41, 42], cnt: [8, 3, 3, 11, 3] },
  { name: "실장",     avg: [14, 13, 14.5, 14, 13], cnt: [3, 1, 4, 4, 1] },
  { name: "인턴",     avg: [0, 0, 0, 0, 0.5], cnt: [0, 0, 5, 8, 0] },
  { name: "사원 4을", avg: [11, 16, 14.5, 12, 11.5], cnt: [1, 1, 6, 2, 2] },
  { name: "사원 5급", avg: [5.5, 4, 3, 3.5, 5.5], cnt: [5, 3, 2, 2, 0] },
  { name: "이사",     avg: [18.5, 19.5, 20, 19.5, 17], cnt: [0, 0, 1, 5, 5] },
  { name: "공장장",   avg: [11, 9.5, 11.5, 12.5, 10.5], cnt: [3, 2, 0, 3, 1] },
  { name: "상무보",   avg: [13.5, 11, 8.5, 8, 8], cnt: [1, 4, 1, 1, 1] },
  { name: "이사부장", avg: [1.5, 0, 0, 0, 0], cnt: [3, 0, 0, 0, 0] },
  { name: "전무",     avg: [6.5, 6.5, 5.5, 5, 5.5], cnt: [0, 1, 1, 1, 0] },
  { name: "상무",     avg: [2.5, 3, 2.5, 3, 3.5], cnt: [0, 0, 1, 1, 0] },
  { name: "사장",     avg: [1, 0.5, 0, 0, 0], cnt: [0, 1, 0, 0, 0] },
  { name: "부사장",   avg: [3, 3, 3.5, 3.5, 3], cnt: [0, 0, 0, 1, 0] },
  { name: "사외이사", avg: [1, 1, 1, 1, 1], cnt: [0, 1, 0, 0, 0] },
  { name: "부회장",   avg: [1, 1, 1, 1, 1], cnt: [0, 1, 0, 0, 0] },
];

/* [④] 퇴사사유 × 연도 — kind 는 기준설정 M열(사유구분) */
const 사유 = [
  { name: "개인사유(이직)", kind: "자발적", cnt: [38, 41, 38, 46, 22] },
  { name: "조직/직무 부적응", kind: "자발적", cnt: [16, 17, 18, 21, 10] },
  { name: "근무조건 불만", kind: "자발적", cnt: [14, 16, 17, 19, 9] },
  { name: "개인사유(가족/육아)", kind: "자발적", cnt: [11, 13, 12, 15, 7] },
  { name: "개인사유(진로변경)", kind: "자발적", cnt: [10, 11, 11, 13, 7] },
  { name: "개인사유(건강)", kind: "자발적", cnt: [7, 8, 8, 8, 4] },
  { name: "개인사유(기타)", kind: "자발적", cnt: [6, 6, 6, 6, 3] },
  { name: "개인사유(학업)", kind: "자발적", cnt: [3, 3, 3, 4, 2] },
  { name: "권고사직", kind: "비자발적", cnt: [18, 22, 21, 35, 10] },
  { name: "구조조정", kind: "비자발적", cnt: [9, 9, 9, 27, 3] },
  { name: "징계해고", kind: "비자발적", cnt: [4, 5, 5, 6, 3] },
  { name: "사망/기타", kind: "기타", cnt: [1, 2, 1, 2, 1] },
  { name: "사유미상", kind: "기타", cnt: [4, 5, 3, 6, 2] },
];
/* RawData [퇴사사유]가 실제로 입력되어 있으면 true. false면 §4와 자발/비자발 KPI는
   가상 분포이므로 화면에 경고를 띄웁니다. 생성기가 자동으로 갱신합니다. */
const 사유_실측 = false;
/* 집계 원본 정보 — 생성기가 채웁니다. */
const 출처 = {
  파일: "퇴직율_분석_raw.xlsx",
  생성시각: "2026-08-25 05:21:33",
  원본행수: 7253,
  본부주석: "이름이 정상이고 연평균 재직인원 5명 이상인 본부만 개별 표시하고, 나머지는 [기타(미매핑)]로 합산합니다.",
};
/* <<< GENERATED-DATA-END >>> */

const KIND_COLOR = { 자발적: C.s1, 비자발적: C.s2, 기타: C.s3 };
const KIND_ORDER = ["자발적", "비자발적", "기타"];

/* ---------------------------------------------------------------------
   3. 계산 헬퍼 — 템플릿 수식과 동일한 규칙
--------------------------------------------------------------------- */
const sum = (a) => a.reduce((x, y) => x + y, 0);
/** IF(분모=0,"",분자/분모) — 분모가 0이면 값 없음(null) */
const ratio = (n, d) => (d > 0 ? n / d : null);
const 평균재직 = (i) => (종합.기초[i] + 종합.기말[i]) / 2;

const pct = (v, d = 1) => (v == null ? "—" : (v * 100).toFixed(d) + "%");
const num = (v, d = 0) =>
  v == null ? "—" : v.toLocaleString("ko-KR", { minimumFractionDigits: d, maximumFractionDigits: d });

/** 연도 인덱스(0~4) 또는 "ALL"(5개년 종합) 기준으로 한 행을 집계 */
function slice(row, yi) {
  if (yi === "ALL") {
    const c = sum(row.cnt);
    const a = row.avg ? sum(row.avg) : null; // 누계 퇴직율 분모 = 연도별 평균인원의 합
    return { cnt: c, avg: a, rate: row.avg ? ratio(c, a) : null };
  }
  const c = row.cnt[yi];
  const a = row.avg ? row.avg[yi] : null;
  return { cnt: c, avg: a, rate: row.avg ? ratio(c, a) : null };
}

/* ---------------------------------------------------------------------
   4. 공통 UI 조각
--------------------------------------------------------------------- */
function Card({ children, style }) {
  return (
    <div
      style={{
        background: C.card,
        border: "1px solid " + C.border,
        borderRadius: 14,
        padding: "20px 22px 18px",
        boxShadow: "0 1px 2px rgba(20,24,38,0.03)",
        ...style,
      }}
    >
      {children}
    </div>
  );
}

function Segmented({ options, value, onChange, size = "md" }) {
  const pad = size === "sm" ? "4px 10px" : "5px 12px";
  return (
    <div style={{ display: "flex", gap: 3, background: C.bg, borderRadius: 8, padding: 3 }}>
      {options.map((o) => {
        const on = o.value === value;
        return (
          <button
            key={String(o.value)}
            onClick={() => onChange(o.value)}
            aria-pressed={on}
            style={{
              padding: pad,
              fontSize: 12.5,
              fontWeight: 600,
              borderRadius: 6,
              border: "none",
              cursor: "pointer",
              whiteSpace: "nowrap",
              background: on ? C.card : "transparent",
              color: on ? C.text : C.sub,
              boxShadow: on ? "0 1px 3px rgba(20,24,38,0.08)" : "none",
            }}
          >
            {o.label}
          </button>
        );
      })}
    </div>
  );
}

function SectionHead({ eyebrow, title, desc, right }) {
  return (
    <div
      style={{
        display: "flex",
        justifyContent: "space-between",
        alignItems: "flex-start",
        gap: 12,
        marginBottom: 16,
        flexWrap: "wrap",
      }}
    >
      <div>
        {eyebrow && (
          <div style={{ fontSize: 11, fontWeight: 700, color: C.faint, letterSpacing: 0.4, marginBottom: 3 }}>
            {eyebrow}
          </div>
        )}
        <div style={{ fontSize: 14.5, fontWeight: 600, color: C.text }}>{title}</div>
        {desc && <div style={{ fontSize: 12, color: C.sub, marginTop: 3, lineHeight: 1.55 }}>{desc}</div>}
      </div>
      {right}
    </div>
  );
}

function Note({ children }) {
  return (
    <div
      style={{
        display: "flex",
        gap: 6,
        alignItems: "flex-start",
        marginTop: 10,
        fontSize: 11.5,
        color: C.faint,
        lineHeight: 1.6,
      }}
    >
      <Info size={12} style={{ marginTop: 2, flexShrink: 0 }} />
      <span>{children}</span>
    </div>
  );
}

/** 데이터가 실측이 아닐 때 띄우는 경고 스트립 */
function SampleWarning({ children }) {
  return (
    <div
      style={{
        display: "flex", gap: 7, alignItems: "flex-start",
        background: "#FBF3E4", border: "1px solid #F0E0C0", color: "#8A5D14",
        borderRadius: 8, padding: "9px 12px", fontSize: 12, lineHeight: 1.6, marginBottom: 14,
      }}
    >
      <AlertTriangle size={13} style={{ marginTop: 2, flexShrink: 0 }} />
      <span>{children}</span>
    </div>
  );
}

/** 시리즈 색을 텍스트에 쓰지 않고, 색 점(swatch) + 회색 텍스트로 정체성을 전달 */
function Swatch({ color, label }) {
  return (
    <span style={{ display: "inline-flex", alignItems: "center", gap: 6, fontSize: 12, color: C.sub }}>
      <span style={{ width: 9, height: 9, borderRadius: 2, background: color, flexShrink: 0 }} />
      {label}
    </span>
  );
}

function ChartTooltip({ active, payload, label, unit = "", fmt }) {
  if (!active || !payload || !payload.length) return null;
  return (
    <div
      style={{
        background: "#20242F",
        borderRadius: 8,
        padding: "10px 12px",
        fontSize: 12,
        color: "#fff",
        boxShadow: "0 6px 20px rgba(0,0,0,0.18)",
        minWidth: 150,
      }}
    >
      <div style={{ fontWeight: 600, marginBottom: 6 }}>{label}</div>
      {payload
        .filter((p) => p.value != null)
        .map((p, i) => (
          <div key={i} style={{ display: "flex", justifyContent: "space-between", gap: 18, color: "#D7DAE3", lineHeight: 1.7 }}>
            <span style={{ display: "flex", alignItems: "center", gap: 6 }}>
              <span style={{ width: 7, height: 7, borderRadius: 99, background: p.color || p.fill }} />
              {p.name}
            </span>
            <span style={{ fontWeight: 600, color: "#fff", fontVariantNumeric: "tabular-nums" }}>
              {fmt ? fmt(p.value) : num(p.value) + unit}
            </span>
          </div>
        ))}
    </div>
  );
}

/** 모든 차트의 표 대응물(table view) — 접근성 및 원표 확인용 */
function DataTable({ head, rows, align = [] }) {
  return (
    <div style={{ overflowX: "auto", marginTop: 2 }}>
      <table
        style={{
          width: "100%",
          borderCollapse: "collapse",
          fontSize: 12.5,
          fontVariantNumeric: "tabular-nums",
          whiteSpace: "nowrap",
        }}
      >
        <thead>
          <tr>
            {head.map((h, i) => (
              <th
                key={i}
                style={{
                  textAlign: i === 0 ? "left" : "right",
                  padding: "8px 10px",
                  borderBottom: "1px solid " + C.border,
                  color: C.sub,
                  fontWeight: 600,
                  fontSize: 12,
                }}
              >
                {h}
              </th>
            ))}
          </tr>
        </thead>
        <tbody>
          {rows.map((r, ri) => (
            <tr key={ri}>
              {r.map((cell, ci) => (
                <td
                  key={ci}
                  style={{
                    textAlign: ci === 0 ? "left" : align[ci] || "right",
                    padding: "7px 10px",
                    borderBottom: "1px solid " + C.grid,
                    color: ci === 0 ? C.text : C.sub,
                    fontWeight: ci === 0 ? 500 : 400,
                  }}
                >
                  {cell}
                </td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

/* ---------------------------------------------------------------------
   5. KPI 카드
--------------------------------------------------------------------- */
function Kpi({ icon: Icon, label, value, unit, delta, deltaGood, sub, badge }) {
  const hasDelta = delta != null && isFinite(delta);
  const up = hasDelta && delta > 0;
  const good = deltaGood === "down" ? !up : up;
  const deltaColor = !hasDelta || delta === 0 ? C.sub : good ? C.good : C.s2;

  return (
    <Card style={{ flex: "1 1 200px", minWidth: 200, padding: "18px 20px 16px" }}>
      <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 12 }}>
        <div
          style={{
            width: 24,
            height: 24,
            borderRadius: 7,
            background: C.bg,
            display: "flex",
            alignItems: "center",
            justifyContent: "center",
            flexShrink: 0,
          }}
        >
          <Icon size={13} color={C.sub} />
        </div>
        <div style={{ fontSize: 12.5, color: C.sub, fontWeight: 500 }}>{label}</div>
      </div>

      <div style={{ display: "flex", alignItems: "baseline", gap: 4 }}>
        {/* 히어로 수치는 본문과 같은 sans + 비례숫자 (tabular-nums 미사용) */}
        <span style={{ fontSize: 28, fontWeight: 600, color: C.text, letterSpacing: -0.6 }}>{value}</span>
        {unit && <span style={{ fontSize: 13.5, color: C.sub, fontWeight: 500 }}>{unit}</span>}
        {badge && (
          <span
            style={{
              marginLeft: 4,
              fontSize: 10.5,
              fontWeight: 700,
              color: C.s3,
              background: "#FBF3E4",
              borderRadius: 5,
              padding: "2px 6px",
            }}
          >
            {badge}
          </span>
        )}
      </div>

      <div style={{ marginTop: 8, fontSize: 12, color: C.sub, lineHeight: 1.5, minHeight: 18 }}>
        {hasDelta && (
          <span style={{ color: deltaColor, fontWeight: 600, marginRight: 6 }}>
            {up ? "▲" : delta === 0 ? "―" : "▼"} {Math.abs(delta).toFixed(1)}
            {deltaGood ? "%p" : ""}
          </span>
        )}
        {sub}
      </div>
    </Card>
  );
}

/* ═══════════════════════════════════════════════════════════════════════
   메인
   ═══════════════════════════════════════════════════════════════════════ */
export default function TurnoverDashboard() {
  const [yi, setYi] = useState(4);            // 0~4 = 연도, "ALL" = 5개년 종합
  const [asTable, setAsTable] = useState(false); // 전체 섹션 표 보기
  const [본부지표, set본부지표] = useState("rate"); // rate | cnt | avg
  const [직급뷰, set직급뷰] = useState("group");    // group | detail

  const isAll = yi === "ALL";
  const yLabel = isAll ? "5개년 종합" : String(기준.연도[yi]);
  const isYtd = !isAll && yi === 기준.진행중;

  /* ---------- 종합대시보드 파생 지표 ---------- */
  const S = useMemo(() => {
    const idx = 기준.연도.map((_, i) => i);
    const 평균 = idx.map(평균재직);

    const 퇴사 = isAll ? sum(종합.퇴사) : 종합.퇴사[yi];
    const 입사 = isAll ? sum(종합.입사) : 종합.입사[yi];
    const 평재 = isAll ? sum(평균) : 평균[yi]; // 누계 퇴직율의 분모 = 연도별 평균인원의 합
    const 퇴직율 = ratio(퇴사, 평재);

    const 조기 = isAll
      ? sum(근속.slice(0, 조기구간수).map((r) => sum(r.cnt)))
      : sum(근속.slice(0, 조기구간수).map((r) => r.cnt[yi]));
    const 조기비중 = ratio(조기, 퇴사);

    const 자발 = isAll
      ? sum(사유.filter((r) => r.kind === "자발적").map((r) => sum(r.cnt)))
      : sum(사유.filter((r) => r.kind === "자발적").map((r) => r.cnt[yi]));
    const 비자발 = isAll
      ? sum(사유.filter((r) => r.kind === "비자발적").map((r) => sum(r.cnt)))
      : sum(사유.filter((r) => r.kind === "비자발적").map((r) => r.cnt[yi]));
    const 기타 = 퇴사 - 자발 - 비자발;

    const 평균근속 = isAll
      ? sum(종합.평균근속년.map((v, i) => v * 종합.퇴사[i])) / sum(종합.퇴사)
      : 종합.평균근속년[yi];

    // 전년 대비 (5개년 종합에서는 비교 대상 없음)
    const prev =
      !isAll && yi > 0 ? ratio(종합.퇴사[yi - 1], 평균[yi - 1]) : null;

    const 연환산 = isYtd && 퇴직율 != null ? (퇴직율 * 365) / 기준.기간일수[yi] : null;

    return { 퇴사, 입사, 평재, 퇴직율, 조기, 조기비중, 자발, 비자발, 기타, 평균근속, prev, 연환산, 평균 };
  }, [yi, isAll, isYtd]);

  /* ---------- §0 추이 ---------- */
  const 추이 = useMemo(
    () =>
      기준.연도.map((y, i) => {
        const r = ratio(종합.퇴사[i], 평균재직(i));
        return {
          연도: String(y),
          퇴직율: r == null ? null : +(r * 100).toFixed(2),
          연환산: i === 기준.진행중 && r != null ? +(((r * 365) / 기준.기간일수[i]) * 100).toFixed(2) : null,
          입사: 종합.입사[i],
          퇴사: 종합.퇴사[i],
        };
      }),
    []
  );

  /* ---------- §1 본부 ---------- */
  const 연수 = isAll ? 기준.연도.length : 1; // ALL 모드의 분모는 5개년 합이므로 연평균으로 환산해 비교
  const 본부행 = useMemo(() => {
    const rows = 본부.map((b) => ({ name: b.name, ...slice(b, yi) }));
    if (본부지표 === "rate") {
      return rows
        .filter((r) => r.rate != null && r.avg / 연수 >= 최소분모)
        .sort((a, b) => b.rate - a.rate);
    }
    return rows.filter((r) => r[본부지표] > 0).sort((a, b) => b[본부지표] - a[본부지표]);
  }, [yi, 본부지표, 연수]);
  const 본부표시 = 본부행.slice(0, 12);
  const 본부소표본 = 본부.filter((b) => {
    const r = slice(b, yi);
    return r.avg / 연수 < 최소분모;
  }).length;

  /* ---------- §2 근속 ---------- */
  const 근속연도 = useMemo(
    () =>
      기준.연도.map((y, i) => {
        const tot = sum(근속.map((r) => r.cnt[i]));
        const row = { 연도: String(y), _tot: tot };
        근속.forEach((r) => {
          row[r.name] = tot > 0 ? +((r.cnt[i] / tot) * 100).toFixed(2) : 0;
          row["_n_" + r.name] = r.cnt[i];
        });
        return row;
      }),
    []
  );

  /* ---------- §3 직급 ---------- */
  const 직급원본 = 직급뷰 === "group" ? 직급그룹 : 직급상세;
  const 직급행 = useMemo(() => {
    return 직급원본
      .map((g) => ({ name: g.name, ...slice(g, yi) }))
      .filter((r) => r.rate != null && r.avg / 연수 >= 최소분모)
      .sort((a, b) => b.rate - a.rate);
  }, [yi, 직급뷰, 직급원본, 연수]);
  const 직급제외 = 직급원본.length - 직급행.length;

  /* ---------- §4 사유 ---------- */
  const 사유행 = useMemo(() => {
    const tot = isAll ? sum(종합.퇴사) : 종합.퇴사[yi];
    return 사유
      .map((r) => {
        const c = isAll ? sum(r.cnt) : r.cnt[yi];
        return {
          name: r.name,
          kind: r.kind,
          cnt: c,
          구성비: ratio(c, tot),
          퇴직율: ratio(c, S.평재),
        };
      })
      .sort((a, b) => b.cnt - a.cnt);
  }, [yi, isAll, S.평재]);

  const 구분추이 = useMemo(
    () =>
      기준.연도.map((y, i) => {
        const row = { 연도: String(y) };
        KIND_ORDER.forEach((k) => {
          row[k] = sum(사유.filter((r) => r.kind === k).map((r) => r.cnt[i]));
        });
        return row;
      }),
    []
  );

  const 전사퇴직율 = S.퇴직율;

  /* ------------------------------------------------------------------ */
  return (
    <div
      style={{
        background: C.bg,
        minHeight: "100%",
        fontFamily: "'Inter','Pretendard','Apple SD Gothic Neo','Malgun Gothic',sans-serif",
        color: C.text,
        padding: "24px 28px 44px",
      }}
    >
      <style>{FONT_IMPORT}</style>

      {/* ── 헤더 ─────────────────────────────────────────────── */}
      <div style={{ marginBottom: 16 }}>
        <div style={{ fontSize: 20, fontWeight: 700, letterSpacing: -0.3 }}>퇴직율 분석 대시보드</div>
        <div style={{ fontSize: 12.5, color: C.sub, marginTop: 4 }}>
          퇴직율_분석_템플릿_2022-2026.xlsx 기준 · 분석 기간 {기준.연도[0]} ~ {기준.연도[기준.연도.length - 1]} ·
          데이터 기준일 {기준.기준일}
        </div>
      </div>

      {/* ── 필터 행 (모든 섹션을 한 번에 스코프) ───────────────── */}
      <div
        style={{
          display: "flex",
          justifyContent: "space-between",
          alignItems: "center",
          gap: 12,
          flexWrap: "wrap",
          marginBottom: 18,
          padding: "10px 12px",
          background: C.card,
          border: "1px solid " + C.border,
          borderRadius: 12,
        }}
      >
        <div style={{ display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap" }}>
          <span style={{ fontSize: 12, color: C.sub, fontWeight: 600 }}>분석 연도</span>
          <Segmented
            value={yi}
            onChange={setYi}
            options={[
              ...기준.연도.map((y, i) => ({ value: i, label: String(y) })),
              { value: "ALL", label: "5개년 종합" },
            ]}
          />
        </div>
        <Segmented
          size="sm"
          value={asTable}
          onChange={setAsTable}
          options={[
            { value: false, label: "차트" },
            { value: true, label: "표" },
          ]}
        />
      </div>

      {isYtd && (
        <div
          style={{
            display: "flex",
            gap: 8,
            alignItems: "center",
            background: "#FBF3E4",
            border: "1px solid #F0E0C0",
            color: "#8A5D14",
            borderRadius: 10,
            padding: "10px 14px",
            fontSize: 12.5,
            marginBottom: 16,
            lineHeight: 1.6,
          }}
        >
          <AlertTriangle size={14} style={{ flexShrink: 0 }} />
          <span>
            <b>{기준.연도[yi]}년은 진행 중인 연도</b>입니다. 기준일({기준.기준일})까지 {기준.기간일수[yi]}일간의
            실적(YTD)이므로, 전년과 비교할 때는 <b>연환산 퇴직율 {pct(S.연환산)}</b>을 함께 보세요.
          </span>
        </div>
      )}

      {/* ── KPI (종합대시보드 요약 표) ─────────────────────────── */}
      <div style={{ display: "flex", gap: 12, flexWrap: "wrap", marginBottom: 20 }}>
        <Kpi
          icon={LogOut}
          label="퇴직율"
          value={pct(S.퇴직율)}
          delta={S.prev != null && S.퇴직율 != null ? (S.퇴직율 - S.prev) * 100 : null}
          deltaGood="down"
          badge={isYtd ? "YTD" : null}
          sub={`퇴사 ${num(S.퇴사)}명 ÷ 평균 재직 ${num(S.평재, 1)}명`}
        />
        <Kpi
          icon={UserMinus}
          label="퇴사자 수"
          value={num(S.퇴사)}
          unit="명"
          sub={`입사 ${num(S.입사)}명 · 순증감 ${S.입사 - S.퇴사 >= 0 ? "+" : ""}${num(S.입사 - S.퇴사)}명`}
        />
        <Kpi
          icon={Users}
          label="평균 재직인원"
          value={num(isAll ? S.평재 / 기준.연도.length : S.평재, 1)}
          unit="명"
          sub={
            isAll
              ? "5개년 연평균 · (기초+기말)÷2"
              : `기초 ${num(종합.기초[yi])}명 → 기말 ${num(종합.기말[yi])}명`
          }
        />
        <Kpi
          icon={AlertTriangle}
          label="1년 이내 조기퇴직 비중"
          value={pct(S.조기비중)}
          sub={`퇴사자 ${num(S.퇴사)}명 중 ${num(S.조기)}명`}
        />
        <Kpi
          icon={Clock}
          label="퇴사자 평균 근속연수"
          value={S.평균근속.toFixed(1)}
          unit="년"
          sub="퇴사일 − 입사일 기준"
        />
        <Kpi
          icon={Building2}
          label="자발적 퇴직율"
          value={pct(ratio(S.자발, S.평재))}
          badge={사유_실측 ? null : "샘플"}
          sub={`자발 ${num(S.자발)} · 비자발 ${num(S.비자발)} · 기타 ${num(S.기타)}명`}
        />
      </div>

      {/* ── §0 전사 추이 ──────────────────────────────────────── */}
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(340px,1fr))", gap: 14, marginBottom: 14 }}>
        <Card>
          <SectionHead
            eyebrow="종합대시보드"
            title="전사 퇴직율 추이"
            desc="퇴사자 수 ÷ 평균 재직인원[(기초+기말)÷2]"
          />
          {asTable ? (
            <DataTable
              head={["연도", "퇴사자", "평균 재직", "퇴직율", "연환산"]}
              rows={기준.연도.map((y, i) => [
                String(y),
                num(종합.퇴사[i]),
                num(평균재직(i), 1),
                pct(ratio(종합.퇴사[i], 평균재직(i))),
                i === 기준.진행중 ? pct((ratio(종합.퇴사[i], 평균재직(i)) * 365) / 기준.기간일수[i]) : "―",
              ])}
            />
          ) : (
            <>
              <ResponsiveContainer width="100%" height={230}>
                <LineChart data={추이} margin={{ top: 14, right: 24, left: -14, bottom: 0 }}>
                  <CartesianGrid stroke={C.grid} vertical={false} />
                  <XAxis dataKey="연도" tick={{ fontSize: 11.5, fill: C.sub }} axisLine={{ stroke: C.border }} tickLine={false} />
                  <YAxis
                    tick={{ fontSize: 11, fill: C.sub }}
                    axisLine={false}
                    tickLine={false}
                    width={44}
                    unit="%"
                    domain={[0, "auto"]}
                  />
                  <Tooltip content={<ChartTooltip fmt={(v) => v.toFixed(1) + "%"} />} />
                  <Legend
                    verticalAlign="bottom"
                    height={26}
                    iconType="plainline"
                    wrapperStyle={{ fontSize: 12, color: C.sub }}
                  />
                  <Line
                    type="monotone"
                    dataKey="퇴직율"
                    name="퇴직율(실적)"
                    stroke={C.s1}
                    strokeWidth={2}
                    dot={{ r: 4, strokeWidth: 2, stroke: C.card, fill: C.s1 }}
                    activeDot={{ r: 6, strokeWidth: 2, stroke: C.card }}
                  >
                    <LabelList
                      dataKey="퇴직율"
                      position="top"
                      offset={10}
                      formatter={(v) => v.toFixed(1)}
                      style={{ fontSize: 11, fill: C.sub, fontWeight: 600 }}
                    />
                  </Line>
                  <Line
                    type="monotone"
                    dataKey="연환산"
                    name={기준.연도[기준.진행중] + " 연환산"}
                    stroke={C.s3}
                    strokeWidth={2}
                    connectNulls={false}
                    dot={{ r: 5, strokeWidth: 2, stroke: C.card, fill: C.s3 }}
                    activeDot={{ r: 7, strokeWidth: 2, stroke: C.card }}
                  >
                    <LabelList
                      dataKey="연환산"
                      position="top"
                      offset={10}
                      formatter={(v) => (v == null ? "" : v.toFixed(1))}
                      style={{ fontSize: 11, fill: C.s3, fontWeight: 700 }}
                    />
                  </Line>
                </LineChart>
              </ResponsiveContainer>
              <Note>
                {기준.연도[기준.진행중]}년은 기준일까지의 실적입니다. 주황 점은 남은 기간을 같은 속도로 가정해
                1년으로 환산한 값(템플릿 [연환산 퇴직율])입니다.
              </Note>
            </>
          )}
        </Card>

        <Card>
          <SectionHead eyebrow="종합대시보드" title="입사 · 퇴사 인원" desc="연도별 유입과 유출의 크기 비교" />
          {asTable ? (
            <DataTable
              head={["연도", "입사자", "퇴사자", "순증감"]}
              rows={기준.연도.map((y, i) => [
                String(y),
                num(종합.입사[i]),
                num(종합.퇴사[i]),
                (종합.입사[i] - 종합.퇴사[i] >= 0 ? "+" : "") + num(종합.입사[i] - 종합.퇴사[i]),
              ])}
            />
          ) : (
            <>
              <ResponsiveContainer width="100%" height={230}>
                <BarChart data={추이} margin={{ top: 14, right: 12, left: -14, bottom: 0 }} barGap={2}>
                  <CartesianGrid stroke={C.grid} vertical={false} />
                  <XAxis dataKey="연도" tick={{ fontSize: 11.5, fill: C.sub }} axisLine={{ stroke: C.border }} tickLine={false} />
                  <YAxis tick={{ fontSize: 11, fill: C.sub }} axisLine={false} tickLine={false} width={44} unit="명" />
                  <Tooltip content={<ChartTooltip unit="명" />} cursor={{ fill: "rgba(25,30,43,0.04)" }} />
                  <Legend verticalAlign="bottom" height={26} iconType="square" wrapperStyle={{ fontSize: 12, color: C.sub }} />
                  <Bar dataKey="입사" name="입사자" fill={C.good} radius={[4, 4, 0, 0]} maxBarSize={22} />
                  <Bar dataKey="퇴사" name="퇴사자" fill={C.s2} radius={[4, 4, 0, 0]} maxBarSize={22} />
                </BarChart>
              </ResponsiveContainer>
              <Note>{기준.연도[기준.진행중]}년 입사·퇴사는 기준일까지의 누계입니다.</Note>
            </>
          )}
        </Card>
      </div>

      {/* ── §1 본부별 ─────────────────────────────────────────── */}
      <Card style={{ marginBottom: 14 }}>
        <SectionHead
          eyebrow="① 본부별 퇴직율"
          title={`본부별 ${본부지표 === "rate" ? "퇴직율" : 본부지표 === "cnt" ? "퇴사자 수" : "평균 재직인원"} · ${yLabel}`}
          desc={
            본부지표 === "rate"
              ? `전사 ${yLabel} 퇴직율 ${pct(전사퇴직율)} 대비 상위 12개 본부입니다. 연평균 재직인원 ${최소분모}명 미만은 제외했습니다.`
              : "상위 12개 본부만 표시합니다. 전체는 [표] 보기에서 확인하세요."
          }
          right={
            <Segmented
              size="sm"
              value={본부지표}
              onChange={set본부지표}
              options={[
                { value: "rate", label: "퇴직율" },
                { value: "cnt", label: "퇴사자 수" },
                { value: "avg", label: "평균 재직인원" },
              ]}
            />
          }
        />

        {asTable ? (
          <DataTable
            head={["본부", "평균 재직인원", "퇴사자 수", "퇴직율"]}
            rows={본부
              .map((b) => ({ name: b.name, ...slice(b, yi) }))
              .sort((a, b) => (b.rate ?? -1) - (a.rate ?? -1))
              .map((r) => [r.name, num(r.avg, 1), num(r.cnt), pct(r.rate)])}
          />
        ) : (
          <>
            <ResponsiveContainer width="100%" height={Math.max(240, 본부표시.length * 27 + 40)}>
              <BarChart
                data={본부표시}
                layout="vertical"
                margin={{ top: 4, right: 54, left: 6, bottom: 6 }}
              >
                <CartesianGrid stroke={C.grid} horizontal={false} />
                <XAxis
                  type="number"
                  tick={{ fontSize: 11, fill: C.sub }}
                  axisLine={false}
                  tickLine={false}
                  tickFormatter={(v) => (본부지표 === "rate" ? (v * 100).toFixed(0) + "%" : num(v))}
                />
                <YAxis
                  type="category"
                  dataKey="name"
                  tick={{ fontSize: 12, fill: C.text }}
                  axisLine={false}
                  tickLine={false}
                  width={104}
                />
                <Tooltip
                  cursor={{ fill: "rgba(25,30,43,0.04)" }}
                  content={
                    <ChartTooltip
                      fmt={(v) => (본부지표 === "rate" ? pct(v) : num(v, 본부지표 === "avg" ? 1 : 0) + "명")}
                    />
                  }
                />
                {본부지표 === "rate" && 전사퇴직율 != null && (
                  <ReferenceLine
                    x={전사퇴직율}
                    stroke={C.faint}
                    strokeDasharray="4 4"
                    label={{
                      value: "전사 " + pct(전사퇴직율),
                      position: "top",
                      fill: C.sub,
                      fontSize: 11,
                      fontWeight: 600,
                    }}
                  />
                )}
                <Bar
                  dataKey={본부지표}
                  name={본부지표 === "rate" ? "퇴직율" : 본부지표 === "cnt" ? "퇴사자 수" : "평균 재직인원"}
                  radius={[0, 4, 4, 0]}
                  maxBarSize={16}
                >
                  {본부표시.map((d, i) => (
                    <Cell
                      key={i}
                      fill={본부지표 === "rate" && 전사퇴직율 != null && d.rate >= 전사퇴직율 ? C.s2 : C.s1}
                    />
                  ))}
                  <LabelList
                    dataKey={본부지표}
                    position="right"
                    offset={8}
                    formatter={(v) => (본부지표 === "rate" ? pct(v) : num(v, 본부지표 === "avg" ? 1 : 0))}
                    style={{ fontSize: 11, fill: C.sub, fontWeight: 600 }}
                  />
                </Bar>
              </BarChart>
            </ResponsiveContainer>

            {본부지표 === "rate" && (
              <div style={{ display: "flex", gap: 16, marginTop: 6, flexWrap: "wrap" }}>
                <Swatch color={C.s2} label={`전사 평균(${pct(전사퇴직율)}) 이상`} />
                <Swatch color={C.s1} label="전사 평균 미만" />
              </div>
            )}
            <Note>
              평균 재직인원이 0인 본부는 퇴직율을 산출하지 않고 “—”로 둡니다(템플릿과 동일). 여기에 더해
              연평균 재직인원 {최소분모}명 미만인 본부는 퇴직율 순위에서 제외했습니다 — 분모가 한 자릿수면
              퇴사 한두 명으로 100%를 넘겨 순위를 뒤엎기 때문입니다.
              {본부소표본 > 0 && ` 이번 화면에서는 ${본부소표본}개 본부가 이에 해당하며, [표] 보기에서는 그대로 보입니다.`}
            </Note>
          </>
        )}
      </Card>

      {/* ── §2 근속구간별 ─────────────────────────────────────── */}
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(380px,1fr))", gap: 14, marginBottom: 14 }}>
        <Card>
          <SectionHead
            eyebrow="② 근속구간별 퇴직 비중"
            title="연도별 퇴사자의 근속구간 구성"
            desc="분모는 해당 연도 퇴사자 수입니다. 구간 이름과 상한 일수는 [기준설정] I·J열을 따릅니다."
          />
          {asTable ? (
            <DataTable
              head={["근속구간", ...기준.연도.map(String), "5개년 합계", "5개년 비중"]}
              rows={근속.map((r) => [
                r.name,
                ...r.cnt.map((v) => num(v)),
                num(sum(r.cnt)),
                pct(ratio(sum(r.cnt), sum(종합.퇴사))),
              ])}
            />
          ) : (
            <>
              <ResponsiveContainer width="100%" height={252}>
                <BarChart data={근속연도} margin={{ top: 4, right: 12, left: -14, bottom: 0 }} stackOffset="expand">
                  <CartesianGrid stroke={C.grid} vertical={false} />
                  <XAxis dataKey="연도" tick={{ fontSize: 11.5, fill: C.sub }} axisLine={{ stroke: C.border }} tickLine={false} />
                  <YAxis
                    tick={{ fontSize: 11, fill: C.sub }}
                    axisLine={false}
                    tickLine={false}
                    width={44}
                    tickFormatter={(v) => Math.round(v * 100) + "%"}
                  />
                  <Tooltip
                    cursor={{ fill: "rgba(25,30,43,0.04)" }}
                    content={<ChartTooltip fmt={(v) => v.toFixed(1) + "%"} />}
                  />
                  <Legend
                    verticalAlign="bottom"
                    height={40}
                    iconType="square"
                    wrapperStyle={{ fontSize: 11.5, color: C.sub }}
                  />
                  {근속.map((r, i) => (
                    <Bar
                      key={r.name}
                      dataKey={r.name}
                      name={r.name}
                      stackId="t"
                      fill={TENURE_RAMP[i]}
                      stroke={C.card}
                      strokeWidth={2}
                      maxBarSize={46}
                    />
                  ))}
                </BarChart>
              </ResponsiveContainer>
              <Note>
                진한 색일수록 짧은 근속입니다. 위 세 구간(3개월·6개월·1년 이내)의 합이 KPI의 [1년 이내 조기퇴직
                비중]과 같습니다.
              </Note>
            </>
          )}
        </Card>

        <Card>
          <SectionHead
            eyebrow="② 표3 · 본부 × 근속구간"
            title="본부별 퇴직 근속 프로파일"
            desc="2022~2026 누계 · 해당 본부 퇴사자 중 비율. 색이 진할수록 그 구간에 퇴사가 몰려 있습니다."
          />
          <div style={{ overflowX: "auto" }}>
            <table
              style={{
                width: "100%",
                borderCollapse: "separate",
                borderSpacing: 2,
                fontSize: 11.5,
                fontVariantNumeric: "tabular-nums",
                whiteSpace: "nowrap",
              }}
            >
              <thead>
                <tr>
                  <th style={{ textAlign: "left", padding: "4px 8px", color: C.sub, fontWeight: 600 }}>본부</th>
                  {근속.map((r) => (
                    <th key={r.name} style={{ textAlign: "center", padding: "4px 6px", color: C.sub, fontWeight: 600 }}>
                      {r.name}
                    </th>
                  ))}
                  <th style={{ textAlign: "right", padding: "4px 8px", color: C.sub, fontWeight: 600 }}>누계</th>
                </tr>
              </thead>
              <tbody>
                {본부X근속.map((b) => {
                  const tot = sum(b.cnt);
                  return (
                    <tr key={b.name}>
                      <td style={{ padding: "5px 8px", color: C.text, fontWeight: 500 }}>{b.name}</td>
                      {b.cnt.map((v, i) => {
                        const share = ratio(v, tot);
                        const step = share == null ? 0 : Math.min(HEAT_RAMP.length - 1, Math.floor(share / 0.1));
                        return (
                          <td
                            key={i}
                            style={{
                              textAlign: "center",
                              padding: "5px 6px",
                              borderRadius: 4,
                              background: HEAT_RAMP[step],
                              color: step >= 3 ? "#FFFFFF" : C.text,
                              fontWeight: step >= 3 ? 600 : 400,
                            }}
                            title={`${b.name} · ${근속[i].name} · ${v}명`}
                          >
                            {share == null ? "—" : Math.round(share * 100)}
                          </td>
                        );
                      })}
                      <td style={{ textAlign: "right", padding: "5px 8px", color: C.sub }}>{num(tot)}</td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
          <div style={{ display: "flex", alignItems: "center", gap: 8, marginTop: 12, fontSize: 11.5, color: C.faint }}>
            <span>0%</span>
            {HEAT_RAMP.map((c, i) => (
              <span key={i} style={{ width: 26, height: 9, borderRadius: 2, background: c }} />
            ))}
            <span>50%+</span>
          </div>
          <Note>셀 값은 % 입니다. 누계 퇴사자가 적은 본부는 한두 명으로도 비중이 크게 튈 수 있습니다.</Note>
        </Card>
      </div>

      {/* ── §3 직급별 ─────────────────────────────────────────── */}
      <Card style={{ marginBottom: 14 }}>
        <SectionHead
          eyebrow="③ 직급별 퇴직율"
          title={`${직급뷰 === "group" ? "직급그룹" : "직급(상세)"}별 퇴직율 · ${yLabel}`}
          desc={`전사 ${yLabel} 퇴직율 ${pct(전사퇴직율)} 대비. 직급그룹 매핑은 [기준설정] D·E열을 따릅니다.`}
          right={
            <Segmented
              size="sm"
              value={직급뷰}
              onChange={set직급뷰}
              options={[
                { value: "group", label: "직급그룹" },
                { value: "detail", label: "직급 상세" },
              ]}
            />
          }
        />
        {asTable ? (
          <DataTable
            head={[직급뷰 === "group" ? "직급그룹" : "직급", "평균 재직인원", "퇴사자 수", "퇴직율"]}
            rows={직급원본
              .map((g) => ({ name: g.name, ...slice(g, yi) }))
              .sort((a, b) => (b.rate ?? -1) - (a.rate ?? -1))
              .map((r) => [r.name, num(r.avg, 1), num(r.cnt), pct(r.rate)])}
          />
        ) : (
          <>
            <ResponsiveContainer width="100%" height={Math.max(220, 직급행.length * 27 + 40)}>
              <BarChart data={직급행} layout="vertical" margin={{ top: 4, right: 54, left: 6, bottom: 6 }}>
                <CartesianGrid stroke={C.grid} horizontal={false} />
                <XAxis
                  type="number"
                  tick={{ fontSize: 11, fill: C.sub }}
                  axisLine={false}
                  tickLine={false}
                  tickFormatter={(v) => (v * 100).toFixed(0) + "%"}
                />
                <YAxis
                  type="category"
                  dataKey="name"
                  tick={{ fontSize: 12, fill: C.text }}
                  axisLine={false}
                  tickLine={false}
                  width={78}
                />
                <Tooltip cursor={{ fill: "rgba(25,30,43,0.04)" }} content={<ChartTooltip fmt={(v) => pct(v)} />} />
                {전사퇴직율 != null && (
                  <ReferenceLine
                    x={전사퇴직율}
                    stroke={C.faint}
                    strokeDasharray="4 4"
                    label={{ value: "전사 " + pct(전사퇴직율), position: "top", fill: C.sub, fontSize: 11, fontWeight: 600 }}
                  />
                )}
                <Bar dataKey="rate" name="퇴직율" radius={[0, 4, 4, 0]} maxBarSize={16}>
                  {직급행.map((d, i) => (
                    <Cell key={i} fill={전사퇴직율 != null && d.rate >= 전사퇴직율 ? C.s2 : C.s1} />
                  ))}
                  <LabelList
                    dataKey="rate"
                    position="right"
                    offset={8}
                    formatter={(v) => pct(v)}
                    style={{ fontSize: 11, fill: C.sub, fontWeight: 600 }}
                  />
                </Bar>
              </BarChart>
            </ResponsiveContainer>
            <div style={{ display: "flex", gap: 16, marginTop: 6, flexWrap: "wrap" }}>
              <Swatch color={C.s2} label={`전사 평균(${pct(전사퇴직율)}) 이상`} />
              <Swatch color={C.s1} label="전사 평균 미만" />
            </div>
            <Note>
              연평균 재직인원 {최소분모}명 미만인 직급은 순위에서 제외했습니다.
              {직급제외 > 0 && ` 이번 화면에서는 ${직급제외}개 직급이 이에 해당하며, [표] 보기에서는 그대로 보입니다.`}
              {" "}폐지 중인 직급(예: 판매사원)은 재직자가 빠져나가며 분모가 먼저 사라지므로, 퇴직율만 보지 말고
              [퇴사자 수]를 함께 확인하세요.
            </Note>
          </>
        )}
      </Card>

      {/* ── §4 퇴사사유별 ─────────────────────────────────────── */}
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(380px,1fr))", gap: 14 }}>
        <Card>
          <SectionHead
            eyebrow="④ 퇴사사유별 퇴직율"
            title={`퇴사사유별 인원 · ${yLabel}`}
            desc="사유 목록과 자발/비자발 구분은 [기준설정] L·M열을 따릅니다."
          />
          {!사유_실측 && (
            <SampleWarning>
              RawData [퇴사사유]가 비어 있어 <b>이 섹션과 [자발적 퇴직율] KPI는 샘플 값</b>입니다. 사유를 입력한 뒤
              생성기를 다시 돌리면 실측으로 바뀝니다.
            </SampleWarning>
          )}
          {asTable ? (
            <DataTable
              head={["퇴사사유", "사유구분", "퇴사자 수", "구성비", "퇴직율"]}
              rows={사유행.map((r) => [r.name, r.kind, num(r.cnt), pct(r.구성비), pct(r.퇴직율, 2)])}
              align={["right", "left"]}
            />
          ) : (
            <>
              <ResponsiveContainer width="100%" height={Math.max(240, 사유행.length * 25 + 40)}>
                <BarChart data={사유행} layout="vertical" margin={{ top: 4, right: 66, left: 6, bottom: 6 }}>
                  <CartesianGrid stroke={C.grid} horizontal={false} />
                  <XAxis type="number" tick={{ fontSize: 11, fill: C.sub }} axisLine={false} tickLine={false} />
                  <YAxis
                    type="category"
                    dataKey="name"
                    tick={{ fontSize: 11.5, fill: C.text }}
                    axisLine={false}
                    tickLine={false}
                    width={126}
                  />
                  <Tooltip cursor={{ fill: "rgba(25,30,43,0.04)" }} content={<ChartTooltip unit="명" />} />
                  <Bar dataKey="cnt" name="퇴사자 수" radius={[0, 4, 4, 0]} maxBarSize={14}>
                    {사유행.map((d, i) => (
                      <Cell key={i} fill={KIND_COLOR[d.kind]} />
                    ))}
                    <LabelList
                      dataKey="cnt"
                      position="right"
                      offset={8}
                      formatter={(v) => num(v) + "명"}
                      style={{ fontSize: 11, fill: C.sub, fontWeight: 600 }}
                    />
                  </Bar>
                </BarChart>
              </ResponsiveContainer>
              <div style={{ display: "flex", gap: 16, marginTop: 4, flexWrap: "wrap" }}>
                {KIND_ORDER.map((k) => (
                  <Swatch key={k} color={KIND_COLOR[k]} label={k} />
                ))}
              </div>
            </>
          )}
          <Note>
            구성비의 분모는 <b>퇴사자 수</b>, 사유별 퇴직율의 분모는 <b>전사 평균 재직인원</b>입니다(템플릿 표2·표3).
            두 값을 섞어 읽지 마세요.
          </Note>
        </Card>

        <Card>
          <SectionHead
            eyebrow="④ 표4 · 자발 / 비자발"
            title="사유구분별 퇴사자 추이"
            desc="자발적 퇴직은 조직이 개선할 여지가 있는 이탈, 비자발적 퇴직은 회사 결정에 따른 이탈입니다."
          />
          {!사유_실측 && <SampleWarning>샘플 값입니다. 실측이 아닙니다.</SampleWarning>}
          {asTable ? (
            <DataTable
              head={["사유구분", ...기준.연도.map(String), "5개년 합계"]}
              rows={[
                ...KIND_ORDER.map((k) => [
                  k,
                  ...기준.연도.map((_, i) => num(구분추이[i][k])),
                  num(sum(기준.연도.map((_, i) => 구분추이[i][k]))),
                ]),
                ["합계", ...종합.퇴사.map((v) => num(v)), num(sum(종합.퇴사))],
              ]}
            />
          ) : (
            <>
              <ResponsiveContainer width="100%" height={252}>
                <BarChart data={구분추이} margin={{ top: 4, right: 12, left: -14, bottom: 0 }}>
                  <CartesianGrid stroke={C.grid} vertical={false} />
                  <XAxis dataKey="연도" tick={{ fontSize: 11.5, fill: C.sub }} axisLine={{ stroke: C.border }} tickLine={false} />
                  <YAxis tick={{ fontSize: 11, fill: C.sub }} axisLine={false} tickLine={false} width={44} unit="명" />
                  <Tooltip cursor={{ fill: "rgba(25,30,43,0.04)" }} content={<ChartTooltip unit="명" />} />
                  <Legend verticalAlign="bottom" height={30} iconType="square" wrapperStyle={{ fontSize: 12, color: C.sub }} />
                  {KIND_ORDER.map((k) => (
                    <Bar
                      key={k}
                      dataKey={k}
                      name={k}
                      stackId="k"
                      fill={KIND_COLOR[k]}
                      stroke={C.card}
                      strokeWidth={2}
                      maxBarSize={46}
                    />
                  ))}
                </BarChart>
              </ResponsiveContainer>
              <Note>
                자발적 퇴직율({pct(ratio(S.자발, S.평재))})은 자발적 퇴사자 ÷ 전사 평균 재직인원입니다.
              </Note>
            </>
          )}
        </Card>
      </div>

      {/* ── 각주 (템플릿 [0_안내] 4항과 동일) ─────────────────── */}
      <div
        style={{
          marginTop: 18,
          padding: "16px 18px",
          background: C.card,
          border: "1px solid " + C.border,
          borderRadius: 12,
          fontSize: 11.5,
          color: C.sub,
          lineHeight: 1.85,
        }}
      >
        <div style={{ fontWeight: 600, color: C.text, fontSize: 12.5, marginBottom: 6 }}>계산 기준</div>
        <div>· 퇴직율 = 해당 연도 퇴사자 수 ÷ 평균 재직인원</div>
        <div>· 평균 재직인원 = (기초 재직인원 + 기말 재직인원) ÷ 2</div>
        <div>· 기초 재직인원 = 1월 1일 시점 재직자 / 기말 재직인원 = 12월 31일(또는 데이터 기준일) 시점 재직자</div>
        <div>· 근속일수 = 퇴사자는 [퇴사일 − 입사일], 재직자는 [기준일 − 입사일]</div>
        <div>· 구성비의 분모는 퇴사자 수, 사유별 퇴직율의 분모는 전사 평균 재직인원입니다.</div>
        <div>· 분모(평균 재직인원 또는 퇴사자 수)가 0이면 값을 산출하지 않고 “—”로 표시합니다.</div>
        <div style={{ marginTop: 8, color: C.faint }}>
          · [5개년 종합]의 퇴직율은 누계 퇴사자 ÷ 연도별 평균 재직인원의 합입니다. 조기퇴직 비중도 누계 기준으로
          계산합니다(엑셀 템플릿 종합대시보드 H21은 연도별 비중의 단순평균이라 값이 조금 다를 수 있습니다).
        </div>
        <div style={{ marginTop: 8, color: C.faint }}>
          · 데이터 출처 : {출처.파일} · {출처.원본행수.toLocaleString("ko-KR")}행 · 생성 {출처.생성시각}
        </div>
        <div style={{ color: C.faint }}>· {출처.본부주석}</div>
        {!사유_실측 && (
          <div style={{ color: C.faint }}>
            · ④ 퇴사사유와 [자발적 퇴직율] KPI는 <b>샘플 값</b>입니다. RawData의 퇴사사유가 비어 있습니다.
          </div>
        )}
      </div>

      <div style={{ marginTop: 14, fontSize: 11.5, color: C.faint, textAlign: "center" }}>
        프로토타입 · 데이터 연동 시 상단 데이터 블록만 교체하면 됩니다
      </div>
    </div>
  );
}
