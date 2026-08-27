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
   퇴직율 분석 대시보드
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
   딥 티일 + 라임 팔레트. 강조는 티일 한 가지로 모으고, 라임은 넓은 면적에만 씁니다.
   시리즈 색은 색각이상(CVD) 분리도·대비 검증을 통과한 조합만 사용합니다.
     범주형 3슬롯 : #12514B / #C1443A / #C98A2C  (자발적 / 비자발적 / 기타)
     순서형 6단계 : 근속구간 램프 (티일→라임, 명도 단조 변화)
--------------------------------------------------------------------- */
const C = {
  bg: "#EFF1F6",
  card: "#FFFFFF",
  border: "#E7EAF1",
  grid: "#EFF2F7",
  text: "#0F1B22",
  sub: "#68737F",   // 본문 대비 4.83:1
  faint: "#848E9C", // 보조 문구 — 카드 위 3.32:1

  s1: "#12514B", // 범주 1 · 기본 시리즈 / 자발적 (딥 티일)
  s2: "#C1443A", // 범주 2 · 퇴사 / 비자발적 / 평균 초과
  s3: "#C98A2C", // 범주 3 · 기타 (대비 2.93:1 → 반드시 직접 라벨 또는 표와 함께 노출)
  good: "#6E9B24", // 입사 등 반대 방향 지표 (올리브 라임 — 작은 글씨에도 읽히는 명도)

  accent: "#12514B",      // 활성 버튼·아이콘 칩 등 강조
  accentSoft: "#E9F1EF",  // 강조의 옅은 배경
  lime: "#C3E14C",        // 하이라이트 — 큰 면적 채움 전용 (작은 글씨 금지)
  limeSoft: "#F3F8E3",
  warn: "#8A5D14",        // 주의 스트립
  warnBg: "#FBF3E4",
  warnBorder: "#F0E0C0",
  tipBg: "#102420",       // 차트 툴팁
  tipText: "#D6E1DE",
  hover: "rgba(18,81,75,0.05)",
};

/* 근속구간 순서형 램프 — 근속이 짧을수록 진하게 (조기 이탈이 무겁게 읽히도록) */
const TENURE_RAMP = ["#0B3B37", "#12514B", "#2E7B6E", "#5E9E70", "#96BE55", "#C3E14C"];
/* 교차표 히트맵 램프 — 단일 티일 계열, 0에 가까울수록 배경으로 후퇴 */
const HEAT_RAMP = ["#F3F7F6", "#DCEBE8", "#B9D8D2", "#8CBFB6", "#2E8177", "#12514B"];
/* 이 단계부터 셀 배경이 충분히 어두워 흰 글씨가 읽힙니다(대비 4.5:1 이상). */
const HEAT_INVERT = 4;

const FONT_IMPORT =
  "@import url('https://cdn.jsdelivr.net/gh/orioncactus/pretendard@v1.3.9/dist/web/static/pretendard.min.css');";
const FONT_STACK =
  "'Pretendard','Pretendard Variable',-apple-system,'Apple SD Gothic Neo','Malgun Gothic',sans-serif";

/* ---------------------------------------------------------------------
   2. 데이터 — 템플릿 워크북에서 생성한 집계값

   ⚠ 아래 GENERATED-DATA 구간은 생성기가 통째로 덮어씁니다. 직접 고치지 마세요.
      갱신:  node tools/build-dashboard-data.mjs 퇴직율_분석_raw.xlsx
      옵션:  --기준일 2026-08-25  --본부최소 5  --keep-reasons
--------------------------------------------------------------------- */
/* <<< GENERATED-DATA-START — 이 아래는 생성기가 덮어씁니다 >>> */
const 기준 = {
  기준일: "2026-08-27",
  연도: [2024, 2025, 2026],
  기간일수: [366, 365, 239], // 기말일 − 기초일 + 1
  진행중: 2, // 기준일까지만 집계된 연도 인덱스 (YTD, 없으면 -1)
};

/* [종합대시보드] 요약 표 */
const 종합 = {
  기초: [700, 737, 694],
  기말: [738, 695, 702],
  입사: [165, 153, 89],
  퇴사: [129, 199, 76],
  평균근속년: [3.69, 4.06, 4.06],
};

/* [①] 본부 × 연도 — avg = 평균 재직인원(명), cnt = 퇴사자 수(명) */
const 본부 = [
  { name: "소싱",       avg: [153.5, 150, 137], cnt: [31, 36, 8] },
  { name: "경영지원",   avg: [101, 99, 96], cnt: [14, 29, 10] },
  { name: "수출",       avg: [94.5, 88, 85.5], cnt: [15, 17, 7] },
  { name: "TT상품기획", avg: [86.5, 87.5, 79.5], cnt: [8, 32, 14] },
  { name: "TT운영",     avg: [69, 61, 62.5], cnt: [22, 43, 12] },
  { name: "FS상품기획", avg: [50, 60, 59], cnt: [3, 6, 5] },
  { name: "E-Biz",      avg: [50, 53, 57.5], cnt: [7, 6, 8] },
  { name: "물류",       avg: [50.5, 48, 46], cnt: [8, 6, 4] },
  { name: "FS운영",     avg: [29.5, 28, 32], cnt: [17, 10, 4] },
  { name: "TT마케팅",   avg: [17.5, 21.5, 20], cnt: [1, 12, 4] },
  { name: "유통",       avg: [16.5, 15.5, 15], cnt: [2, 2, 0] },
  { name: "FS마케팅",   avg: [0.5, 4.5, 8], cnt: [1, 0, 0] },
];

/* [②] 근속구간 × 연도 퇴사자 수 — 순서는 기준설정 I열 순서 그대로 */
const 근속 = [
  { name: "3개월 이내", cnt: [9, 25, 6] },
  { name: "6개월 이내", cnt: [12, 15, 7] },
  { name: "1년 이내",   cnt: [9, 12, 2] },
  { name: "1~3년",      cnt: [49, 64, 27] },
  { name: "3~5년",      cnt: [15, 36, 16] },
  { name: "5년 이상",   cnt: [35, 47, 18] },
];
const 조기구간수 = 3; // 앞 3개 구간 = "1년 이내 조기퇴직"

/* 퇴직율 순위 차트에서 제외할 소표본 기준 (연평균 재직인원, 명).
   분모가 한 자릿수면 한두 명의 퇴사로 퇴직율이 100%를 넘어 순위가 무의미해집니다.
   제외된 조직·직급도 [표] 보기에서는 그대로 보입니다. */
const 최소분모 = 10;

/* [②-표3] 본부 × 근속구간 — cnt[근속구간 인덱스][연도 인덱스] · 근속구간 순서는 위 [근속] 배열과 같습니다 */
const 본부X근속 = [
  { name: "TT운영", cnt: [[2, 6, 1], [3, 4, 2], [1, 2, 2], [5, 16, 4], [3, 7, 0], [8, 8, 3]] },
  { name: "소싱", cnt: [[4, 3, 0], [2, 2, 0], [1, 2, 0], [19, 16, 2], [2, 5, 4], [3, 8, 2]] },
  { name: "TT상품기획", cnt: [[0, 6, 1], [5, 3, 4], [0, 1, 0], [0, 12, 5], [1, 4, 3], [2, 6, 1]] },
  { name: "경영지원", cnt: [[0, 7, 2], [2, 4, 0], [1, 1, 0], [2, 5, 3], [0, 6, 2], [9, 6, 3]] },
  { name: "수출", cnt: [[0, 0, 0], [0, 1, 0], [2, 2, 0], [4, 2, 3], [4, 4, 0], [5, 8, 4]] },
  { name: "FS운영", cnt: [[1, 1, 0], [0, 0, 0], [3, 0, 0], [8, 2, 3], [1, 3, 1], [4, 4, 0]] },
  { name: "E-Biz", cnt: [[1, 0, 1], [0, 0, 0], [0, 0, 0], [4, 2, 6], [1, 2, 1], [1, 2, 0]] },
  { name: "물류", cnt: [[0, 0, 0], [0, 0, 0], [0, 1, 0], [4, 2, 0], [3, 1, 1], [1, 2, 3]] },
  { name: "TT마케팅", cnt: [[0, 2, 0], [0, 0, 1], [0, 2, 0], [1, 4, 0], [0, 2, 2], [0, 2, 1]] },
  { name: "FS상품기획", cnt: [[0, 0, 1], [0, 1, 0], [1, 1, 0], [1, 2, 1], [0, 2, 2], [1, 0, 1]] },
  { name: "유통", cnt: [[0, 0, 0], [0, 0, 0], [0, 0, 0], [1, 1, 0], [0, 0, 0], [1, 1, 0]] },
  { name: "FS마케팅", cnt: [[1, 0, 0], [0, 0, 0], [0, 0, 0], [0, 0, 0], [0, 0, 0], [0, 0, 0]] },
];

/* [③-표1] 직급그룹 × 연도 */
const 직급그룹 = [
  { name: "임원",   avg: [0.5, 0, 0], cnt: [1, 0, 0] },
  { name: "부장급", avg: [68, 66, 64.5], cnt: [7, 17, 5] },
  { name: "차장급", avg: [112, 105.5, 96.5], cnt: [10, 19, 10] },
  { name: "과장급", avg: [202.5, 199, 190], cnt: [20, 38, 18] },
  { name: "대리급", avg: [207.5, 199, 189], cnt: [45, 53, 20] },
  { name: "사원급", avg: [128, 143.5, 157.5], cnt: [41, 64, 23] },
  { name: "인턴",   avg: [0.5, 3, 0.5], cnt: [5, 8, 0] },
];

/* [③-표4] 직급 상세 × 연도 */
const 직급상세 = [
  { name: "대리",     avg: [207.5, 199, 189], cnt: [45, 53, 20] },
  { name: "사원",     avg: [110, 128, 140.5], cnt: [33, 60, 21] },
  { name: "과장",     avg: [202.5, 199, 190], cnt: [20, 38, 18] },
  { name: "차장",     avg: [112, 105.5, 96.5], cnt: [10, 19, 10] },
  { name: "부장",     avg: [42.5, 40, 41], cnt: [3, 11, 3] },
  { name: "인턴",     avg: [0.5, 3, 0.5], cnt: [5, 8, 0] },
  { name: "사원 4을", avg: [15, 12, 11.5], cnt: [6, 2, 2] },
  { name: "실장",     avg: [14, 13.5, 13], cnt: [4, 3, 1] },
  { name: "공장장",   avg: [11.5, 12.5, 10.5], cnt: [0, 3, 1] },
  { name: "사원 5급", avg: [3, 3.5, 5.5], cnt: [2, 2, 0] },
  { name: "상무보",   avg: [0.5, 0, 0], cnt: [1, 0, 0] },
];

/* [④] 퇴사사유 × 연도 — kind(사유구분)는 [기준설정] M열, 없으면 생성기의 기본_사유구분 */
const 사유 = [
  { name: "처우불만(급여/보상/복리후생)", kind: "자발적", cnt: [0, 17, 8] },
  { name: "워라밸불만", kind: "자발적", cnt: [0, 3, 2] },
  { name: "경력개발/성장기회부족", kind: "자발적", cnt: [0, 26, 11] },
  { name: "직무불만/업직종변경희망", kind: "자발적", cnt: [0, 45, 12] },
  { name: "리더십불만", kind: "자발적", cnt: [0, 14, 4] },
  { name: "조직문화불만", kind: "자발적", cnt: [0, 6, 8] },
  { name: "대인갈등", kind: "자발적", cnt: [0, 3, 0] },
  { name: "출퇴근/근무지문제", kind: "자발적", cnt: [0, 1, 2] },
  { name: "원거리발령", kind: "자발적", cnt: [0, 1, 0] },
  { name: "출산/육아", kind: "자발적", cnt: [0, 2, 4] },
  { name: "건강/휴식", kind: "자발적", cnt: [0, 16, 4] },
  { name: "이주/가족돌봄", kind: "자발적", cnt: [0, 5, 6] },
  { name: "권고사직", kind: "비자발적", cnt: [0, 40, 11] },
  { name: "계약만료/수습종료/정년퇴직", kind: "비자발적", cnt: [0, 13, 3] },
  { name: "사간전출", kind: "비자발적", cnt: [0, 7, 1] },
];

/* [④-표5] 본부 × 퇴사사유 — cnt[사유 인덱스][연도 인덱스] · 사유 순서는 위 [사유] 배열과 같습니다 */
const 본부X사유 = [
  { name: "TT운영", cnt: [[0, 2, 2], [0, 2, 0], [0, 4, 2], [0, 10, 2], [0, 1, 1], [0, 5, 1], [0, 1, 0], [0, 0, 0], [0, 0, 0], [0, 0, 0], [0, 1, 1], [0, 1, 1], [0, 15, 1], [0, 1, 0], [0, 0, 1]] },
  { name: "소싱", cnt: [[0, 5, 1], [0, 1, 0], [0, 8, 2], [0, 9, 2], [0, 3, 1], [0, 1, 0], [0, 2, 0], [0, 0, 0], [0, 1, 0], [0, 0, 1], [0, 2, 0], [0, 0, 0], [0, 1, 1], [0, 1, 0], [0, 2, 0]] },
  { name: "TT상품기획", cnt: [[0, 1, 3], [0, 0, 0], [0, 3, 1], [0, 6, 1], [0, 2, 1], [0, 0, 2], [0, 0, 0], [0, 0, 0], [0, 0, 0], [0, 0, 0], [0, 0, 1], [0, 1, 2], [0, 12, 2], [0, 5, 1], [0, 2, 0]] },
  { name: "경영지원", cnt: [[0, 2, 0], [0, 0, 0], [0, 1, 3], [0, 4, 3], [0, 7, 1], [0, 0, 0], [0, 0, 0], [0, 1, 0], [0, 0, 0], [0, 0, 0], [0, 8, 0], [0, 0, 1], [0, 2, 1], [0, 2, 1], [0, 2, 0]] },
  { name: "수출", cnt: [[0, 0, 0], [0, 0, 1], [0, 3, 0], [0, 3, 1], [0, 1, 0], [0, 0, 1], [0, 0, 0], [0, 0, 0], [0, 0, 0], [0, 1, 2], [0, 2, 0], [0, 2, 0], [0, 5, 2], [0, 0, 0], [0, 0, 0]] },
  { name: "FS운영", cnt: [[0, 2, 0], [0, 0, 0], [0, 2, 1], [0, 3, 2], [0, 0, 0], [0, 0, 0], [0, 0, 0], [0, 0, 0], [0, 0, 0], [0, 0, 0], [0, 0, 0], [0, 0, 1], [0, 2, 0], [0, 1, 0], [0, 0, 0]] },
  { name: "E-Biz", cnt: [[0, 2, 0], [0, 0, 1], [0, 1, 2], [0, 2, 1], [0, 0, 0], [0, 0, 2], [0, 0, 0], [0, 0, 1], [0, 0, 0], [0, 0, 0], [0, 0, 0], [0, 1, 0], [0, 0, 1], [0, 0, 0], [0, 0, 0]] },
  { name: "물류", cnt: [[0, 3, 1], [0, 0, 0], [0, 1, 0], [0, 2, 0], [0, 0, 0], [0, 0, 0], [0, 0, 0], [0, 0, 0], [0, 0, 0], [0, 0, 0], [0, 0, 1], [0, 0, 0], [0, 0, 1], [0, 0, 1], [0, 0, 0]] },
  { name: "TT마케팅", cnt: [[0, 0, 1], [0, 0, 0], [0, 1, 0], [0, 5, 0], [0, 0, 0], [0, 0, 1], [0, 0, 0], [0, 0, 0], [0, 0, 0], [0, 0, 1], [0, 1, 1], [0, 0, 0], [0, 3, 0], [0, 2, 0], [0, 0, 0]] },
  { name: "FS상품기획", cnt: [[0, 0, 0], [0, 0, 0], [0, 2, 0], [0, 1, 0], [0, 0, 0], [0, 0, 1], [0, 0, 0], [0, 0, 1], [0, 0, 0], [0, 1, 0], [0, 1, 0], [0, 0, 1], [0, 0, 2], [0, 0, 0], [0, 1, 0]] },
  { name: "유통", cnt: [[0, 0, 0], [0, 0, 0], [0, 0, 0], [0, 0, 0], [0, 0, 0], [0, 0, 0], [0, 0, 0], [0, 0, 0], [0, 0, 0], [0, 0, 0], [0, 1, 0], [0, 0, 0], [0, 0, 0], [0, 1, 0], [0, 0, 0]] },
  { name: "FS마케팅", cnt: [[0, 0, 0], [0, 0, 0], [0, 0, 0], [0, 0, 0], [0, 0, 0], [0, 0, 0], [0, 0, 0], [0, 0, 0], [0, 0, 0], [0, 0, 0], [0, 0, 0], [0, 0, 0], [0, 0, 0], [0, 0, 0], [0, 0, 0]] },
];
/* RawData [퇴사사유]가 실제로 입력되어 있으면 true. false면 §4와 자발/비자발 KPI는
   가상 분포이므로 화면에 경고를 띄웁니다. 생성기가 자동으로 갱신합니다. */
const 사유_실측 = true;
/* 집계 원본 정보 — 생성기가 채웁니다. */
const 출처 = {
  파일: "퇴직율_분석_raw.xlsx",
  생성시각: "2026-08-27 01:51:37",
  원본행수: 1109,
  본부주석: "이름이 정상이고 연평균 재직인원 5명 이상인 본부만 개별 표시하고, 나머지는 [기타(미매핑)]로 합산합니다.",
};
/* <<< GENERATED-DATA-END >>> */

const KIND_COLOR = { 자발적: C.s1, 비자발적: C.s2, 기타: C.s3 };
const KIND_ORDER = ["자발적", "비자발적", "기타"];

/* 분석 기간 라벨 — 연도 배열 길이에서 파생하므로 기간이 바뀌어도 문구가 따라갑니다. */
const 연도수 = 기준.연도.length;
const 전체라벨 = 연도수 + "개년 종합";
const 전체기간 = 기준.연도[0] + "~" + 기준.연도[연도수 - 1];

/* §4 퇴사사유 집계 구간 — 사유가 입력된 최근 N개년만 봅니다.
   사유.cnt 배열은 기준.연도와 같은 길이로 두고, 화면에서만 이 구간으로 잘라 씁니다.
   집계 연도를 늘리려면 아래 숫자만 바꾸면 됩니다. */
/* 본부 표시 순서 — 조직도 순서로 고정해 어느 표·차트에서나 같은 자리에 오게 합니다.
   이 목록에 없는 본부(신설·미매핑 등)는 뒤에 원래 순서대로 붙습니다. */
const 본부순서 = [
  "경영지원", "소싱", "수출", "E-Biz", "TT상품기획", "TT운영",
  "TT마케팅", "FS상품기획", "FS운영", "FS마케팅", "물류", "유통",
];
const 본부자리 = (name) => {
  const i = 본부순서.findIndex((v) => v.toLowerCase() === String(name).toLowerCase());
  return i < 0 ? 본부순서.length : i;
};
/** 본부 행 목록을 조직도 순서로 다시 세웁니다(정렬은 안정적이라 목록 밖 본부의 상대 순서는 유지). */
const 본부정렬 = (rows) => [...rows].sort((a, b) => 본부자리(a.name) - 본부자리(b.name));

const 사유연도수 = Math.min(2, 연도수);
const 사유시작 = 연도수 - 사유연도수;               // 전체 연도 배열에서 사유 집계가 시작되는 인덱스
const 사유연도 = 기준.연도.slice(사유시작);
const 사유범위 = 사유연도[0] + "~" + 사유연도[사유연도수 - 1];

/* ---------------------------------------------------------------------
   3. 계산 헬퍼 — 템플릿 수식과 동일한 규칙
--------------------------------------------------------------------- */
const sum = (a) => a.reduce((x, y) => x + y, 0);
/** IF(분모=0,"",분자/분모) — 분모가 0이면 값 없음(null) */
const ratio = (n, d) => (d > 0 ? n / d : null);
const 평균재직 = (i) => (종합.기초[i] + 종합.기말[i]) / 2;

const pct = (v, d = 1) => (v == null ? "—" : (v * 100).toFixed(d) + "%");
/** 교차표 머리글용 짧은 이름 — 괄호 앞까지만 쓰고, 길면 말줄임 */
const 짧은이름 = (s, n = 7) => {
  const t = String(s).split("(")[0].trim();
  return t.length > n ? t.slice(0, n) + "…" : t;
};
const num = (v, d = 0) =>
  v == null ? "—" : v.toLocaleString("ko-KR", { minimumFractionDigits: d, maximumFractionDigits: d });

/** 연도 인덱스 또는 "ALL"(전체 기간 종합) 기준으로 한 행을 집계 */
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
        borderRadius: 18,
        padding: "20px 22px 18px",
        boxShadow: "0 1px 2px rgba(15,27,34,0.04)",
        ...style,
      }}
    >
      {children}
    </div>
  );
}

function Segmented({ options, value, onChange, size = "md" }) {
  const pad = size === "sm" ? "5px 12px" : "6px 15px";
  return (
    <div style={{ display: "flex", gap: 3, background: C.bg, borderRadius: 999, padding: 3 }}>
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
              fontFamily: "inherit",
              borderRadius: 999,
              border: "none",
              cursor: "pointer",
              whiteSpace: "nowrap",
              background: on ? C.accent : "transparent",
              color: on ? "#FFFFFF" : C.sub,
              boxShadow: on ? "0 1px 3px rgba(15,27,34,0.16)" : "none",
              transition: "background 120ms ease, color 120ms ease",
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
          <div style={{ fontSize: 10.5, fontWeight: 700, color: C.accent, letterSpacing: 0.5, marginBottom: 4, opacity: 0.75 }}>
            {eyebrow}
          </div>
        )}
        <div style={{ fontSize: 15.5, fontWeight: 700, color: C.text, letterSpacing: -0.2 }}>{title}</div>
        {desc && <div style={{ fontSize: 12, color: C.sub, marginTop: 3, lineHeight: 1.55 }}>{desc}</div>}
      </div>
      {right}
    </div>
  );
}

/** 교차표 한 칸 — 인원(명)이 주 값, 비중(%)은 괄호로 덧붙이고 색은 비중을 따릅니다. */
function HeatCell({ v, share, title, strong = false, style }) {
  const step = share == null ? 0 : Math.min(HEAT_RAMP.length - 1, Math.floor(share / 0.1));
  const 흰글씨 = step >= HEAT_INVERT;
  return (
    <td
      title={title}
      style={{
        textAlign: "center",
        padding: strong ? "3px 5px" : "2px 5px",
        borderRadius: 4,
        background: HEAT_RAMP[step],
        color: 흰글씨 ? "#FFFFFF" : C.text,
        lineHeight: 1.5,
        ...style,
      }}
    >
      {v ? (
        <>
          <span style={{ fontWeight: strong || 흰글씨 ? 700 : 600 }}>{num(v)}</span>
          <span style={{ fontSize: 9.5, fontWeight: 500, marginLeft: 2, opacity: 흰글씨 ? 0.85 : 0.6 }}>
            ({share == null ? "—" : Math.round(share * 100)}%)
          </span>
        </>
      ) : (
        <span style={{ color: C.faint }}>—</span>
      )}
    </td>
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
        background: C.warnBg, border: "1px solid " + C.warnBorder, color: C.warn,
        borderRadius: 10, padding: "9px 12px", fontSize: 12, lineHeight: 1.6, marginBottom: 14,
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
        background: C.tipBg,
        borderRadius: 10,
        padding: "10px 12px",
        fontSize: 12,
        color: "#fff",
        fontFamily: FONT_STACK,
        boxShadow: "0 8px 24px rgba(15,27,34,0.22)",
        minWidth: 150,
      }}
    >
      <div style={{ fontWeight: 600, marginBottom: 6 }}>{label}</div>
      {payload
        .filter((p) => p.value != null)
        .map((p, i) => (
          <div key={i} style={{ display: "flex", justifyContent: "space-between", gap: 18, color: C.tipText, lineHeight: 1.7 }}>
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
                  padding: "9px 10px",
                  background: C.bg,
                  borderBottom: "1px solid " + C.border,
                  borderTopLeftRadius: i === 0 ? 8 : 0,
                  borderTopRightRadius: i === head.length - 1 ? 8 : 0,
                  color: C.sub,
                  fontWeight: 700,
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
            width: 30,
            height: 30,
            borderRadius: 10,
            background: C.accentSoft,
            display: "flex",
            alignItems: "center",
            justifyContent: "center",
            flexShrink: 0,
          }}
        >
          <Icon size={15} color={C.accent} />
        </div>
        <div style={{ fontSize: 12.5, color: C.sub, fontWeight: 600 }}>{label}</div>
      </div>

      <div style={{ display: "flex", alignItems: "baseline", gap: 4 }}>
        {/* 히어로 수치는 본문과 같은 sans + 비례숫자 (tabular-nums 미사용) */}
        <span style={{ fontSize: 31, fontWeight: 700, color: C.text, letterSpacing: -0.9 }}>{value}</span>
        {unit && <span style={{ fontSize: 13.5, color: C.sub, fontWeight: 500 }}>{unit}</span>}
        {badge && (
          <span
            style={{
              marginLeft: 4,
              fontSize: 10.5,
              fontWeight: 700,
              color: C.s3,
              background: C.warnBg,
              borderRadius: 999,
              padding: "2px 8px",
            }}
          >
            {badge}
          </span>
        )}
      </div>

      <div style={{ marginTop: 8, fontSize: 12, color: C.sub, lineHeight: 1.5, minHeight: 18 }}>
        {hasDelta && (
          <span
            style={{
              display: "inline-block",
              color: deltaColor,
              background: delta === 0 ? C.bg : good ? C.limeSoft : C.warnBg,
              borderRadius: 999,
              padding: "1px 7px",
              fontWeight: 700,
              fontSize: 11,
              marginRight: 6,
            }}
          >
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
  const [yi, setYi] = useState(연도수 - 1);    // 숫자 = 연도 인덱스, "ALL" = 전체 기간 종합
  const [asTable, setAsTable] = useState(false); // 전체 섹션 표 보기
  const [본부지표, set본부지표] = useState("rate"); // rate | cnt | avg
  const [직급뷰, set직급뷰] = useState("group");    // group | detail
  const [사유뷰, set사유뷰] = useState("전사");      // 전사 | 본부
  const [본부순, set본부순] = useState("조직");      // 조직 = 조직도 순서 | 인원 = 많은 순

  /** 본부 행 정렬 — [조직순]이면 고정 순서, [많은순]이면 값이 큰 순서. */
  const 본부순정렬 = (rows, 값) =>
    본부순 === "조직" ? 본부정렬(rows) : [...rows].sort((a, b) => (값(b) ?? -1) - (값(a) ?? -1));

  const isAll = yi === "ALL";
  const yLabel = isAll ? 전체라벨 : String(기준.연도[yi]);
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

    /* §4 사유는 사유 집계 구간(최근 사유연도수 개년)만 씁니다.
       분자를 잘랐으면 분모도 같은 구간으로 맞춰야 비율이 어긋나지 않습니다. */
    const 사유구간 = isAll ? idx.slice(사유시작) : yi >= 사유시작 ? [yi] : [];
    const 사유퇴사 = sum(사유구간.map((i) => 종합.퇴사[i]));
    const 사유평재 = sum(사유구간.map((i) => 평균[i]));
    const 구분합 = (k) =>
      sum(사유.filter((r) => r.kind === k).map((r) => sum(사유구간.map((i) => r.cnt[i]))));
    const 자발 = 구분합("자발적");
    const 비자발 = 구분합("비자발적");
    const 기타 = 사유퇴사 - 자발 - 비자발;

    const 평균근속 = isAll
      ? sum(종합.평균근속년.map((v, i) => v * 종합.퇴사[i])) / sum(종합.퇴사)
      : 종합.평균근속년[yi];

    // 전년 대비 (전체 기간 종합에서는 비교 대상 없음)
    const prev =
      !isAll && yi > 0 ? ratio(종합.퇴사[yi - 1], 평균[yi - 1]) : null;

    const 연환산 = isYtd && 퇴직율 != null ? (퇴직율 * 365) / 기준.기간일수[yi] : null;

    return {
      퇴사, 입사, 평재, 퇴직율, 조기, 조기비중, 평균근속, prev, 연환산, 평균,
      자발, 비자발, 기타, 사유구간, 사유퇴사, 사유평재,
    };
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
  const 연수 = isAll ? 기준.연도.length : 1; // ALL 모드의 분모는 전체 기간 합이므로 연평균으로 환산해 비교
  const 본부행 = useMemo(() => {
    const rows = 본부.map((b) => ({ name: b.name, ...slice(b, yi) }));
    const 남길 = 본부지표 === "rate"
      ? rows.filter((r) => r.rate != null && r.avg / 연수 >= 최소분모)
      : rows.filter((r) => r[본부지표] > 0);
    return 본부순정렬(남길, (r) => r[본부지표]);
  }, [yi, 본부지표, 연수, 본부순]);
  const 본부표시 = 본부행;
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

  /* [②-표3] 본부 × 근속구간 — 상단 [분석 연도] 선택에 맞춰 해당 연도만, [종합]이면 전 기간을 합칩니다.
     열합은 시점(근속구간)별 누계로 표 하단 합계 행에 씁니다. */
  const 본부근속 = useMemo(() => {
    const idx = isAll ? 기준.연도.map((_, i) => i) : [yi];
    const rows = 본부X근속
      .map((b) => {
        const cnt = 근속.map((_, si) => sum(idx.map((i) => b.cnt[si][i])));
        return { name: b.name, cnt, tot: sum(cnt) };
      })
      .filter((r) => r.tot > 0);
    const 열합 = 근속.map((_, si) => sum(rows.map((r) => r.cnt[si])));
    return { rows: 본부순정렬(rows, (r) => r.tot), 열합, 총계: sum(열합) };
  }, [yi, isAll, 본부순]);

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
  const 사유범위밖 = S.사유구간.length === 0;
  const 사유행 = useMemo(() => {
    return 사유
      .map((r) => {
        const c = sum(S.사유구간.map((i) => r.cnt[i]));
        return {
          name: r.name,
          kind: r.kind,
          cnt: c,
          구성비: ratio(c, S.사유퇴사),
          퇴직율: ratio(c, S.사유평재),
        };
      })
      .filter((r) => r.cnt > 0)
      .sort((a, b) => b.cnt - a.cnt);
  }, [S.사유구간, S.사유퇴사, S.사유평재]);

  const 구분추이 = useMemo(
    () =>
      사유연도.map((y, k) => {
        const i = 사유시작 + k;
        const row = { 연도: String(y) };
        KIND_ORDER.forEach((kd) => {
          row[kd] = sum(사유.filter((r) => r.kind === kd).map((r) => r.cnt[i]));
        });
        return row;
      }),
    []
  );

  /* [④-표5] 본부별 퇴사사유 — 상위 사유만 열로 세우고 나머지는 [기타 사유]로 묶습니다.
     행은 본부X사유에 담긴 본부(퇴사 누계 상위)만 있으므로 전사 합계와는 다를 수 있습니다. */
  const 사유열수 = 8;
  const 본부사유 = useMemo(() => {
    const g = S.사유구간;
    if (!본부X사유.length || !g.length) return null;
    const 구간합 = (b, si) => sum(g.map((i) => b.cnt[si][i]));
    const 후보 = 사유
      .map((r, si) => ({ si, name: r.name, kind: r.kind, tot: sum(본부X사유.map((b) => 구간합(b, si))) }))
      .filter((c) => c.tot > 0)
      .sort((a, b) => b.tot - a.tot);
    if (!후보.length) return null;
    const 표시 = 후보.slice(0, 사유열수);
    const 나머지 = 후보.slice(사유열수);
    const cols = 나머지.length
      ? [...표시, { si: -1, name: "기타 사유", kind: "기타", tot: sum(나머지.map((c) => c.tot)) }]
      : 표시;
    const rows = 본부X사유
      .map((b) => {
        const cnt = cols.map((c) =>
          c.si >= 0 ? 구간합(b, c.si) : sum(나머지.map((x) => 구간합(b, x.si))));
        const tot = sum(cnt);
        let top = -1;
        cnt.forEach((v, i) => { if (v > 0 && (top < 0 || v > cnt[top])) top = i; });
        return { name: b.name, cnt, tot, top };
      })
      .filter((r) => r.tot > 0);
    const 열합 = cols.map((_, ci) => sum(rows.map((r) => r.cnt[ci])));
    return { cols, rows: 본부순정렬(rows, (r) => r.tot), 열합, 총계: sum(열합) };
  }, [S.사유구간, 본부순]);

  /* [본부별] 버튼이 눌려 있고, 실제로 보여 줄 데이터가 있을 때만 본부별 보기로 넘어갑니다. */
  const 본부사유보기 = 사유뷰 === "본부" && !사유범위밖;

  /* 사유 구간에 한 건도 없는 구분은 범례·표에서 뺍니다 — 0건 항목이 유령으로 남지 않게. */
  const 사용구분 = KIND_ORDER.filter((k) => 구분추이.some((row) => row[k] > 0));

  const 전사퇴직율 = S.퇴직율;

  /* ------------------------------------------------------------------ */
  return (
    <div
      style={{
        background: C.bg,
        minHeight: "100%",
        fontFamily: FONT_STACK,
        color: C.text,
        letterSpacing: -0.1,
        padding: "26px 30px 48px",
      }}
    >
      <style>{FONT_IMPORT}</style>

      {/* ── 헤더 ─────────────────────────────────────────────── */}
      <div style={{ marginBottom: 20 }}>
        <div style={{ fontSize: 30, fontWeight: 700, letterSpacing: -1.1, lineHeight: 1.2 }}>
          퇴직율 분석 대시보드
        </div>
        <div style={{ fontSize: 12.5, color: C.sub, marginTop: 6 }}>
          {출처.파일} 기준 · 분석 기간 {기준.연도[0]} ~ {기준.연도[연도수 - 1]} ·
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
              { value: "ALL", label: 전체라벨 },
            ]}
          />
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap" }}>
          <span style={{ fontSize: 12, color: C.sub, fontWeight: 600 }}>본부 정렬</span>
          <Segmented
            size="sm"
            value={본부순}
            onChange={set본부순}
            options={[
              { value: "조직", label: "조직순" },
              { value: "인원", label: "많은순" },
            ]}
          />
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
      </div>

      {isYtd && (
        <div
          style={{
            display: "flex",
            gap: 8,
            alignItems: "center",
            background: C.warnBg,
            border: "1px solid " + C.warnBorder,
            color: C.warn,
            borderRadius: 12,
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
              ? 연도수 + "개년 연평균 · (기초+기말)÷2"
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
          value={pct(ratio(S.자발, S.사유평재))}
          badge={사유_실측 ? null : "샘플"}
          sub={
            사유범위밖
              ? `퇴사사유는 ${사유범위}만 집계됩니다`
              : `자발 ${num(S.자발)} · 비자발 ${num(S.비자발)}` +
                (S.기타 > 0 ? ` · 기타 ${num(S.기타)}` : "") + "명" +
                (isAll ? ` (${사유범위})` : "")
          }
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
                  <Tooltip content={<ChartTooltip unit="명" />} cursor={{ fill: C.hover }} />
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
              ? `${본부순 === "조직" ? "조직도 순서" : "퇴직율이 높은 순서"}로 세운 전 본부입니다. 점선은 전사 ${yLabel} 퇴직율 ${pct(전사퇴직율)}이고, 연평균 재직인원 ${최소분모}명 미만은 제외했습니다.`
              : `${본부순 === "조직" ? "조직도 순서" : "값이 큰 순서"}로 세운 전 본부입니다. 값이 0인 본부는 빠집니다.`
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
            rows={본부순정렬(본부.map((b) => ({ name: b.name, ...slice(b, yi) })), (r) => r.rate)
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
                  cursor={{ fill: C.hover }}
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
              head={["근속구간", ...기준.연도.map(String), 연도수 + "개년 합계", 연도수 + "개년 비중"]}
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
                    cursor={{ fill: C.hover }}
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
            desc={
              (isAll ? 전체기간 + " 누계" : yLabel + "년") +
              " · 인원(명)과 그 본부 퇴사자 중 비중(%). 색이 진할수록 그 구간에 퇴사가 몰려 있습니다."
            }
          />
          <div style={{ overflowX: "auto" }}>
            <table
              style={{
                width: "100%",
                borderCollapse: "separate",
                borderSpacing: "2px 1px",
                fontSize: 11,
                fontVariantNumeric: "tabular-nums",
                whiteSpace: "nowrap",
              }}
            >
              <thead>
                <tr>
                  <th style={{ textAlign: "left", padding: "2px 8px", color: C.sub, fontWeight: 600 }}>본부</th>
                  {근속.map((r) => (
                    <th key={r.name} style={{ textAlign: "center", padding: "2px 6px", color: C.sub, fontWeight: 600 }}>
                      {r.name}
                    </th>
                  ))}
                  <th style={{ textAlign: "right", padding: "2px 8px", color: C.sub, fontWeight: 600 }}>누계</th>
                </tr>
              </thead>
              <tbody>
                {본부근속.rows.map((b) => (
                  <tr key={b.name}>
                    <td style={{ padding: "2px 8px", color: C.text, fontWeight: 500 }}>{b.name}</td>
                    {b.cnt.map((v, i) => (
                      <HeatCell
                        key={i}
                        v={v}
                        share={ratio(v, b.tot)}
                        title={`${b.name} · ${근속[i].name} · ${v}명`}
                      />
                    ))}
                    <td style={{ textAlign: "right", padding: "2px 8px", color: C.sub }}>{num(b.tot)}</td>
                  </tr>
                ))}
              </tbody>
              <tfoot>
                <tr>
                  <td
                    style={{
                      padding: "3px 8px",
                      color: C.text,
                      fontWeight: 700,
                      borderTop: "1px solid " + C.border,
                    }}
                  >
                    시점별 누계
                  </td>
                  {본부근속.열합.map((v, i) => (
                    <HeatCell
                      key={i}
                      v={v}
                      share={ratio(v, 본부근속.총계)}
                      strong
                      style={{ borderTop: "1px solid " + C.border }}
                      title={`${근속[i].name} · 표 전체 누계 ${v}명`}
                    />
                  ))}
                  <td
                    style={{
                      textAlign: "right",
                      padding: "3px 8px",
                      color: C.text,
                      fontWeight: 700,
                      borderTop: "1px solid " + C.border,
                    }}
                  >
                    {num(본부근속.총계)}
                  </td>
                </tr>
              </tfoot>
            </table>
          </div>
          <div style={{ display: "flex", alignItems: "center", gap: 8, marginTop: 8, fontSize: 11, color: C.faint }}>
            <span>0%</span>
            {HEAT_RAMP.map((c, i) => (
              <span key={i} style={{ width: 22, height: 8, borderRadius: 2, background: c }} />
            ))}
            <span>50%+</span>
          </div>
          <Note>
            셀은 <b>퇴사자 수(명)</b>이고 괄호 안은 비중(%) — 본부 행은 그 본부 퇴사자 중 비중, 맨 아래 [시점별
            누계] 행은 표에 실린 전 본부를 합친 비중입니다. 색이 진할수록 비중이 높습니다. 누계 퇴사자가 적은
            본부는 한두 명으로도 비중이 크게 튈 수 있습니다.
          </Note>
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
                <Tooltip cursor={{ fill: C.hover }} content={<ChartTooltip fmt={(v) => pct(v)} />} />
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
        <Card style={본부사유보기 ? { gridColumn: "1 / -1" } : undefined}>
          <SectionHead
            eyebrow="④ 퇴사사유별 퇴직율"
            title={`${본부사유보기 ? "본부별 퇴사사유" : "퇴사사유별 인원"} · ${사유범위밖 ? "집계 구간 밖" : isAll ? 사유범위 + " 누계" : yLabel}`}
            desc={
              본부사유보기
                ? `본부마다 어떤 사유로 많이 나갔는지 봅니다. 상위 ${사유열수}개 사유만 열로 세우고 나머지는 [기타 사유]로 묶었습니다.`
                : `퇴사사유는 ${사유범위} ${사유연도수}개년만 집계합니다 — 그 이전은 RawData [퇴사사유]가 비어 있습니다. 사유 목록은 [코드설정] J열, 자발/비자발 구분은 생성기의 기본 분류를 따릅니다.`
            }
            right={
              사유_실측 && !사유범위밖 && 본부X사유.length ? (
                <Segmented
                  size="sm"
                  value={사유뷰}
                  onChange={set사유뷰}
                  options={[
                    { value: "전사", label: "전사" },
                    { value: "본부", label: "본부별" },
                  ]}
                />
              ) : null
            }
          />
          {!사유_실측 && (
            <SampleWarning>
              RawData [퇴사사유]가 비어 있어 <b>이 섹션과 [자발적 퇴직율] KPI는 샘플 값</b>입니다. 사유를 입력한 뒤
              생성기를 다시 돌리면 실측으로 바뀝니다.
            </SampleWarning>
          )}
          {사유범위밖 ? (
            <Note>
              <b>{yLabel}년</b>은 퇴사사유 집계 구간({사유범위}) 밖입니다. 상단 [분석 연도]에서 {사유범위} 중 한 해
              또는 [{전체라벨}]을 선택하면 사유가 표시됩니다.
            </Note>
          ) : 본부사유보기 ? (
            !본부사유 ? (
              <Note>이 구간에 사유가 입력된 퇴사자가 없어 본부별로 나눌 수 없습니다.</Note>
            ) : asTable ? (
              <DataTable
                head={["본부", ...본부사유.cols.map((c) => c.name), "누계", "최다 사유"]}
                rows={[
                  ...본부사유.rows.map((r) => [
                    r.name,
                    ...r.cnt.map((v) => (v ? num(v) : "—")),
                    num(r.tot),
                    r.top < 0 ? "—" : 본부사유.cols[r.top].name,
                  ]),
                  ["합계", ...본부사유.열합.map((v) => num(v)), num(본부사유.총계), ""],
                ]}
                align={[..."x".repeat(본부사유.cols.length + 2)].map(() => "right").concat("left")}
              />
            ) : (
              <>
                <div style={{ overflowX: "auto" }}>
                  <table
                    style={{
                      width: "100%",
                      borderCollapse: "separate",
                      borderSpacing: "2px 1px",
                      fontSize: 11,
                      fontVariantNumeric: "tabular-nums",
                      whiteSpace: "nowrap",
                    }}
                  >
                    <thead>
                      <tr>
                        <th style={{ textAlign: "left", padding: "2px 8px", color: C.sub, fontWeight: 600 }}>본부</th>
                        {본부사유.cols.map((c) => (
                          <th
                            key={c.name}
                            title={`${c.name} · ${c.kind} · ${c.tot}명`}
                            style={{
                              textAlign: "center",
                              padding: "2px 6px",
                              color: KIND_COLOR[c.kind] || C.sub,
                              fontWeight: 600,
                            }}
                          >
                            {짧은이름(c.name)}
                          </th>
                        ))}
                        <th style={{ textAlign: "right", padding: "2px 8px", color: C.sub, fontWeight: 600 }}>누계</th>
                        <th style={{ textAlign: "left", padding: "2px 8px", color: C.sub, fontWeight: 600 }}>최다 사유</th>
                      </tr>
                    </thead>
                    <tbody>
                      {본부사유.rows.map((b) => (
                        <tr key={b.name}>
                          <td style={{ padding: "2px 8px", color: C.text, fontWeight: 500 }}>{b.name}</td>
                          {b.cnt.map((v, i) => (
                            <HeatCell
                              key={i}
                              v={v}
                              share={ratio(v, b.tot)}
                              title={`${b.name} · ${본부사유.cols[i].name} · ${v}명`}
                            />
                          ))}
                          <td style={{ textAlign: "right", padding: "2px 8px", color: C.sub }}>{num(b.tot)}</td>
                          <td
                            style={{
                              padding: "2px 8px",
                              color: b.top < 0 ? C.faint : KIND_COLOR[본부사유.cols[b.top].kind] || C.text,
                              fontWeight: 600,
                            }}
                          >
                            {b.top < 0 ? "—" : 본부사유.cols[b.top].name}
                          </td>
                        </tr>
                      ))}
                    </tbody>
                    <tfoot>
                      <tr>
                        <td style={{ padding: "3px 8px", color: C.text, fontWeight: 700, borderTop: "1px solid " + C.border }}>
                          합계
                        </td>
                        {본부사유.열합.map((v, i) => (
                          <HeatCell
                            key={i}
                            v={v}
                            share={ratio(v, 본부사유.총계)}
                            strong
                            style={{ borderTop: "1px solid " + C.border }}
                            title={`${본부사유.cols[i].name} · 표 전체 누계 ${v}명`}
                          />
                        ))}
                        <td style={{ textAlign: "right", padding: "3px 8px", color: C.text, fontWeight: 700, borderTop: "1px solid " + C.border }}>
                          {num(본부사유.총계)}
                        </td>
                        <td style={{ borderTop: "1px solid " + C.border }} />
                      </tr>
                    </tfoot>
                  </table>
                </div>
                <div style={{ display: "flex", gap: 16, marginTop: 12, flexWrap: "wrap", alignItems: "center" }}>
                  {사용구분.map((k) => (
                    <Swatch key={k} color={KIND_COLOR[k]} label={k + " 사유"} />
                  ))}
                  <span style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 11.5, color: C.faint }}>
                    <span>0%</span>
                    {HEAT_RAMP.map((c, i) => (
                      <span key={i} style={{ width: 20, height: 9, borderRadius: 2, background: c }} />
                    ))}
                    <span>50%+</span>
                  </span>
                </div>
              </>
            )
          ) : asTable ? (
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
                  <Tooltip cursor={{ fill: C.hover }} content={<ChartTooltip unit="명" />} />
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
                {사용구분.map((k) => (
                  <Swatch key={k} color={KIND_COLOR[k]} label={k} />
                ))}
              </div>
            </>
          )}
          {본부사유보기 ? (
            <Note>
              셀은 <b>퇴사자 수(명)</b>이고 괄호 안은 그 본부 퇴사자 중 해당 사유의 비중(%)입니다. 맨 아래 [합계]
              행은 표에 실린 전 본부를 합친 값이고, 색이 진할수록 비중이 높습니다. 머리글 색은 자발·비자발 구분을
              나타냅니다. 이 표에는 퇴사 누계 상위 본부만 실려 있어 합계가 전사 퇴사자 수보다 적을 수 있습니다.
            </Note>
          ) : (
            <Note>
              구성비의 분모는 <b>퇴사자 수</b>, 사유별 퇴직율의 분모는 <b>전사 평균 재직인원</b>입니다(템플릿 표2·표3).
              두 값을 섞어 읽지 마세요.
            </Note>
          )}
        </Card>

        <Card>
          <SectionHead
            eyebrow="④ 표4 · 자발 / 비자발"
            title={`사유구분별 퇴사자 추이 · ${사유범위}`}
            desc="자발적 퇴직은 조직이 개선할 여지가 있는 이탈, 비자발적 퇴직은 회사 결정에 따른 이탈입니다. 상단 [분석 연도] 선택과 무관하게 사유 집계 구간 전체를 보여줍니다."
          />
          {!사유_실측 && <SampleWarning>샘플 값입니다. 실측이 아닙니다.</SampleWarning>}
          {asTable ? (
            <DataTable
              head={["사유구분", ...사유연도.map(String), 사유연도수 + "개년 합계"]}
              rows={[
                ...사용구분.map((k) => [
                  k,
                  ...사유연도.map((_, i) => num(구분추이[i][k])),
                  num(sum(사유연도.map((_, i) => 구분추이[i][k]))),
                ]),
                [
                  "합계",
                  ...종합.퇴사.slice(사유시작).map((v) => num(v)),
                  num(sum(종합.퇴사.slice(사유시작))),
                ],
              ]}
            />
          ) : (
            <>
              <ResponsiveContainer width="100%" height={Math.max(240, 사유행.length * 25 + 40)}>
                <BarChart data={구분추이} margin={{ top: 4, right: 12, left: -14, bottom: 0 }}>
                  <CartesianGrid stroke={C.grid} vertical={false} />
                  <XAxis dataKey="연도" tick={{ fontSize: 11.5, fill: C.sub }} axisLine={{ stroke: C.border }} tickLine={false} />
                  <YAxis tick={{ fontSize: 11, fill: C.sub }} axisLine={false} tickLine={false} width={44} unit="명" />
                  <Tooltip cursor={{ fill: C.hover }} content={<ChartTooltip unit="명" />} />
                  <Legend verticalAlign="bottom" height={30} iconType="square" wrapperStyle={{ fontSize: 12, color: C.sub }} />
                  {사용구분.map((k) => (
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
                자발적 퇴직율({pct(ratio(S.자발, S.사유평재))})은 자발적 퇴사자 ÷ 전사 평균 재직인원입니다. 분자·분모
                모두 {사유범위} 구간 기준입니다.
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
          · [{전체라벨}]의 퇴직율은 누계 퇴사자 ÷ 연도별 평균 재직인원의 합입니다. 조기퇴직 비중도 누계 기준으로
          계산합니다(엑셀 템플릿 종합대시보드 H21은 연도별 비중의 단순평균이라 값이 조금 다를 수 있습니다).
        </div>
        <div style={{ marginTop: 8, color: C.faint }}>
          · ④ 퇴사사유와 [자발적 퇴직율] KPI는 <b>{사유범위} {사유연도수}개년</b>만 집계합니다. 나머지 섹션(①~③)은
          {" "}{전체기간} {연도수}개년 전체입니다. 연도를 {사유연도[0]}년 이전으로 선택하면 ④와 [자발적 퇴직율]은 값을 내지 않습니다.
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
    </div>
  );
}
