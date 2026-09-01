import fs from "node:fs/promises";
import path from "node:path";
import { Workbook, SpreadsheetFile } from "@oai/artifact-tool";

const outDir = path.resolve("outputs");
await fs.mkdir(outDir, { recursive: true });

const wb = Workbook.create();
const guide = wb.worksheets.add("사용안내");
const settings = wb.worksheets.add("기준설정");
const prior = wb.worksheets.add("전기원장");
const current = wb.worksheets.add("당기원장");
const compare = wb.worksheets.add("계정비교");
const detail = wb.worksheets.add("세부원인분석");
const checks = wb.worksheets.add("점검");

const navy = "#17365D";
const blue = "#D9EAF7";
const paleBlue = "#EEF5FB";
const orange = "#FCE4D6";
const paleOrange = "#FFF4EC";
const green = "#E2F0D9";
const red = "#F4CCCC";
const gray = "#E7E6E6";
const dark = "#1F2937";
const white = "#FFFFFF";
const moneyFmt = "#,##0;[Red](#,##0);-";

function title(sheet, range, textValue, subtitle) {
  sheet.showGridLines = false;
  sheet.getRange(range).merge();
  sheet.getRange(range).values = [[textValue]];
  sheet.getRange(range).format = {
    fill: navy,
    font: { bold: true, color: white, size: 18 },
    verticalAlignment: "center",
    horizontalAlignment: "left",
  };
  sheet.getRange(range).format.rowHeight = 34;
  if (subtitle) {
    const startCol = range.split(":")[0].replace(/[0-9]/g, "");
    sheet.getRange(`${startCol}2`).values = [[subtitle]];
    sheet.getRange(`${startCol}2`).format = { font: { color: "#5B6573", italic: true, size: 10 } };
  }
}

function header(rng) {
  rng.format = {
    fill: navy,
    font: { bold: true, color: white },
    horizontalAlignment: "center",
    verticalAlignment: "center",
    wrapText: true,
    borders: { preset: "inside", style: "thin", color: "#B4C7E7" },
  };
  rng.format.rowHeight = 30;
}

// 사용안내
title(guide, "A1:H1", "당기·전기 계정원장 비교분석", "계정별 증감과 중요 변동의 원인을 검토하는 재사용형 템플릿");
guide.getRange("A4:H4").merge();
guide.getRange("A4:H4").values = [["사용 순서"]];
guide.getRange("A4:H4").format = { fill: blue, font: { bold: true, color: dark, size: 13 } };
guide.getRange("A6:B12").values = [
  ["1", "기준설정 시트에서 증감률·증감액 기준을 확인합니다. 기본값은 20%, 100,000,000원입니다."],
  ["2", "전기원장·당기원장 시트의 예시 행을 지우고 A:G 열에 데이터를 붙여 넣습니다."],
  ["3", "A 계정코드, B 계정명, C 전표일자, D 결의서번호, E 결의서명, F 차변, G 대변 순서를 유지합니다."],
  ["4", "H:J 열은 계산 열입니다. 1,000건을 초과하면 마지막 행의 수식을 아래로 복사합니다."],
  ["5", "계정비교 시트에서 신규·소멸·금액변동과 원인분석 대상을 확인합니다."],
  ["6", "세부원인분석 시트에서 대상 계정의 결의서명·원인분류·붙임 확인 여부를 검토합니다."],
  ["7", "결의서명만으로 원인이 명확하지 않은 거래는 붙임 문서를 확인하고 검토메모를 남깁니다."],
];
guide.getRange("A6:A12").format = { fill: orange, font: { bold: true, color: "#C65911" }, horizontalAlignment: "center" };
guide.getRange("B6:B12").format = { wrapText: true, verticalAlignment: "center" };
guide.getRange("A6:B12").format.borders = { preset: "inside", style: "thin", color: "#D9E2F3" };
guide.getRange("A15:H15").merge();
guide.getRange("A15:H15").values = [["판정 로직"]];
guide.getRange("A15:H15").format = { fill: blue, font: { bold: true, color: dark, size: 13 } };
guide.getRange("A17:C21").values = [
  ["유형", "정의", "원인분석 대상 판정"],
  ["신규", "전기 금액 0, 당기 금액 존재", "증감액 절대값이 기준 이상이면 대상"],
  ["소멸", "전기 금액 존재, 당기 금액 0", "증감액 절대값이 기준 이상이면 대상"],
  ["금액변동", "전기·당기 모두 금액 존재", "증감률과 증감액 절대값이 모두 기준 이상이면 대상"],
  ["변동없음", "전기와 당기 순액이 동일", "일반"],
];
header(guide.getRange("A17:C17"));
guide.getRange("A18:C21").format = { wrapText: true, borders: { preset: "inside", style: "thin", color: "#D9E2F3" } };
guide.getRange("A24:H26").merge();
guide.getRange("A24:H26").values = [["주의: 증감률은 |당기-전기| ÷ |전기|로 계산합니다. 전기 금액이 0인 신규 계정과 당기 금액이 0인 소멸 계정은 증감률 대신 유형으로 판정합니다. 차변-대변 순액 기준이 회사의 원장 부호 체계와 다르면 H열 수식을 조정하십시오."]];
guide.getRange("A24:H26").format = { fill: paleOrange, font: { color: "#9C5700" }, wrapText: true, verticalAlignment: "center" };
guide.getRange("A:A").format.columnWidth = 8;
guide.getRange("B:B").format.columnWidth = 82;
guide.getRange("C:C").format.columnWidth = 42;
guide.getRange("D:H").format.columnWidth = 12;

// 기준설정
title(settings, "A1:F1", "기준설정", "파란색 입력 셀은 사용자가 변경할 수 있습니다.");
settings.getRange("A4:C7").values = [
  ["항목", "입력값", "설명"],
  ["증감률 기준", 0.2, "전기 대비 절대 증감률"],
  ["증감액 기준(원)", 100000000, "절대 증감액"],
  ["원장 입력 최대행", 1000, "기본 수식이 준비된 거래행 수"],
];
header(settings.getRange("A4:C4"));
settings.getRange("B5:B7").format = { fill: paleBlue, font: { color: "#0070C0", bold: true } };
settings.getRange("B5").format.numberFormat = "0.0%";
settings.getRange("B6:B7").format.numberFormat = "#,##0";
settings.getRange("A4:C7").format.borders = { preset: "inside", style: "thin", color: "#D9E2F3" };
settings.getRange("A10:F10").merge();
settings.getRange("A10:F10").values = [["결의서명 기반 1차 원인분류 키워드"]];
settings.getRange("A10:F10").format = { fill: blue, font: { bold: true, color: dark } };
settings.getRange("A12:B17").values = [
  ["원인분류", "결의서명 포함 키워드 예시"],
  ["인건비", "급여"],
  ["임차·시설", "임차"],
  ["광고·행사", "광고"],
  ["용역·수수료", "용역"],
  ["기타/불명", "위 키워드에 해당하지 않거나 명칭이 불충분한 경우 — 붙임 확인"],
];
header(settings.getRange("A12:B12"));
settings.getRange("A13:B17").format = { wrapText: true, borders: { preset: "inside", style: "thin", color: "#D9E2F3" } };
settings.getRange("A:A").format.columnWidth = 22;
settings.getRange("B:B").format.columnWidth = 24;
settings.getRange("C:C").format.columnWidth = 44;

const ledgerHeaders = [["계정코드", "계정명", "전표일자", "결의서번호", "결의서명", "차변", "대변", "순액(차변-대변)", "1차 원인분류", "붙임 확인", "검토메모/붙임경로"]];
const priorData = [
  ["510100", "급여", new Date("2025-01-31"), "P-001", "1월 정규직 급여 지급", 350000000, 0, null, null, null, ""],
  ["520300", "광고선전비", new Date("2025-03-20"), "P-032", "브랜드 캠페인 광고 용역", 120000000, 0, null, null, null, ""],
  ["530100", "임차료", new Date("2025-02-10"), "P-020", "본사 사무실 임차료", 240000000, 0, null, null, null, ""],
  ["540200", "지급수수료", new Date("2025-05-15"), "P-088", "기타 정산", 150000000, 0, null, null, null, ""],
];
const currentData = [
  ["510100", "급여", new Date("2026-01-31"), "C-001", "1월 정규직 급여 및 신규채용 반영", 470000000, 0, null, null, null, ""],
  ["520300", "광고선전비", new Date("2026-03-22"), "C-041", "신제품 출시 광고 캠페인", 260000000, 0, null, null, null, ""],
  ["530100", "임차료", new Date("2026-02-10"), "C-022", "본사 사무실 임차료", 240000000, 0, null, null, null, ""],
  ["550500", "소프트웨어", new Date("2026-04-01"), "C-067", "ERP 시스템 구축 용역", 180000000, 0, null, null, null, ""],
];

function buildLedger(sheet, titleText, sample) {
  title(sheet, "A1:K1", titleText, "A:G 입력 / H:J 자동계산 / K 검토자 입력");
  sheet.getRange("A4:K4").values = ledgerHeaders;
  header(sheet.getRange("A4:K4"));
  sheet.getRange(`A5:K${4 + sample.length}`).values = sample;
  sheet.getRange("H5").formulas = [["=IF(A5=\"\",\"\",F5-G5)"]];
  sheet.getRange("H5:H1004").fillDown();
  sheet.getRange("I5").formulas = [["=IF(A5=\"\",\"\",IF(E5=\"\",\"기타/불명\",IF(IFERROR(SEARCH(\"급여\",E5),0)>0,\"인건비\",IF(IFERROR(SEARCH(\"임차\",E5),0)>0,\"임차·시설\",IF(IFERROR(SEARCH(\"광고\",E5),0)>0,\"광고·행사\",IF(IFERROR(SEARCH(\"용역\",E5),0)>0,\"용역·수수료\",\"기타/불명\"))))))"]];
  sheet.getRange("I5:I1004").fillDown();
  sheet.getRange("J5").formulas = [["=IF(A5=\"\",\"\",IF(I5=\"기타/불명\",\"확인 필요\",\"결의서명으로 1차 확인\"))"]];
  sheet.getRange("J5:J1004").fillDown();
  sheet.getRange("C5:C1004").format.numberFormat = "yyyy-mm-dd";
  sheet.getRange("F5:H1004").format.numberFormat = moneyFmt;
  sheet.getRange("H5:J1004").format.fill = "#F2F2F2";
  sheet.getRange("K5:K1004").format.fill = paleBlue;
  sheet.getRange("A4:K1004").format.borders = { insideHorizontal: { style: "hair", color: "#E7E6E6" } };
  sheet.getRange("J5:J1004").conditionalFormats.add("containsText", { text: "확인 필요", format: { fill: red, font: { bold: true, color: "#9C0006" } } });
  sheet.freezePanes.freezeRows(4);
  sheet.getRange("A:A").format.columnWidth = 13;
  sheet.getRange("B:B").format.columnWidth = 20;
  sheet.getRange("C:C").format.columnWidth = 13;
  sheet.getRange("D:D").format.columnWidth = 16;
  sheet.getRange("E:E").format.columnWidth = 38;
  sheet.getRange("F:H").format.columnWidth = 17;
  sheet.getRange("I:I").format.columnWidth = 18;
  sheet.getRange("J:J").format.columnWidth = 22;
  sheet.getRange("K:K").format.columnWidth = 34;
}

buildLedger(prior, "전기 계정원장", priorData);
buildLedger(current, "당기 계정원장", currentData);

// 계정비교
title(compare, "A1:K1", "계정별 증감 비교", "전기·당기 계정코드의 합집합을 기준으로 신규·소멸·금액변동을 판정합니다.");
compare.getRange("A4:K4").values = [["계정코드", "계정명", "전기 순액", "당기 순액", "증감액", "증감률", "변동유형", "판정", "검토 안내", "붙임 확인 필요 건수", "원인분석 메모"]];
header(compare.getRange("A4:K4"));
compare.getRange("A5").formulas = [["=SORT(UNIQUE(VSTACK(FILTER('전기원장'!$A$5:$A$1004,'전기원장'!$A$5:$A$1004<>\"\"),FILTER('당기원장'!$A$5:$A$1004,'당기원장'!$A$5:$A$1004<>\"\"))))"]];
compare.getRange("B5").formulas = [["=IF(A5=\"\",\"\",IFERROR(XLOOKUP(A5,'당기원장'!$A$5:$A$1004,'당기원장'!$B$5:$B$1004),XLOOKUP(A5,'전기원장'!$A$5:$A$1004,'전기원장'!$B$5:$B$1004,\"\")))"]];
compare.getRange("B5:B204").fillDown();
compare.getRange("C5").formulas = [["=IF(A5=\"\",\"\",SUMIF('전기원장'!$A$5:$A$1004,A5,'전기원장'!$H$5:$H$1004))"]];
compare.getRange("C5:C204").fillDown();
compare.getRange("D5").formulas = [["=IF(A5=\"\",\"\",SUMIF('당기원장'!$A$5:$A$1004,A5,'당기원장'!$H$5:$H$1004))"]];
compare.getRange("D5:D204").fillDown();
compare.getRange("E5").formulas = [["=IF(A5=\"\",\"\",D5-C5)"]];
compare.getRange("E5:E204").fillDown();
compare.getRange("F5").formulas = [["=IF(A5=\"\",\"\",IF(C5=0,IF(D5=0,0,1),ABS(E5)/ABS(C5)))"]];
compare.getRange("F5:F204").fillDown();
compare.getRange("G5").formulas = [["=IF(A5=\"\",\"\",IF(AND(C5=0,D5<>0),\"신규\",IF(AND(C5<>0,D5=0),\"소멸\",IF(E5=0,\"변동없음\",\"금액변동\"))))"]];
compare.getRange("G5:G204").fillDown();
compare.getRange("H5").formulas = [["=IF(A5=\"\",\"\",IF(AND(ABS(E5)>='기준설정'!$B$6,OR(G5=\"신규\",G5=\"소멸\",F5>='기준설정'!$B$5)),\"원인분석 대상\",\"일반\"))"]];
compare.getRange("H5:H204").fillDown();
compare.getRange("I5").formulas = [["=IF(A5=\"\",\"\",IF(H5=\"원인분석 대상\",\"세부원인분석 시트에서 거래별 원인을 검토\",\"-\"))"]];
compare.getRange("I5:I204").fillDown();
compare.getRange("J5").formulas = [["=IF(A5=\"\",\"\",COUNTIFS('전기원장'!$A$5:$A$1004,A5,'전기원장'!$J$5:$J$1004,\"확인 필요\")+COUNTIFS('당기원장'!$A$5:$A$1004,A5,'당기원장'!$J$5:$J$1004,\"확인 필요\"))"]];
compare.getRange("J5:J204").fillDown();
compare.getRange("K5").formulas = [["=IF(A5=\"\",\"\",IF(H5<>\"원인분석 대상\",\"\",IF(J5>0,\"붙임 문서 확인 후 원인 확정 필요\",\"결의서명 기반 원인분류 검토\")))"]];
compare.getRange("K5:K204").fillDown();
compare.getRange("C5:E204").format.numberFormat = moneyFmt;
compare.getRange("F5:F204").format.numberFormat = "0.0%";
compare.getRange("A4:K204").format.borders = { insideHorizontal: { style: "hair", color: "#E7E6E6" } };
compare.getRange("H5:H204").conditionalFormats.add("containsText", { text: "원인분석 대상", format: { fill: red, font: { bold: true, color: "#9C0006" } } });
compare.getRange("J5:J204").conditionalFormats.add("cellIs", { operator: "greaterThan", formula: 0, format: { fill: orange, font: { bold: true, color: "#C65911" } } });
compare.freezePanes.freezeRows(4);
compare.getRange("A:A").format.columnWidth = 13;
compare.getRange("B:B").format.columnWidth = 20;
compare.getRange("C:E").format.columnWidth = 17;
compare.getRange("F:F").format.columnWidth = 12;
compare.getRange("G:H").format.columnWidth = 17;
compare.getRange("I:I").format.columnWidth = 35;
compare.getRange("J:J").format.columnWidth = 17;
compare.getRange("K:K").format.columnWidth = 34;

// 세부원인분석
title(detail, "A1:K1", "세부 원인분석", "원인분석 대상 계정의 전기·당기 거래를 함께 표시합니다.");
detail.getRange("A3:K3").merge();
detail.getRange("A3:K3").values = [["※ '확인 필요' 거래는 결의서명만으로 판단하지 말고 붙임 문서(계약서·정산서·세금계산서 등)를 확인한 후 원인분석 메모를 확정하십시오."]];
detail.getRange("A3:K3").format = { fill: paleOrange, font: { color: "#9C5700", bold: true }, wrapText: true };
detail.getRange("A5:K5").values = [["기간", "계정코드", "계정명", "전표일자", "결의서번호", "결의서명", "순액", "1차 원인분류", "붙임 확인", "검토메모/붙임경로", "분석 포인트"]];
header(detail.getRange("A5:K5"));
detail.getRange("B6").formulas = [["=LET(t,VSTACK(HSTACK('전기원장'!$A$5:$A$1004,'전기원장'!$B$5:$B$1004,'전기원장'!$C$5:$C$1004,'전기원장'!$D$5:$D$1004,'전기원장'!$E$5:$E$1004,'전기원장'!$H$5:$H$1004,'전기원장'!$I$5:$I$1004,'전기원장'!$J$5:$J$1004,'전기원장'!$K$5:$K$1004),HSTACK('당기원장'!$A$5:$A$1004,'당기원장'!$B$5:$B$1004,'당기원장'!$C$5:$C$1004,'당기원장'!$D$5:$D$1004,'당기원장'!$E$5:$E$1004,'당기원장'!$H$5:$H$1004,'당기원장'!$I$5:$I$1004,'당기원장'!$J$5:$J$1004,'당기원장'!$K$5:$K$1004)),FILTER(t,(CHOOSECOLS(t,1)<>\"\")*ISNUMBER(XMATCH(CHOOSECOLS(t,1),FILTER('계정비교'!$A$5:$A$204,'계정비교'!$H$5:$H$204=\"원인분석 대상\"))),\"대상 없음\"))"]];
detail.getRange("A6").formulas = [["=IF(B6=\"\",\"\",IF(COUNTIFS('전기원장'!$A$5:$A$1004,B6,'전기원장'!$D$5:$D$1004,E6)>0,\"전기\",\"당기\"))"]];
detail.getRange("A6:A2005").fillDown();
detail.getRange("K6").formulas = [["=IF(B6=\"\",\"\",IF(I6=\"확인 필요\",\"붙임 문서 확인 후 거래 성격·기간·상대방·일회성 여부 확인\",\"결의서명과 전기 대비 발생 건수·단가·시점 차이 검토\"))"]];
detail.getRange("K6:K2005").fillDown();
detail.getRange("D6:D2005").format.numberFormat = "yyyy-mm-dd";
detail.getRange("G6:G2005").format.numberFormat = moneyFmt;
detail.getRange("I6:I2005").conditionalFormats.add("containsText", { text: "확인 필요", format: { fill: red, font: { bold: true, color: "#9C0006" } } });
detail.freezePanes.freezeRows(5);
detail.getRange("A:A").format.columnWidth = 9;
detail.getRange("B:B").format.columnWidth = 13;
detail.getRange("C:C").format.columnWidth = 20;
detail.getRange("D:D").format.columnWidth = 13;
detail.getRange("E:E").format.columnWidth = 16;
detail.getRange("F:F").format.columnWidth = 38;
detail.getRange("G:G").format.columnWidth = 17;
detail.getRange("H:H").format.columnWidth = 18;
detail.getRange("I:I").format.columnWidth = 22;
detail.getRange("J:J").format.columnWidth = 34;
detail.getRange("K:K").format.columnWidth = 52;

// 점검
title(checks, "A1:F1", "점검 및 통제", "분석 결과를 사용하기 전에 아래 상태가 PASS인지 확인하십시오.");
checks.getRange("A4:B4").values = [["MODEL STATUS", null]];
checks.getRange("A4").format = { fill: navy, font: { bold: true, color: white } };
checks.getRange("B4").formulas = [["=IF(COUNTIF(B8:B13,\"FAIL\")=0,\"PASS\",\"FAIL\")"]];
checks.getRange("B4").format = { fill: gray, font: { bold: true, color: dark, size: 14 }, horizontalAlignment: "center" };
checks.getRange("B4").conditionalFormats.add("containsText", { text: "FAIL", format: { fill: red, font: { bold: true, color: "#9C0006" } } });
checks.getRange("B4").conditionalFormats.add("containsText", { text: "PASS", format: { fill: green, font: { bold: true, color: "#006100" } } });
checks.getRange("A7:D13").values = [
  ["점검항목", "상태", "수치", "조치"],
  ["전기 원장 입력 건수", null, null, "0건이면 전기원장을 확인"],
  ["당기 원장 입력 건수", null, null, "0건이면 당기원장을 확인"],
  ["원인분석 대상 계정 수", null, null, "대상 계정의 원인 메모 검토"],
  ["붙임 확인 필요 거래 수", null, null, "붙임 문서 확인 후 메모 입력"],
  ["계정코드 누락 거래", null, null, "A열 계정코드 보완"],
  ["계정명 누락 거래", null, null, "B열 계정명 보완"],
];
header(checks.getRange("A7:D7"));
checks.getRange("C8").formulas = [["=COUNTIF('전기원장'!$A$5:$A$1004,\"<>\")"]];
checks.getRange("C9").formulas = [["=COUNTIF('당기원장'!$A$5:$A$1004,\"<>\")"]];
checks.getRange("C10").formulas = [["=COUNTIF('계정비교'!$H$5:$H$204,\"원인분석 대상\")"]];
checks.getRange("C11").formulas = [["=COUNTIF('전기원장'!$J$5:$J$1004,\"확인 필요\")+COUNTIF('당기원장'!$J$5:$J$1004,\"확인 필요\")"]];
checks.getRange("C12").formulas = [["=COUNTIFS('전기원장'!$B$5:$B$1004,\"<>\",'전기원장'!$A$5:$A$1004,\"\")+COUNTIFS('당기원장'!$B$5:$B$1004,\"<>\",'당기원장'!$A$5:$A$1004,\"\")"]];
checks.getRange("C13").formulas = [["=COUNTIFS('전기원장'!$A$5:$A$1004,\"<>\",'전기원장'!$B$5:$B$1004,\"\")+COUNTIFS('당기원장'!$A$5:$A$1004,\"<>\",'당기원장'!$B$5:$B$1004,\"\")"]];
checks.getRange("B8").formulas = [["=IF(C8>0,\"PASS\",\"FAIL\")"]];
checks.getRange("B9").formulas = [["=IF(C9>0,\"PASS\",\"FAIL\")"]];
checks.getRange("B10").formulas = [["=IF(C10>=0,\"PASS\",\"FAIL\")"]];
checks.getRange("B11").formulas = [["=IF(C11=0,\"PASS\",\"FAIL\")"]];
checks.getRange("B12").formulas = [["=IF(C12=0,\"PASS\",\"FAIL\")"]];
checks.getRange("B13").formulas = [["=IF(C13=0,\"PASS\",\"FAIL\")"]];
checks.getRange("B8:B13").conditionalFormats.add("containsText", { text: "FAIL", format: { fill: red, font: { bold: true, color: "#9C0006" } } });
checks.getRange("B8:B13").conditionalFormats.add("containsText", { text: "PASS", format: { fill: green, font: { bold: true, color: "#006100" } } });
checks.getRange("A7:D13").format.borders = { preset: "inside", style: "thin", color: "#D9E2F3" };
checks.getRange("A:A").format.columnWidth = 30;
checks.getRange("B:B").format.columnWidth = 13;
checks.getRange("C:C").format.columnWidth = 15;
checks.getRange("D:D").format.columnWidth = 44;

const inspect = await wb.inspect({ kind: "table", range: "계정비교!A1:K12", include: "values,formulas", tableMaxRows: 12, tableMaxCols: 11, maxChars: 8000 });
console.log(inspect.ndjson);
const formulaErrors = await wb.inspect({ kind: "match", searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A", options: { useRegex: true, maxResults: 100 }, summary: "final formula error scan" });
console.log(formulaErrors.ndjson);

for (const [sheetName, range, file] of [
  ["사용안내", "A1:H26", "preview_guide.png"],
  ["기준설정", "A1:F20", "preview_settings.png"],
  ["전기원장", "A1:K12", "preview_prior.png"],
  ["당기원장", "A1:K12", "preview_current.png"],
  ["계정비교", "A1:K14", "preview_compare.png"],
  ["세부원인분석", "A1:K18", "preview_detail.png"],
  ["점검", "A1:F14", "preview_checks.png"],
]) {
  const image = await wb.render({ sheetName, range, scale: 1.3, format: "png" });
  await fs.writeFile(path.join("work", file), new Uint8Array(await image.arrayBuffer()));
}

const xlsx = await SpreadsheetFile.exportXlsx(wb);
const outputPath = path.join(outDir, "당기_전기_계정원장_비교분석_템플릿.xlsx");
await xlsx.save(outputPath);
console.log(`OUTPUT=${outputPath}`);

