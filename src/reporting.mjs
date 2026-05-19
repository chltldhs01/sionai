import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const CHANNELS = ["파워링크", "쇼핑검색"];

function parseCsv(csvText) {
  const rows = [];
  let field = "";
  let row = [];
  let inQuotes = false;

  for (let i = 0; i < csvText.length; i += 1) {
    const char = csvText[i];
    const next = csvText[i + 1];

    if (char === '"') {
      if (inQuotes && next === '"') {
        field += '"';
        i += 1;
      } else {
        inQuotes = !inQuotes;
      }
      continue;
    }

    if (char === "," && !inQuotes) {
      row.push(field);
      field = "";
      continue;
    }

    if ((char === "\n" || char === "\r") && !inQuotes) {
      if (char === "\r" && next === "\n") {
        i += 1;
      }
      row.push(field);
      rows.push(row);
      row = [];
      field = "";
      continue;
    }

    field += char;
  }

  if (field.length > 0 || row.length > 0) {
    row.push(field);
    rows.push(row);
  }

  return rows.filter((currentRow) => currentRow.some((cell) => cell !== ""));
}

function normalizeNumber(value) {
  if (value === null || value === undefined || value === "") {
    return 0;
  }
  const cleaned = String(value).replace(/,/g, "").trim();
  const num = Number(cleaned);
  return Number.isFinite(num) ? num : 0;
}

function parseDateLabel(value) {
  const match = String(value).trim().match(/^(\d{4})\.(\d{2})\.(\d{2})\.$/);
  if (!match) {
    return null;
  }
  return new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]));
}

function parseInputDate(value) {
  const match = String(value).trim().match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!match) {
    return new Date(value);
  }
  return new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]));
}

function formatDate(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}.${month}.${day}`;
}

function formatDateForFile(date) {
  return formatDate(date).replace(/\./g, "");
}

function currency(value) {
  return `${Math.round(value).toLocaleString("ko-KR")}원`;
}

function safeRatio(numerator, denominator, factor = 100, digits = 2) {
  if (!denominator) {
    return 0;
  }
  return Number(((numerator / denominator) * factor).toFixed(digits));
}

function safeValue(numerator, denominator) {
  if (!denominator) {
    return 0;
  }
  return Math.round(numerator / denominator);
}

function sanitizeSheetName(name) {
  return name.replace(/[\\/*?:[\]]/g, "_").slice(0, 28);
}

function columnLettersToIndex(letters) {
  return letters.split("").reduce((acc, char) => acc * 26 + char.charCodeAt(0) - 64, 0) - 1;
}

function columnIndexToLetters(index) {
  let value = index + 1;
  let output = "";
  while (value > 0) {
    const remainder = (value - 1) % 26;
    output = String.fromCharCode(65 + remainder) + output;
    value = Math.floor((value - 1) / 26);
  }
  return output;
}

function parseCellRef(cellRef) {
  const match = cellRef.match(/^([A-Z]+)(\d+)$/);
  if (!match) {
    throw new Error(`지원하지 않는 셀 주소입니다: ${cellRef}`);
  }
  return {
    col: columnLettersToIndex(match[1]),
    row: Number(match[2]) - 1,
  };
}

function cellRef(rowIndex, colIndex) {
  return `${columnIndexToLetters(colIndex)}${rowIndex + 1}`;
}

function groupBy(items, getKey) {
  const map = new Map();
  for (const item of items) {
    const key = getKey(item);
    if (!map.has(key)) {
      map.set(key, []);
    }
    map.get(key).push(item);
  }
  return map;
}

function buildRecords(csvText) {
  const rows = parseCsv(csvText.replace(/^\uFEFF/, ""));
  if (rows.length < 3) {
    throw new Error("CSV 구조를 인식하지 못했습니다. 네이버 SA 보고서 형식인지 확인해주세요.");
  }

  const header = rows[1];
  const rawRows = rows.slice(2);

  return rawRows
    .filter((row) => row.length >= header.length)
    .map((row) => Object.fromEntries(header.map((key, index) => [key, row[index] ?? ""])))
    .map((row) => {
      const date = parseDateLabel(row["일별"]);
      if (!date) {
        return null;
      }

      return {
        date,
        dateLabel: formatDate(date),
        campaign: row["캠페인"] || "",
        campaignType: row["캠페인유형"] || "",
        adGroup: row["광고그룹"] || "",
        adGroupType: row["광고그룹유형"] || "",
        query: row["검색어"] || "",
        impressions: normalizeNumber(row["노출수"]),
        clicks: normalizeNumber(row["클릭수"]),
        ctrRaw: normalizeNumber(row["클릭률(%)"]),
        cpcRaw: normalizeNumber(row["평균 CPC"]),
        cost: normalizeNumber(row["총비용"]),
        conversions: normalizeNumber(row["총 전환수"]),
        directConversions: normalizeNumber(row["직접전환수"]),
        indirectConversions: normalizeNumber(row["간접전환수"]),
        conversionRateRaw: normalizeNumber(row["총 전환율(%)"]),
        revenue: normalizeNumber(row["총 전환매출액(원)"]),
        directRevenue: normalizeNumber(row["직접전환매출액(원)"]),
        indirectRevenue: normalizeNumber(row["간접전환매출액(원)"]),
        cpaRaw: normalizeNumber(row["총 전환당비용(원)"]),
        roasRaw: normalizeNumber(row["총 광고수익률(%)"]),
        purchaseConversions: normalizeNumber(row["구매완료 전환수"]),
        purchaseRevenue: normalizeNumber(row["구매완료 전환매출액(원)"]),
        purchaseRoasRaw: normalizeNumber(row["구매완료 광고수익률(%)"]),
      };
    })
    .filter(Boolean);
}

function sumMetrics(records) {
  const totals = {
    impressions: 0,
    clicks: 0,
    cost: 0,
    conversions: 0,
    revenue: 0,
    purchaseConversions: 0,
    purchaseRevenue: 0,
  };

  for (const record of records) {
    totals.impressions += record.impressions;
    totals.clicks += record.clicks;
    totals.cost += record.cost;
    totals.conversions += record.conversions;
    totals.revenue += record.revenue;
    totals.purchaseConversions += record.purchaseConversions;
    totals.purchaseRevenue += record.purchaseRevenue;
  }

  return {
    ...totals,
    ctr: safeRatio(totals.clicks, totals.impressions),
    cpc: safeValue(totals.cost, totals.clicks),
    cvr: safeRatio(totals.conversions, totals.clicks),
    purchaseCvr: safeRatio(totals.purchaseConversions, totals.clicks),
    cpa: safeValue(totals.cost, totals.conversions),
    purchaseCpa: safeValue(totals.cost, totals.purchaseConversions),
    roas: safeRatio(totals.revenue, totals.cost, 100, 0),
    purchaseRoas: safeRatio(totals.purchaseRevenue, totals.cost, 100, 0),
  };
}

function summarizeByChannel(records) {
  return CHANNELS.map((channel) => {
    const channelRecords = records.filter((record) => record.campaignType === channel);
    return {
      name: channel,
      ...sumMetrics(channelRecords),
    };
  }).filter((channel) => channel.cost > 0 || channel.impressions > 0);
}

function summarizeDaily(records) {
  return Array.from(groupBy(records, (record) => record.dateLabel).entries())
    .sort(([left], [right]) => left.localeCompare(right))
    .map(([dateLabel, dayRecords]) => ({
      dateLabel,
      ...sumMetrics(dayRecords),
    }));
}

function summarizeDailyByChannel(records) {
  return Array.from(groupBy(records, (record) => `${record.dateLabel}|||${record.campaignType}`).entries())
    .map(([key, groupedRecords]) => {
      const [dateLabel, channel] = key.split("|||");
      return {
        dateLabel,
        channel,
        ...sumMetrics(groupedRecords),
      };
    })
    .filter((row) => CHANNELS.includes(row.channel))
    .sort((left, right) => {
      const dateCompare = left.dateLabel.localeCompare(right.dateLabel);
      return dateCompare === 0 ? left.channel.localeCompare(right.channel) : dateCompare;
    });
}

function summarizeKeyword(records) {
  return Array.from(
    groupBy(records.filter((record) => CHANNELS.includes(record.campaignType)), (record) => `${record.query}|||${record.campaignType}`),
  )
    .map(([key, groupedRecords]) => {
      const [query, channel] = key.split("|||");
      return {
        query,
        channel,
        ...sumMetrics(groupedRecords),
      };
    })
    .sort((left, right) => right.cost - left.cost);
}

function summarizeCampaign(records, targetChannel) {
  return Array.from(
    groupBy(
      records.filter((record) => record.campaignType === targetChannel),
      (record) => `${record.campaign}|||${record.adGroup}`,
    ),
  )
    .map(([key, groupedRecords]) => {
      const [campaign, adGroup] = key.split("|||");
      return {
        campaign,
        adGroup,
        ...sumMetrics(groupedRecords),
      };
    })
    .sort((left, right) => right.cost - left.cost);
}

function parseMemoEntries(memoText) {
  return memoText
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean);
}

function setHeaderStyle(range, fill = "#12343B") {
  range.format = {
    fill,
    font: { bold: true, color: "#FFFFFF" },
    wrapText: true,
    verticalAlignment: "center",
  };
}

function setBodyTableStyle(range) {
  range.format = {
    verticalAlignment: "center",
  };
}

function addKpiCard(sheet, cell, title, value, note, tone) {
  const fillMap = {
    orange: "#FFF1EB",
    teal: "#EAF8F6",
    navy: "#EEF4F8",
    sand: "#F7F1E8",
  };
  const { row, col } = parseCellRef(cell);
  const titleRange = sheet.getRange(cellRef(row, col));
  const valueCell = sheet.getRange(cellRef(row + 1, col));
  const noteCell = sheet.getRange(cellRef(row + 2, col));
  const block = sheet.getRange(`${cellRef(row, col)}:${cellRef(row + 2, col + 1)}`);

  block.merge(true);
  block.format = {
    fill: fillMap[tone] || "#FFFFFF",
  };

  titleRange.values = [[title]];
  titleRange.format = {
    font: { bold: true, color: "#5A6B7B", size: 10 },
  };

  valueCell.values = [[value]];
  valueCell.format = {
    font: { bold: true, color: "#153147", size: 18 },
  };

  noteCell.values = [[note]];
  noteCell.format = {
    font: { color: "#5A6B7B", size: 10 },
    wrapText: true,
  };
}

function addSectionTitle(sheet, rangeRef, text) {
  const range = sheet.getRange(rangeRef);
  range.merge();
  range.values = [[text]];
  range.format = {
    fill: "#12343B",
    font: { bold: true, color: "#FFFFFF", size: 12 },
    verticalAlignment: "center",
  };
}

function buildNarrative(brand, startDate, endDate, overall, channelSummary, topKeywords, memoEntries) {
  const powerlink = channelSummary.find((row) => row.name === "파워링크");
  const shopping = channelSummary.find((row) => row.name === "쇼핑검색");
  const leadingKeywords = topKeywords
    .filter((row) => row.purchaseRevenue > 0)
    .slice(0, 3)
    .map((row) => `${row.query}(${row.channel})`);

  return [
    `${formatDate(startDate)}부터 ${formatDate(endDate)}까지 ${brand} 네이버 SA는 총 광고비 ${currency(overall.cost)}를 집행했고, 구매완료 ${overall.purchaseConversions.toLocaleString("ko-KR")}건과 구매완료 매출 ${currency(overall.purchaseRevenue)}를 기록했습니다.`,
    `광고주 관점의 핵심 효율 지표인 구매완료 ROAS는 ${overall.purchaseRoas.toLocaleString("ko-KR")}%이며, 구매완료 CPA는 ${currency(overall.purchaseCpa)} 수준입니다.`,
    powerlink
      ? `파워링크는 구매완료 ROAS ${powerlink.purchaseRoas.toLocaleString("ko-KR")}%로 효율 역할을 맡고 있습니다.`
      : "파워링크 데이터는 이번 보고 구간에 없습니다.",
    shopping
      ? `쇼핑검색은 구매완료 매출 ${currency(shopping.purchaseRevenue)}로 볼륨 확보에 기여했습니다.`
      : "쇼핑검색 데이터는 이번 보고 구간에 없습니다.",
    leadingKeywords.length
      ? `구매 기여 키워드는 ${leadingKeywords.join(", ")} 중심으로 확인됩니다.`
      : "키워드별 구매 기여는 추가 확인이 필요합니다.",
    memoEntries.length
      ? `운영 메모 기준으로는 ${memoEntries.slice(0, 3).join(" / ")} 등의 조정이 반영되었습니다.`
      : "운영 메모는 별도 입력되지 않았습니다.",
  ];
}

function writeTable(sheet, startCell, headers, rows, numberFormats = {}) {
  const safeRows = rows.length ? rows : [Array(headers.length).fill("")];
  const { row, col } = parseCellRef(startCell);
  const headerStart = cellRef(row, col);
  const headerEnd = cellRef(row, col + headers.length - 1);
  const dataStartRow = row + 1;
  const dataEndRow = row + safeRows.length;
  const fullEnd = cellRef(dataEndRow, col + headers.length - 1);

  sheet.getRange(`${headerStart}:${fullEnd}`).values = [headers, ...safeRows];
  setHeaderStyle(sheet.getRange(`${headerStart}:${headerEnd}`));
  setBodyTableStyle(sheet.getRange(`${cellRef(dataStartRow, col)}:${fullEnd}`));

  for (const [colIndex, format] of Object.entries(numberFormats)) {
    const columnNumber = Number(colIndex);
    const targetTop = cellRef(dataStartRow, col + columnNumber);
    const targetBottom = cellRef(dataEndRow, col + columnNumber);
    sheet.getRange(`${targetTop}:${targetBottom}`).format.numberFormat = format;
  }
}

export async function buildWorkbookFromPayload(payload) {
  const {
    csvText,
    brand = "바이맘",
    reportStartDate,
    reportEndDate,
    memoText = "",
  } = payload;

  if (!csvText) {
    throw new Error("CSV 파일 내용이 비어 있습니다.");
  }

  const startDate = parseInputDate(reportStartDate);
  const endDate = parseInputDate(reportEndDate);
  if (Number.isNaN(startDate.getTime()) || Number.isNaN(endDate.getTime())) {
    throw new Error("보고 기간을 정확히 입력해주세요.");
  }

  const allRecords = buildRecords(csvText);
  const reportRecords = allRecords.filter((record) => record.date >= startDate && record.date <= endDate);
  if (!reportRecords.length) {
    throw new Error("선택한 기간에 해당하는 데이터가 없습니다.");
  }

  const advertiserRecords = reportRecords.filter((record) => CHANNELS.includes(record.campaignType));
  const overall = sumMetrics(advertiserRecords);
  const channelSummary = summarizeByChannel(advertiserRecords);
  const dailySummary = summarizeDaily(advertiserRecords);
  const dailyChannelSummary = summarizeDailyByChannel(advertiserRecords);
  const keywordSummary = summarizeKeyword(advertiserRecords);
  const powerlinkSummary = summarizeCampaign(advertiserRecords, "파워링크");
  const shoppingSummary = summarizeCampaign(advertiserRecords, "쇼핑검색");
  const memoEntries = parseMemoEntries(memoText);
  const narrative = buildNarrative(
    brand,
    startDate,
    endDate,
    overall,
    channelSummary,
    keywordSummary,
    memoEntries,
  );

  const workbook = Workbook.create();
  const dashboardSheet = workbook.worksheets.add("광고주 요약");
  const overallSheet = workbook.worksheets.add("전체 요약");
  const dailySheet = workbook.worksheets.add("일자별 추이");
  const powerlinkSheet = workbook.worksheets.add("파워링크 상세");
  const shoppingSheet = workbook.worksheets.add("쇼핑검색 상세");
  const keywordSheet = workbook.worksheets.add("키워드 상세");
  const memoSheet = workbook.worksheets.add("운영 메모");
  const sourceSheet = workbook.worksheets.add("원본 데이터");

  dashboardSheet.getRange("A1:N2").merge();
  dashboardSheet.getRange("A1").values = [[`${brand} 네이버 SA 주간보고서`]];
  dashboardSheet.getRange("A1").format = {
    fill: "#12343B",
    font: { bold: true, color: "#FFFFFF", size: 18 },
    horizontalAlignment: "center",
    verticalAlignment: "center",
  };

  dashboardSheet.getRange("A3:N3").merge();
  dashboardSheet.getRange("A3").values = [[`보고 기간: ${formatDate(startDate)} ~ ${formatDate(endDate)}`]];
  dashboardSheet.getRange("A3").format = {
    fill: "#EEF4F8",
    font: { bold: true, color: "#153147", size: 11 },
    horizontalAlignment: "center",
  };

  addKpiCard(
    dashboardSheet,
    "A5",
    "총 광고비",
    currency(overall.cost),
    `전체 클릭 ${overall.clicks.toLocaleString("ko-KR")}회`,
    "orange",
  );
  addKpiCard(
    dashboardSheet,
    "C5",
    "총 노출수",
    overall.impressions.toLocaleString("ko-KR"),
    `CTR ${overall.ctr.toFixed(2)}%`,
    "navy",
  );
  addKpiCard(
    dashboardSheet,
    "E5",
    "구매완료 수",
    `${overall.purchaseConversions.toLocaleString("ko-KR")}건`,
    `구매완료 CVR ${overall.purchaseCvr.toFixed(2)}%`,
    "teal",
  );
  addKpiCard(
    dashboardSheet,
    "G5",
    "구매완료 매출",
    currency(overall.purchaseRevenue),
    `총 전환매출 ${currency(overall.revenue)}`,
    "teal",
  );
  addKpiCard(
    dashboardSheet,
    "I5",
    "구매완료 ROAS",
    `${overall.purchaseRoas.toLocaleString("ko-KR")}%`,
    `전체 ROAS ${overall.roas.toLocaleString("ko-KR")}%`,
    "sand",
  );
  addKpiCard(
    dashboardSheet,
    "K5",
    "구매완료 CPA",
    currency(overall.purchaseCpa),
    `전체 CPA ${currency(overall.cpa)}`,
    "orange",
  );

  addSectionTitle(dashboardSheet, "A10:N10", "광고주용 핵심 요약");
  dashboardSheet.getRange(`A11:N${10 + narrative.length}`).merge(true);
  dashboardSheet.getRange(`A11:A${10 + narrative.length}`).values = narrative.map((line) => [line]);
  dashboardSheet.getRange(`A11:A${10 + narrative.length}`).format = {
    wrapText: true,
    verticalAlignment: "top",
    font: { color: "#243B53" },
  };

  addSectionTitle(dashboardSheet, "A16:G16", "매체별 핵심 지표");
  const channelRows = channelSummary.map((row) => [
    row.name,
    row.cost,
    row.purchaseConversions,
    row.purchaseRevenue,
    row.purchaseRoas / 100,
    row.purchaseCpa,
    row.ctr / 100,
  ]);
  writeTable(
    dashboardSheet,
    "A17",
    ["매체", "광고비", "구매완료 수", "구매완료 매출", "구매완료 ROAS", "구매완료 CPA", "CTR"],
    channelRows,
    {
      1: "#,##0",
      2: "#,##0",
      3: "#,##0",
      4: "0.0%",
      5: "#,##0",
      6: "0.0%",
    },
  );

  addSectionTitle(dashboardSheet, "H16:N16", "상위 키워드");
  writeTable(
    dashboardSheet,
    "H17",
    ["검색어", "매체", "구매완료 수", "구매완료 매출", "구매완료 ROAS"],
    keywordSummary.slice(0, 5).map((row) => [
      row.query,
      row.channel,
      row.purchaseConversions,
      row.purchaseRevenue,
      row.purchaseRoas / 100,
    ]),
    {
      2: "#,##0",
      3: "#,##0",
      4: "0.0%",
    },
  );

  addSectionTitle(dashboardSheet, "A26:N26", "일자별 구매 흐름");
  writeTable(
    dashboardSheet,
    "A27",
    ["일자", "광고비", "총 클릭수", "구매완료 수", "구매완료 매출", "구매완료 ROAS", "CTR"],
    dailySummary.map((row) => [
      row.dateLabel,
      row.cost,
      row.clicks,
      row.purchaseConversions,
      row.purchaseRevenue,
      row.purchaseRoas / 100,
      row.ctr / 100,
    ]),
    {
      1: "#,##0",
      2: "#,##0",
      3: "#,##0",
      4: "#,##0",
      5: "0.0%",
      6: "0.0%",
    },
  );

  dashboardSheet.freezePanes.freezeRows(3);

  const overallRows = [
    ["지표", "값", "설명"],
    ["총 광고비", overall.cost, "파워링크 + 쇼핑검색 기준 총 집행 비용"],
    ["총 노출수", overall.impressions, "광고 노출 합계"],
    ["총 클릭수", overall.clicks, "광고 클릭 합계"],
    ["CTR", overall.ctr / 100, "클릭률"],
    ["평균 CPC", overall.cpc, "클릭당 평균 비용"],
    ["총 전환수", overall.conversions, "플랫폼 전체 전환 기준"],
    ["총 전환매출", overall.revenue, "플랫폼 전체 전환매출"],
    ["구매완료 수", overall.purchaseConversions, "광고주 기준 최중요 전환"],
    ["구매완료 매출", overall.purchaseRevenue, "구매완료 전환매출"],
    ["구매완료 ROAS", overall.purchaseRoas / 100, "광고주 보고용 핵심 효율 지표"],
    ["구매완료 CPA", overall.purchaseCpa, "구매완료 1건당 비용"],
  ];
  overallSheet.getRange("A1:C12").values = overallRows;
  setHeaderStyle(overallSheet.getRange("A1:C1"));
  overallSheet.getRange("B2:B12").format.numberFormat = "#,##0";
  overallSheet.getRange("B5").format.numberFormat = "0.0%";
  overallSheet.getRange("B11").format.numberFormat = "0.0%";

  writeTable(
    overallSheet,
    "E1",
    ["매체", "광고비", "구매완료 수", "구매완료 매출", "구매완료 ROAS", "구매완료 CPA", "CTR", "CVR"],
    channelSummary.map((row) => [
      row.name,
      row.cost,
      row.purchaseConversions,
      row.purchaseRevenue,
      row.purchaseRoas / 100,
      row.purchaseCpa,
      row.ctr / 100,
      row.purchaseCvr / 100,
    ]),
    {
      1: "#,##0",
      2: "#,##0",
      3: "#,##0",
      4: "0.0%",
      5: "#,##0",
      6: "0.0%",
      7: "0.0%",
    },
  );
  overallSheet.freezePanes.freezeRows(1);

  writeTable(
    dailySheet,
    "A1",
    ["일자", "광고비", "총 클릭수", "구매완료 수", "구매완료 매출", "구매완료 ROAS", "CTR"],
    dailySummary.map((row) => [
      row.dateLabel,
      row.cost,
      row.clicks,
      row.purchaseConversions,
      row.purchaseRevenue,
      row.purchaseRoas / 100,
      row.ctr / 100,
    ]),
    {
      1: "#,##0",
      2: "#,##0",
      3: "#,##0",
      4: "#,##0",
      5: "0.0%",
      6: "0.0%",
    },
  );

  writeTable(
    dailySheet,
    "J1",
    ["일자", "매체", "광고비", "구매완료 수", "구매완료 매출", "구매완료 ROAS"],
    dailyChannelSummary.map((row) => [
      row.dateLabel,
      row.channel,
      row.cost,
      row.purchaseConversions,
      row.purchaseRevenue,
      row.purchaseRoas / 100,
    ]),
    {
      2: "#,##0",
      3: "#,##0",
      4: "#,##0",
      5: "0.0%",
    },
  );
  dailySheet.freezePanes.freezeRows(1);

  writeTable(
    powerlinkSheet,
    "A1",
    ["캠페인", "광고그룹", "광고비", "클릭수", "구매완료 수", "구매완료 매출", "구매완료 ROAS", "구매완료 CPA"],
    powerlinkSummary.map((row) => [
      row.campaign,
      row.adGroup,
      row.cost,
      row.clicks,
      row.purchaseConversions,
      row.purchaseRevenue,
      row.purchaseRoas / 100,
      row.purchaseCpa,
    ]),
    {
      2: "#,##0",
      3: "#,##0",
      4: "#,##0",
      5: "#,##0",
      6: "0.0%",
      7: "#,##0",
    },
  );
  powerlinkSheet.freezePanes.freezeRows(1);

  writeTable(
    shoppingSheet,
    "A1",
    ["캠페인", "광고그룹", "광고비", "클릭수", "구매완료 수", "구매완료 매출", "구매완료 ROAS", "구매완료 CPA"],
    shoppingSummary.map((row) => [
      row.campaign,
      row.adGroup,
      row.cost,
      row.clicks,
      row.purchaseConversions,
      row.purchaseRevenue,
      row.purchaseRoas / 100,
      row.purchaseCpa,
    ]),
    {
      2: "#,##0",
      3: "#,##0",
      4: "#,##0",
      5: "#,##0",
      6: "0.0%",
      7: "#,##0",
    },
  );
  shoppingSheet.freezePanes.freezeRows(1);

  writeTable(
    keywordSheet,
    "A1",
    ["검색어", "매체", "광고비", "클릭수", "구매완료 수", "구매완료 매출", "구매완료 ROAS", "구매완료 CPA"],
    keywordSummary.map((row) => [
      row.query,
      row.channel,
      row.cost,
      row.clicks,
      row.purchaseConversions,
      row.purchaseRevenue,
      row.purchaseRoas / 100,
      row.purchaseCpa,
    ]),
    {
      2: "#,##0",
      3: "#,##0",
      4: "#,##0",
      5: "#,##0",
      6: "0.0%",
      7: "#,##0",
    },
  );
  keywordSheet.freezePanes.freezeRows(1);

  memoSheet.getRange("A1:C1").values = [["구분", "내용", "비고"]];
  setHeaderStyle(memoSheet.getRange("A1:C1"));
  const memoRows = memoEntries.length
    ? memoEntries.map((entry) => ["운영 메모", entry, "직접 입력"])
    : [["운영 메모", "입력된 메모가 없습니다.", ""]];
  memoSheet.getRange(`A2:C${memoRows.length + 1}`).values = memoRows;
  memoSheet.getRange(`B2:B${memoRows.length + 1}`).format = { wrapText: true };
  memoSheet.freezePanes.freezeRows(1);

  writeTable(
    sourceSheet,
    "A1",
    [
      "일자",
      "캠페인",
      "캠페인유형",
      "광고그룹",
      "광고그룹유형",
      "검색어",
      "노출수",
      "클릭수",
      "총비용",
      "총 전환수",
      "총 전환매출액",
      "구매완료 전환수",
      "구매완료 전환매출액",
    ],
    reportRecords.map((row) => [
      row.dateLabel,
      row.campaign,
      row.campaignType,
      row.adGroup,
      row.adGroupType,
      row.query,
      row.impressions,
      row.clicks,
      row.cost,
      row.conversions,
      row.revenue,
      row.purchaseConversions,
      row.purchaseRevenue,
    ]),
    {
      6: "#,##0",
      7: "#,##0",
      8: "#,##0",
      9: "#,##0",
      10: "#,##0",
      11: "#,##0",
      12: "#,##0",
    },
  );
  sourceSheet.freezePanes.freezeRows(1);

  const sheets = [
    dashboardSheet,
    overallSheet,
    dailySheet,
    powerlinkSheet,
    shoppingSheet,
    keywordSheet,
    memoSheet,
    sourceSheet,
  ];

  for (const sheet of sheets) {
    sheet.showGridLines = false;
    const used = sheet.getUsedRange();
    if (used) {
      used.format.autofitColumns();
      used.format.autofitRows();
    }
  }

  dashboardSheet.getRange("A1:N40").format.wrapText = true;

  const fileName = `${sanitizeSheetName(brand)}_주간보고서_${formatDateForFile(startDate)}_${formatDateForFile(endDate)}.xlsx`;
  const xlsx = await SpreadsheetFile.exportXlsx(workbook);
  const bytes = Buffer.from(xlsx.data);

  return {
    fileName,
    bytes,
    contentType: xlsx.mime,
    summary: {
      brand,
      reportRange: `${formatDate(startDate)} ~ ${formatDate(endDate)}`,
      spend: currency(overall.cost),
      conversions: overall.purchaseConversions,
      revenue: currency(overall.purchaseRevenue),
      roas: `${overall.purchaseRoas.toLocaleString("ko-KR")}%`,
      raw: {
        cost: overall.cost,
        impressions: overall.impressions,
        clicks: overall.clicks,
        conversions: overall.purchaseConversions,
        revenue: overall.purchaseRevenue,
        ctr: overall.ctr,
        cpc: overall.cpc,
        cvr: overall.purchaseCvr,
        cpa: overall.purchaseCpa,
        roas: overall.purchaseRoas,
      },
      channels: channelSummary.map((entry) => ({
        name: entry.name,
        cost: entry.cost,
        clicks: entry.clicks,
        conversions: entry.purchaseConversions,
        revenue: entry.purchaseRevenue,
        roas: entry.purchaseRoas,
        ctr: entry.ctr,
        cvr: entry.purchaseCvr,
      })),
      topKeywords: keywordSummary.slice(0, 5).map((entry) => ({
        query: entry.query,
        channel: entry.channel,
        cost: entry.cost,
        conversions: entry.purchaseConversions,
        revenue: entry.purchaseRevenue,
        roas: entry.purchaseRoas,
      })),
      memoCount: memoEntries.length,
    },
  };
}
