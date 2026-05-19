import ExcelJS from "exceljs";

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

function parseCellRef(cellAddress) {
  const match = cellAddress.match(/^([A-Z]+)(\d+)$/);
  if (!match) {
    throw new Error(`지원하지 않는 셀 주소입니다: ${cellAddress}`);
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
    throw new Error("CSV 구조를 인식하지 못했습니다. 네이버 SA 보고서 형식인지 확인해 주세요.");
  }

  const header = rows[1];
  const rawRows = rows.slice(2);

  return rawRows
    .filter((row) => row.length >= header.length)
    .map((row) => Object.fromEntries(header.map((key, index) => [key, row[index] ?? ""])))
    .map((row) => {
      const date = parseDateLabel(row["일계"]);
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

function buildNarrative(brand, startDate, endDate, overall, channelSummary, topKeywords, memoEntries) {
  const powerlink = channelSummary.find((row) => row.name === "파워링크");
  const shopping = channelSummary.find((row) => row.name === "쇼핑검색");
  const leadingKeywords = topKeywords
    .filter((row) => row.purchaseRevenue > 0)
    .slice(0, 3)
    .map((row) => `${row.query}(${row.channel})`);

  return [
    `${formatDate(startDate)}부터 ${formatDate(endDate)}까지 ${brand} 네이버 SA의 총 광고비는 ${currency(overall.cost)}이며, 구매완료 ${overall.purchaseConversions.toLocaleString("ko-KR")}건과 구매완료 매출 ${currency(overall.purchaseRevenue)}를 기록했습니다.`,
    `광고주 기준 핵심 효율 지표인 구매완료 ROAS는 ${overall.purchaseRoas.toLocaleString("ko-KR")}%이며, 구매완료 CPA는 ${currency(overall.purchaseCpa)}입니다.`,
    powerlink
      ? `파워링크는 구매완료 ROAS ${powerlink.purchaseRoas.toLocaleString("ko-KR")}%로 효율을 만들고 있습니다.`
      : "파워링크 데이터는 이번 보고 구간에 없습니다.",
    shopping
      ? `쇼핑검색은 구매완료 매출 ${currency(shopping.purchaseRevenue)}로 볼륨 확보에 기여했습니다.`
      : "쇼핑검색 데이터는 이번 보고 구간에 없습니다.",
    leadingKeywords.length
      ? `구매 기여 키워드는 ${leadingKeywords.join(", ")} 중심으로 확인됩니다.`
      : "키워드별 구매 기여는 추가 확인이 필요합니다.",
    memoEntries.length
      ? `운영 메모 기준으로는 ${memoEntries.slice(0, 3).join(" / ")} 조정이 반영되었습니다.`
      : "운영 메모는 별도 입력되지 않았습니다.",
  ];
}

function toArgb(hex) {
  return `FF${hex.replace("#", "").toUpperCase()}`;
}

function applyFill(cell, hex) {
  cell.fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: toArgb(hex) },
  };
}

function applyBorder(cell, color = "D9E2EC") {
  cell.border = {
    top: { style: "thin", color: { argb: `FF${color}` } },
    left: { style: "thin", color: { argb: `FF${color}` } },
    bottom: { style: "thin", color: { argb: `FF${color}` } },
    right: { style: "thin", color: { argb: `FF${color}` } },
  };
}

function styleRange(worksheet, startRow, startCol, endRow, endCol, callback) {
  for (let row = startRow; row <= endRow; row += 1) {
    for (let col = startCol; col <= endCol; col += 1) {
      callback(worksheet.getCell(row, col), row, col);
    }
  }
}

function addMergedSectionTitle(worksheet, rangeRef, text) {
  worksheet.mergeCells(rangeRef);
  const cell = worksheet.getCell(rangeRef.split(":")[0]);
  cell.value = text;
  applyFill(cell, "#12343B");
  cell.font = { bold: true, color: { argb: toArgb("#FFFFFF") }, size: 12 };
  cell.alignment = { vertical: "middle", horizontal: "left" };
}

function addKpiCard(worksheet, cellAddress, title, value, note, tone) {
  const fillMap = {
    orange: "#FFF1EB",
    teal: "#EAF8F6",
    navy: "#EEF4F8",
    sand: "#F7F1E8",
  };
  const { row, col } = parseCellRef(cellAddress);

  worksheet.mergeCells(row + 1, col + 1, row + 1, col + 2);
  worksheet.mergeCells(row + 2, col + 1, row + 2, col + 2);
  worksheet.mergeCells(row + 3, col + 1, row + 3, col + 2);

  styleRange(worksheet, row + 1, col + 1, row + 3, col + 2, (cell) => {
    applyFill(cell, fillMap[tone] || "#FFFFFF");
    applyBorder(cell);
    cell.alignment = { vertical: "middle", horizontal: "left", wrapText: true };
  });

  const titleCell = worksheet.getCell(row + 1, col + 1);
  titleCell.value = title;
  titleCell.font = { bold: true, color: { argb: toArgb("#5A6B7B") }, size: 10 };

  const valueCell = worksheet.getCell(row + 2, col + 1);
  valueCell.value = value;
  valueCell.font = { bold: true, color: { argb: toArgb("#153147") }, size: 18 };

  const noteCell = worksheet.getCell(row + 3, col + 1);
  noteCell.value = note;
  noteCell.font = { color: { argb: toArgb("#5A6B7B") }, size: 10 };
}

function writeTable(worksheet, startCell, headers, rows, numberFormats = {}) {
  const safeRows = rows.length ? rows : [Array(headers.length).fill("")];
  const { row, col } = parseCellRef(startCell);
  const headerRowNumber = row + 1;
  const dataStartRowNumber = headerRowNumber + 1;

  headers.forEach((header, index) => {
    const cell = worksheet.getCell(headerRowNumber, col + index + 1);
    cell.value = header;
    applyFill(cell, "#12343B");
    applyBorder(cell, "C9D2DB");
    cell.font = { bold: true, color: { argb: toArgb("#FFFFFF") } };
    cell.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
  });

  safeRows.forEach((currentRow, rowIndex) => {
    currentRow.forEach((value, cellIndex) => {
      const cell = worksheet.getCell(dataStartRowNumber + rowIndex, col + cellIndex + 1);
      cell.value = value;
      applyBorder(cell);
      cell.alignment = {
        vertical: "middle",
        horizontal: typeof value === "number" ? "right" : "left",
        wrapText: true,
      };
    });
  });

  for (const [colIndex, format] of Object.entries(numberFormats)) {
    const targetColumn = col + Number(colIndex) + 1;
    for (let rowNumber = dataStartRowNumber; rowNumber < dataStartRowNumber + safeRows.length; rowNumber += 1) {
      worksheet.getCell(rowNumber, targetColumn).numFmt = format;
    }
  }

  return {
    endRow: dataStartRowNumber + safeRows.length - 1,
  };
}

function autoFitWorksheet(worksheet, minWidth = 12, maxWidth = 36) {
  worksheet.columns.forEach((column) => {
    let maxLength = minWidth;
    column.eachCell({ includeEmpty: true }, (cell) => {
      let value = cell.value;
      if (value && typeof value === "object") {
        if ("richText" in value) {
          value = value.richText.map((item) => item.text).join("");
        } else if ("text" in value) {
          value = value.text;
        } else if ("result" in value) {
          value = value.result;
        }
      }
      const lines = String(value ?? "").split("\n");
      for (const line of lines) {
        maxLength = Math.max(maxLength, Math.min(maxWidth, line.length + 2));
      }
    });
    column.width = maxLength;
  });
}

function configureSheetView(worksheet, freezeRows = 1) {
  worksheet.views = [{ state: "frozen", ySplit: freezeRows, showGridLines: false }];
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
    throw new Error("보고 기간을 정확하게 입력해 주세요.");
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

  const workbook = new ExcelJS.Workbook();
  workbook.creator = "sionAI";
  workbook.created = new Date();

  const dashboardSheet = workbook.addWorksheet("광고주 요약");
  const overallSheet = workbook.addWorksheet("전체 요약");
  const dailySheet = workbook.addWorksheet("일자별 추이");
  const powerlinkSheet = workbook.addWorksheet("파워링크 상세");
  const shoppingSheet = workbook.addWorksheet("쇼핑검색 상세");
  const keywordSheet = workbook.addWorksheet("키워드 상세");
  const memoSheet = workbook.addWorksheet("운영 메모");
  const sourceSheet = workbook.addWorksheet("원본 데이터");

  dashboardSheet.mergeCells("A1:N2");
  dashboardSheet.getCell("A1").value = `${brand} 네이버 SA 주간보고서`;
  applyFill(dashboardSheet.getCell("A1"), "#12343B");
  dashboardSheet.getCell("A1").font = { bold: true, color: { argb: toArgb("#FFFFFF") }, size: 18 };
  dashboardSheet.getCell("A1").alignment = { vertical: "middle", horizontal: "center" };

  dashboardSheet.mergeCells("A3:N3");
  dashboardSheet.getCell("A3").value = `보고 기간: ${formatDate(startDate)} ~ ${formatDate(endDate)}`;
  applyFill(dashboardSheet.getCell("A3"), "#EEF4F8");
  dashboardSheet.getCell("A3").font = { bold: true, color: { argb: toArgb("#153147") }, size: 11 };
  dashboardSheet.getCell("A3").alignment = { horizontal: "center" };

  addKpiCard(dashboardSheet, "A5", "총 광고비", currency(overall.cost), `전체 클릭 ${overall.clicks.toLocaleString("ko-KR")}회`, "orange");
  addKpiCard(dashboardSheet, "C5", "총 노출수", overall.impressions.toLocaleString("ko-KR"), `CTR ${overall.ctr.toFixed(2)}%`, "navy");
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

  addMergedSectionTitle(dashboardSheet, "A10:N10", "광고주용 전달 요약");
  dashboardSheet.mergeCells("A11:N16");
  dashboardSheet.getCell("A11").value = narrative.join("\n");
  dashboardSheet.getCell("A11").alignment = { vertical: "top", horizontal: "left", wrapText: true };
  dashboardSheet.getCell("A11").font = { color: { argb: toArgb("#243B53") } };
  applyBorder(dashboardSheet.getCell("A11"));

  addMergedSectionTitle(dashboardSheet, "A18:G18", "매체별 전달 지표");
  writeTable(
    dashboardSheet,
    "A19",
    ["매체", "광고비", "구매완료 수", "구매완료 매출", "구매완료 ROAS", "구매완료 CPA", "CTR"],
    channelSummary.map((row) => [
      row.name,
      row.cost,
      row.purchaseConversions,
      row.purchaseRevenue,
      row.purchaseRoas / 100,
      row.purchaseCpa,
      row.ctr / 100,
    ]),
    {
      1: "#,##0",
      2: "#,##0",
      3: "#,##0",
      4: "0.0%",
      5: "#,##0",
      6: "0.0%",
    },
  );

  addMergedSectionTitle(dashboardSheet, "H18:N18", "상위 키워드");
  writeTable(
    dashboardSheet,
    "H19",
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

  addMergedSectionTitle(dashboardSheet, "A28:N28", "일자별 구매 흐름");
  writeTable(
    dashboardSheet,
    "A29",
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

  const overallRows = [
    ["지표", "값", "설명"],
    ["총 광고비", overall.cost, "파워링크 + 쇼핑검색 기준 총 집행 비용"],
    ["총 노출수", overall.impressions, "광고 노출 합계"],
    ["총 클릭수", overall.clicks, "광고 클릭 합계"],
    ["CTR", overall.ctr / 100, "클릭률"],
    ["평균 CPC", overall.cpc, "클릭당 평균 비용"],
    ["총 전환수", overall.conversions, "플랫폼 전체 전환 기준"],
    ["총 전환매출", overall.revenue, "플랫폼 전체 전환매출"],
    ["구매완료 수", overall.purchaseConversions, "광고주 기준 최종 전환"],
    ["구매완료 매출", overall.purchaseRevenue, "구매완료 전환매출"],
    ["구매완료 ROAS", overall.purchaseRoas / 100, "광고주 보고용 핵심 효율 지표"],
    ["구매완료 CPA", overall.purchaseCpa, "구매완료 1건당 비용"],
  ];
  overallRows.forEach((row, rowIndex) => {
    row.forEach((value, colIndex) => {
      overallSheet.getCell(rowIndex + 1, colIndex + 1).value = value;
    });
  });
  styleRange(overallSheet, 1, 1, 1, 3, (cell) => {
    applyFill(cell, "#12343B");
    applyBorder(cell, "C9D2DB");
    cell.font = { bold: true, color: { argb: toArgb("#FFFFFF") } };
    cell.alignment = { vertical: "middle", horizontal: "center" };
  });
  styleRange(overallSheet, 2, 1, overallRows.length, 3, (cell) => {
    applyBorder(cell);
    cell.alignment = { vertical: "middle", horizontal: typeof cell.value === "number" ? "right" : "left", wrapText: true };
  });
  overallSheet.getColumn(2).numFmt = "#,##0";
  overallSheet.getCell("B5").numFmt = "0.0%";
  overallSheet.getCell("B11").numFmt = "0.0%";

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

  writeTable(
    memoSheet,
    "A1",
    ["구분", "내용", "비고"],
    memoEntries.length
      ? memoEntries.map((entry) => ["운영 메모", entry, "직접 입력"])
      : [["운영 메모", "입력된 메모가 없습니다.", ""]],
  );

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

  configureSheetView(dashboardSheet, 3);
  configureSheetView(overallSheet, 1);
  configureSheetView(dailySheet, 1);
  configureSheetView(powerlinkSheet, 1);
  configureSheetView(shoppingSheet, 1);
  configureSheetView(keywordSheet, 1);
  configureSheetView(memoSheet, 1);
  configureSheetView(sourceSheet, 1);

  dashboardSheet.getColumn(1).width = 16;
  dashboardSheet.getColumn(7).width = 16;

  [
    dashboardSheet,
    overallSheet,
    dailySheet,
    powerlinkSheet,
    shoppingSheet,
    keywordSheet,
    memoSheet,
    sourceSheet,
  ].forEach((worksheet) => autoFitWorksheet(worksheet));

  const fileName = `${sanitizeSheetName(brand)}_주간보고서_${formatDateForFile(startDate)}_${formatDateForFile(endDate)}.xlsx`;
  const buffer = await workbook.xlsx.writeBuffer();

  return {
    fileName,
    bytes: Buffer.from(buffer),
    contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
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
