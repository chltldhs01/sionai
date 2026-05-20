import ExcelJS from "exceljs";
import sharp from "sharp";

const CHANNELS = ["파워링크", "쇼핑검색"];
const DETAIL_ROW_LIMIT = 200;
const CHART_COLORS = ["#4C8DF6", "#13A7A0", "#FF9F6E", "#8F7CFF"];

function parseCsv(csvText) {
  const rows = [];
  let field = "";
  let row = [];
  let inQuotes = false;

  for (let index = 0; index < csvText.length; index += 1) {
    const char = csvText[index];
    const next = csvText[index + 1];

    if (char === '"') {
      if (inQuotes && next === '"') {
        field += '"';
        index += 1;
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
        index += 1;
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
  const numeric = Number(cleaned);
  return Number.isFinite(numeric) ? numeric : 0;
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
  return `₩${Math.round(value || 0).toLocaleString("ko-KR")}`;
}

function formatPercent(value, digits = 2) {
  return `${Number(value || 0).toFixed(digits)}%`;
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
  return name.replace(/[\\/*?:[\]]/g, "_").slice(0, 24);
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

function pickField(row, names) {
  for (const name of names) {
    if (row[name] !== undefined) {
      return row[name];
    }
  }
  return "";
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
      const dateValue = pickField(row, ["일계", "일별", "날짜"]);
      const date = parseDateLabel(dateValue);
      if (!date) {
        return null;
      }

      return {
        date,
        dateLabel: formatDate(date),
        campaign: pickField(row, ["캠페인"]),
        campaignType: pickField(row, ["캠페인유형"]),
        adGroup: pickField(row, ["광고그룹"]),
        adGroupType: pickField(row, ["광고그룹유형"]),
        query: pickField(row, ["검색어"]),
        impressions: normalizeNumber(pickField(row, ["노출수"])),
        clicks: normalizeNumber(pickField(row, ["클릭수"])),
        ctrRaw: normalizeNumber(pickField(row, ["클릭률(%)"])),
        cpcRaw: normalizeNumber(pickField(row, ["평균 CPC"])),
        cost: normalizeNumber(pickField(row, ["총비용"])),
        conversions: normalizeNumber(pickField(row, ["총 전환수", "총전환수"])),
        directConversions: normalizeNumber(pickField(row, ["직접전환수"])),
        indirectConversions: normalizeNumber(pickField(row, ["간접전환수"])),
        conversionRateRaw: normalizeNumber(pickField(row, ["총 전환율(%)", "총전환율(%)"])),
        revenue: normalizeNumber(pickField(row, ["총 전환매출액(원)", "총전환매출액(원)"])),
        directRevenue: normalizeNumber(pickField(row, ["직접전환매출액(원)"])),
        indirectRevenue: normalizeNumber(pickField(row, ["간접전환매출액(원)"])),
        cpaRaw: normalizeNumber(pickField(row, ["총 전환당비용(원)", "총전환당비용(원)"])),
        roasRaw: normalizeNumber(pickField(row, ["총 광고수익률(%)", "총광고수익률(%)"])),
        purchaseConversions: normalizeNumber(pickField(row, ["구매완료 전환수", "구매완료전환수"])),
        purchaseRevenue: normalizeNumber(
          pickField(row, ["구매완료 전환매출액(원)", "구매완료전환매출액(원)"]),
        ),
        purchaseRoasRaw: normalizeNumber(
          pickField(row, ["구매완료 광고수익률(%)", "구매완료광고수익률(%)"]),
        ),
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
  }).filter((item) => item.cost > 0 || item.impressions > 0);
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
    .filter((item) => CHANNELS.includes(item.channel))
    .sort((left, right) => {
      const dateCompare = left.dateLabel.localeCompare(right.dateLabel);
      return dateCompare === 0 ? left.channel.localeCompare(right.channel) : dateCompare;
    });
}

function summarizeKeyword(records) {
  return Array.from(
    groupBy(
      records.filter((record) => CHANNELS.includes(record.campaignType)),
      (record) => `${record.query || "(검색어 없음)"}|||${record.campaignType}`,
    ).entries(),
  )
    .map(([key, groupedRecords]) => {
      const [query, channel] = key.split("|||");
      return {
        query,
        channel,
        ...sumMetrics(groupedRecords),
      };
    })
    .sort((left, right) => {
      if (right.purchaseRevenue !== left.purchaseRevenue) {
        return right.purchaseRevenue - left.purchaseRevenue;
      }
      return right.cost - left.cost;
    });
}

function summarizeCampaign(records, targetChannel) {
  return Array.from(
    groupBy(
      records.filter((record) => record.campaignType === targetChannel),
      (record) => `${record.campaign}|||${record.adGroup}`,
    ).entries(),
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

function buildNarrative(brand, startDate, endDate, overall, channelSummary, keywordSummary, memoEntries) {
  const topChannel = [...channelSummary].sort((left, right) => right.purchaseRevenue - left.purchaseRevenue)[0];
  const topKeyword = keywordSummary.find((item) => item.purchaseRevenue > 0);

  return [
    `${formatDate(startDate)}부터 ${formatDate(endDate)}까지 ${brand} 네이버 SA의 총 광고비는 ${currency(overall.cost)}, 구매완료 ${overall.purchaseConversions.toLocaleString(
      "ko-KR",
    )}건, 구매완료 매출 ${currency(overall.purchaseRevenue)}입니다.`,
    `구매완료 기준 핵심 효율 지표는 ROAS ${overall.purchaseRoas.toLocaleString("ko-KR")}%와 CPA ${currency(
      overall.purchaseCpa,
    )}입니다.`,
    topChannel
      ? `${topChannel.name}이 구매완료 매출 ${currency(topChannel.purchaseRevenue)}로 가장 큰 기여를 보였습니다.`
      : "활성 매체 데이터가 없어 매체별 비교는 제한적입니다.",
    topKeyword
      ? `상위 키워드는 ${topKeyword.query}이며 구매완료 매출 ${currency(topKeyword.purchaseRevenue)}를 기록했습니다.`
      : "구매완료 매출이 발생한 키워드가 없어 키워드 인사이트는 제한적입니다.",
    memoEntries.length
      ? `운영 메모 ${memoEntries.length}건이 함께 정리되어 수치 외 운영 배경까지 한 번에 전달할 수 있습니다.`
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

function applyBorder(cell, color = "DCE5EF") {
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
  applyFill(cell, "#17324A");
  cell.font = { bold: true, color: { argb: toArgb("#FFFFFF") }, size: 12 };
  cell.alignment = { vertical: "middle", horizontal: "left" };
}

function addKpiCard(worksheet, startColumn, title, value, note, tone = "blue") {
  const rowStart = 5;
  const rowEnd = 8;
  const fillMap = {
    blue: "#EFF6FF",
    teal: "#ECFCFA",
    warm: "#FFF4EC",
    violet: "#F4F1FF",
  };

  worksheet.mergeCells(rowStart, startColumn, rowStart + 1, startColumn + 2);
  worksheet.mergeCells(rowStart + 2, startColumn, rowStart + 2, startColumn + 2);
  worksheet.mergeCells(rowStart + 3, startColumn, rowStart + 3, startColumn + 2);

  styleRange(worksheet, rowStart, startColumn, rowEnd, startColumn + 2, (cell) => {
    applyFill(cell, fillMap[tone] || "#FFFFFF");
    applyBorder(cell);
    cell.alignment = { vertical: "middle", horizontal: "left", wrapText: true };
  });

  const titleCell = worksheet.getCell(rowStart, startColumn);
  titleCell.value = title;
  titleCell.font = { bold: true, color: { argb: toArgb("#5A6B7B") }, size: 10 };

  const valueCell = worksheet.getCell(rowStart + 2, startColumn);
  valueCell.value = value;
  valueCell.font = { bold: true, color: { argb: toArgb("#17324A") }, size: 18 };

  const noteCell = worksheet.getCell(rowStart + 3, startColumn);
  noteCell.value = note;
  noteCell.font = { color: { argb: toArgb("#5A6B7B") }, size: 10 };
}

function writeTable(worksheet, startRow, startCol, headers, rows, options = {}) {
  const {
    numberFormats = {},
    widths = [],
    zebra = true,
    filter = true,
    tableName = null,
  } = options;
  const safeRows = rows.length ? rows : [Array(headers.length).fill("")];
  const headerRow = startRow;
  const dataStartRow = startRow + 1;
  const endRow = dataStartRow + safeRows.length - 1;
  const endCol = startCol + headers.length - 1;

  headers.forEach((header, index) => {
    const cell = worksheet.getCell(headerRow, startCol + index);
    cell.value = header;
    applyFill(cell, "#17324A");
    applyBorder(cell, "C9D6E5");
    cell.font = { bold: true, color: { argb: toArgb("#FFFFFF") } };
    cell.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
  });
  worksheet.getRow(headerRow).height = 24;

  safeRows.forEach((currentRow, rowIndex) => {
    currentRow.forEach((value, cellIndex) => {
      const rowNumber = dataStartRow + rowIndex;
      const colNumber = startCol + cellIndex;
      const cell = worksheet.getCell(rowNumber, colNumber);
      cell.value = value;
      applyBorder(cell);
      if (zebra && rowIndex % 2 === 0) {
        applyFill(cell, "#F8FBFF");
      }
      cell.alignment = {
        vertical: "middle",
        horizontal: typeof value === "number" ? "right" : "left",
        wrapText: true,
      };
    });
    worksheet.getRow(dataStartRow + rowIndex).height = 22;
  });

  for (const [relativeIndex, format] of Object.entries(numberFormats)) {
    const colNumber = startCol + Number(relativeIndex);
    for (let rowNumber = dataStartRow; rowNumber <= endRow; rowNumber += 1) {
      worksheet.getCell(rowNumber, colNumber).numFmt = format;
    }
  }

  widths.forEach((width, index) => {
    if (width) {
      worksheet.getColumn(startCol + index).width = width;
    }
  });

  if (filter) {
    worksheet.autoFilter = {
      from: { row: headerRow, column: startCol },
      to: { row: endRow, column: endCol },
    };
  }

  if (tableName) {
    worksheet.getCell(headerRow - 1, startCol).value = tableName;
    worksheet.getCell(headerRow - 1, startCol).font = { bold: true, size: 12, color: { argb: toArgb("#17324A") } };
  }

  return { endRow, endCol };
}

function autoFitWorksheet(worksheet, minWidth = 12, maxWidth = 30, maxRowsToScan = 500) {
  worksheet.columns.forEach((column) => {
    if (column.width && column.width > minWidth) {
      return;
    }

    let maxLength = minWidth;
    let scannedRows = 0;

    column.eachCell({ includeEmpty: true }, (cell) => {
      if (scannedRows >= maxRowsToScan) {
        return;
      }
      scannedRows += 1;

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
  worksheet.properties.defaultRowHeight = 22;
}

function escapeXml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&apos;");
}

async function renderSvgToPng(svg, width, height) {
  return sharp(Buffer.from(svg)).png().resize(width, height).toBuffer();
}

function buildDailyComboChartSvg(dailySummary) {
  const width = 620;
  const height = 320;
  const padding = { top: 34, right: 64, bottom: 42, left: 64 };
  const innerWidth = width - padding.left - padding.right;
  const innerHeight = height - padding.top - padding.bottom;
  const maxCost = Math.max(...dailySummary.map((item) => item.cost), 1);
  const maxCtr = Math.max(...dailySummary.map((item) => item.ctr), 1);
  const stepX = dailySummary.length > 1 ? innerWidth / (dailySummary.length - 1) : 0;
  const barWidth = Math.max(12, Math.min(28, innerWidth / Math.max(dailySummary.length * 1.5, 8)));

  const linePoints = dailySummary
    .map((item, index) => {
      const x = padding.left + stepX * index;
      const y = padding.top + innerHeight - (item.ctr / maxCtr) * innerHeight;
      return `${x},${y}`;
    })
    .join(" ");

  const areaPoints = `${padding.left},${padding.top + innerHeight} ${linePoints} ${
    padding.left + stepX * (dailySummary.length - 1)
  },${padding.top + innerHeight}`;

  const grid = [0, 0.25, 0.5, 0.75, 1]
    .map((ratio) => {
      const y = padding.top + innerHeight - innerHeight * ratio;
      const costValue = Math.round(maxCost * ratio);
      const ctrValue = (maxCtr * ratio).toFixed(1);
      return `
        <line x1="${padding.left}" y1="${y}" x2="${width - padding.right}" y2="${y}" stroke="#E2EAF3" />
        <text x="${padding.left - 12}" y="${y + 4}" text-anchor="end" font-size="11" fill="#7B8DA1">${costValue === 0 ? "0" : `${Math.round(costValue / 10000)}만`}</text>
        <text x="${width - padding.right + 12}" y="${y + 4}" font-size="11" fill="#7B8DA1">${ctrValue}%</text>
      `;
    })
    .join("");

  const bars = dailySummary
    .map((item, index) => {
      const x = padding.left + stepX * index - barWidth / 2;
      const barHeight = (item.cost / maxCost) * innerHeight;
      const y = padding.top + innerHeight - barHeight;
      return `<rect x="${x.toFixed(1)}" y="${y.toFixed(1)}" width="${barWidth.toFixed(1)}" height="${barHeight.toFixed(
        1,
      )}" rx="8" fill="#CFE2FF" />`;
    })
    .join("");

  const dots = dailySummary
    .map((item, index) => {
      const x = padding.left + stepX * index;
      const y = padding.top + innerHeight - (item.ctr / maxCtr) * innerHeight;
      return `<circle cx="${x.toFixed(1)}" cy="${y.toFixed(1)}" r="4.5" fill="#FFFFFF" stroke="#13A7A0" stroke-width="3" />`;
    })
    .join("");

  const labels = dailySummary
    .map((item, index) => {
      if (dailySummary.length > 10 && index % 2 === 1 && index !== dailySummary.length - 1) {
        return "";
      }
      const x = padding.left + stepX * index;
      return `<text x="${x.toFixed(1)}" y="${height - 14}" text-anchor="middle" font-size="11" fill="#66788B">${escapeXml(
        item.dateLabel.slice(5),
      )}</text>`;
    })
    .join("");

  return `
    <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${width} ${height}">
      <defs>
        <linearGradient id="dailyArea" x1="0" y1="0" x2="0" y2="1">
          <stop offset="0%" stop-color="#8BE4DB" stop-opacity="0.38" />
          <stop offset="100%" stop-color="#8BE4DB" stop-opacity="0.02" />
        </linearGradient>
      </defs>
      <rect width="${width}" height="${height}" rx="22" fill="#FFFFFF" />
      <text x="${padding.left}" y="20" font-size="15" font-weight="700" fill="#17324A">일자별 광고비 / CTR 추이</text>
      <text x="${padding.left}" y="40" font-size="11" fill="#66788B">막대는 광고비, 선은 CTR을 의미합니다.</text>
      ${grid}
      ${bars}
      <polygon points="${areaPoints}" fill="url(#dailyArea)" />
      <polyline points="${linePoints}" fill="none" stroke="#13A7A0" stroke-width="4" stroke-linecap="round" stroke-linejoin="round" />
      ${dots}
      ${labels}
    </svg>
  `;
}

function buildChannelDonutSvg(channelSummary) {
  const size = 440;
  const center = 160;
  const radius = 90;
  const stroke = 34;
  const circumference = 2 * Math.PI * radius;
  const totalRevenue = channelSummary.reduce((sum, item) => sum + item.purchaseRevenue, 0);
  let offset = 0;

  const donutSlices = channelSummary
    .map((item, index) => {
      const ratio = totalRevenue ? item.purchaseRevenue / totalRevenue : 0;
      const dash = ratio * circumference;
      const currentOffset = offset;
      offset += dash;
      return `
        <circle
          cx="${center}"
          cy="${center}"
          r="${radius}"
          fill="none"
          stroke="${CHART_COLORS[index % CHART_COLORS.length]}"
          stroke-width="${stroke}"
          stroke-dasharray="${dash} ${circumference - dash}"
          stroke-dashoffset="${-currentOffset}"
          stroke-linecap="round"
          transform="rotate(-90 ${center} ${center})"
        />
      `;
    })
    .join("");

  const legends = channelSummary
    .map((item, index) => {
      const ratio = totalRevenue ? (item.purchaseRevenue / totalRevenue) * 100 : 0;
      const y = 72 + index * 54;
      return `
        <rect x="262" y="${y}" width="14" height="14" rx="7" fill="${CHART_COLORS[index % CHART_COLORS.length]}" />
        <text x="286" y="${y + 12}" font-size="13" font-weight="700" fill="#17324A">${escapeXml(item.name)}</text>
        <text x="286" y="${y + 30}" font-size="12" fill="#66788B">${escapeXml(currency(item.purchaseRevenue))} · ${ratio.toFixed(1)}%</text>
      `;
    })
    .join("");

  return `
    <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${size} 320">
      <rect width="${size}" height="320" rx="22" fill="#FFFFFF" />
      <text x="24" y="24" font-size="15" font-weight="700" fill="#17324A">매체별 구매완료 매출 비중</text>
      <circle cx="${center}" cy="${center}" r="${radius}" fill="none" stroke="#E7EEF6" stroke-width="${stroke}" />
      ${donutSlices}
      <text x="${center}" y="${center - 6}" text-anchor="middle" font-size="14" fill="#66788B">총 매출</text>
      <text x="${center}" y="${center + 24}" text-anchor="middle" font-size="24" font-weight="700" fill="#17324A">${escapeXml(
        currency(totalRevenue),
      )}</text>
      ${legends}
    </svg>
  `;
}

function buildInsightRows(narrative) {
  return narrative.map((text, index) => [`인사이트 ${index + 1}`, text]);
}

function addDashboardHeader(worksheet, brand, startDate, endDate) {
  worksheet.mergeCells("A1:N2");
  const titleCell = worksheet.getCell("A1");
  titleCell.value = `${brand} 네이버 SA 주간보고서`;
  applyFill(titleCell, "#17324A");
  titleCell.font = { bold: true, color: { argb: toArgb("#FFFFFF") }, size: 20 };
  titleCell.alignment = { vertical: "middle", horizontal: "center" };

  worksheet.mergeCells("A3:N3");
  const periodCell = worksheet.getCell("A3");
  periodCell.value = `보고 기간: ${formatDate(startDate)} ~ ${formatDate(endDate)}`;
  applyFill(periodCell, "#EAF3FF");
  periodCell.font = { bold: true, color: { argb: toArgb("#17324A") }, size: 11 };
  periodCell.alignment = { vertical: "middle", horizontal: "center" };
}

function addNumberFormatsToSheets(sheets) {
  for (const sheet of sheets) {
    configureSheetView(sheet, 1);
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

  const dashboardSheet = workbook.addWorksheet("대시보드");
  const overallSheet = workbook.addWorksheet("전체 요약");
  const dailySheet = workbook.addWorksheet("일자별 추이");
  const powerlinkSheet = workbook.addWorksheet("파워링크 상세");
  const shoppingSheet = workbook.addWorksheet("쇼핑검색 상세");
  const keywordSheet = workbook.addWorksheet("키워드 상세");
  const memoSheet = workbook.addWorksheet("운영 메모");

  dashboardSheet.columns = Array.from({ length: 14 }, () => ({ width: 14 }));
  addDashboardHeader(dashboardSheet, brand, startDate, endDate);
  addKpiCard(dashboardSheet, 1, "총 광고비", currency(overall.cost), `총 클릭 ${overall.clicks.toLocaleString("ko-KR")}회`, "warm");
  addKpiCard(dashboardSheet, 4, "총 노출수", overall.impressions.toLocaleString("ko-KR"), `CTR ${formatPercent(overall.ctr)}`, "blue");
  addKpiCard(
    dashboardSheet,
    7,
    "구매완료 전환수",
    `${overall.purchaseConversions.toLocaleString("ko-KR")}건`,
    `구매완료 CVR ${formatPercent(overall.purchaseCvr)}`,
    "teal",
  );
  addKpiCard(
    dashboardSheet,
    10,
    "구매완료 매출",
    currency(overall.purchaseRevenue),
    `구매완료 ROAS ${overall.purchaseRoas.toLocaleString("ko-KR")}%`,
    "violet",
  );

  addMergedSectionTitle(dashboardSheet, "A10:F10", "광고주용 핵심 인사이트");
  dashboardSheet.mergeCells("A11:F16");
  const insightCell = dashboardSheet.getCell("A11");
  insightCell.value = narrative.join("\n");
  insightCell.alignment = { vertical: "top", horizontal: "left", wrapText: true };
  insightCell.font = { color: { argb: toArgb("#243B53") }, size: 11 };
  applyBorder(insightCell);
  applyFill(insightCell, "#FFFFFF");

  addMergedSectionTitle(dashboardSheet, "G10:N10", "운영 메모 요약");
  dashboardSheet.mergeCells("G11:N16");
  const memoCell = dashboardSheet.getCell("G11");
  memoCell.value = memoEntries.length ? memoEntries.join("\n") : "입력된 운영 메모가 없습니다.";
  memoCell.alignment = { vertical: "top", horizontal: "left", wrapText: true };
  memoCell.font = { color: { argb: toArgb("#243B53") }, size: 11 };
  applyBorder(memoCell);
  applyFill(memoCell, "#FFFFFF");

  const dailyChartBuffer = await renderSvgToPng(buildDailyComboChartSvg(dailySummary), 620, 320);
  const dailyChartId = workbook.addImage({ buffer: dailyChartBuffer, extension: "png" });
  dashboardSheet.addImage(dailyChartId, {
    tl: { col: 0, row: 17 },
    ext: { width: 620, height: 320 },
  });

  const donutChartBuffer = await renderSvgToPng(buildChannelDonutSvg(channelSummary), 420, 320);
  const donutChartId = workbook.addImage({ buffer: donutChartBuffer, extension: "png" });
  dashboardSheet.addImage(donutChartId, {
    tl: { col: 8.4, row: 17 },
    ext: { width: 420, height: 320 },
  });

  addMergedSectionTitle(dashboardSheet, "A34:G34", "매체별 성과 비교");
  writeTable(
    dashboardSheet,
    35,
    1,
    ["매체", "광고비", "구매완료 전환수", "구매완료 매출", "구매완료 ROAS", "구매완료 CPA", "CTR"],
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
      numberFormats: {
        2: '"₩"#,##0',
        3: "#,##0",
        4: '"₩"#,##0',
        5: "0.0%",
        6: '"₩"#,##0',
        7: "0.0%",
      },
      widths: [16, 16, 16, 18, 14, 16, 12],
      filter: false,
    },
  );

  addMergedSectionTitle(dashboardSheet, "H34:N34", "상위 키워드");
  writeTable(
    dashboardSheet,
    35,
    8,
    ["검색어", "매체", "광고비", "구매완료 전환수", "구매완료 매출", "구매완료 ROAS", "CTR"],
    keywordSummary.slice(0, 6).map((row) => [
      row.query,
      row.channel,
      row.cost,
      row.purchaseConversions,
      row.purchaseRevenue,
      row.purchaseRoas / 100,
      row.ctr / 100,
    ]),
    {
      numberFormats: {
        3: '"₩"#,##0',
        4: "#,##0",
        5: '"₩"#,##0',
        6: "0.0%",
        7: "0.0%",
      },
      widths: [24, 14, 16, 16, 18, 14, 12],
      filter: false,
    },
  );

  addMergedSectionTitle(dashboardSheet, "A44:N44", "일자별 핵심 수치");
  writeTable(
    dashboardSheet,
    45,
    1,
    ["일자", "광고비", "총 클릭수", "구매완료 전환수", "구매완료 매출", "구매완료 ROAS", "CTR"],
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
      numberFormats: {
        2: '"₩"#,##0',
        3: "#,##0",
        4: "#,##0",
        5: '"₩"#,##0',
        6: "0.0%",
        7: "0.0%",
      },
      widths: [14, 16, 14, 16, 18, 14, 12],
      filter: false,
    },
  );

  const overallRows = [
    ["지표", "값", "설명"],
    ["총 광고비", overall.cost, "파워링크와 쇼핑검색 기준 총 집행 비용"],
    ["총 노출수", overall.impressions, "광고 노출 합계"],
    ["총 클릭수", overall.clicks, "광고 클릭 합계"],
    ["CTR", overall.ctr / 100, "클릭률"],
    ["평균 CPC", overall.cpc, "클릭당 평균 비용"],
    ["총 전환수", overall.conversions, "플랫폼 기준 전체 전환"],
    ["총 전환매출", overall.revenue, "플랫폼 기준 전체 전환매출"],
    ["구매완료 전환수", overall.purchaseConversions, "광고주 기준 최종 전환"],
    ["구매완료 매출", overall.purchaseRevenue, "광고주 기준 전환매출"],
    ["구매완료 ROAS", overall.purchaseRoas / 100, "광고주 보고 기준 효율 지표"],
    ["구매완료 CPA", overall.purchaseCpa, "구매완료 1건당 비용"],
  ];
  overallRows.forEach((row, rowIndex) => {
    row.forEach((value, colIndex) => {
      overallSheet.getCell(rowIndex + 1, colIndex + 1).value = value;
    });
  });
  styleRange(overallSheet, 1, 1, 1, 3, (cell) => {
    applyFill(cell, "#17324A");
    applyBorder(cell, "C9D6E5");
    cell.font = { bold: true, color: { argb: toArgb("#FFFFFF") } };
    cell.alignment = { vertical: "middle", horizontal: "center" };
  });
  styleRange(overallSheet, 2, 1, overallRows.length, 3, (cell, row, col) => {
    applyBorder(cell);
    if ((row - 2) % 2 === 0) {
      applyFill(cell, "#F8FBFF");
    }
    cell.alignment = { vertical: "middle", horizontal: col === 2 ? "right" : "left", wrapText: true };
  });
  overallSheet.getColumn(1).width = 18;
  overallSheet.getColumn(2).width = 18;
  overallSheet.getColumn(3).width = 34;
  overallSheet.getCell("B5").numFmt = "0.0%";
  overallSheet.getCell("B6").numFmt = '"₩"#,##0';
  overallSheet.getCell("B8").numFmt = '"₩"#,##0';
  overallSheet.getCell("B10").numFmt = '"₩"#,##0';
  overallSheet.getCell("B11").numFmt = "0.0%";
  overallSheet.getCell("B12").numFmt = '"₩"#,##0';

  writeTable(
    overallSheet,
    1,
    5,
    ["매체", "광고비", "구매완료 전환수", "구매완료 매출", "구매완료 ROAS", "구매완료 CPA", "CTR", "CVR"],
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
      numberFormats: {
        2: '"₩"#,##0',
        3: "#,##0",
        4: '"₩"#,##0',
        5: "0.0%",
        6: '"₩"#,##0',
        7: "0.0%",
        8: "0.0%",
      },
      widths: [14, 16, 16, 18, 14, 16, 12, 12],
    },
  );

  writeTable(
    overallSheet,
    12,
    5,
    ["항목", "내용"],
    buildInsightRows(narrative),
    {
      widths: [14, 58],
      filter: false,
    },
  );

  writeTable(
    dailySheet,
    1,
    1,
    ["일자", "광고비", "총 클릭수", "구매완료 전환수", "구매완료 매출", "구매완료 ROAS", "CTR"],
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
      numberFormats: {
        2: '"₩"#,##0',
        3: "#,##0",
        4: "#,##0",
        5: '"₩"#,##0',
        6: "0.0%",
        7: "0.0%",
      },
      widths: [14, 16, 14, 16, 18, 14, 12],
    },
  );

  writeTable(
    dailySheet,
    1,
    10,
    ["일자", "매체", "광고비", "구매완료 전환수", "구매완료 매출", "구매완료 ROAS"],
    dailyChannelSummary.map((row) => [
      row.dateLabel,
      row.channel,
      row.cost,
      row.purchaseConversions,
      row.purchaseRevenue,
      row.purchaseRoas / 100,
    ]),
    {
      numberFormats: {
        3: '"₩"#,##0',
        4: "#,##0",
        5: '"₩"#,##0',
        6: "0.0%",
      },
      widths: [14, 14, 16, 16, 18, 14],
    },
  );

  writeTable(
    powerlinkSheet,
    1,
    1,
    ["캠페인", "광고그룹", "광고비", "클릭수", "구매완료 전환수", "구매완료 매출", "구매완료 ROAS", "구매완료 CPA"],
    powerlinkSummary.slice(0, DETAIL_ROW_LIMIT).map((row) => [
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
      numberFormats: {
        3: '"₩"#,##0',
        4: "#,##0",
        5: "#,##0",
        6: '"₩"#,##0',
        7: "0.0%",
        8: '"₩"#,##0',
      },
      widths: [30, 24, 16, 12, 16, 18, 14, 16],
    },
  );

  writeTable(
    shoppingSheet,
    1,
    1,
    ["캠페인", "광고그룹", "광고비", "클릭수", "구매완료 전환수", "구매완료 매출", "구매완료 ROAS", "구매완료 CPA"],
    shoppingSummary.slice(0, DETAIL_ROW_LIMIT).map((row) => [
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
      numberFormats: {
        3: '"₩"#,##0',
        4: "#,##0",
        5: "#,##0",
        6: '"₩"#,##0',
        7: "0.0%",
        8: '"₩"#,##0',
      },
      widths: [30, 24, 16, 12, 16, 18, 14, 16],
    },
  );

  writeTable(
    keywordSheet,
    1,
    1,
    ["검색어", "매체", "광고비", "클릭수", "구매완료 전환수", "구매완료 매출", "구매완료 ROAS", "구매완료 CPA"],
    keywordSummary.slice(0, DETAIL_ROW_LIMIT).map((row) => [
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
      numberFormats: {
        3: '"₩"#,##0',
        4: "#,##0",
        5: "#,##0",
        6: '"₩"#,##0',
        7: "0.0%",
        8: '"₩"#,##0',
      },
      widths: [24, 14, 16, 12, 16, 18, 14, 16],
    },
  );

  writeTable(
    memoSheet,
    1,
    1,
    ["구분", "내용"],
    memoEntries.length
      ? memoEntries.map((entry) => ["운영 메모", entry])
      : [["운영 메모", "입력된 메모가 없습니다."]],
    {
      widths: [16, 72],
      filter: false,
    },
  );

  configureSheetView(dashboardSheet, 4);
  addNumberFormatsToSheets([overallSheet, dailySheet, powerlinkSheet, shoppingSheet, keywordSheet, memoSheet]);

  autoFitWorksheet(dashboardSheet, 12, 24, 200);
  autoFitWorksheet(overallSheet, 12, 32, 200);
  autoFitWorksheet(dailySheet, 12, 24, 240);
  autoFitWorksheet(powerlinkSheet, 12, 26, 240);
  autoFitWorksheet(shoppingSheet, 12, 26, 240);
  autoFitWorksheet(keywordSheet, 12, 24, 240);
  autoFitWorksheet(memoSheet, 14, 60, 180);

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
      daily: dailySummary.map((entry) => ({
        dateLabel: entry.dateLabel,
        cost: entry.cost,
        clicks: entry.clicks,
        ctr: entry.ctr,
        conversions: entry.purchaseConversions,
        revenue: entry.purchaseRevenue,
        roas: entry.purchaseRoas,
      })),
      topKeywords: keywordSummary.slice(0, 6).map((entry) => ({
        query: entry.query,
        channel: entry.channel,
        cost: entry.cost,
        conversions: entry.purchaseConversions,
        revenue: entry.purchaseRevenue,
        roas: entry.purchaseRoas,
        ctr: entry.ctr,
      })),
      memoCount: memoEntries.length,
      narrative,
    },
  };
}
