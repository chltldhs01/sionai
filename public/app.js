const form = document.querySelector("#report-form");
const resultPanel = document.querySelector("#result-panel");
const downloadLink = document.querySelector("#download-link");
const submitButton = document.querySelector("#submit-button");
const kpiGrid = document.querySelector("#kpi-grid");
const channelBars = document.querySelector("#channel-bars");
const keywordList = document.querySelector("#keyword-list");
const trendChart = document.querySelector("#trend-chart");
const donutChart = document.querySelector("#donut-chart");
const insightList = document.querySelector("#insight-list");
const summaryBrand = document.querySelector("#summary-brand");
const summaryRange = document.querySelector("#summary-range");
const summaryMemoCount = document.querySelector("#summary-memo-count");

const chartPalette = ["#4c8df6", "#13a7a0", "#ff9f6e", "#8f7cff"];

function formatNumber(value) {
  return Number(value || 0).toLocaleString("ko-KR");
}

function formatCurrency(value) {
  return `₩${formatNumber(Math.round(value || 0))}`;
}

function formatPercent(value, digits = 2) {
  return `${Number(value || 0).toFixed(digits)}%`;
}

function buildKpis(summary) {
  return [
    {
      label: "총 광고비",
      value: summary.spend,
      note: `총 클릭 ${formatNumber(summary.raw.clicks)}회`,
      tone: "warm",
    },
    {
      label: "총 노출수",
      value: formatNumber(summary.raw.impressions),
      note: `CTR ${formatPercent(summary.raw.ctr)}`,
      tone: "cool",
    },
    {
      label: "구매완료 전환수",
      value: `${formatNumber(summary.conversions)}건`,
      note: `구매완료 CVR ${formatPercent(summary.raw.cvr)}`,
      tone: "teal",
    },
    {
      label: "구매완료 매출",
      value: summary.revenue,
      note: `구매완료 ROAS ${summary.roas}`,
      tone: "accent",
    },
    {
      label: "평균 CPC",
      value: formatCurrency(summary.raw.cpc),
      note: "클릭당 평균 비용",
      tone: "cool",
    },
    {
      label: "구매완료 CPA",
      value: formatCurrency(summary.raw.cpa),
      note: "전환 1건당 비용",
      tone: "warm",
    },
    {
      label: "총 클릭수",
      value: formatNumber(summary.raw.clicks),
      note: "광고 반응량",
      tone: "teal",
    },
    {
      label: "매체 수",
      value: `${summary.channels.length}개`,
      note: "활성 매체 기준",
      tone: "accent",
    },
  ];
}

function renderKpis(summary) {
  kpiGrid.innerHTML = "";

  for (const item of buildKpis(summary)) {
    const card = document.createElement("article");
    card.className = `kpi-card tone-${item.tone}`;
    card.innerHTML = `
      <span class="kpi-label">${item.label}</span>
      <strong class="kpi-value">${item.value}</strong>
      <span class="kpi-note">${item.note}</span>
    `;
    kpiGrid.append(card);
  }
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function createTrendSvg(daily) {
  if (!daily.length) {
    return '<div class="chart-caption">표시할 일자별 데이터가 없습니다.</div>';
  }

  const width = 860;
  const height = 320;
  const padding = { top: 24, right: 58, bottom: 40, left: 58 };
  const innerWidth = width - padding.left - padding.right;
  const innerHeight = height - padding.top - padding.bottom;
  const maxCost = Math.max(...daily.map((item) => item.cost), 1);
  const maxCtr = Math.max(...daily.map((item) => item.ctr), 1);
  const stepX = daily.length > 1 ? innerWidth / (daily.length - 1) : 0;

  const barWidth = Math.max(12, Math.min(28, innerWidth / Math.max(daily.length * 1.5, 8)));
  const linePoints = daily
    .map((item, index) => {
      const x = padding.left + stepX * index;
      const y = padding.top + innerHeight - (item.ctr / maxCtr) * innerHeight;
      return `${x},${y}`;
    })
    .join(" ");

  const areaPoints = `${padding.left},${padding.top + innerHeight} ${linePoints} ${padding.left + stepX * (daily.length - 1)},${padding.top + innerHeight}`;

  const bars = daily
    .map((item, index) => {
      const x = padding.left + stepX * index - barWidth / 2;
      const barHeight = (item.cost / maxCost) * innerHeight;
      const y = padding.top + innerHeight - barHeight;
      return `<rect x="${x.toFixed(1)}" y="${y.toFixed(1)}" width="${barWidth.toFixed(1)}" height="${barHeight.toFixed(
        1,
      )}" rx="8" fill="rgba(76,141,246,0.22)" />`;
    })
    .join("");

  const labels = daily
    .map((item, index) => {
      if (daily.length > 10 && index % 2 === 1 && index !== daily.length - 1) {
        return "";
      }
      const x = padding.left + stepX * index;
      return `<text x="${x.toFixed(1)}" y="${height - 14}" text-anchor="middle" font-size="11" fill="#66788b">${escapeHtml(
        item.dateLabel.slice(5),
      )}</text>`;
    })
    .join("");

  const valueLabels = daily
    .map((item, index) => {
      if (daily.length > 12 && index % 3 !== 0 && index !== daily.length - 1) {
        return "";
      }
      const x = padding.left + stepX * index;
      const y = padding.top + innerHeight - (item.ctr / maxCtr) * innerHeight - 10;
      return `<text x="${x.toFixed(1)}" y="${y.toFixed(1)}" text-anchor="middle" font-size="11" fill="#2f6fdb">${formatPercent(
        item.ctr,
      )}</text>`;
    })
    .join("");

  const horizontalGrid = [0, 0.25, 0.5, 0.75, 1]
    .map((ratio) => {
      const y = padding.top + innerHeight - innerHeight * ratio;
      const costValue = Math.round(maxCost * ratio);
      const ctrValue = (maxCtr * ratio).toFixed(1);
      return `
        <line x1="${padding.left}" y1="${y}" x2="${width - padding.right}" y2="${y}" stroke="rgba(24,50,74,0.08)" />
        <text x="${padding.left - 12}" y="${y + 4}" text-anchor="end" font-size="11" fill="#8aa">${costValue === 0 ? "0" : `${Math.round(costValue / 10000)}만`}</text>
        <text x="${width - padding.right + 12}" y="${y + 4}" font-size="11" fill="#8aa">${ctrValue}%</text>
      `;
    })
    .join("");

  return `
    <svg viewBox="0 0 ${width} ${height}" role="img" aria-label="일자별 광고비와 CTR 추이 차트">
      <defs>
        <linearGradient id="lineFill" x1="0" y1="0" x2="0" y2="1">
          <stop offset="0%" stop-color="rgba(19,167,160,0.28)" />
          <stop offset="100%" stop-color="rgba(19,167,160,0.02)" />
        </linearGradient>
      </defs>
      <rect width="${width}" height="${height}" fill="transparent" />
      ${horizontalGrid}
      ${bars}
      <polygon points="${areaPoints}" fill="url(#lineFill)" />
      <polyline points="${linePoints}" fill="none" stroke="#13a7a0" stroke-width="4" stroke-linecap="round" stroke-linejoin="round" />
      ${daily
        .map((item, index) => {
          const x = padding.left + stepX * index;
          const y = padding.top + innerHeight - (item.ctr / maxCtr) * innerHeight;
          return `<circle cx="${x.toFixed(1)}" cy="${y.toFixed(1)}" r="4.5" fill="#ffffff" stroke="#13a7a0" stroke-width="3" />`;
        })
        .join("")}
      ${labels}
      ${valueLabels}
      <text x="${padding.left}" y="18" font-size="12" fill="#66788b">막대: 광고비</text>
      <text x="${padding.left + 90}" y="18" font-size="12" fill="#66788b">선: CTR</text>
    </svg>
  `;
}

function renderInsights(summary) {
  insightList.innerHTML = "";
  const items = [
    {
      title: "핵심 성과",
      body: `${summary.reportRange} 동안 구매완료 ${formatNumber(summary.conversions)}건, 구매완료 매출 ${summary.revenue}, 구매완료 ROAS ${summary.roas}를 기록했습니다.`,
    },
    {
      title: "매체 포인트",
      body:
        summary.channels.length > 0
          ? `${summary.channels[0].name} 포함 총 ${summary.channels.length}개 매체 데이터가 집계되었고, 각 매체별 광고비와 매출 비중을 아래에서 바로 비교할 수 있습니다.`
          : "활성 매체 데이터가 없습니다.",
    },
    {
      title: "운영 메모 반영",
      body:
        summary.memoCount > 0
          ? `운영 메모 ${summary.memoCount}건이 보고서에 반영되어, 단순 수치 외에도 운영 이슈를 함께 전달할 수 있습니다.`
          : "운영 메모가 없어 수치 중심 보고서로 생성됩니다.",
    },
  ];

  for (const item of items) {
    const card = document.createElement("article");
    card.className = "insight-item";
    card.innerHTML = `
      <strong>${item.title}</strong>
      <p>${item.body}</p>
    `;
    insightList.append(card);
  }
}

function renderChannels(summary) {
  channelBars.innerHTML = "";
  const maxRevenue = Math.max(...summary.channels.map((item) => item.revenue), 1);
  const maxCost = Math.max(...summary.channels.map((item) => item.cost), 1);

  for (const channel of summary.channels) {
    const row = document.createElement("article");
    row.className = "channel-row";
    row.innerHTML = `
      <div class="channel-top">
        <div>
          <strong>${channel.name}</strong>
          <span>구매완료 ROAS ${formatPercent(channel.roas, 0)}</span>
        </div>
        <div class="channel-metrics">
          <span>광고비 ${formatCurrency(channel.cost)}</span>
          <span>매출 ${formatCurrency(channel.revenue)}</span>
        </div>
      </div>
      <div class="bar-stack">
        <div class="bar-track">
          <div class="bar bar-cost" style="width:${Math.max((channel.cost / maxCost) * 100, 6)}%"></div>
        </div>
        <div class="bar-track">
          <div class="bar bar-revenue" style="width:${Math.max((channel.revenue / maxRevenue) * 100, 6)}%"></div>
        </div>
      </div>
      <div class="bar-legend">
        <span><i class="dot cost"></i>광고비</span>
        <span><i class="dot revenue"></i>구매완료 매출</span>
      </div>
    `;
    channelBars.append(row);
  }
}

function createDonutSvg(channels) {
  const total = channels.reduce((sum, item) => sum + item.revenue, 0);
  if (!total) {
    return '<div class="chart-caption">매출 비중을 표시할 데이터가 없습니다.</div>';
  }

  const size = 340;
  const radius = 92;
  const stroke = 34;
  const center = size / 2;
  const circumference = 2 * Math.PI * radius;
  let offset = 0;

  const circles = channels
    .map((channel, index) => {
      const value = channel.revenue / total;
      const dash = value * circumference;
      const node = `
        <circle
          cx="${center}"
          cy="${center}"
          r="${radius}"
          fill="none"
          stroke="${chartPalette[index % chartPalette.length]}"
          stroke-width="${stroke}"
          stroke-dasharray="${dash} ${circumference - dash}"
          stroke-dashoffset="${-offset}"
          stroke-linecap="round"
          transform="rotate(-90 ${center} ${center})"
        />
      `;
      offset += dash;
      return node;
    })
    .join("");

  const legend = channels
    .map((channel, index) => {
      const ratio = channel.revenue / total;
      return `
        <div class="insight-item">
          <strong style="display:flex; align-items:center; gap:10px;">
            <span style="display:inline-block; width:12px; height:12px; border-radius:999px; background:${chartPalette[index % chartPalette.length]}"></span>
            ${channel.name}
          </strong>
          <p>${formatCurrency(channel.revenue)} · ${formatPercent(ratio * 100, 1)}</p>
        </div>
      `;
    })
    .join("");

  return `
    <div style="display:grid; gap:16px; align-items:center; min-height:320px; grid-template-columns:minmax(0, 1fr); padding:20px;">
      <svg viewBox="0 0 ${size} ${size}" aria-label="매체별 매출 비중 도넛 차트">
        <circle cx="${center}" cy="${center}" r="${radius}" fill="none" stroke="rgba(24,50,74,0.08)" stroke-width="${stroke}" />
        ${circles}
        <text x="${center}" y="${center - 8}" text-anchor="middle" font-size="16" fill="#66788b">총 매출</text>
        <text x="${center}" y="${center + 24}" text-anchor="middle" font-size="28" font-weight="700" fill="#18324a">${escapeHtml(
          formatCurrency(total),
        )}</text>
      </svg>
      <div style="display:grid; gap:12px;">${legend}</div>
    </div>
  `;
}

function renderKeywords(summary) {
  keywordList.innerHTML = "";

  for (const keyword of summary.topKeywords) {
    const item = document.createElement("article");
    item.className = "keyword-card";
    item.innerHTML = `
      <div class="keyword-head">
        <strong>${keyword.query}</strong>
        <span>${keyword.channel}</span>
      </div>
      <dl>
        <div>
          <dt>광고비</dt>
          <dd>${formatCurrency(keyword.cost)}</dd>
        </div>
        <div>
          <dt>구매완료 전환수</dt>
          <dd>${formatNumber(keyword.conversions)}건</dd>
        </div>
        <div>
          <dt>구매완료 매출</dt>
          <dd>${formatCurrency(keyword.revenue)}</dd>
        </div>
        <div>
          <dt>구매완료 ROAS</dt>
          <dd>${formatPercent(keyword.roas, 0)}</dd>
        </div>
      </dl>
    `;
    keywordList.append(item);
  }
}

function renderCharts(summary) {
  trendChart.innerHTML = createTrendSvg(summary.daily || []);
  donutChart.innerHTML = createDonutSvg(summary.channels || []);
}

form.addEventListener("submit", async (event) => {
  event.preventDefault();

  const fileInput = document.querySelector("#csv-file");
  const file = fileInput.files?.[0];
  if (!file) {
    alert("CSV 파일을 선택해 주세요.");
    return;
  }

  submitButton.disabled = true;
  submitButton.textContent = "생성 중...";

  try {
    const csvText = await file.text();
    const payload = {
      brand: document.querySelector("#brand").value.trim() || "바이맘",
      reportStartDate: document.querySelector("#start-date").value,
      reportEndDate: document.querySelector("#end-date").value,
      memoText: document.querySelector("#memo-text").value,
      csvText,
    };

    const response = await fetch("/api/generate", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify(payload),
    });

    const rawText = await response.text();
    let result = {};
    if (rawText) {
      try {
        result = JSON.parse(rawText);
      } catch {
        throw new Error("서버 응답을 읽는 중 문제가 발생했습니다. 잠시 후 다시 시도해 주세요.");
      }
    }

    if (!response.ok) {
      throw new Error(result.error || "보고서 생성 중 서버 오류가 발생했습니다. 잠시 후 다시 시도해 주세요.");
    }

    const { summary } = result;
    downloadLink.href = result.downloadUrl;
    downloadLink.download = result.fileName;
    summaryBrand.textContent = summary.brand;
    summaryRange.textContent = summary.reportRange;
    summaryMemoCount.textContent = `${summary.memoCount}건`;

    renderKpis(summary);
    renderInsights(summary);
    renderCharts(summary);
    renderChannels(summary);
    renderKeywords(summary);

    resultPanel.classList.remove("hidden");
    resultPanel.scrollIntoView({ behavior: "smooth", block: "start" });
  } catch (error) {
    alert(error instanceof Error ? error.message : "알 수 없는 오류가 발생했습니다.");
  } finally {
    submitButton.disabled = false;
    submitButton.textContent = "주간보고서 생성";
  }
});
