const form = document.querySelector("#report-form");
const resultPanel = document.querySelector("#result-panel");
const downloadLink = document.querySelector("#download-link");
const submitButton = document.querySelector("#submit-button");
const kpiGrid = document.querySelector("#kpi-grid");
const channelBars = document.querySelector("#channel-bars");
const keywordList = document.querySelector("#keyword-list");
const summaryBrand = document.querySelector("#summary-brand");
const summaryRange = document.querySelector("#summary-range");
const summaryMemoCount = document.querySelector("#summary-memo-count");

function formatNumber(value) {
  return Number(value || 0).toLocaleString("ko-KR");
}

function formatCurrency(value) {
  return `${formatNumber(Math.round(value || 0))}원`;
}

function formatPercent(value, digits = 2) {
  return `${Number(value || 0).toFixed(digits)}%`;
}

function buildKpis(summary) {
  return [
    { label: "총 광고비", value: summary.spend, tone: "warm" },
    { label: "구매완료 수", value: `${formatNumber(summary.conversions)}건`, tone: "cool" },
    { label: "구매완료 매출", value: summary.revenue, tone: "neutral" },
    { label: "구매완료 ROAS", value: summary.roas, tone: "accent" },
    { label: "CTR", value: formatPercent(summary.raw.ctr), tone: "neutral" },
    { label: "구매완료 CVR", value: formatPercent(summary.raw.cvr), tone: "cool" },
    { label: "평균 CPC", value: formatCurrency(summary.raw.cpc), tone: "neutral" },
    { label: "구매완료 CPA", value: formatCurrency(summary.raw.cpa), tone: "warm" },
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
    `;
    kpiGrid.append(card);
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
          <span>구매완료 매출 ${formatCurrency(channel.revenue)}</span>
        </div>
      </div>
      <div class="bar-stack">
        <div class="bar-track">
          <div class="bar bar-cost" style="width:${Math.max((channel.cost / maxCost) * 100, 4)}%"></div>
        </div>
        <div class="bar-track">
          <div class="bar bar-revenue" style="width:${Math.max((channel.revenue / maxRevenue) * 100, 4)}%"></div>
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
        <div><dt>광고비</dt><dd>${formatCurrency(keyword.cost)}</dd></div>
        <div><dt>구매완료 수</dt><dd>${formatNumber(keyword.conversions)}건</dd></div>
        <div><dt>구매완료 매출</dt><dd>${formatCurrency(keyword.revenue)}</dd></div>
        <div><dt>구매완료 ROAS</dt><dd>${formatPercent(keyword.roas, 0)}</dd></div>
      </dl>
    `;
    keywordList.append(item);
  }
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
