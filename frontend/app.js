const form = document.querySelector("#proposalForm");
const saveStatus = document.querySelector("#saveStatus");
const searchStatus = document.querySelector("#searchStatus");
const resultsBox = document.querySelector("#companyResults");
const findInnBtn = document.querySelector("#findInnBtn");
const companyNameInput = document.querySelector("#companyName");
const innInput = document.querySelector("#inn");
const websiteInput = form.elements.website;
const existingCompanySelect = document.querySelector("#existingCompanySelect");
const selectExistingCompany = document.querySelector("#selectExistingCompany");
const buildPresentationBtn = document.querySelector("#buildPresentationBtn");
const prepareAiSelectionBtn = document.querySelector("#prepareAiSelectionBtn");
const finalizeAiSelectionBtn = document.querySelector("#finalizeAiSelectionBtn");
const approveAiSelectionBtn = document.querySelector("#approveAiSelectionBtn");
const aiSelectionStatus = document.querySelector("#aiSelectionStatus");
const aiSelectionPanel = document.querySelector("#aiSelectionPanel");
const aiSelectionSource = document.querySelector("#aiSelectionSource");
const aiSelectionRationale = document.querySelector("#aiSelectionRationale");
const aiStatsChoice = document.querySelector("#aiStatsChoice");
const aiPhotosChoice = document.querySelector("#aiPhotosChoice");
const aiReviewsChoice = document.querySelector("#aiReviewsChoice");
const aiFinalSection = document.querySelector("#aiFinalSection");
const aiFinalPanel = document.querySelector("#aiFinalPanel");
const aiFinalSource = document.querySelector("#aiFinalSource");
const aiFinalRationale = document.querySelector("#aiFinalRationale");
const aiFinalStats = document.querySelector("#aiFinalStats");
const aiFinalPhotos = document.querySelector("#aiFinalPhotos");
const aiFinalReviews = document.querySelector("#aiFinalReviews");
const approveAiFinalBtn = document.querySelector("#approveAiFinalBtn");
const aiRequestFlow = document.querySelector("#aiRequestFlow");
const aiThinkingTimer = document.querySelector("#aiThinkingTimer");
const aiDebugPanel = document.querySelector("#aiDebugPanel");
const aiDebugStatus = document.querySelector("#aiDebugStatus");
const aiDebugJson = document.querySelector("#aiDebugJson");
const aiRawResponsePanel = document.querySelector("#aiRawResponsePanel");
const aiRawResponseStatus = document.querySelector("#aiRawResponseStatus");
const aiRawResponseJson = document.querySelector("#aiRawResponseJson");
const presentationProgress = document.querySelector("#presentationProgress");
const presentationProgressText = document.querySelector("#presentationProgressText");
const presentationProgressPercent = document.querySelector("#presentationProgressPercent");
const presentationProgressBar = document.querySelector("#presentationProgressBar");
const dadataDetails = document.querySelector("#dadataDetails");
const dadataGrid = document.querySelector("#dadataGrid");
const dadataSummary = document.querySelector("#dadataSummary");
const dadataUpdatedAt = document.querySelector("#dadataUpdatedAt");
const toggleDadataBtn = document.querySelector("#toggleDadataBtn");
const refreshDadataBtn = document.querySelector("#refreshDadataBtn");
const buildPresentationDefaultText = buildPresentationBtn.textContent;
let existingCards = [];
let activePresentationJobId = null;
let activeSearchRequestId = 0;
let activeAiSelection = null;
let approvedAiSelectionId = "";
let aiRequestStartedAt = 0;
let aiThinkingInterval = null;
let dadataExpanded = false;

const optionIds = {
  industries: "industries",
  activities: "activities",
  productCategories: "productCategories",
};
const ADD_COMPANY_TYPE = "__add_company_type__";
const dadataFields = [
  ["dadataName", "Название"],
  ["dadataStatus", "Статус"],
  ["dadataLegalAddress", "Юр адрес"],
  ["dadataMainActivity", "Основной вид деятельности"],
  ["dadataEmployeeCount", "Среднесписочная численность"],
  ["dadataCapital", "Уставной капитал"],
  ["dadataSmb", "Реестр МСП"],
  ["dadataTaxSystem", "Спец. налоговый режим"],
  ["dadataIncome", "Доходы"],
  ["dadataExpense", "Расходы"],
  ["dadataDebt", "Недоимки"],
  ["dadataPenalty", "Штрафы"],
  ["dadataManager", "Ген. директор"],
];
const dadataHeaderByField = {
  dadataName: "DaData: название",
  dadataStatus: "DaData: статус",
  dadataLegalAddress: "DaData: юр адрес",
  dadataMainActivity: "DaData: основной вид деятельности",
  dadataEmployeeCount: "DaData: среднесписочная численность",
  dadataCapital: "DaData: уставной капитал",
  dadataSmb: "DaData: реестр МСП",
  dadataTaxSystem: "DaData: специальный налоговый режим",
  dadataIncome: "DaData: доходы",
  dadataExpense: "DaData: расходы",
  dadataDebt: "DaData: недоимки",
  dadataPenalty: "DaData: штрафы",
  dadataManager: "DaData: ген. директор",
  dadataUpdatedAt: "DaData: обновлено",
};

function fillDatalist(id, values) {
  const list = document.querySelector(`#${id}`);
  list.innerHTML = "";
  values.forEach((value) => {
    const option = document.createElement("option");
    option.value = value;
    list.append(option);
  });
}

function fillSelect(id, values) {
  const select = document.querySelector(`#${id}`);
  select.innerHTML = "";
  values.forEach((value) => {
    const option = document.createElement("option");
    option.value = value;
    option.textContent = value;
    select.append(option);
  });
}

function fillCompanyTypes(values) {
  const select = document.querySelector("#companyTypeSelect");
  const hidden = document.querySelector("#companyType");
  const custom = document.querySelector("#companyTypeCustom");
  const previous = hidden.value || select.value;

  select.innerHTML = "";
  values.forEach((value) => {
    const option = document.createElement("option");
    option.value = value;
    option.textContent = value;
    select.append(option);
  });

  const addOption = document.createElement("option");
  addOption.value = ADD_COMPANY_TYPE;
  addOption.textContent = "Добавить новый тип...";
  select.append(addOption);

  if (previous && values.includes(previous)) {
    select.value = previous;
  }
  syncCompanyType();
}

function syncCompanyType() {
  const select = document.querySelector("#companyTypeSelect");
  const hidden = document.querySelector("#companyType");
  const custom = document.querySelector("#companyTypeCustom");
  const isCustom = select.value === ADD_COMPANY_TYPE;

  custom.classList.toggle("is-visible", isCustom);
  custom.required = isCustom;
  hidden.value = isCustom ? custom.value.trim() : select.value;
}

function fillPreferredNetworkCategories(values) {
  const host = document.querySelector("#preferredNetworkCategories");
  host.innerHTML = "";
  values.forEach((value) => {
    const label = document.createElement("label");
    label.className = "chip";
    label.innerHTML = `<input type="checkbox" name="preferredNetworkCategories" value="${value}"><span>${value}</span>`;
    host.append(label);
  });
}

function fillPreferredNetworks(values) {
  const host = document.querySelector("#preferredNetworks");
  host.innerHTML = "";
  values.forEach((value) => {
    const label = document.createElement("label");
    label.className = "chip";
    label.innerHTML = `<input type="checkbox" name="preferredNetworks" value="${value}"><span>${value}</span>`;
    host.append(label);
  });
}

function setStatus(element, text, isError = false) {
  element.textContent = text;
  element.classList.toggle("error", isError);
}

function setSearchLoading(isLoading) {
  searchStatus.classList.toggle("is-loading", isLoading);
}

function selectedAiProvider() {
  return form.querySelector("input[name='aiProvider']:checked")?.value || "rules";
}

function setAiRequestStage(stage) {
  if (!aiRequestFlow) {
    return;
  }
  const stages = ["send", "think", "receive"];
  aiRequestFlow.hidden = false;
  aiRequestFlow.querySelectorAll(".ai-request-step").forEach((item) => {
    const itemStage = item.dataset.aiStage;
    item.classList.toggle("is-active", itemStage === stage);
    item.classList.toggle("is-done", stages.indexOf(itemStage) < stages.indexOf(stage));
  });
}

function stopAiRequestFlow() {
  if (aiThinkingInterval) {
    window.clearInterval(aiThinkingInterval);
    aiThinkingInterval = null;
  }
  aiRequestStartedAt = 0;
  if (!aiRequestFlow) {
    return;
  }
  aiRequestFlow.hidden = true;
  aiRequestFlow.querySelectorAll(".ai-request-step").forEach((item) => {
    item.classList.remove("is-active", "is-done");
  });
}

function startAiRequestFlow() {
  stopAiRequestFlow();
  aiRequestStartedAt = Date.now();
  aiFinalSection.hidden = false;
  aiFinalPanel.hidden = false;
  aiFinalSource.textContent = `Источник: ${selectedAiProvider()}`;
  aiFinalRationale.textContent = "ИИ получает JSON с параметрами компании, кандидатами фотографий и отзывов.";
  aiFinalStats.innerHTML = "";
  aiFinalPhotos.innerHTML = "";
  aiFinalReviews.innerHTML = "";
  setAiRequestStage("send");
  if (aiThinkingTimer) {
    aiThinkingTimer.textContent = "0 сек.";
  }
  window.setTimeout(() => {
    if (aiRequestStartedAt) {
      setAiRequestStage("think");
    }
  }, 450);
  aiThinkingInterval = window.setInterval(() => {
    if (!aiThinkingTimer || !aiRequestStartedAt) {
      return;
    }
    const seconds = Math.max(0, Math.round((Date.now() - aiRequestStartedAt) / 1000));
    aiThinkingTimer.textContent = `${seconds} сек.`;
  }, 500);
}

function finishAiRequestFlow() {
  if (aiThinkingInterval) {
    window.clearInterval(aiThinkingInterval);
    aiThinkingInterval = null;
  }
  setAiRequestStage("receive");
}

function clearAiDebugPanel() {
  if (!aiDebugPanel) {
    return;
  }
  aiDebugPanel.hidden = true;
  aiDebugStatus.textContent = "";
  aiDebugJson.textContent = "";
}

function clearAiRawResponsePanel() {
  if (!aiRawResponsePanel) {
    return;
  }
  aiRawResponsePanel.hidden = true;
  aiRawResponseStatus.textContent = "";
  aiRawResponseJson.textContent = "";
}

function showAiRawResponse(rawResponse, statusText = "") {
  if (!aiRawResponsePanel) {
    return;
  }
  aiRawResponsePanel.hidden = false;
  aiRawResponseStatus.textContent = statusText || rawResponse?.provider || "Ответ получен";
  aiRawResponseJson.textContent = typeof rawResponse === "string"
    ? rawResponse
    : JSON.stringify(rawResponse || {}, null, 2);
}

async function refreshAiDebugPayload() {
  if (!activeAiSelection?.id || !aiDebugPanel) {
    clearAiDebugPanel();
    return;
  }
  aiDebugPanel.hidden = false;
  aiDebugStatus.textContent = "Готовлю preview...";
  try {
    const response = await fetch(`/api/ai-selection/${encodeURIComponent(activeAiSelection.id)}/request-preview`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ provider: selectedAiProvider(), ...collectManualChoices() }),
    });
    const payload = await response.json();
    if (!response.ok) {
      throw new Error(payload.detail || "Не удалось подготовить JSON запроса");
    }
    aiDebugStatus.textContent = payload.url ? `${payload.provider} · ${payload.url}` : payload.message || payload.provider;
    aiDebugJson.textContent = JSON.stringify(payload.request || payload, null, 2);
  } catch (error) {
    aiDebugStatus.textContent = "Ошибка preview";
    aiDebugJson.textContent = error.message;
  }
}

function clearAiSelection() {
  activeAiSelection = null;
  approvedAiSelectionId = "";
  stopAiRequestFlow();
  clearAiDebugPanel();
  clearAiRawResponsePanel();
  approveAiSelectionBtn.disabled = true;
  finalizeAiSelectionBtn.disabled = true;
  approveAiFinalBtn.disabled = true;
  aiSelectionPanel.hidden = true;
  aiFinalSection.hidden = true;
  aiFinalPanel.hidden = true;
  aiSelectionSource.textContent = "";
  aiSelectionRationale.textContent = "";
  aiFinalSource.textContent = "";
  aiFinalRationale.textContent = "";
  aiStatsChoice.innerHTML = "";
  aiPhotosChoice.innerHTML = "";
  aiReviewsChoice.innerHTML = "";
  aiFinalStats.innerHTML = "";
  aiFinalPhotos.innerHTML = "";
  aiFinalReviews.innerHTML = "";
  aiSelectionStatus.textContent = "";
  aiSelectionStatus.classList.remove("error");
}

function appendChoice(host, title, text) {
  const item = document.createElement("div");
  item.className = "ai-choice-item";
  const strong = document.createElement("strong");
  strong.textContent = title;
  const body = document.createElement("span");
  body.textContent = text || "Нет данных";
  item.append(strong, body);
  host.append(item);
}

function canFinalizeWithAi() {
  return ["openai", "deepseek"].includes(selectedAiProvider());
}

function previewUrlForPath(path) {
  return `/api/image-preview?path=${encodeURIComponent(path)}`;
}

function collectManualChoices() {
  return {
    photoIds: [...document.querySelectorAll("input[data-ai-photo-id]:checked")].map((input) => input.value),
    reviewIds: [...document.querySelectorAll("input[data-ai-review-id]:checked")].map((input) => input.value),
  };
}

function selectedCounts() {
  return {
    photoIds: [...document.querySelectorAll("input[data-ai-photo-id]:checked")].map((input) => input.value),
    reviewIds: [...document.querySelectorAll("input[data-ai-review-id]:checked")].map((input) => input.value),
  };
}

function syncSelectionApprovalState() {
  if (!activeAiSelection?.selection) {
    approveAiSelectionBtn.disabled = true;
    return;
  }
  if (activeAiSelection?.approved) {
    approveAiSelectionBtn.disabled = true;
    return;
  }
  const counts = selectedCounts();
  const requiredPhotos = Number(activeAiSelection.selection.required_photo_count || 0);
  const requiredReviews = Number(activeAiSelection.selection.required_review_count || 0);
  const photosReady = counts.photoIds.length === requiredPhotos;
  const reviewsReady = counts.reviewIds.length === requiredReviews;
  approveAiSelectionBtn.disabled = !(photosReady && reviewsReady);
}

function renderCandidateOptions(host, candidates, kind, requiredCount, selectedIds) {
  host.innerHTML = "";
  if (!candidates.length) {
    appendChoice(host, kind === "photo" ? "Фотографии" : "Отзывы", "Варианты не найдены");
    return;
  }

  const header = document.createElement("div");
  header.className = "ai-option-group";
  const title = document.createElement("strong");
  title.className = "ai-option-group-title";
  title.textContent = kind === "photo"
    ? `Выберите ${requiredCount} из ${candidates.length} фотографий`
    : `Выберите ${requiredCount} из ${candidates.length} отзывов`;
  header.append(title);
  const caption = document.createElement("span");
  caption.className = "ai-option-group-caption";
  caption.textContent = kind === "photo"
    ? `Сейчас отмечено ${selectedIds.length} из ${requiredCount}`
    : `Сейчас отмечено ${selectedIds.length} из ${requiredCount}`;
  header.append(caption);
  host.append(header);

  candidates.forEach((option, index) => {
    const label = document.createElement("label");
    label.className = `ai-option-card${kind === "photo" ? " is-photo" : ""}`;
    const input = document.createElement("input");
    input.type = "checkbox";
    input.value = option.id;
    input.checked = selectedIds.includes(option.id);
    input.dataset.aiKind = kind;
    input.dataset.aiLimit = String(requiredCount);
    if (kind === "photo") {
      input.dataset.aiPhotoId = option.id;
    } else {
      input.dataset.aiReviewId = option.id;
    }

    if (kind === "photo" && option.path) {
      const preview = document.createElement("img");
      preview.className = "ai-option-preview";
      preview.src = previewUrlForPath(option.path);
      preview.alt = option.network || `Фото ${index + 1}`;
      label.append(preview);
    }

    const body = document.createElement("div");
    body.className = "ai-option-card-body";
    const heading = document.createElement("strong");
    heading.textContent = kind === "photo"
      ? (option.network || option.path || `Фото ${index + 1}`)
      : (option.company || `Отзыв ${index + 1}`);
    const meta = document.createElement("span");
    meta.textContent = kind === "photo"
      ? [option.network_type, option.region, option.price_segment, option.path].filter(Boolean).join(" · ")
      : [option.person, option.group, option.text].filter(Boolean).join(" · ");
    const reason = document.createElement("small");
    reason.textContent = [option.reason, option.score ? `score ${option.score}` : ""].filter(Boolean).join(" · ");
    body.append(heading, meta, reason);
    label.append(input, body);
    host.append(label);
  });
}

function renderStats(selection, host, rationaleText = "") {
  host.innerHTML = "";
  const stats = selection.stats || {};
  appendChoice(host, "Категория", stats.category || selection.category || "");
  appendChoice(host, "Переговоров", stats.negotiations || "");
  appendChoice(host, "Интерес / чемпионы", `${stats.interest || ""}${stats.champions ? ` / ${stats.champions}` : ""}`);
  if (rationaleText) {
    appendChoice(host, "Почему", rationaleText);
  }
}

function renderAiSelection(payload) {
  const selection = payload.selection || {};
  const rationale = selection.rationale || {};
  activeAiSelection = payload;
  approvedAiSelectionId = payload.approved ? payload.id : "";
  approveAiSelectionBtn.disabled = Boolean(payload.approved);
  finalizeAiSelectionBtn.disabled = !canFinalizeWithAi();
  approveAiFinalBtn.disabled = !payload.finalSelection;

  aiSelectionSource.textContent = `Источник: ${selection.source || payload.provider || "rules"}`;
  aiSelectionRationale.textContent = rationale.summary || "Алгоритм подготовил предварительный пул материалов.";

  renderStats(selection, aiStatsChoice, rationale.stats || "");
  renderCandidateOptions(
    aiPhotosChoice,
    selection.photo_candidates || [],
    "photo",
    Number(selection.required_photo_count || 0),
    selection.selected_photo_ids || [],
  );
  renderCandidateOptions(
    aiReviewsChoice,
    selection.review_candidates || [],
    "review",
    Number(selection.required_review_count || 0),
    selection.selected_review_ids || [],
  );
  aiSelectionPanel.hidden = false;
  syncSelectionApprovalState();
  refreshAiDebugPayload();

  if (payload.finalSelection) {
    renderAiFinalSelection(payload.finalSelection);
  }
}

function renderAiFinalSelection(selection) {
  const rationale = selection?.rationale || {};
  aiFinalSection.hidden = false;
  aiFinalPanel.hidden = false;
  aiFinalSource.textContent = `Источник: ${selection?.source || "ai"}`;
  aiFinalRationale.textContent = rationale.summary || "ИИ подготовил окончательный выбор.";
  renderStats(selection || {}, aiFinalStats, rationale.stats || "");

  aiFinalPhotos.innerHTML = "";
  const photoMap = new Map((selection?.photo_candidates || []).map((item) => [item.id, item]));
  (selection?.selected_photo_ids || []).forEach((photoId, index) => {
    const item = photoMap.get(photoId);
    aiFinalPhotos.append(renderAiFinalPhotoCard(item, photoId, index));
  });
  if (!aiFinalPhotos.children.length) {
    appendChoice(aiFinalPhotos, "Фотографии", "Фотографии не выбраны");
  }
  if (rationale.photos) {
    appendChoice(aiFinalPhotos, "Почему", rationale.photos);
  }

  aiFinalReviews.innerHTML = "";
  const reviewMap = new Map((selection?.review_candidates || []).map((item) => [item.id, item]));
  (selection?.selected_review_ids || []).forEach((reviewId, index) => {
    const review = reviewMap.get(reviewId);
    aiFinalReviews.append(renderAiFinalReviewCard(review, reviewId, index));
  });
  if (!aiFinalReviews.children.length) {
    appendChoice(aiFinalReviews, "Отзывы", "Отзывы не выбраны");
  }
  if (rationale.reviews) {
    appendChoice(aiFinalReviews, "Почему", rationale.reviews);
  }
}

function renderAiFinalPhotoCard(item, photoId, index) {
  const card = document.createElement("article");
  card.className = "ai-final-card is-photo";
  if (item?.path) {
    const preview = document.createElement("img");
    preview.className = "ai-final-preview";
    preview.src = previewUrlForPath(item.path);
    preview.alt = item.network || `Фото ${index + 1}`;
    card.append(preview);
  }

  const body = document.createElement("div");
  body.className = "ai-final-card-body";
  const title = document.createElement("strong");
  title.textContent = item?.network || `Фото ${index + 1}`;
  const meta = document.createElement("span");
  meta.textContent = item
    ? [item.network_type, item.region, item.price_segment].filter(Boolean).join(" · ")
    : photoId;
  const detail = document.createElement("small");
  detail.textContent = item?.path || photoId;
  body.append(title, meta, detail);
  card.append(body);
  return card;
}

function renderAiFinalReviewCard(review, reviewId, index) {
  const card = document.createElement("article");
  card.className = "ai-final-card";
  const body = document.createElement("div");
  body.className = "ai-final-card-body";
  const title = document.createElement("strong");
  title.textContent = review?.company || `Отзыв ${index + 1}`;
  const meta = document.createElement("span");
  meta.textContent = [review?.person, review?.group, review?.networks].filter(Boolean).join(" · ");
  const detail = document.createElement("small");
  detail.textContent = review?.text || reviewId;
  body.append(title, meta, detail);
  card.append(body);
  return card;
}

async function prepareAiSelection() {
  const company = companyNameInput.value.trim();
  if (!company) {
    setStatus(aiSelectionStatus, "Сначала выберите или заполните компанию.", true);
    return;
  }

  clearAiSelection();
  prepareAiSelectionBtn.disabled = true;
  setStatus(aiSelectionStatus, "Готовлю список вариантов...");
  try {
    const response = await fetch("/api/ai-selection", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ companyName: company, provider: selectedAiProvider(), draftPayload: formPayload() }),
    });
    const payload = await response.json();
    if (!response.ok) {
      throw new Error(payload.detail || "Не удалось подготовить выбор материалов");
    }
    renderAiSelection(payload);
    setStatus(aiSelectionStatus, canFinalizeWithAi()
      ? "Выберите материалы, затем одобрите выбор или отдайте его ИИ."
      : "Выберите материалы и одобрите выбор.");
  } catch (error) {
    setStatus(aiSelectionStatus, error.message, true);
  } finally {
    prepareAiSelectionBtn.disabled = false;
  }
}

async function approveAiSelection(mode = "manual") {
  if (!activeAiSelection?.id) {
    setStatus(aiSelectionStatus, "Сначала подготовьте выбор материалов.", true);
    return;
  }

  approveAiSelectionBtn.disabled = true;
  approveAiFinalBtn.disabled = true;
  setStatus(aiSelectionStatus, mode === "ai" ? "Одобряю ответ ИИ..." : "Одобряю выбор...");
  try {
    const response = await fetch(`/api/ai-selection/${encodeURIComponent(activeAiSelection.id)}/approve`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ mode, ...collectManualChoices() }),
    });
    const payload = await response.json();
    if (!response.ok) {
      throw new Error(payload.detail || "Не удалось одобрить выбор");
    }
    approvedAiSelectionId = payload.id;
    activeAiSelection.approved = true;
    if (payload.selection) {
      activeAiSelection.selection = payload.selection;
    }
    setStatus(aiSelectionStatus, "Выбор одобрен. Запускаю формирование презентации...");
    await buildPresentation();
  } catch (error) {
    setStatus(aiSelectionStatus, error.message, true);
  } finally {
    approveAiSelectionBtn.disabled = Boolean(activeAiSelection?.approved);
    approveAiFinalBtn.disabled = !activeAiSelection?.finalSelection;
  }
}

async function finalizeAiSelection() {
  if (!activeAiSelection?.id) {
    setStatus(aiSelectionStatus, "Сначала подготовьте выбор материалов.", true);
    return;
  }
  if (!canFinalizeWithAi()) {
    setStatus(aiSelectionStatus, "Для ответа ИИ выберите OpenAI или DeepSeek.", true);
    return;
  }

  finalizeAiSelectionBtn.disabled = true;
  clearAiRawResponsePanel();
  startAiRequestFlow();
  setStatus(aiSelectionStatus, "Отправляю JSON с компанией, фотографиями и отзывами в ИИ...");
  try {
    const response = await fetch(`/api/ai-selection/${encodeURIComponent(activeAiSelection.id)}/finalize`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ provider: selectedAiProvider(), ...collectManualChoices() }),
    });
    const payload = await response.json();
    if (!response.ok) {
      const detail = payload.detail || "Не удалось получить ответ ИИ";
      if (detail.rawResponse) {
        showAiRawResponse(detail.rawResponse, "Ошибка разбора ответа");
      } else if (typeof detail !== "string") {
        showAiRawResponse(detail, "Ошибка ИИ");
      }
      throw new Error(typeof detail === "string" ? detail : detail.message || "Не удалось получить ответ ИИ");
    }
    finishAiRequestFlow();
    activeAiSelection = payload;
    renderAiFinalSelection(payload.finalSelection || {});
    if (payload.finalSelection?.aiRawResponse) {
      showAiRawResponse(payload.finalSelection.aiRawResponse, "Ответ разобран");
    }
    approveAiFinalBtn.disabled = false;
    setStatus(aiSelectionStatus, "ИИ вернул окончательный вариант. Проверьте его ниже.");
  } catch (error) {
    stopAiRequestFlow();
    if (!activeAiSelection?.finalSelection) {
      aiFinalSection.hidden = true;
      aiFinalPanel.hidden = true;
    }
    setStatus(aiSelectionStatus, error.message, true);
  } finally {
    finalizeAiSelectionBtn.disabled = !canFinalizeWithAi();
  }
}

function clearDadataDetails() {
  dadataDetails.hidden = true;
  dadataExpanded = false;
  dadataDetails.classList.remove("is-expanded");
  dadataGrid.innerHTML = "";
  dadataSummary.textContent = "";
  dadataUpdatedAt.textContent = "";
  toggleDadataBtn.textContent = "Развернуть";
  toggleDadataBtn.setAttribute("aria-expanded", "false");
}

function setDadataExpanded(isExpanded) {
  dadataExpanded = isExpanded;
  dadataDetails.classList.toggle("is-expanded", isExpanded);
  toggleDadataBtn.textContent = isExpanded ? "Свернуть" : "Развернуть";
  toggleDadataBtn.setAttribute("aria-expanded", String(isExpanded));
}

function renderDadataDetails(details = {}) {
  const visibleItems = dadataFields
    .map(([field, label]) => ({ field, label, value: details[field] || "нет данных в ответе DaData" }));
  dadataGrid.innerHTML = "";
  if (!Object.values(details).some((value) => value)) {
    clearDadataDetails();
    return;
  }

  visibleItems.forEach((item) => {
    const wrapper = document.createElement("dl");
    wrapper.className = "dadata-item";
    const term = document.createElement("dt");
    term.textContent = item.label;
    const value = document.createElement("dd");
    value.textContent = item.value;
    wrapper.append(term, value);
    dadataGrid.append(wrapper);
  });
  dadataSummary.textContent = details.dadataName || "Название не получено";
  dadataUpdatedAt.textContent = details.dadataUpdatedAt ? `Обновлено: ${details.dadataUpdatedAt}` : "";
  setDadataExpanded(false);
  dadataDetails.hidden = false;
}

function dadataDetailsFromCard(card) {
  const details = {};
  Object.entries(dadataHeaderByField).forEach(([field, header]) => {
    details[field] = valueBy(card, header);
  });
  return details;
}

async function loadDadataByInn(inn, { rowNumber = "", silent = false } = {}) {
  const normalizedInn = String(inn || "").replace(/\D+/g, "");
  if (!normalizedInn) {
    clearDadataDetails();
    return null;
  }
  refreshDadataBtn.disabled = true;
  if (!silent) {
    setStatus(searchStatus, "Получаю сведения DaData...");
  }
  try {
    const response = await fetch("/api/dadata-company", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ inn: normalizedInn, rowNumber }),
    });
    const payload = await response.json();
    if (!response.ok) {
      throw new Error(payload.detail || "Не удалось получить сведения DaData");
    }
    renderDadataDetails(payload.items || {});
    const card = existingCards.find((item) => String(item._rowNumber) === String(rowNumber));
    if (card && payload.items) {
      Object.entries(dadataHeaderByField).forEach(([field, header]) => {
        card[header] = payload.items[field] || "";
      });
    }
    if (!silent) {
      setStatus(searchStatus, "Сведения DaData получены.");
    }
    return payload.items || {};
  } catch (error) {
    if (!silent) {
      setStatus(searchStatus, error.message, true);
    }
    return null;
  } finally {
    refreshDadataBtn.disabled = false;
  }
}

function valueBy(card, header) {
  return card[header] || card[header.trim()] || "";
}

function innBy(card) {
  return valueBy(card, "ИНН");
}

function contactNameBy(card) {
  return valueBy(card, "ФИО контакта");
}

function currentRowNumber() {
  return document.querySelector("#rowNumber").value.trim();
}

function renderExistingCompanyOptions() {
  existingCompanySelect.innerHTML = "";
  existingCards.forEach((card) => {
    const option = document.createElement("option");
    option.value = card._rowNumber;
    option.textContent = `${valueBy(card, "Название компании") || "Компания без названия"}${innBy(card) ? ` · ${innBy(card)}` : ""}`;
    existingCompanySelect.append(option);
  });
}

async function rememberInnForExistingCard(inn) {
  const normalizedInn = String(inn || "").replace(/\D+/g, "");
  const rowNumber = currentRowNumber();
  if (!rowNumber || !normalizedInn) {
    return;
  }

  const card = existingCards.find((item) => String(item._rowNumber) === String(rowNumber));
  if (card && innBy(card) === normalizedInn) {
    return;
  }

  const response = await fetch(`/api/cards/${encodeURIComponent(rowNumber)}/inn`, {
    method: "PATCH",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ inn: normalizedInn }),
  });
  const payload = await response.json();
  if (!response.ok) {
    throw new Error(payload.detail || "Не удалось сохранить ИНН в карточке");
  }

  if (card) {
    card["ИНН"] = normalizedInn;
  }
  renderExistingCompanyOptions();
  existingCompanySelect.value = rowNumber;
  setStatus(saveStatus, `ИНН ${normalizedInn} сохранен в строке ${rowNumber}.`);
}

function setFormValue(name, value) {
  const field = form.elements[name];
  if (field) {
    field.value = value || "";
  }
}

function setCompanyType(value) {
  const select = document.querySelector("#companyTypeSelect");
  const custom = document.querySelector("#companyTypeCustom");
  const options = [...select.options].map((option) => option.value);
  if (options.includes(value)) {
    select.value = value;
    custom.value = "";
  } else {
    select.value = ADD_COMPANY_TYPE;
    custom.value = value || "";
  }
  syncCompanyType();
}

function setPreferredNetworkCategories(value) {
  const selected = new Set((value || "").split(",").map((item) => item.trim()).filter(Boolean));
  form.querySelectorAll("input[name='preferredNetworkCategories']").forEach((checkbox) => {
    checkbox.checked = selected.has(checkbox.value);
  });
}

function setPreferredNetworks(value) {
  const selected = new Set((value || "").split(",").map((item) => item.trim()).filter(Boolean));
  form.querySelectorAll("input[name='preferredNetworks']").forEach((checkbox) => {
    checkbox.checked = selected.has(checkbox.value);
  });
}

function applyExistingCard(rowNumber) {
  const card = existingCards.find((item) => String(item._rowNumber) === String(rowNumber));
  if (!card) {
    return;
  }

  resetPresentationProgress();
  clearAiSelection();
  setFormValue("_rowNumber", card._rowNumber);
  setFormValue("companyName", valueBy(card, "Название компании"));
  setFormValue("inn", innBy(card));
  setCompanyType(valueBy(card, "Тип компании"));
  setFormValue("contactName", contactNameBy(card));
  setFormValue("contactPosition", valueBy(card, "Должность контакта"));
  setFormValue("industry", valueBy(card, "Отрасль"));
  setFormValue("activity", valueBy(card, "Сфера деятельности"));
  setFormValue("networkCategories", valueBy(card, "Категория по работе с сетями"));
  setFormValue("website", valueBy(card, "Сайт"));
  setFormValue("country", valueBy(card, "Страна"));
  setFormValue("city", valueBy(card, "Город"));
  setFormValue("productName", valueBy(card, "Название товара"));
  setFormValue("productCategory", valueBy(card, "Категория товара"));
  setFormValue("priceCategory", valueBy(card, "Ценовая категория"));
  setFormValue("productDescription", valueBy(card, "Краткое описание продукции"));
  setPreferredNetworkCategories(valueBy(card, "Предпочтительные категории сетей") || valueBy(card, "Предпочтительные сети"));
  setPreferredNetworks(valueBy(card, "Предпочтительные сети"));
  setStatus(saveStatus, `Редактируется строка ${card._rowNumber}: ${valueBy(card, "Название компании")}`);
  renderDadataDetails(dadataDetailsFromCard(card));
  if (innInput.value.trim()) {
    loadDadataByInn(innInput.value, { rowNumber: card._rowNumber, silent: true });
  } else {
    searchCompany({ auto: true });
  }
}

function clearSelectedCompany() {
  form.reset();
  document.querySelector("#rowNumber").value = "";
  document.querySelector("input[name='country']").value = "Россия";
  setCompanyType("");
  setPreferredNetworkCategories("");
  setPreferredNetworks("");
  clearDadataDetails();
  clearAiSelection();
  saveStatus.textContent = "";
}

async function loadExistingCards() {
  const response = await fetch("/api/cards");
  if (!response.ok) {
    throw new Error("Не удалось загрузить существующие компании");
  }
  const payload = await response.json();
  existingCards = payload.items || [];
  renderExistingCompanyOptions();
}

async function loadOptions() {
  const response = await fetch("/api/options");
  if (!response.ok) {
    throw new Error("Не удалось загрузить справочники");
  }
  const options = await response.json();

  fillCompanyTypes(options.companyTypes || []);
  Object.entries(optionIds).forEach(([key, id]) => fillDatalist(id, options[key] || []));
  fillSelect("priceCategory", options.priceCategories || []);
  fillSelect("networkWorkCategory", options.networkWorkCategories || []);
  fillPreferredNetworkCategories(options.preferredNetworkCategories || []);
  fillPreferredNetworks(options.preferredNetworks || []);
  document.querySelector("#savedCount").textContent = options.cardCount || 0;
  await loadExistingCards();
}

function renderCompanyResults(items, siteInns = []) {
  resultsBox.innerHTML = "";
  const siteInnSet = new Set(siteInns.map((item) => item.inn));
  items.forEach((item) => {
    const matchedBySite = item.inn && siteInnSet.has(item.inn);
    const button = document.createElement("button");
    button.type = "button";
    button.className = "result-item";
    button.innerHTML = `
      <strong>${item.name || "Компания без названия"}</strong>
      <small>ИНН ${item.inn || "не указан"}${matchedBySite ? " · подтверждено сайтом" : ""}${item.address ? ` · ${item.address}` : ""}${item.status ? ` · ${item.status}` : ""}</small>
    `;
    button.addEventListener("click", async () => {
      companyNameInput.value = item.name || companyNameInput.value;
      innInput.value = item.inn || innInput.value;
      resultsBox.innerHTML = "";
      setStatus(searchStatus, matchedBySite ? "Компания выбрана из базы ФНС, ИНН подтвержден на сайте." : "Компания выбрана из базы ФНС.");
      try {
        await rememberInnForExistingCard(item.inn);
        await loadDadataByInn(item.inn, { rowNumber: currentRowNumber() });
      } catch (error) {
        setStatus(searchStatus, error.message, true);
      }
    });
    resultsBox.append(button);
  });
}

function normalizeCompanySearchPayload(payload) {
  if (!payload || typeof payload !== "object") {
    return { items: [], detail: "" };
  }
  return {
    ...payload,
    items: Array.isArray(payload.items) ? payload.items : [],
  };
}

async function searchSiteInn(website) {
  if (!website) {
    return { items: [] };
  }
  const response = await fetch(`/api/site-inn-search?website=${encodeURIComponent(website)}&selectedInn=${encodeURIComponent(innInput.value.trim())}`);
  const payload = await response.json();
  if (!response.ok) {
    throw new Error(payload.detail || "Не удалось проверить ИНН на сайте компании");
  }
  return payload;
}

async function searchCompany(options = {}) {
  const existingInn = innInput.value.trim().replace(/\D+/g, "");
  if (existingInn) {
    resultsBox.innerHTML = "";
    if (!options.auto) {
      setStatus(searchStatus, "ИНН уже указан, поиск не требуется.");
    }
    return;
  }

  const query = companyNameInput.value.trim();
  const website = websiteInput.value.trim();
  if (query.length < 3 && !website) {
    if (!options.auto) {
      setStatus(searchStatus, "Введите хотя бы 3 символа названия или ИНН.", true);
    }
    return;
  }

  const requestId = ++activeSearchRequestId;
  findInnBtn.disabled = true;
  setSearchLoading(true);
  resultsBox.innerHTML = "";
  setStatus(
    searchStatus,
    query.length >= 3 && website
      ? `Ищу "${query}" в ФНС и проверяю ИНН на сайте...`
      : query.length >= 3
        ? `Ищу "${query}" в открытой базе ФНС...`
        : "Проверяю ИНН на сайте компании...",
  );
  try {
    const companyRequest = query.length >= 3
      ? fetch(`/api/company-search?query=${encodeURIComponent(query)}`)
      : Promise.resolve(null);
    const [response, sitePayload] = await Promise.all([
      companyRequest,
      searchSiteInn(website).catch(() => ({ items: [], error: true })),
    ]);
    if (requestId !== activeSearchRequestId) {
      return;
    }

    const payload = normalizeCompanySearchPayload(response ? await response.json() : { items: [] });
    const fnsUnavailable = Boolean(response && !response.ok);
    const siteInns = sitePayload.items || [];
    const singleSiteInn = siteInns.length === 1 ? siteInns[0].inn : "";
    const matchingItem = singleSiteInn ? payload.items.find((item) => item.inn === singleSiteInn) : null;
    if (singleSiteInn && !innInput.value.trim()) {
      innInput.value = singleSiteInn;
      await rememberInnForExistingCard(singleSiteInn);
      await loadDadataByInn(singleSiteInn, { rowNumber: currentRowNumber(), silent: true });
    }
    if (!payload.items.length) {
      if (singleSiteInn) {
        setStatus(
          searchStatus,
          fnsUnavailable
            ? `ФНС сейчас недоступна. ИНН найден на сайте: ${singleSiteInn}.`
            : `В ФНС варианты по названию не найдены. ИНН найден на сайте: ${singleSiteInn}.`,
          fnsUnavailable,
        );
      } else if (fnsUnavailable) {
        setStatus(searchStatus, payload.detail || "ФНС сейчас недоступна. ИНН можно ввести вручную или повторить поиск позже.", true);
      } else {
        setStatus(searchStatus, query.length >= 3 ? "Варианты не найдены. ИНН можно ввести вручную." : "На сайте ИНН не найден.", true);
      }
      return;
    }
    if (matchingItem) {
      setStatus(searchStatus, `ИНН ${singleSiteInn} найден в ФНС и подтвержден сайтом. Выберите подходящую компанию.`);
    } else if (singleSiteInn) {
      setStatus(searchStatus, `На сайте найден ИНН ${singleSiteInn}, но среди вариантов ФНС по названию точного совпадения нет. Проверьте компанию перед выбором.`, true);
    } else {
      setStatus(searchStatus, payload.items.length > 1 ? "Выберите подходящую компанию." : "Найден один вариант.");
    }
    renderCompanyResults(payload.items, siteInns);
  } catch (error) {
    setStatus(searchStatus, error.message, true);
  } finally {
    if (requestId === activeSearchRequestId) {
      setSearchLoading(false);
      findInnBtn.disabled = false;
    }
  }
}

function formPayload() {
  const data = Object.fromEntries(new FormData(form).entries());
  data.preferredNetworkCategories = [...form.querySelectorAll("input[name='preferredNetworkCategories']:checked")]
    .map((item) => item.value)
    .join(", ");
  data.preferredNetworks = [...form.querySelectorAll("input[name='preferredNetworks']:checked")]
    .map((item) => item.value)
    .join(", ");
  return data;
}

async function saveCard(event) {
  event.preventDefault();
  setStatus(saveStatus, "Сохраняю параметры...");

  const body = new FormData();
  body.append("payload", JSON.stringify(formPayload()));
  const photo = document.querySelector("#photo").files[0];
  if (photo) {
    body.append("photo", photo);
  }

  const submit = form.querySelector("button[type='submit']");
  submit.disabled = true;
  try {
    const response = await fetch("/api/cards", { method: "POST", body });
    const payload = await response.json();
    if (!response.ok) {
      throw new Error(payload.detail || "Не удалось сохранить параметры");
    }
    setStatus(saveStatus, `Готово: запись добавлена в строку ${payload.row}.`);
    const wasUpdate = Boolean(document.querySelector("#rowNumber").value);
    form.reset();
    document.querySelector("#rowNumber").value = "";
    document.querySelector("input[name='country']").value = "Россия";
    resultsBox.innerHTML = "";
    clearDadataDetails();
    clearAiSelection();
    searchStatus.textContent = "";
    await loadOptions();
    if (wasUpdate) {
      setStatus(saveStatus, `Готово: строка ${payload.row} обновлена.`);
    }
  } catch (error) {
    setStatus(saveStatus, error.message, true);
  } finally {
    submit.disabled = false;
  }
}

function updatePresentationProgress(job) {
  const progress = Math.max(0, Math.min(100, Number(job.progress) || 0));
  const message = job.message || "Формирую презентацию";
  presentationProgress.hidden = false;
  presentationProgressText.textContent = message;
  presentationProgressPercent.textContent = `${progress}%`;
  presentationProgressBar.style.width = `${progress}%`;
  if (job.status === "done") {
    buildPresentationBtn.textContent = "Презентация готова";
  } else if (message.toLowerCase().includes("остановлена")) {
    buildPresentationBtn.textContent = "Сборка остановлена";
  } else {
    buildPresentationBtn.textContent = `Формируется... ${progress}%`;
  }
}

function resetPresentationProgress() {
  activePresentationJobId = null;
  presentationProgress.hidden = true;
  presentationProgressText.textContent = "Подготовка";
  presentationProgressPercent.textContent = "0%";
  presentationProgressBar.style.width = "0%";
  buildPresentationBtn.textContent = buildPresentationDefaultText;
}

async function pollPresentationJob(jobId) {
  while (true) {
    await new Promise((resolve) => setTimeout(resolve, 900));
    if (activePresentationJobId !== jobId) {
      return;
    }
    const response = await fetch(`/api/presentations/jobs/${jobId}`);
    const job = await response.json();
    if (!response.ok) {
      throw new Error(job.detail || "Не удалось получить статус сборки");
    }

    updatePresentationProgress(job);
    if (job.status === "done") {
      saveStatus.classList.remove("error");
      const openText = job.openStatus === "opened" ? " Файл открыт на компьютере." : "";
      saveStatus.innerHTML = `Готово: <a href="${job.downloadUrl}" target="_blank" rel="noreferrer">${job.fileName}</a>${openText}`;
      return;
    }
    if (job.status === "error") {
      throw new Error(job.message || "Не удалось сформировать презентацию");
    }
  }
}

findInnBtn.addEventListener("click", searchCompany);
form.addEventListener("submit", saveCard);
document.querySelector("#companyTypeSelect").addEventListener("change", syncCompanyType);
document.querySelector("#companyTypeCustom").addEventListener("input", syncCompanyType);
selectExistingCompany.addEventListener("click", () => {
  if (existingCompanySelect.value) {
    applyExistingCard(existingCompanySelect.value);
  }
});
refreshDadataBtn.addEventListener("click", () => {
  loadDadataByInn(innInput.value, { rowNumber: currentRowNumber() });
});
toggleDadataBtn.addEventListener("click", () => {
  setDadataExpanded(!dadataExpanded);
});
prepareAiSelectionBtn.addEventListener("click", prepareAiSelection);
finalizeAiSelectionBtn.addEventListener("click", finalizeAiSelection);
approveAiSelectionBtn.addEventListener("click", () => approveAiSelection("manual"));
approveAiFinalBtn.addEventListener("click", () => approveAiSelection("ai"));
form.querySelectorAll("input[name='aiProvider']").forEach((item) => {
  item.addEventListener("change", clearAiSelection);
});
aiSelectionPanel.addEventListener("change", (event) => {
  if (!(event.target instanceof HTMLInputElement) || event.target.type !== "checkbox") {
    return;
  }
  const limit = Number(event.target.dataset.aiLimit || 0);
  const selector = event.target.dataset.aiKind === "photo" ? "input[data-ai-photo-id]" : "input[data-ai-review-id]";
  const checked = [...aiSelectionPanel.querySelectorAll(`${selector}:checked`)];
  if (limit > 0 && checked.length > limit) {
    event.target.checked = false;
    setStatus(
      aiSelectionStatus,
      event.target.dataset.aiKind === "photo"
        ? `Можно выбрать только ${limit} фотографий.`
        : `Можно выбрать только ${limit} отзывов.`,
      true,
    );
    syncSelectionApprovalState();
    return;
  }
  approvedAiSelectionId = "";
  if (activeAiSelection) {
    activeAiSelection.approved = false;
  }
  syncSelectionApprovalState();
  refreshAiDebugPayload();
  const counts = selectedCounts();
  const requiredPhotos = Number(activeAiSelection?.selection?.required_photo_count || 0);
  const requiredReviews = Number(activeAiSelection?.selection?.required_review_count || 0);
  if (counts.photoIds.length === requiredPhotos && counts.reviewIds.length === requiredReviews) {
    setStatus(aiSelectionStatus, "Выбор изменен. Одобрите его заново.");
  } else {
    setStatus(aiSelectionStatus, `Сейчас выбрано фото: ${counts.photoIds.length}/${requiredPhotos}, отзывы: ${counts.reviewIds.length}/${requiredReviews}.`);
  }
});

async function buildPresentation() {
  const company = companyNameInput.value.trim();
  if (!company) {
    setStatus(saveStatus, "Сначала выберите или заполните компанию.", true);
    return;
  }
  if (!approvedAiSelectionId) {
    setStatus(saveStatus, "Сначала подготовьте и одобрите выбор материалов.", true);
    return;
  }
  buildPresentationBtn.disabled = true;
  const requestId = `pending-${Date.now()}-${Math.random()}`;
  activePresentationJobId = requestId;
  let startedJobId = null;
  setStatus(saveStatus, `Формирую презентацию для ${company}...`);
  updatePresentationProgress({ progress: 3, message: "Отправляю задачу на сборку" });
  try {
    const response = await fetch("/api/presentations", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ companyName: company, approvedSelectionId: approvedAiSelectionId }),
    });
    const payload = await response.json();
    if (!response.ok) {
      throw new Error(payload.detail || "Не удалось сформировать презентацию");
    }

    if (activePresentationJobId !== requestId) {
      return;
    }
    startedJobId = payload.id;
    activePresentationJobId = startedJobId;
    updatePresentationProgress(payload);
    await pollPresentationJob(startedJobId);
  } catch (error) {
    setStatus(saveStatus, error.message, true);
    if (activePresentationJobId === requestId || activePresentationJobId === startedJobId) {
      updatePresentationProgress({ progress: 100, message: "Сборка остановлена" });
    }
  } finally {
    buildPresentationBtn.disabled = false;
    if (!activePresentationJobId || presentationProgress.hidden) {
      buildPresentationBtn.textContent = buildPresentationDefaultText;
    }
  }
}

buildPresentationBtn.addEventListener("click", buildPresentation);

loadOptions().catch((error) => setStatus(saveStatus, error.message, true));
