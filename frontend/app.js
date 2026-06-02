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
const approveAiSelectionBtn = document.querySelector("#approveAiSelectionBtn");
const aiSelectionStatus = document.querySelector("#aiSelectionStatus");
const aiSelectionPanel = document.querySelector("#aiSelectionPanel");
const aiSelectionSource = document.querySelector("#aiSelectionSource");
const aiSelectionRationale = document.querySelector("#aiSelectionRationale");
const aiStatsChoice = document.querySelector("#aiStatsChoice");
const aiPhotosChoice = document.querySelector("#aiPhotosChoice");
const aiReviewsChoice = document.querySelector("#aiReviewsChoice");
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

function clearAiSelection() {
  activeAiSelection = null;
  approvedAiSelectionId = "";
  approveAiSelectionBtn.disabled = true;
  aiSelectionPanel.hidden = true;
  aiSelectionSource.textContent = "";
  aiSelectionRationale.textContent = "";
  aiStatsChoice.innerHTML = "";
  aiPhotosChoice.innerHTML = "";
  aiReviewsChoice.innerHTML = "";
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

function renderAiSelection(payload) {
  const selection = payload.selection || {};
  const rationale = selection.rationale || {};
  activeAiSelection = payload;
  approvedAiSelectionId = payload.approved ? payload.id : "";
  approveAiSelectionBtn.disabled = Boolean(payload.approved);

  aiSelectionSource.textContent = `Источник: ${selection.source || payload.provider || "rules"}`;
  aiSelectionRationale.textContent = rationale.summary || "Выбор подготовлен.";

  aiStatsChoice.innerHTML = "";
  aiPhotosChoice.innerHTML = "";
  aiReviewsChoice.innerHTML = "";

  const stats = selection.stats || {};
  appendChoice(aiStatsChoice, "Категория", stats.category || selection.category || "");
  appendChoice(aiStatsChoice, "Переговоров", stats.negotiations || "");
  appendChoice(aiStatsChoice, "Интерес / чемпионы", `${stats.interest || ""}${stats.champions ? ` / ${stats.champions}` : ""}`);
  if (rationale.stats) {
    appendChoice(aiStatsChoice, "Почему", rationale.stats);
  }

  Object.entries(selection.planned_images || {}).forEach(([token, path]) => {
    appendChoice(aiPhotosChoice, token, path);
  });
  if (!aiPhotosChoice.children.length) {
    appendChoice(aiPhotosChoice, "Фото", "Фотографии не выбраны");
  }
  if (rationale.photos) {
    appendChoice(aiPhotosChoice, "Почему", rationale.photos);
  }

  (selection.reviews || []).forEach((review, index) => {
    const name = review.company || `Отзыв ${index + 1}`;
    const text = [review.person, review.text].filter(Boolean).join(" - ");
    appendChoice(aiReviewsChoice, name, text);
  });
  if (!aiReviewsChoice.children.length) {
    appendChoice(aiReviewsChoice, "Отзывы", "Отзывы не выбраны");
  }
  if (rationale.reviews) {
    appendChoice(aiReviewsChoice, "Почему", rationale.reviews);
  }

  aiSelectionPanel.hidden = false;
}

async function prepareAiSelection() {
  const company = companyNameInput.value.trim();
  if (!company) {
    setStatus(aiSelectionStatus, "Сначала выберите или заполните компанию.", true);
    return;
  }

  clearAiSelection();
  prepareAiSelectionBtn.disabled = true;
  setStatus(aiSelectionStatus, "Готовлю выбор материалов...");
  try {
    const response = await fetch("/api/ai-selection", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ companyName: company, provider: selectedAiProvider() }),
    });
    const payload = await response.json();
    if (!response.ok) {
      throw new Error(payload.detail || "Не удалось подготовить выбор материалов");
    }
    renderAiSelection(payload);
    setStatus(aiSelectionStatus, "Проверьте выбор и нажмите «Одобрить выбор».");
  } catch (error) {
    setStatus(aiSelectionStatus, error.message, true);
  } finally {
    prepareAiSelectionBtn.disabled = false;
  }
}

async function approveAiSelection() {
  if (!activeAiSelection?.id) {
    setStatus(aiSelectionStatus, "Сначала подготовьте выбор материалов.", true);
    return;
  }

  approveAiSelectionBtn.disabled = true;
  setStatus(aiSelectionStatus, "Одобряю выбор...");
  try {
    const response = await fetch(`/api/ai-selection/${encodeURIComponent(activeAiSelection.id)}/approve`, {
      method: "POST",
    });
    const payload = await response.json();
    if (!response.ok) {
      throw new Error(payload.detail || "Не удалось одобрить выбор");
    }
    approvedAiSelectionId = payload.id;
    activeAiSelection.approved = true;
    setStatus(aiSelectionStatus, "Выбор одобрен. Можно формировать презентацию.");
  } catch (error) {
    approveAiSelectionBtn.disabled = false;
    setStatus(aiSelectionStatus, error.message, true);
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
  return valueBy(card, "ИНН") || valueBy(card, "РРќРќ") || valueBy(card, "Р ВР РќР Рќ");
}

function contactNameBy(card) {
  return (
    valueBy(card, "ФИО контакта") ||
    valueBy(card, "Р¤РРћ контакта") ||
    valueBy(card, "Р¤РРћ РєРѕРЅС‚Р°РєС‚Р°") ||
    valueBy(card, "Р В¤Р ВР С› РєРѕРЅС‚Р°РєС‚Р°")
  );
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
      saveStatus.innerHTML = `Готово: <a href="${job.downloadUrl}" target="_blank" rel="noreferrer">${job.fileName}</a>`;
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
approveAiSelectionBtn.addEventListener("click", approveAiSelection);
form.querySelectorAll("input[name='aiProvider']").forEach((item) => {
  item.addEventListener("change", clearAiSelection);
});
buildPresentationBtn.addEventListener("click", async () => {
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
});

loadOptions().catch((error) => setStatus(saveStatus, error.message, true));
