const form = document.querySelector("#proposalForm");
const saveStatus = document.querySelector("#saveStatus");
const searchStatus = document.querySelector("#searchStatus");
const resultsBox = document.querySelector("#companyResults");
const findInnBtn = document.querySelector("#findInnBtn");
const companyNameInput = document.querySelector("#companyName");
const innInput = document.querySelector("#inn");
const existingCompanySelect = document.querySelector("#existingCompanySelect");
const selectExistingCompany = document.querySelector("#selectExistingCompany");
const buildPresentationBtn = document.querySelector("#buildPresentationBtn");
const presentationProgress = document.querySelector("#presentationProgress");
const presentationProgressText = document.querySelector("#presentationProgressText");
const presentationProgressPercent = document.querySelector("#presentationProgressPercent");
const presentationProgressBar = document.querySelector("#presentationProgressBar");
const buildPresentationDefaultText = buildPresentationBtn.textContent;
let existingCards = [];
let activePresentationJobId = null;

const optionIds = {
  industries: "industries",
  activities: "activities",
  productCategories: "productCategories",
};
const ADD_COMPANY_TYPE = "__add_company_type__";

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

function valueBy(card, header) {
  return card[header] || card[header.trim()] || "";
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
  setFormValue("_rowNumber", card._rowNumber);
  setFormValue("companyName", valueBy(card, "Название компании"));
  setFormValue("inn", valueBy(card, "РРќРќ"));
  setCompanyType(valueBy(card, "Тип компании"));
  setFormValue("contactName", valueBy(card, "Р¤РРћ контакта"));
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
}

function clearSelectedCompany() {
  form.reset();
  document.querySelector("#rowNumber").value = "";
  document.querySelector("input[name='country']").value = "Россия";
  setCompanyType("");
  setPreferredNetworks("");
  saveStatus.textContent = "";
}

async function loadExistingCards() {
  const response = await fetch("/api/cards");
  if (!response.ok) {
    throw new Error("Не удалось загрузить существующие компании");
  }
  const payload = await response.json();
  existingCards = payload.items || [];
  existingCompanySelect.innerHTML = "";
  existingCards.forEach((card) => {
    const option = document.createElement("option");
    option.value = card._rowNumber;
    option.textContent = `${valueBy(card, "Название компании") || "Компания без названия"}${valueBy(card, "РРќРќ") ? ` · ${valueBy(card, "РРќРќ")}` : ""}`;
    existingCompanySelect.append(option);
  });
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

function renderCompanyResults(items) {
  resultsBox.innerHTML = "";
  items.forEach((item) => {
    const button = document.createElement("button");
    button.type = "button";
    button.className = "result-item";
    button.innerHTML = `
      <strong>${item.name || "Компания без названия"}</strong>
      <small>РРќРќ ${item.inn || "не указан"}${item.address ? ` · ${item.address}` : ""}${item.status ? ` · ${item.status}` : ""}</small>
    `;
    button.addEventListener("click", () => {
      companyNameInput.value = item.name || companyNameInput.value;
      innInput.value = item.inn || innInput.value;
      resultsBox.innerHTML = "";
      setStatus(searchStatus, "Компания выбрана из базы ФНС.");
    });
    resultsBox.append(button);
  });
}

async function searchCompany() {
  const query = companyNameInput.value.trim() || innInput.value.trim();
  if (query.length < 3) {
    setStatus(searchStatus, "Введите хотя бы 3 символа названия или РРќРќ.", true);
    return;
  }

  findInnBtn.disabled = true;
  resultsBox.innerHTML = "";
  setStatus(searchStatus, "РС‰Сѓ компанию в открытой базе ФНС...");
  try {
    const response = await fetch(`/api/company-search?query=${encodeURIComponent(query)}`);
    const payload = await response.json();
    if (!response.ok) {
      throw new Error(payload.detail || "Поиск временно недоступен");
    }
    if (!payload.items.length) {
      setStatus(searchStatus, "Варианты не найдены. РРќРќ можно ввести вручную.", true);
      return;
    }
    setStatus(searchStatus, payload.items.length > 1 ? "Выберите подходящую компанию." : "Найден один вариант.");
    renderCompanyResults(payload.items);
  } catch (error) {
    setStatus(searchStatus, error.message, true);
  } finally {
    findInnBtn.disabled = false;
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
buildPresentationBtn.addEventListener("click", async () => {
  const company = companyNameInput.value.trim();
  if (!company) {
    setStatus(saveStatus, "Сначала выберите или заполните компанию.", true);
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
      body: JSON.stringify({ companyName: company }),
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
