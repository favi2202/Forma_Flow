const state = {
  sessionId: null,
  columns: [],
  files: [],
  datasetGroups: [],
  activeDatasetId: null,
  rowCount: 0,
  language: localStorage.getItem("formaFlowLanguage") || "en",
  previewTimer: null,
};

const translations = {
  en: {
    version: "Local Document Intelligence v0.6.0",
    language: "Language",
    privacyChip: "● Files stay on this computer",
    eyebrow: "Teacher paperwork, minus the copy-paste",
    heroTitle: "Understand the document first. Build the right table second.",
    heroText: "FormaFlow classifies school files, separates different table types, repairs difficult rosters, and lets you export only the dataset you actually need.",
    worksLocally: "Works locally",
    formats: "XLSX · XLS · CSV · DOCX · DOC · PDF → XLSX · CSV · DOCX",
    step1: "Step 1",
    chooseFiles: "Choose school documents",
    checkingServer: "Checking local server…",
    inputFiles: "Excel, Word, CSV, or PDF files",
    uploadHint: "Select one or many files. FormaFlow will classify them and separate incompatible tables instead of mixing everything blindly.",
    processFiles: "Analyze files",
    processing: "Analyzing…",
    stepDataset: "Detected data",
    chooseDataset: "Choose which dataset to work with",
    datasetHelp: "Different structures remain separate, so a financial table or monitoring report cannot accidentally enter a class roster.",
    dataset: "Dataset",
    noDataset: "No reliable data table was found. The file was still classified and its warnings are shown above.",
    step2: "Step 2",
    chooseColumns: "Choose, rename, and reorder columns",
    toggleAll: "Toggle all",
    step3: "Step 3",
    buildColumns: "Add fixed and calculated columns",
    fixedValues: "Same value for every row",
    fixedHelp: "Example: academic year, school, or teacher.",
    addFixed: "+ Fixed column",
    calculatedColumns: "Calculated columns",
    calculatedHelp: "Create next class, row number, copies, or numeric changes.",
    addNextClass: "+ Next class",
    addCalculated: "+ Other",
    nextClassExample: "Next class rule: 5-B → 6-B, 8-A → 9-A, while grades 9 and 11 become empty.",
    cleaning: "Data cleaning",
    cleaningHelp: "Optional corrections applied only to the exported copy.",
    trimWhitespace: "Trim and normalize spaces",
    removeDuplicates: "Remove exact duplicate output rows",
    skipBlank: "Skip rows where this field is blank",
    disabled: "Disabled",
    sortBy: "Sort by",
    originalOrder: "Original order",
    step4: "Step 4",
    previewExport: "Preview and export",
    downloadExcel: "Download Excel",
    downloadCsv: "Download CSV",
    downloadWord: "Download Word",
    clearData: "Clear data",
    serverConnected: "Local server connected",
    serverNotConnected: "Server not connected",
    serverHelp: "Start the app with run.bat or python app.py, then open http://127.0.0.1:8000",
    chooseAtLeastOne: "Choose at least one file.",
    rowsLoaded: "{count} rows loaded from the selected dataset.",
    ready: "Ready",
    error: "Error",
    rows: "{count} rows",
    previewShowing: "Showing {shown} of {count} result rows.",
    selectColumn: "Select at least one column.",
    uploadFirst: "Upload files first.",
    downloadStarted: "{format} download started",
    cleared: "Local page data cleared",
    untitled: "Untitled",
    outputName: "Output column name",
    moveUp: "Move up",
    moveDown: "Move down",
    fixedName: "Column name",
    fixedValue: "Value for every row",
    newColumn: "New column",
    nextClass: "Next class",
    rowNumber: "No.",
    calculatedColumn: "Calculated column",
    rule: "Rule",
    sourceColumn: "Source column",
    stopGrades: "Grades that become empty",
    startNumber: "Start number",
    amount: "Amount to add",
    kindNextClass: "Next class / grade",
    kindSequence: "Sequential number",
    kindAddNumber: "Add number to source",
    kindCopy: "Copy source column",
    remove: "Remove",
    academicYear: "Academic year",
    studentName: "Student name",
    firstName: "First name",
    lastName: "Last name",
    class: "Class",
    promotedClass: "Promoted / next class",
    grade: "Grade",
    age: "Age",
    gender: "Gender",
    birthDate: "Birth date",
    email: "Email",
    phone: "Phone",
    address: "Address",
    parentName: "Relative / guardian",
    parentPhone: "Relative phone",
    parentEmail: "Relative email",
    pinfl: "JSHSHIR",
    total: "Total",
    percentage: "Percentage",
    subject: "Subject",
    teacher: "Teacher",
    place: "Place",
    score: "Score",
    previewLoading: "Updating preview…",
    previewFailed: "Preview could not be updated.",
    unknown: "Unknown",
    methodAlias: "Recognized",
    methodToken: "Similar phrase",
    methodFuzzy: "Fuzzy match",
    methodCustom: "Custom field",
    methodInferred: "Inferred from values",
    methodQuestion: "Question column",
    sensitiveField: "Sensitive field",
    filesAnalyzed: "{count} file(s) analyzed",
    datasetsFound: "{count} compatible dataset group(s) found",
    sources: "{count} source file(s)",
    tables: "{count} table/sheet(s)",
    cleanup: "Cleanup",
    duplicateColumnsRemoved: "{count} duplicated merged column(s) repaired",
    blankRowsRemoved: "{count} blank row(s) removed",
    nonStudentRowsRemoved: "{count} footer/header row(s) removed",
    continuationRowsMerged: "{count} continuation row(s) merged",
    documentTypes: {
      student_roster_document: "Student roster",
      promotion_document: "Class-promotion document",
      monitoring_document: "Monitoring document",
      assessment_results_document: "Assessment results",
      financial_document: "Financial document",
      meeting_minutes: "Meeting minutes / report",
      methodical_document: "Methodical document",
      assessment_material: "Test / lesson material",
      data_document: "Data document",
      unknown_document: "Unknown document",
    },
    datasetTypes: {
      student_roster: "Student rosters",
      promotion_table: "Class promotion tables",
      monitoring_table: "Monitoring tables",
      assessment_results: "Assessment results",
      financial_table: "Financial tables",
      generic_table: "Other data tables",
      unknown_table: "Unclassified tables",
    },
  },
  uz: {
    version: "Mahalliy Document Intelligence v0.6.0",
    language: "Til",
    privacyChip: "● Fayllar shu kompyuterda qoladi",
    eyebrow: "O‘qituvchi hujjatlari — ortiqcha nusxalashsiz",
    heroTitle: "Avval hujjatni tushuning. Keyin kerakli jadvalni yarating.",
    heroText: "FormaFlow maktab fayllarini tasniflaydi, turli jadvallarni aralashtirmaydi, murakkab ro‘yxatlarni tuzatadi va faqat kerakli ma’lumotni eksport qiladi.",
    worksLocally: "Mahalliy ishlaydi",
    formats: "XLSX · XLS · CSV · DOCX · DOC · PDF → XLSX · CSV · DOCX",
    step1: "1-qadam",
    chooseFiles: "Maktab hujjatlarini tanlang",
    checkingServer: "Mahalliy server tekshirilmoqda…",
    inputFiles: "Excel, Word, CSV yoki PDF fayllari",
    uploadHint: "Bir yoki bir nechta faylni tanlang. FormaFlow ularni tasniflaydi va mos kelmaydigan jadvallarni alohida saqlaydi.",
    processFiles: "Fayllarni tahlil qilish",
    processing: "Tahlil qilinmoqda…",
    stepDataset: "Aniqlangan ma’lumot",
    chooseDataset: "Ishlanadigan ma’lumotlar to‘plamini tanlang",
    datasetHelp: "Turli tuzilmalar alohida qoladi, shuning uchun moliyaviy jadval yoki monitoring ro‘yxatga tasodifan qo‘shilmaydi.",
    dataset: "Ma’lumotlar to‘plami",
    noDataset: "Ishonchli jadval topilmadi. Fayl tasniflandi va ogohlantirishlar yuqorida ko‘rsatildi.",
    step2: "2-qadam",
    chooseColumns: "Ustunlarni tanlang, nomlang va tartiblang",
    toggleAll: "Barchasini almashtirish",
    step3: "3-qadam",
    buildColumns: "Doimiy va hisoblanadigan ustunlar qo‘shing",
    fixedValues: "Har bir qator uchun bir xil qiymat",
    fixedHelp: "Masalan: o‘quv yili, maktab yoki o‘qituvchi.",
    addFixed: "+ Doimiy ustun",
    calculatedColumns: "Hisoblanadigan ustunlar",
    calculatedHelp: "Keyingi sinf, tartib raqami, nusxa yoki sonli o‘zgarish yarating.",
    addNextClass: "+ Keyingi sinf",
    addCalculated: "+ Boshqa",
    nextClassExample: "Keyingi sinf qoidasi: 5-B → 6-B, 8-A → 9-A; 9 va 11-sinflar bo‘sh qoladi.",
    cleaning: "Ma’lumotlarni tozalash",
    cleaningHelp: "Faqat eksport qilinadigan nusxaga qo‘llanadigan ixtiyoriy tuzatishlar.",
    trimWhitespace: "Ortiqcha bo‘sh joylarni tozalash",
    removeDuplicates: "Bir xil natija qatorlarini o‘chirish",
    skipBlank: "Ushbu maydon bo‘sh bo‘lgan qatorlarni tashlab ketish",
    disabled: "O‘chirilgan",
    sortBy: "Saralash",
    originalOrder: "Asl tartib",
    step4: "4-qadam",
    previewExport: "Ko‘rib chiqish va eksport",
    downloadExcel: "Excel yuklash",
    downloadCsv: "CSV yuklash",
    downloadWord: "Word yuklash",
    clearData: "Ma’lumotlarni tozalash",
    serverConnected: "Mahalliy server ulandi",
    serverNotConnected: "Server ulanmagan",
    serverHelp: "run.bat yoki python app.py ni ishga tushiring, keyin http://127.0.0.1:8000 manzilini oching",
    chooseAtLeastOne: "Kamida bitta fayl tanlang.",
    rowsLoaded: "Tanlangan to‘plamdan {count} qator yuklandi.",
    ready: "Tayyor",
    error: "Xato",
    rows: "{count} qator",
    previewShowing: "Natijaning {count} qatoridan {shown} tasi ko‘rsatilmoqda.",
    selectColumn: "Kamida bitta ustunni tanlang.",
    uploadFirst: "Avval fayllarni yuklang.",
    downloadStarted: "{format} yuklanishi boshlandi",
    cleared: "Mahalliy sahifa ma’lumotlari tozalandi",
    untitled: "Nomsiz",
    outputName: "Natija ustuni nomi",
    moveUp: "Yuqoriga",
    moveDown: "Pastga",
    fixedName: "Ustun nomi",
    fixedValue: "Har bir qator uchun qiymat",
    newColumn: "Yangi ustun",
    nextClass: "Keyingi sinf",
    rowNumber: "№",
    calculatedColumn: "Hisoblangan ustun",
    rule: "Qoida",
    sourceColumn: "Manba ustuni",
    stopGrades: "Bo‘sh qoladigan sinflar",
    startNumber: "Boshlang‘ich raqam",
    amount: "Qo‘shiladigan son",
    kindNextClass: "Keyingi sinf",
    kindSequence: "Tartib raqami",
    kindAddNumber: "Manba qiymatiga son qo‘shish",
    kindCopy: "Manba ustunini nusxalash",
    remove: "O‘chirish",
    academicYear: "O‘quv yili",
    studentName: "O‘quvchi F.I.Sh.",
    firstName: "Ism",
    lastName: "Familiya",
    class: "Sinf",
    promotedClass: "Ko‘chirilgan / keyingi sinf",
    grade: "Baho",
    age: "Yosh",
    gender: "Jinsi",
    birthDate: "Tug‘ilgan sana",
    email: "Email",
    phone: "Telefon",
    address: "Manzil",
    parentName: "Qarindosh / ota-ona",
    parentPhone: "Qarindosh telefoni",
    parentEmail: "Qarindosh e-pochtasi",
    pinfl: "JSHSHIR",
    total: "Jami",
    percentage: "Foiz",
    subject: "Fan",
    teacher: "O‘qituvchi",
    place: "O‘rni",
    score: "Ball",
    previewLoading: "Ko‘rib chiqish yangilanmoqda…",
    previewFailed: "Ko‘rib chiqishni yangilab bo‘lmadi.",
    unknown: "Noma’lum",
    methodAlias: "Aniqlandi",
    methodToken: "O‘xshash ibora",
    methodFuzzy: "Taxminiy moslik",
    methodCustom: "Maxsus maydon",
    methodInferred: "Qiymatlardan aniqlandi",
    methodQuestion: "Savol ustuni",
    sensitiveField: "Maxfiy maydon",
    filesAnalyzed: "{count} ta fayl tahlil qilindi",
    datasetsFound: "{count} ta mos ma’lumot guruhi topildi",
    sources: "{count} ta manba fayli",
    tables: "{count} ta jadval/varaq",
    cleanup: "Tozalash",
    duplicateColumnsRemoved: "{count} ta takroriy birlashtirilgan ustun tuzatildi",
    blankRowsRemoved: "{count} ta bo‘sh qator o‘chirildi",
    nonStudentRowsRemoved: "{count} ta sarlavha/pastki qator o‘chirildi",
    continuationRowsMerged: "{count} ta davomiy qator birlashtirildi",
    documentTypes: {
      student_roster_document: "O‘quvchilar ro‘yxati",
      promotion_document: "Sinfdan-sinfga ko‘chirish hujjati",
      monitoring_document: "Monitoring hujjati",
      assessment_results_document: "Baholash natijalari",
      financial_document: "Moliyaviy hujjat",
      meeting_minutes: "Yig‘ilish bayonnomasi / hisobot",
      methodical_document: "Metodik hujjat",
      assessment_material: "Test / dars materiali",
      data_document: "Ma’lumot hujjati",
      unknown_document: "Noma’lum hujjat",
    },
    datasetTypes: {
      student_roster: "O‘quvchilar ro‘yxatlari",
      promotion_table: "Sinfdan-sinfga ko‘chirish jadvallari",
      monitoring_table: "Monitoring jadvallari",
      assessment_results: "Baholash natijalari",
      financial_table: "Moliyaviy jadvallar",
      generic_table: "Boshqa ma’lumot jadvallari",
      unknown_table: "Tasniflanmagan jadvallar",
    },
  },
  ru: {
    version: "Локальный Document Intelligence v0.6.0",
    language: "Язык",
    privacyChip: "● Файлы остаются на этом компьютере",
    eyebrow: "Школьные документы без бесконечного копирования",
    heroTitle: "Сначала понять документ. Затем собрать нужную таблицу.",
    heroText: "FormaFlow классифицирует школьные файлы, разделяет разные типы таблиц, исправляет сложные списки и экспортирует только нужный набор данных.",
    worksLocally: "Работает локально",
    formats: "XLSX · XLS · CSV · DOCX · DOC · PDF → XLSX · CSV · DOCX",
    step1: "Шаг 1",
    chooseFiles: "Выберите школьные документы",
    checkingServer: "Проверка локального сервера…",
    inputFiles: "Файлы Excel, Word, CSV или PDF",
    uploadHint: "Выберите один или несколько файлов. FormaFlow классифицирует их и не смешивает несовместимые таблицы.",
    processFiles: "Анализировать файлы",
    processing: "Анализ…",
    stepDataset: "Обнаруженные данные",
    chooseDataset: "Выберите набор данных для работы",
    datasetHelp: "Разные структуры остаются отдельными, поэтому финансовая или мониторинговая таблица не попадёт в список класса случайно.",
    dataset: "Набор данных",
    noDataset: "Надёжная таблица не найдена. Файл классифицирован, предупреждения показаны выше.",
    step2: "Шаг 2",
    chooseColumns: "Выберите, переименуйте и расположите столбцы",
    toggleAll: "Переключить все",
    step3: "Шаг 3",
    buildColumns: "Добавьте постоянные и вычисляемые столбцы",
    fixedValues: "Одинаковое значение для каждой строки",
    fixedHelp: "Например: учебный год, школа или учитель.",
    addFixed: "+ Постоянный столбец",
    calculatedColumns: "Вычисляемые столбцы",
    calculatedHelp: "Создайте следующий класс, номер строки, копию или числовое изменение.",
    addNextClass: "+ Следующий класс",
    addCalculated: "+ Другое",
    nextClassExample: "Правило следующего класса: 5-Б → 6-Б, 8-А → 9-А; для 9 и 11 классов значение пустое.",
    cleaning: "Очистка данных",
    cleaningHelp: "Необязательные исправления применяются только к экспортируемой копии.",
    trimWhitespace: "Убирать лишние пробелы",
    removeDuplicates: "Удалять одинаковые строки результата",
    skipBlank: "Пропускать строки, где это поле пустое",
    disabled: "Отключено",
    sortBy: "Сортировать по",
    originalOrder: "Исходный порядок",
    step4: "Шаг 4",
    previewExport: "Предпросмотр и экспорт",
    downloadExcel: "Скачать Excel",
    downloadCsv: "Скачать CSV",
    downloadWord: "Скачать Word",
    clearData: "Очистить данные",
    serverConnected: "Локальный сервер подключен",
    serverNotConnected: "Сервер не подключен",
    serverHelp: "Запустите run.bat или python app.py и откройте http://127.0.0.1:8000",
    chooseAtLeastOne: "Выберите хотя бы один файл.",
    rowsLoaded: "Загружено строк из выбранного набора: {count}.",
    ready: "Готово",
    error: "Ошибка",
    rows: "{count} строк",
    previewShowing: "Показано {shown} из {count} строк результата.",
    selectColumn: "Выберите хотя бы один столбец.",
    uploadFirst: "Сначала загрузите файлы.",
    downloadStarted: "Загрузка {format} началась",
    cleared: "Локальные данные страницы очищены",
    untitled: "Без названия",
    outputName: "Название выходного столбца",
    moveUp: "Выше",
    moveDown: "Ниже",
    fixedName: "Название столбца",
    fixedValue: "Значение для каждой строки",
    newColumn: "Новый столбец",
    nextClass: "Следующий класс",
    rowNumber: "№",
    calculatedColumn: "Вычисляемый столбец",
    rule: "Правило",
    sourceColumn: "Исходный столбец",
    stopGrades: "Классы с пустым результатом",
    startNumber: "Начальный номер",
    amount: "Прибавляемое число",
    kindNextClass: "Следующий класс",
    kindSequence: "Порядковый номер",
    kindAddNumber: "Прибавить число",
    kindCopy: "Копировать столбец",
    remove: "Удалить",
    academicYear: "Учебный год",
    studentName: "Ф.И.О. ученика",
    firstName: "Имя",
    lastName: "Фамилия",
    class: "Класс",
    promotedClass: "Переведён / следующий класс",
    grade: "Оценка",
    age: "Возраст",
    gender: "Пол",
    birthDate: "Дата рождения",
    email: "Email",
    phone: "Телефон",
    address: "Адрес",
    parentName: "Родственник / родитель",
    parentPhone: "Телефон родственника",
    parentEmail: "Email родственника",
    pinfl: "ПИНФЛ",
    total: "Итого",
    percentage: "Процент",
    subject: "Предмет",
    teacher: "Учитель",
    place: "Место",
    score: "Балл",
    previewLoading: "Обновление предпросмотра…",
    previewFailed: "Не удалось обновить предпросмотр.",
    unknown: "Неизвестно",
    methodAlias: "Распознано",
    methodToken: "Похожая фраза",
    methodFuzzy: "Нечёткое совпадение",
    methodCustom: "Пользовательское поле",
    methodInferred: "Определено по значениям",
    methodQuestion: "Столбец вопроса",
    sensitiveField: "Конфиденциальное поле",
    filesAnalyzed: "Проанализировано файлов: {count}",
    datasetsFound: "Найдено совместимых наборов: {count}",
    sources: "Исходных файлов: {count}",
    tables: "Таблиц/листов: {count}",
    cleanup: "Очистка",
    duplicateColumnsRemoved: "Исправлено повторяющихся объединённых столбцов: {count}",
    blankRowsRemoved: "Удалено пустых строк: {count}",
    nonStudentRowsRemoved: "Удалено заголовков/служебных строк: {count}",
    continuationRowsMerged: "Объединено продолжений строк: {count}",
    documentTypes: {
      student_roster_document: "Список учеников",
      promotion_document: "Документ о переводе классов",
      monitoring_document: "Мониторинговый документ",
      assessment_results_document: "Результаты оценивания",
      financial_document: "Финансовый документ",
      meeting_minutes: "Протокол / отчёт",
      methodical_document: "Методический документ",
      assessment_material: "Тест / учебный материал",
      data_document: "Документ с данными",
      unknown_document: "Неизвестный документ",
    },
    datasetTypes: {
      student_roster: "Списки учеников",
      promotion_table: "Таблицы перевода классов",
      monitoring_table: "Мониторинговые таблицы",
      assessment_results: "Результаты оценивания",
      financial_table: "Финансовые таблицы",
      generic_table: "Другие таблицы данных",
      unknown_table: "Неклассифицированные таблицы",
    },
  },
};

const $ = (id) => document.getElementById(id);
const uploadForm = $("uploadForm");
const fileInput = $("fileInput");
const processButton = $("processButton");
const uploadMessage = $("uploadMessage");
const fileResults = $("fileResults");
const datasetPanel = $("datasetPanel");
const datasetSelect = $("datasetSelect");
const datasetInfo = $("datasetInfo");
const builder = $("builder");
const columnList = $("columnList");
const fixedList = $("fixedList");
const derivedList = $("derivedList");
const skipBlankKey = $("skipBlankKey");
const sortKey = $("sortKey");
const previewHead = $("previewHead");
const previewBody = $("previewBody");
const previewNote = $("previewNote");
const rowCount = $("rowCount");
const toast = $("toast");

function t(key, params = {}) {
  let value = translations[state.language]?.[key] ?? translations.en[key] ?? key;
  if (typeof value !== "string") return value;
  for (const [name, replacement] of Object.entries(params)) {
    value = value.replaceAll(`{${name}}`, String(replacement));
  }
  return value;
}

function translatedMap(group, key) {
  return translations[state.language]?.[group]?.[key] ?? translations.en[group]?.[key] ?? key.replaceAll("_", " ");
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function showToast(message) {
  toast.textContent = message;
  toast.classList.add("show");
  clearTimeout(showToast.timer);
  showToast.timer = setTimeout(() => toast.classList.remove("show"), 2600);
}

function showMessage(message, kind = "info") {
  uploadMessage.textContent = message;
  uploadMessage.className = `message ${kind}`;
}

function hideMessage() {
  uploadMessage.className = "message hidden";
  uploadMessage.textContent = "";
}

function applyTranslations() {
  document.documentElement.lang = state.language;
  document.querySelectorAll("[data-i18n]").forEach((element) => {
    const key = element.dataset.i18n;
    if (translations[state.language]?.[key]) element.textContent = t(key);
  });
  $("languageSelect").value = state.language;
  const defaultAcademic = fixedList.querySelector('[data-default-key="academicYear"]');
  if (defaultAcademic && defaultAcademic.dataset.auto !== "false") defaultAcademic.value = t("academicYear");
  renderDatasetPanel();
  renderFiles();
  renderColumns();
  populateSourceSelects();
  renderDerivedRows(true);
}

$("languageSelect").addEventListener("change", (event) => {
  state.language = event.target.value;
  localStorage.setItem("formaFlowLanguage", state.language);
  applyTranslations();
  renderPreviewDebounced();
});

async function checkServer() {
  const status = $("serverStatus");
  try {
    const response = await fetch("/health");
    if (!response.ok) throw new Error("health failed");
    status.textContent = t("serverConnected");
    status.className = "status good";
  } catch {
    status.textContent = t("serverNotConnected");
    status.className = "status bad";
    showMessage(t("serverHelp"), "error");
  }
}

uploadForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  if (!fileInput.files.length) {
    showMessage(t("chooseAtLeastOne"), "error");
    return;
  }

  hideMessage();
  processButton.disabled = true;
  processButton.textContent = t("processing");
  builder.classList.add("hidden");
  datasetPanel.classList.add("hidden");

  const formData = new FormData();
  [...fileInput.files].forEach((file) => formData.append("files", file));

  try {
    const response = await fetch("/api/upload", { method: "POST", body: formData });
    const payload = await response.json();
    if (!response.ok) {
      const detail = payload.detail;
      if (detail?.files) {
        state.files = detail.files;
        renderFiles();
        fileResults.classList.remove("hidden");
      }
      throw new Error(detail?.message || detail || t("previewFailed"));
    }

    state.sessionId = payload.session_id;
    state.files = payload.files || [];
    state.datasetGroups = payload.dataset_groups || [];
    state.activeDatasetId = payload.active_dataset_id;
    renderFiles();
    renderDatasetPanel();
    fileResults.classList.remove("hidden");

    if (state.activeDatasetId) {
      applyActiveDataset({
        columns: payload.columns || [],
        row_count: payload.row_count || 0,
        active_dataset_id: payload.active_dataset_id,
      });
      showMessage(t("rowsLoaded", { count: payload.row_count }), "info");
    } else {
      state.columns = [];
      state.rowCount = 0;
      builder.classList.add("hidden");
      showMessage(t("noDataset"), "info");
    }
  } catch (error) {
    showMessage(error.message || String(error), "error");
  } finally {
    processButton.disabled = false;
    processButton.textContent = t("processFiles");
  }
});

function renderFiles() {
  if (!state.files.length) return;
  fileResults.innerHTML = state.files.map((file) => {
    if (!file.ok) {
      return `<article class="file-card error">
        <span class="file-icon">!</span>
        <div><strong>${escapeHtml(file.name)}</strong><small>${escapeHtml(file.error || t("error"))}</small></div>
        <span class="status bad">${escapeHtml(t("error"))}</span>
      </article>`;
    }

    const documentType = translatedMap("documentTypes", file.document_type);
    const datasetLabels = (file.dataset_types || []).map((type) => translatedMap("datasetTypes", type)).join(", ");
    const notes = [];
    if (file.rows) notes.push(t("rows", { count: file.rows }));
    if (file.table_count != null) notes.push(t("tables", { count: file.table_count }));
    if (file.sheet_count != null) notes.push(t("tables", { count: file.sheet_count }));
    const warnings = (file.warnings || []).map((warning) => `<span class="file-warning">${escapeHtml(warning)}</span>`).join("");
    const reasons = (file.classification_reasons || []).slice(0, 2).map(escapeHtml).join(" · ");

    return `<article class="file-card file-card-rich">
      <span class="file-icon ${file.source_kind === "word" ? "word-icon" : file.source_kind === "pdf" ? "pdf-icon" : ""}">${file.source_kind === "word" ? "W" : file.source_kind === "pdf" ? "P" : "X"}</span>
      <div>
        <strong>${escapeHtml(file.name)}</strong>
        <small>${escapeHtml(documentType)} · ${Math.round((file.document_confidence || 0) * 100)}%</small>
        <div class="file-notes">
          ${datasetLabels ? `<span>${escapeHtml(datasetLabels)}</span>` : `<span>${escapeHtml(t("noDataset"))}</span>`}
          ${notes.length ? `<span>${escapeHtml(notes.join(" · "))}</span>` : ""}
          ${reasons ? `<span class="reason-note">${reasons}</span>` : ""}
          ${warnings}
        </div>
      </div>
      <span class="status good">${escapeHtml(t("ready"))}</span>
    </article>`;
  }).join("");
}

function renderDatasetPanel() {
  if (!state.datasetGroups.length) {
    datasetPanel.classList.add("hidden");
    datasetSelect.innerHTML = "";
    datasetInfo.innerHTML = "";
    return;
  }

  datasetPanel.classList.remove("hidden");
  datasetSelect.innerHTML = state.datasetGroups.map((group) => {
    const label = translatedMap("datasetTypes", group.dataset_type);
    return `<option value="${escapeHtml(group.id)}" ${group.id === state.activeDatasetId ? "selected" : ""}>${escapeHtml(label)} — ${group.row_count} rows — ${group.source_count} source(s)</option>`;
  }).join("");
  renderDatasetInfo();
}

function renderDatasetInfo() {
  const group = state.datasetGroups.find((item) => item.id === state.activeDatasetId);
  if (!group) {
    datasetInfo.innerHTML = "";
    return;
  }
  const diagnostics = group.diagnostics || {};
  const chips = [
    t("rows", { count: group.row_count }),
    t("sources", { count: group.source_count }),
    t("tables", { count: group.table_count }),
  ];
  if (diagnostics.duplicate_columns_removed) chips.push(t("duplicateColumnsRemoved", { count: diagnostics.duplicate_columns_removed }));
  if (diagnostics.blank_rows_removed) chips.push(t("blankRowsRemoved", { count: diagnostics.blank_rows_removed }));
  if (diagnostics.non_student_rows_removed) chips.push(t("nonStudentRowsRemoved", { count: diagnostics.non_student_rows_removed }));
  if (diagnostics.continuation_rows_merged) chips.push(t("continuationRowsMerged", { count: diagnostics.continuation_rows_merged }));

  datasetInfo.innerHTML = `<div class="dataset-summary">
    <div><strong>${escapeHtml(translatedMap("datasetTypes", group.dataset_type))}</strong><span>${escapeHtml(t("dataset"))}</span></div>
    <div><strong>${escapeHtml(group.row_count)}</strong><span>${escapeHtml(t("rows", { count: group.row_count }))}</span></div>
    <div><strong>${escapeHtml(group.source_count)}</strong><span>${escapeHtml(t("sources", { count: group.source_count }))}</span></div>
    <div><strong>${escapeHtml(group.table_count)}</strong><span>${escapeHtml(t("tables", { count: group.table_count }))}</span></div>
  </div>
  <div class="reason-note">${escapeHtml((group.sources || []).join(" · "))}</div>
  <div class="diagnostic-chips">${chips.map((chip) => `<span>${escapeHtml(chip)}</span>`).join("")}</div>`;
}

datasetSelect.addEventListener("change", async () => {
  if (!state.sessionId) return;
  const datasetId = datasetSelect.value;
  datasetSelect.disabled = true;
  try {
    const response = await fetch("/api/select-dataset", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ session_id: state.sessionId, dataset_id: datasetId }),
    });
    const payload = await response.json();
    if (!response.ok) throw new Error(payload.detail || t("previewFailed"));
    state.activeDatasetId = payload.active_dataset_id;
    applyActiveDataset(payload);
    renderDatasetPanel();
    showToast(t("rowsLoaded", { count: payload.row_count }));
  } catch (error) {
    showToast(error.message || String(error));
  } finally {
    datasetSelect.disabled = false;
  }
});

function applyActiveDataset(payload) {
  state.columns = payload.columns || [];
  state.rowCount = payload.row_count || 0;
  state.activeDatasetId = payload.active_dataset_id || state.activeDatasetId;
  derivedList.innerHTML = "";
  renderColumns();
  populateSourceSelects();
  ensureDefaultNextClass();
  builder.classList.remove("hidden");
  renderDatasetInfo();
  renderPreviewDebounced();
}

function canonicalLabel(column) {
  const base = column.key.replace(/_\d+$/, "");
  const mapping = {
    row_number: "rowNumber",
    student_name: "studentName",
    first_name: "firstName",
    last_name: "lastName",
    class: "class",
    promoted_class: "promotedClass",
    grade: "grade",
    age: "age",
    gender: "gender",
    birth_date: "birthDate",
    email: "email",
    phone: "phone",
    address: "address",
    parent_name: "parentName",
    parent_phone: "parentPhone",
    parent_email: "parentEmail",
    pinfl: "pinfl",
    total: "total",
    percentage: "percentage",
    subject: "subject",
    teacher: "teacher",
    place: "place",
    score: "score",
  };
  return mapping[base] ? t(mapping[base]) : column.label || column.originals?.[0] || t("untitled");
}

function methodLabel(method) {
  const mapping = {
    alias: "methodAlias",
    token: "methodToken",
    fuzzy: "methodFuzzy",
    custom: "methodCustom",
    inferred: "methodInferred",
    question: "methodQuestion",
  };
  return t(mapping[method] || "methodCustom");
}

function renderColumns() {
  if (!columnList || !state.columns.length) {
    if (columnList) columnList.innerHTML = "";
    return;
  }
  columnList.innerHTML = state.columns.map((column) => {
    const selected = Boolean(column.default_selected);
    const sourceNames = (column.originals || []).join(" / ") || column.label;
    return `<article class="column-card ${selected ? "selected" : ""}">
      <input class="column-check" type="checkbox" data-key="${escapeHtml(column.key)}" ${selected ? "checked" : ""}>
      <div class="column-main">
        <div class="column-original">
          <span>
            <strong>${escapeHtml(sourceNames)}</strong>
            <small>${escapeHtml(methodLabel(column.method))} · ${Math.round((column.confidence || 0) * 100)}%</small>
          </span>
          <span class="move-actions">
            <button class="move-button" data-move="up" type="button" title="${escapeHtml(t("moveUp"))}">↑</button>
            <button class="move-button" data-move="down" type="button" title="${escapeHtml(t("moveDown"))}">↓</button>
          </span>
        </div>
        <label>
          <input class="column-name" value="${escapeHtml(canonicalLabel(column))}" aria-label="${escapeHtml(t("outputName"))}">
          ${column.sensitive ? `<span class="sensitive-badge">${escapeHtml(t("sensitiveField"))}</span>` : ""}
        </label>
      </div>
    </article>`;
  }).join("");

  columnList.querySelectorAll(".column-check").forEach((checkbox) => {
    checkbox.addEventListener("change", () => {
      checkbox.closest(".column-card").classList.toggle("selected", checkbox.checked);
      renderPreviewDebounced();
    });
  });
  columnList.querySelectorAll(".column-name").forEach((input) => input.addEventListener("input", renderPreviewDebounced));
  columnList.querySelectorAll(".move-button").forEach((button) => {
    button.addEventListener("click", () => {
      const card = button.closest(".column-card");
      if (button.dataset.move === "up" && card.previousElementSibling) columnList.insertBefore(card, card.previousElementSibling);
      if (button.dataset.move === "down" && card.nextElementSibling) columnList.insertBefore(card.nextElementSibling, card);
      renderPreviewDebounced();
    });
  });
}

function getSelectedColumns() {
  return [...columnList.querySelectorAll(".column-card")]
    .filter((card) => card.querySelector(".column-check").checked)
    .map((card) => ({
      key: card.querySelector(".column-check").dataset.key,
      name: card.querySelector(".column-name").value.trim() || t("untitled"),
    }));
}

function getFixedColumns() {
  return [...fixedList.querySelectorAll(".fixed-row")]
    .map((row) => ({
      name: row.querySelector(".fixed-name").value.trim(),
      value: row.querySelector(".fixed-value").value,
    }))
    .filter((item) => item.name);
}

function sourceOptions(selected = "") {
  return state.columns.map((column) => `<option value="${escapeHtml(column.key)}" ${column.key === selected ? "selected" : ""}>${escapeHtml(canonicalLabel(column))}</option>`).join("");
}

function kindOptions(kind) {
  return [
    ["next_class", "kindNextClass"],
    ["sequence", "kindSequence"],
    ["add_number", "kindAddNumber"],
    ["copy", "kindCopy"],
  ].map(([value, key]) => `<option value="${value}" ${value === kind ? "selected" : ""}>${escapeHtml(t(key))}</option>`).join("");
}

function createDerivedRow(kind = "sequence", sourceKey = "", name = "") {
  const defaultName = name || (kind === "next_class" ? t("nextClass") : kind === "sequence" ? t("rowNumber") : t("calculatedColumn"));
  const row = document.createElement("div");
  row.className = "derived-row";
  row.dataset.autoName = "true";
  row.innerHTML = `
    <input class="derived-name" value="${escapeHtml(defaultName)}" aria-label="${escapeHtml(t("outputName"))}">
    <select class="derived-kind" aria-label="${escapeHtml(t("rule"))}">${kindOptions(kind)}</select>
    <select class="derived-source" aria-label="${escapeHtml(t("sourceColumn"))}">${sourceOptions(sourceKey)}</select>
    <label class="derived-option stop-option"><span>${escapeHtml(t("stopGrades"))}</span><input class="derived-stops" value="9,11"></label>
    <label class="derived-option start-option"><span>${escapeHtml(t("startNumber"))}</span><input class="derived-start" type="number" value="1"></label>
    <label class="derived-option amount-option"><span>${escapeHtml(t("amount"))}</span><input class="derived-amount" type="number" step="any" value="1"></label>
    <button class="remove" type="button" aria-label="${escapeHtml(t("remove"))}">×</button>`;
  derivedList.appendChild(row);
  updateDerivedVisibility(row);
  bindDerivedRow(row);
  renderPreviewDebounced();
  return row;
}

function bindDerivedRow(row) {
  row.querySelector(".derived-kind").addEventListener("change", () => {
    row.dataset.autoName = "true";
    const kind = row.querySelector(".derived-kind").value;
    row.querySelector(".derived-name").value = kind === "next_class" ? t("nextClass") : kind === "sequence" ? t("rowNumber") : t("calculatedColumn");
    updateDerivedVisibility(row);
    renderPreviewDebounced();
  });
  row.querySelector(".derived-name").addEventListener("input", () => {
    row.dataset.autoName = "false";
    renderPreviewDebounced();
  });
  row.querySelectorAll("select,input").forEach((control) => control.addEventListener("change", renderPreviewDebounced));
  row.querySelector(".remove").addEventListener("click", () => {
    row.remove();
    renderPreviewDebounced();
  });
}

function updateDerivedVisibility(row) {
  const kind = row.querySelector(".derived-kind").value;
  row.querySelector(".derived-source").classList.toggle("hidden-control", kind === "sequence");
  row.querySelector(".stop-option").classList.toggle("hidden-control", kind !== "next_class");
  row.querySelector(".start-option").classList.toggle("hidden-control", kind !== "sequence");
  row.querySelector(".amount-option").classList.toggle("hidden-control", kind !== "add_number");
}

function renderDerivedRows(updateNames = true) {
  [...derivedList.querySelectorAll(".derived-row")].forEach((row) => {
    const kind = row.querySelector(".derived-kind").value;
    row.querySelector(".derived-kind").innerHTML = kindOptions(kind);
    const source = row.querySelector(".derived-source").value;
    row.querySelector(".derived-source").innerHTML = sourceOptions(source);
    row.querySelector(".stop-option span").textContent = t("stopGrades");
    row.querySelector(".start-option span").textContent = t("startNumber");
    row.querySelector(".amount-option span").textContent = t("amount");
    if (updateNames && row.dataset.autoName === "true") {
      row.querySelector(".derived-name").value = kind === "next_class" ? t("nextClass") : kind === "sequence" ? t("rowNumber") : t("calculatedColumn");
    }
    updateDerivedVisibility(row);
  });
}

function getDerivedColumns() {
  return [...derivedList.querySelectorAll(".derived-row")].map((row) => ({
    name: row.querySelector(".derived-name").value.trim() || t("untitled"),
    kind: row.querySelector(".derived-kind").value,
    source_key: row.querySelector(".derived-source").value || null,
    stop_grades: row.querySelector(".derived-stops").value.split(",").map((value) => Number(value.trim())).filter(Number.isInteger),
    amount: Number(row.querySelector(".derived-amount").value) || 0,
    start: Number.parseInt(row.querySelector(".derived-start").value, 10) || 1,
  }));
}

function populateSourceSelects() {
  if (!skipBlankKey || !sortKey) return;
  const blank = `<option value="">${escapeHtml(t("disabled"))}</option>`;
  skipBlankKey.innerHTML = blank + sourceOptions();
  sortKey.innerHTML = `<option value="">${escapeHtml(t("originalOrder"))}</option>` + sourceOptions();
  renderDerivedRows(false);
}

function ensureDefaultNextClass() {
  if (derivedList.children.length) return;
  const classColumn = state.columns.find((column) => column.key === "class");
  const alreadyPromoted = state.columns.some((column) => column.key === "promoted_class");
  if (classColumn && !alreadyPromoted) createDerivedRow("next_class", classColumn.key, t("nextClass"));
}

function getOptions() {
  return {
    trim_whitespace: $("trimWhitespace").checked,
    remove_duplicates: $("removeDuplicates").checked,
    skip_blank_key: skipBlankKey.value || null,
    sort_key: sortKey.value || null,
  };
}

function buildPayload() {
  return {
    session_id: state.sessionId,
    columns: getSelectedColumns(),
    fixed_columns: getFixedColumns(),
    derived_columns: getDerivedColumns(),
    options: getOptions(),
  };
}

function renderPreviewDebounced() {
  if (!state.sessionId || !state.activeDatasetId) return;
  clearTimeout(state.previewTimer);
  previewNote.textContent = t("previewLoading");
  state.previewTimer = setTimeout(updatePreview, 180);
}

async function updatePreview() {
  const columns = getSelectedColumns();
  if (!columns.length) {
    previewHead.innerHTML = "";
    previewBody.innerHTML = `<tr><td>${escapeHtml(t("selectColumn"))}</td></tr>`;
    previewNote.textContent = "";
    return;
  }
  try {
    const response = await fetch("/api/preview", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ ...buildPayload(), limit: 50 }),
    });
    const payload = await response.json();
    if (!response.ok) throw new Error(payload.detail || t("previewFailed"));
    rowCount.textContent = t("rows", { count: payload.row_count });
    previewNote.textContent = t("previewShowing", { shown: payload.preview_count, count: payload.row_count });
    previewHead.innerHTML = payload.headers.length ? `<tr>${payload.headers.map((header) => `<th>${escapeHtml(header)}</th>`).join("")}</tr>` : "";
    previewBody.innerHTML = payload.rows.length
      ? payload.rows.map((row) => `<tr>${payload.headers.map((header) => `<td>${escapeHtml(row[header] ?? "")}</td>`).join("")}</tr>`).join("")
      : `<tr><td>${escapeHtml(t("selectColumn"))}</td></tr>`;
  } catch (error) {
    previewNote.textContent = error.message || t("previewFailed");
  }
}

$("toggleAll").addEventListener("click", () => {
  const boxes = [...columnList.querySelectorAll(".column-check")];
  const next = boxes.some((box) => !box.checked);
  boxes.forEach((box) => {
    box.checked = next;
    box.closest(".column-card").classList.toggle("selected", next);
  });
  renderPreviewDebounced();
});

$("addFixed").addEventListener("click", () => {
  fixedList.insertAdjacentHTML("beforeend", `<div class="fixed-row"><input class="fixed-name" value="${escapeHtml(t("newColumn"))}" aria-label="${escapeHtml(t("fixedName"))}"><input class="fixed-value" value="" aria-label="${escapeHtml(t("fixedValue"))}"><button class="remove" type="button" aria-label="${escapeHtml(t("remove"))}">×</button></div>`);
  renderPreviewDebounced();
});

$("addNextClass").addEventListener("click", () => {
  const classColumn = state.columns.find((column) => column.key === "class") || state.columns[0];
  createDerivedRow("next_class", classColumn?.key || "", t("nextClass"));
});

$("addDerived").addEventListener("click", () => createDerivedRow("sequence", "", t("rowNumber")));

fixedList.addEventListener("input", (event) => {
  if (event.target.dataset.defaultKey) event.target.dataset.auto = "false";
  renderPreviewDebounced();
});

fixedList.addEventListener("click", (event) => {
  if (event.target.classList.contains("remove")) {
    event.target.closest(".fixed-row").remove();
    renderPreviewDebounced();
  }
});

[$("trimWhitespace"), $("removeDuplicates"), skipBlankKey, sortKey].forEach((control) => control.addEventListener("change", renderPreviewDebounced));

document.querySelectorAll(".export").forEach((button) => button.addEventListener("click", () => downloadExport(button.dataset.format)));

async function downloadExport(format) {
  if (!state.sessionId || !state.activeDatasetId) {
    showToast(t("uploadFirst"));
    return;
  }
  const columns = getSelectedColumns();
  if (!columns.length) {
    showToast(t("selectColumn"));
    return;
  }
  const response = await fetch("/api/export", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ ...buildPayload(), format }),
  });
  if (!response.ok) {
    const payload = await response.json();
    showToast(typeof payload.detail === "string" ? payload.detail : JSON.stringify(payload.detail));
    return;
  }
  const blob = await response.blob();
  const disposition = response.headers.get("content-disposition") || "";
  const match = disposition.match(/filename="?([^";]+)"?/);
  const filename = match ? match[1] : `FormaFlow_Output.${format}`;
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
  showToast(t("downloadStarted", { format: format.toUpperCase() }));
}

$("clearButton").addEventListener("click", () => {
  state.sessionId = null;
  state.columns = [];
  state.files = [];
  state.datasetGroups = [];
  state.activeDatasetId = null;
  state.rowCount = 0;
  fileInput.value = "";
  builder.classList.add("hidden");
  datasetPanel.classList.add("hidden");
  fileResults.classList.add("hidden");
  derivedList.innerHTML = "";
  hideMessage();
  showToast(t("cleared"));
});

applyTranslations();
checkServer();
