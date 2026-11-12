// JAI Lorebook Converter for SillyTavern (v0.2.1)
//
// Converts JanitorAI lorebooks and scripts to SillyTavern World Info during import.
// Supported inputs (same as the web converter): JAI lorebook JSON, loreEntries script,
// Advanced Lore script, if-chain script, and any of those pasted into a .docx.
// Conversion uses the web converter's own code and its "Simple mode" defaults.
//
// How it hooks in: SillyTavern's backend only checks `'entries' in json`, which is also true
// for any JSON array, so a JAI lorebook gets saved as-is and shows up empty. A capture-phase
// listener on the World Info import input runs before ST's own handler, looks at the file,
// and either hands it back to ST untouched (anything ST can read) or offers to convert it.
//
// Safety: JSON files and if-chain scripts are only parsed as text. loreEntries and Advanced Lore
// scripts are read by running the script's array/object literals with `new Function` (same as the
// web converter), which executes file text inside SillyTavern's page. That is why those two
// formats sit behind the "Advanced scripts" toggle, which is off by default.
//
// Only the public SillyTavern.getContext() API is used (no imports from ST internals).

const MODULE_NAME = 'jai_lorebook_converter';
const DEFAULT_SETTINGS = Object.freeze({ enabled: true, advancedScripts: false });
const IMPORT_INPUT_ID = 'world_import_file';
const EXTRA_ACCEPT = ['.txt', '.js', '.docx'];
// Formats that are read by executing the file's text as JavaScript.
const OPT_IN_FORMATS = new Set(['jai_script', 'dynamic_lore']);

// ---------------------------------------------------------------------------
// Strings (en / ru). Russian is used when ST's locale starts with "ru".
// The converter strings below are copied from the web converter's lang files.
// ---------------------------------------------------------------------------
const ruPlural = (n, one, few, many) => {
    const mod10 = n % 10, mod100 = n % 100;
    if (mod10 === 1 && mod100 !== 11) return one;
    if (mod10 >= 2 && mod10 <= 4 && (mod100 < 10 || mod100 >= 20)) return few;
    return many;
};

const CONVERTER_STRINGS_EN = {
    keysSplit: (n, k) => "<strong>" + n + " " + (n === 1 ? "entry" : "entries") + "</strong> had keywords packed into a single item (like <code>/a/b/c/</code> or <code>a, b, c</code>) — split into proper separate keywords (" + k + " total).",
    badKeywords: (n, d) => "<strong>" + n + " " + (n === 1 ? "entry" : "entries") + "</strong> had a keyword field that looked like pasted content, not trigger words (" + d + " item" + (d === 1 ? "" : "s") + ", too long or containing newlines) — moved into the entry's content and the keyword field cleared for you to fill in.",
    jaiPosSet: "Position set to <strong>Before Char (0)</strong> for all entries. Adjust in ST as needed.",
    scriptSplit: "Each script entry generates up to <strong>2 ST entries</strong> (scenario + personality). Adjust positions in the panel if needed.",
    scriptRepaired: (s) => {
    const parts = [];
    if (s.unclosedStrings) parts.push(s.unclosedStrings + " unclosed string" + (s.unclosedStrings === 1 ? "" : "s"));
    if (s.missingCommas) parts.push(s.missingCommas + " missing comma" + (s.missingCommas === 1 ? "" : "s") + " between list items");
    if (s.prematureCloses) parts.push(s.prematureCloses + " array" + (s.prematureCloses === 1 ? "" : "s") + " that got closed early, orphaning entries after it");
    return "<strong>Auto-repaired " + parts.length + " syntax issue" + (parts.length === 1 ? "" : "s") + "</strong> in your script so it could still parse: " + parts.join(", ") + ". Double-check the affected entries came through correctly.";
  },
    outletReminder: "For any entries mapped to <strong>7 / Outlet</strong>, make sure to add <code>{{outlet::name}}</code> to your character card.",
    filtersUsed: (n) => "<strong>" + n + " entries</strong> use filters — mapped to ST secondary keys.",
    ifChainDup: "If-chain scripts may have duplicate keyword sets. They have been grouped under the <strong>If-Chain</strong> category.",
    dynLoreInlined: (n) => "<strong>" + n + " trigger-only entr" + (n === 1 ? "y" : "ies") + "</strong> (tag-based, no keywords) " + (n === 1 ? "was" : "were") + " merged into whichever entry emits its tag — ST has no cross-entry tag system, so if two entries share a trigger tag, that content may now appear twice in the same turn instead of once.",
    dynLoreShifts: (n) => "<strong>" + n + " Shift" + (n === 1 ? "" : "s") + "</strong> became separate ST entries gated on the parent's keywords via AND ALL — an approximation of \"only after the parent fires.\"",
    dynLoreGates: (n) => "<strong>" + n + " entr" + (n === 1 ? "y" : "ies") + "</strong> used gates ST can't express (tag-gates, emotion detection, maxMessages, notAll, or prev.* targeting) — those conditions were dropped; keyword/priority logic still applies.",
    dynLoreNameBlock: (n) => "<strong>" + n + " nameBlock entr" + (n === 1 ? "y" : "ies") + "</strong> mapped to ST's Character Filter (exclude).",
    dynLoreRelationships: (n) => "<strong>" + n + " relationship injection" + (n === 1 ? "" : "s") + "</strong> converted to entries requiring both characters' names (AND ALL) — the original requireTags condition couldn't be preserved.",
    scriptParseFailed: (msg) => "Script parse failed: " + msg,
    ifChainParseFailed: (msg) => "if-chain parse failed: " + msg,
    noEntriesLoreEntries: () => "No entries found in loreEntries.",
    noKeywordEntries: () => "No keyword-based entries found. Check the script uses lastMessage.includes(…) and context.character.personality/scenario +=",
    docxReadFailed: (msg) => "DOCX Read Failed: " + msg,
};

const CONVERTER_STRINGS_RU = {
    keysSplit: (n, k) => "<strong>" + n + " " + ruPlural(n, "запись", "записи", "записей") + "</strong> содержали ключевые слова, склеенные в один элемент (например, <code>/a/b/c/</code> или <code>a, b, c</code>) — разделены на отдельные ключевые слова (всего " + k + ").",
    badKeywords: (n, d) => "<strong>" + n + " " + ruPlural(n, "запись", "записи", "записей") + "</strong> содержали поле ключевых слов, похожее на вставленный текст, а не на триггерные слова (" + d + " " + ruPlural(d, "элемент", "элемента", "элементов") + ", слишком длинные или с переносами строки) — перенесено в поле содержимого записи, а поле ключевых слов очищено для ручного заполнения.",
    jaiPosSet: "Для всех записей выставлена позиция <strong>Перед описанием персонажа (0)</strong>. При необходимости измените её в ST.",
    scriptSplit: "Каждая запись скрипта создаёт до <strong>2 записей ST</strong> (scenario + personality). При необходимости измените позиции в панели выше.",
    scriptRepaired: (s) => {
    const parts = [];
    if (s.unclosedStrings) parts.push(s.unclosedStrings + " " + ruPlural(s.unclosedStrings, "незакрытая строка", "незакрытые строки", "незакрытых строк"));
    if (s.missingCommas) parts.push(s.missingCommas + " " + ruPlural(s.missingCommas, "пропущенная запятая", "пропущенные запятые", "пропущенных запятых") + " между элементами списка");
    if (s.prematureCloses) parts.push(s.prematureCloses + " " + ruPlural(s.prematureCloses, "преждевременно закрытый массив", "преждевременно закрытых массива", "преждевременно закрытых массивов") + ", из-за чего часть записей осталась «снаружи»");
    return "<strong>Автоматически исправлено " + parts.length + " " + ruPlural(parts.length, "синтаксическая ошибка", "синтаксические ошибки", "синтаксических ошибок") + "</strong> в скрипте, чтобы его вообще можно было разобрать: " + parts.join(", ") + ". Проверьте, что затронутые записи собрались верно.";
  },
    outletReminder: "Для всех записей, привязанных к позиции <strong>7 / Outlet</strong>, добавьте <code>{{outlet::name}}</code> в карточку персонажа.",
    filtersUsed: (n) => "<strong>" + n + " " + ruPlural(n, "запись", "записи", "записей") + "</strong> используют фильтры — они сопоставлены с дополнительными ключами (secondary) в ST.",
    ifChainDup: "Скрипты if-chain могут содержать повторяющиеся наборы ключевых слов. Они сгруппированы в категорию <strong>If-Chain</strong>.",
    dynLoreInlined: (n) => "<strong>" + n + " " + ruPlural(n, "запись-триггер", "записи-триггера", "записей-триггеров") + "</strong> (по тегу, без ключевых слов) объединены с записью, которая испускает этот тег — в ST нет межзаписевой системы тегов, поэтому если тег испускают две записи, этот текст может появиться дважды за один ход вместо одного раза.",
    dynLoreShifts: (n) => "<strong>" + n + " " + ruPlural(n, "Shift", "Shift'а", "Shift'ов") + "</strong> стали отдельными записями ST с условием AND ALL по ключевым словам родителя — приближение к «только после срабатывания родителя».",
    dynLoreGates: (n) => "<strong>" + n + " " + ruPlural(n, "запись", "записи", "записей") + "</strong> использовали условия, которые ST не может выразить (теги, эмоции, maxMessages, notAll или таргетинг prev.*) — эти условия были отброшены; логика ключевых слов и приоритета сохраняется.",
    dynLoreNameBlock: (n) => "<strong>" + n + " " + ruPlural(n, "запись nameBlock", "записи nameBlock", "записей nameBlock") + "</strong> сопоставлены с фильтром персонажа ST (исключение).",
    dynLoreRelationships: (n) => "<strong>" + n + " " + ruPlural(n, "триггер отношений", "триггера отношений", "триггеров отношений") + "</strong> преобразованы в записи, требующие оба имени персонажей (AND ALL) — исходное условие requireTags сохранить не удалось.",
    scriptParseFailed: (msg) => "Ошибка разбора скрипта: " + msg,
    ifChainParseFailed: (msg) => "Ошибка разбора if-chain: " + msg,
    noEntriesLoreEntries: () => "В loreEntries не найдено записей.",
    noKeywordEntries: () => "Не найдено записей на основе ключевых слов. Убедитесь, что скрипт использует lastMessage.includes(…) и context.character.personality/scenario +=",
    docxReadFailed: (msg) => "Не удалось прочитать DOCX: " + msg,
};

const UI_STRINGS_EN = {
    title: 'JAI Lorebook Converter',
    enable: 'Offer to convert JanitorAI lorebooks and scripts on import',
    enableScripts: 'Advanced scripts (loreEntries, Advanced Lore)',
    enableScriptsHint: 'These two script types can only be read by running the script text as JavaScript inside SillyTavern. Enable only for files you trust. If-chain scripts are converted without running anything.',
    popupTitle: 'Convert to SillyTavern format?',
    popupBody: (name, fmt, src, st) => 'The file "' + name + '" is a ' + fmt + ' (' + src + ' source entr' + (src === 1 ? 'y' : 'ies') + ' \u2192 ' + st + ' SillyTavern entr' + (st === 1 ? 'y' : 'ies') + '). SillyTavern cannot read this format directly.',
    convert: 'Convert & import',
    cancel: 'Cancel',
    convertFailed: 'Could not convert this file. Nothing was imported.',
    scriptsDisabled: (name, fmt) => 'The file "' + name + '" is a ' + fmt + '. Turn on "Advanced scripts" in the JAI Lorebook Converter settings to convert it.',
    formats: { jai_json: 'JanitorAI lorebook (JSON)', jai_script: 'JanitorAI loreEntries script', dynamic_lore: 'JanitorAI Advanced Lore script', if_chain: 'JanitorAI if-chain script' },
};

const UI_STRINGS_RU = {
    title: 'Конвертер лорбуков JAI',
    enable: 'Предлагать конвертацию лорбуков и скриптов JanitorAI при импорте',
    enableScripts: 'Продвинутые скрипты (loreEntries, Advanced Lore)',
    enableScriptsHint: 'Эти два типа скриптов можно прочитать только запустив текст скрипта как JavaScript внутри SillyTavern. Включайте только для файлов, которым доверяете. Скрипты if-chain конвертируются без запуска кода.',
    popupTitle: 'Конвертировать в формат SillyTavern?',
    popupBody: (name, fmt, src, st) => 'Файл «' + name + '» — ' + fmt + ' (' + src + ' ' + ruPlural(src, 'исходная запись', 'исходные записи', 'исходных записей') + ' → ' + st + ' ' + ruPlural(st, 'запись', 'записи', 'записей') + ' SillyTavern). SillyTavern не умеет читать этот формат напрямую.',
    convert: 'Конвертировать и импортировать',
    cancel: 'Отмена',
    convertFailed: 'Не удалось сконвертировать файл. Ничего не импортировано.',
    scriptsDisabled: (name, fmt) => 'Файл «' + name + '» — ' + fmt + '. Включите «Продвинутые скрипты» в настройках конвертера JAI, чтобы его конвертировать.',
    formats: { jai_json: 'лорбук JanitorAI (JSON)', jai_script: 'скрипт JanitorAI (loreEntries)', dynamic_lore: 'скрипт JanitorAI (Advanced Lore)', if_chain: 'скрипт JanitorAI (if-chain)' },
};

const STRINGS = {
    en: { ...CONVERTER_STRINGS_EN, ...UI_STRINGS_EN },
    ru: { ...CONVERTER_STRINGS_RU, ...UI_STRINGS_RU },
};

function getStrings() {
    const locale = String(SillyTavern.getContext().getCurrentLocale?.() || 'en').toLowerCase();
    return locale.startsWith('ru') ? STRINGS.ru : STRINGS.en;
}

// ---------------------------------------------------------------------------
// Settings
// ---------------------------------------------------------------------------
function getSettings() {
    const { extensionSettings } = SillyTavern.getContext();
    if (!extensionSettings[MODULE_NAME]) extensionSettings[MODULE_NAME] = {};
    const s = extensionSettings[MODULE_NAME];
    for (const [key, value] of Object.entries(DEFAULT_SETTINGS)) {
        if (!Object.hasOwn(s, key)) s[key] = value;
    }
    return s;
}

// ---------------------------------------------------------------------------
// Converter code, copied unchanged from the web converter's script.js
// (module-level `last*Stats` variables are reset at the start of every conversion).
// ---------------------------------------------------------------------------
let lastScriptRepairStats = null;

// --- Keyword sanitization ---
// If a keyword field is unnaturally big (any item too long or containing a
// newline), the whole field is treated as corrupted — not item by item. The
// entire thing is moved into the entry's content (nothing discarded), `key`
// is left empty, and the entry is flagged so the fix-it popup lets the user
// type real trigger words in by hand.
let lastBadKeyStats = null;

const KEY_MAX_LEN = 80;

// Turns messy keyword lists into proper ones:
//  - "/a/b/c/"        (slash-delimited, leading slash)  -> a, b, c
//  - "a, b, c"        (comma-joined into one item)      -> a, b, c
// Trims, drops empties and removes duplicates (case-insensitive). Items with newlines are left
// alone so the "pasted content" check below can still catch them.
let lastKeySplitStats = null;

let lastDynLoreStats = null;

function skipStringOrComment(src, i) {
  const c = src[i];
  if (c === "/" && src[i + 1] === "/") {
    const nl = src.indexOf("\n", i);
    return nl === -1 ? src.length : nl + 1;
  }
  if (c === "/" && src[i + 1] === "*") {
    const end = src.indexOf("*/", i + 2);
    return end === -1 ? src.length : end + 2;
  }
  if (c === '"' || c === "'" || c === "`") {
    let j = i + 1;
    while (j < src.length) {
      if (src[j] === "\\") { j += 2; continue; }
      if (src[j] === c) return j + 1;
      j++;
    }
    return j;
  }
  return i;
}

function detectFormat(text) {
  const t = text.trim();
  if (t.startsWith("{") || t.startsWith("[")) {
    try {
      const p = JSON.parse(t);
      if (Array.isArray(p) && p.length && ("keys" in p[0] || "keyword" in p[0] || "keywords" in p[0] || "content" in p[0]))
        return { type: "jai_json", parsed: p };
      if (p && p.entries && typeof p.entries === "object" && !Array.isArray(p.entries))
        return { type: "st_json", parsed: p };
      const ks = Object.keys(p);
      if (ks.length && ks.every((k) => !isNaN(parseInt(k))))
        return { type: "st_json", parsed: { entries: p } };
    } catch (e) {}
  }
  if (/loreEntries\s*=\s*\[/.test(t)) return { type: "jai_script", parsed: null };
  if (detectDynamicLore(t)) return { type: "dynamic_lore", parsed: null };
  const msgVar = detectMsgVar(t);
  const msgVarRe = new RegExp(msgVar.replace(/[.*+?^${}()|[\]\\]/g, "\\$&") + "\\.includes\\s*\\(");
  if (msgVarRe.test(t) && /context\.character\.(personality|scenario)\s*(\+=|\.replace\s*\()/.test(t))
    return { type: "if_chain", parsed: null };
  return null;
}

async function extractDocxText(arrayBuffer) {
  const view = new DataView(arrayBuffer);
  const bytes = new Uint8Array(arrayBuffer);
  let eocdOffset = -1;
  for (let i = bytes.length - 22; i >= 0; i--) {
    if (view.getUint32(i, true) === 0x06054b50) { eocdOffset = i; break; }
  }
  if (eocdOffset === -1) throw new Error("Not a valid ZIP/DOCX file");
  const cdOffset = view.getUint32(eocdOffset + 16, true);
  const cdEntries = view.getUint16(eocdOffset + 10, true);
  let offset = cdOffset, targetOffset = -1, targetCompMethod = -1, targetCompSize = -1;
  for (let i = 0; i < cdEntries; i++) {
    if (view.getUint32(offset, true) !== 0x02014b50) break;
    const compMethod = view.getUint16(offset + 10, true);
    const compSize = view.getUint32(offset + 20, true);
    const nameLen = view.getUint16(offset + 28, true);
    const extraLen = view.getUint16(offset + 30, true);
    const commentLen = view.getUint16(offset + 32, true);
    const localHeaderOffset = view.getUint32(offset + 42, true);
    const nameBytes = bytes.subarray(offset + 46, offset + 46 + nameLen);
    const name = new TextDecoder().decode(nameBytes);
    if (name === "word/document.xml") {
      targetOffset = localHeaderOffset; targetCompMethod = compMethod; targetCompSize = compSize; break;
    }
    offset += 46 + nameLen + extraLen + commentLen;
  }
  if (targetOffset === -1) throw new Error("word/document.xml not found in DOCX");
  if (view.getUint32(targetOffset, true) !== 0x04034b50) throw new Error("Invalid local file header");
  const lfNameLen = view.getUint16(targetOffset + 26, true);
  const lfExtraLen = view.getUint16(targetOffset + 28, true);
  const dataOffset = targetOffset + 30 + lfNameLen + lfExtraLen;
  const compressedData = bytes.subarray(dataOffset, dataOffset + targetCompSize);

  let xmlString = "";
  if (targetCompMethod === 0) {
    xmlString = new TextDecoder().decode(compressedData);
  } else if (targetCompMethod === 8) {
    if (typeof DecompressionStream === "undefined") throw new Error("Browser doesn't support local DOCX extraction. Try copy-pasting instead.");
    const ds = new DecompressionStream("deflate-raw");
    const writer = ds.writable.getWriter();
    writer.write(compressedData); writer.close();
    const reader = ds.readable.getReader();
    const chunks = [];
    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      chunks.push(value);
    }
    const totalLen = chunks.reduce((acc, c) => acc + c.length, 0);
    const decompressed = new Uint8Array(totalLen);
    let ptr = 0;
    for (const c of chunks) { decompressed.set(c, ptr); ptr += c.length; }
    xmlString = new TextDecoder().decode(decompressed);
  } else {
    throw new Error("Unsupported DOCX compression method");
  }

  let paragraphs = xmlString.split("</w:p>");
  let text = paragraphs.map((p) => {
    let cleanP = p.replace(/<w:br[^>]*>/g, "<w:t>\n</w:t>");
    let tMatch = cleanP.match(/<w:t[^>]*>([\s\S]*?)<\/w:t>/g);
    if (!tMatch) return "";
    let line = tMatch.map((t) => t.replace(/<w:t[^>]*>/, "").replace(/<\/w:t>/, "")).join("");
    return line.replace(/&amp;/g, "&").replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&apos;/g, "'");
  }).filter((p) => p.trim() !== "").join("\n\n");
  return text.replace(/[\u2018\u2019]/g, "'").replace(/[\u201C\u201D]/g, '"');
}

function sanitizeQuotes(code) {
  let out = "", i = 0;
  while (i < code.length) {
    if (code[i] === "/" && code[i + 1] === "/") {
      const end = code.indexOf("\n", i);
      const e = end === -1 ? code.length : end + 1;
      out += code.slice(i, e); i = e; continue;
    }
    if (code[i] === "/" && code[i + 1] === "*") {
      const end = code.indexOf("*/", i + 2);
      const e = end === -1 ? code.length : end + 2;
      out += code.slice(i, e); i = e; continue;
    }
    if (code[i] === "`") {
      let j = i + 1;
      while (j < code.length) {
        if (code[j] === "\\") { j += 2; continue; }
        if (code[j] === "`") { j++; break; }
        j++;
      }
      out += code.slice(i, j); i = j; continue;
    }
    if (code[i] === '"') {
      let j = i + 1;
      while (j < code.length) {
        if (code[j] === "\\") { j += 2; continue; }
        if (code[j] === '"') { j++; break; }
        j++;
      }
      out += code.slice(i, j); i = j; continue;
    }
    if (code[i] === "'") {
      let j = i + 1, inner = "";
      while (j < code.length) {
        if (code[j] === "\\" && j + 1 < code.length) {
          const nx = code[j + 1];
          if (nx === "'") inner += "'";
          else if (nx === '"') inner += '\\"';
          else inner += code[j] + nx;
          j += 2; continue;
        }
        if (code[j] === "'") { j++; break; }
        if (code[j] === '"') inner += '\\"';
        else inner += code[j];
        j++;
      }
      out += '"' + inner + '"'; i = j; continue;
    }
    out += code[i]; i++;
  }
  return out;
}

function isBadKeyword(s) {
  if (typeof s !== "string") return true;
  if (s.includes("\n")) return true;
  if (s.length > KEY_MAX_LEN) return true;
  return false;
}

function normalizeKeyList(arr) {
  if (!Array.isArray(arr)) return arr;
  const out = [];
  const seen = new Set();
  let split = false;
  const add = (t) => {
    const low = typeof t === "string" ? t.toLowerCase() : t;
    if (typeof t === "string" && seen.has(low)) return;
    if (typeof t === "string") seen.add(low);
    out.push(t);
  };
  for (const k of arr) {
    if (typeof k !== "string" || k.includes("\n")) { add(k); continue; }
    const t = k.trim();
    let parts = null;
    if (t.startsWith("/") && t.indexOf("/", 1) !== -1) parts = t.split("/");
    else if (t.includes(",")) parts = t.split(",");
    if (parts) {
      split = true;
      for (const piece of parts) { const x = piece.trim(); if (x) add(x); }
    } else if (t) add(t);
  }
  if (split) {
    if (!lastKeySplitStats) lastKeySplitStats = { entries: 0, keys: 0 };
    lastKeySplitStats.entries++;
    lastKeySplitStats.keys += out.length;
  }
  return out;
}

function sanitizeKeyList(arr) {
  if (!Array.isArray(arr) || !arr.length) return { keys: arr, movedText: "" };
  const expanded = normalizeKeyList(arr);
  if (!expanded.some(isBadKeyword)) return { keys: expanded, movedText: "" };
  if (!lastBadKeyStats) lastBadKeyStats = { entries: 0, dropped: 0 };
  lastBadKeyStats.entries++;
  lastBadKeyStats.dropped += arr.length;
  return { keys: [], movedText: arr.join(", ") };
}

function fixUnclosedLineStrings(code) {
  const lines = code.split("\n");
  let fixed = 0;
  for (let li = 0; li < lines.length; li++) {
    let line = lines[li];
    const m = line.match(/^(\s*[A-Za-z_$][\w$]*\s*:\s*)(['"])/);
    if (!m) continue;
    const q = m[2];
    const rest = line.slice(m[0].length);
    let closed = false, esc = false;
    for (const ch of rest) {
      if (esc) { esc = false; continue; }
      if (ch === "\\") { esc = true; continue; }
      if (ch === q) { closed = true; break; }
    }
    if (!closed) {
      if (/\[/.test(rest) && !/\]/.test(rest)) continue;
      const trailMatch = line.match(/,\s*$/);
      if (trailMatch) line = line.slice(0, line.length - trailMatch[0].length) + q + ",";
      else line = line.replace(/\s*$/, "") + q;
      lines[li] = line; fixed++;
    }
  }
  return { code: lines.join("\n"), fixed };
}

function fixMissingCommasInLists(code) {
  let fixed = 0;
  const out = code.replace(/\[([^[\]]*)\]/gs, (whole, inner) => {
    const fixedInner = inner.replace(/(['"])(\s+)(?=['"])/g, (m2, q, ws) => { fixed++; return q + "," + ws; });
    return "[" + fixedInner + "]";
  });
  return { code: out, fixed };
}

function repairScriptText(code) {
  const r1 = fixUnclosedLineStrings(code);
  const r2 = fixMissingCommasInLists(r1.code);
  if (!lastScriptRepairStats) lastScriptRepairStats = { unclosedStrings: 0, missingCommas: 0, prematureCloses: 0 };
  lastScriptRepairStats.unclosedStrings += r1.fixed;
  lastScriptRepairStats.missingCommas += r2.fixed;
  return r2.code;
}

function findEntryArrayEnd(code, start) {
  let depth = 0, i = start;
  const falseCloses = [];
  while (i < code.length) {
    const nextI = skipStringOrComment(code, i);
    if (nextI > i) { i = nextI; continue; }
    const c = code[i];
    if (c === "[") { depth++; i++; continue; }
    if (c === "]") {
      depth--;
      if (depth === 0) {
        let j = i + 1;
        while (j < code.length) {
          const skip = skipStringOrComment(code, j);
          if (skip > j) { j = skip; continue; }
          if (/\s/.test(code[j]) || code[j] === ";") { j++; continue; }
          break;
        }
        if (code[j] === "{" && /\b(keywords|tag)\s*:/.test(code.slice(j, j + 400))) {
          falseCloses.push([i, j]); depth = 1; i = j; continue;
        }
        return { end: i, falseCloses };
      }
    }
    i++;
  }
  return { end: -1, falseCloses };
}

function extractRepairedArray(code, start) {
  const { end, falseCloses } = findEntryArrayEnd(code, start);
  if (end === -1) return null;
  let arrStr = code.slice(start, end + 1);
  for (let k = falseCloses.length - 1; k >= 0; k--) {
    const [closeIdx, resumeIdx] = falseCloses[k];
    arrStr = arrStr.slice(0, closeIdx - start) + "," + arrStr.slice(resumeIdx - start);
  }
  if (!lastScriptRepairStats) lastScriptRepairStats = { unclosedStrings: 0, missingCommas: 0, prematureCloses: 0 };
  lastScriptRepairStats.prematureCloses = falseCloses.length;
  return arrStr;
}

function parseScript(code) {
  const startIdx = code.indexOf("loreEntries");
  if (startIdx === -1) throw new Error("loreEntries not found");
  const arrStart = code.indexOf("[", code.indexOf("=", startIdx));
  if (arrStart === -1) throw new Error("Array not found");
  lastScriptRepairStats = null;
  const arrStr = extractRepairedArray(code, arrStart);
  if (arrStr === null) throw new Error("Unclosed bracket");

  try {
    const r = new Function('"use strict";return(' + arrStr + ")")();
    if (Array.isArray(r)) return r;
  } catch (_) {}
  try {
    const r = new Function('"use strict";return(' + sanitizeQuotes(arrStr) + ")")();
    if (Array.isArray(r)) return r;
  } catch (_) {}

  const repairedCode = repairScriptText(code);
  const repairedStart = repairedCode.indexOf("[", repairedCode.indexOf("=", repairedCode.indexOf("loreEntries")));
  const repairedArrStr = repairedStart === -1 ? null : extractRepairedArray(repairedCode, repairedStart);
  if (repairedArrStr === null) throw new Error("Unclosed bracket");
  try {
    const r = new Function('"use strict";return(' + repairedArrStr + ")")();
    if (Array.isArray(r)) return r;
  } catch (_) {}
  try {
    const r = new Function('"use strict";return(' + sanitizeQuotes(repairedArrStr) + ")")();
    if (Array.isArray(r)) return r;
    throw new Error("Result is not an array");
  } catch (e) { throw new Error("Parse failed: " + e.message); }
}

function detectMsgVar(code) {
  const m = code.match(/\b(?:const|let|var)\s+(\w+)\s*=[^;]*context\.chat\.last_messages?\b/);
  return m ? m[1] : "lastMessage";
}

function extractKeywordConstants(code) {
  const map = {};
  const declRe = /\b(?:const|let|var)\s+(\w+)\s*=\s*([(\[])/g;
  let m;
  while ((m = declRe.exec(code)) !== null) {
    const name = m[1];
    const openChar = m[2];
    const closeChar = openChar === "(" ? ")" : "]";
    let i = declRe.lastIndex - 1, depth = 0;
    const bodyStart = i;
    while (i < code.length) {
      const nextI = skipStringOrComment(code, i);
      if (nextI > i) { i = nextI; continue; }
      const c = code[i];
      if (c === openChar) depth++;
      else if (c === closeChar) { depth--; if (depth === 0) { i++; break; } }
      i++;
    }
    const body = code.slice(bodyStart, i);
    const kws = [];
    const strRe = /'([^'\\]*(?:\\.[^'\\]*)*)'|"([^"\\]*(?:\\.[^"\\]*)*)"/g;
    let sm;
    while ((sm = strRe.exec(body)) !== null) kws.push(sm[1] !== undefined ? sm[1] : sm[2]);
    if (kws.length) map[name] = kws;
    declRe.lastIndex = i;
  }
  return map;
}

function parseIfChain(code) {
  const entries = [];
  const msgVar = detectMsgVar(code);
  const kwConsts = extractKeywordConstants(code);
  // name -> string value for `const x = `...`` / "..." / '...' declarations,
  // so `context.character.scenario += x;` can be resolved.
  const strConsts = {};
  {
    const sRe = /\b(?:const|let|var)\s+(\w+)\s*=\s*(?=[`'"])/g;
    let sm0;
    while ((sm0 = sRe.exec(code)) !== null) {
      const r = extractStrEnd(code, sRe.lastIndex);
      if (r) { strConsts[sm0[1]] = r.value; sRe.lastIndex = r.end; }
    }
  }

  function getBlock(src, braceIdx) {
    let depth = 0, i = braceIdx;
    while (i < src.length) {
      const nextI = skipStringOrComment(src, i);
      if (nextI > i) { i = nextI; continue; }
      const c = src[i];
      if (c === "{") depth++;
      else if (c === "}") { depth--; if (depth === 0) return { body: src.slice(braceIdx + 1, i), end: i }; }
      i++;
    }
    return null;
  }

  function extractStrEnd(src, from) {
    let i = from;
    while (i < src.length && src[i] !== "'" && src[i] !== '"' && src[i] !== "`") i++;
    if (i >= src.length) return null;
    const q = src[i];
    i++;
    let val = "";
    const ESCAPES = { n: "\n", t: "\t", r: "\r" };
    while (i < src.length) {
      if (src[i] === "\\" && i + 1 < src.length) {
        const esc = src[i + 1]; val += ESCAPES[esc] !== undefined ? ESCAPES[esc] : esc; i += 2; continue;
      }
      if (src[i] === q) { i++; break; }
      val += src[i]; i++;
    }
    return { value: val, end: i };
  }

  function extractStr(src, from) {
    const r = extractStrEnd(src, from);
    return r ? r.value : "";
  }

  function getKws(cond) {
    const msgVarEsc = msgVar.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    const kws = [];
    // Pattern: someKeywords.some(k => lastMessage.includes(k))  (also function form, or an inline [..] array)
    const someRe = new RegExp("(\\w+|\\[[^\\]]*\\])\\s*\\.some\\s*\\(\\s*(?:function\\s*)?\\(?\\s*(\\w+)\\s*\\)?\\s*(?:=>|\\{\\s*return)\\s*" + msgVarEsc + "\\.includes\\s*\\(\\s*\\2\\s*\\)", "g");
    let sm;
    while ((sm = someRe.exec(cond)) !== null) {
      const srcArr = sm[1];
      if (srcArr[0] === "[") {
        const strRe = /'([^'\\]*(?:\\.[^'\\]*)*)'|"([^"\\]*(?:\\.[^"\\]*)*)"/g;
        let x;
        while ((x = strRe.exec(srcArr)) !== null) kws.push(x[1] !== undefined ? x[1] : x[2]);
      } else if (kwConsts[srcArr]) kws.push(...kwConsts[srcArr]);
    }
    cond = cond.replace(someRe, " ");
    const re = new RegExp(msgVarEsc + "\\.includes\\s*\\(([^)]*)\\)", "g");
    let m;
    while ((m = re.exec(cond)) !== null) {
      const parts = m[1].split("||").map((s) => s.trim());
      parts.forEach((p) => {
        const sm = p.match(/^['"](.+)['"]$/);
        if (sm) kws.push(sm[1]);
        else if (kwConsts[p]) kws.push(...kwConsts[p]);
      });
    }
    return [...new Set(kws)];
  }

  function extractReplaceArg(body, callIdx) {
    const openParen = body.indexOf("(", callIdx);
    if (openParen === -1) return "";
    const first = extractStrEnd(body, openParen + 1);
    if (!first) return "";
    const second = extractStrEnd(body, first.end);
    return second ? second.value : "";
  }

  function getContent(body) {
    function readValue(src, from) {
      let i = from;
      while (i < src.length && /\s/.test(src[i])) i++;
      const c = src[i];
      if (c === "'" || c === '"' || c === "`") return extractStr(src, i);
      const idm = /^[A-Za-z_$][\w$]*/.exec(src.slice(i));
      if (idm && strConsts[idm[0]] !== undefined) return strConsts[idm[0]];
      return extractStr(src, from);
    }
    function collect(field) {
      const re = new RegExp("context\\.character\\." + field + "\\s*\\+=", "g");
      const parts = [];
      let m;
      while ((m = re.exec(body)) !== null) { const v = readValue(body, re.lastIndex); if (v) parts.push(v); }
      return parts.join("");
    }
    let p = collect("personality");
    let s = collect("scenario");
    if (!p) {
      const pr = body.indexOf("context.character.personality.replace(");
      if (pr !== -1) p = extractReplaceArg(body, pr);
    }
    if (!s) {
      const sr = body.indexOf("context.character.scenario.replace(");
      if (sr !== -1) s = extractReplaceArg(body, sr);
    }
    return { p, s };
  }

  function walk(src, minMsg) {
    let i = 0;
    while (i < src.length) {
      const nextI = skipStringOrComment(src, i);
      if (nextI > i) { i = nextI; continue; }
      const prevOk = i === 0 || /\W/.test(src[i - 1]);
      const nextOk = /[\s(]/.test(src[i + 2] || "");
      if (src.slice(i, i + 2) === "if" && prevOk && nextOk) {
        let j = i + 2;
        while (j < src.length && src[j] !== "(") j++;
        if (j >= src.length) { i++; continue; }
        let depth = 0, k = j;
        while (k < src.length) {
          if (src[k] === "(") depth++;
          else if (src[k] === ")") { depth--; if (depth === 0) break; }
          k++;
        }
        const cond = src.slice(j + 1, k);
        let bi = k + 1;
        while (bi < src.length && src[bi] !== "{") bi++;
        if (bi >= src.length) { i++; continue; }
        const block = getBlock(src, bi);
        if (!block) { i++; continue; }

        const mg = cond.match(/context\.chat\.message_count\s*[>>=]+\s*(\d+)/);
        if (mg) {
          walk(block.body, Math.max(minMsg, parseInt(mg[1]) + 1));
          i = block.end + 1; continue;
        }

        const kws = getKws(cond);
        if (kws.length) {
          const { p, s } = getContent(block.body);
          if (p || s) entries.push({ keywords: kws, personality: p, scenario: s, minMessages: minMsg, category: "If-Chain" });
          walk(block.body, minMsg);
          i = block.end + 1; continue;
        }
        i = block.end + 1; continue;
      }
      i++;
    }
  }
  walk(code, 0);
  return entries;
}

function dynArr(x) { return Array.isArray(x) ? x : x == null ? [] : [x]; }

function dynProbToPercent(v) {
  if (v == null) return null;
  if (typeof v === "number") return Math.round(v <= 1 ? v * 100 : v);
  const s = String(v).trim();
  const n = parseFloat(s.replace("%", ""));
  if (!isFinite(n)) return null;
  return Math.round(s.indexOf("%") !== -1 ? n : n * 100);
}

function extractTopLevelLiteral(code, varName, openChar) {
  const closeChar = openChar === "[" ? "]" : "}";
  const declRe = new RegExp("\\b(?:const|let|var)\\s+" + varName + "\\s*=\\s*");
  const m = declRe.exec(code);
  if (!m) return null;
  const start = code.indexOf(openChar, m.index + m[0].length);
  if (start === -1) return null;
  let depth = 0, i = start, end = -1;
  while (i < code.length) {
    const nextI = skipStringOrComment(code, i);
    if (nextI > i) { i = nextI; continue; }
    const c = code[i];
    if (c === openChar) depth++;
    else if (c === closeChar) { depth--; if (depth === 0) { end = i; break; } }
    i++;
  }
  if (end === -1) return null;
  return code.slice(start, end + 1);
}

function dynSafeEval(str) {
  if (!str) return null;
  try { return new Function('"use strict";return(' + str + ")")(); } catch (_) {}
  try { return new Function('"use strict";return(' + sanitizeQuotes(str) + ")")(); } catch (_) { return null; }
}

function dynHasUnsupportedGates(e) {
  return !!(e.notAll || e.andAnyTags || e.andAllTags || e.notAnyTags || e.notAllTags || e.requireEmotion || e.blockEmotion || e.maxMessages != null || e["prev.requireAny"] || e["prev.requireAll"] || e["prev.requireNone"] || e["prev.andAny"] || e["prev.andAll"] || e["prev.notAny"]);
}

function dynWordGates(e) {
  const r = e.requires || {};
  return {
    any: [].concat(dynArr(e.requireAny), dynArr(e.andAny), dynArr(r.any)),
    all: [].concat(dynArr(e.requireAll), dynArr(e.andAll), dynArr(r.all)),
    none: [].concat(dynArr(e.requireNone), dynArr(e.notAny), dynArr(e.block), dynArr(r.none)),
  };
}

function detectDynamicLore(code) { return /\bdynamicLore\s*=\s*\[/.test(code); }

function parseDynamicLore(code) {
  lastScriptRepairStats = null;
  let loreStr = extractTopLevelLiteral(code, "dynamicLore", "[");
  let dynamicLore = loreStr && dynSafeEval(loreStr);
  let workingCode = code;
  if (!Array.isArray(dynamicLore)) {
    workingCode = repairScriptText(code);
    loreStr = extractTopLevelLiteral(workingCode, "dynamicLore", "[");
    dynamicLore = loreStr && dynSafeEval(loreStr);
  }
  if (!Array.isArray(dynamicLore)) throw new Error("Could not parse dynamicLore array");

  const entityStr = extractTopLevelLiteral(workingCode, "ENTITY_DB", "{");
  const ENTITY_DB = (entityStr && dynSafeEval(entityStr)) || {};
  const relStr = extractTopLevelLiteral(workingCode, "RELATIONSHIP_DB", "[");
  const RELATIONSHIP_DB = (relStr && dynSafeEval(relStr)) || [];

  let all = dynamicLore.slice();
  for (const name in ENTITY_DB) {
    if (!Object.prototype.hasOwnProperty.call(ENTITY_DB, name)) continue;
    const ent = ENTITY_DB[name];
    if (ent && Array.isArray(ent.lore)) {
      for (const le of ent.lore) all.push(Object.assign({ _entityOwner: name }, le));
    }
  }

  const entityKwRe = /^char\.([a-z0-9_]+)$/i;
  function expandKeywords(list) {
    if (!Array.isArray(list)) return [];
    const out = [];
    for (const kw of list) {
      const m = String(kw).match(entityKwRe);
      if (m && ENTITY_DB[m[1].toLowerCase()]) {
        const ent = ENTITY_DB[m[1].toLowerCase()];
        out.push(m[1].toLowerCase());
        if (Array.isArray(ent.aliases)) out.push(...ent.aliases);
      } else out.push(kw);
    }
    return [...new Set(out)];
  }
  for (const e of all) {
    if (e.keywords) e.keywords = expandKeywords(e.keywords);
    if (Array.isArray(e.Shifts)) {
      for (const sh of e.Shifts) if (sh.keywords) sh.keywords = expandKeywords(sh.keywords);
    }
  }

  const byTag = {};
  for (const e of all) { if (e.tag) (byTag[e.tag] = byTag[e.tag] || []).push(e); }

  const stats = { inlined: 0, shifts: 0, nameBlocked: 0, relationships: 0, unsupportedGates: 0 };
  function collectTriggered(triggers, depth, seen) {
    let p = "", s = "";
    if (!triggers || !triggers.length || depth > 4) return { p, s };
    for (const tag of triggers) {
      const list = byTag[tag];
      if (!list) continue;
      for (const te of list) {
        if (seen.has(te)) continue;
        seen.add(te);
        if (dynHasUnsupportedGates(te)) stats.unsupportedGates++;
        if (te.personality) p += "\n\n" + te.personality;
        if (te.scenario) s += "\n\n" + te.scenario;
        stats.inlined++;
        if (Array.isArray(te.triggers) && te.triggers.length) {
          const nested = collectTriggered(te.triggers, depth + 1, seen);
          p += nested.p; s += nested.s;
        }
      }
    }
    return { p, s };
  }

  function toFlat(e, ownerLabel) {
    if (dynHasUnsupportedGates(e)) stats.unsupportedGates++;
    let keywords = e.keywords && e.keywords.length ? e.keywords : e["prev.keywords"] && e["prev.keywords"].length ? e["prev.keywords"] : [];
    const g = dynWordGates(e);
    const filters = {};
    if (keywords.length) {
      if (g.none.length) filters.notWith = g.none;
      else if (g.all.length) filters.requiresAll = g.all;
      else if (g.any.length) filters.requiresAny = g.any;
    } else if (g.any.length) {
      keywords = g.any.slice();
      if (g.none.length) filters.notWith = g.none;
    } else if (g.all.length) {
      keywords = [g.all[0]];
      if (g.all.length > 1) filters.requiresAll = g.all.slice(1);
      if (g.none.length) filters.notWith = g.none;
    } else if (g.none.length) {
      filters.notWith = g.none;
      stats.unsupportedGates++;
    }
    const { p: trigP, s: trigS } = collectTriggered(e.triggers, 0, new Set());
    const flat = {
      keywords, personality: (e.personality || "") + trigP, scenario: (e.scenario || "") + trigS,
      priority: typeof e.priority === "number" ? e.priority : 3, minMessages: typeof e.minMessages === "number" ? e.minMessages : 0, category: ownerLabel || "Advanced Lore", filters,
    };
    const prob = dynProbToPercent(e.probability);
    if (prob != null) flat.probability = prob;
    if (Array.isArray(e.nameBlock) && e.nameBlock.length) { flat.nameBlock = e.nameBlock.slice(); stats.nameBlocked++; }
    return flat;
  }

  const out = [];
  for (const e of all) {
    if (e.tag && !(e.keywords && e.keywords.length) && !e["prev.keywords"]) continue;
    const ownerLabel = e._entityOwner ? "Character \u2014 " + e._entityOwner : "Advanced Lore";
    out.push(toFlat(e, ownerLabel));
    if (Array.isArray(e.Shifts) && e.Shifts.length) {
      for (const sh of e.Shifts) {
        stats.shifts++;
        if (dynHasUnsupportedGates(sh)) stats.unsupportedGates++;
        const shKeys = sh.keywords && sh.keywords.length ? sh.keywords : e.keywords || [];
        const shFilters = e.keywords && e.keywords.length ? { requiresAll: e.keywords.slice() } : {};
        const shFlat = {
          keywords: shKeys, personality: sh.personality || "", scenario: sh.scenario || "",
          priority: typeof e.priority === "number" ? e.priority : 3, minMessages: typeof e.minMessages === "number" ? e.minMessages : 0, category: ownerLabel + " (Shift)", filters: shFilters,
        };
        const shProb = dynProbToPercent(sh.probability);
        if (shProb != null) shFlat.probability = shProb;
        if (Array.isArray(sh.nameBlock) && sh.nameBlock.length) { shFlat.nameBlock = sh.nameBlock.slice(); stats.nameBlocked++; }
        out.push(shFlat);
      }
    }
  }

  for (const rel of RELATIONSHIP_DB) {
    if (!Array.isArray(rel.pair) || rel.pair.length < 2) continue;
    stats.relationships++;
    const [a, b] = rel.pair;
    const aEnt = ENTITY_DB[a] || {}; const bEnt = ENTITY_DB[b] || {};
    const aKeys = [a].concat(Array.isArray(aEnt.aliases) ? aEnt.aliases : []);
    const bKeys = [b].concat(Array.isArray(bEnt.aliases) ? bEnt.aliases : []);
    out.push({
      keywords: aKeys, personality: rel.injection || "", scenario: "", priority: 3, minMessages: 0, category: "Relationship", filters: { requiresAll: bKeys },
    });
  }
  lastDynLoreStats = stats;
  return out;
}

function smartScenPos(category) {
  const c = (category || "").toLowerCase().replace(/[_-]/g, " ");
  if (/\bnpc\b|character/.test(c)) return "1";
  if (/location|place|dorm|building|region/.test(c)) return "1";
  return "0";
}

function catLabel(category) {
  const c = (category || "").toLowerCase().replace(/[_-]/g, " ").trim();
  if (/^world/.test(c)) return "World";
  if (/hierarchy/.test(c)) return "Social Hierarchy";
  if (/^culture/.test(c)) return "Culture";
  if (/^location/.test(c)) return "Location";
  if (/\bnpc\b|^character/.test(c)) return "NPC";
  if (/underground/.test(c)) return "Underground";
  if (/faction/.test(c)) return "Faction";
  if (/empire/.test(c)) return "Empire";
  if (/species/.test(c)) return "Species";
  if (/creature/.test(c)) return "Creature";
  if (/mechanic/.test(c)) return "Mechanics";
  if (/secret/.test(c)) return "Secret";
  if (/cosm/.test(c)) return "Cosmology";
  if (/if\s*chain/.test(c)) return "If-Chain";
  return c.split(" ").map((w) => (w ? w[0].toUpperCase() + w.slice(1) : "")).join(" ");
}

function entryLabel(entry) {
  const cat = (entry.category || "").toLowerCase();
  if (/\bnpc\b|^character/.test(cat) && entry.personality) {
    const m = entry.personality.match(/knows?\s+([A-Z][a-zA-Záàâäéèêëíìîïóòôöúùûüñç]+(?:\s+[A-Z][a-zA-Záàâäéèêëíìîïóòôöúùûüñç]+)*)\s+(?:very well|:)/i);
    if (m) {
      const pts = m[1].trim().split(/\s+/);
      return pts.length > 2 ? pts[0] + " " + pts[pts.length - 1] : m[1].trim();
    }
  }
  if (!(entry.keywords && entry.keywords[0])) {
    const text = (entry.personality || entry.scenario || "").replace(/\{\{[^}]*\}\}/g, "").trim();
    if (text) {
      const words = text.split(/\s+/).filter(Boolean);
      const snippet = words.slice(0, 5).join(" ").replace(/[.,;:!?]+$/, "");
      if (snippet) return snippet + (words.length > 5 ? "\u2026" : "");
    }
  }
  const kw = entry.keywords && entry.keywords[0] ? entry.keywords[0] : entry.category || "Entry";
  return kw.split(/\s+/).map((w) => (w ? w[0].toUpperCase() + w.slice(1) : "")).join(" ");
}

function cleanPersonality(text) {
  if (!text) return "";
  let t = text.trim().replace(/^,\s*/, "");
  t = t.charAt(0).toUpperCase() + t.slice(1);
  if (!/[.!?]$/.test(t)) t += ".";
  return t;
}

function cleanScenario(text) { return (text || "").trim(); }

function makeSTEntry(uid, key, keysecondary, selective, selectiveLogic, comment, content, posKey, order, delay, outletName, depthNum, extra) {
  const parts = String(posKey).split(":");
  const position = parseInt(parts[0]) || 0;
  const role = parts.length > 1 ? parseInt(parts[1]) : position === 7 ? null : 0;
  const isOutlet = position === 7;
  const ex = extra || {};
  return {
    uid, key, keysecondary, comment, content,
    constant: !!ex.constant, vectorized: false, selective, selectiveLogic, addMemo: true, order, position,
    disable: false, ignoreBudget: false, excludeRecursion: false, preventRecursion: false, matchPersonaDescription: false,
    matchCharacterDescription: false, matchCharacterPersonality: false, matchCharacterDepthPrompt: false, matchScenario: false,
    matchCreatorNotes: false, delayUntilRecursion: false, probability: ex.probability ?? 100, useProbability: ex.probability != null ? true : false,
    depth: depthNum ?? 4, outletName: isOutlet ? outletName || "" : "", group: "", groupOverride: false, groupWeight: 100,
    scanDepth: null, caseSensitive: null, matchWholeWords: null, useGroupScoring: null, automationId: "", role, sticky: 0, cooldown: 0, delay, triggers: [], displayIndex: uid,
    characterFilter: ex.characterFilter || { isExclude: false, names: [], tags: [] },
  };
}

function scriptToST(entries, catMap, outlet, opts) {
  const { merge, scenPref, persPref } = opts;
  const result = {};
  let uid = 0;
  entries.forEach((e) => {
    const keys = Array.isArray(e.keywords) ? e.keywords : [];
    let sec = [], selective = false, selectiveLogic = 0;
    if (e.filters) {
      if (e.filters.notWith && e.filters.notWith.length) { sec = e.filters.notWith; selectiveLogic = 2; selective = true; }
      else if (e.filters.requiresAll && e.filters.requiresAll.length) { sec = e.filters.requiresAll; selectiveLogic = 1; selective = true; }
      else if (e.filters.requiresAny && e.filters.requiresAny.length) { sec = e.filters.requiresAny; selectiveLogic = 0; selective = true; }
    }
    const order = (e.priority ?? 5) * 10, delay = e.minMessages ?? 0;
    const cmap = catMap[e.category] || { scen: "0", scenDepth: 4, pers: "7", persDepth: 4 };
    let label = entryLabel(e);
    if (e.category && e.category !== "If-Chain") label = catLabel(e.category) + " \u2014 " + label;
    const extra = {};
    if (typeof e.probability === "number") extra.probability = e.probability;
    if (Array.isArray(e.nameBlock) && e.nameBlock.length) { extra.characterFilter = { isExclude: true, names: e.nameBlock.slice(), tags: [] }; }
    if (!keys.length) extra.constant = true;
    let sText = e.scenario ? cleanScenario(e.scenario) : "";
    let pText = e.personality ? cleanPersonality(e.personality) : "";
    if (sText && scenPref) sText = scenPref.replace(/\\n/g, "\n") + sText;
    if (pText && persPref) pText = persPref.replace(/\\n/g, "\n") + pText;

    if (merge) {
      const combined = [sText, pText].filter(Boolean).join("\n\n");
      if (combined) {
        const st = makeSTEntry(uid, keys, sec, selective, selectiveLogic, label, combined, cmap.scen, order, delay, outlet, cmap.scenDepth, extra);
        st._type = "merged"; result[uid] = st; uid++;
      }
    } else {
      if (sText) {
        const st = makeSTEntry(uid, keys, sec, selective, selectiveLogic, label, sText, cmap.scen, order, delay, outlet, cmap.scenDepth, extra);
        st._type = "scenario"; result[uid] = st; uid++;
      }
      if (pText) {
        const st = makeSTEntry(uid, keys, sec, selective, selectiveLogic, label, pText, cmap.pers, order, delay, outlet, cmap.persDepth, extra);
        st._type = "personality"; result[uid] = st; uid++;
      }
    }
  });
  return { entries: result, count: uid };
}

function jaiToSt(e, uid) {
  const label = e.comment || e.name || e.title || "";
  let k = e.keys || e.keywords || e.keyword || e.key || [];
  if (typeof k === "string") k = k.split(",").map((s) => s.trim()).filter(Boolean);
  else if (!Array.isArray(k)) k = [k];
  const sanitized = sanitizeKeyList(k);
  k = sanitized.keys;
  let ks = e.keysecondary || e.secondary_keys || [];
  if (typeof ks === "string") ks = ks.split(",").map((s) => s.trim()).filter(Boolean);
  else if (!Array.isArray(ks)) ks = [ks];
  let content = e.content || "";
  if (sanitized.movedText) content = sanitized.movedText + (content ? "\n\n" + content : "");
  const st = makeSTEntry(uid, k, ks, e.selective ?? true, e.selectiveLogic ?? 0, label, content, "0", e.order ?? 100, e.minMessages ?? 0, "", 4);
  st._type = "jai";
  st._flaggedKeys = !!sanitized.movedText;
  const out = { ...e, ...st };
  // keep the editor's raw-text mirrors of the key fields in sync with the cleaned keys
  if (!sanitized.movedText) {
    if ("keysRaw" in out) out.keysRaw = k.join(", ");
    if ("keywordsRaw" in out) out.keywordsRaw = k.join(", ");
  } else {
    if ("keysRaw" in out) out.keysRaw = "";
    if ("keywordsRaw" in out) out.keywordsRaw = "";
  }
  return out;
}

// ---------------------------------------------------------------------------
// Extension glue
// ---------------------------------------------------------------------------
function resetStats() {
    lastScriptRepairStats = null;
    lastBadKeyStats = null;
    lastKeySplitStats = null;
    lastDynLoreStats = null;
}

// Same output shape as the web converter's buildExportJson (compact), minus its internal markers.
function buildExportJson(entriesObj) {
    const exportObj = { entries: {} };
    Object.keys(entriesObj).forEach((k) => {
        exportObj.entries[k] = { ...entriesObj[k] };
        delete exportObj.entries[k]._type;
        delete exportObj.entries[k]._flaggedKeys;
    });
    return JSON.stringify(exportObj);
}

const baseName = (fileName) => fileName.replace(/\.[^./\\]+$/, '');

// Same rule as the web converter's getOutlet(), with the file name standing in for the lorebook name.
function outletFromName(fileName) {
    const n = baseName(fileName).trim();
    return n ? n.toLowerCase().replace(/[^a-z0-9]/g, '_') + '_knowledge' : 'lorebook_knowledge';
}

// Same warnings as the web converter's buildWarnings(), for the formats the extension converts.
function buildNotes(fmt, sourceEntries, stEntries, L) {
    const w = [];
    if (lastKeySplitStats && lastKeySplitStats.entries) w.push(L.keysSplit(lastKeySplitStats.entries, lastKeySplitStats.keys));
    if (fmt === 'jai_json') {
        if (lastBadKeyStats && lastBadKeyStats.entries) w.push(L.badKeywords(lastBadKeyStats.entries, lastBadKeyStats.dropped));
        w.push(L.jaiPosSet);
        return w;
    }
    if (lastScriptRepairStats && (lastScriptRepairStats.unclosedStrings || lastScriptRepairStats.missingCommas || lastScriptRepairStats.prematureCloses)) {
        w.push(L.scriptRepaired(lastScriptRepairStats));
    }
    w.push(L.scriptSplit);
    if (Object.values(stEntries).some((e) => e.position === 7)) w.push(L.outletReminder);
    if (fmt === 'jai_script') {
        const f = sourceEntries.filter((e) => e.filters && Object.keys(e.filters).length > 0);
        if (f.length) w.push(L.filtersUsed(f.length));
    }
    if (fmt === 'if_chain') w.push(L.ifChainDup);
    if (fmt === 'dynamic_lore' && lastDynLoreStats) {
        const s = lastDynLoreStats;
        if (s.inlined) w.push(L.dynLoreInlined(s.inlined));
        if (s.shifts) w.push(L.dynLoreShifts(s.shifts));
        if (s.unsupportedGates) w.push(L.dynLoreGates(s.unsupportedGates));
        if (s.nameBlocked) w.push(L.dynLoreNameBlock(s.nameBlocked));
        if (s.relationships) w.push(L.dynLoreRelationships(s.relationships));
    }
    return w;
}

// Anything ST already reads itself: an object with its own `entries` (NOT `'entries' in x`, which
// is true for arrays), or the Novel / Agnai / Risu formats that ST converts before importing.
function isNativeStFormat(parsed) {
    if (!parsed || typeof parsed !== 'object' || Array.isArray(parsed)) return false;
    return Object.prototype.hasOwnProperty.call(parsed, 'entries')
        || parsed.lorebookVersion !== undefined
        || parsed.kind === 'memory'
        || parsed.type === 'risu';
}

const PASS = { action: 'pass' };
const CANCEL = { action: 'cancel' };
let passthrough = false;

/**
 * Decides what to do with the picked text. Returns PASS (let ST handle the file), CANCEL
 * (show a message, import nothing) or { action: 'ask', ... } with the converted result.
 */
function analyze(text, fileName, L) {
    resetStats();
    const trimmed = text.trim();
    let parsed = null;
    if (trimmed.startsWith('{') || trimmed.startsWith('[')) {
        try { parsed = JSON.parse(trimmed); } catch { /* not JSON */ }
    }
    if (isNativeStFormat(parsed)) return PASS;

    const det = detectFormat(text);
    if (!det || det.type === 'st_json') return PASS;

    const formatLabel = L.formats[det.type];
    if (OPT_IN_FORMATS.has(det.type) && !getSettings().advancedScripts) {
        toastr.warning(L.scriptsDisabled(fileName, formatLabel));
        return CANCEL;
    }

    let stEntries, srcEntries, stCount;
    if (det.type === 'jai_json') {
        srcEntries = det.parsed.filter((x) => x && typeof x === 'object');
        if (!srcEntries.length) return PASS;
        stEntries = {};
        srcEntries.forEach((e, i) => { stEntries[i] = jaiToSt(e, i); });
        stCount = srcEntries.length;
    } else {
        try {
            srcEntries = det.type === 'jai_script' ? parseScript(text)
                : det.type === 'dynamic_lore' ? parseDynamicLore(text)
                    : parseIfChain(text);
        } catch (err) {
            toastr.error(det.type === 'if_chain' ? L.ifChainParseFailed(err.message) : L.scriptParseFailed(err.message));
            return CANCEL;
        }
        if (!srcEntries.length) {
            toastr.warning(det.type === 'if_chain' ? L.noKeywordEntries() : L.noEntriesLoreEntries());
            return CANCEL;
        }
        // Defaults of the web converter's Simple mode: scenario text by category, personality text at 4:0 (@ Depth, System role, depth 4).
        const catMap = {};
        for (const cat of new Set(srcEntries.map((e) => e.category || ''))) {
            catMap[cat] = { scen: smartScenPos(cat), scenDepth: 4, pers: '4:0', persDepth: 4 };
        }
        const res = scriptToST(srcEntries, catMap, outletFromName(fileName), { merge: false, scenPref: '', persPref: '' });
        stEntries = res.entries;
        stCount = res.count;
    }
    if (!stCount) {
        toastr.warning(L.noKeywordEntries());
        return CANCEL;
    }
    return {
        action: 'ask',
        json: buildExportJson(stEntries),
        formatLabel,
        srcCount: srcEntries.length,
        stCount,
        notes: buildNotes(det.type, srcEntries, stEntries, L),
    };
}

function buildPopupContent(fileName, result, L) {
    const root = document.createElement('div');
    const heading = document.createElement('h3');
    heading.textContent = L.popupTitle;
    const body = document.createElement('p');
    body.textContent = L.popupBody(fileName, result.formatLabel, result.srcCount, result.stCount);
    root.append(heading, body);
    if (result.notes.length) {
        const list = document.createElement('ul');
        for (const note of result.notes) {
            const li = document.createElement('li');
            li.innerHTML = note; // note strings are built only from the converter's own constants and numbers
            list.append(li);
        }
        root.append(list);
    }
    return root;
}

async function readImportText(file) {
    if (/\.docx$/i.test(file.name)) return extractDocxText(await file.arrayBuffer());
    return file.text();
}

async function prepareImport(file) {
    const L = getStrings();
    let text;
    try {
        text = await readImportText(file);
    } catch (err) {
        if (/\.docx$/i.test(file.name)) {
            toastr.error(L.docxReadFailed(err.message));
            return CANCEL;
        }
        return PASS;
    }

    let result;
    try {
        result = analyze(text, file.name, L);
    } catch (err) {
        console.error(`[${MODULE_NAME}] conversion failed`, err);
        toastr.error(L.convertFailed);
        return CANCEL; // importing the raw file would not give a usable lorebook
    }
    if (result.action !== 'ask') return result;

    const { callGenericPopup, POPUP_TYPE, POPUP_RESULT } = SillyTavern.getContext();
    const content = buildPopupContent(file.name, result, L);
    const answer = await callGenericPopup(content, POPUP_TYPE.CONFIRM, '', { okButton: L.convert, cancelButton: L.cancel, wide: true, allowVerticalScrolling: true });
    if (answer !== POPUP_RESULT.AFFIRMATIVE) return CANCEL;

    // ST derives the lorebook name from everything before the last dot in the file name.
    return { action: 'convert', file: new File([result.json], baseName(file.name) + '.json', { type: 'application/json' }) };
}

function refire(input, replacementFile) {
    if (replacementFile) {
        const transfer = new DataTransfer();
        transfer.items.add(replacementFile);
        input.files = transfer.files;
    }
    passthrough = true;
    try {
        input.dispatchEvent(new Event('change', { bubbles: true }));
    } finally {
        passthrough = false;
    }
}

async function onWorldImportChange(event) {
    if (passthrough) return;
    const input = event.target;
    if (!(input instanceof HTMLInputElement) || input.id !== IMPORT_INPUT_ID) return;
    if (!getSettings().enabled) return;
    const file = input.files?.[0];
    if (!file || /\.png$/i.test(file.name)) return; // ST reads PNG lorebooks itself

    // Reading the file is async, but ST's handler has to be stopped synchronously.
    // So it is always stopped here, and the change event is always re-fired below
    // (with the original file or the converted one), so an import is never lost.
    event.stopPropagation();

    let decision = PASS;
    try {
        decision = await prepareImport(file);
    } catch (err) {
        console.error(`[${MODULE_NAME}]`, err);
    }

    if (decision.action === 'cancel') {
        input.value = ''; // lets the same file be picked again
        return;
    }
    refire(input, decision.action === 'convert' ? decision.file : null);
}

// ---------------------------------------------------------------------------
// ST's import picker only offers .json/.lorebook/.png. `accept` is just a file-picker filter,
// so while the converter is on, the formats this extension converts are added to it.
// ---------------------------------------------------------------------------
let originalAccept = null;
function syncAcceptAttribute() {
    const input = document.getElementById(IMPORT_INPUT_ID);
    if (!input) return;
    if (originalAccept === null) originalAccept = input.getAttribute('accept') ?? '';
    if (!getSettings().enabled) {
        input.setAttribute('accept', originalAccept);
        return;
    }
    const list = new Set(originalAccept.split(',').map((s) => s.trim()).filter(Boolean));
    EXTRA_ACCEPT.forEach((ext) => list.add(ext));
    input.setAttribute('accept', [...list].join(','));
}

// ---------------------------------------------------------------------------
// Settings panel (uses ST's own classes; no stylesheet of our own)
// ---------------------------------------------------------------------------
function addSettingsPanel() {
    const L = getStrings();
    const html = `
    <div id="jai_lorebook_converter_settings">
        <div class="inline-drawer">
            <div class="inline-drawer-toggle inline-drawer-header">
                <b class="jai_conv_title"></b>
                <div class="inline-drawer-icon fa-solid fa-circle-chevron-down down"></div>
            </div>
            <div class="inline-drawer-content">
                <label class="checkbox_label" for="jai_conv_enabled">
                    <input id="jai_conv_enabled" type="checkbox" />
                    <span class="jai_conv_enable_text"></span>
                </label>
                <label class="checkbox_label" for="jai_conv_scripts">
                    <input id="jai_conv_scripts" type="checkbox" />
                    <span class="jai_conv_scripts_text"></span>
                </label>
                <small class="jai_conv_scripts_hint"></small>
            </div>
        </div>
    </div>`;
    const $panel = $(html);
    $panel.find('.jai_conv_title').text(L.title);
    $panel.find('.jai_conv_enable_text').text(L.enable);
    $panel.find('.jai_conv_scripts_text').text(L.enableScripts);
    $panel.find('.jai_conv_scripts_hint').text(L.enableScriptsHint);
    $panel.find('#jai_conv_enabled').prop('checked', getSettings().enabled).on('change', function () {
        getSettings().enabled = this.checked;
        SillyTavern.getContext().saveSettingsDebounced();
        syncAcceptAttribute();
    });
    $panel.find('#jai_conv_scripts').prop('checked', getSettings().advancedScripts).on('change', function () {
        getSettings().advancedScripts = this.checked;
        SillyTavern.getContext().saveSettingsDebounced();
    });
    $('#extensions_settings2').append($panel);
}

jQuery(() => {
    addSettingsPanel();
    syncAcceptAttribute();
    // Capture phase on document: runs before ST's own jQuery handler on #world_import_file.
    document.addEventListener('change', onWorldImportChange, true);
});
