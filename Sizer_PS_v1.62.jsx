#target photoshop
app.bringToFront();

var TARGET_DPI = 300;
var RESAMPLE = ResampleMethod.BICUBIC;
var REPORT_FLUSH_INTERVAL = 5;
var CHECKPOINT_REPORT_WRITES_ENABLED = false;
var FILE_WRITE_RETRY_COUNT = 3;
var FILE_WRITE_RETRY_DELAY_MS = 80;

function trimStr(s) { return String(s).replace(/^\s+|\s+$/g, ""); }
function round2(n) { return Math.round(n * 100) / 100; }
function roundMoney(n) { return Math.round((n + 0.0000001) * 100) / 100; }
function pad2(n) { return (n < 10 ? "0" : "") + n; }

function sleepMs(ms) {
    try { $.sleep(ms); } catch (e) {}
}

function formatDiagnosticValue(value) {
    if (value === null || value === undefined) return "";
    if (typeof value === "number") return isNaN(value) ? "NaN" : String(value);
    if (typeof value === "boolean") return value ? "true" : "false";
    return String(value).replace(/\r?\n/g, " \\n ");
}

function safeErrorMessage(err) {
    if (err === null || err === undefined) return "";
    try {
        if (err.message) return String(err.message);
    } catch (eMessage) {}
    try {
        return String(err);
    } catch (eString) {}
    return "Unknown error";
}

function safeErrorStack(err) {
    if (!err) return "";
    try {
        if (err.stack) return String(err.stack);
    } catch (eStack) {}
    try {
        var parts = [];
        if (err.fileName) parts.push(String(err.fileName));
        if (err.line) parts.push("line " + String(err.line));
        return parts.join(" | ");
    } catch (eFile) {}
    return "";
}

function initDiagnosticsState() {
    return {
        startedAt: (new Date()).toString(),
        lines: [],
        events: []
    };
}

function addDiagnostic(state, level, eventName, details) {
    if (!state) return;
    var timeStamp = (new Date()).toString();
    var parts = [];
    var eventDetails = {};
    if (details) {
        for (var key in details) {
            if (!details.hasOwnProperty(key)) continue;
            eventDetails[key] = details[key];
            parts.push(key + "=" + formatDiagnosticValue(details[key]));
        }
    }
    state.lines.push("[" + String(level || "info").toUpperCase() + "] " + timeStamp + " | " + String(eventName || "") + (parts.length ? " | " + parts.join(" | ") : ""));
    state.events.push({
        time: timeStamp,
        level: String(level || "info"),
        event: String(eventName || ""),
        details: eventDetails
    });
}

function jsonEscapeString(s) {
    s = String(s);
    return s.replace(/\\/g, "\\\\").replace(/"/g, "\\\"").replace(/\r/g, "\\r").replace(/\n/g, "\\n").replace(/\t/g, "\\t");
}

function jsonStringifyLoose(value) {
    if (value === null) return "null";
    if (value === undefined) return "null";
    var t = typeof value;
    if (t === "string") return "\"" + jsonEscapeString(value) + "\"";
    if (t === "number") return isNaN(value) || !isFinite(value) ? "\"" + String(value) + "\"" : String(value);
    if (t === "boolean") return value ? "true" : "false";
    if (value instanceof Array) {
        var arr = [];
        for (var i = 0; i < value.length; i++) arr.push(jsonStringifyLoose(value[i]));
        return "[" + arr.join(",") + "]";
    }
    var props = [];
    for (var key in value) {
        if (!value.hasOwnProperty(key)) continue;
        props.push("\"" + jsonEscapeString(key) + "\":" + jsonStringifyLoose(value[key]));
    }
    return "{" + props.join(",") + "}";
}

function buildDiagnosticsText(state) {
    var lines = [];
    lines.push("Sizer Diagnostics");
    lines.push("Started: " + formatDiagnosticValue(state ? state.startedAt : ""));
    lines.push("");
    if (state && state.lines && state.lines.length) {
        for (var i = 0; i < state.lines.length; i++) lines.push(state.lines[i]);
    } else {
        lines.push("[INFO] No diagnostics events captured.");
    }
    return lines.join("\r\n") + "\r\n";
}

function buildDiagnosticsJson(state) {
    return jsonStringifyLoose({
        app: "Photoshop",
        startedAt: state ? state.startedAt : "",
        events: state ? state.events : []
    });
}

function createManagedWriteDescriptor(defaultFileObj, result) {
    var actualPath = result && result.path ? result.path : defaultFileObj.fsName;
    var fallbackUsed = actualPath !== defaultFileObj.fsName;
    return {
        ok: !!(result && result.ok),
        path: actualPath,
        name: File(actualPath).name,
        warning: result && result.warning ? result.warning : "",
        error: result && result.error ? result.error : "",
        fallbackUsed: fallbackUsed
    };
}

function describeManagedWrite(descriptor) {
    if (!descriptor || !descriptor.path) return "";
    if (descriptor.fallbackUsed) return descriptor.name + " (fallback: " + descriptor.path + ")";
    return descriptor.name;
}

function makeTimestampTag() {
    var d = new Date();
    return d.getFullYear() + pad2(d.getMonth() + 1) + pad2(d.getDate()) + "_" + pad2(d.getHours()) + pad2(d.getMinutes()) + pad2(d.getSeconds());
}

function stripExt(name) {
    var s = String(name);
    var i = s.lastIndexOf(".");
    return (i > 0) ? s.substring(0, i) : s;
}

function safeOpen(fileObj) {
    try { return app.open(fileObj); }
    catch (e) { return null; }
}

function ensureRgbDocument(doc) {
    if (!doc) return false;
    try {
        if (doc.mode !== DocumentMode.RGB) doc.changeMode(ChangeMode.RGB);
        return doc.mode === DocumentMode.RGB;
    } catch (e) {
        return false;
    }
}

function makeBaseWithQtyOption(qty, base, option) {
    qty = parseInt(qty, 10);
    if (isNaN(qty) || qty < 1) qty = 1;
    if (option === "filenameQty") return base + "___" + qty;
    if (option === "qtyFilename") return qty + "___" + base;
    return base;
}

function decodeNumericEntitiesLoose(s) {
    s = String(s);
    s = s.replace(/&#(\d+);?/g, function (_, num) {
        var code = parseInt(num, 10);
        if (isNaN(code)) return _;
        try { return String.fromCharCode(code); } catch (e) { return _; }
    });
    s = s.replace(/&#x([0-9a-fA-F]+);?/g, function (_, hex) {
        var code = parseInt(hex, 16);
        if (isNaN(code)) return _;
        try { return String.fromCharCode(code); } catch (e) { return _; }
    });
    return s;
}

function decodePercentEscapesLoose(s) {
    s = String(s);
    return s.replace(/(?:%[0-9A-Fa-f]{2})+/g, function (chunk) {
        try { return decodeURIComponent(chunk); } catch (e) { return chunk; }
    });
}

function normalizeForMatch(s) {
    s = trimStr(s);
    s = decodeNumericEntitiesLoose(s);
    s = decodePercentEscapesLoose(s);
    s = s.replace(/[\u2010\u2011\u2012\u2013\u2014\u2015\u2212]/g, "-");
    s = s.replace(/[\u00A0\u2000-\u200B\u202F\u205F\u3000]/g, " ");
    s = trimStr(s).replace(/\s+/g, " ");
    return s;
}

function canonicalKey(s) {
    s = normalizeForMatch(s).toLowerCase();
    s = s.replace(/[\s_\-]+/g, "-");
    s = s.replace(/\u00D7/g, "x");
    s = s.replace(/^-+/, "").replace(/-+$/, "");
    return s;
}

function ultraLooseKey(s) {
    s = canonicalKey(s);
    s = s.replace(/[\-_ ]+/g, "");
    return s;
}

function commonPrefixLen(a, b) {
    var m = Math.min(a.length, b.length);
    var i = 0;
    while (i < m && a.charAt(i) === b.charAt(i)) i++;
    return i;
}

function similarityScore(a, b) {
    a = String(a);
    b = String(b);
    if (!a || !b) return 0;
    if (a === b) return 9999;
    return (commonPrefixLen(a, b) * 3) + ((a.indexOf(b) >= 0 || b.indexOf(a) >= 0) ? 10 : 0) - Math.abs(a.length - b.length);
}

function findFileMatchByEmailName(fileList, emailFileName) {
    var emailRaw = String(emailFileName);
    var emailNorm = normalizeForMatch(emailRaw);
    var emailCanon = canonicalKey(emailRaw);
    var emailLoose = ultraLooseKey(emailRaw);
    var i, f, c, l;

    for (i = 0; i < fileList.length; i++) {
        f = fileList[i];
        if (f.name === emailRaw) return { file: f, matchType: "exact", suggested: "" };
    }
    for (i = 0; i < fileList.length; i++) {
        f = fileList[i];
        if (normalizeForMatch(f.name) === emailNorm) return { file: f, matchType: "normalized", suggested: "" };
    }
    for (i = 0; i < fileList.length; i++) {
        f = fileList[i];
        c = canonicalKey(f.name);
        if (c === emailCanon) return { file: f, matchType: "canonical", suggested: "" };
    }
    for (i = 0; i < fileList.length; i++) {
        f = fileList[i];
        l = ultraLooseKey(f.name);
        if (l === emailLoose) return { file: f, matchType: "ultraLoose", suggested: "" };
    }

    var bestFile = null;
    var bestScore = -999999;
    for (i = 0; i < fileList.length; i++) {
        f = fileList[i];
        c = canonicalKey(f.name);
        l = ultraLooseKey(f.name);
        var sc = similarityScore(emailCanon, c) + similarityScore(emailLoose, l);
        if (sc > bestScore) {
            bestScore = sc;
            bestFile = f;
        }
    }

    if (bestFile && bestScore >= 8) return { file: null, matchType: "suggestion", suggested: bestFile.name };
    return { file: null, matchType: "missing", suggested: "" };
}

function normalizeCurrencyLabel(token) {
    token = trimStr(String(token || ""));
    if (!token) return "$";
    var upper = token.toUpperCase();
    if (upper === "$" || upper === "CAD") return "$";
    if (upper === "USD") return "USD";
    if (upper === "EUR" || token === "€") return "EUR";
    if (upper === "GBP" || token === "£") return "GBP";
    return token;
}

function parseMoneyNumber(text) {
    var clean = String(text || "").replace(/,/g, "");
    var n = parseFloat(clean);
    return isNaN(n) ? NaN : roundMoney(n);
}

function parseMoneyToken(text) {
    var src = String(text || "");
    var m = /(?:CAD|USD|EUR|GBP|\$|€|£)\s*([\d,]+(?:\.\d+)?)/i.exec(src);
    if (m) {
        return { currency: normalizeCurrencyLabel(m[0].replace(/[\d,.\s]+/g, "")), amount: parseMoneyNumber(m[1]) };
    }
    m = /([\d,]+(?:\.\d+)?)\s*(CAD|USD|EUR|GBP|\$|€|£)/i.exec(src);
    if (m) {
        return { currency: normalizeCurrencyLabel(m[2]), amount: parseMoneyNumber(m[1]) };
    }
    return { currency: "$", amount: NaN };
}

function formatMoney(amount, currency) {
    if (isNaN(amount)) return "";
    var cur = normalizeCurrencyLabel(currency);
    var absNum = Math.abs(roundMoney(amount)).toFixed(2);
    var sign = amount < 0 ? "-" : "";
    if (cur === "$" || cur === "€" || cur === "£") return sign + cur + absNum;
    return sign + cur + " " + absNum;
}

function extractQtyAndPrice(blockText) {
    var lines = String(blockText || "").split(/\r?\n/);
    for (var i = lines.length - 1; i >= 0; i--) {
        var line = trimStr(lines[i]);
        if (!line) continue;
        var m = /^(\d{1,4})\s+(CAD|USD|EUR|GBP|\$|€|£)\s*([\d,]+(?:\.\d+)?)$/i.exec(line);
        if (m) return { qty: parseInt(m[1], 10), price: parseMoneyNumber(m[3]), currency: normalizeCurrencyLabel(m[2]) };
        m = /^(\d{1,4})\s+([\d,]+(?:\.\d+)?)\s*(CAD|USD|EUR|GBP|\$|€|£)$/i.exec(line);
        if (m) return { qty: parseInt(m[1], 10), price: parseMoneyNumber(m[2]), currency: normalizeCurrencyLabel(m[3]) };
    }

    var qtyMatch = /(?:^|\s)(\d{1,4})\s*(?=(?:CAD|USD|EUR|GBP|\$|€|£)\s*[\d.,]+)/i.exec(String(blockText));
    var money = parseMoneyToken(blockText);
    var qty = (qtyMatch && qtyMatch[1]) ? parseInt(qtyMatch[1], 10) : 1;
    if (isNaN(qty) || qty < 1) qty = 1;
    return { qty: qty, price: money.amount, currency: money.currency };
}

function extractQtyBeforePrice(blockText) {
    return extractQtyAndPrice(blockText).qty;
}

function extractMoneyAfterLabel(text, label) {
    var safe = String(label).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    var re = new RegExp(safe + "\\s*:?\\s*([^\\r\\n]*)", "i");
    var m = re.exec(String(text || ""));
    if (!m || !m[1]) return { currency: "$", amount: NaN };
    return parseMoneyToken(m[1]);
}

function parseEmailFinancials(text, items, orderFormat, diagnosticsState) {
    var subtotal = extractMoneyAfterLabel(text, "Subtotal");
    var shipping = extractMoneyAfterLabel(text, "Shipping");
    var tax = extractMoneyAfterLabel(text, "HST");
    var total = extractMoneyAfterLabel(text, "Total");
    var taxRateMatch = /HST\s*\(([\d.]+)%\)/i.exec(String(text || ""));
    if (orderFormat === "us") {
        var available = !isNaN(subtotal.amount) || !isNaN(shipping.amount) || !isNaN(tax.amount) || !isNaN(total.amount);
        if (!available) addDiagnostic(diagnosticsState, "warn", "financials_unavailable", { format: orderFormat, reason: "No reliable totals found in pasted text" });
        return {
            currency: subtotal.currency || shipping.currency || tax.currency || total.currency || "$",
            subtotal: subtotal.amount,
            shipping: shipping.amount,
            tax: tax.amount,
            total: total.amount,
            taxRate: taxRateMatch ? parseFloat(taxRateMatch[1]) : NaN
        };
    }
    var subtotalSum = 0;
    var subtotalCount = 0;
    var currency = subtotal.currency || shipping.currency || tax.currency || total.currency || "$";
    for (var i = 0; i < items.length; i++) {
        if (!isNaN(items[i].price)) {
            subtotalSum += items[i].price;
            subtotalCount++;
        }
        if (!currency && items[i].currency) currency = items[i].currency;
    }
    return {
        currency: currency || "$",
        subtotal: !isNaN(subtotal.amount) ? subtotal.amount : (subtotalCount ? roundMoney(subtotalSum) : NaN),
        shipping: shipping.amount,
        tax: tax.amount,
        total: total.amount,
        taxRate: taxRateMatch ? parseFloat(taxRateMatch[1]) : 13
    };
}

function calculateAdjustedPrice(orderW, orderH, actualW, actualH, currentPrice) {
    if (isNaN(orderW) || isNaN(orderH) || isNaN(actualW) || isNaN(actualH) || isNaN(currentPrice)) return NaN;
    var oldArea = orderW * orderH;
    var newArea = actualW * actualH;
    if (oldArea <= 0 || newArea < 0) return NaN;
    return roundMoney(currentPrice * (newArea / oldArea));
}

function formatResizeModeLabel(mode) {
    if (mode === "respectWidth") return "Respect Width";
    if (mode === "respectHeight") return "Respect Height";
    if (mode === "stretch") return "Stretch";
    return String(mode || "");
}

function formatFilenameFormatLabel(mode) {
    if (mode === "filenameQty") return "Filename___Qty";
    if (mode === "qtyFilename") return "Qty___Filename";
    if (mode === "filename") return "Filename";
    return String(mode || "");
}

function formatPrintTypeModeLabel(mode) {
    if (mode === "folder") return "Folder";
    if (mode === "prefix") return "Prefix";
    if (mode === "none") return "None";
    return String(mode || "");
}

var PRINT_TYPE_RULES = [
    { type: "UV", re: /\b(?:UV\s*DTF|DTF\s*UV)\b/i },
    { type: "COOL", re: /\b(?:COOL\s*DTF|DTF\s*COOL)\b/i },
    { type: "HEAT", re: /\b(?:HEAT\s*DTF|DTF\s*HEAT)\b/i },
    { type: "Glitter", re: /\b(?:GLITTER\s*DTF|DTF\s*GLITTER)\b/i },
    { type: "Dyeblocker", re: /\b(?:DYE[\s-]*BLOCKER\s*DTF|DTF\s*DYE[\s-]*BLOCKER|DYEBLOCKER\s*DTF|DTF\s*DYEBLOCKER)\b/i }
];

function detectPrintType(text) {
    var hay = String(text || "");
    for (var i = 0; i < PRINT_TYPE_RULES.length; i++) {
        if (PRINT_TYPE_RULES[i].re.test(hay)) return PRINT_TYPE_RULES[i].type;
    }
    return "";
}

function lineStartAt(text, index) {
    var i = Math.min(Math.max(0, index), String(text).length);
    while (i > 0 && text.charAt(i - 1) !== "\n") i--;
    return i;
}

function previousNonEmptyLine(text, beforeIndex) {
    var end = Math.min(String(text).length, Math.max(0, beforeIndex));
    while (end > 0) {
        var start = lineStartAt(text, end);
        var line = trimStr(text.substring(start, end).replace(/\r/g, ""));
        if (line) return { start: start, end: end, text: line };
        end = start > 0 ? start - 1 : 0;
    }
    return { start: 0, end: 0, text: "" };
}

function findSectionEnd(text, fromIndex) {
    var m = /(?:^|\r?\n)(?:Subtotal:|Shipping:|Rush order:|GST\s*\(|HST\s*\(|PST\s*\(|QST\s*\(|Total:|Billing address|Shipping address)\b/i.exec(String(text).substring(fromIndex));
    return m ? (fromIndex + m.index) : String(text).length;
}

function extractItemDimension(blockText, labels) {
    var block = String(blockText || "");
    for (var i = 0; i < labels.length; i++) {
        var re = new RegExp(labels[i] + "\\s*:\\s*([\\d.]+)", "i");
        var m = re.exec(block);
        if (m && m[1]) {
            var n = parseFloat(m[1]);
            if (!isNaN(n)) return n;
        }
    }
    return NaN;
}

function findNextNonEmptyLine(lines, startIndex) {
    for (var i = Math.max(0, startIndex); i < lines.length; i++) {
        if (trimStr(lines[i])) return { index: i, text: trimStr(lines[i]) };
    }
    return { index: -1, text: "" };
}

function findLineIndexEquals(lines, label, startIndex) {
    var needle = trimStr(String(label || "")).toLowerCase();
    for (var i = Math.max(0, startIndex || 0); i < lines.length; i++) {
        if (trimStr(lines[i]).toLowerCase() === needle) return i;
    }
    return -1;
}

function parseCustomSizeLine(text) {
    var m = /([\d.]+)\s*[x×]\s*([\d.]+)/i.exec(String(text || ""));
    if (!m) return { width: NaN, height: NaN };
    return { width: parseFloat(m[1]), height: parseFloat(m[2]) };
}

function isProbablyUSOrderText(text) {
    var src = String(text || "");
    if (/Get Order Details\s*-\s*Wemust US/i.test(src)) return true;
    return /Orders\s*#\d+\s*Details/i.test(src) && /Custom Size/i.test(src) && /File Name/i.test(src) && /(?:^|\r?\n)\s*#\d+\s*(?:\r?\n|$)/i.test(src);
}

function parseEmailItemsClassic(emailText, diagnosticsState) {
    var text = String(emailText || "");
    var hits = [];
    var re = /Width:/gi;
    var m;
    while ((m = re.exec(text)) !== null) hits.push(m.index);
    if (!hits.length) return [];

    var markers = [];
    for (var i = 0; i < hits.length; i++) {
        var widthIdx = hits[i];
        var widthLineStart = lineStartAt(text, widthIdx);
        var productLine = previousNonEmptyLine(text, widthLineStart - 1);
        if (!productLine.text) continue;
        markers.push({ itemStart: productLine.start, widthIndex: widthIdx, productLabel: productLine.text });
    }
    if (!markers.length) return [];

    var sectionEnd = findSectionEnd(text, markers[0].itemStart);
    var scopedMarkers = [];
    for (var mi = 0; mi < markers.length; mi++) {
        if (markers[mi].itemStart >= sectionEnd) break;
        scopedMarkers.push(markers[mi]);
    }
    if (!scopedMarkers.length) return [];

    var items = [];
    var fRe = /Image file upload:\s*([^\r\n]+)/i;

    for (var j = 0; j < scopedMarkers.length; j++) {
        var itemStart = scopedMarkers[j].itemStart;
        var itemEnd = (j + 1 < scopedMarkers.length) ? scopedMarkers[j + 1].itemStart : sectionEnd;
        var block = text.substring(itemStart, itemEnd);
        var fm = fRe.exec(block);
        if (!fm) {
            addDiagnostic(diagnosticsState, "warn", "classic_item_skipped", { index: j + 1, reason: "Missing uploaded file label" });
            continue;
        }

        var widthIn = extractItemDimension(block, ["Width"]);
        var heightIn = extractItemDimension(block, ["Height", "Length"]);
        var fileName = trimStr(fm[1]);
        var qp = extractQtyAndPrice(block);
        if (isNaN(widthIn) || isNaN(heightIn) || !fileName) {
            addDiagnostic(diagnosticsState, "warn", "classic_item_skipped", { index: j + 1, reason: "Incomplete dimensions or filename", file: fileName });
            continue;
        }

        items.push({
            qty: qp.qty,
            width: widthIn,
            height: heightIn,
            file: fileName,
            productLabel: scopedMarkers[j].productLabel,
            printType: detectPrintType(scopedMarkers[j].productLabel),
            note: extractNoteForBlock(block),
            price: qp.price,
            currency: qp.currency,
            matchInfo: null
        });
    }
    return items;
}

function parseEmailItemsUS(emailText, diagnosticsState) {
    var text = String(emailText || "").replace(/\r/g, "");
    var markerRe = /(?:^|\n)\s*#(\d+)\s*(?=\n|$)/g;
    var markers = [];
    var markerMatch;
    while ((markerMatch = markerRe.exec(text)) !== null) {
        markers.push({ index: markerMatch.index, itemNo: parseInt(markerMatch[1], 10) });
    }
    if (!markers.length) return [];

    var items = [];
    for (var i = 0; i < markers.length; i++) {
        var blockStart = markers[i].index;
        if (text.charAt(blockStart) === "\n") blockStart++;
        var blockEnd = (i + 1 < markers.length) ? markers[i + 1].index : text.length;
        var block = text.substring(blockStart, blockEnd);
        var lines = block.split("\n");
        if (!lines.length) continue;

        var markerLine = trimStr(lines[0]);
        var productInfo = findNextNonEmptyLine(lines, 1);
        var fileLabelIdx = findLineIndexEquals(lines, "File Name", 0);
        var sizeLabelIdx = findLineIndexEquals(lines, "Custom Size", 0);
        var noteLabelIdx = findLineIndexEquals(lines, "Note for designers", 0);
        var fileInfo = fileLabelIdx >= 0 ? findNextNonEmptyLine(lines, fileLabelIdx + 1) : { index: -1, text: "" };
        var sizeInfo = sizeLabelIdx >= 0 ? findNextNonEmptyLine(lines, sizeLabelIdx + 1) : { index: -1, text: "" };
        var qty = 1;
        for (var q = lines.length - 1; q >= 0; q--) {
            var candidate = trimStr(lines[q]);
            if (/^\d+$/.test(candidate)) {
                qty = parseInt(candidate, 10);
                break;
            }
        }

        var note = "";
        if (noteLabelIdx >= 0) {
            var noteStop = lines.length;
            if (sizeLabelIdx >= 0 && sizeLabelIdx > noteLabelIdx) noteStop = Math.min(noteStop, sizeLabelIdx);
            var fileStopIdx = fileLabelIdx >= 0 && fileLabelIdx > noteLabelIdx ? fileLabelIdx : noteStop;
            if (fileStopIdx < noteStop) noteStop = fileStopIdx;
            for (var nl = noteLabelIdx + 1; nl < noteStop; nl++) {
                var noteCandidate = trimStr(lines[nl]);
                if (!noteCandidate) continue;
                note = noteCandidate;
                break;
            }
        }

        var dims = parseCustomSizeLine(sizeInfo.text);
        var productLabel = productInfo.text;
        var fileName = fileInfo.text;
        var nearbyText = block;

        if (!productLabel || !fileName || isNaN(dims.width) || isNaN(dims.height)) {
            addDiagnostic(diagnosticsState, "warn", "us_item_skipped", {
                marker: markerLine,
                product: productLabel,
                file: fileName,
                reason: "Malformed US item block"
            });
            continue;
        }

        items.push({
            qty: qty,
            width: dims.width,
            height: dims.height,
            file: fileName,
            productLabel: productLabel,
            printType: detectPrintType(productLabel + "\n" + nearbyText),
            note: note,
            price: NaN,
            currency: "$",
            matchInfo: null
        });
    }
    return items;
}

function parseEmailItems(emailText, diagnosticsState) {
    var text = String(emailText || "");
    var probablyUS = isProbablyUSOrderText(text);
    var items = [];
    var formatName = "classic";

    if (probablyUS) {
        items = parseEmailItemsUS(text, diagnosticsState);
        formatName = "us";
        if (items.length === 0) {
            addDiagnostic(diagnosticsState, "warn", "parser_fallback", { from: "us", to: "classic", reason: "US parser returned no items" });
            items = parseEmailItemsClassic(text, diagnosticsState);
            formatName = "classic";
        }
    } else {
        items = parseEmailItemsClassic(text, diagnosticsState);
        formatName = "classic";
        if (items.length === 0) {
            var usItems = parseEmailItemsUS(text, diagnosticsState);
            if (usItems.length > 0) {
                addDiagnostic(diagnosticsState, "info", "parser_fallback", { from: "classic", to: "us", reason: "Classic parser returned no items" });
                items = usItems;
                formatName = "us";
            }
        }
    }

    addDiagnostic(diagnosticsState, "info", "parser_selected", { format: formatName, items: items.length });
    return { items: items, formatName: formatName };
}

function extractNoteForBlock(blockText) {
    var m = /Message:\s*([\s\S]*?)(?=(?:\r?\n){2,}\s*(?:\d{1,4}\s*(?=(?:CAD|USD|EUR|GBP|\$|€|£))|Subtotal:|Total:)|$)/i.exec(String(blockText));
    if (!m || !m[1]) return "";
    return trimStr(String(m[1]).replace(/\r?\n+/g, " "));
}

function getFilesInFolder(folderObj) {
    return folderObj.getFiles(function (f) { return f instanceof File; });
}

function getRelativePath(fileObj, rootFolder) {
    var full = String(fileObj.fsName);
    var root = String(rootFolder.fsName);
    if (full.indexOf(root) === 0) {
        var rel = full.substring(root.length);
        if (rel.charAt(0) === "\\" || rel.charAt(0) === "/") rel = rel.substring(1);
        return rel.replace(/\\/g, "/");
    }
    return String(fileObj.name);
}

function toUrlPath(relPath) {
    var clean = String(relPath || "").replace(/\\/g, "/");
    try { return encodeURI(clean); } catch (e) { return clean; }
}

function ensureFolder(folderObj) {
    if (folderObj.exists) return true;
    try { return folderObj.create(); } catch (e) {}
    return folderObj.exists;
}

function getOutputFolderByPrintType(exportRoot, printTypeMode, printType) {
    if (printTypeMode !== "folder") return exportRoot;
    var bucket = printType ? printType : "Other";
    var destFolder = new Folder(exportRoot.fsName + "/" + bucket);
    ensureFolder(destFolder);
    return destFolder;
}

function runActionIfNeeded(enableAction, actionSetName, actionName) {
    if (!enableAction) return { ok: true, msg: "" };
    try {
        app.doAction(actionName, actionSetName);
        return { ok: true, msg: "" };
    } catch (e) {
        return { ok: false, msg: String(e) };
    }
}

function escHtml(s) {
    s = (s === null || s === undefined) ? "" : String(s);
    return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&#039;");
}

function escJsSingleQuoted(s) {
    s = (s === null || s === undefined) ? "" : String(s);
    return s.replace(/\\/g, "\\\\").replace(/'/g, "\\'").replace(/\r/g, "\\r").replace(/\n/g, "\\n");
}

function formatSigned(n) {
    var v = round2(n);
    return (v > 0 ? "+" : "") + v;
}

function formatSignedPercent(n) {
    return formatSigned(n) + "%";
}

function formatSize(w, h) {
    if (isNaN(w) || isNaN(h)) return "";
    return round2(w) + " x " + round2(h) + " in";
}

function percentDiff(actual, expected) {
    if (!expected) return 0;
    return ((actual - expected) / expected) * 100;
}

function getSeverityByPercent(absPct) {
    if (absPct > 10) return "not_ok";
    if (absPct >= 5) return "warn";
    return "ok";
}

function statusSortValue(status) {
    if (status === "MISSING_FILE") return 10;
    if (status === "QUEUED") return 15;
    if (status === "NOT OK") return 30;
    if (status === "CHECK") return 40;
    if (status === "OK") return 50;
    return 20;
}

function visualSortValue(direction) {
    if (direction === "grow") return 10;
    if (direction === "shrink") return 20;
    if (direction === "mixed") return 30;
    return 40;
}

function buildMeasuredVisualState(resizeMode, widthPct, heightPct) {
    var severity = "ok";
    var direction = "neutral";
    var eps = 0.0001;
    var hasPos = false;
    var hasNeg = false;

    if (resizeMode === "respectWidth") {
        severity = getSeverityByPercent(Math.abs(heightPct));
        if (heightPct > eps) direction = "grow";
        else if (heightPct < -eps) direction = "shrink";
    } else if (resizeMode === "respectHeight") {
        severity = getSeverityByPercent(Math.abs(widthPct));
        if (widthPct > eps) direction = "grow";
        else if (widthPct < -eps) direction = "shrink";
    } else {
        severity = worstSeverity(getSeverityByPercent(Math.abs(widthPct)), getSeverityByPercent(Math.abs(heightPct)));
        hasPos = widthPct > eps || heightPct > eps;
        hasNeg = widthPct < -eps || heightPct < -eps;
        if (hasPos && hasNeg) direction = "mixed";
        else if (hasPos) direction = "grow";
        else if (hasNeg) direction = "shrink";
    }

    var rowClass = "";
    if (severity === "warn") {
        if (direction === "grow") rowClass = "row-grow-warn";
        else if (direction === "shrink") rowClass = "row-shrink-warn";
        else if (direction === "mixed") rowClass = "row-mixed-warn";
    } else if (severity === "not_ok") {
        if (direction === "grow") rowClass = "row-grow-not-ok";
        else if (direction === "shrink") rowClass = "row-shrink-not-ok";
        else if (direction === "mixed") rowClass = "row-mixed-not-ok";
    }

    var primaryPct = resizeMode === "respectWidth" ? heightPct : (resizeMode === "respectHeight" ? widthPct : Math.max(Math.abs(widthPct), Math.abs(heightPct)));
    if (resizeMode === "stretch") primaryPct = hasPos && hasNeg ? Math.max(Math.abs(widthPct), Math.abs(heightPct)) : (hasPos ? Math.max(widthPct, heightPct) : Math.min(widthPct, heightPct));

    return { severity: severity, direction: direction, rowClass: rowClass, comparePct: primaryPct, visualSort: visualSortValue(direction) };
}

function severityRank(severity) {
    if (severity === "error") return 3;
    if (severity === "not_ok") return 2;
    if (severity === "warn") return 1;
    return 0;
}

function worstSeverity(a, b) {
    return severityRank(a) >= severityRank(b) ? a : b;
}

function makeMatchSummary(orderFile, matchInfo) {
    matchInfo = matchInfo || { file: null, matchType: "missing", suggested: "" };
    var label = "Missing";
    if (matchInfo.matchType === "exact") label = "Exact";
    else if (matchInfo.matchType === "normalized") label = "Normalized";
    else if (matchInfo.matchType === "canonical") label = "Canonical";
    else if (matchInfo.matchType === "ultraLoose") label = "Loose";

    if (matchInfo.file && matchInfo.file.name !== orderFile) return label + " -> " + matchInfo.file.name;
    if (!matchInfo.file && matchInfo.suggested) return label + " -> maybe: " + matchInfo.suggested;
    return label;
}

function makeMeasuredRow(emailFileName, qty, printType, note, price, currency, matchInfo, resizeMode, orderW, orderH, outW, outH, thumbPath, outputFsPath) {
    var widthDiff = outW - orderW;
    var heightDiff = outH - orderH;
    var widthPct = percentDiff(outW, orderW);
    var heightPct = percentDiff(outH, orderH);
    var delta = "";
    var visual = buildMeasuredVisualState(resizeMode, widthPct, heightPct);

    if (resizeMode === "respectWidth") {
        delta = "H " + formatSigned(heightDiff) + " in (" + formatSignedPercent(heightPct) + ")";
    } else if (resizeMode === "respectHeight") {
        delta = "W " + formatSigned(widthDiff) + " in (" + formatSignedPercent(widthPct) + ")";
    } else {
        delta = "W " + formatSigned(widthDiff) + " in (" + formatSignedPercent(widthPct) + ") | H " + formatSigned(heightDiff) + " in (" + formatSignedPercent(heightPct) + ")";
    }

    return {
        file: emailFileName,
        qty: qty,
        printType: printType || "",
        note: note || "",
        price: isNaN(price) ? NaN : roundMoney(price),
        currency: normalizeCurrencyLabel(currency),
        match: makeMatchSummary(emailFileName, matchInfo),
        orderW: round2(orderW),
        orderH: round2(orderH),
        orderSize: formatSize(orderW, orderH),
        outputSize: formatSize(outW, outH),
        outputW: round2(outW),
        outputH: round2(outH),
        thumbPath: thumbPath || "",
        outputFsPath: outputFsPath || "",
        delta: delta,
        status: visual.severity === "not_ok" ? "NOT OK" : (visual.severity === "warn" ? "CHECK" : "OK"),
        rowClass: visual.rowClass,
        statusSort: statusSortValue(visual.severity === "not_ok" ? "NOT OK" : (visual.severity === "warn" ? "CHECK" : "OK")),
        visualSort: visual.visualSort,
        deltaSort: Math.abs(isNaN(visual.comparePct) ? 0 : visual.comparePct),
        printSort: (printType || "").toLowerCase()
    };
}

function makeStatusRow(emailFileName, qty, printType, note, price, currency, matchInfo, orderW, orderH, statusCode) {
    return {
        file: emailFileName,
        qty: qty,
        printType: printType || "",
        note: note || "",
        price: isNaN(price) ? NaN : roundMoney(price),
        currency: normalizeCurrencyLabel(currency),
        match: makeMatchSummary(emailFileName, matchInfo),
        orderW: round2(orderW),
        orderH: round2(orderH),
        orderSize: formatSize(orderW, orderH),
        outputSize: "",
        outputW: "",
        outputH: "",
        thumbPath: "",
        outputFsPath: "",
        delta: "",
        status: statusCode,
        rowClass: "row-error",
        statusSort: statusSortValue(statusCode),
        visualSort: 90,
        deltaSort: 999999,
        printSort: (printType || "").toLowerCase()
    };
}

function makeQueuedRow(emailFileName, qty, printType, note, price, currency, matchInfo, orderW, orderH) {
    return {
        file: emailFileName,
        qty: qty,
        printType: printType || "",
        note: note || "",
        price: isNaN(price) ? NaN : roundMoney(price),
        currency: normalizeCurrencyLabel(currency),
        match: makeMatchSummary(emailFileName, matchInfo),
        orderW: round2(orderW),
        orderH: round2(orderH),
        orderSize: formatSize(orderW, orderH),
        outputSize: "",
        outputW: "",
        outputH: "",
        thumbPath: "",
        outputFsPath: "",
        delta: "",
        status: "QUEUED",
        rowClass: "row-pending",
        statusSort: statusSortValue("QUEUED"),
        visualSort: 80,
        deltaSort: -1,
        printSort: (printType || "").toLowerCase()
    };
}

function buildInitialReportRows(items) {
    var rows = [];
    for (var i = 0; i < items.length; i++) {
        var item = items[i];
        var matchInfo = item.matchInfo || { file: null, matchType: "missing", suggested: "" };
        if (isNaN(item.width) || isNaN(item.height)) {
            rows.push(makeStatusRow(item.file, item.qty, item.printType, item.note, item.price, item.currency, matchInfo, item.width, item.height, "BAD_WIDTH_HEIGHT"));
        } else if (!matchInfo.file || !matchInfo.file.exists) {
            rows.push(makeStatusRow(item.file, item.qty, item.printType, item.note, item.price, item.currency, matchInfo, item.width, item.height, "MISSING_FILE"));
        } else {
            rows.push(makeQueuedRow(item.file, item.qty, item.printType, item.note, item.price, item.currency, matchInfo, item.width, item.height));
        }
    }
    return rows;
}

function buildReportStats(reportRows) {
    var stats = { exported: 0, check: 0, notOk: 0, errors: 0, queued: 0 };
    for (var i = 0; i < reportRows.length; i++) {
        var row = reportRows[i];
        if (row.status === "OK") stats.exported++;
        else if (row.status === "CHECK") { stats.exported++; stats.check++; }
        else if (row.status === "NOT OK") { stats.exported++; stats.notOk++; }
        else if (row.status === "QUEUED") stats.queued++;
        else stats.errors++;
    }
    return stats;
}

function buildReportHtml(reportMeta, reportRows) {
    var stats = buildReportStats(reportRows);
    var html = [];
    html.push("<!doctype html>");
    html.push("<html><head><meta charset='utf-8'><title>DTF QC Report</title>");
    html.push("<style>");
    html.push("body{font-family:Arial,Helvetica,sans-serif;font-size:13px;padding:18px;color:#1f1f1f;background:#fbfbfb;}.page{max-width:1680px;margin:0 auto;}h1{font-size:22px;margin:0 0 10px 0;}.meta,.summary{margin-bottom:12px;line-height:1.6;}.summary strong{display:inline-block;min-width:84px;}.table-wrap{max-height:86vh;overflow:auto;border:1px solid #d7d7d7;background:#fff;box-shadow:0 1px 2px rgba(0,0,0,.03);}table{border-collapse:collapse;width:100%;background:#fff;min-width:1180px;}th,td{border:1px solid #d7d7d7;padding:7px 9px;vertical-align:top;text-align:left;}th{background:#f3f3f3;position:sticky;top:0;z-index:4;cursor:pointer;white-space:nowrap;}th .sort-label{display:inline-flex;align-items:center;gap:6px;}th .sort-ind{font-size:10px;color:#666;}tr:nth-child(even) td{background:#fafafa;}.row-grow-warn td{background:#ffdede !important;}.row-grow-not-ok td{background:#ffbcbc !important;}.row-shrink-warn td{background:#ffe7cf !important;}.row-shrink-not-ok td{background:#ffc999 !important;}.row-mixed-warn td{background:#fff6b8 !important;}.row-mixed-not-ok td{background:#ffe17a !important;}.row-error td{background:#ffe7e7 !important;}.row-pending td{background:#eef2f6 !important;}.review-muted td{background:#fff !important;}.mono{font-family:Consolas,Monaco,monospace;}.thumb-link{text-decoration:none;cursor:pointer;color:inherit;}.thumb-wrap{display:inline-flex;align-items:center;justify-content:center;width:84px;height:84px;border:1px solid #cfcfcf;background:linear-gradient(135deg,#e6c1ca 0%,#f0d7dd 48%,#f8f1f3 100%);box-shadow:inset 0 0 0 1px rgba(255,255,255,.35);overflow:hidden;}.thumb-box{display:flex;align-items:center;justify-content:center;width:80px;height:80px;overflow:hidden;}.thumb-img{border:0;display:block;}.note-cell{min-width:180px;max-width:280px;}.review-cell{white-space:nowrap;text-align:center;}.legend{margin-bottom:12px;font-size:12px;color:#555;line-height:1.7;} .legend span{display:inline-block;margin-right:14px;padding:2px 8px;border-radius:10px;} .lg-rw{background:#ffdede;} .lg-rn{background:#ffbcbc;} .lg-sw{background:#ffe7cf;} .lg-sn{background:#ffc999;} .lg-mw{background:#fff6b8;} .lg-er{background:#ffe7e7;} .lg-pd{background:#eef2f6;}");
    html.push("</style>");
    html.push("<script language='javascript'>");
    html.push("function fitThumb(img,maxW,maxH){try{var w=img.width||img.offsetWidth||1;var h=img.height||img.offsetHeight||1;if(!w||!h)return;var r=Math.min(maxW/w,maxH/h);img.width=Math.max(1,Math.round(w*r));img.height=Math.max(1,Math.round(h*r));}catch(e){}}");
    html.push("var sortState={key:'',asc:true};function statusSortValue(s){if(s==='MISSING_FILE')return 10;if(s==='QUEUED')return 15;if(s==='NOT OK')return 30;if(s==='CHECK')return 40;if(s==='OK')return 50;return 20;}function sortReport(key){var tbody=document.getElementById('report-body');var rows=[];for(var i=0;i<tbody.rows.length;i++)rows.push(tbody.rows[i]);if(sortState.key===key)sortState.asc=!sortState.asc;else{sortState.key=key;sortState.asc=true;}rows.sort(function(a,b){var av='',bv='',ac=0,bc=0;if(key==='status'){av=parseFloat(a.getAttribute('data-status-sort')||'999');bv=parseFloat(b.getAttribute('data-status-sort')||'999');if(av!==bv)return sortState.asc?(av-bv):(bv-av);ac=parseFloat(a.getAttribute('data-visual-sort')||'999');bc=parseFloat(b.getAttribute('data-visual-sort')||'999');if(ac!==bc)return sortState.asc?(ac-bc):(bc-ac);}else if(key==='print'){av=(a.getAttribute('data-print-sort')||'').toLowerCase();bv=(b.getAttribute('data-print-sort')||'').toLowerCase();}else if(key==='delta'){av=parseFloat(a.getAttribute('data-delta-sort')||'0');bv=parseFloat(b.getAttribute('data-delta-sort')||'0');}else if(key==='qty'){av=parseFloat(a.getAttribute('data-qty-sort')||'0');bv=parseFloat(b.getAttribute('data-qty-sort')||'0');}else if(key==='price'){av=parseFloat(a.getAttribute('data-price-sort')||'0');bv=parseFloat(b.getAttribute('data-price-sort')||'0');}else if(key==='file'){av=(a.getAttribute('data-file-sort')||'').toLowerCase();bv=(b.getAttribute('data-file-sort')||'').toLowerCase();}else return 0;if(av<bv)return sortState.asc?-1:1;if(av>bv)return sortState.asc?1:-1;var ai=parseInt(a.getAttribute('data-row-index')||'0',10);var bi=parseInt(b.getAttribute('data-row-index')||'0',10);return ai-bi;});for(var j=0;j<rows.length;j++){tbody.appendChild(rows[j]);rows[j].cells[0].innerHTML=String(j+1);}var headers=document.getElementsByTagName('th');for(var h=0;h<headers.length;h++){var k=headers[h].getAttribute('data-key');var ind=headers[h].getElementsByClassName('sort-ind');if(ind&&ind[0])ind[0].innerHTML=(k===sortState.key?(sortState.asc?'&uarr;':'&darr;'):'');}}function toggleReviewed(cb){var tr=cb;while(tr&&tr.tagName!=='TR')tr=tr.parentNode;if(!tr)return;if(cb.checked)tr.className+=' review-muted';else tr.className=tr.className.replace(/\\breview-muted\\b/g,'').replace(/\\s+/g,' ').replace(/^\\s+|\\s+$/g,'');}");
    html.push("</script></head><body><div class='page'>");
    html.push("<h1>DTF Export Report</h1>");
    html.push("<div class='meta'>");
    html.push("<div><strong>App:</strong> " + escHtml(reportMeta.appName) + "</div>");
    html.push("<div><strong>Date:</strong> " + escHtml(reportMeta.date) + "</div>");
    html.push("<div><strong>Mode:</strong> " + escHtml(reportMeta.resizeMode) + "</div>");
    html.push("<div><strong>DPI:</strong> " + escHtml(reportMeta.dpi) + "</div>");
    html.push("<div><strong>Naming:</strong> " + escHtml(reportMeta.filenameFormat) + "</div>");
    html.push("<div><strong>Print Sort:</strong> " + escHtml(reportMeta.printTypeMode) + "</div>");
    html.push("<div><strong>Action:</strong> " + escHtml(reportMeta.actionSummary) + "</div>");
    html.push("<div><strong>Folder:</strong> " + escHtml(reportMeta.exportFolder) + "</div>");
    html.push("</div><div class='summary'>");
    html.push("<div><strong>Items:</strong> " + escHtml(reportMeta.itemsFound) + "</div>");
    html.push("<div><strong>Exported:</strong> " + escHtml(stats.exported) + "</div>");
    html.push("<div><strong>Check:</strong> " + escHtml(stats.check) + "</div>");
    html.push("<div><strong>Not OK:</strong> " + escHtml(stats.notOk) + "</div>");
    html.push("<div><strong>Queued:</strong> " + escHtml(stats.queued) + "</div>");
    html.push("<div><strong>Errors:</strong> " + escHtml(stats.errors) + "</div></div>");
    html.push("<div class='legend'><span class='lg-rw'>Bigger 5-10%</span><span class='lg-rn'>Bigger 10%+</span><span class='lg-sw'>Smaller 5-10%</span><span class='lg-sn'>Smaller 10%+</span><span class='lg-mw'>Mixed Stretch</span><span class='lg-er'>Error / Missing</span><span class='lg-pd'>Queued / Not Reached</span></div>");
    html.push("<div class='table-wrap'><table><thead><tr><th data-key='row'><span class='sort-label'>#<span class='sort-ind'></span></span></th><th data-key='thumb'><span class='sort-label'>Thumb<span class='sort-ind'></span></span></th><th data-key='file' onclick=\"sortReport('file')\"><span class='sort-label'>File<span class='sort-ind'></span></span></th><th data-key='qty' onclick=\"sortReport('qty')\"><span class='sort-label'>Qty<span class='sort-ind'></span></span></th><th data-key='print' onclick=\"sortReport('print')\"><span class='sort-label'>Print<span class='sort-ind'></span></span></th><th data-key='price' onclick=\"sortReport('price')\"><span class='sort-label'>Price<span class='sort-ind'></span></span></th><th data-key='note'><span class='sort-label'>Note<span class='sort-ind'></span></span></th><th data-key='match'><span class='sort-label'>Match<span class='sort-ind'></span></span></th><th data-key='order'><span class='sort-label'>Order<span class='sort-ind'></span></span></th><th data-key='output'><span class='sort-label'>Output<span class='sort-ind'></span></span></th><th data-key='delta' onclick=\"sortReport('delta')\"><span class='sort-label'>Delta<span class='sort-ind'></span></span></th><th data-key='status' onclick=\"sortReport('status')\"><span class='sort-label'>Status<span class='sort-ind'></span></span></th><th data-key='review'><span class='sort-label'>Reviewed<span class='sort-ind'></span></span></th></tr></thead><tbody id='report-body'>");
    for (var i = 0; i < reportRows.length; i++) {
        var row = reportRows[i];
        var thumbHtml = "";
        var fileHtml = escHtml(row.file);
        if (row.thumbPath) {
            var url = toUrlPath(row.thumbPath);
            var outputUrl = row.outputFsPath ? toUrlPath(row.outputFsPath) : url;
            thumbHtml = "<a class='thumb-link' href='" + escHtml(outputUrl) + "' target='_blank'><span class='thumb-wrap'><span class='thumb-box'><img class='thumb-img' src='" + escHtml(url) + "' alt='thumb' onload='fitThumb(this,80,80)'></span></span></a>";
            fileHtml = "<a href='" + escHtml(outputUrl) + "' target='_blank'>" + escHtml(row.file) + "</a>";
        }
        html.push("<tr class='" + escHtml(row.rowClass) + "' data-row-index='" + escHtml(i) + "' data-file-sort='" + escHtml((row.file || '').toLowerCase()) + "' data-qty-sort='" + escHtml(row.qty) + "' data-print-sort='" + escHtml(row.printSort || '') + "' data-price-sort='" + escHtml(isNaN(row.price) ? -1 : row.price) + "' data-delta-sort='" + escHtml(row.deltaSort || 0) + "' data-status-sort='" + escHtml(row.statusSort || 999) + "' data-visual-sort='" + escHtml(row.visualSort || 0) + "'><td>" + escHtml(i + 1) + "</td><td>" + thumbHtml + "</td><td class='mono'>" + fileHtml + "</td><td>" + escHtml(row.qty) + "</td><td>" + escHtml(row.printType) + "</td><td>" + escHtml(formatMoney(row.price, row.currency)) + "</td><td class='note-cell'>" + escHtml(row.note) + "</td><td>" + escHtml(row.match) + "</td><td>" + escHtml(row.orderSize) + "</td><td>" + escHtml(row.outputSize) + "</td><td>" + escHtml(row.delta) + "</td><td>" + escHtml(row.status) + "</td><td class='review-cell'><input type='checkbox' onclick='toggleReviewed(this)'></td></tr>");
    }
    html.push("</tbody></table></div></div></body></html>");
    return html.join("\r\n");
}

function buildProofHtml(reportMeta, reportRows) {
    var html = [];
    html.push("<!doctype html>");
    html.push("<html><head><meta charset='utf-8'><title>Customer Proof</title><style>");
    html.push("body{font-family:Arial,Helvetica,sans-serif;margin:0;padding:24px;background:#f6f4ef;color:#1f1f1f;}h1{margin:0 0 8px 0;font-size:26px;}p.meta{margin:0 0 18px 0;color:#555;} .grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(360px,1fr));gap:18px;} .card{background:#fff;border:1px solid #ddd;padding:16px;box-shadow:0 1px 3px rgba(0,0,0,.06);break-inside:avoid;overflow:hidden;} .name{font-family:Consolas,Monaco,monospace;font-size:16px;font-weight:bold;margin-bottom:12px;word-break:break-word;} .proof-wrap{display:flex;justify-content:center;margin-top:8px;overflow:hidden;} .proof-frame{position:relative;display:inline-block;padding-left:44px;padding-bottom:36px;max-width:100%;box-sizing:border-box;} .stage{display:inline-flex;align-items:center;justify-content:center;border:1px solid #d4d4d4;cursor:pointer;line-height:0;overflow:hidden;max-width:320px;max-height:320px;box-sizing:border-box;} .stage img{display:block;max-width:320px;max-height:320px;width:auto;height:auto;vertical-align:bottom;} .grid-light{background-color:#f7f7f7;background-image:linear-gradient(45deg,#ececec 25%,transparent 25%),linear-gradient(-45deg,#ececec 25%,transparent 25%),linear-gradient(45deg,transparent 75%,#ececec 75%),linear-gradient(-45deg,transparent 75%,#ececec 75%);} .grid-medium{background-color:#ececec;background-image:linear-gradient(45deg,#cfcfcf 25%,transparent 25%),linear-gradient(-45deg,#cfcfcf 25%,transparent 25%),linear-gradient(45deg,transparent 75%,#cfcfcf 75%),linear-gradient(-45deg,transparent 75%,#cfcfcf 75%);} .grid-dark{background-color:#c8c8c8;background-image:linear-gradient(45deg,#8f8f8f 25%,transparent 25%),linear-gradient(-45deg,#8f8f8f 25%,transparent 25%),linear-gradient(45deg,transparent 75%,#8f8f8f 75%),linear-gradient(-45deg,transparent 75%,#8f8f8f 75%);} .grid-light,.grid-medium,.grid-dark{background-size:24px 24px;background-position:0 0,0 12px,12px -12px,-12px 0;} .width-guide{position:absolute;left:44px;right:0;bottom:0;text-align:center;font-size:20px;font-weight:bold;line-height:1.1;} .height-guide{position:absolute;left:0;top:0;bottom:36px;width:34px;display:flex;align-items:center;justify-content:center;} .height-guide span{display:inline-block;writing-mode:vertical-rl;transform:rotate(180deg);font-size:20px;font-weight:bold;line-height:1.1;} .empty{padding:24px;background:#fff;border:1px dashed #ccc;}");
    html.push("</style><script>");
    html.push("function cycleGrid(el){var cls=['grid-light','grid-medium','grid-dark'];var idx=parseInt(el.getAttribute('data-grid-index')||'0',10);el.className=el.className.replace(/grid-light|grid-medium|grid-dark/g,'').replace(/\\s+/g,' ').replace(/^\\s+|\\s+$/g,'');idx=(idx+1)%cls.length;el.className+=(el.className?' ':'')+cls[idx];el.setAttribute('data-grid-index',String(idx));}");
    html.push("</script></head><body>");
    html.push("<h1>Customer Proof</h1>");
    html.push("<p class='meta'>App: " + escHtml(reportMeta.appName) + " | Date: " + escHtml(reportMeta.date) + "</p>");
    html.push("<div class='grid'>");
    var count = 0;
    for (var i = 0; i < reportRows.length; i++) {
        var row = reportRows[i];
        if (!row.thumbPath || row.status === "BAD_WIDTH_HEIGHT" || row.status === "MISSING_FILE" || row.status === "OPEN_FAIL" || row.status === "ACTION_FAIL" || row.status === "PROCESS_ERROR") continue;
        var proofUrl = toUrlPath(row.thumbPath);
        html.push("<section class='card'><div class='name'>" + escHtml(row.file) + "</div><div class='proof-wrap'><div class='proof-frame'><div class='height-guide'><span>" + escHtml(row.outputH) + " in</span></div><div class='stage grid-medium' data-grid-index='1' onclick='cycleGrid(this)'><img src='" + escHtml(proofUrl) + "' alt='proof'></div><div class='width-guide'>" + escHtml(row.outputW) + " in</div></div></div></section>");
        count++;
    }
    if (!count) html.push("<div class='empty'>No exported images available for proof.</div>");
    html.push("</div></body></html>");
    return html.join("\r\n");
}

function buildPricingAuditHtml(reportMeta, reportRows) {
    var html = [];
    html.push("<!doctype html>");
    html.push("<html><head><meta charset='utf-8'><title>Pricing Audit</title><style>");
    html.push("body{font-family:Arial,Helvetica,sans-serif;margin:0;padding:20px;background:#f5f5f5;color:#1f1f1f;}h1{margin:0 0 6px 0;font-size:24px;}p.meta{margin:0 0 8px 0;color:#555;line-height:1.5;}p.help{margin:0 0 14px 0;color:#555;font-size:12px;} .summary{margin:0 0 12px 0;background:#fff;border:1px solid #ddd;} .summary table{border-collapse:collapse;width:100%;} .summary td{padding:8px 10px;border-right:1px solid #e6e6e6;vertical-align:top;} .summary td:last-child{border-right:0;} .sum-label{display:block;font-size:11px;color:#666;margin-bottom:3px;text-transform:uppercase;letter-spacing:.03em;} .sum-value{display:block;font-size:16px;font-weight:bold;} .sum-meta{margin:6px 10px 10px 10px;font-size:11px;color:#666;} table{border-collapse:collapse;width:100%;background:#fff;} th,td{border:1px solid #d7d7d7;padding:7px 8px;vertical-align:top;text-align:left;font-size:12px;} th{background:#f3f3f3;} tr:nth-child(even) td{background:#fafafa;} .mono{font-family:Consolas,Monaco,monospace;} .num-input{width:64px;padding:4px 5px;font:inherit;} .money-pos{color:#9a1b1b;font-weight:bold;} .money-neg{color:#0c6b35;font-weight:bold;} .money-zero{color:#333;} .small{font-size:11px;color:#666;margin-top:3px;} .right{text-align:right;} .qty-old{white-space:nowrap;} ");
    html.push("</style><script>");
    html.push("function roundMoney(n){return Math.round((n + 0.0000001) * 100) / 100;}");
    html.push("function formatMoney(n,cur){if(isNaN(n))return '--';var sign=n<0?'-':'';var s=Math.abs(roundMoney(n)).toFixed(2);if(cur==='$'||cur==='€'||cur==='£')return sign+cur+s;return sign+cur+' '+s;}");
    html.push("function parseNum(v){var n=parseFloat(v);return isNaN(n)?NaN:n;}");
    html.push("var sortTimer=null;function sortRows(){var tbody=document.getElementById('audit-body');var rows=[];for(var i=0;i<tbody.rows.length;i++)rows.push(tbody.rows[i]);rows.sort(function(a,b){var ad=parseFloat(a.getAttribute('data-diff-sort')||'0');var bd=parseFloat(b.getAttribute('data-diff-sort')||'0');if(bd!==ad)return bd-ad;var ai=parseInt(a.getAttribute('data-sort-index')||'0',10);var bi=parseInt(b.getAttribute('data-sort-index')||'0',10);return ai-bi;});for(var j=0;j<rows.length;j++){tbody.appendChild(rows[j]);rows[j].cells[0].innerHTML=String(j+1);}}");
    html.push("function recalcRow(row){var ow=parseNum(row.getAttribute('data-order-w'));var oh=parseNum(row.getAttribute('data-order-h'));var oq=parseNum(row.getAttribute('data-order-qty'));var price=parseNum(row.getAttribute('data-price'));var cur=row.getAttribute('data-currency')||'$';var aw=parseNum(row.getElementsByClassName('adj-w')[0].value);var ah=parseNum(row.getElementsByClassName('adj-h')[0].value);var nq=parseNum(row.getElementsByClassName('adj-q')[0].value);var adjusted=NaN;var diff=NaN;if(!isNaN(ow)&&!isNaN(oh)&&!isNaN(oq)&&!isNaN(price)&&!isNaN(aw)&&!isNaN(ah)&&!isNaN(nq)&&ow>0&&oh>0&&oq>0&&aw>=0&&ah>=0&&nq>=0){adjusted=roundMoney(price*((aw*ah*nq)/(ow*oh*oq)));diff=roundMoney(adjusted-price);}row.setAttribute('data-adjusted-price',isNaN(adjusted)?'':String(adjusted));row.setAttribute('data-diff',isNaN(diff)?'':String(diff));row.setAttribute('data-diff-sort',isNaN(diff)?'-1':String(Math.abs(diff)));var adjCell=row.getElementsByClassName('adjusted-price')[0];var diffCell=row.getElementsByClassName('price-diff')[0];adjCell.innerHTML=formatMoney(adjusted,cur);diffCell.innerHTML=formatMoney(diff,cur);diffCell.className='price-diff';if(isNaN(diff))diffCell.className+=' money-zero';else if(diff>0.004)diffCell.className+=' money-pos';else if(diff<-0.004)diffCell.className+=' money-neg';else diffCell.className+=' money-zero';}");
    html.push("function updateSummary(){var rows=document.getElementById('audit-body').rows;var ordered=0,adjusted=0,calc=0;for(var i=0;i<rows.length;i++){var price=parseNum(rows[i].getAttribute('data-price'));var adj=parseNum(rows[i].getAttribute('data-adjusted-price'));if(!isNaN(price))ordered+=price;if(!isNaN(adj)){adjusted+=adj;calc++;}}var cur=document.body.getAttribute('data-currency')||'$';var taxRate=parseNum(document.body.getAttribute('data-tax-rate'));if(isNaN(taxRate))taxRate=0;ordered=roundMoney(ordered);adjusted=roundMoney(adjusted);var diff=roundMoney(adjusted-ordered);var diffTax=roundMoney(diff*taxRate/100);var due=roundMoney(diff+diffTax);document.getElementById('sum-price-diff').innerHTML=formatMoney(diff,cur);document.getElementById('sum-tax').innerHTML=formatMoney(diffTax,cur);document.getElementById('sum-due').innerHTML=formatMoney(due,cur);document.getElementById('sum-meta').innerHTML='Old subtotal: '+formatMoney(ordered,cur)+' | New subtotal: '+formatMoney(adjusted,cur)+' | Recalculated rows: '+calc+' / '+rows.length;}");
    html.push("function recalcAll(noSort){var rows=document.getElementById('audit-body').rows;for(var i=0;i<rows.length;i++)recalcRow(rows[i]);if(!noSort)sortRows();updateSummary();}");
    html.push("function scheduleSort(){if(sortTimer)clearTimeout(sortTimer);sortTimer=setTimeout(function(){recalcAll(false);},700);}");
    html.push("function initAudit(){var inputs=document.getElementsByTagName('input');for(var i=0;i<inputs.length;i++){if(inputs[i].className.indexOf('num-input')>=0){inputs[i].oninput=function(){recalcAll(true);scheduleSort();};inputs[i].onchange=function(){if(sortTimer)clearTimeout(sortTimer);recalcAll(false);};inputs[i].onblur=function(){if(sortTimer)clearTimeout(sortTimer);recalcAll(false);};}}recalcAll(false);}");
    html.push("</script></head><body data-currency='" + escHtml(reportMeta.currency || "$") + "' data-tax-rate='" + escHtml(isNaN(reportMeta.taxRate) ? "" : reportMeta.taxRate) + "' onload='initAudit()'>");
    html.push("<h1>Pricing Audit</h1>");
    html.push("<p class='meta'>App: " + escHtml(reportMeta.appName) + " | Date: " + escHtml(reportMeta.date) + " | Folder: " + escHtml(reportMeta.exportFolder) + "</p>");
    html.push("<p class='help'>Adjusted width/height defaults to the exported size when available. If a file was missing or failed, the inputs start from the order size so you can enter a manual value.</p>");
    html.push("<div class='summary'>");
    html.push("<table><tr><td><span class='sum-label'>Price Adjustment</span><span class='sum-value' id='sum-price-diff'>--</span></td><td><span class='sum-label'>HST On Adjustment</span><span class='sum-value' id='sum-tax'>--</span></td><td><span class='sum-label'>Amount Due</span><span class='sum-value' id='sum-due'>--</span></td></tr></table>");
    html.push("<div class='sum-meta' id='sum-meta'>--</div>");
    html.push("</div>");
    html.push("<table><thead><tr><th>#</th><th>Print</th><th>File</th><th>W Old</th><th>H Old</th><th>Qty Old</th><th>Detected</th><th>W New</th><th>H New</th><th>Qty New</th><th>Current Price</th><th>Adjusted Price</th><th>Diff</th></tr></thead><tbody id='audit-body'>");
    for (var i = 0; i < reportRows.length; i++) {
        var row = reportRows[i];
        var defaultW = row.outputW !== "" && !isNaN(row.outputW) ? row.outputW : (!isNaN(row.orderW) ? row.orderW : "");
        var defaultH = row.outputH !== "" && !isNaN(row.outputH) ? row.outputH : (!isNaN(row.orderH) ? row.orderH : "");
        var detected = (row.outputW !== "" && row.outputH !== "") ? formatSize(row.outputW, row.outputH) : "--";
        html.push("<tr data-sort-index='" + escHtml(i) + "' data-order-w='" + escHtml(row.orderW) + "' data-order-h='" + escHtml(row.orderH) + "' data-order-qty='" + escHtml(row.qty) + "' data-price='" + escHtml(isNaN(row.price) ? "" : roundMoney(row.price)) + "' data-currency='" + escHtml(row.currency || reportMeta.currency || "$") + "'>");
        html.push("<td>" + escHtml(i + 1) + "</td><td>" + escHtml(row.printType) + "</td><td class='mono'>" + escHtml(row.file) + "<div class='small'>" + escHtml(row.note || "") + "</div></td><td>" + escHtml(row.orderW) + "</td><td>" + escHtml(row.orderH) + "</td><td class='qty-old'>" + escHtml(row.qty) + "</td><td>" + escHtml(detected) + "</td><td><input class='num-input adj-w' type='number' step='0.01' value='" + escHtml(defaultW) + "'></td><td><input class='num-input adj-h' type='number' step='0.01' value='" + escHtml(defaultH) + "'></td><td><input class='num-input adj-q' type='number' step='1' value='" + escHtml(row.qty) + "'></td><td class='right'>" + escHtml(formatMoney(row.price, row.currency || reportMeta.currency)) + "</td><td class='adjusted-price right'>--</td><td class='price-diff right money-zero'>--</td>");
        html.push("</tr>");
    }
    html.push("</tbody></table></body></html>");
    return html.join("\r\n");
}

function buildLogHeader(reportMeta) {
    var lines = [];
    lines.push("DTF Export Log");
    lines.push("App: " + reportMeta.appName);
    lines.push("Date: " + reportMeta.date);
    lines.push("Mode: " + reportMeta.resizeMode);
    lines.push("DPI: " + reportMeta.dpi);
    lines.push("Naming: " + reportMeta.filenameFormat);
    lines.push("Print Sort: " + reportMeta.printTypeMode);
    lines.push("Action: " + reportMeta.actionSummary);
    lines.push("Folder: " + reportMeta.exportFolder);
    lines.push("Items: " + reportMeta.itemsFound);
    lines.push("");
    lines.push("Index | Status | Print | File | Qty | Order | Output | Delta | Match | Note");
    lines.push("----- | ------ | ----- | ---- | --- | ----- | ------ | ----- | ----- | ----");
    return lines.join("\r\n") + "\r\n";
}

function formatLogRow(index, row) {
    return [
        index,
        row.status || "",
        row.printType || "",
        row.file || "",
        row.qty || "",
        row.orderSize || "",
        row.outputSize || "",
        row.delta || "",
        row.match || "",
        row.note || ""
    ].join(" | ") + "\r\n";
}

function buildLogSummary(reportMeta, reportRows) {
    var stats = buildReportStats(reportRows);
    return "\r\nSummary | Exported: " + stats.exported + " | Check: " + stats.check + " | Not OK: " + stats.notOk + " | Queued: " + stats.queued + " | Errors: " + stats.errors + "\r\n";
}

function buildFinalLogText(logBufferParts, reportMeta, reportRows) {
    var parts = logBufferParts ? logBufferParts.slice(0) : [];
    for (var i = 0; i < reportRows.length; i++) {
        parts.push(formatLogRow(i + 1, reportRows[i]));
    }
    parts.push(buildLogSummary(reportMeta, reportRows));
    return parts.join("");
}

function writeTextFile(fileObj, text) {
    var lastError = "";
    for (var attempt = 0; attempt < FILE_WRITE_RETRY_COUNT; attempt++) {
        var opened = false;
        try {
            fileObj.encoding = "UTF-8";
            if (!fileObj.open("w")) {
                lastError = "open failed: " + fileObj.error;
            } else {
                opened = true;
                if (!fileObj.write(text)) {
                    lastError = "write failed: " + fileObj.error;
                    try { fileObj.close(); } catch (eWriteClose) {}
                    opened = false;
                } else if (!fileObj.close()) {
                    lastError = "close failed: " + fileObj.error;
                    opened = false;
                } else {
                    return { ok: true, error: "" };
                }
            }
        } catch (e) {
            lastError = String(e);
            if (opened) { try { fileObj.close(); } catch (eClose) {} }
        }
        if (attempt < FILE_WRITE_RETRY_COUNT - 1) sleepMs(FILE_WRITE_RETRY_DELAY_MS);
    }
    return { ok: false, error: lastError || "write failed" };
}

function appendTextFile(fileObj, text) {
    var lastError = "";
    for (var attempt = 0; attempt < FILE_WRITE_RETRY_COUNT; attempt++) {
        var opened = false;
        try {
            fileObj.encoding = "UTF-8";
            if (!fileObj.open("a")) {
                lastError = "open failed: " + fileObj.error;
            } else {
                opened = true;
                if (!fileObj.write(text)) {
                    lastError = "write failed: " + fileObj.error;
                    try { fileObj.close(); } catch (eWriteClose) {}
                    opened = false;
                } else if (!fileObj.close()) {
                    lastError = "close failed: " + fileObj.error;
                    opened = false;
                } else {
                    return { ok: true, error: "" };
                }
            }
        } catch (e) {
            lastError = String(e);
            if (opened) { try { fileObj.close(); } catch (eClose) {} }
        }
        if (attempt < FILE_WRITE_RETRY_COUNT - 1) sleepMs(FILE_WRITE_RETRY_DELAY_MS);
    }
    return { ok: false, error: lastError || "append failed" };
}

function writeManagedTextFile(fileObj, text) {
    var primary = writeTextFile(fileObj, text);
    if (primary.ok) return { ok: true, error: "", warning: "", path: fileObj.fsName };

    var baseName = stripExt(fileObj.name);
    var extIndex = fileObj.name.lastIndexOf(".");
    var ext = (extIndex >= 0) ? fileObj.name.substring(extIndex) : "";
    var stamp = makeTimestampTag();
    var sameFolderFile = new File(fileObj.parent.fsName + "/" + baseName + "__" + stamp + ext);
    var sameFolder = writeTextFile(sameFolderFile, text);
    if (sameFolder.ok) {
        return { ok: true, error: "", warning: "Primary path unavailable (" + primary.error + "). Wrote fallback file instead.", path: sameFolderFile.fsName };
    }

    var tempDir = new Folder(Folder.temp.fsName + "/Sizer_Reports");
    ensureFolder(tempDir);
    var tempFile = new File(tempDir.fsName + "/" + baseName + "__" + stamp + ext);
    var tempResult = writeTextFile(tempFile, text);
    if (tempResult.ok) {
        return { ok: true, error: "", warning: "Primary path unavailable (" + primary.error + "). Wrote fallback file to temp instead.", path: tempFile.fsName };
    }

    return {
        ok: false,
        error: primary.error + " | same-folder fallback failed: " + sameFolder.error + " | temp fallback failed: " + tempResult.error,
        warning: "",
        path: fileObj.fsName
    };
}

function initManagedLog(fileObj, headerText) {
    var result = writeManagedTextFile(fileObj, headerText);
    return { path: result.path, warning: result.warning || "", error: result.ok ? "" : result.error };
}

function appendManagedLog(logState, text) {
    if (!logState || !logState.path) return;
    var currentFile = new File(logState.path);
    var result = appendTextFile(currentFile, text);
    if (result.ok) return;

    var stamp = makeTimestampTag();
    var tempDir = new Folder(Folder.temp.fsName + "/Sizer_Logs");
    ensureFolder(tempDir);
    var fallbackFile = new File(tempDir.fsName + "/" + stripExt(currentFile.name) + "__" + stamp + ".txt");
    var fallbackWrite = writeTextFile(fallbackFile, "[Log continued after append failure]\r\n" + text);
    if (fallbackWrite.ok) {
        if (!logState.warning) logState.warning = "Primary log path became unavailable. Continued in fallback log.";
        logState.path = fallbackFile.fsName;
    } else if (!logState.error) {
        logState.error = result.error + " | fallback append failed: " + fallbackWrite.error;
    }
}

function writeFinalLog(logFileObj, logBufferParts, reportMeta, reportRows) {
    return writeManagedTextFile(logFileObj, buildFinalLogText(logBufferParts, reportMeta, reportRows));
}

function writeDiagnosticsFiles(exportFolder, diagnosticsState, baseName) {
    ensureFolder(exportFolder);
    var textFile = new File(exportFolder.fsName + "/" + baseName + ".txt");
    var jsonFile = new File(exportFolder.fsName + "/" + baseName + ".json");
    var textResult = writeManagedTextFile(textFile, buildDiagnosticsText(diagnosticsState));
    var jsonResult = writeManagedTextFile(jsonFile, buildDiagnosticsJson(diagnosticsState));
    return {
        text: createManagedWriteDescriptor(textFile, textResult),
        json: createManagedWriteDescriptor(jsonFile, jsonResult)
    };
}

function writeHtmlReport(exportFolder, reportMeta, reportRows, targetPath) {
    ensureFolder(exportFolder);
    var reportFile = targetPath ? new File(targetPath) : new File(exportFolder.fsName + "/_Export_REPORT.html");
    return writeManagedTextFile(reportFile, buildReportHtml(reportMeta, reportRows));
}

function writeProofHtml(exportFolder, reportMeta, reportRows, targetPath) {
    ensureFolder(exportFolder);
    var proofFile = targetPath ? new File(targetPath) : new File(exportFolder.fsName + "/_Customer_Proof.html");
    return writeManagedTextFile(proofFile, buildProofHtml(reportMeta, reportRows));
}

function writePricingAuditHtml(exportFolder, reportMeta, reportRows, targetPath) {
    ensureFolder(exportFolder);
    var auditFile = targetPath ? new File(targetPath) : new File(exportFolder.fsName + "/_Pricing_Audit.html");
    return writeManagedTextFile(auditFile, buildPricingAuditHtml(reportMeta, reportRows));
}

var dlg = new Window("dialog", "DTF Batch Export - Photoshop");
dlg.orientation = "column";
dlg.alignChildren = "fill";
dlg.margins = 16;

var mainRow = dlg.add("group");
mainRow.orientation = "row";
mainRow.alignChildren = "fill";
mainRow.spacing = 12;

var leftCol = mainRow.add("group");
leftCol.orientation = "column";
leftCol.alignChildren = "fill";
leftCol.spacing = 10;

var filenameFormatPanel = leftCol.add("panel", undefined, "Filename Format");
filenameFormatPanel.orientation = "column";
filenameFormatPanel.alignChildren = "left";
filenameFormatPanel.margins = 12;
var rbFilenameQty = filenameFormatPanel.add("radiobutton", undefined, "Filename___Qty");
var rbQtyFilename = filenameFormatPanel.add("radiobutton", undefined, "Qty___Filename");
var rbFilenameOnly = filenameFormatPanel.add("radiobutton", undefined, "Filename");
rbQtyFilename.value = true;

var resizePanel = leftCol.add("panel", undefined, "Resize Mode");
resizePanel.orientation = "column";
resizePanel.alignChildren = "left";
resizePanel.margins = 12;
var rbW = resizePanel.add("radiobutton", undefined, "Respect Width (keep aspect ratio)");
var rbH = resizePanel.add("radiobutton", undefined, "Respect Height (keep aspect ratio)");
var rbS = resizePanel.add("radiobutton", undefined, "Stretch (exact Width + Height)");
rbW.value = true;

var printTypePanel = leftCol.add("panel", undefined, "Sort by Print Type");
printTypePanel.orientation = "column";
printTypePanel.alignChildren = "left";
printTypePanel.margins = 12;
var rbPrintFolder = printTypePanel.add("radiobutton", undefined, "Folder (subfolders COOL, HEAT, UV...)");
var rbPrintPrefix = printTypePanel.add("radiobutton", undefined, "Prefix (print type before filename)");
var rbPrintNone = printTypePanel.add("radiobutton", undefined, "None");
rbPrintNone.value = true;

var actionPanel = leftCol.add("panel", undefined, "Actions");
actionPanel.orientation = "column";
actionPanel.alignChildren = "left";
actionPanel.margins = 12;
var runWeMustChk = actionPanel.add("checkbox", undefined, "Run Action: WeMust / WeMust");
runWeMustChk.value = false;
var generateProofChk = actionPanel.add("checkbox", undefined, "Generate Customer Proof HTML");
generateProofChk.value = false;
var generatePricingChk = actionPanel.add("checkbox", undefined, "Generate Pricing Audit HTML");
generatePricingChk.value = false;

var dpiPanel = leftCol.add("panel", undefined, "Output DPI");
dpiPanel.orientation = "row";
dpiPanel.alignChildren = "left";
dpiPanel.margins = 12;
dpiPanel.add("statictext", undefined, "Fixed at 300 DPI");

var folderPanel = leftCol.add("panel", undefined, "Files Folder");
folderPanel.orientation = "row";
folderPanel.alignChildren = "fill";
folderPanel.margins = 12;
var folderPath = folderPanel.add("edittext", undefined, "");
folderPath.characters = 30;
var browseBtn = folderPanel.add("button", undefined, "Browse...");
browseBtn.onClick = function () {
    var f = Folder.selectDialog("Select files folder (downloaded assets)");
    if (f) folderPath.text = f.fsName;
};

var rightCol = mainRow.add("panel", undefined, "Paste Order Email");
rightCol.orientation = "column";
rightCol.alignChildren = "fill";
rightCol.margins = 12;
var emailInput = rightCol.add("edittext", undefined, "", { multiline: true, scrolling: true });
emailInput.preferredSize = [650, 430];

var btns = dlg.add("group");
btns.alignment = "right";
btns.add("button", undefined, "Cancel");
btns.add("button", undefined, "Run", { name: "ok" });

if (dlg.show() !== 1) {
    alert("Operation cancelled");
} else {
    var qtyInFilename = rbFilenameQty.value ? "filenameQty" : (rbQtyFilename.value ? "qtyFilename" : "filename");
    var resizeMode = rbH.value ? "respectHeight" : (rbS.value ? "stretch" : "respectWidth");
    var printTypeMode = rbPrintFolder.value ? "folder" : (rbPrintPrefix.value ? "prefix" : "none");
    var runWeMustAction = !!runWeMustChk.value;
    var generateProof = !!generateProofChk.value;
    var generatePricing = !!generatePricingChk.value;

    if (!folderPath.text) {
        alert("Please select a folder.");
    } else {
        var inputFolder = new Folder(folderPath.text);
        if (!inputFolder.exists) {
            alert("Selected folder does not exist.");
        } else if (!emailInput.text || emailInput.text.length < 10) {
            alert("Please paste the order email text.");
        } else {
            var exportFolder = new Folder(inputFolder.fsName + "/Export");
            ensureFolder(exportFolder);

            var reportRows = [];
            var reportWriteError = "";
            var reportWriteWarning = "";
            var reportPath = new File(exportFolder.fsName + "/_Export_REPORT.html").fsName;
            var logPath = new File(exportFolder.fsName + "/_Export_LOG.txt").fsName;
            var logWriteWarning = "";
            var logWriteError = "";
            var proofWriteError = "";
            var proofWriteWarning = "";
            var proofPath = new File(exportFolder.fsName + "/_Customer_Proof.html").fsName;
            var pricingWriteError = "";
            var pricingWriteWarning = "";
            var pricingPath = new File(exportFolder.fsName + "/_Pricing_Audit.html").fsName;
            var diagnosticsWriteWarning = "";
            var diagnosticsWriteError = "";
            var reportDirtyCount = 0;
            var missing = [];
            var processed = 0;
            var diagnosticsState = initDiagnosticsState();
            var logBufferParts = [];
            var diagnosticsFiles = null;
            var logDescriptor = null;
            var reportDescriptor = null;
            var proofDescriptor = null;
            var pricingDescriptor = null;
            var orderFormat = "classic";
            var fatalRunError = "";
            var cancelledByUser = false;

            var emailText = emailInput.text;
            addDiagnostic(diagnosticsState, "info", "script_start", { app: "Photoshop", folder: inputFolder.fsName });
            addDiagnostic(diagnosticsState, "info", "user_settings", {
                resizeMode: resizeMode,
                filenameFormat: qtyInFilename,
                printTypeMode: printTypeMode,
                runWeMustAction: runWeMustAction,
                generateProof: generateProof,
                generatePricing: generatePricing
            });

            var parseResult = parseEmailItems(emailText, diagnosticsState);
            var items = parseResult.items;
            orderFormat = parseResult.formatName;

            if (items.length === 0) {
                addDiagnostic(diagnosticsState, "error", "no_items_found", { format: orderFormat });
                diagnosticsFiles = writeDiagnosticsFiles(exportFolder, diagnosticsState, "_Diagnostics");
                if (!diagnosticsFiles.text.ok) diagnosticsWriteError = diagnosticsFiles.text.error;
                else if (diagnosticsFiles.text.warning) diagnosticsWriteWarning = diagnosticsFiles.text.warning;
                alert("No valid items found in the pasted email.");
            } else {
                addDiagnostic(diagnosticsState, "info", "items_detected", { count: items.length, format: orderFormat });
                var emailFinancials = parseEmailFinancials(emailText, items, orderFormat, diagnosticsState);
                var reportMeta = {
                    appName: "Photoshop",
                    date: (new Date()).toString(),
                    resizeMode: formatResizeModeLabel(resizeMode),
                    dpi: TARGET_DPI,
                    filenameFormat: formatFilenameFormatLabel(qtyInFilename),
                    printTypeMode: formatPrintTypeModeLabel(printTypeMode),
                    actionSummary: runWeMustAction ? "WeMust / WeMust" : "None",
                    exportFolder: exportFolder.fsName,
                    itemsFound: items.length,
                    currency: emailFinancials.currency,
                    subtotal: emailFinancials.subtotal,
                    shipping: emailFinancials.shipping,
                    tax: emailFinancials.tax,
                    total: emailFinancials.total,
                    taxRate: emailFinancials.taxRate
                };

                function flushReport(force) {
                    if (!force && !(CHECKPOINT_REPORT_WRITES_ENABLED && reportDirtyCount >= REPORT_FLUSH_INTERVAL)) return null;
                    var result = writeHtmlReport(exportFolder, reportMeta, reportRows, reportPath);
                    reportWriteError = result.ok ? "" : result.error;
                    reportWriteWarning = result.warning || reportWriteWarning;
                    reportPath = result.path;
                    reportDirtyCount = 0;
                    addDiagnostic(diagnosticsState, result.ok ? "info" : "error", "report_write_attempt", {
                        checkpoint: force ? "final" : "checkpoint",
                        ok: result.ok,
                        path: result.path,
                        warning: result.warning || "",
                        error: result.error || ""
                    });
                    return result;
                }

                function updateReportRow(index, row) {
                    reportRows[index] = row;
                    reportDirtyCount++;
                    addDiagnostic(diagnosticsState, (row.status === "OK" || row.status === "CHECK" || row.status === "NOT OK") ? "info" : "warn", "item_result", {
                        index: index + 1,
                        file: row.file,
                        status: row.status,
                        match: row.match,
                        output: row.outputSize
                    });
                    flushReport(false);
                }

                logBufferParts = [buildLogHeader(reportMeta)];

                var allFiles = getFilesInFolder(inputFolder);
                var missingCount = 0;
                for (var m = 0; m < items.length; m++) {
                    items[m].matchInfo = findFileMatchByEmailName(allFiles, items[m].file);
                    if (!items[m].matchInfo.file) missingCount++;
                }
                reportRows = buildInitialReportRows(items);
                addDiagnostic(diagnosticsState, "info", "preflight_complete", { items: items.length, missingCount: missingCount });

                var shouldProcess = true;
                if (missingCount > 0) {
                    shouldProcess = confirm(
                        "Preflight result:\n" +
                        "Items found in email: " + items.length + "\n" +
                        "Items missing/unmatched: " + missingCount + "\n\n" +
                        "Continue anyway?"
                    );
                    if (!shouldProcess) {
                        cancelledByUser = true;
                        addDiagnostic(diagnosticsState, "warn", "user_cancelled_after_preflight", { missingCount: missingCount });
                        alert("Operation cancelled");
                    }
                }

                try {
                    if (shouldProcess) {
                        var oldUnits = app.preferences.rulerUnits;
                        try {
                            app.preferences.rulerUnits = Units.PIXELS;

                            for (var i = 0; i < items.length; i++) {
                                var item = items[i];
                                var matchInfo = item.matchInfo || { file: null, matchType: "missing", suggested: "" };
                                addDiagnostic(diagnosticsState, "info", "item_start", {
                                    index: i + 1,
                                    file: item.file,
                                    qty: item.qty,
                                    width: item.width,
                                    height: item.height
                                });

                                if (isNaN(item.width) || isNaN(item.height)) {
                                    missing.push(item.file + " (bad width/height)");
                                    addDiagnostic(diagnosticsState, "warn", "item_bad_dimensions", { index: i + 1, file: item.file });
                                    updateReportRow(i, makeStatusRow(item.file, item.qty, item.printType, item.note, item.price, item.currency, matchInfo, item.width, item.height, "BAD_WIDTH_HEIGHT"));
                                    continue;
                                }

                                if (!matchInfo.file || !matchInfo.file.exists) {
                                    var missMsg = item.file;
                                    if (matchInfo.suggested) missMsg += " (did you mean: " + matchInfo.suggested + ")";
                                    missing.push(missMsg);
                                    addDiagnostic(diagnosticsState, "warn", "item_missing_file", { index: i + 1, file: item.file, suggested: matchInfo.suggested || "" });
                                    updateReportRow(i, makeStatusRow(item.file, item.qty, item.printType, item.note, item.price, item.currency, matchInfo, item.width, item.height, "MISSING_FILE"));
                                    continue;
                                }

                                var doc = safeOpen(matchInfo.file);
                                if (!doc) {
                                    missing.push(item.file + " (failed to open)");
                                    addDiagnostic(diagnosticsState, "error", "item_open_failed", { index: i + 1, file: item.file, path: matchInfo.file.fsName });
                                    updateReportRow(i, makeStatusRow(item.file, item.qty, item.printType, item.note, item.price, item.currency, matchInfo, item.width, item.height, "OPEN_FAIL"));
                                    continue;
                                }

                                try {
                                    if (!ensureRgbDocument(doc)) {
                                        missing.push(item.file + " (rgb conversion failed)");
                                        addDiagnostic(diagnosticsState, "error", "item_rgb_failed", { index: i + 1, file: item.file });
                                        updateReportRow(i, makeStatusRow(item.file, item.qty, item.printType, item.note, item.price, item.currency, matchInfo, item.width, item.height, "RGB_FAIL"));
                                        try { doc.close(SaveOptions.DONOTSAVECHANGES); } catch (eCloseRgbOpen) {}
                                        continue;
                                    }

                                    try { doc.trim(TrimType.TRANSPARENT, true, true, true, true); } catch (eTrim) {
                                        addDiagnostic(diagnosticsState, "warn", "item_trim_skipped", { index: i + 1, file: item.file, error: safeErrorMessage(eTrim) });
                                    }

                                    var targetWidthPx = Math.round(item.width * TARGET_DPI);
                                    var targetHeightPx = Math.round(item.height * TARGET_DPI);
                                    if (resizeMode === "stretch") doc.resizeImage(UnitValue(targetWidthPx, "px"), UnitValue(targetHeightPx, "px"), TARGET_DPI, RESAMPLE);
                                    else if (resizeMode === "respectWidth") doc.resizeImage(UnitValue(targetWidthPx, "px"), undefined, TARGET_DPI, RESAMPLE);
                                    else doc.resizeImage(undefined, UnitValue(targetHeightPx, "px"), TARGET_DPI, RESAMPLE);

                                    if (runWeMustAction) {
                                        var actionResult = runActionIfNeeded(true, "WeMust", "WeMust");
                                        if (!actionResult.ok) {
                                            missing.push(item.file + " (action WeMust failed)");
                                            addDiagnostic(diagnosticsState, "error", "item_action_failed", { index: i + 1, file: item.file, error: actionResult.msg });
                                            updateReportRow(i, makeStatusRow(item.file, item.qty, item.printType, item.note, item.price, item.currency, matchInfo, item.width, item.height, "ACTION_FAIL"));
                                            try { doc.close(SaveOptions.DONOTSAVECHANGES); } catch (eCloseAction) {}
                                            continue;
                                        }
                                    }

                                    if (!ensureRgbDocument(doc)) {
                                        missing.push(item.file + " (rgb conversion failed)");
                                        addDiagnostic(diagnosticsState, "error", "item_rgb_failed", { index: i + 1, file: item.file, stage: "post_action" });
                                        updateReportRow(i, makeStatusRow(item.file, item.qty, item.printType, item.note, item.price, item.currency, matchInfo, item.width, item.height, "RGB_FAIL"));
                                        try { doc.close(SaveOptions.DONOTSAVECHANGES); } catch (eCloseRgbFinal) {}
                                        continue;
                                    }

                                    var outW = doc.width.as("px") / TARGET_DPI;
                                    var outH = doc.height.as("px") / TARGET_DPI;
                                    var base = makeBaseWithQtyOption(item.qty, stripExt(item.file), qtyInFilename);
                                    if (printTypeMode === "prefix" && item.printType) base = item.printType + "___" + base;

                                    var destFolder = getOutputFolderByPrintType(exportFolder, printTypeMode, item.printType);
                                    var outputFile = new File(destFolder.fsName + "/" + base + ".png");
                                    doc.saveAs(outputFile, new PNGSaveOptions(), true);

                                    processed++;
                                    addDiagnostic(diagnosticsState, "info", "item_exported", {
                                        index: i + 1,
                                        file: item.file,
                                        output: outputFile.fsName,
                                        outW: round2(outW),
                                        outH: round2(outH)
                                    });
                                    updateReportRow(i, makeMeasuredRow(item.file, item.qty, item.printType, item.note, item.price, item.currency, matchInfo, resizeMode, item.width, item.height, outW, outH, getRelativePath(outputFile, exportFolder), outputFile.fsName));
                                } catch (eProc) {
                                    missing.push(item.file + " (process/export error)");
                                    addDiagnostic(diagnosticsState, "error", "item_process_failed", {
                                        index: i + 1,
                                        file: item.file,
                                        error: safeErrorMessage(eProc),
                                        stack: safeErrorStack(eProc)
                                    });
                                    updateReportRow(i, makeStatusRow(item.file, item.qty, item.printType, item.note, item.price, item.currency, matchInfo, item.width, item.height, "PROCESS_ERROR"));
                                }

                                try { doc.close(SaveOptions.DONOTSAVECHANGES); } catch (eClose) {
                                    addDiagnostic(diagnosticsState, "warn", "item_close_failed", { index: i + 1, file: item.file, error: safeErrorMessage(eClose) });
                                }
                            }
                        } finally {
                            app.preferences.rulerUnits = oldUnits;
                        }
                    }
                } catch (eRun) {
                    fatalRunError = safeErrorMessage(eRun);
                    addDiagnostic(diagnosticsState, "error", "fatal_run_error", {
                        error: fatalRunError,
                        stack: safeErrorStack(eRun)
                    });
                }

                if (!cancelledByUser) {
                    var finalReportResult = flushReport(true);
                    if (!finalReportResult) finalReportResult = { ok: false, error: "Report write was not attempted", warning: "", path: reportPath };
                    reportDescriptor = createManagedWriteDescriptor(new File(exportFolder.fsName + "/_Export_REPORT.html"), finalReportResult);
                    reportWriteWarning = reportDescriptor.warning || reportWriteWarning;
                    reportWriteError = reportDescriptor.error || reportWriteError;

                    var finalLogResult = writeFinalLog(new File(exportFolder.fsName + "/_Export_LOG.txt"), logBufferParts, reportMeta, reportRows);
                    logDescriptor = createManagedWriteDescriptor(new File(exportFolder.fsName + "/_Export_LOG.txt"), finalLogResult);
                    logPath = logDescriptor.path;
                    logWriteWarning = logDescriptor.warning || logWriteWarning;
                    logWriteError = logDescriptor.error || logWriteError;
                    addDiagnostic(diagnosticsState, finalLogResult.ok ? "info" : "error", "log_write_attempt", {
                        ok: finalLogResult.ok,
                        path: finalLogResult.path,
                        warning: finalLogResult.warning || "",
                        error: finalLogResult.error || ""
                    });

                    if (generateProof) {
                        var proofResult = writeProofHtml(exportFolder, reportMeta, reportRows, proofPath);
                        proofDescriptor = createManagedWriteDescriptor(new File(exportFolder.fsName + "/_Customer_Proof.html"), proofResult);
                        proofPath = proofDescriptor.path;
                        proofWriteWarning = proofDescriptor.warning || proofWriteWarning;
                        proofWriteError = proofDescriptor.error || proofWriteError;
                        addDiagnostic(diagnosticsState, proofResult.ok ? "info" : "error", "proof_write_attempt", {
                            ok: proofResult.ok,
                            path: proofResult.path,
                            warning: proofResult.warning || "",
                            error: proofResult.error || ""
                        });
                    }

                    if (generatePricing) {
                        var pricingResult = writePricingAuditHtml(exportFolder, reportMeta, reportRows, pricingPath);
                        pricingDescriptor = createManagedWriteDescriptor(new File(exportFolder.fsName + "/_Pricing_Audit.html"), pricingResult);
                        pricingPath = pricingDescriptor.path;
                        pricingWriteWarning = pricingDescriptor.warning || pricingWriteWarning;
                        pricingWriteError = pricingDescriptor.error || pricingWriteError;
                        addDiagnostic(diagnosticsState, pricingResult.ok ? "info" : "error", "pricing_write_attempt", {
                            ok: pricingResult.ok,
                            path: pricingResult.path,
                            warning: pricingResult.warning || "",
                            error: pricingResult.error || ""
                        });
                    }
                }

                diagnosticsFiles = writeDiagnosticsFiles(exportFolder, diagnosticsState, "_Diagnostics");
                if (!diagnosticsFiles.text.ok) diagnosticsWriteError = diagnosticsFiles.text.error;
                else if (diagnosticsFiles.text.warning) diagnosticsWriteWarning = diagnosticsFiles.text.warning;
                if (!diagnosticsFiles.json.ok && !diagnosticsWriteError) diagnosticsWriteError = diagnosticsFiles.json.error;
                else if (diagnosticsFiles.json.warning && !diagnosticsWriteWarning) diagnosticsWriteWarning = diagnosticsFiles.json.warning;

                if (!cancelledByUser) {
                    var finalStats = buildReportStats(reportRows);
                    var failedItems = finalStats.errors + finalStats.queued;
                    var msg = "Done!\n"
                        + "Mode: " + formatResizeModeLabel(resizeMode) + "\n"
                        + "DPI: " + TARGET_DPI + "\n"
                        + "Order format: " + orderFormat + "\n"
                        + "Items found: " + items.length + "\n"
                        + "Files processed: " + processed + "\n"
                        + "Failed items: " + failedItems + "\n"
                        + "Report: " + describeManagedWrite(reportDescriptor) + "\n"
                        + "Log: " + describeManagedWrite(logDescriptor) + "\n"
                        + "Diagnostics: " + (diagnosticsFiles ? describeManagedWrite(diagnosticsFiles.text) : "_Diagnostics.txt");
                    if (diagnosticsFiles) msg += "\nDiagnostics JSON: " + describeManagedWrite(diagnosticsFiles.json);
                    if (generateProof) msg += "\nProof: " + describeManagedWrite(proofDescriptor);
                    if (generatePricing) msg += "\nPricing audit: " + describeManagedWrite(pricingDescriptor);
                    if (fatalRunError) msg += "\n\nFatal run issue:\n- " + fatalRunError;
                    if (missing.length > 0) msg += "\n\nMissing/Problem (" + missing.length + "):\n- " + missing.join("\n- ");
                    if (logWriteWarning) msg += "\n\nLog write note:\n- " + logWriteWarning;
                    if (logWriteError) msg += "\n\nLog write issue:\n- " + logWriteError;
                    if (reportWriteWarning) msg += "\n\nReport write note:\n- " + reportWriteWarning;
                    if (reportWriteError) msg += "\n\nReport write issue:\n- " + reportWriteError;
                    if (proofWriteWarning) msg += "\n\nProof write note:\n- " + proofWriteWarning;
                    if (proofWriteError) msg += "\n\nProof write issue:\n- " + proofWriteError;
                    if (pricingWriteWarning) msg += "\n\nPricing write note:\n- " + pricingWriteWarning;
                    if (pricingWriteError) msg += "\n\nPricing write issue:\n- " + pricingWriteError;
                    if (diagnosticsWriteWarning) msg += "\n\nDiagnostics write note:\n- " + diagnosticsWriteWarning;
                    if (diagnosticsWriteError) msg += "\n\nDiagnostics write issue:\n- " + diagnosticsWriteError;
                    alert(msg);
                }
            }
        }
    }
}
