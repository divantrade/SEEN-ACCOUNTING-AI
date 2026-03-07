// ==================== محرك الذكاء الاصطناعي ====================
/**
 * ملف محرك AI لنظام SEEN المحاسبي
 * يستخدم Gemini Flash لتحليل النصوص الطبيعية واستخراج الحركات المالية
 */

// ==================== الاتصال بـ Gemini API ====================

/**
 * استدعاء Gemini API لتحليل النص
 * @param {string} userMessage - رسالة المستخدم
 * @param {Object} context - السياق (المشاريع، البنود، الأطراف)
 * @returns {Object} - نتيجة التحليل
 */
/**
 * استدعاء Gemini API لتحليل النص
 * @param {string} userMessage - رسالة المستخدم
 * @param {Object} context - السياق (المشاريع، البنود، الأطراف)
 * @returns {Object} - نتيجة التحليل
 */
function callGemini(userMessage, context) {
    const apiKey = getGeminiApiKey();
    const props = PropertiesService.getScriptProperties();
    const CACHED_MODEL_KEY = 'GEMINI_WORKING_MODEL';

    // 1. محاولة استخدام الموديل المخزن سابقاً (للسرعة)
    let cachedModel = props.getProperty(CACHED_MODEL_KEY);
    const fallbackModels = AI_CONFIG.GEMINI.FALLBACK_MODELS;

    // ترتيب الموديلات: الموديل المخزن أولاً، ثم الباقي
    let modelsToTry = [];
    if (cachedModel) {
        modelsToTry.push(cachedModel);
        // إضافة الباقي مع تجنب التكرار
        fallbackModels.forEach(m => {
            if (m !== cachedModel) modelsToTry.push(m);
        });
    } else {
        modelsToTry = fallbackModels;
    }

    let lastError = null;

    // محاولة الاتصال بالموديلات
    for (let i = 0; i < modelsToTry.length; i++) {
        const modelName = modelsToTry[i];
        Logger.log(`🔄 جاري تجربة الموديل: ${modelName} (${i + 1}/${modelsToTry.length})`);

        try {
            const url = `${AI_CONFIG.GEMINI.BASE_URL}${modelName}:generateContent?key=${apiKey}`;

            // بناء الـ prompt مع السياق
            const fullPrompt = buildFullPrompt(userMessage, context);

            const payload = {
                contents: [{
                    parts: [{
                        text: fullPrompt
                    }]
                }],
                generationConfig: AI_CONFIG.GEMINI.GENERATION_CONFIG,
                safetySettings: AI_CONFIG.GEMINI.SAFETY_SETTINGS
            };

            const options = {
                method: 'post',
                contentType: 'application/json',
                payload: JSON.stringify(payload),
                muteHttpExceptions: true
            };

            const response = UrlFetchApp.fetch(url, options);
            const responseCode = response.getResponseCode();
            const responseText = response.getContentText();

            if (responseCode === 200) {
                Logger.log(`✅ نجاح الاتصال بالموديل: ${modelName}`);

                // حفظ الموديل الناجح للمستقبل إذا كان مختلفاً
                if (modelName !== cachedModel) {
                    props.setProperty(CACHED_MODEL_KEY, modelName);
                    Logger.log(`💾 تم حفظ الموديل ${modelName} كخيار افتراضي سريع.`);
                }

                const result = JSON.parse(responseText);
                Logger.log('📋 Gemini response structure: ' + JSON.stringify(Object.keys(result)));

                // استخراج النص من الرد
                if (result.candidates && result.candidates[0] && result.candidates[0].content) {
                    const text = result.candidates[0].content.parts[0].text;
                    return parseGeminiResponse(text);
                } else if (result.candidates && result.candidates[0] && result.candidates[0].finishReason) {
                    // حالة الحظر أو عدم وجود محتوى
                    Logger.log('⚠️ Gemini blocked or no content: ' + result.candidates[0].finishReason);
                    lastError = 'تم حظر الرد من قبل Gemini: ' + result.candidates[0].finishReason;
                } else {
                    Logger.log('⚠️ Unexpected response structure: ' + responseText.substring(0, 200));
                    lastError = 'بنية رد غير متوقعة من Gemini';
                }
            } else {
                Logger.log(`❌ فشل الموديل ${modelName}: ${responseCode}`);
                if (responseCode === 404) {
                    // إذا كان 404، نعتبره غير موجود ونكمل
                } else {
                    lastError = `خطأ (${responseCode}): ${responseText.substring(0, 100)}`;
                }
            }

        } catch (error) {
            Logger.log(`❌ خطأ استثنائي مع ${modelName}: ${error.message}`);
            lastError = error.message;
        }
    }

    // إذا فشلت جميع المحاولات - تحليل الخطأ الأخير
    let friendlyError = 'عذراً، لم أتمكن من الاتصال بأي موديل ذكاء اصطناعي حالياً.';

    if (lastError && lastError.includes('403')) {
        friendlyError += '\n🔒 السبب: مفتاح API غير صالح أو غير مفعل (403).';
    } else if (lastError && lastError.includes('429')) {
        friendlyError += '\n⏳ السبب: تجاوز حد الاستخدام المجاني (429). حاول لاحقاً.';
    } else if (lastError && lastError.includes('400')) {
        friendlyError += '\n❌ السبب: طلب غير صحيح (400). ربما المفتاح لا يدعم هذا الموديل.';
    } else {
        friendlyError += `\n⚠️ التفاصيل التقنية: ${lastError}`;
    }

    return {
        success: false,
        error: friendlyError,
        details: 'تمت تجربة الموديلات: ' + modelsToTry.join(', ') + '\nآخر خطأ: ' + lastError
    };
}

/**
 * بناء الـ prompt الكامل مع السياق
 */
function buildFullPrompt(userMessage, context) {
    // ⭐ استبدال تواريخ ديناميكية في البرومبت
    let systemPrompt = AI_CONFIG.SYSTEM_PROMPT;
    systemPrompt = systemPrompt.replace('__DATE_SECTION__', generateDateSection_());

    let prompt = systemPrompt + '\n\n';

    // ⭐ إضافة قوائم طبيعة الحركة والتصنيف (إلزامية من شيت البنود)
    if (context.natures && context.natures.length > 0) {
        prompt += '## ⚠️ طبيعة الحركة المسموحة (استخدم إحدى هذه القيم فقط):\n';
        context.natures.forEach(n => {
            prompt += `- ${n}\n`;
        });
        prompt += '\n';
    }

    if (context.classifications && context.classifications.length > 0) {
        prompt += '## ⚠️ تصنيف الحركة المسموح (استخدم إحدى هذه القيم فقط):\n';
        context.classifications.forEach(c => {
            prompt += `- ${c}\n`;
        });
        prompt += '\n';
    }

    // إضافة قائمة المشاريع مع أكوادها
    if (context.projects && context.projects.length > 0) {
        prompt += '## المشاريع المتاحة (الكود - الاسم):\n';
        context.projects.forEach(p => {
            if (typeof p === 'object') {
                prompt += `- ${p.code} - ${p.name}\n`;
            } else {
                prompt += `- ${p}\n`;
            }
        });
        prompt += '\n';
    }

    // إضافة قائمة البنود
    if (context.items && context.items.length > 0) {
        prompt += '## البنود المتاحة:\n';
        context.items.forEach(i => {
            prompt += `- ${i}\n`;
        });
        prompt += '\n';
    }

    // إضافة قائمة الأطراف
    if (context.parties && context.parties.length > 0) {
        prompt += '## الأطراف المسجلين:\n';
        context.parties.forEach(p => {
            prompt += `- ${p.name} (${p.type})\n`;
        });
        prompt += '\n';
    }

    // إضافة رسالة المستخدم
    prompt += '## نص المستخدم:\n';
    prompt += userMessage + '\n\n';
    prompt += '## المطلوب:\nحلل النص أعلاه واستخرج بيانات الحركة المالية بصيغة JSON. استخدم فقط القيم المتاحة أعلاه لطبيعة الحركة والتصنيف.';

    return prompt;
}

/**
 * ⭐ توليد قسم التواريخ الديناميكي للبرومبت
 * يحسب التواريخ بناءً على اليوم الحالي
 */
function generateDateSection_() {
    const now = new Date();
    const tz = 'Asia/Istanbul';

    // تنسيق التاريخ بـ YYYY-MM-DD
    function fmt(d) {
        const y = d.getFullYear();
        const m = String(d.getMonth() + 1).padStart(2, '0');
        const day = String(d.getDate()).padStart(2, '0');
        return y + '-' + m + '-' + day;
    }

    function addDays(d, days) {
        const r = new Date(d);
        r.setDate(r.getDate() + days);
        return r;
    }

    function lastDayOfMonth(d) {
        return new Date(d.getFullYear(), d.getMonth() + 1, 0);
    }

    function firstDayNextMonth(d) {
        return new Date(d.getFullYear(), d.getMonth() + 1, 1);
    }

    function lastDayNextMonth(d) {
        return new Date(d.getFullYear(), d.getMonth() + 2, 0);
    }

    const dayNames = ['الأحد', 'الاثنين', 'الثلاثاء', 'الأربعاء', 'الخميس', 'الجمعة', 'السبت'];
    const dayName = dayNames[now.getDay()];
    const today = fmt(now);
    const year = now.getFullYear();

    let section = '';
    section += `- اليوم الحالي: ${today} (${dayName})\n`;
    section += '- حوّل الأرقام العربية (٠١٢٣٤٥٦٧٨٩) إلى إنجليزية (0123456789)\n\n';

    section += '- **تواريخ ماضية (للحقل due_date):**\n';
    section += `  - "اليوم" أو "النهارده" = ${today}\n`;
    section += `  - "أمس" أو "امبارح" = ${fmt(addDays(now, -1))}\n`;
    section += `  - "قبل يومين" = ${fmt(addDays(now, -2))}\n`;
    section += `  - "قبل أسبوع" = ${fmt(addDays(now, -7))}\n\n`;

    section += '- **⭐ تواريخ مستقبلية (للحقل payment_term_date) - مهم جداً!:**\n';
    section += `  - "بعد شهر" = ${fmt(addDays(now, 30))} (اليوم + 30 يوم)\n`;
    section += `  - "بعد شهرين" = ${fmt(addDays(now, 60))} (اليوم + 60 يوم)\n`;
    section += `  - "بعد 60 يوم" أو "بعد ٦٠ يوم" = ${fmt(addDays(now, 60))} (اليوم + 60 يوم)\n`;
    section += `  - "بعد أسبوع" = ${fmt(addDays(now, 7))} (اليوم + 7 أيام)\n`;
    section += `  - "بعد أسبوعين" = ${fmt(addDays(now, 14))} (اليوم + 14 يوم)\n`;
    section += `  - "بعد 15 يوم" = ${fmt(addDays(now, 15))} (اليوم + 15 يوم)\n`;
    section += `  - "نهاية الشهر" = ${fmt(lastDayOfMonth(now))}\n`;
    section += `  - "نهاية الشهر الجاي" = ${fmt(lastDayNextMonth(now))}\n`;
    section += `  - "أول الشهر الجاي" = ${fmt(firstDayNextMonth(now))}\n\n`;

    section += '- **تواريخ محددة:**\n';
    section += '  - حوّل أي تاريخ مذكور إلى صيغة YYYY-MM-DD\n';
    section += `  - السنة الحالية: ${year}\n`;

    return section;
}

/**
 * تحليل رد Gemini واستخراج JSON
 */
function parseGeminiResponse(text) {
    try {
        Logger.log('📥 Raw AI Response (first 500 chars): ' + (text || '').substring(0, 500));

        if (!text || text.trim() === '') {
            Logger.log('❌ Empty response from AI');
            return {
                success: false,
                error: 'رد فارغ من AI'
            };
        }

        // محاولة استخراج JSON من النص
        let jsonStr = text;

        // إزالة markdown code blocks إن وجدت
        const jsonMatch = text.match(/```(?:json)?\s*([\s\S]*?)\s*```/);
        if (jsonMatch) {
            jsonStr = jsonMatch[1];
            Logger.log('📋 Extracted from code block');
        }

        // تنظيف النص
        jsonStr = jsonStr.trim();

        // البحث عن أول { وآخر }
        const startIndex = jsonStr.indexOf('{');
        const endIndex = jsonStr.lastIndexOf('}');

        if (startIndex === -1 || endIndex === -1) {
            Logger.log('❌ No JSON object found in response');
            return {
                success: false,
                error: 'لم يتم العثور على JSON في رد AI',
                rawResponse: text.substring(0, 200)
            };
        }

        jsonStr = jsonStr.substring(startIndex, endIndex + 1);
        Logger.log('📋 JSON to parse (first 300 chars): ' + jsonStr.substring(0, 300));

        // ⭐ تنظيف JSON من المشاكل الشائعة قبل التحليل
        jsonStr = cleanJsonString_(jsonStr);

        const parsed = JSON.parse(jsonStr);
        Logger.log('✅ JSON parsed successfully');
        return parsed;

    } catch (error) {
        Logger.log('❌ JSON Parse Error: ' + error.message);
        Logger.log('Raw text (first 500 chars): ' + (text || '').substring(0, 500));

        // ⭐ محاولة إصلاح متقدمة للـ JSON
        try {
            const fixedJson = advancedJsonFix_(text);
            if (fixedJson) {
                Logger.log('✅ JSON fixed and parsed successfully (advanced fix)');
                return fixedJson;
            }
        } catch (fixError) {
            Logger.log('❌ Advanced JSON fix also failed: ' + fixError.message);
        }

        return {
            success: false,
            error: 'فشل في تحليل رد AI: ' + error.message,
            rawResponse: (text || '').substring(0, 200)
        };
    }
}

/**
 * ⭐ تنظيف JSON من المشاكل الشائعة التي ينتجها AI
 */
function cleanJsonString_(jsonStr) {
    if (!jsonStr) return jsonStr;

    // 1. إزالة الأحرف غير المرئية (control characters)
    jsonStr = jsonStr.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, '');

    // 2. إزالة الفواصل الزائدة قبل } أو ] (السبب الأكثر شيوعاً للخطأ)
    // مثل: [1, 2, 3,] أو {"a": 1,}
    jsonStr = jsonStr.replace(/,\s*([}\]])/g, '$1');

    // 3. إزالة تعليقات JavaScript خارج النصوص فقط
    // نزيل فقط التعليقات التي تبدأ من بداية السطر أو بعد مسافة (ليست داخل نصوص)
    jsonStr = jsonStr.replace(/^\s*\/\/[^\n]*/gm, '');
    jsonStr = jsonStr.replace(/\/\*[\s\S]*?\*\//g, '');

    // 4. إعادة إزالة الفواصل الزائدة (قد تظهر بعد إزالة التعليقات)
    jsonStr = jsonStr.replace(/,\s*([}\]])/g, '$1');

    return jsonStr;
}

/**
 * ⭐ محاولة إصلاح متقدمة للـ JSON الفاسد
 */
function advancedJsonFix_(text) {
    if (!text) return null;

    let jsonStr = text;

    // استخراج JSON من markdown
    const jsonMatch = text.match(/```(?:json)?\s*([\s\S]*?)\s*```/);
    if (jsonMatch) {
        jsonStr = jsonMatch[1];
    }

    // البحث عن أول { وآخر }
    const startIndex = jsonStr.indexOf('{');
    const endIndex = jsonStr.lastIndexOf('}');
    if (startIndex === -1 || endIndex === -1) return null;

    jsonStr = jsonStr.substring(startIndex, endIndex + 1);

    // تنظيف أساسي
    jsonStr = cleanJsonString_(jsonStr);

    // محاولة 1: حذف آخر عنصر في المصفوفة المسببة للمشكلة
    // إذا كان الخطأ في مصفوفة، نحاول إغلاقها
    let attempts = [
        jsonStr,
        // محاولة 2: إغلاق أي مصفوفات أو كائنات مفتوحة
        jsonStr.replace(/,\s*$/, '') + '}',
        // محاولة 3: إصلاح النصوص غير المكتملة
        fixUnclosedStrings_(jsonStr)
    ];

    for (let attempt of attempts) {
        try {
            if (attempt) {
                // موازنة الأقواس
                attempt = balanceBrackets_(attempt);
                const result = JSON.parse(attempt);
                return result;
            }
        } catch (e) {
            continue;
        }
    }

    return null;
}

/**
 * ⭐ إصلاح النصوص غير المغلقة في JSON
 */
function fixUnclosedStrings_(jsonStr) {
    if (!jsonStr) return jsonStr;

    let result = '';
    let inString = false;
    let escaped = false;

    for (let i = 0; i < jsonStr.length; i++) {
        const ch = jsonStr[i];

        if (escaped) {
            result += ch;
            escaped = false;
            continue;
        }

        if (ch === '\\' && inString) {
            result += ch;
            escaped = true;
            continue;
        }

        if (ch === '"') {
            inString = !inString;
            result += ch;
            continue;
        }

        // إذا وجدنا سطر جديد داخل نص مفتوح، نغلقه ونفتح واحد جديد
        if (inString && (ch === '\n' || ch === '\r')) {
            result += ' ';
            continue;
        }

        result += ch;
    }

    // إذا النص لا يزال مفتوحاً، أغلقه
    if (inString) {
        result += '"';
    }

    return result;
}

/**
 * ⭐ موازنة الأقواس في JSON
 */
function balanceBrackets_(jsonStr) {
    if (!jsonStr) return jsonStr;

    let openBraces = 0;
    let openBrackets = 0;
    let inString = false;
    let escaped = false;

    for (let i = 0; i < jsonStr.length; i++) {
        const ch = jsonStr[i];

        if (escaped) { escaped = false; continue; }
        if (ch === '\\' && inString) { escaped = true; continue; }
        if (ch === '"') { inString = !inString; continue; }
        if (inString) continue;

        if (ch === '{') openBraces++;
        else if (ch === '}') openBraces--;
        else if (ch === '[') openBrackets++;
        else if (ch === ']') openBrackets--;
    }

    // إضافة الأقواس الناقصة
    let result = jsonStr;
    while (openBrackets > 0) { result += ']'; openBrackets--; }
    while (openBraces > 0) { result += '}'; openBraces--; }

    return result;
}


// ==================== تحميل السياق ====================

/**
 * ⚡ Cache ذكي مقسّم حسب نوع البيانات
 * المشاريع + البنود + التصنيفات + الطبيعة → Cache لمدة 30 دقيقة (نادراً ما تتغير)
 * الأطراف → بدون Cache (تتغير مع كل حركة جديدة بطرف جديد)
 */
var _AI_CONTEXT_CACHE_TTL = 1800; // 30 دقيقة بالثواني

/**
 * تحميل المشاريع مع Cache
 */
function loadProjectsCached() {
    const cache = CacheService.getScriptCache();
    const cacheKey = 'AI_CTX_PROJECTS';
    const cached = cache.get(cacheKey);

    if (cached) {
        try {
            return JSON.parse(cached);
        } catch (e) {
            // Cache تالف - نتجاهله ونقرأ من الشيت
        }
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const projects = loadProjects(ss);

    try {
        cache.put(cacheKey, JSON.stringify(projects), _AI_CONTEXT_CACHE_TTL);
    } catch (e) {
        // فشل الحفظ في Cache (ربما حجم البيانات كبير) - لا مشكلة
        Logger.log('⚠️ Cache put failed for projects: ' + e.message);
    }

    return projects;
}

/**
 * ⭐ جلب البنود المستخدمة سابقاً مع طرف معين من دفتر الحركات
 * @param {string} partyName - اسم الطرف
 * @returns {string[]} - قائمة البنود الفريدة المستخدمة مع هذا الطرف
 */
function getPartyItems(partyName) {
    if (!partyName) return [];

    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const sheet = ss.getSheetByName(CONFIG.SHEETS.TRANSACTIONS);
        if (!sheet) return [];

        const lastRow = sheet.getLastRow();
        if (lastRow < 2) return [];

        // G = البند (عمود 7)، I = اسم الطرف (عمود 9)
        const partyCol = sheet.getRange(2, 9, lastRow - 1, 1).getValues(); // I
        const itemCol = sheet.getRange(2, 7, lastRow - 1, 1).getValues();  // G

        const normalizedParty = normalizeArabicText(partyName);
        const itemsSet = new Set();

        for (let i = 0; i < partyCol.length; i++) {
            const cellParty = partyCol[i][0];
            if (!cellParty) continue;

            const normalizedCell = normalizeArabicText(cellParty.toString().trim());
            if (normalizedCell === normalizedParty ||
                normalizedCell.includes(normalizedParty) ||
                normalizedParty.includes(normalizedCell)) {
                const item = itemCol[i][0];
                if (item && item.toString().trim()) {
                    itemsSet.add(item.toString().trim());
                }
            }
        }

        Logger.log('📋 بنود ' + partyName + ': ' + itemsSet.size + ' بند');
        return Array.from(itemsSet);

    } catch (error) {
        Logger.log('⚠️ getPartyItems Error: ' + error.message);
        return [];
    }
}

/**
 * تحميل البنود + الطبيعة + التصنيفات مع Cache
 */
function loadItemsCached() {
    const cache = CacheService.getScriptCache();
    const cacheKey = 'AI_CTX_ITEMS';
    const cached = cache.get(cacheKey);

    if (cached) {
        try {
            const parsedCache = JSON.parse(cached);
            // ⭐ ضمان وجود 'تغيير عملة' حتى في البيانات المحفوظة في الكاش
            if (parsedCache.natures && !parsedCache.natures.includes('تغيير عملة')) {
                parsedCache.natures.push('تغيير عملة');
            }
            if (parsedCache.classifications && !parsedCache.classifications.includes('بيع دولار')) {
                parsedCache.classifications.push('بيع دولار');
            }
            if (parsedCache.classifications && !parsedCache.classifications.includes('شراء دولار')) {
                parsedCache.classifications.push('شراء دولار');
            }
            return parsedCache;
        } catch (e) {
            // Cache تالف
        }
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const itemsData = loadItems(ss);

    try {
        cache.put(cacheKey, JSON.stringify(itemsData), _AI_CONTEXT_CACHE_TTL);
    } catch (e) {
        Logger.log('⚠️ Cache put failed for items: ' + e.message);
    }

    return itemsData;
}

/**
 * تحميل السياق الكامل (المشاريع، البنود، الأطراف، الأنواع، التصنيفات)
 * ⚡ تحسين: المشاريع + البنود + التصنيفات من Cache، الأطراف دائماً من الشيت
 */
function loadAIContext() {
    const context = {
        projects: [],
        items: [],
        parties: [],
        natures: [],        // أنواع الحركات من شيت البنود
        classifications: [] // تصنيفات الحركات من شيت البنود
    };

    try {
        // ⚡ المشاريع من Cache (نادراً ما تتغير)
        context.projects = loadProjectsCached();

        // ⚡ البنود + الأنواع + التصنيفات من Cache (نادراً ما تتغير)
        const itemsData = loadItemsCached();
        context.items = itemsData.items;
        context.natures = itemsData.natures;
        context.classifications = itemsData.classifications;

        // ⚠️ الأطراف دائماً من الشيت مباشرة (تتغير بشكل متكرر مع إضافة أطراف جديدة)
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        context.parties = loadParties(ss);

        Logger.log('✅ AI Context loaded: ' + context.projects.length + ' projects, ' +
                   context.items.length + ' items, ' + context.parties.length + ' parties, ' +
                   context.natures.length + ' natures, ' + context.classifications.length + ' classifications');

    } catch (error) {
        Logger.log('Context Load Error: ' + error.message);
    }

    return context;
}

/**
 * ⚡ مسح Cache السياق يدوياً - استخدمها بعد تعديل المشاريع أو البنود
 */
function clearAIContextCache() {
    const cache = CacheService.getScriptCache();
    cache.removeAll(['AI_CTX_PROJECTS', 'AI_CTX_ITEMS']);
    Logger.log('✅ تم مسح Cache السياق (المشاريع والبنود)');
}

/**
 * تحميل قائمة المشاريع
 */
function loadProjects(ss) {
    const projects = [];

    try {
        const sheet = ss.getSheetByName(CONFIG.SHEETS.PROJECTS);
        if (!sheet) return projects;

        const data = sheet.getDataRange().getValues();

        // تخطي الهيدر، العمود الأول = كود، العمود الثاني = اسم المشروع
        for (let i = 1; i < data.length; i++) {
            const projectCode = data[i][0]; // كود المشروع
            const projectName = data[i][1]; // اسم المشروع
            if (projectName && projectName.toString().trim()) {
                projects.push({
                    code: projectCode ? projectCode.toString().trim() : '',
                    name: projectName.toString().trim()
                });
            }
        }

    } catch (error) {
        Logger.log('Load Projects Error: ' + error.message);
    }

    return projects;
}

/**
 * تحميل قائمة البنود مع طبيعة الحركة والتصنيف
 */
function loadItems(ss) {
    const result = {
        items: [],
        natures: [],
        classifications: []
    };

    try {
        const sheet = ss.getSheetByName(CONFIG.SHEETS.ITEMS);
        if (!sheet) return result;

        const data = sheet.getDataRange().getValues();
        const naturesSet = new Set();
        const classificationsSet = new Set();

        // تخطي الهيدر
        // A = البند، B = طبيعة الحركة، C = تصنيف الحركة
        for (let i = 1; i < data.length; i++) {
            const itemName = data[i][0];
            const nature = data[i][1];
            const classification = data[i][2];

            if (itemName && itemName.toString().trim()) {
                result.items.push(itemName.toString().trim());
            }
            if (nature && nature.toString().trim()) {
                naturesSet.add(nature.toString().trim());
            }
            if (classification && classification.toString().trim()) {
                classificationsSet.add(classification.toString().trim());
            }
        }

        result.natures = Array.from(naturesSet);
        result.classifications = Array.from(classificationsSet);

        // ⭐ ضمان وجود "تغيير عملة" في القائمة حتى لو لم يكن في شيت البنود
        if (!result.natures.includes('تغيير عملة')) {
            result.natures.push('تغيير عملة');
        }
        // ⭐ ضمان تصنيفات تصريف العملات
        if (!result.classifications.includes('بيع دولار')) {
            result.classifications.push('بيع دولار');
        }
        if (!result.classifications.includes('شراء دولار')) {
            result.classifications.push('شراء دولار');
        }

        Logger.log('📋 Loaded from Items sheet: ' + result.items.length + ' items, ' +
                   result.natures.length + ' natures, ' + result.classifications.length + ' classifications');

    } catch (error) {
        Logger.log('Load Items Error: ' + error.message);
    }

    return result;
}

/**
 * تحميل قائمة الأطراف من الشيتين (الرئيسي + أطراف البوت)
 */
function loadParties(ss) {
    const parties = [];
    const addedNames = new Set(); // لتجنب التكرار

    try {
        // ⭐ 1. تحميل من شيت الأطراف الرئيسي
        const mainSheet = ss.getSheetByName(CONFIG.SHEETS.PARTIES);
        if (mainSheet) {
            const mainData = mainSheet.getDataRange().getValues();
            for (let i = 1; i < mainData.length; i++) {
                const partyName = mainData[i][0];
                const partyType = mainData[i][1];

                if (partyName && partyName.toString().trim()) {
                    const name = partyName.toString().trim();
                    if (!addedNames.has(name.toLowerCase())) {
                        parties.push({
                            name: name,
                            type: partyType ? partyType.toString().trim() : 'مورد'
                        });
                        addedNames.add(name.toLowerCase());
                    }
                }
            }
        }

        // ⭐ 2. تحميل من شيت أطراف البوت (الأطراف الجديدة)
        const botPartiesSheet = ss.getSheetByName(CONFIG.SHEETS.BOT_PARTIES);
        if (botPartiesSheet) {
            const botData = botPartiesSheet.getDataRange().getValues();
            for (let i = 1; i < botData.length; i++) {
                const partyName = botData[i][0]; // اسم الطرف
                const partyType = botData[i][1]; // نوع الطرف

                if (partyName && partyName.toString().trim()) {
                    const name = partyName.toString().trim();
                    // إضافة فقط إذا لم يكن موجوداً بالفعل
                    if (!addedNames.has(name.toLowerCase())) {
                        parties.push({
                            name: name,
                            type: partyType ? partyType.toString().trim() : 'مورد'
                        });
                        addedNames.add(name.toLowerCase());
                        Logger.log('📋 Added bot party: ' + name);
                    }
                }
            }
        }

        Logger.log('✅ Loaded ' + parties.length + ' parties (main + bot)');

    } catch (error) {
        Logger.log('Load Parties Error: ' + error.message);
    }

    return parties;
}


// ==================== المطابقة الذكية ====================

/**
 * البحث الذكي في مصفوفة نصية
 * يعيد النتائج مع نسب التشابه
 */
function fuzzySearchInArray(searchText, array, minScore = 0.5) {
    if (!searchText || !array || array.length === 0) {
        return [];
    }

    const results = [];
    const normalizedSearch = normalizeArabicText(searchText);

    array.forEach(item => {
        const normalizedItem = normalizeArabicText(item);

        // التحقق من التطابق التام
        if (normalizedItem === normalizedSearch) {
            results.push({ item: item, score: 1.0 });
            return;
        }

        // التحقق من الاحتواء
        if (normalizedItem.includes(normalizedSearch) || normalizedSearch.includes(normalizedItem)) {
            results.push({ item: item, score: 0.9 });
            return;
        }

        // التحقق من بداية الكلمة
        if (normalizedItem.startsWith(normalizedSearch)) {
            results.push({ item: item, score: 0.95 });
            return;
        }

        // حساب التشابه
        const similarity = calculateSimilarity(item, searchText);
        if (similarity >= minScore) {
            results.push({ item: item, score: similarity });
        }
    });

    // ترتيب حسب النتيجة
    results.sort((a, b) => b.score - a.score);

    return results;
}

/**
 * مطابقة اسم المشروع مع المشاريع الموجودة
 * يدعم المشاريع ككائنات {code, name} أو كنصوص
 */
function matchProject(projectName, projectsList) {
    if (!projectName || !projectsList || projectsList.length === 0) {
        return { found: false, matches: [] };
    }

    // تحويل القائمة لأسماء فقط للبحث
    const isObjectList = projectsList.length > 0 && typeof projectsList[0] === 'object';

    // بحث مطابق تماماً
    if (isObjectList) {
        const exactMatch = projectsList.find(p =>
            normalizeArabicText(p.name) === normalizeArabicText(projectName)
        );
        if (exactMatch) {
            return {
                found: true,
                match: exactMatch.name,
                code: exactMatch.code,
                score: 1.0
            };
        }
    } else {
        const exactMatch = projectsList.find(p =>
            normalizeArabicText(p) === normalizeArabicText(projectName)
        );
        if (exactMatch) {
            return { found: true, match: exactMatch, score: 1.0 };
        }
    }

    // بحث ذكي
    const searchList = isObjectList ? projectsList.map(p => p.name) : projectsList;
    const results = fuzzySearchInArray(projectName, searchList, 0.5);

    if (results.length > 0) {
        const matchedName = results[0].item;
        let matchedCode = null;

        if (isObjectList) {
            const matchedProject = projectsList.find(p => p.name === matchedName);
            matchedCode = matchedProject ? matchedProject.code : null;
        }

        return {
            found: true,
            match: matchedName,
            code: matchedCode,
            score: results[0].score,
            alternatives: results.slice(1, 4).map(r => {
                if (isObjectList) {
                    const proj = projectsList.find(p => p.name === r.item);
                    return { name: r.item, code: proj ? proj.code : null };
                }
                return r.item;
            })
        };
    }

    return { found: false, matches: [] };
}

/**
 * مطابقة اسم الطرف مع الأطراف الموجودة
 * يبحث عن تطابق تام أو جزئي ويُرجع الاقتراحات
 */
function matchParty(partyName, partiesList) {
    if (!partyName || !partiesList || partiesList.length === 0) {
        return { found: false, matches: [] };
    }

    const normalizedInput = normalizeArabicText(partyName);
    const partyNames = partiesList.map(p => p.name);

    // 1. بحث مطابق تماماً
    const exactIndex = partyNames.findIndex(p =>
        normalizeArabicText(p) === normalizedInput
    );

    if (exactIndex !== -1) {
        return {
            found: true,
            match: partiesList[exactIndex],
            score: 1.0
        };
    }

    // 2. بحث جزئي (contains) - للأسماء التي تحتوي على النص المدخل
    const containsMatches = partiesList.filter(p => {
        const normalizedName = normalizeArabicText(p.name);
        return normalizedName.includes(normalizedInput) || normalizedInput.includes(normalizedName);
    });

    // 3. بحث ضبابي بعتبة منخفضة لإيجاد المتشابهين
    const fuzzyResults = fuzzySearchInArray(partyName, partyNames, 0.4);
    const fuzzyMatches = fuzzyResults.map(r => {
        const party = partiesList.find(p => p.name === r.item);
        return { ...party, score: r.score };
    });

    // 4. دمج النتائج وإزالة التكرارات
    const allSuggestions = [];
    const addedNames = new Set();

    // أضف نتائج البحث الجزئي أولاً (الأكثر صلة)
    containsMatches.forEach(p => {
        if (!addedNames.has(p.name)) {
            allSuggestions.push({ ...p, score: 0.95 });
            addedNames.add(p.name);
        }
    });

    // أضف نتائج البحث الضبابي
    fuzzyMatches.forEach(p => {
        if (!addedNames.has(p.name)) {
            allSuggestions.push(p);
            addedNames.add(p.name);
        }
    });

    // 5. إذا وجدنا اقتراحات
    if (allSuggestions.length > 0) {
        // ترتيب حسب الدرجة
        allSuggestions.sort((a, b) => b.score - a.score);

        // إذا كان التطابق عالي جداً (> 0.95) نعتبره نفس الشخص
        if (allSuggestions[0].score > 0.95) {
            return {
                found: true,
                match: allSuggestions[0],
                score: allSuggestions[0].score,
                alternatives: allSuggestions.slice(1, 5)
            };
        }

        // تطابق متوسط أو جزئي - نقترح ونترك للمستخدم الاختيار
        return {
            found: false,
            suggestions: allSuggestions.slice(0, 5) // أعلى 5 اقتراحات
        };
    }

    return { found: false, matches: [], suggestions: [] };
}

/**
 * مطابقة البند مع البنود الموجودة
 */
function matchItem(itemName, itemsList) {
    if (!itemName || !itemsList || itemsList.length === 0) {
        return { found: false, matches: [] };
    }

    // بحث مطابق تماماً
    const exactMatch = itemsList.find(i =>
        normalizeArabicText(i) === normalizeArabicText(itemName)
    );

    if (exactMatch) {
        return { found: true, match: exactMatch, score: 1.0 };
    }

    // بحث ذكي
    const results = fuzzySearchInArray(itemName, itemsList, 0.4);

    if (results.length > 0) {
        return {
            found: true,
            match: results[0].item,
            score: results[0].score,
            alternatives: results.slice(1, 4).map(r => r.item)
        };
    }

    return { found: false, matches: [] };
}


// ==================== التحقق من الحركة ====================

/**
 * التحقق من اكتمال الحركة وتحديد الحقول الناقصة
 */
function validateTransaction(transaction, context) {
    Logger.log('═══════════════════════════════════════');
    Logger.log('🔍 validateTransaction STARTED');
    Logger.log('🔍 Transaction input: ' + JSON.stringify(transaction).substring(0, 500));
    Logger.log('🔍 classification: "' + transaction.classification + '"');
    Logger.log('🔍 project: "' + transaction.project + '"');

    const validation = {
        isValid: true,
        missingRequired: [],
        warnings: [],
        enriched: { ...transaction }
    };

    // التحقق من الحقول الإلزامية الأساسية
    if (!transaction.amount || transaction.amount <= 0) {
        validation.missingRequired.push({
            field: 'amount',
            label: 'المبلغ',
            message: 'يرجى تحديد المبلغ'
        });
    }

    // التحويل الداخلي وتصريف العملات لا يحتاجان طرف، المصاريف البنكية الطرف اختياري
    const isInternalTransfer = (transaction.nature || '').includes('تحويل داخلي');
    // ⭐ كشف تغيير العملة: بالطبيعة أو بالتصنيف (بيع/شراء دولار = تغيير عملة حتى لو Gemini أرجع طبيعة خاطئة)
    const classificationVal = (transaction.classification || '').trim();
    const isCurrencyExchange = (transaction.nature || '').includes('تغيير عملة')
        || (transaction.nature || '').includes('تصريف عملات')
        || classificationVal === 'بيع دولار'
        || classificationVal === 'شراء دولار';
    // ⭐ توحيد: إذا كشفنا تغيير عملة بأي طريقة، نصحح الطبيعة لتطابق قاعدة البنود
    if (isCurrencyExchange && !(transaction.nature || '').includes('تغيير عملة')) {
        Logger.log('💱 تم تصحيح الطبيعة من "' + transaction.nature + '" → "تغيير عملة" (كُشفت من التصنيف: ' + classificationVal + ')');
        transaction.nature = 'تغيير عملة';
        validation.enriched.nature = 'تغيير عملة';
    }
    // ⭐ كشف المصاريف البنكية: بالبند أو بالطبيعة القديمة (للتوافق) أو بالكلمات المفتاحية
    const isBankFees = (transaction.item || '').includes('مصاريف بنكية')
        || (transaction.nature || '').includes('مصاريف بنكية')
        || (transaction.nature || '').includes('عمولة بنكية')
        || (transaction.nature || '').includes('رسوم بنكية');

    // ⭐ تصحيح: المصاريف البنكية طبيعتها "دفعة مصروف" وليست طبيعة مستقلة
    if (isBankFees) {
        transaction.nature = 'دفعة مصروف';
        validation.enriched.nature = 'دفعة مصروف';
        transaction.item = 'مصاريف بنكية';
        validation.enriched.item = 'مصاريف بنكية';
        validation.enriched.isBankFees = true;
        Logger.log('🏦 مصاريف بنكية: تصحيح الطبيعة إلى "دفعة مصروف" والبند إلى "مصاريف بنكية"');
    }
    if (!transaction.party && !isInternalTransfer && !isCurrencyExchange && !isBankFees) {
        validation.missingRequired.push({
            field: 'party',
            label: 'الطرف',
            message: 'يرجى تحديد اسم الطرف (المورد/العميل/الممول)'
        });
    }

    // ⭐ تصحيح مبكر للتصنيف: المرتبات والرواتب = مصروفات عمومية (قبل فحص المشروع)
    var salaryKeywordsEarly = ['مرتبات', 'مرتب', 'راتب', 'رواتب', 'أجور', 'أجر', 'مكافأة', 'مكافآت', 'حوافز', 'بدلات'];
    var itemTextEarly = (transaction.item || '');
    var detailsTextEarly = (transaction.details || '');
    var isSalaryEarly = salaryKeywordsEarly.some(function(keyword) {
        return itemTextEarly.includes(keyword) || detailsTextEarly.includes(keyword);
    });

    if (isSalaryEarly && transaction.classification === 'مصروفات مباشرة') {
        Logger.log('⚠️ تصحيح مبكر: المرتبات/الرواتب → مصروفات عمومية');
        transaction.classification = 'مصروفات عمومية';
        validation.enriched.classification = 'مصروفات عمومية';
    }

    // التحقق من المشروع للمصروفات المباشرة والإيرادات
    // ⭐ المشروع اختياري - نسأل عنه لكن يمكن التخطي
    // ⭐ الدفعات (دفعة مصروف / تحصيل إيراد) تحتاج مشروع دائماً لضمان التوزيع الصحيح
    const txNature = transaction.nature || '';
    const isPaymentNature = txNature.includes('دفعة مصروف') || txNature.includes('تحصيل إيراد');
    const needsProject = ['مصروفات مباشرة', 'ايراد'].includes(transaction.classification) || isPaymentNature;
    Logger.log('🔍 needsProject check:');
    Logger.log('   - classification: "' + transaction.classification + '"');
    Logger.log('   - needsProject: ' + needsProject);
    Logger.log('   - project value: ' + transaction.project);
    Logger.log('   - !transaction.project: ' + !transaction.project);

    if (needsProject && !transaction.project) {
        // ⭐ إذا كان هناك مشروع واحد فقط، نختاره تلقائياً بدلاً من سؤال المستخدم
        if (context.projects && context.projects.length === 1) {
            const singleProject = context.projects[0];
            const projectName = typeof singleProject === 'object' ? singleProject.name : singleProject;
            const projectCode = typeof singleProject === 'object' ? (singleProject.code || '') : '';
            transaction.project = projectName;
            transaction.project_code = projectCode;
            validation.enriched.project = projectName;
            validation.enriched.project_code = projectCode;
            Logger.log('✅ مشروع واحد فقط متاح - تم اختياره تلقائياً: ' + projectName);
        } else {
            // أكثر من مشروع - نسأل المستخدم
            Logger.log('📋 Project is optional - will ask with skip option');
            validation.needsProjectSelection = true;
        }
    }

    // مطابقة المشروع
    if (transaction.project && context.projects) {
        const projectMatch = matchProject(transaction.project, context.projects);
        if (projectMatch.found) {
            // ⭐ دائماً نستخدم الاسم والكود من قاعدة البيانات (أحدث نسخة)
            validation.enriched.project = projectMatch.match;
            // الكود من قاعدة البيانات له الأولوية المطلقة (حتى لو كان فارغ)
            validation.enriched.project_code = (projectMatch.code != null && projectMatch.code !== undefined) ? projectMatch.code : (transaction.project_code || '');
            validation.enriched.projectScore = projectMatch.score;
            Logger.log('✅ مشروع من قاعدة البيانات: ' + projectMatch.match + ' (كود: ' + validation.enriched.project_code + ')');
            if (projectMatch.score < 0.9) {
                validation.warnings.push({
                    field: 'project',
                    message: `هل تقصد "${projectMatch.match}"؟`,
                    alternatives: projectMatch.alternatives
                });
            }
        } else {
            validation.warnings.push({
                field: 'project',
                message: `المشروع "${transaction.project}" غير موجود`,
                suggestions: context.projects.slice(0, 5).map(p => typeof p === 'object' ? p.name : p)
            });
        }
    }

    // ⭐ شروط الدفع (تعتمد على نوع الحركة)
    const nature = transaction.nature || '';
    const isPayment = nature.includes('دفعة') || nature.includes('تحصيل') || nature.includes('سداد') || nature.includes('استلام') || nature.includes('تسوية') || nature.includes('دخول قرض');

    if (isPayment) {
        // الدفعات الفعلية: شرط الدفع "فوري" تلقائياً (تم الدفع بتاريخ الحركة)
        validation.enriched.payment_term = 'فوري';
        validation.enriched.payment_term_weeks = '';
        validation.enriched.payment_term_date = '';
    } else if (transaction.payment_term) {
        // استحقاق مع شرط دفع محدد من المستخدم
        validation.enriched.payment_term = transaction.payment_term;
        validation.enriched.payment_term_weeks = transaction.payment_term_weeks || '';
        validation.enriched.payment_term_date = transaction.payment_term_date || '';

        // إذا كان "بعد التسليم" ولم يُحدد عدد الأسابيع
        if (transaction.payment_term === 'بعد التسليم' && !transaction.payment_term_weeks) {
            validation.needsPaymentTermWeeks = true;
        }
        // إذا كان "تاريخ مخصص" ولم يُحدد التاريخ
        if (transaction.payment_term === 'تاريخ مخصص' && !transaction.payment_term_date) {
            validation.needsPaymentTermDate = true;
        }
    } else {
        // استحقاق بدون شرط دفع - نسأل المستخدم
        validation.needsPaymentTerm = true;
        validation.enriched.payment_term = null;
        validation.enriched.payment_term_weeks = '';
        validation.enriched.payment_term_date = '';
    }

    // مطابقة الطرف (التحويل الداخلي وتصريف العملات لا يحتاجان طرف، المصاريف البنكية الطرف اختياري)
    if (isInternalTransfer || isCurrencyExchange) {
        validation.enriched.party = '';
        validation.enriched.isNewParty = false;
        Logger.log('🔄 ' + (isCurrencyExchange ? 'تغيير عملة' : 'تحويل داخلي') + ' - تخطي مطابقة الطرف');
    } else if (isBankFees) {
        // المصاريف البنكية: تعيين التصنيف وطريقة الدفع تلقائياً
        validation.enriched.classification = 'مصروفات عمومية';
        validation.enriched.payment_method = 'تحويل بنكي';
        validation.enriched.payment_term = 'فوري';
        validation.enriched.payment_term_weeks = '';
        validation.enriched.payment_term_date = '';
        // ⭐ المصاريف البنكية: البند دائماً "مصاريف بنكية" بغض النظر عن ما رجعه Gemini
        validation.enriched.item = 'مصاريف بنكية';
        // المصاريف البنكية: الطرف اختياري - إذا ذُكر طرف في النص يتم مطابقته
        if (transaction.party && context.parties) {
            const partyMatch = matchParty(transaction.party, context.parties);
            if (partyMatch.found) {
                validation.enriched.party = partyMatch.match.name;
                validation.enriched.partyType = partyMatch.match.type;
                validation.enriched.partyScore = partyMatch.score;
                Logger.log('🏦 مصاريف بنكية - تم ربط الطرف: ' + partyMatch.match.name);
            } else {
                // الطرف غير موجود - يجب طلب تأكيد
                validation.enriched.isNewParty = true;
                validation.enriched.newPartyName = transaction.party;
                validation.needsPartyConfirmation = true;
                if (partyMatch.suggestions && partyMatch.suggestions.length > 0) {
                    validation.warnings.push({
                        field: 'party',
                        message: 'الطرف "' + transaction.party + '" غير موجود في المصاريف البنكية. هل تقصد أحد هؤلاء؟',
                        suggestions: partyMatch.suggestions,
                        isNew: true
                    });
                } else {
                    validation.warnings.push({
                        field: 'party',
                        message: 'الطرف "' + transaction.party + '" غير موجود. هل تريد إضافته كطرف جديد؟',
                        isNew: true
                    });
                }
                Logger.log('🏦 مصاريف بنكية - طرف جديد يحتاج تأكيد: ' + transaction.party);
            }
        } else {
            validation.enriched.party = '';
            validation.enriched.isNewParty = false;
            Logger.log('🏦 مصاريف بنكية - بدون طرف');
        }
    } else if (transaction.party && context.parties) {
        const partyMatch = matchParty(transaction.party, context.parties);
        if (partyMatch.found) {
            validation.enriched.party = partyMatch.match.name;
            validation.enriched.partyType = partyMatch.match.type;
            validation.enriched.partyScore = partyMatch.score;
        } else {
            // الطرف غير موجود - يجب طلب تأكيد من المستخدم
            validation.enriched.isNewParty = true;
            validation.enriched.newPartyName = transaction.party;
            validation.needsPartyConfirmation = true;

            // إذا وجدت اقتراحات مشابهة
            if (partyMatch.suggestions && partyMatch.suggestions.length > 0) {
                validation.warnings.push({
                    field: 'party',
                    message: `الطرف "${transaction.party}" غير موجود. هل تقصد أحد هؤلاء؟`,
                    suggestions: partyMatch.suggestions,
                    isNew: true
                });
            } else {
                validation.warnings.push({
                    field: 'party',
                    message: `الطرف "${transaction.party}" غير موجود في قاعدة البيانات`,
                    isNew: true
                });
            }
        }
    }

    // مطابقة البند (تخطي المصاريف البنكية والتحويل الداخلي وتغيير العملة - البند يتعيّن تلقائياً)
    if (transaction.item && context.items && !isBankFees && !isInternalTransfer && !isCurrencyExchange) {
        const itemMatch = matchItem(transaction.item, context.items);
        if (itemMatch.found) {
            validation.enriched.item = itemMatch.match;
            validation.enriched.itemScore = itemMatch.score;
            // ⭐ تحديث البند في الحركة أيضاً ليظهر في التأكيد والحفظ
            transaction.item = itemMatch.match;
        }
    }

    // ⭐ التحقق من البند - إذا لم يُذكر البند نسأل عنه (تخطي الحالات الخاصة)
    if (!transaction.item && !isBankFees && !isInternalTransfer && !isCurrencyExchange) {
        Logger.log('📋 Item is missing - will ask user to select');
        validation.needsItemSelection = true;
    }

    // تحويل التاريخ
    if (transaction.due_date === 'TODAY' || !transaction.due_date) {
        validation.enriched.due_date = Utilities.formatDate(new Date(), 'Asia/Istanbul', 'yyyy-MM-dd');
    }

    // ⭐ التحقق من تاريخ استحقاق السلفة/التمويل
    // إذا كانت الحركة تمويل (دخول قرض) أو سلفة، يجب السؤال عن تاريخ السداد
    const classification = transaction.classification || '';
    const isLoan = (nature.includes('تمويل') && !nature.includes('سداد تمويل')) || classification.includes('سلفة');
    if (isLoan && !transaction.loan_due_date) {
        validation.needsLoanDueDate = true;
        validation.enriched.loan_due_date = null;
    } else if (transaction.loan_due_date) {
        validation.enriched.loan_due_date = transaction.loan_due_date;
    }

    // ⭐ التحقق من طريقة الدفع (يجب أن تكون محددة)
    // التحويل الداخلي: تحويل للخزنة = نقدي، تحويل للبنك = تحويل بنكي
    if (isInternalTransfer) {
        const classification = (transaction.classification || '').trim();
        if (classification.includes('خزنة')) {
            validation.enriched.payment_method = 'نقدي';
        } else {
            validation.enriched.payment_method = 'تحويل بنكي';
        }
        // ⭐ البند = "تحويل داخلي" (مطابق لقاعدة البنود) - الاتجاه يُحدد بالتصنيف
        validation.enriched.item = 'تحويل داخلي';
        validation.enriched.payment_term = 'فوري';
        validation.enriched.payment_term_weeks = '';
        validation.enriched.payment_term_date = '';
        // ⭐ التحويل الداخلي فوري دائماً - لا نسأل عن شرط الدفع
        validation.needsPaymentTerm = false;
        validation.needsPaymentMethod = false;
        validation.needsProjectSelection = false;
        // ⭐ التحويل الداخلي من نفس العملة لا يحتاج سعر صرف
        // (ليرة→ليرة أو دولار→دولار: المرجع هو المبلغ بالعملة الأصلية)
        validation.needsExchangeRate = false;
        // ⭐ سعر الصرف = 0 للتحويل الداخلي بنفس العملة (حتى لا يُحسب بالدولار خطأ)
        validation.enriched.exchangeRate = 0;
        validation.enriched.exchange_rate = 0;
    }
    // ⭐ تغيير عملة: يحتاج سعر صرف دائماً + لا يحتاج طرف أو مشروع
    else if (isCurrencyExchange) {
        const classification = (transaction.classification || '').trim();
        if (classification.includes('خزنة') || classification.includes('نقد') || classification.includes('كاش')) {
            validation.enriched.payment_method = 'نقدي';
        } else if (transaction.payment_method && (transaction.payment_method.includes('نقد') || transaction.payment_method.includes('كاش'))) {
            validation.enriched.payment_method = 'نقدي';
        } else {
            validation.enriched.payment_method = transaction.payment_method || 'نقدي';
        }
        // ⭐ البند = "تغيير عملة" (مطابق لقاعدة البنود)
        validation.enriched.item = 'تغيير عملة';
        validation.enriched.payment_term = 'فوري';
        validation.enriched.payment_term_weeks = '';
        validation.enriched.payment_term_date = '';
        validation.needsPaymentTerm = false;
        validation.needsPaymentMethod = false;
        validation.needsProjectSelection = false;
        validation.enriched.party = '';
        validation.enriched.isNewParty = false;
        // تغيير العملة يحتاج سعر صرف دائماً
        const rateVal = Number(transaction.exchange_rate) || 0;
        if (rateVal <= 1) {
            validation.needsExchangeRate = true;
        }
        // ⭐ الحفاظ على العملة الأصلية - لا نحوّل لدولار
        // BotSheets.js يحسب عمود M (القيمة بالدولار) تلقائياً
        validation.enriched.exchangeRate = rateVal;
        Logger.log('💱 تغيير عملة: المبلغ ' + transaction.amount + ' ' + transaction.currency + ' بسعر ' + rateVal);
    }
    // القيم المسموحة: نقدي، تحويل بنكي، شيك، بطاقة، أخرى
    else if (!transaction.payment_method) {
        // لم تحدد طريقة الدفع - نسأل المستخدم
        validation.needsPaymentMethod = true;
        validation.enriched.payment_method = null;
    } else {
        // تحويل القيم المختلفة للقيم الصحيحة
        const method = transaction.payment_method.toLowerCase();
        if (method.includes('نقد') || method.includes('كاش') || method.includes('يد')) {
            validation.enriched.payment_method = 'نقدي';
        } else if (method.includes('بنك') || method.includes('تحويل') || method.includes('حوال')) {
            validation.enriched.payment_method = 'تحويل بنكي';
        } else if (method.includes('شيك') || method.includes('check')) {
            validation.enriched.payment_method = 'شيك';
        } else if (method.includes('بطاق') || method.includes('كارت') || method.includes('فيزا') || method.includes('card')) {
            validation.enriched.payment_method = 'بطاقة';
        } else {
            // قيمة غير معروفة - نسأل المستخدم
            validation.needsPaymentMethod = true;
            validation.enriched.payment_method = null;
        }
    }

    // ⭐ التحقق من العملة (يجب أن تكون محددة)
    if (!transaction.currency) {
        validation.needsCurrency = true;
        validation.enriched.currency = null;
    } else {
        validation.enriched.currency = transaction.currency;

        // ⭐ إذا كانت العملة غير دولار ولا يوجد سعر صرف صحيح
        const rateValue = Number(transaction.exchange_rate) || 0;
        if (transaction.currency !== 'USD' && (rateValue <= 1) && !isInternalTransfer) {
            if (isCurrencyExchange) {
                // تغيير العملة: يجب أن يحدد المستخدم سعر الصرف الفعلي
                validation.needsExchangeRate = true;
            } else {
                // ⭐ الدفعات العادية بالليرة/الجنيه: نستخدم سعر الصرف الافتراضي تلقائياً (لا نسأل المستخدم)
                const defaultRate = getDefaultExchangeRate(transaction.currency);
                validation.enriched.exchangeRate = defaultRate;
                validation.enriched.exchange_rate = defaultRate;
                Logger.log('💱 استخدام سعر صرف افتراضي: ' + defaultRate + ' لعملة ' + transaction.currency);
            }
        }
        // ⭐ تغيير العملة: دائماً يحتاج سعر صرف (حتى لو العملة USD في شراء دولار)
        if (isCurrencyExchange && !transaction.exchange_rate) {
            validation.needsExchangeRate = true;
        }
    }

    // تحويل سعر الصرف من snake_case إلى camelCase
    // ⭐ فقط إذا كان سعر الصرف الأصلي صحيحاً (> 1 لغير الدولار) ولم نحدد سعر افتراضي بالفعل
    if (transaction.exchange_rate && Number(transaction.exchange_rate) > 1) {
        validation.enriched.exchangeRate = transaction.exchange_rate;
        validation.enriched.exchange_rate = transaction.exchange_rate;
    }

    validation.isValid = validation.missingRequired.length === 0;

    Logger.log('🔍 validateTransaction FINISHED');
    Logger.log('🔍 isValid: ' + validation.isValid);
    Logger.log('🔍 missingRequired.length: ' + validation.missingRequired.length);
    Logger.log('🔍 missingRequired: ' + JSON.stringify(validation.missingRequired));
    Logger.log('🔍 needsProjectSelection: ' + validation.needsProjectSelection);
    Logger.log('🔍 needsPaymentMethod: ' + validation.needsPaymentMethod);
    Logger.log('🔍 needsCurrency: ' + validation.needsCurrency);
    Logger.log('🔍 needsExchangeRate: ' + validation.needsExchangeRate);
    Logger.log('🔍 enriched.exchangeRate: ' + validation.enriched.exchangeRate);
    Logger.log('🔍 enriched.exchange_rate: ' + validation.enriched.exchange_rate);
    Logger.log('🔍 enriched.payment_method: ' + validation.enriched.payment_method);
    Logger.log('🔍 needsPartyConfirmation: ' + validation.needsPartyConfirmation);
    Logger.log('🔍 needsLoanDueDate: ' + validation.needsLoanDueDate);
    Logger.log('═══════════════════════════════════════');

    return validation;
}


// ==================== تحليل الحركة الكامل ====================

/**
 * تحليل نص المستخدم واستخراج الحركة المالية
 * @param {string} userMessage - رسالة المستخدم
 * @returns {Object} - نتيجة التحليل مع الحركة
 */
function analyzeTransaction(userMessage) {
    Logger.log('═══════════════════════════════════════');
    Logger.log('🤖 analyzeTransaction STARTED');
    Logger.log('🤖 userMessage: ' + userMessage);

    try {
        // تحميل السياق
        Logger.log('🤖 Loading AI context...');
        const context = loadAIContext();

        // استدعاء Gemini
        Logger.log('🤖 Calling Gemini...');
        const aiResult = callGemini(userMessage, context);
        Logger.log('🤖 Gemini returned');
        Logger.log('🤖 aiResult.success: ' + aiResult.success);
        Logger.log('🤖 aiResult keys: ' + Object.keys(aiResult).join(', '));

        // ⭐ تحقق أكثر ذكاءً - إذا كان الرد يحتوي على بيانات الحركة فهو ناجح
        // حتى لو لم يحتوي على success: true صراحة
        const hasTransactionData = aiResult.nature || aiResult.classification || aiResult.amount || aiResult.party;
        const isExplicitFailure = aiResult.success === false;

        Logger.log('🤖 hasTransactionData: ' + hasTransactionData);
        Logger.log('🤖 isExplicitFailure: ' + isExplicitFailure);

        if (isExplicitFailure || (!hasTransactionData && !aiResult.success)) {
            Logger.log('❌ AI result indicates failure or no transaction data');
            Logger.log('❌ aiResult.error: ' + aiResult.error);

            // ⭐ خطة بديلة: إذا فشل Gemini والنص يبدو كمصاريف بنكية، حلّله يدوياً
            const bankFeesFallback = tryParseBankFees_(userMessage);
            if (bankFeesFallback) {
                Logger.log('✅ Bank fees fallback parser succeeded');
                const validation = validateTransaction(bankFeesFallback, context);
                return {
                    success: true,
                    transaction: validation.enriched,
                    validation: validation,
                    needsInput: validation.missingRequired.length > 0,
                    missingFields: validation.missingRequired,
                    warnings: validation.warnings,
                    confidence: 0.85
                };
            }

            // ⭐ خطة بديلة: إذا فشل Gemini والنص يبدو كتغيير عملة، حلّله يدوياً
            const exchangeFallback = tryParseCurrencyExchange_(userMessage);
            if (exchangeFallback) {
                Logger.log('✅ Currency exchange fallback parser succeeded');
                const validation = validateTransaction(exchangeFallback, context);
                return {
                    success: true,
                    transaction: validation.enriched,
                    validation: validation,
                    needsInput: validation.missingRequired.length > 0,
                    missingFields: validation.missingRequired,
                    warnings: validation.warnings,
                    confidence: 0.85
                };
            }

            return {
                success: false,
                error: aiResult.error || AI_CONFIG.AI_MESSAGES.ERROR_PARSE,
                suggestion: aiResult.suggestion
            };
        }

        // التحقق من الحركة وإثرائها
        Logger.log('🤖 Calling validateTransaction...');
        const validation = validateTransaction(aiResult, context);
        Logger.log('🤖 validateTransaction returned');

        const result = {
            success: true,
            transaction: validation.enriched,
            validation: validation,
            needsInput: validation.missingRequired.length > 0,
            missingFields: validation.missingRequired,
            warnings: validation.warnings,
            confidence: aiResult.confidence || 0.8
        };

        Logger.log('🤖 analyzeTransaction result:');
        Logger.log('🤖 - success: ' + result.success);
        Logger.log('🤖 - needsInput: ' + result.needsInput);
        Logger.log('🤖 - missingFields.length: ' + result.missingFields.length);
        Logger.log('🤖 analyzeTransaction FINISHED');
        Logger.log('═══════════════════════════════════════');

        return result;
    } catch (error) {
        Logger.log('❌ Analyze Transaction Error: ' + error.message);
        Logger.log('❌ Stack: ' + error.stack);
        return {
            success: false,
            error: 'خطأ في تحليل البيانات الداخلية: ' + error.message
        };
    }
}


// ==================== دوال مساعدة ====================

/**
 * ⭐ محلل بديل للمصاريف البنكية - يعمل عند فشل Gemini
 * يحلل النص يدوياً ويستخرج البيانات الأساسية
 */
function tryParseBankFees_(text) {
    if (!text) return null;

    // كشف المصاريف البنكية بالكلمات المفتاحية
    const bankKeywords = ['مصاريف بنكية', 'عمولة بنكية', 'رسوم بنكية', 'مصاريف تحويل',
        'عمولة تحويل', 'رسوم حوالة', 'مصاريف البنك', 'عمولة البنك', 'رسوم البنك', 'مصاريف بنك', 'رسوم مصرفية'];
    const isBankFee = bankKeywords.some(kw => text.includes(kw));
    if (!isBankFee) return null;

    Logger.log('🏦 Bank fees fallback parser: detected bank fees in text');

    // استخراج المبلغ (أرقام عربية وإنجليزية)
    let amount = null;
    const normalizedText = text.replace(/[٠-٩]/g, d => '٠١٢٣٤٥٦٧٨٩'.indexOf(d));
    const amountMatch = normalizedText.match(/(\d+(?:[.,]\d+)?)\s*(?:دولار|ليرة|جنيه|USD|TRY|EGP|\$)/i)
        || normalizedText.match(/(?:بقيمة|بمبلغ)\s*(\d+(?:[.,]\d+)?)/i)
        || normalizedText.match(/(\d+(?:[.,]\d+)?)/);
    if (amountMatch) {
        amount = parseFloat(amountMatch[1].replace(',', '.'));
    }
    if (!amount) return null;

    // استخراج العملة
    let currency = 'USD';
    if (/ليرة|TRY|TL/i.test(text)) currency = 'TRY';
    else if (/جنيه|EGP/i.test(text)) currency = 'EGP';

    // استخراج التاريخ
    let dueDate = 'TODAY';
    const dateMatch = text.match(/(?:بتاريخ\s*)?(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
    if (dateMatch) {
        const day = dateMatch[1].padStart(2, '0');
        const month = dateMatch[2].padStart(2, '0');
        const year = dateMatch[3];
        dueDate = `${year}-${month}-${day}`;
    }

    // استخراج الطرف (بعد "مقابل استلام فاتورة X" أو "مقابل تحويل لـ X")
    let party = null;
    let details = '';
    const contextMatch = text.match(/مقابل\s+(.*?)(?:\s+بتاريخ|\s*$)/);
    if (contextMatch) {
        details = contextMatch[1].trim();
        // استخراج اسم الطرف من السياق
        const partyPatterns = [
            /(?:استلام\s+فاتورة|فاتورة)\s+(\S+)/,
            /(?:تحويل\s+ل|حوالة\s+ل|تحويل\s+إلى|حوالة\s+إلى)\s*(.+?)(?:\s+فيلم|\s+مشروع|\s*$)/,
            /(?:حوالة|تحويل)\s+(\S+)/
        ];
        for (const pat of partyPatterns) {
            const m = details.match(pat);
            if (m) {
                party = m[1].trim();
                break;
            }
        }
    }

    Logger.log('🏦 Parsed: amount=' + amount + ', currency=' + currency + ', date=' + dueDate + ', party=' + party + ', details=' + details);

    return {
        success: true,
        is_shared_order: false,
        nature: 'دفعة مصروف',
        classification: 'مصروفات عمومية',
        project: null,
        project_code: null,
        item: 'مصاريف بنكية',
        party: party,
        amount: amount,
        currency: currency,
        payment_method: 'تحويل بنكي',
        payment_term: 'فوري',
        payment_term_weeks: null,
        payment_term_date: null,
        due_date: dueDate,
        details: details,
        unit_count: null,
        exchange_rate: null,
        confidence: 0.85
    };
}

/**
 * ⭐ محلل بديل لتغيير العملة - يعمل عند فشل Gemini
 * يكتشف عمليات الصرافة (صرفت دولار، غيرت عملة، تصريف، إلخ)
 */
function tryParseCurrencyExchange_(text) {
    if (!text) return null;

    // كشف تغيير العملة بالكلمات المفتاحية
    const exchangeKeywords = ['تغيير عملة', 'تصريف عملات', 'غيرت عملة', 'صرفت', 'صرافة', 'تصريف',
        'غيرت دولار', 'غيرت ليرة', 'حولت دولار', 'حولت ليرة',
        'بعت دولار', 'شريت دولار', 'اشتريت دولار',
        'صرفت دولار', 'صرفت ليرة', 'exchange'];
    const isExchange = exchangeKeywords.some(kw => text.includes(kw));
    if (!isExchange) return null;

    Logger.log('💱 Currency exchange fallback parser: detected exchange in text');

    // استخراج المبلغ
    let amount = null;
    const normalizedText = text.replace(/[٠-٩]/g, d => '٠١٢٣٤٥٦٧٨٩'.indexOf(d));
    const amountMatch = normalizedText.match(/(\d+(?:[.,]\d+)?)\s*(?:دولار|ليرة|جنيه|USD|TRY|EGP|\$)/i)
        || normalizedText.match(/(?:بقيمة|بمبلغ|صرفت|غيرت)\s*(\d+(?:[.,]\d+)?)/i)
        || normalizedText.match(/(\d+(?:[.,]\d+)?)/);
    if (amountMatch) {
        amount = parseFloat(amountMatch[1].replace(',', '.'));
    }
    if (!amount) return null;

    // تحديد نوع العملية (بيع أو شراء دولار)
    let classification = 'بيع دولار'; // الافتراضي
    if (/شريت|اشتريت|شراء/.test(text)) {
        classification = 'شراء دولار';
    } else if (/بعت|بيع|صرفت\s*دولار|غيرت\s*دولار|حولت\s*دولار/.test(text)) {
        classification = 'بيع دولار';
    }

    // استخراج العملة المصدر
    let currency = 'USD';
    if (classification === 'بيع دولار') {
        currency = 'USD'; // المبلغ بالدولار
    } else {
        // شراء دولار - المبلغ قد يكون بالليرة
        if (/ليرة|TRY|TL/i.test(text)) currency = 'TRY';
        else currency = 'USD';
    }

    // استخراج سعر الصرف
    let exchangeRate = null;
    const rateMatch = normalizedText.match(/(?:بسعر|بكورس|سعر الصرف|الكورس|على سعر|ع سعر)\s*(\d+(?:[.,]\d+)?)/i)
        || normalizedText.match(/(\d+(?:[.,]\d+)?)\s*(?:كورس|سعر)/i);
    if (rateMatch) {
        exchangeRate = parseFloat(rateMatch[1].replace(',', '.'));
    }

    // استخراج التاريخ
    let dueDate = 'TODAY';
    const dateMatch = text.match(/(?:بتاريخ\s*)?(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
    if (dateMatch) {
        const day = dateMatch[1].padStart(2, '0');
        const month = dateMatch[2].padStart(2, '0');
        const year = dateMatch[3];
        dueDate = `${year}-${month}-${day}`;
    }

    // استخراج التفاصيل
    let details = text.trim();

    Logger.log('💱 Parsed exchange: amount=' + amount + ', currency=' + currency +
        ', classification=' + classification + ', rate=' + exchangeRate + ', date=' + dueDate);

    return {
        success: true,
        is_shared_order: false,
        nature: 'تغيير عملة',
        classification: classification,
        project: null,
        project_code: null,
        item: 'تغيير عملة',
        party: null,
        amount: amount,
        currency: currency,
        payment_method: null,
        payment_term: 'فوري',
        payment_term_weeks: null,
        payment_term_date: null,
        due_date: dueDate,
        details: details,
        unit_count: null,
        exchange_rate: exchangeRate,
        confidence: 0.85
    };
}

/**
 * استنتاج نوع الطرف من طبيعة الحركة
 */
function inferPartyType(nature, classification) {
    if (nature.includes('إيراد') || nature.includes('تحصيل')) {
        return 'عميل';
    }
    if (nature.includes('تمويل') || classification.includes('تمويل') || classification.includes('سلفة')) {
        return 'ممول';
    }
    return 'مورد';
}

/**
 * استنتاج نوع الحركة (مدين/دائن/تسوية)
 */
function inferMovementType(nature) {
    // دائن تسوية = خصم/تسوية من استحقاق بدون حركة نقدية
    const settlementNatures = [
        'تسوية استحقاق مصروف',
        'تسوية استحقاق إيراد'
    ];
    if (settlementNatures.includes(nature)) return CONFIG.MOVEMENT.SETTLEMENT;

    // دائن دفعة = دفع/خروج نقدية أو تحصيل/دخول نقدية
    const creditNatures = [
        'دفعة مصروف',
        'تحصيل إيراد',
        'استلام تمويل',
        'سداد تمويل',
        'تأمين مدفوع للقناة',
        'تحويل داخلي',
        'مصاريف بنكية',
        'استرداد تأمين'
    ];
    if (creditNatures.includes(nature)) return CONFIG.MOVEMENT.CREDIT;

    // مدين استحقاق = دين/التزام ورقي بدون حركة نقدية فعلية
    return CONFIG.MOVEMENT.DEBIT;
}

/**
 * حساب القيمة بالدولار
 */
function calculateUSDAmount(amount, currency, exchangeRate) {
    if (currency === 'USD') {
        return amount;
    }

    const rate = exchangeRate || getDefaultExchangeRate(currency);
    return amount / rate;
}

/**
 * الحصول على سعر الصرف الافتراضي
 */
function getDefaultExchangeRate(currency) {
    // ⭐ أسعار تقريبية افتراضية - يمكن تحديثها من مصدر خارجي
    // آخر تحديث: فبراير 2026
    const rates = {
        'TRY': 38.0,
        'EGP': 50.0,
        'USD': 1.0
    };
    return rates[currency] || 1.0;
}

/**
 * بناء رسالة التأكيد للحركة
 */
function buildTransactionSummary(transaction) {
    // ⭐ دالة مساعدة لتنظيف النصوص من أحرف Markdown التي تسبب أخطاء في Telegram
    function esc(val) {
        if (!val) return '';
        return String(val).replace(/[*_`\[]/g, '');
    }

    const emoji = getTransactionEmoji(transaction.nature);
    const typeLabel = getTypeLabel(transaction.nature);

    let summary = `${emoji} *${esc(typeLabel)}*\n`;
    summary += '━━━━━━━━━━━━━━━━\n';

    // عرض المشروع مع الكود
    if (transaction.project) {
        let projectDisplay = esc(transaction.project);
        if (transaction.project_code) {
            projectDisplay = `${esc(transaction.project)} (${esc(transaction.project_code)})`;
        }
        summary += `🎬 *المشروع:* ${projectDisplay}\n`;
    }

    summary += `📁 *التصنيف:* ${esc(transaction.classification)}\n`;

    if (transaction.item) {
        summary += `📋 *البند:* ${esc(transaction.item)}\n`;
    }

    if (transaction.party) {
        summary += `👤 *الطرف:* ${esc(transaction.party)}`;
        if (transaction.isNewParty) {
            summary += ' (جديد)';
        }
        summary += '\n';
    }

    summary += `💰 *المبلغ:* ${formatNumber(transaction.amount)} ${esc(transaction.currency)}\n`;

    if (transaction.currency !== 'USD') {
        const usdAmount = calculateUSDAmount(transaction.amount, transaction.currency, transaction.exchangeRate);
        summary += `💵 *بالدولار:* ${formatNumber(usdAmount)} USD\n`;
    }

    summary += `📅 *التاريخ:* ${esc(transaction.due_date)}\n`;

    if (transaction.payment_method) {
        summary += `💳 *طريقة الدفع:* ${esc(transaction.payment_method)}\n`;
    }

    // عرض شرط الدفع
    if (transaction.payment_term) {
        let termDisplay = esc(transaction.payment_term);
        if (transaction.payment_term === 'بعد التسليم' && transaction.payment_term_weeks) {
            termDisplay = `بعد التسليم (${transaction.payment_term_weeks} أسبوع)`;
        } else if (transaction.payment_term === 'تاريخ مخصص' && transaction.payment_term_date) {
            termDisplay = `تاريخ مخصص: ${esc(transaction.payment_term_date)}`;
        }
        summary += `⏰ *شرط الدفع:* ${termDisplay}\n`;
    }

    // ⭐ عرض تاريخ استحقاق السلفة/التمويل
    if (transaction.loan_due_date) {
        summary += `📆 *تاريخ السداد:* ${esc(transaction.loan_due_date)}\n`;
    }

    // عرض عدد الوحدات إذا موجود
    if (transaction.unit_count && transaction.unit_count > 0) {
        summary += `📊 *عدد الوحدات:* ${transaction.unit_count}\n`;
    }

    if (transaction.details) {
        summary += `📝 *التفاصيل:* ${esc(transaction.details)}\n`;
    }

    summary += '━━━━━━━━━━━━━━━━';

    return summary;
}

/**
 * الحصول على إيموجي نوع الحركة
 */
function getTransactionEmoji(nature) {
    const emojis = {
        'استحقاق مصروف': '📤',
        'دفعة مصروف': '💸',
        'استحقاق إيراد': '📥',
        'تحصيل إيراد': '💰',
        'تمويل': '🏦',
        'سداد تمويل': '💳'
    };
    return emojis[nature] || '📋';
}

/**
 * الحصول على عنوان نوع الحركة
 */
function getTypeLabel(nature) {
    return nature || 'حركة مالية';
}

/**
 * تنسيق الأرقام
 */
function formatNumber(num) {
    if (!num) return '0';
    return Number(num).toLocaleString('en-US', {
        minimumFractionDigits: 2,
        maximumFractionDigits: 2
    });
}


// ==================== اختبار الـ Agent ====================

/**
 * اختبار تحليل نص
 */
function testAIAgent() {
    const testMessages = [
        'اتفقت مع احمد نايل على رسم فيلم بل كلنتون بقيمة 400 دولار يوم 1/2/2025',
        'دفعت لأحمد 500 دولار رسوم رسم',
        'استلمت 1000 دولار من قناة الجزيرة مقابل فيلم الوثائقي'
    ];

    testMessages.forEach((msg, index) => {
        Logger.log(`\n=== اختبار ${index + 1} ===`);
        Logger.log('النص: ' + msg);

        try {
            const result = analyzeTransaction(msg);
            Logger.log('النتيجة: ' + JSON.stringify(result, null, 2));
        } catch (error) {
            Logger.log('خطأ: ' + error.message);
        }
    });
}

/**
 * اختبار اتصال Gemini API مباشرة
 * شغّل هذه الدالة للتشخيص
 */
function testGeminiConnection() {
    Logger.log('═══════════════════════════════════════');
    Logger.log('=== اختبار اتصال Gemini API ===');
    Logger.log('═══════════════════════════════════════');

    try {
        const apiKey = getGeminiApiKey();
        Logger.log('✅ تم الحصول على API Key');
        Logger.log('API Key (أول 10 أحرف): ' + apiKey.substring(0, 10) + '...');

        const url = `${AI_CONFIG.GEMINI.API_URL}?key=${apiKey}`;
        Logger.log('URL: ' + AI_CONFIG.GEMINI.API_URL);

        // طلب بسيط جداً
        const payload = {
            contents: [{
                parts: [{
                    text: 'قل مرحبا'
                }]
            }]
        };

        Logger.log('جاري إرسال الطلب...');

        const response = UrlFetchApp.fetch(url, {
            method: 'post',
            contentType: 'application/json',
            payload: JSON.stringify(payload),
            muteHttpExceptions: true
        });

        const responseCode = response.getResponseCode();
        const responseText = response.getContentText();

        Logger.log('Response Code: ' + responseCode);

        if (responseCode === 200) {
            Logger.log('✅ الاتصال ناجح!');
            const result = JSON.parse(responseText);
            if (result.candidates && result.candidates[0]) {
                const text = result.candidates[0].content.parts[0].text;
                Logger.log('رد Gemini: ' + text);
            }
        } else {
            Logger.log('❌ خطأ في الاتصال');
            Logger.log('Response: ' + responseText);
        }

    } catch (error) {
        Logger.log('❌ خطأ: ' + error.message);
    }

    Logger.log('═══════════════════════════════════════');
}

/**
 * عرض الموديلات المتاحة لمفتاح API
 * شغّل هذه الدالة لمعرفة الموديلات التي يمكنك استخدامها
 */
function listAvailableModels() {
    Logger.log('═══════════════════════════════════════');
    Logger.log('=== قائمة الموديلات المتاحة (Gemini) ===');

    try {
        const apiKey = getGeminiApiKey();
        const url = `https://generativelanguage.googleapis.com/v1/models?key=${apiKey}`;

        const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
        const result = JSON.parse(response.getContentText());

        if (result.models) {
            Logger.log('✅ تم جلب القائمة بنجاح:');
            result.models.forEach(model => {
                // تصفية الموديلات التي تدعم generateContent
                if (model.supportedGenerationMethods && model.supportedGenerationMethods.includes('generateContent')) {
                    Logger.log(`• ${model.name.replace('models/', '')} (${model.displayName})`);
                }
            });
        } else {
            Logger.log('❌ لم يتم العثور على موديلات: ' + JSON.stringify(result));
        }

    } catch (e) {
        Logger.log('❌ خطأ: ' + e.message);
    }
    Logger.log('═══════════════════════════════════════');
}

