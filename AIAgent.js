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
    let prompt = AI_CONFIG.SYSTEM_PROMPT + '\n\n';

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

        const parsed = JSON.parse(jsonStr);
        Logger.log('✅ JSON parsed successfully');
        return parsed;

    } catch (error) {
        Logger.log('❌ JSON Parse Error: ' + error.message);
        Logger.log('Raw text (first 500 chars): ' + (text || '').substring(0, 500));
        return {
            success: false,
            error: 'فشل في تحليل رد AI: ' + error.message,
            rawResponse: (text || '').substring(0, 200)
        };
    }
}


// ==================== تحميل السياق ====================

/**
 * تحميل السياق الكامل (المشاريع، البنود، الأطراف، الأنواع، التصنيفات)
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
        const ss = SpreadsheetApp.getActiveSpreadsheet();

        // تحميل المشاريع
        context.projects = loadProjects(ss);

        // تحميل البنود + الأنواع + التصنيفات من شيت البنود
        const itemsData = loadItems(ss);
        context.items = itemsData.items;
        context.natures = itemsData.natures;
        context.classifications = itemsData.classifications;

        // تحميل الأطراف
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

    if (!transaction.party) {
        validation.missingRequired.push({
            field: 'party',
            label: 'الطرف',
            message: 'يرجى تحديد اسم الطرف (المورد/العميل/الممول)'
        });
    }

    // التحقق من المشروع للمصروفات المباشرة والإيرادات
    // ⭐ المشروع اختياري - نسأل عنه لكن يمكن التخطي
    const needsProject = ['مصروفات مباشرة', 'ايراد'].includes(transaction.classification);
    Logger.log('🔍 needsProject check:');
    Logger.log('   - classification: "' + transaction.classification + '"');
    Logger.log('   - needsProject: ' + needsProject);
    Logger.log('   - project value: ' + transaction.project);
    Logger.log('   - !transaction.project: ' + !transaction.project);

    if (needsProject && !transaction.project) {
        // ⭐ لا نضيفه كحقل إلزامي، بل كسؤال اختياري
        Logger.log('📋 Project is optional - will ask with skip option');
        validation.needsProjectSelection = true;
    }

    // مطابقة المشروع
    if (transaction.project && context.projects) {
        const projectMatch = matchProject(transaction.project, context.projects);
        if (projectMatch.found) {
            validation.enriched.project = projectMatch.match;
            validation.enriched.project_code = projectMatch.code || transaction.project_code || '';
            validation.enriched.projectScore = projectMatch.score;
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
    const isPayment = nature.includes('دفعة') || nature.includes('تحصيل') || nature.includes('سداد') || nature.includes('استلام');

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

    // مطابقة الطرف
    if (transaction.party && context.parties) {
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

    // مطابقة البند
    if (transaction.item && context.items) {
        const itemMatch = matchItem(transaction.item, context.items);
        if (itemMatch.found) {
            validation.enriched.item = itemMatch.match;
            validation.enriched.itemScore = itemMatch.score;
        }
    }

    // تحويل التاريخ
    if (transaction.due_date === 'TODAY' || !transaction.due_date) {
        validation.enriched.due_date = Utilities.formatDate(new Date(), 'Asia/Istanbul', 'yyyy-MM-dd');
    }

    // ⭐ التحقق من طريقة الدفع (يجب أن تكون محددة)
    if (!transaction.payment_method || transaction.payment_method === 'تحويل بنكي') {
        // إذا لم تحدد أو كانت القيمة الافتراضية، نحتاج تأكيد من المستخدم
        validation.needsPaymentMethod = true;
        validation.enriched.payment_method = transaction.payment_method || null;
    } else {
        // تحويل القيم المختلفة لـ "بنك" أو "خزنة"
        const method = transaction.payment_method.toLowerCase();
        if (method.includes('نقد') || method.includes('كاش') || method.includes('خزن') || method.includes('يد')) {
            validation.enriched.payment_method = 'خزنة';
        } else if (method.includes('بنك') || method.includes('تحويل') || method.includes('حوال')) {
            validation.enriched.payment_method = 'بنك';
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

        // ⭐ إذا كانت العملة غير دولار، يجب تحديد سعر الصرف
        if (transaction.currency !== 'USD' && !transaction.exchange_rate) {
            validation.needsExchangeRate = true;
        }
    }

    // تحويل سعر الصرف من snake_case إلى camelCase
    if (transaction.exchange_rate) {
        validation.enriched.exchangeRate = transaction.exchange_rate;
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
    Logger.log('🔍 needsPartyConfirmation: ' + validation.needsPartyConfirmation);
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
 * استنتاج نوع الحركة (مدين/دائن)
 */
function inferMovementType(nature) {
    // دائن دفعة = دفع/خروج نقدية أو تحصيل/دخول نقدية
    const creditNatures = [
        'دفعة مصروف',
        'تحصيل إيراد',
        'استلام تمويل',
        'سداد تمويل',
        'تأمين مدفوع للقناة',
        'تحويل داخلي'
    ];
    return creditNatures.includes(nature) ? 'دائن دفعة' : 'مدين استحقاق';
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
    // يمكن تحديث هذه القيم أو جلبها من مصدر خارجي
    const rates = {
        'TRY': 32.0,
        'EGP': 50.0,
        'USD': 1.0
    };
    return rates[currency] || 1.0;
}

/**
 * بناء رسالة التأكيد للحركة
 */
function buildTransactionSummary(transaction) {
    const emoji = getTransactionEmoji(transaction.nature);
    const typeLabel = getTypeLabel(transaction.nature);

    let summary = `${emoji} *${typeLabel}*\n`;
    summary += '━━━━━━━━━━━━━━━━\n';

    // عرض المشروع مع الكود
    if (transaction.project) {
        let projectDisplay = transaction.project;
        if (transaction.project_code) {
            projectDisplay = `${transaction.project} (${transaction.project_code})`;
        }
        summary += `🎬 *المشروع:* ${projectDisplay}\n`;
    }

    summary += `📁 *التصنيف:* ${transaction.classification}\n`;

    if (transaction.item) {
        summary += `📋 *البند:* ${transaction.item}\n`;
    }

    summary += `👤 *الطرف:* ${transaction.party}`;
    if (transaction.isNewParty) {
        summary += ' _(جديد)_';
    }
    summary += '\n';

    summary += `💰 *المبلغ:* ${formatNumber(transaction.amount)} ${transaction.currency}\n`;

    if (transaction.currency !== 'USD') {
        const usdAmount = calculateUSDAmount(transaction.amount, transaction.currency, transaction.exchangeRate);
        summary += `💵 *بالدولار:* ${formatNumber(usdAmount)} USD\n`;
    }

    summary += `📅 *التاريخ:* ${transaction.due_date}\n`;

    if (transaction.payment_method) {
        summary += `💳 *طريقة الدفع:* ${transaction.payment_method}\n`;
    }

    // عرض شرط الدفع
    if (transaction.payment_term) {
        let termDisplay = transaction.payment_term;
        if (transaction.payment_term === 'بعد التسليم' && transaction.payment_term_weeks) {
            termDisplay = `بعد التسليم (${transaction.payment_term_weeks} أسبوع)`;
        } else if (transaction.payment_term === 'تاريخ مخصص' && transaction.payment_term_date) {
            termDisplay = `تاريخ مخصص: ${transaction.payment_term_date}`;
        }
        summary += `⏰ *شرط الدفع:* ${termDisplay}\n`;
    }

    if (transaction.details) {
        summary += `📝 *التفاصيل:* ${transaction.details}\n`;
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

