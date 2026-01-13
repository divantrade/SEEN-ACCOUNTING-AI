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
function callGemini(userMessage, context) {
    try {
        const apiKey = getGeminiApiKey();
        const url = `${AI_CONFIG.GEMINI.API_URL}?key=${apiKey}`;

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

        if (responseCode !== 200) {
            Logger.log('Gemini API Error: ' + responseCode + ' - ' + responseText);
            let errorDetails = responseText;
            try {
                const jsonError = JSON.parse(responseText);
                if (jsonError.error && jsonError.error.message) {
                    errorDetails = jsonError.error.message;
                }
            } catch (e) { }

            return {
                success: false,
                error: `خطأ في الاتصال بـ Gemini API (${responseCode}): ${errorDetails}`,
                details: responseText
            };
        }

        const result = JSON.parse(responseText);

        // استخراج النص من الرد
        if (result.candidates && result.candidates[0] && result.candidates[0].content) {
            const text = result.candidates[0].content.parts[0].text;
            return parseGeminiResponse(text);
        }

        return {
            success: false,
            error: 'رد غير متوقع من Gemini'
        };

    } catch (error) {
        Logger.log('Gemini Error: ' + error.message);
        return {
            success: false,
            error: error.message
        };
    }
}

/**
 * بناء الـ prompt الكامل مع السياق
 */
function buildFullPrompt(userMessage, context) {
    let prompt = AI_CONFIG.SYSTEM_PROMPT + '\n\n';

    // إضافة قائمة المشاريع
    if (context.projects && context.projects.length > 0) {
        prompt += '## المشاريع المتاحة:\n';
        context.projects.forEach(p => {
            prompt += `- ${p}\n`;
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
    prompt += '## المطلوب:\nحلل النص أعلاه واستخرج بيانات الحركة المالية بصيغة JSON.';

    return prompt;
}

/**
 * تحليل رد Gemini واستخراج JSON
 */
function parseGeminiResponse(text) {
    try {
        // محاولة استخراج JSON من النص
        let jsonStr = text;

        // إزالة markdown code blocks إن وجدت
        const jsonMatch = text.match(/```(?:json)?\s*([\s\S]*?)\s*```/);
        if (jsonMatch) {
            jsonStr = jsonMatch[1];
        }

        // تنظيف النص
        jsonStr = jsonStr.trim();

        // البحث عن أول { وآخر }
        const startIndex = jsonStr.indexOf('{');
        const endIndex = jsonStr.lastIndexOf('}');

        if (startIndex !== -1 && endIndex !== -1) {
            jsonStr = jsonStr.substring(startIndex, endIndex + 1);
        }

        const parsed = JSON.parse(jsonStr);
        return parsed;

    } catch (error) {
        Logger.log('JSON Parse Error: ' + error.message);
        Logger.log('Raw text: ' + text);
        return {
            success: false,
            error: 'فشل في تحليل رد AI',
            rawResponse: text
        };
    }
}


// ==================== تحميل السياق ====================

/**
 * تحميل السياق الكامل (المشاريع، البنود، الأطراف)
 */
function loadAIContext() {
    const context = {
        projects: [],
        items: [],
        parties: []
    };

    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();

        // تحميل المشاريع
        context.projects = loadProjects(ss);

        // تحميل البنود
        context.items = loadItems(ss);

        // تحميل الأطراف
        context.parties = loadParties(ss);

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
            const projectName = data[i][1]; // اسم المشروع في العمود الثاني
            if (projectName && projectName.toString().trim()) {
                projects.push(projectName.toString().trim());
            }
        }

    } catch (error) {
        Logger.log('Load Projects Error: ' + error.message);
    }

    return projects;
}

/**
 * تحميل قائمة البنود
 */
function loadItems(ss) {
    const items = [];

    try {
        const sheet = ss.getSheetByName(CONFIG.SHEETS.ITEMS);
        if (!sheet) return items;

        const data = sheet.getDataRange().getValues();

        // تخطي الهيدر، العمود الأول = اسم البند
        for (let i = 1; i < data.length; i++) {
            const itemName = data[i][0];
            if (itemName && itemName.toString().trim()) {
                items.push(itemName.toString().trim());
            }
        }

    } catch (error) {
        Logger.log('Load Items Error: ' + error.message);
    }

    return items;
}

/**
 * تحميل قائمة الأطراف
 */
function loadParties(ss) {
    const parties = [];

    try {
        const sheet = ss.getSheetByName(CONFIG.SHEETS.PARTIES);
        if (!sheet) return parties;

        const data = sheet.getDataRange().getValues();

        // تخطي الهيدر
        for (let i = 1; i < data.length; i++) {
            const partyName = data[i][0]; // اسم الطرف
            const partyType = data[i][1]; // نوع الطرف

            if (partyName && partyName.toString().trim()) {
                parties.push({
                    name: partyName.toString().trim(),
                    type: partyType ? partyType.toString().trim() : 'مورد'
                });
            }
        }

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
 */
function matchProject(projectName, projectsList) {
    if (!projectName || !projectsList || projectsList.length === 0) {
        return { found: false, matches: [] };
    }

    // بحث مطابق تماماً
    const exactMatch = projectsList.find(p =>
        normalizeArabicText(p) === normalizeArabicText(projectName)
    );

    if (exactMatch) {
        return { found: true, match: exactMatch, score: 1.0 };
    }

    // بحث ذكي
    const results = fuzzySearchInArray(projectName, projectsList, 0.5);

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

/**
 * مطابقة اسم الطرف مع الأطراف الموجودة
 */
function matchParty(partyName, partiesList) {
    if (!partyName || !partiesList || partiesList.length === 0) {
        return { found: false, matches: [] };
    }

    const partyNames = partiesList.map(p => p.name);

    // بحث مطابق تماماً
    const exactIndex = partyNames.findIndex(p =>
        normalizeArabicText(p) === normalizeArabicText(partyName)
    );

    if (exactIndex !== -1) {
        return {
            found: true,
            match: partiesList[exactIndex],
            score: 1.0
        };
    }

    // بحث ذكي
    const results = fuzzySearchInArray(partyName, partyNames, 0.5);

    if (results.length > 0) {
        const matchedParties = results.map(r => {
            const party = partiesList.find(p => p.name === r.item);
            return { ...party, score: r.score };
        });

        return {
            found: true,
            match: matchedParties[0],
            score: matchedParties[0].score,
            alternatives: matchedParties.slice(1, 4)
        };
    }

    return { found: false, matches: [] };
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
    const needsProject = ['مصروفات مباشرة', 'ايراد'].includes(transaction.classification);
    if (needsProject && !transaction.project) {
        validation.missingRequired.push({
            field: 'project',
            label: 'المشروع',
            message: 'يرجى تحديد المشروع'
        });
    }

    // مطابقة المشروع
    if (transaction.project && context.projects) {
        const projectMatch = matchProject(transaction.project, context.projects);
        if (projectMatch.found) {
            validation.enriched.project = projectMatch.match;
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
                suggestions: context.projects.slice(0, 5)
            });
        }
    }

    // مطابقة الطرف
    if (transaction.party && context.parties) {
        const partyMatch = matchParty(transaction.party, context.parties);
        if (partyMatch.found) {
            validation.enriched.party = partyMatch.match.name;
            validation.enriched.partyType = partyMatch.match.type;
            validation.enriched.partyScore = partyMatch.score;
            if (partyMatch.score < 0.9) {
                validation.warnings.push({
                    field: 'party',
                    message: `هل تقصد "${partyMatch.match.name}"؟`,
                    alternatives: partyMatch.alternatives
                });
            }
        } else {
            validation.enriched.isNewParty = true;
            validation.warnings.push({
                field: 'party',
                message: `الطرف "${transaction.party}" غير موجود - سيتم إضافته كطرف جديد`,
                isNew: true
            });
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

    // تعيين القيم الافتراضية
    if (!validation.enriched.currency) {
        validation.enriched.currency = AI_CONFIG.INFERENCE_RULES.DEFAULTS.CURRENCY;
    }
    if (!validation.enriched.payment_method) {
        validation.enriched.payment_method = AI_CONFIG.INFERENCE_RULES.DEFAULTS.PAYMENT_METHOD;
    }

    validation.isValid = validation.missingRequired.length === 0;

    return validation;
}


// ==================== تحليل الحركة الكامل ====================

/**
 * تحليل نص المستخدم واستخراج الحركة المالية
 * @param {string} userMessage - رسالة المستخدم
 * @returns {Object} - نتيجة التحليل مع الحركة
 */
function analyzeTransaction(userMessage) {
    // تحميل السياق
    const context = loadAIContext();

    // استدعاء Gemini
    const aiResult = callGemini(userMessage, context);

    if (!aiResult.success) {
        return {
            success: false,
            error: aiResult.error || AI_CONFIG.AI_MESSAGES.ERROR_PARSE,
            suggestion: aiResult.suggestion
        };
    }

    // التحقق من الحركة وإثرائها
    const validation = validateTransaction(aiResult, context);

    return {
        success: true,
        transaction: validation.enriched,
        validation: validation,
        needsInput: validation.missingRequired.length > 0,
        missingFields: validation.missingRequired,
        warnings: validation.warnings,
        confidence: aiResult.confidence || 0.8
    };
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
    const creditNatures = ['دفعة مصروف', 'تحصيل إيراد', 'استلام تمويل', 'سداد تمويل'];
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

    if (transaction.project) {
        summary += `🎬 *المشروع:* ${transaction.project}\n`;
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

