// ==================== المحرك الذكي الجديد (Smart Agent) ====================
/**
 * محرك AI Agent لنظام SEEN المحاسبي
 * يستخدم Gemini Function Calling بدلاً من تحليل النص + if/else
 *
 * الفرق عن النظام القديم:
 * - القديم: Gemini يعيد JSON ← الكود يصحح ← الكود يسأل المستخدم
 * - الجديد: Gemini يقرر ← يبحث ← يسأل ← يحفظ (كله عبر أدوات)
 *
 * ⚠️ هذا الملف مستقل تماماً عن AIAgent.js - لا يعدله ولا يتعارض معه
 */

// ==================== الثوابت ====================

// ⚡ معرف المستخدم الحالي (يُعيّن قبل كل تحليل لجلب تفضيلاته)
var CURRENT_AGENT_CHAT_ID_ = null;

var SMART_AGENT_CONFIG = {
    // الحد الأقصى لدورات التفكير (لمنع الحلقات اللانهائية)
    MAX_TURNS: 12,

    // الحد الأقصى لسؤال المستخدم
    MAX_USER_ASKS: 3,

    // إعدادات Gemini المحسّنة للـ Agent
    GENERATION_CONFIG: {
        temperature: 0.1,
        topP: 0.8,
        topK: 40,
        maxOutputTokens: 4096
    }
};

// ==================== System Prompt للـ Agent ====================

/**
 * بناء system prompt مُحسّن للـ Agent
 * أقصر وأدق من البرومبت القديم لأن الأدوات تتولى التفاصيل
 */
function buildAgentSystemPrompt_() {
    var today = Utilities.formatDate(new Date(), CONFIG.COMPANY.TIMEZONE, 'yyyy-MM-dd');
    var dayName = Utilities.formatDate(new Date(), CONFIG.COMPANY.TIMEZONE, 'EEEE');

    // تحميل قوائم مختصرة لمساعدة Gemini
    var contextSummary = buildContextSummary_();

    return `أنت محاسب ذكي في نظام SEEN لإنتاج الأفلام الوثائقية.
مهمتك: تحليل رسائل المستخدم العربية واستخراج الحركات المالية.

📅 اليوم: ${today} (${dayName})

${contextSummary}

## قواعد ذهبية:
1. "اتفقت/حجزت/طلبت/يحاسب بعد X يوم" = استحقاق (دين ورقي، ليس دفع فعلي)
2. "دفعت/سددت/حولت/أرسلت" = دفعة (حركة نقدية فعلية)
3. "استلمت/قبضت/وصلني" = تحصيل إيراد
4. المرتبات والرواتب = دائماً "مصروفات عمومية" (وليس مصروفات مباشرة)
5. المصاريف البنكية = "دفعة مصروف" + "مصروفات عمومية" + بند "مصاريف بنكية"
6. تغيير العملة = لا يحتاج طرف ولا مشروع، يحتاج سعر الصرف
7. التحويل الداخلي = لا يحتاج طرف ولا مشروع
8. طريقة الدفع = لا تفترض أبداً، اسأل المستخدم إذا لم يذكرها (إلا في الدفعات البنكية)
9. المصروفات المباشرة والإيرادات = تحتاج مشروع
10. "بعد التسليم" = شرط دفع يحتاج تحديد عدد أسابيع

## أنواع الحركات (طبيعة الحركة):
- استحقاق مصروف | دفعة مصروف | تسوية استحقاق مصروف
- استحقاق إيراد | تحصيل إيراد | تسوية استحقاق إيراد
- تمويل | استلام تمويل | سداد تمويل
- تأمين مدفوع للقناة | استرداد تأمين من القناة
- تحويل داخلي | تغيير عملة | مصاريف بنكية

## التصنيفات المسموحة:
- مصروفات مباشرة | مصروفات عمومية
- إيرادات أفلام | إيرادات خدمات
- تمويل طويل الأجل | سلفة قصيرة
- تأمينات مدفوعة | تحويل داخلي
- بيع دولار | شراء دولار

## العملات: USD (دولار) | TRY (ليرة تركية) | EGP (جنيه مصري)
## طرق الدفع: نقدي | تحويل بنكي | شيك | بطاقة
## شروط الدفع: فوري | بعد التسليم | تاريخ مخصص

## وحدات الإنتاج:
- تصوير → مقابلة | مونتاج/مكساج/دوبلاج/تلوين → دقيقة
- جرافيك/رسم → رسمة | فيكسر → ضيف | تعليق صوتي → دقيقة
- ساعة = 60 دقيقة | ساعتين = 120 دقيقة

## تعليمات العمل (اتبعها بالترتيب بدقة):
1. حلل رسالة المستخدم واستخرج البيانات المالية
2. **إلزامي**: إذا ذُكر اسم طرف، استخدم search_parties فوراً قبل أي شيء آخر
   - النتيجة ستعطيك: نوع الطرف + تخصصه + المشاريع التي عمل فيها
   - استخدم التخصص لتحديد البند تلقائياً (مثال: تخصص "مونتاج" = بند "مونتاج")
   - استخدم نوع الطرف لتحديد التصنيف (مورد = مصروفات مباشرة عادةً)
   - إذا وجدت مشاريع سابقة (party_projects)، اعرضها على المستخدم للاختيار
3. استخدم search_items للتأكد من البند (إذا لم يُستنتج من تخصص الطرف)
4. استخدم search_projects للتأكد من اسم المشروع (إذا اختار المستخدم أو ذكر مشروعاً)
5. **مهم**: لا تسأل المستخدم عن معلومة يمكنك استنتاجها من نتائج البحث
6. إذا بقيت معلومات ناقصة لا يمكن استنتاجها، استخدم ask_user (اجمع عدة أسئلة في سؤال واحد)
7. عندما تكتمل البيانات، استخدم show_confirmation لعرض الملخص
8. لا تحفظ إلا بعد تأكيد المستخدم عبر show_confirmation

## قواعد الاستنتاج التلقائي:
- طرف تخصصه "مونتاج/مكساج/تلوين/جرافيك/تصوير" = التصنيف "مصروفات مباشرة"
- طرف نوعه "ممول" = الطبيعة "تمويل" أو "سداد تمويل"
- طرف نوعه "عميل" = الطبيعة "إيراد"
- إذا الطرف عمل في مشروع واحد فقط، اقترحه تلقائياً
- إذا المستخدم قال "شوف أفلام X" أو "مشاريع X"، ابحث في party_projects

## التعامل مع نتائج البحث:
- إذا search_parties أرجع partial_match=true مع similar_parties: اعرض الأسماء المشابهة واسأل المستخدم "هل تقصد أحد هؤلاء؟"
- إذا search_parties أرجع found=true: استخدم أول نتيجة واستنتج التصنيف من التخصص
- إذا search_parties أرجع found=false بدون similar_parties: اسأل المستخدم هل يريد إضافته كطرف جديد
- **ممنوع** أن تقول "لا يمكنني البحث" - دائماً استخدم الأدوات المتاحة
- **ممنوع** أن تسأل عن التصنيف أو البند إذا كان تخصص الطرف يحدده

## ملاحظات:
- الأرقام العربية (٠١٢٣٤٥٦٧٨٩) = أرقام إنجليزية (0123456789)
- "بكرة/غداً" = ${getNextDay_()}
- "أمس/إمبارح" = ${getYesterday_()}
- "بعد شهر" = تاريخ مخصص: ${getDateAfterMonths_(1)}
- "بعد أسبوع" = تاريخ مخصص: ${getDateAfterWeeks_(1)}
- العملة الافتراضية USD إذا لم تُذكر
- سعر الصرف الافتراضي: ${getDefaultExchangeRates_()}
${getUserPreferencesPrompt_()}`;
}

// ==================== المحرك الرئيسي ====================

/**
 * تحليل رسالة المستخدم باستخدام الـ Agent الذكي
 * هذه الدالة تحل محل analyzeTransaction() + validateTransaction() + handleMissingFields()
 *
 * @param {string} userMessage - رسالة المستخدم
 * @param {Object} conversationHistory - سجل المحادثة السابقة (اختياري)
 * @returns {Object} - نتيجة التحليل مع الإجراء المطلوب
 */
function smartAnalyze(userMessage, conversationHistory) {
    Logger.log('🧠 Smart Agent: بدء تحليل "' + userMessage + '"');

    try {
        // بناء سجل المحادثة
        var contents = buildConversationContents_(userMessage, conversationHistory);

        // استدعاء Gemini مع Function Calling
        var result = callGeminiWithTools_(contents);

        if (!result.success) {
            Logger.log('❌ Smart Agent: فشل الاتصال - ' + result.error);
            // fallback للنظام القديم
            return fallbackToLegacy_(userMessage);
        }

        // معالجة رد Gemini
        return processAgentResponse_(result, contents);

    } catch (e) {
        Logger.log('❌ Smart Agent خطأ: ' + e.message);
        return fallbackToLegacy_(userMessage);
    }
}

/**
 * متابعة المحادثة بعد رد المستخدم
 * تُستدعى عندما يرد المستخدم على سؤال من الـ Agent
 *
 * @param {string} userReply - رد المستخدم
 * @param {Object} session - جلسة المحادثة المحفوظة
 * @returns {Object} - الإجراء التالي
 */
function smartContinue(userReply, session) {
    Logger.log('🧠 Smart Agent: متابعة بـ "' + userReply + '"');

    try {
        // إضافة رد المستخدم لسجل المحادثة
        var history = session.agentHistory || [];
        history.push({
            role: 'user',
            parts: [{ text: userReply }]
        });

        // استدعاء Gemini مع السجل الكامل
        var result = callGeminiWithTools_(history);

        if (!result.success) {
            return { action: 'ERROR', error: result.error };
        }

        return processAgentResponse_(result, history);

    } catch (e) {
        Logger.log('❌ Smart Agent Continue خطأ: ' + e.message);
        return { action: 'ERROR', error: e.message };
    }
}

// ==================== استدعاء Gemini مع Function Calling ====================

/**
 * استدعاء Gemini API مع دعم Function Calling
 * @param {Array} contents - محتوى المحادثة
 * @returns {Object} - رد Gemini
 */
function callGeminiWithTools_(contents) {
    var apiKey = getGeminiApiKey();
    var props = PropertiesService.getScriptProperties();
    var CACHED_MODEL_KEY = 'GEMINI_WORKING_MODEL';
    var cachedModel = props.getProperty(CACHED_MODEL_KEY);

    // الموديلات التي تدعم Function Calling
    var modelsToTry = [];
    var fcModels = [
        'gemini-2.0-flash',
        'gemini-2.0-flash-001',
        'gemini-1.5-flash',
        'gemini-1.5-pro'
    ];

    if (cachedModel && fcModels.indexOf(cachedModel) !== -1) {
        modelsToTry.push(cachedModel);
    }
    fcModels.forEach(function(m) {
        if (modelsToTry.indexOf(m) === -1) modelsToTry.push(m);
    });

    var lastError = null;

    for (var i = 0; i < modelsToTry.length; i++) {
        var modelName = modelsToTry[i];
        Logger.log('🔄 Smart Agent: تجربة ' + modelName);

        try {
            var url = AI_CONFIG.GEMINI.BASE_URL + modelName + ':generateContent?key=' + apiKey;

            var payload = {
                contents: contents,
                tools: [{
                    functionDeclarations: AGENT_TOOL_DECLARATIONS
                }],
                toolConfig: {
                    functionCallingConfig: {
                        mode: 'AUTO'
                    }
                },
                systemInstruction: {
                    parts: [{ text: buildAgentSystemPrompt_() }]
                },
                generationConfig: SMART_AGENT_CONFIG.GENERATION_CONFIG,
                safetySettings: AI_CONFIG.GEMINI.SAFETY_SETTINGS
            };

            var response = UrlFetchApp.fetch(url, {
                method: 'post',
                contentType: 'application/json',
                payload: JSON.stringify(payload),
                muteHttpExceptions: true
            });

            var responseCode = response.getResponseCode();

            if (responseCode === 200) {
                var result = JSON.parse(response.getContentText());

                if (result.candidates && result.candidates[0] && result.candidates[0].content) {
                    // حفظ الموديل الناجح
                    if (modelName !== cachedModel) {
                        props.setProperty(CACHED_MODEL_KEY, modelName);
                    }

                    Logger.log('✅ Smart Agent: نجاح مع ' + modelName);
                    return {
                        success: true,
                        content: result.candidates[0].content,
                        model: modelName
                    };
                }

                lastError = 'بنية رد غير متوقعة';
            } else {
                lastError = 'خطأ ' + responseCode;
                Logger.log('❌ Smart Agent: فشل ' + modelName + ' - ' + responseCode);
            }

        } catch (e) {
            lastError = e.message;
            Logger.log('❌ Smart Agent: خطأ مع ' + modelName + ': ' + e.message);
        }
    }

    return { success: false, error: 'فشل الاتصال بجميع الموديلات: ' + lastError };
}

// ==================== معالجة رد الـ Agent ====================

/**
 * معالجة رد Gemini واتخاذ الإجراء المناسب
 * @param {Object} geminiResult - رد Gemini
 * @param {Array} currentContents - سجل المحادثة الحالي
 * @returns {Object} - الإجراء المطلوب
 */
function processAgentResponse_(geminiResult, currentContents) {
    var content = geminiResult.content;
    var parts = content.parts || [];

    // تحديث سجل المحادثة برد الـ model
    var updatedHistory = currentContents.slice();
    updatedHistory.push({ role: 'model', parts: parts });

    // ⚡ تجميع كل function calls في رد واحد وتنفيذها دفعة واحدة
    var functionCalls = [];
    var textMessage = null;

    for (var i = 0; i < parts.length; i++) {
        var part = parts[i];
        if (part.functionCall) {
            functionCalls.push(part.functionCall);
        } else if (part.text) {
            textMessage = part.text;
        }
    }

    // إذا لا يوجد function calls - أرسل النص
    if (functionCalls.length === 0) {
        if (textMessage) {
            return {
                action: 'SEND_MESSAGE',
                message: textMessage,
                agentHistory: updatedHistory
            };
        }
        return { action: 'ERROR', error: 'رد فارغ من AI' };
    }

    // ⚡ تنفيذ جميع الأدوات دفعة واحدة
    return handleBatchFunctionCalls_(functionCalls, updatedHistory);
}

/**
 * ⚡ معالجة عدة استدعاءات أدوات دفعة واحدة
 * بدلاً من: tool1 → Gemini → tool2 → Gemini → tool3 → Gemini
 * الآن: tool1 + tool2 + tool3 → Gemini (طلب واحد)
 *
 * @param {Array} functionCalls - قائمة استدعاءات الأدوات
 * @param {Array} history - سجل المحادثة
 * @returns {Object} - الإجراء المطلوب
 */
function handleBatchFunctionCalls_(functionCalls, history) {
    var toolResults = [];
    var userAction = null;

    // تنفيذ جميع الأدوات
    for (var i = 0; i < functionCalls.length; i++) {
        var fc = functionCalls[i];
        var toolName = fc.name;
        var args = fc.args || {};

        Logger.log('🔧 Smart Agent يستدعي: ' + toolName + (functionCalls.length > 1 ? ' (' + (i + 1) + '/' + functionCalls.length + ')' : ''));

        var toolResult = executeAgentTool_(toolName, args);

        // أدوات خاصة تحتاج تفاعل المستخدم - أوقف التنفيذ
        if (toolResult.action === 'WAIT_FOR_USER') {
            userAction = {
                action: 'ASK_USER',
                question: toolResult.question,
                options: toolResult.options,
                field_name: toolResult.field_name,
                agentHistory: history,
                pendingToolResponse: { name: toolName, response: toolResult }
            };
            break;
        }

        if (toolResult.action === 'SHOW_CONFIRMATION') {
            userAction = {
                action: 'SHOW_CONFIRMATION',
                transaction: toolResult.transaction,
                agentHistory: history
            };
            break;
        }

        // أداة عادية - اجمع النتيجة
        toolResults.push({ name: toolName, result: toolResult });
    }

    // إذا أداة تحتاج تفاعل المستخدم
    if (userAction) return userAction;

    // ⚡ إرسال كل النتائج لـ Gemini في طلب واحد
    return continueAfterBatchExecution_(toolResults, history);
}

/**
 * ⚡ متابعة بعد تنفيذ عدة أدوات (إرسال كل النتائج لـ Gemini مرة واحدة)
 */
function continueAfterBatchExecution_(toolResults, history) {
    var updatedHistory = history.slice();

    // إضافة كل نتائج الأدوات في رسالة واحدة (الصيغة الصحيحة لـ Gemini API)
    var functionParts = [];
    for (var i = 0; i < toolResults.length; i++) {
        functionParts.push({
            functionResponse: {
                name: toolResults[i].name,
                response: toolResults[i].result
            }
        });
    }
    updatedHistory.push({
        role: 'function',
        parts: functionParts
    });

    // حماية من الحلقات اللانهائية
    var turnCount = 0;
    for (var j = 0; j < updatedHistory.length; j++) {
        if (updatedHistory[j].role === 'model') turnCount++;
    }
    if (turnCount >= SMART_AGENT_CONFIG.MAX_TURNS) {
        Logger.log('⚠️ Smart Agent: وصل الحد الأقصى للدورات');
        return {
            action: 'ERROR',
            error: 'تم تجاوز الحد الأقصى لمحاولات التحليل',
            agentHistory: updatedHistory
        };
    }

    // استدعاء Gemini مرة واحدة مع كل النتائج
    var nextResult = callGeminiWithTools_(updatedHistory);

    if (!nextResult.success) {
        return { action: 'ERROR', error: nextResult.error, agentHistory: updatedHistory };
    }

    return processAgentResponse_(nextResult, updatedHistory);
}

/**
 * متابعة المحادثة بعد تنفيذ أداة واحدة (تستخدم الدالة الموحدة)
 */
function continueAfterToolExecution_(toolName, toolResult, history) {
    return continueAfterBatchExecution_([{ name: toolName, result: toolResult }], history);
}

// ==================== بناء المحادثة ====================

/**
 * بناء محتوى المحادثة الأولى
 */
function buildConversationContents_(userMessage, conversationHistory) {
    if (conversationHistory && conversationHistory.length > 0) {
        // متابعة محادثة سابقة
        var contents = conversationHistory.slice();
        contents.push({
            role: 'user',
            parts: [{ text: userMessage }]
        });
        return contents;
    }

    // محادثة جديدة
    return [{
        role: 'user',
        parts: [{ text: userMessage }]
    }];
}

// ==================== بناء ملخص السياق ====================

/**
 * بناء ملخص مختصر بالمشاريع والأطراف لمساعدة Gemini في اتخاذ قرارات أفضل
 * بدلاً من الاعتماد فقط على أدوات البحث (التي تسبب دورات كثيرة)
 */
function buildContextSummary_() {
    var cache = CacheService.getScriptCache();
    var cached = cache.get('AGENT_CONTEXT_SUMMARY');
    if (cached) return cached;

    var summary = '';
    try {
        var ss = SpreadsheetApp.getActiveSpreadsheet();

        // تحميل الأطراف (اسم + نوع + تخصص)
        var partiesSheet = ss.getSheetByName(CONFIG.SHEETS.PARTIES);
        if (partiesSheet && partiesSheet.getLastRow() > 1) {
            var partiesData = partiesSheet.getRange(2, 1, Math.min(partiesSheet.getLastRow() - 1, 50), 3).getValues();
            var partiesList = [];
            for (var i = 0; i < partiesData.length; i++) {
                var name = String(partiesData[i][0] || '').trim();
                var type = String(partiesData[i][1] || '').trim();
                var spec = String(partiesData[i][2] || '').trim();
                if (name) {
                    partiesList.push(name + ' (' + type + (spec ? '/' + spec : '') + ')');
                }
            }
            if (partiesList.length > 0) {
                summary += '## الأطراف المسجلين:\n' + partiesList.join(' | ') + '\n\n';
            }
        }

        // تحميل المشاريع (كود + اسم)
        var projectsSheet = ss.getSheetByName(CONFIG.SHEETS.PROJECTS);
        if (projectsSheet && projectsSheet.getLastRow() > 1) {
            var projData = projectsSheet.getRange(2, 1, Math.min(projectsSheet.getLastRow() - 1, 30), 2).getValues();
            var projList = [];
            for (var j = 0; j < projData.length; j++) {
                var code = String(projData[j][0] || '').trim();
                var pname = String(projData[j][1] || '').trim();
                if (code && pname) projList.push(code + '-' + pname);
            }
            if (projList.length > 0) {
                summary += '## المشاريع المتاحة:\n' + projList.join(' | ') + '\n\n';
            }
        }

    } catch (e) {
        Logger.log('⚠️ buildContextSummary_ error: ' + e.message);
    }

    // Cache لمدة 10 دقائق
    if (summary) cache.put('AGENT_CONTEXT_SUMMARY', summary, 600);
    return summary;
}

// ==================== Fallback للنظام القديم ====================

/**
 * في حالة فشل الـ Agent، نستخدم النظام القديم
 */
function fallbackToLegacy_(userMessage) {
    Logger.log('⚠️ Smart Agent: fallback للنظام القديم');
    try {
        var result = analyzeTransaction(userMessage);
        if (result && result.success) {
            return {
                action: 'LEGACY_RESULT',
                result: result
            };
        }
    } catch (e) {
        Logger.log('❌ Legacy fallback فشل أيضاً: ' + e.message);
    }

    return {
        action: 'ERROR',
        error: 'عذراً، لم أتمكن من فهم الحركة. حاول كتابتها بصيغة أوضح.'
    };
}

// ==================== تعلم تفضيلات المستخدم ====================

/**
 * جلب تفضيلات المستخدم من آخر 5 حركات محفوظة له
 * يساعد الذكاء الاصطناعي على تقليل الأسئلة المتكررة
 */
function getUserPreferencesPrompt_() {
    if (!CURRENT_AGENT_CHAT_ID_) return '';

    var cache = CacheService.getScriptCache();
    var cacheKey = 'USER_PREFS_' + CURRENT_AGENT_CHAT_ID_;
    var cached = cache.get(cacheKey);
    if (cached) return cached;

    try {
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var sheet = ss.getSheetByName(CONFIG.SHEETS.TRANSACTIONS);
        if (!sheet) return '';

        var data = sheet.getDataRange().getValues();
        var userTransactions = [];

        // البحث عن حركات هذا المستخدم (من الأحدث للأقدم)
        for (var i = data.length - 1; i >= 1 && userTransactions.length < 5; i--) {
            // عمود المُدخل أو chat_id - نبحث في التفاصيل أو نستخدم آخر 5 حركات عموماً
            var nature = data[i][2];    // طبيعة الحركة
            var project = data[i][5];   // اسم المشروع
            var item = data[i][6];      // البند
            var party = data[i][8];     // الطرف
            var currency = data[i][11]; // العملة
            var payMethod = data[i][15]; // طريقة الدفع

            if (nature && party) {
                userTransactions.push({
                    nature: nature, project: project, item: item,
                    party: party, currency: currency, payMethod: payMethod
                });
            }
        }

        if (userTransactions.length === 0) return '';

        // تحليل الأنماط
        var partyCount = {};
        var currencyCount = {};
        var payMethodCount = {};
        var projectCount = {};

        for (var j = 0; j < userTransactions.length; j++) {
            var t = userTransactions[j];
            if (t.party) partyCount[t.party] = (partyCount[t.party] || 0) + 1;
            if (t.currency) currencyCount[t.currency] = (currencyCount[t.currency] || 0) + 1;
            if (t.payMethod) payMethodCount[t.payMethod] = (payMethodCount[t.payMethod] || 0) + 1;
            if (t.project) projectCount[t.project] = (projectCount[t.project] || 0) + 1;
        }

        var topParties = Object.keys(partyCount).sort(function(a, b) { return partyCount[b] - partyCount[a]; }).slice(0, 3);
        var topCurrency = Object.keys(currencyCount).sort(function(a, b) { return currencyCount[b] - currencyCount[a]; })[0];
        var topPayMethod = Object.keys(payMethodCount).sort(function(a, b) { return payMethodCount[b] - payMethodCount[a]; })[0];
        var topProjects = Object.keys(projectCount).sort(function(a, b) { return projectCount[b] - projectCount[a]; }).slice(0, 3);

        var prompt = '\n## تفضيلات المستخدم (من آخر الحركات):';
        if (topParties.length > 0) prompt += '\n- الأطراف المتكررة: ' + topParties.join('، ');
        if (topProjects.length > 0) prompt += '\n- المشاريع المتكررة: ' + topProjects.join('، ');
        if (topCurrency) prompt += '\n- العملة الأكثر استخداماً: ' + topCurrency;
        if (topPayMethod) prompt += '\n- طريقة الدفع المعتادة: ' + topPayMethod;
        prompt += '\n- استخدم هذه كقيم مقترحة إذا لم يحدد المستخدم، لكن لا تفترض بدون سؤال في الحقول الإلزامية';

        cache.put(cacheKey, prompt, 1800); // cache لمدة 30 دقيقة
        return prompt;

    } catch (e) {
        Logger.log('⚠️ فشل جلب تفضيلات المستخدم: ' + e.message);
        return '';
    }
}

// ==================== أسعار الصرف الديناميكية ====================

/**
 * جلب أسعار الصرف من آخر حركة تغيير عملة في دفتر الحركات
 * مع cache لمدة ساعة لتجنب القراءة المتكررة
 */
function getDefaultExchangeRates_() {
    var cache = CacheService.getScriptCache();
    var cached = cache.get('EXCHANGE_RATES_PROMPT');
    if (cached) return cached;

    var rates = { TRY: 38.0, EGP: 50.0 }; // قيم افتراضية احتياطية

    try {
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var sheet = ss.getSheetByName(CONFIG.SHEETS.TRANSACTIONS);
        if (!sheet) return 'TRY=' + rates.TRY + ', EGP=' + rates.EGP;

        var data = sheet.getDataRange().getValues();
        // البحث من الأسفل لأعلى عن آخر حركة بسعر صرف
        for (var i = data.length - 1; i >= 1; i--) {
            var currency = data[i][11]; // عمود العملة (L)
            var exRate = data[i][12];   // عمود سعر الصرف (M)
            if (exRate && Number(exRate) > 0) {
                if (currency === 'TRY' || currency === 'ليرة') {
                    rates.TRY = Number(exRate);
                } else if (currency === 'EGP' || currency === 'جنيه مصري') {
                    rates.EGP = Number(exRate);
                }
                // وجدنا كلا السعرين
                if (rates.TRY !== 38.0 && rates.EGP !== 50.0) break;
            }
        }
    } catch (e) {
        Logger.log('⚠️ فشل جلب أسعار الصرف: ' + e.message);
    }

    var result = 'TRY=' + rates.TRY + ', EGP=' + rates.EGP;
    cache.put('EXCHANGE_RATES_PROMPT', result, 3600); // cache لمدة ساعة
    return result;
}

// ==================== دوال مساعدة للتواريخ ====================

function getNextDay_() {
    var d = new Date();
    d.setDate(d.getDate() + 1);
    return Utilities.formatDate(d, CONFIG.COMPANY.TIMEZONE, 'yyyy-MM-dd');
}

function getYesterday_() {
    var d = new Date();
    d.setDate(d.getDate() - 1);
    return Utilities.formatDate(d, CONFIG.COMPANY.TIMEZONE, 'yyyy-MM-dd');
}

function getDateAfterMonths_(months) {
    var d = new Date();
    d.setMonth(d.getMonth() + months);
    return Utilities.formatDate(d, CONFIG.COMPANY.TIMEZONE, 'yyyy-MM-dd');
}

function getDateAfterWeeks_(weeks) {
    var d = new Date();
    d.setDate(d.getDate() + (weeks * 7));
    return Utilities.formatDate(d, CONFIG.COMPANY.TIMEZONE, 'yyyy-MM-dd');
}

// ==================== دالة اختبار ====================

/**
 * اختبار الـ Smart Agent
 * شغّل هذه الدالة من محرر Apps Script للتجربة
 */
function testSmartAgent() {
    var testMessages = [
        'اتفقت مع أحمد نايل على رسم فيلم بل كلنتون بقيمة 400 دولار',
        'دفعت 500 دولار لأحمد مونتاج نقدي',
        'مصاريف بنكية 15 دولار على حوالة',
        'صرفت 1000 دولار بسعر 38.5',
        'استلمت 5000 دولار من قناة الجزيرة'
    ];

    testMessages.forEach(function(msg, i) {
        Logger.log('\n' + '═'.repeat(50));
        Logger.log('📝 اختبار ' + (i + 1) + ': ' + msg);
        Logger.log('═'.repeat(50));

        var result = smartAnalyze(msg);
        Logger.log('📋 النتيجة: ' + JSON.stringify(result, null, 2));
    });
}
