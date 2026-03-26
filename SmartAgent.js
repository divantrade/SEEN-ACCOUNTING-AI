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

var SMART_AGENT_CONFIG = {
    // الحد الأقصى لدورات التفكير (لمنع الحلقات اللانهائية)
    MAX_TURNS: 8,

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

    return `أنت محاسب ذكي في نظام SEEN لإنتاج الأفلام الوثائقية.
مهمتك: تحليل رسائل المستخدم العربية واستخراج الحركات المالية.

📅 اليوم: ${today} (${dayName})

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

## تعليمات العمل:
1. حلل رسالة المستخدم واستخرج البيانات المالية
2. استخدم أداة search_projects للتأكد من اسم المشروع
3. استخدم أداة search_parties للتأكد من اسم الطرف
4. استخدم أداة search_items للتأكد من البند
5. إذا كانت هناك معلومات ناقصة لا يمكن استنتاجها، استخدم ask_user (اجمع عدة أسئلة في سؤال واحد)
6. عندما تكتمل البيانات، استخدم show_confirmation لعرض الملخص
7. لا تحفظ إلا بعد تأكيد المستخدم عبر show_confirmation

## ملاحظات:
- الأرقام العربية (٠١٢٣٤٥٦٧٨٩) = أرقام إنجليزية (0123456789)
- "بكرة/غداً" = ${getNextDay_()}
- "أمس/إمبارح" = ${getYesterday_()}
- "بعد شهر" = تاريخ مخصص: ${getDateAfterMonths_(1)}
- "بعد أسبوع" = تاريخ مخصص: ${getDateAfterWeeks_(1)}
- العملة الافتراضية USD إذا لم تُذكر
- سعر الصرف الافتراضي: TRY=38.0, EGP=50.0`;
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

    // فحص الأجزاء - هل فيها function call أو نص؟
    for (var i = 0; i < parts.length; i++) {
        var part = parts[i];

        // 1. Function Call - الـ AI يريد استخدام أداة
        if (part.functionCall) {
            return handleFunctionCall_(part.functionCall, updatedHistory);
        }

        // 2. نص عادي - الـ AI يريد التواصل مع المستخدم مباشرة
        if (part.text) {
            return {
                action: 'SEND_MESSAGE',
                message: part.text,
                agentHistory: updatedHistory
            };
        }
    }

    return { action: 'ERROR', error: 'رد فارغ من AI' };
}

/**
 * معالجة استدعاء أداة من الـ AI
 * @param {Object} functionCall - بيانات الاستدعاء
 * @param {Array} history - سجل المحادثة
 * @returns {Object} - الإجراء المطلوب
 */
function handleFunctionCall_(functionCall, history) {
    var toolName = functionCall.name;
    var args = functionCall.args || {};

    Logger.log('🔧 Smart Agent يستدعي: ' + toolName);

    // تنفيذ الأداة
    var toolResult = executeAgentTool_(toolName, args);

    // أدوات خاصة تحتاج تفاعل المستخدم
    if (toolResult.action === 'WAIT_FOR_USER') {
        return {
            action: 'ASK_USER',
            question: toolResult.question,
            options: toolResult.options,
            field_name: toolResult.field_name,
            agentHistory: history,
            // إضافة نتيجة الأداة للسجل
            pendingToolResponse: {
                name: toolName,
                response: toolResult
            }
        };
    }

    if (toolResult.action === 'SHOW_CONFIRMATION') {
        return {
            action: 'SHOW_CONFIRMATION',
            transaction: toolResult.transaction,
            agentHistory: history
        };
    }

    // أدوات عادية (بحث، رصيد) - نعيد النتيجة لـ Gemini ونتابع
    return continueAfterToolExecution_(toolName, toolResult, history);
}

/**
 * متابعة المحادثة بعد تنفيذ أداة (إعادة النتيجة لـ Gemini)
 * @param {string} toolName - اسم الأداة
 * @param {Object} toolResult - نتيجة التنفيذ
 * @param {Array} history - سجل المحادثة
 * @returns {Object} - الإجراء التالي
 */
function continueAfterToolExecution_(toolName, toolResult, history) {
    // إضافة نتيجة الأداة لسجل المحادثة
    var updatedHistory = history.slice();
    updatedHistory.push({
        role: 'function',
        parts: [{
            functionResponse: {
                name: toolName,
                response: toolResult
            }
        }]
    });

    // حماية من الحلقات اللانهائية
    var turnCount = 0;
    for (var i = 0; i < updatedHistory.length; i++) {
        if (updatedHistory[i].role === 'model') turnCount++;
    }
    if (turnCount >= SMART_AGENT_CONFIG.MAX_TURNS) {
        Logger.log('⚠️ Smart Agent: وصل الحد الأقصى للدورات');
        return {
            action: 'ERROR',
            error: 'تم تجاوز الحد الأقصى لمحاولات التحليل',
            agentHistory: updatedHistory
        };
    }

    // استدعاء Gemini مرة أخرى مع النتيجة
    var nextResult = callGeminiWithTools_(updatedHistory);

    if (!nextResult.success) {
        return { action: 'ERROR', error: nextResult.error, agentHistory: updatedHistory };
    }

    return processAgentResponse_(nextResult, updatedHistory);
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
