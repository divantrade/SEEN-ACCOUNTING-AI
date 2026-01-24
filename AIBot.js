// ==================== البوت الذكي لنظام SEEN ====================
/**
 * البوت الذكي الرئيسي - يفهم اللغة الطبيعية ويحولها لحركات مالية
 * يعمل بالتوازي مع البوت التقليدي دون التأثير عليه
 */

// ==================== تخزين جلسات المستخدمين ====================
const aiUserSessions = {};

// ==================== معالجة التحديثات ====================

/**
 * معالجة تحديثات البوت الذكي (Long Polling)
 * يتم استدعاؤها بواسطة Time-driven Trigger كل دقيقة
 */
function processAIBotUpdates() {
    // استخدام Lock لمنع التنفيذ المتزامن
    const lock = LockService.getScriptLock();
    const hasLock = lock.tryLock(1000);

    if (!hasLock) {
        Logger.log('⏭️ AI Bot: Instance أخرى تعمل - تخطي');
        return;
    }

    try {
        // التحقق من إعداد البوت
        const setup = checkAIBotSetup();
        if (!setup.ready) {
            Logger.log('البوت الذكي غير جاهز - يرجى إعداد المفاتيح أولاً');
            return;
        }

        const token = getAIBotToken();
        const startTime = Date.now();
        const MAX_TIME = 55000; // 55 ثانية

        Logger.log('🤖 البوت الذكي يعمل...');

        // حلقة polling لمدة 55 ثانية
        while (Date.now() - startTime < MAX_TIME) {
            const lastUpdateId = getAILastUpdateId();

            // جلب التحديثات مع timeout قصير
            const url = `https://api.telegram.org/bot${token}/getUpdates?offset=${lastUpdateId + 1}&timeout=5`;

            const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
            const data = JSON.parse(response.getContentText());

            if (!data.ok) {
                Logger.log('AI Bot Error: ' + JSON.stringify(data));
                Utilities.sleep(1000);
                continue;
            }

            const updates = data.result;

            if (updates.length > 0) {
                Logger.log('📥 استلام ' + updates.length + ' تحديث');

                // معالجة كل تحديث
                updates.forEach(update => {
                    try {
                        if (update.message) {
                            handleAIMessage(update.message);
                        } else if (update.callback_query) {
                            handleAICallback(update.callback_query);
                        }
                    } catch (error) {
                        Logger.log('Update Processing Error: ' + error.message);
                    }
                });

                // حفظ آخر update_id
                const lastId = updates[updates.length - 1].update_id;
                setAILastUpdateId(lastId);
            }
        }

        Logger.log('⏹️ انتهى وقت البوت');

    } catch (error) {
        Logger.log('AI Bot Main Error: ' + error.message);
    } finally {
        lock.releaseLock();
    }
}

/**
 * الحصول على آخر update_id
 */
function getAILastUpdateId() {
    const id = PropertiesService.getScriptProperties().getProperty('AI_BOT_LAST_UPDATE_ID');
    return id ? parseInt(id) : 0;
}

/**
 * حفظ آخر update_id
 */
function setAILastUpdateId(id) {
    PropertiesService.getScriptProperties().setProperty('AI_BOT_LAST_UPDATE_ID', id.toString());
}

/**
 * فحص وإصلاح آخر update_id - شغّل هذه الدالة إذا البوت لا يستجيب
 */
function fixLastUpdateId() {
    const token = PropertiesService.getScriptProperties().getProperty('AI_BOT_TOKEN');
    const currentId = PropertiesService.getScriptProperties().getProperty('AI_BOT_LAST_UPDATE_ID');

    Logger.log('📋 آخر Update ID المحفوظ: ' + (currentId || 'غير موجود'));

    // جلب آخر update من تليجرام
    const url = `https://api.telegram.org/bot${token}/getUpdates?limit=1&offset=-1`;
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const data = JSON.parse(response.getContentText());

    Logger.log('📥 استجابة تليجرام: ' + JSON.stringify(data));

    if (data.ok && data.result && data.result.length > 0) {
        const latestUpdate = data.result[0];
        Logger.log('📨 آخر رسالة في تليجرام:');
        Logger.log('   Update ID: ' + latestUpdate.update_id);
        if (latestUpdate.message) {
            Logger.log('   النص: ' + (latestUpdate.message.text || '[بدون نص]'));
            Logger.log('   من: ' + (latestUpdate.message.from?.first_name || 'مجهول'));
        }

        // إعادة تعيين الـ offset ليبدأ من آخر رسالة
        const newOffset = latestUpdate.update_id;
        PropertiesService.getScriptProperties().setProperty('AI_BOT_LAST_UPDATE_ID', newOffset.toString());
        Logger.log('✅ تم تحديث Last Update ID إلى: ' + newOffset);
        Logger.log('🔄 الرسائل الجديدة القادمة ستتم معالجتها');
    } else {
        Logger.log('ℹ️ لا توجد رسائل في انتظار المعالجة');
        // إعادة تعيين إلى 0 لبدء من جديد
        PropertiesService.getScriptProperties().setProperty('AI_BOT_LAST_UPDATE_ID', '0');
        Logger.log('✅ تم إعادة تعيين Last Update ID إلى 0');
    }
}

/**
 * فحص معلومات البوت - للتحقق من هوية البوت
 * قم بتشغيل هذه الدالة يدوياً للتحقق من أنك تراسل البوت الصحيح
 */
function checkBotInfo() {
    const token = PropertiesService.getScriptProperties().getProperty('AI_BOT_TOKEN');
    if (!token) {
        Logger.log('❌ لم يتم تعيين AI_BOT_TOKEN');
        return;
    }

    const url = `https://api.telegram.org/bot${token}/getMe`;
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const data = JSON.parse(response.getContentText());

    if (data.ok) {
        Logger.log('✅ معلومات البوت:');
        Logger.log('📛 الاسم: ' + data.result.first_name);
        Logger.log('🔗 Username: @' + data.result.username);
        Logger.log('🆔 Bot ID: ' + data.result.id);
        Logger.log('🤖 Is Bot: ' + data.result.is_bot);
    } else {
        Logger.log('❌ خطأ في الاتصال بالبوت: ' + JSON.stringify(data));
    }

    return data;
}


// ==================== معالجة الرسائل ====================

/**
 * معالجة الرسائل النصية
 */
function handleAIMessage(message) {
    const chatId = message.chat.id;
    const text = message.text;
    const user = message.from;

    Logger.log('═══════════════════════════════════════');
    Logger.log('📨 NEW MESSAGE RECEIVED');
    Logger.log('📨 chatId: ' + chatId);
    Logger.log('📨 text: "' + text + '"');
    Logger.log('═══════════════════════════════════════');

    // التحقق من الصلاحيات
    const permission = checkAIUserPermission(chatId, user);
    if (!permission.authorized) {
        Logger.log('❌ User not authorized');
        sendAIMessage(chatId, AI_CONFIG.AI_MESSAGES.UNAUTHORIZED);
        return;
    }

    // معالجة الأوامر
    if (text && text.startsWith('/')) {
        Logger.log('📍 Processing as command');
        handleAICommand(chatId, text, user);
        return;
    }

    // الحصول على جلسة المستخدم
    const session = getAIUserSession(chatId);

    // ⭐ تسجيل حالة الجلسة للتصحيح - مفصّل
    Logger.log('═══════════════════════════════════════');
    Logger.log('📍 SESSION DEBUG INFO');
    Logger.log('📍 Session state: "' + (session.state || 'undefined/null') + '"');
    Logger.log('📍 Expected WAITING_EXCHANGE_RATE: "' + AI_CONFIG.AI_CONVERSATION_STATES.WAITING_EXCHANGE_RATE + '"');
    Logger.log('📍 States match: ' + (session.state === AI_CONFIG.AI_CONVERSATION_STATES.WAITING_EXCHANGE_RATE));
    Logger.log('📍 Has transaction: ' + (session.transaction ? 'yes' : 'no'));
    Logger.log('📍 Has validation: ' + (session.validation ? 'yes' : 'no'));
    Logger.log('═══════════════════════════════════════');

    // معالجة حسب حالة المحادثة
    switch (session.state) {
        case AI_CONFIG.AI_CONVERSATION_STATES.WAITING_MISSING_FIELD:
            handleMissingFieldInput(chatId, text, session);
            break;

        case AI_CONFIG.AI_CONVERSATION_STATES.WAITING_PROJECT_SELECTION:
            handleAIProjectSelection(chatId, text, session);
            break;

        case AI_CONFIG.AI_CONVERSATION_STATES.WAITING_PARTY_SELECTION:
            handleAIPartySelection(chatId, text, session);
            break;

        case AI_CONFIG.AI_CONVERSATION_STATES.WAITING_EDIT:
            handleEditInput(chatId, text, session);
            break;

        case AI_CONFIG.AI_CONVERSATION_STATES.WAITING_EXCHANGE_RATE:
            // ⭐ معالجة إدخال سعر الصرف
            handleAIExchangeRateInput(chatId, text, session);
            break;

        case AI_CONFIG.AI_CONVERSATION_STATES.WAITING_PAYMENT_TERM:
            // ⭐ معالجة إدخال شرط الدفع (أسابيع أو تاريخ)
            handlePaymentTermInput(chatId, text, session);
            break;

        case AI_CONFIG.AI_CONVERSATION_STATES.WAITING_LOAN_DUE_DATE:
            // ⭐ معالجة إدخال تاريخ استحقاق السلفة/التمويل
            handleLoanDueDateInput(chatId, text, session);
            break;

        // ⭐ حالات انتظار ضغط الأزرار - إذا أرسل المستخدم نص، نطلب منه الضغط على الزر
        case AI_CONFIG.AI_CONVERSATION_STATES.WAITING_PAYMENT_METHOD:
            Logger.log('⚠️ User sent text while waiting for payment method button');
            sendAIMessage(chatId, '⚠️ يرجى اختيار طريقة الدفع من الأزرار أعلاه (تحويل بنكي / نقدي)', { parse_mode: 'Markdown' });
            // إعادة إرسال الأزرار
            askPaymentMethod(chatId, session);
            break;

        case AI_CONFIG.AI_CONVERSATION_STATES.WAITING_CURRENCY:
            Logger.log('⚠️ User sent text while waiting for currency button');
            sendAIMessage(chatId, '⚠️ يرجى اختيار العملة من الأزرار أعلاه', { parse_mode: 'Markdown' });
            askCurrency(chatId, session);
            break;

        case AI_CONFIG.AI_CONVERSATION_STATES.WAITING_NEW_PARTY_CONFIRM:
            Logger.log('⚠️ User sent text while waiting for party confirmation');
            sendAIMessage(chatId, '⚠️ يرجى اختيار أحد الخيارات من الأزرار أعلاه', { parse_mode: 'Markdown' });
            askNewPartyConfirmation(chatId, session);
            break;

        case AI_CONFIG.AI_CONVERSATION_STATES.WAITING_CONFIRMATION:
        case AI_CONFIG.AI_CONVERSATION_STATES.CONFIRM_WAIT:
            Logger.log('⚠️ User sent text while waiting for confirmation');
            sendAIMessage(chatId, '⚠️ يرجى تأكيد الحركة أو تعديلها من الأزرار أعلاه', { parse_mode: 'Markdown' });
            showTransactionConfirmation(chatId, session);
            break;

        default:
            // ⭐ التحقق من وضع التقارير أولاً
            if (isInReportMode(session)) {
                if (session.reportState === REPORTS_CONFIG.STATES.WAITING_PARTY_NAME) {
                    handleReportPartySearch(chatId, text, session);
                    return;
                }
            }
            // تحليل النص كحركة مالية جديدة
            Logger.log('⚠️ DEFAULT CASE - Processing as new transaction');
            Logger.log('⚠️ Session state was: "' + session.state + '"');
            processNewTransaction(chatId, text, user);
    }
}

/**
 * معالجة الأوامر
 */
function handleAICommand(chatId, command, user) {
    const cmd = command.split(' ')[0].toLowerCase();

    switch (cmd) {
        case '/start':
            sendAIMessage(chatId, AI_CONFIG.AI_MESSAGES.WELCOME, { parse_mode: 'Markdown' });
            resetAIUserSession(chatId);
            break;

        case '/help':
        case '/مساعدة':
            sendAIMessage(chatId, AI_CONFIG.AI_MESSAGES.HELP, { parse_mode: 'Markdown' });
            break;

        case '/cancel':
        case '/الغاء':
            sendAIMessage(chatId, AI_CONFIG.AI_MESSAGES.CANCELLED);
            resetAIUserSession(chatId);
            break;

        case '/status':
        case '/حالة':
            showUserTransactionStatus(chatId, user);
            break;

        case '/reports':
        case '/تقارير':
        case '/report':
        case '/تقرير':
            // ⭐ أمر التقارير وكشوف الحساب
            handleReportsCommand(chatId);
            break;

        case '/statement':
        case '/كشف':
            // ⭐ كشف حساب مباشر
            handleReportsCommand(chatId);
            break;

        default:
            // إذا كان الأمر غير معروف، حاول تحليله كحركة
            processNewTransaction(chatId, command, user);
    }
}


// ==================== معالجة الحركات الجديدة ====================

/**
 * تحليل ومعالجة حركة جديدة
 */
function processNewTransaction(chatId, text, user) {
    Logger.log('═══════════════════════════════════════');
    Logger.log('📊 processNewTransaction STARTED');
    Logger.log('📊 chatId: ' + chatId);
    Logger.log('📊 text: ' + text);

    // إرسال رسالة "جاري التحليل"
    const loadingMsg = sendAIMessage(chatId, AI_CONFIG.AI_MESSAGES.ANALYZING, { parse_mode: 'Markdown' });

    try {
        // تحليل النص
        Logger.log('📊 Calling analyzeTransaction...');
        const result = analyzeTransaction(text);
        Logger.log('📊 analyzeTransaction returned');
        Logger.log('📊 result.success: ' + result.success);
        Logger.log('📊 result.needsInput: ' + result.needsInput);
        Logger.log('📊 result.missingFields: ' + JSON.stringify(result.missingFields));

        if (!result.success) {
            Logger.log('❌ result.success is false, sending error');
            sendAIMessage(chatId, result.error || AI_CONFIG.AI_MESSAGES.ERROR_PARSE, { parse_mode: 'Markdown' });
            return;
        }

        // ⭐ التحقق من الأوردر المشترك
        if (result.transaction && result.transaction.is_shared_order) {
            Logger.log('✅ Detected SHARED ORDER');
            handleSharedOrder(chatId, result.transaction, user);
            return;
        }

        // حفظ الحركة في الجلسة
        Logger.log('📊 Getting session...');
        const session = getAIUserSession(chatId);
        session.transaction = result.transaction;
        session.validation = result.validation;
        session.originalText = text;

        // حفظ التغييرات في الكاش
        saveAIUserSession(chatId, session);
        Logger.log('📊 Session saved');

        // التحقق من الحقول الناقصة
        Logger.log('📊 Checking missing fields...');
        Logger.log('📊 result.needsInput: ' + result.needsInput);
        Logger.log('📊 result.missingFields?.length: ' + (result.missingFields ? result.missingFields.length : 'undefined'));

        if (result.needsInput && result.missingFields.length > 0) {
            Logger.log('✅ Has missing fields, calling handleMissingFields');
            handleMissingFields(chatId, result.missingFields, session);
            return;
        }

        // التحقق من طرف جديد يحتاج تأكيد
        Logger.log('📊 Checking needsPartyConfirmation: ' + (result.validation ? result.validation.needsPartyConfirmation : 'no validation'));
        if (result.validation && result.validation.needsPartyConfirmation) {
            Logger.log('✅ Needs party confirmation');
            askNewPartyConfirmation(chatId, session);
            return;
        }

        // ⭐ التحقق من المشروع (اختياري - يمكن التخطي)
        Logger.log('📊 Checking needsProjectSelection: ' + (result.validation ? result.validation.needsProjectSelection : 'no validation'));
        if (result.validation && result.validation.needsProjectSelection) {
            Logger.log('✅ Needs project selection (optional)');
            askProjectSelection(chatId, session);
            return;
        }

        // ⭐ التحقق من طريقة الدفع
        Logger.log('📊 Checking needsPaymentMethod: ' + (result.validation ? result.validation.needsPaymentMethod : 'no validation'));
        if (result.validation && result.validation.needsPaymentMethod) {
            Logger.log('✅ Needs payment method');
            askPaymentMethod(chatId, session);
            return;
        }

        // ⭐ التحقق من العملة
        Logger.log('📊 Checking needsCurrency: ' + (result.validation ? result.validation.needsCurrency : 'no validation'));
        if (result.validation && result.validation.needsCurrency) {
            Logger.log('✅ Needs currency');
            askCurrency(chatId, session);
            return;
        }

        // ⭐ التحقق من سعر الصرف (إذا العملة غير دولار)
        Logger.log('📊 Checking needsExchangeRate: ' + (result.validation ? result.validation.needsExchangeRate : 'no validation'));
        if (result.validation && result.validation.needsExchangeRate) {
            Logger.log('✅ Needs exchange rate');
            askExchangeRate(chatId, session);
            return;
        }

        // ⭐ التحقق من شرط الدفع (للاستحقاقات فقط)
        Logger.log('📊 Checking needsPaymentTerm: ' + (result.validation ? result.validation.needsPaymentTerm : 'no validation'));
        if (result.validation && result.validation.needsPaymentTerm) {
            Logger.log('✅ Needs payment term');
            askPaymentTerm(chatId, session);
            return;
        }

        // ⭐ التحقق من عدد الأسابيع (لشرط بعد التسليم)
        Logger.log('📊 Checking needsPaymentTermWeeks: ' + (result.validation ? result.validation.needsPaymentTermWeeks : 'no validation'));
        if (result.validation && result.validation.needsPaymentTermWeeks) {
            Logger.log('✅ Needs payment term weeks');
            askPaymentTermWeeks(chatId, session);
            return;
        }

        // ⭐ التحقق من تاريخ الدفع المخصص
        Logger.log('📊 Checking needsPaymentTermDate: ' + (result.validation ? result.validation.needsPaymentTermDate : 'no validation'));
        if (result.validation && result.validation.needsPaymentTermDate) {
            Logger.log('✅ Needs payment term date');
            askPaymentTermDate(chatId, session);
            return;
        }

        // ⭐ التحقق من تاريخ استحقاق السلفة/التمويل
        Logger.log('📊 Checking needsLoanDueDate: ' + (result.validation ? result.validation.needsLoanDueDate : 'no validation'));
        if (result.validation && result.validation.needsLoanDueDate) {
            Logger.log('✅ Needs loan due date');
            askLoanDueDate(chatId, session);
            return;
        }

        // عرض ملخص للتأكيد
        Logger.log('📊 All checks passed, showing confirmation');
        showTransactionConfirmation(chatId, session);

    } catch (error) {
        Logger.log('❌ Process Transaction Error: ' + error.message);
        Logger.log('Stack: ' + error.stack);
        sendAIMessage(chatId, `❌ *حدث خطأ غير متوقع:*\n${error.message}\n\nيرجى إعادة المحاولة أو التواصل مع الدعم التقني.`);
    }

    Logger.log('📊 processNewTransaction ENDED');
    Logger.log('═══════════════════════════════════════');
}

/**
 * معالجة الحقول الناقصة
 */
function handleMissingFields(chatId, missingFields, session) {
    session.missingFields = missingFields;
    session.currentMissingIndex = 0;
    session.state = AI_CONFIG.AI_CONVERSATION_STATES.WAITING_MISSING_FIELD;
    saveAIUserSession(chatId, session);

    askForMissingField(chatId, session);
}

/**
 * السؤال عن حقل ناقص
 */
function askForMissingField(chatId, session) {
    const field = session.missingFields[session.currentMissingIndex];

    let message = '';
    let keyboard = null;

    switch (field.field) {
        case 'project':
            message = AI_CONFIG.AI_MESSAGES.ASK_PROJECT;
            session.state = AI_CONFIG.AI_CONVERSATION_STATES.WAITING_PROJECT_SELECTION;
            // بناء لوحة المشاريع
            keyboard = buildProjectsKeyboard();
            saveAIUserSession(chatId, session);
            break;

        case 'party':
            message = AI_CONFIG.AI_MESSAGES.ASK_PARTY;
            session.state = AI_CONFIG.AI_CONVERSATION_STATES.WAITING_PARTY_SELECTION;
            saveAIUserSession(chatId, session);
            break;

        case 'amount':
            message = AI_CONFIG.AI_MESSAGES.ASK_AMOUNT;
            break;

        default:
            message = field.message;
    }

    sendAIMessage(chatId, message, {
        parse_mode: 'Markdown',
        reply_markup: keyboard ? JSON.stringify(keyboard) : null
    });
}

/**
 * معالجة إدخال حقل ناقص
 */
function handleMissingFieldInput(chatId, text, session) {
    const field = session.missingFields[session.currentMissingIndex];

    // تحديث الحركة
    session.transaction[field.field] = text;

    // الانتقال للحقل التالي
    session.currentMissingIndex++;

    saveAIUserSession(chatId, session);

    if (session.currentMissingIndex < session.missingFields.length) {
        askForMissingField(chatId, session);
    } else {
        // اكتملت الحقول - عرض التأكيد
        session.state = AI_CONFIG.AI_CONVERSATION_STATES.CONFIRM_WAIT;
        showTransactionConfirmation(chatId, session);
    }
}

/**
 * طلب تأكيد إضافة طرف جديد أو اختيار من المتشابهين
 */
function askNewPartyConfirmation(chatId, session) {
    const partyName = session.validation.enriched.newPartyName || session.transaction.party;
    const suggestions = session.validation.warnings?.find(w => w.field === 'party')?.suggestions || [];

    // تحديد نوع الطرف بناءً على نوع الحركة
    let partyType = 'مورد';
    const nature = session.transaction.nature || '';
    if (nature.includes('إيراد')) {
        partyType = 'عميل';
    } else if (nature.includes('تمويل')) {
        partyType = 'ممول';
    }

    session.newPartyName = partyName;
    session.newPartyType = partyType;
    session.partySuggestions = suggestions; // حفظ الاقتراحات
    session.state = AI_CONFIG.AI_CONVERSATION_STATES.WAITING_NEW_PARTY_CONFIRM;
    saveAIUserSession(chatId, session);

    // إذا وجدت أطراف متشابهة، اعرضها للاختيار
    if (suggestions && suggestions.length > 0) {
        let message = `🔍 *وجدت أطراف متشابهة لـ "${partyName}"*\n\n`;
        message += `اختر أحد الأطراف التالية أو أضف طرف جديد:\n\n`;

        const keyboard = { inline_keyboard: [] };

        // إضافة زر لكل طرف مقترح (أقصى 5)
        suggestions.slice(0, 5).forEach((s, index) => {
            const name = s.name || s;
            const type = s.type || '';
            keyboard.inline_keyboard.push([
                { text: `👤 ${name}${type ? ' (' + type + ')' : ''}`, callback_data: `ai_select_party_${index}` }
            ]);
        });

        // إضافة خيارات إضافية
        keyboard.inline_keyboard.push([
            { text: '➕ إضافة كطرف جديد', callback_data: 'ai_add_party_yes' }
        ]);
        keyboard.inline_keyboard.push([
            { text: '✏️ تعديل الاسم', callback_data: 'ai_add_party_edit' },
            { text: '❌ إلغاء', callback_data: 'ai_add_party_no' }
        ]);

        sendAIMessage(chatId, message, {
            parse_mode: 'Markdown',
            reply_markup: JSON.stringify(keyboard)
        });
    } else {
        // لا توجد اقتراحات - اعرض خيار الإضافة فقط
        const message = `⚠️ *الطرف غير موجود في قاعدة البيانات*

👤 الاسم: *${partyName}*
📋 النوع المقترح: ${partyType}

هل تريد إضافة هذا الطرف الجديد؟`;

        const keyboard = {
            inline_keyboard: [
                [
                    { text: '✅ نعم، أضف الطرف', callback_data: 'ai_add_party_yes' },
                    { text: '❌ لا، إلغاء', callback_data: 'ai_add_party_no' }
                ],
                [
                    { text: '✏️ تعديل الاسم', callback_data: 'ai_add_party_edit' },
                    { text: '🔄 تغيير النوع', callback_data: 'ai_add_party_type' }
                ]
            ]
        };

        sendAIMessage(chatId, message, {
            parse_mode: 'Markdown',
            reply_markup: JSON.stringify(keyboard)
        });
    }
}

/**
 * معالجة تأكيد إضافة طرف جديد
 */
function handleNewPartyConfirmation(chatId, action, session) {
    switch (action) {
        case 'ai_add_party_yes':
            // إضافة الطرف لقاعدة البيانات
            const added = addNewParty(session.newPartyName, session.newPartyType);
            if (added) {
                sendAIMessage(chatId, `✅ تم إضافة الطرف *${session.newPartyName}* بنجاح!`, { parse_mode: 'Markdown' });

                // تحديث الجلسة والمتابعة
                session.transaction.party = session.newPartyName;
                session.validation.enriched.party = session.newPartyName;
                session.validation.enriched.partyType = session.newPartyType;
                session.validation.needsPartyConfirmation = false;
                session.transaction.isNewParty = true;
                saveAIUserSession(chatId, session);

                // ⭐ التحقق من باقي الحقول المطلوبة
                continueValidation(chatId, session);
            } else {
                sendAIMessage(chatId, '❌ فشل في إضافة الطرف. يرجى المحاولة مرة أخرى.');
            }
            break;

        case 'ai_add_party_no':
            sendAIMessage(chatId, AI_CONFIG.AI_MESSAGES.CANCELLED);
            resetAIUserSession(chatId);
            break;

        case 'ai_add_party_edit':
            session.state = AI_CONFIG.AI_CONVERSATION_STATES.WAITING_PARTY_SELECTION;
            saveAIUserSession(chatId, session);
            sendAIMessage(chatId, '✏️ أدخل اسم الطرف الصحيح:');
            break;

        case 'ai_add_party_type':
            showPartyTypeSelection(chatId, session);
            break;
    }
}

/**
 * عرض اختيار نوع الطرف
 */
function showPartyTypeSelection(chatId, session) {
    const keyboard = {
        inline_keyboard: [
            [
                { text: '🏭 مورد', callback_data: 'ai_party_type_مورد' },
                { text: '👥 عميل', callback_data: 'ai_party_type_عميل' }
            ],
            [
                { text: '💰 ممول', callback_data: 'ai_party_type_ممول' }
            ]
        ]
    };

    sendAIMessage(chatId, `اختر نوع الطرف *${session.newPartyName}*:`, {
        parse_mode: 'Markdown',
        reply_markup: JSON.stringify(keyboard)
    });
}

/**
 * ⭐ معالجة اختيار طرف من قائمة الاقتراحات
 */
function handleSelectPartyFromSuggestions(chatId, index, session) {
    const suggestions = session.partySuggestions || [];

    if (index >= 0 && index < suggestions.length) {
        const selectedParty = suggestions[index];
        const partyName = selectedParty.name || selectedParty;
        const partyType = selectedParty.type || 'مورد';

        // تحديث الجلسة بالطرف المختار
        session.transaction.party = partyName;
        session.validation.enriched.party = partyName;
        session.validation.enriched.partyType = partyType;
        session.validation.needsPartyConfirmation = false;
        session.validation.enriched.isNewParty = false;
        saveAIUserSession(chatId, session);

        sendAIMessage(chatId, `✅ تم اختيار الطرف: *${partyName}* (${partyType})`, { parse_mode: 'Markdown' });

        // ⭐ التحقق من باقي الحقول المطلوبة (طريقة الدفع، العملة، سعر الصرف، إلخ)
        continueValidation(chatId, session);
    } else {
        sendAIMessage(chatId, '❌ حدث خطأ في اختيار الطرف. يرجى المحاولة مرة أخرى.');
    }
}

/**
 * ⭐ السؤال عن المشروع (اختياري - يمكن التخطي)
 */
function askProjectSelection(chatId, session) {
    session.state = AI_CONFIG.AI_CONVERSATION_STATES.WAITING_PROJECT_SELECTION;
    saveAIUserSession(chatId, session);

    const keyboard = buildProjectsKeyboard(true); // true = include skip option

    sendAIMessage(chatId, '🎬 *اختر المشروع:*\n\n_(يمكنك التخطي إذا لم تكن الحركة مرتبطة بمشروع محدد)_', {
        parse_mode: 'Markdown',
        reply_markup: JSON.stringify(keyboard)
    });
}

/**
 * ⭐ السؤال عن طريقة الدفع
 */
function askPaymentMethod(chatId, session) {
    session.state = AI_CONFIG.AI_CONVERSATION_STATES.WAITING_PAYMENT_METHOD;
    saveAIUserSession(chatId, session);

    sendAIMessage(chatId, AI_CONFIG.AI_MESSAGES.ASK_PAYMENT_METHOD, {
        parse_mode: 'Markdown',
        reply_markup: JSON.stringify(AI_CONFIG.AI_KEYBOARDS.PAYMENT_METHOD)
    });
}

/**
 * ⭐ معالجة اختيار طريقة الدفع
 */
function handleAIPaymentMethodSelection(chatId, method, session) {
    Logger.log('🔧 handleAIPaymentMethodSelection called with method: ' + method);

    // ⭐ الفحوصات الرئيسية موجودة في handleAICallback، هذه فقط للأمان
    if (!session || !session.transaction || !session.validation) {
        Logger.log('❌ Session data incomplete in handleAIPaymentMethodSelection');
        sendAIMessage(chatId, '⚠️ حدث خطأ. يرجى إعادة إرسال الحركة.');
        resetAIUserSession(chatId);
        return;
    }

    // ⭐ إنشاء enriched إذا لم يكن موجوداً
    if (!session.validation.enriched) {
        session.validation.enriched = {};
    }

    // ⭐ حفظ طريقة الدفع (القيم الآن تأتي مطابقة للشيت مباشرة)
    session.transaction.payment_method = method;
    session.validation.enriched.payment_method = method;
    session.validation.needsPaymentMethod = false;
    saveAIUserSession(chatId, session);

    const emoji = method === 'نقدي' ? '💵' : '🏦';
    sendAIMessage(chatId, `✅ تم تحديد طريقة الدفع: *${emoji} ${method}*`, { parse_mode: 'Markdown' });

    Logger.log('✅ Payment method saved: ' + method + ', calling continueValidation');
    // التحقق من الخطوات التالية
    continueValidation(chatId, session);
}

/**
 * ⭐ السؤال عن العملة
 */
function askCurrency(chatId, session) {
    session.state = AI_CONFIG.AI_CONVERSATION_STATES.WAITING_CURRENCY;
    saveAIUserSession(chatId, session);

    sendAIMessage(chatId, AI_CONFIG.AI_MESSAGES.ASK_CURRENCY, {
        parse_mode: 'Markdown',
        reply_markup: JSON.stringify(AI_CONFIG.AI_KEYBOARDS.CURRENCY)
    });
}

/**
 * ⭐ معالجة اختيار العملة
 */
function handleAICurrencySelection(chatId, currency, session) {
    session.transaction.currency = currency;
    session.validation.enriched.currency = currency;
    session.validation.needsCurrency = false;
    saveAIUserSession(chatId, session);

    const currencyNames = { 'USD': 'دولار', 'TRY': 'ليرة', 'EGP': 'جنيه' };
    sendAIMessage(chatId, `✅ تم تحديد العملة: *${currencyNames[currency] || currency}*`, { parse_mode: 'Markdown' });

    // إذا كانت العملة غير دولار، نسأل عن سعر الصرف
    if (currency !== 'USD') {
        session.validation.needsExchangeRate = true;
        saveAIUserSession(chatId, session);
        askExchangeRate(chatId, session);
    } else {
        // التحقق من الخطوات التالية
        continueValidation(chatId, session);
    }
}

/**
 * ⭐ السؤال عن سعر الصرف
 */
function askExchangeRate(chatId, session) {
    Logger.log('═══════════════════════════════════════');
    Logger.log('📤 askExchangeRate CALLED');
    Logger.log('📤 Setting state to: ' + AI_CONFIG.AI_CONVERSATION_STATES.WAITING_EXCHANGE_RATE);

    session.state = AI_CONFIG.AI_CONVERSATION_STATES.WAITING_EXCHANGE_RATE;
    saveAIUserSession(chatId, session);

    Logger.log('📤 State saved. Verifying...');
    const verifySession = getAIUserSession(chatId);
    Logger.log('📤 Verified state: ' + verifySession.state);
    Logger.log('═══════════════════════════════════════');

    const currency = session.transaction.currency || session.validation.enriched.currency;
    const currencyNames = { 'TRY': 'الليرة التركية', 'EGP': 'الجنيه المصري' };
    const currencyName = currencyNames[currency] || currency;

    const message = AI_CONFIG.AI_MESSAGES.ASK_EXCHANGE_RATE.replace('{currency}', currencyName);
    sendAIMessage(chatId, message, { parse_mode: 'Markdown' });
}

/**
 * ⭐ معالجة إدخال سعر الصرف للبوت الذكي
 */
function handleAIExchangeRateInput(chatId, text, session) {
    Logger.log('═══════════════════════════════════════');
    Logger.log('📥 handleAIExchangeRateInput CALLED');
    Logger.log('📥 Input text: "' + text + '"');
    Logger.log('📥 Session exists: ' + (session ? 'yes' : 'no'));
    Logger.log('📥 Session.transaction: ' + (session && session.transaction ? 'exists' : 'UNDEFINED'));
    Logger.log('📥 Session.validation: ' + (session && session.validation ? 'exists' : 'UNDEFINED'));

    // ⭐ التحقق من وجود البيانات أولاً - قبل أي شيء آخر
    if (!session || !session.transaction || !session.validation) {
        Logger.log('❌ SESSION DATA MISSING!');
        Logger.log('❌ session: ' + JSON.stringify(session ? Object.keys(session) : 'null'));
        sendAIMessage(chatId, '⚠️ حدث خطأ في الجلسة. يرجى إعادة إرسال الحركة من البداية.');
        resetAIUserSession(chatId);
        return;
    }

    if (!session.validation.enriched) {
        session.validation.enriched = {};
    }

    Logger.log('✅ Session data validated');

    // ⭐ تحويل الأرقام العربية للإنجليزية
    const arabicNumerals = '٠١٢٣٤٥٦٧٨٩';
    const englishNumerals = '0123456789';
    let convertedText = text;
    for (let i = 0; i < arabicNumerals.length; i++) {
        convertedText = convertedText.replace(new RegExp(arabicNumerals[i], 'g'), englishNumerals[i]);
    }
    // تحويل الفاصلة العربية للنقطة
    convertedText = convertedText.replace(/٫/g, '.');
    convertedText = convertedText.replace(/،/g, '.');

    Logger.log('📥 Converted text: "' + convertedText + '"');

    // استخراج الرقم من النص
    const rate = parseFloat(convertedText.replace(/[^0-9.]/g, ''));
    Logger.log('📥 Parsed rate: ' + rate);

    if (isNaN(rate) || rate <= 0) {
        Logger.log('❌ Invalid rate: ' + rate);
        sendAIMessage(chatId, '❌ سعر الصرف غير صحيح. يرجى إدخال رقم صحيح (مثال: 32.5 أو ٣٢٫٥):');
        return;
    }

    Logger.log('✅ Rate is valid: ' + rate);

    session.transaction.exchangeRate = rate;
    session.transaction.exchange_rate = rate;
    session.validation.enriched.exchangeRate = rate;
    session.validation.needsExchangeRate = false;
    saveAIUserSession(chatId, session);

    sendAIMessage(chatId, `✅ تم تحديد سعر الصرف: *${rate}*`, { parse_mode: 'Markdown' });

    Logger.log('✅ Exchange rate saved: ' + rate);
    Logger.log('═══════════════════════════════════════');

    // التحقق من الخطوات التالية
    continueValidation(chatId, session);
}

/**
 * ⭐ السؤال عن شرط الدفع (للاستحقاقات فقط)
 */
function askPaymentTerm(chatId, session) {
    session.state = AI_CONFIG.AI_CONVERSATION_STATES.WAITING_PAYMENT_TERM;
    saveAIUserSession(chatId, session);

    sendAIMessage(chatId, AI_CONFIG.AI_MESSAGES.ASK_PAYMENT_TERM, {
        parse_mode: 'Markdown',
        reply_markup: JSON.stringify(AI_CONFIG.AI_KEYBOARDS.PAYMENT_TERM)
    });
}

/**
 * ⭐ معالجة اختيار شرط الدفع
 */
function handleAIPaymentTermSelection(chatId, term, session) {
    session.transaction.payment_term = term;
    session.validation.enriched.payment_term = term;
    session.validation.needsPaymentTerm = false;
    saveAIUserSession(chatId, session);

    const termLabels = { 'فوري': '⚡ فوري', 'بعد التسليم': '📦 بعد التسليم', 'تاريخ مخصص': '📅 تاريخ مخصص' };
    sendAIMessage(chatId, `✅ تم تحديد شرط الدفع: *${termLabels[term] || term}*`, { parse_mode: 'Markdown' });

    // إذا كان "بعد التسليم"، نسأل عن عدد الأسابيع
    if (term === 'بعد التسليم') {
        session.validation.needsPaymentTermWeeks = true;
        saveAIUserSession(chatId, session);
        askPaymentTermWeeks(chatId, session);
    } else if (term === 'تاريخ مخصص') {
        // إذا كان "تاريخ مخصص"، نسأل عن التاريخ
        session.validation.needsPaymentTermDate = true;
        saveAIUserSession(chatId, session);
        askPaymentTermDate(chatId, session);
    } else {
        // فوري - نكمل
        continueValidation(chatId, session);
    }
}

/**
 * ⭐ السؤال عن عدد الأسابيع بعد التسليم
 */
function askPaymentTermWeeks(chatId, session) {
    session.state = AI_CONFIG.AI_CONVERSATION_STATES.WAITING_PAYMENT_TERM;
    session.waitingFor = 'weeks';
    saveAIUserSession(chatId, session);

    sendAIMessage(chatId, AI_CONFIG.AI_MESSAGES.ASK_PAYMENT_TERM_WEEKS, { parse_mode: 'Markdown' });
}

/**
 * ⭐ السؤال عن تاريخ الدفع المخصص
 */
function askPaymentTermDate(chatId, session) {
    session.state = AI_CONFIG.AI_CONVERSATION_STATES.WAITING_PAYMENT_TERM;
    session.waitingFor = 'date';
    saveAIUserSession(chatId, session);

    sendAIMessage(chatId, AI_CONFIG.AI_MESSAGES.ASK_PAYMENT_TERM_DATE, { parse_mode: 'Markdown' });
}

/**
 * ⭐ السؤال عن تاريخ استحقاق السلفة/التمويل
 */
function askLoanDueDate(chatId, session) {
    session.state = AI_CONFIG.AI_CONVERSATION_STATES.WAITING_LOAN_DUE_DATE;
    saveAIUserSession(chatId, session);

    sendAIMessage(chatId, '📅 *متى موعد سداد هذه السلفة/التمويل؟*\n\nاكتب التاريخ (مثال: 15/3/2026 أو بعد شهر أو بعد 3 أشهر):', { parse_mode: 'Markdown' });
}

/**
 * ⭐ معالجة إدخال تاريخ استحقاق السلفة
 */
function handleLoanDueDateInput(chatId, text, session) {
    Logger.log('📥 Loan due date input: "' + text + '"');

    // ⭐ تحويل الأرقام العربية للإنجليزية
    const arabicNumerals = '٠١٢٣٤٥٦٧٨٩';
    const englishNumerals = '0123456789';
    let convertedText = text;
    for (let i = 0; i < arabicNumerals.length; i++) {
        convertedText = convertedText.replace(new RegExp(arabicNumerals[i], 'g'), englishNumerals[i]);
    }

    // تحويل النص إلى تاريخ
    let dueDate = null;
    const today = new Date();

    // التعامل مع "بعد شهر" أو "بعد X أشهر"
    const monthMatch = convertedText.match(/بعد\s*(\d+)?\s*(شهر|أشهر)/);
    if (monthMatch) {
        const months = monthMatch[1] ? parseInt(monthMatch[1]) : 1;
        dueDate = new Date(today);
        dueDate.setMonth(dueDate.getMonth() + months);
    }

    // التعامل مع "بعد X يوم"
    const dayMatch = convertedText.match(/بعد\s*(\d+)\s*(يوم|أيام)/);
    if (!dueDate && dayMatch) {
        const days = parseInt(dayMatch[1]);
        dueDate = new Date(today);
        dueDate.setDate(dueDate.getDate() + days);
    }

    // التعامل مع تاريخ صريح
    if (!dueDate) {
        dueDate = parseArabicDate(convertedText);
    }

    if (!dueDate) {
        sendAIMessage(chatId, '❌ لم أفهم التاريخ. يرجى كتابته بوضوح:\n• مثال: 15/3/2026\n• أو: بعد شهر\n• أو: بعد 3 أشهر');
        return;
    }

    // حفظ تاريخ الاستحقاق
    const formattedDate = Utilities.formatDate(dueDate, 'Asia/Istanbul', 'yyyy-MM-dd');
    session.transaction.loan_due_date = formattedDate;

    if (!session.validation.enriched) {
        session.validation.enriched = {};
    }
    session.validation.enriched.loan_due_date = formattedDate;
    session.validation.needsLoanDueDate = false;

    saveAIUserSession(chatId, session);

    const displayDate = Utilities.formatDate(dueDate, 'Asia/Istanbul', 'dd-MM-yyyy');
    sendAIMessage(chatId, `✅ تم تحديد تاريخ السداد: *${displayDate}*`, { parse_mode: 'Markdown' });

    continueValidation(chatId, session);
}

/**
 * ⭐ معالجة إدخال بيانات شرط الدفع (أسابيع أو تاريخ)
 */
function handlePaymentTermInput(chatId, text, session) {
    Logger.log('📥 Payment term input: "' + text + '", waitingFor: ' + session.waitingFor);

    // ⭐ تحويل الأرقام العربية للإنجليزية
    const arabicNumerals = '٠١٢٣٤٥٦٧٨٩';
    const englishNumerals = '0123456789';
    let convertedText = text;
    for (let i = 0; i < arabicNumerals.length; i++) {
        convertedText = convertedText.replace(new RegExp(arabicNumerals[i], 'g'), englishNumerals[i]);
    }
    Logger.log('📥 Converted text: "' + convertedText + '"');

    // ⭐ التحقق من وجود البيانات
    if (!session.transaction || !session.validation) {
        Logger.log('❌ Session data missing in handlePaymentTermInput');
        sendAIMessage(chatId, '⚠️ حدث خطأ. يرجى إعادة إرسال الحركة.');
        resetAIUserSession(chatId);
        return;
    }

    if (!session.validation.enriched) {
        session.validation.enriched = {};
    }

    if (session.waitingFor === 'weeks') {
        // إدخال عدد الأسابيع
        const weeks = parseInt(convertedText.replace(/[^0-9]/g, ''));
        Logger.log('📥 Parsed weeks: ' + weeks);

        if (isNaN(weeks) || weeks <= 0) {
            sendAIMessage(chatId, '❌ يرجى إدخال رقم صحيح (مثال: 2 أو ٢):');
            return;
        }

        session.transaction.payment_term_weeks = weeks;
        session.validation.enriched.payment_term_weeks = weeks;
        session.validation.needsPaymentTermWeeks = false;
        delete session.waitingFor;
        saveAIUserSession(chatId, session);

        sendAIMessage(chatId, `✅ تم تحديد: الدفع بعد *${weeks}* أسبوع من التسليم`, { parse_mode: 'Markdown' });
        continueValidation(chatId, session);

    } else if (session.waitingFor === 'date') {
        // إدخال تاريخ مخصص
        const parsedDate = parseArabicDate(convertedText);
        if (!parsedDate) {
            sendAIMessage(chatId, '❌ تنسيق التاريخ غير صحيح. يرجى الإدخال بصيغة: 15/2/2026 أو ١٥/٢/٢٠٢٦');
            return;
        }

        session.transaction.payment_term_date = parsedDate;
        session.validation.enriched.payment_term_date = parsedDate;
        session.validation.needsPaymentTermDate = false;
        delete session.waitingFor;
        saveAIUserSession(chatId, session);

        sendAIMessage(chatId, `✅ تم تحديد تاريخ الدفع: *${parsedDate}*`, { parse_mode: 'Markdown' });
        continueValidation(chatId, session);
    }
}

/**
 * ⭐ متابعة التحقق من البيانات بعد إكمال حقل
 */
function continueValidation(chatId, session) {
    // ⭐ التحقق من المشروع (اختياري)
    if (session.validation.needsProjectSelection) {
        askProjectSelection(chatId, session);
        return;
    }

    // التحقق من طريقة الدفع
    if (session.validation.needsPaymentMethod) {
        askPaymentMethod(chatId, session);
        return;
    }

    // التحقق من العملة
    if (session.validation.needsCurrency) {
        askCurrency(chatId, session);
        return;
    }

    // التحقق من سعر الصرف
    if (session.validation.needsExchangeRate) {
        askExchangeRate(chatId, session);
        return;
    }

    // التحقق من شرط الدفع (للاستحقاقات)
    if (session.validation.needsPaymentTerm) {
        askPaymentTerm(chatId, session);
        return;
    }

    // التحقق من عدد الأسابيع
    if (session.validation.needsPaymentTermWeeks) {
        askPaymentTermWeeks(chatId, session);
        return;
    }

    // التحقق من تاريخ الدفع المخصص
    if (session.validation.needsPaymentTermDate) {
        askPaymentTermDate(chatId, session);
        return;
    }

    // ⭐ التحقق من تاريخ استحقاق السلفة/التمويل
    if (session.validation.needsLoanDueDate) {
        askLoanDueDate(chatId, session);
        return;
    }

    // التحقق من الطرف الجديد
    if (session.validation.needsPartyConfirmation) {
        askNewPartyConfirmation(chatId, session);
        return;
    }

    // كل شيء تمام - عرض التأكيد
    session.state = AI_CONFIG.AI_CONVERSATION_STATES.CONFIRM_WAIT;
    saveAIUserSession(chatId, session);
    showTransactionConfirmation(chatId, session);
}

/**
 * إضافة طرف جديد لقاعدة البيانات
 */
function addNewParty(name, type) {
    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        // جرب أطراف البوت أولاً، ثم قاعدة بيانات الأطراف
        let partiesSheet = ss.getSheetByName('أطراف البوت');
        if (!partiesSheet) {
            partiesSheet = ss.getSheetByName('قاعدة بيانات الأطراف');
        }
        if (!partiesSheet) {
            partiesSheet = ss.getSheetByName('الأطراف');
        }

        if (!partiesSheet) {
            Logger.log('❌ شيت الأطراف غير موجود');
            return false;
        }

        // إضافة الطرف في نهاية الشيت
        const lastRow = partiesSheet.getLastRow() + 1;
        partiesSheet.getRange(lastRow, 1).setValue(name);
        partiesSheet.getRange(lastRow, 2).setValue(type);

        Logger.log(`✅ تم إضافة الطرف: ${name} (${type}) في شيت: ${partiesSheet.getName()}`);
        return true;

    } catch (error) {
        Logger.log('Error adding party: ' + error.message);
        return false;
    }
}


// ==================== عرض التأكيد ====================

/**
 * عرض ملخص الحركة للتأكيد
 */
function showTransactionConfirmation(chatId, session) {
    const summary = buildTransactionSummary(session.transaction);

    session.state = AI_CONFIG.AI_CONVERSATION_STATES.WAITING_CONFIRMATION;

    sendAIMessage(chatId, summary, {
        parse_mode: 'Markdown',
        reply_markup: JSON.stringify(AI_CONFIG.AI_KEYBOARDS.CONFIRMATION)
    });
}


// ==================== معالجة الأزرار ====================

/**
 * معالجة ضغطات الأزرار
 */
function handleAICallback(callbackQuery) {
    // ⭐ فحص وجود البيانات المطلوبة
    if (!callbackQuery || !callbackQuery.message || !callbackQuery.data) {
        Logger.log('⚠️ Callback query missing required data: ' + JSON.stringify(callbackQuery));
        return;
    }

    const chatId = callbackQuery.message.chat.id;
    const messageId = callbackQuery.message.message_id;
    const data = callbackQuery.data;
    const user = callbackQuery.from;

    // ⭐ منع الضغط المتكرر (قفل لمدة 3 ثواني)
    const cache = CacheService.getScriptCache();
    const lockKey = `CALLBACK_LOCK_${chatId}_${data}`;
    const isLocked = cache.get(lockKey);

    if (isLocked) {
        // الرد على الـ callback برسالة "جاري المعالجة"
        answerAICallback(callbackQuery.id, '⏳ جاري المعالجة...');
        Logger.log('⚠️ Duplicate callback ignored: ' + data);
        return;
    }

    // تفعيل القفل لمدة 3 ثواني
    cache.put(lockKey, 'locked', 3);

    // الرد على الـ callback فوراً
    answerAICallback(callbackQuery.id, '✅');

    // التحقق من الصلاحيات
    const permission = checkAIUserPermission(chatId, user);
    if (!permission.authorized) {
        return;
    }

    const session = getAIUserSession(chatId);

    // ⭐ فحص وجود الجلسة والبيانات المطلوبة
    if (!session) {
        Logger.log('⚠️ Session missing for callback: ' + data);
        sendAIMessage(chatId, '⚠️ انتهت الجلسة. يرجى إعادة إرسال الحركة.');
        resetAIUserSession(chatId);
        return;
    }

    // ⭐ السماح بـ cancel حتى بدون بيانات كاملة
    if (data === 'ai_cancel') {
        Logger.log('📥 Cancel callback - processing immediately');
        sendAIMessage(chatId, AI_CONFIG.AI_MESSAGES.CANCELLED);
        resetAIUserSession(chatId);
        return;
    }

    // ⭐ معالجة callbacks التقارير (لا تحتاج validation أو transaction)
    if (isReportCallback(data)) {
        Logger.log('📊 Report callback detected: ' + data);
        handleReportCallback(chatId, data, session);
        return;
    }

    // ⭐ معالجة callbacks الأوردر المشترك
    if (data.startsWith('shared_')) {
        Logger.log('📦 Shared order callback detected: ' + data);
        handleSharedOrderCallback(chatId, messageId, data, session);
        return;
    }

    // ⭐ العمليات التي لا تحتاج session (تعمل مباشرة من الشيت)
    const isEditResend = data === 'edit_resend' || data.startsWith('edit_resend_');
    const isDeleteRejected = data === 'delete_rejected' || data.startsWith('delete_rejected_');
    const noSessionRequired = isEditResend || isDeleteRejected;

    // ⭐ فحص وجود الـ validation للعمليات الأخرى
    if (!session.validation && !noSessionRequired) {
        Logger.log('⚠️ Validation missing for callback: ' + data);
        sendAIMessage(chatId, '⚠️ انتهت الجلسة. يرجى إعادة إرسال الحركة.');
        resetAIUserSession(chatId);
        return;
    }

    // ⭐ فحص وجود transaction (للعمليات التي تحتاجه)
    if (!session.transaction && !data.startsWith('ai_cancel') && !noSessionRequired) {
        Logger.log('⚠️ Transaction missing for callback: ' + data);
        sendAIMessage(chatId, '⚠️ لا توجد حركة للمعالجة. يرجى إعادة إرسال الحركة.');
        resetAIUserSession(chatId);
        return;
    }

    // ⭐ إنشاء enriched إذا لم يكن موجوداً
    if (session.validation && !session.validation.enriched) {
        Logger.log('⚠️ Creating session.validation.enriched in callback handler');
        session.validation.enriched = {};
        saveAIUserSession(chatId, session);
    }

    // معالجة حسب نوع الـ callback
    if (data.startsWith('ai_confirm')) {
        handleAIConfirmation(chatId, session, user);
    } else if (data.startsWith('ai_edit_item_')) {
        // ✅ اختيار بند من الاقتراحات أثناء التعديل
        const item = data.replace('ai_edit_item_', '');
        session.transaction.item = item;
        saveAIUserSession(chatId, session);
        sendAIMessage(chatId, `✅ تم اختيار البند: ${item}\n\nهل تريد تعديل حقل آخر؟`, {
            parse_mode: 'Markdown',
            reply_markup: JSON.stringify(AI_CONFIG.AI_KEYBOARDS.EDIT_FIELDS)
        });
    } else if (data.startsWith('ai_edit_project_')) {
        // ✅ اختيار مشروع من الاقتراحات أثناء التعديل
        const project = data.replace('ai_edit_project_', '');
        const context = loadAIContext();
        const projMatch = context.projects.find(p =>
            (typeof p === 'object' ? p.name : p) === project
        );
        session.transaction.project = project;
        session.transaction.project_code = projMatch && typeof projMatch === 'object' ? projMatch.code : '';
        saveAIUserSession(chatId, session);
        sendAIMessage(chatId, `✅ تم اختيار المشروع: ${project}\n\nهل تريد تعديل حقل آخر؟`, {
            parse_mode: 'Markdown',
            reply_markup: JSON.stringify(AI_CONFIG.AI_KEYBOARDS.EDIT_FIELDS)
        });
    } else if (data === 'ai_edit_cancel') {
        // ✅ إلغاء اقتراح التعديل
        sendAIMessage(chatId, '❌ تم الإلغاء\n\nهل تريد تعديل حقل آخر؟', {
            parse_mode: 'Markdown',
            reply_markup: JSON.stringify(AI_CONFIG.AI_KEYBOARDS.EDIT_FIELDS)
        });
    } else if (data.startsWith('ai_editnature_')) {
        // ✅ اختيار طبيعة الحركة أثناء التعديل
        const nature = data.replace('ai_editnature_', '');
        session.transaction.nature = nature;
        // مسح التصنيف لأنه قد لا يتوافق مع الطبيعة الجديدة
        session.transaction.classification = '';
        saveAIUserSession(chatId, session);
        sendAIMessage(chatId, `✅ تم تغيير طبيعة الحركة إلى: ${nature}\n\n⚠️ يرجى تحديث التصنيف ليتوافق مع الطبيعة الجديدة\n\nهل تريد تعديل حقل آخر؟`, {
            parse_mode: 'Markdown',
            reply_markup: JSON.stringify(AI_CONFIG.AI_KEYBOARDS.EDIT_FIELDS)
        });
    } else if (data.startsWith('ai_editclass_')) {
        // ✅ اختيار التصنيف أثناء التعديل
        const classification = data.replace('ai_editclass_', '');
        session.transaction.classification = classification;
        saveAIUserSession(chatId, session);
        sendAIMessage(chatId, `✅ تم تغيير التصنيف إلى: ${classification}\n\nهل تريد تعديل حقل آخر؟`, {
            parse_mode: 'Markdown',
            reply_markup: JSON.stringify(AI_CONFIG.AI_KEYBOARDS.EDIT_FIELDS)
        });
    } else if (data.startsWith('ai_edit')) {
        handleEditRequest(chatId, data, session, messageId);
    } else if (data === 'ai_skip_project') {
        // ⭐ تخطي اختيار المشروع (بدون مشروع)
        Logger.log('📥 Skip project callback - continuing without project');
        handleSkipProject(chatId, session);
    } else if (data.startsWith('ai_project_')) {
        const project = data.replace('ai_project_', '');
        handleProjectCallback(chatId, project, session);
    } else if (data.startsWith('ai_select_party_')) {
        // ⭐ معالجة اختيار طرف من القائمة
        const index = parseInt(data.replace('ai_select_party_', ''));
        handleSelectPartyFromSuggestions(chatId, index, session);
    } else if (data.startsWith('ai_payment_')) {
        // ⭐ معالجة اختيار طريقة الدفع
        Logger.log('═══════════════════════════════════════');
        Logger.log('📥 PAYMENT METHOD CALLBACK');
        Logger.log('📥 Full callback_data: ' + data);
        Logger.log('📥 chatId: ' + chatId);
        const method = data.replace('ai_payment_', '');
        Logger.log('📥 Extracted method: "' + method + '"');
        Logger.log('📥 Session state: ' + (session ? session.state : 'null'));
        Logger.log('📥 Has transaction: ' + (session && session.transaction ? 'yes' : 'no'));
        Logger.log('📥 Has validation: ' + (session && session.validation ? 'yes' : 'no'));
        Logger.log('═══════════════════════════════════════');
        handleAIPaymentMethodSelection(chatId, method, session);
        Logger.log('✅ Payment method handler completed');
    } else if (data.startsWith('ai_currency_')) {
        // ⭐ معالجة اختيار العملة
        const currency = data.replace('ai_currency_', '');
        // ✅ التحقق من وضع التعديل
        if (session.state === AI_CONFIG.AI_CONVERSATION_STATES.WAITING_EDIT && session.editingField === 'currency') {
            session.transaction.currency = currency;
            saveAIUserSession(chatId, session);
            sendAIMessage(chatId, `✅ تم تغيير العملة إلى: ${currency}\n\nهل تريد تعديل حقل آخر؟`, {
                parse_mode: 'Markdown',
                reply_markup: JSON.stringify(AI_CONFIG.AI_KEYBOARDS.EDIT_FIELDS)
            });
        } else {
            handleAICurrencySelection(chatId, currency, session);
        }
    } else if (data.startsWith('ai_term_')) {
        // ⭐ معالجة اختيار شرط الدفع
        const term = data.replace('ai_term_', '');
        handleAIPaymentTermSelection(chatId, term, session);
    } else if (data.startsWith('ai_add_party_')) {
        // معالجة تأكيد إضافة طرف جديد
        handleNewPartyConfirmation(chatId, data, session);
    } else if (data.startsWith('ai_party_type_')) {
        // معالجة اختيار نوع الطرف الجديد
        const partyType = data.replace('ai_party_type_', '');
        session.newPartyType = partyType;
        saveAIUserSession(chatId, session);
        sendAIMessage(chatId, `✅ تم تحديد النوع: ${partyType}`);
        // إعادة عرض رسالة التأكيد
        askNewPartyConfirmation(chatId, session);
    } else if (data.startsWith('ai_party_')) {
        const party = data.replace('ai_party_', '');
        handlePartyCallback(chatId, party, session);
    } else if (data.startsWith('ai_partytype_')) {
        const partyType = data.replace('ai_partytype_', '');
        handleAINewPartyType(chatId, partyType, session);
    } else if (data === 'ai_add_party') {
        showNewPartyTypeSelection(chatId, session);
    } else if (data === 'ai_edit_done') {
        showTransactionConfirmation(chatId, session);
    } else if (isEditResend) {
        // ⭐ تعديل وإعادة إرسال حركة مرفوضة
        const transactionId = data.startsWith('edit_resend_') ? data.replace('edit_resend_', '') : null;
        handleEditRejectedTransaction(chatId, session, callbackQuery.from, transactionId);
    } else if (isDeleteRejected) {
        // ⭐ حذف حركة مرفوضة
        const transactionId = data.startsWith('delete_rejected_') ? data.replace('delete_rejected_', '') : null;
        handleDeleteRejectedTransaction(chatId, transactionId);
    }
}

/**
 * ⭐ معالجة تعديل حركة مرفوضة
 * @param {string} transactionId - رقم الحركة المحددة (اختياري - إذا لم يتم تحديده يبحث عن آخر حركة مرفوضة)
 */
function handleEditRejectedTransaction(chatId, session, user, transactionId) {
    try {
        Logger.log('📝 handleEditRejectedTransaction called for chatId: ' + chatId + ', transactionId: ' + transactionId);

        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const sheet = ss.getSheetByName(CONFIG.SHEETS.BOT_TRANSACTIONS);

        if (!sheet) {
            sendAIMessage(chatId, '❌ لم يتم العثور على شيت الحركات');
            return;
        }

        const data = sheet.getDataRange().getValues();
        const columns = BOT_CONFIG.BOT_TRANSACTIONS_COLUMNS;

        let rejectedTransaction = null;
        let rejectedRowIndex = -1;

        // ⭐ البحث بطريقتين:
        // 1. إذا تم تحديد رقم الحركة، نبحث عنه مباشرة
        // 2. إذا لم يتم تحديده، نبحث عن آخر حركة مرفوضة للمستخدم (للتوافق مع الإصدارات القديمة)
        for (let i = data.length - 1; i >= 1; i--) {
            const row = data[i];
            const rowTransactionId = String(row[columns.TRANSACTION_ID.index - 1] || '');
            const rowChatId = String(row[columns.TELEGRAM_CHAT_ID.index - 1] || '');
            const status = row[columns.REVIEW_STATUS.index - 1];

            // إذا تم تحديد رقم الحركة
            if (transactionId) {
                if (rowTransactionId === transactionId && status === CONFIG.TELEGRAM_BOT.REVIEW_STATUS.REJECTED) {
                    rejectedTransaction = row;
                    rejectedRowIndex = i + 1;
                    break;
                }
            } else {
                // البحث عن آخر حركة مرفوضة للمستخدم
                if (rowChatId === String(chatId) && status === CONFIG.TELEGRAM_BOT.REVIEW_STATUS.REJECTED) {
                    rejectedTransaction = row;
                    rejectedRowIndex = i + 1;
                    break;
                }
            }
        }

        if (!rejectedTransaction) {
            sendAIMessage(chatId, '❌ لم يتم العثور على حركة مرفوضة للتعديل.\n\nيرجى إرسال الحركة من جديد.');
            return;
        }

        // استخراج بيانات الحركة المرفوضة
        const transactionData = {
            nature: rejectedTransaction[columns.NATURE.index - 1] || '',
            classification: rejectedTransaction[columns.CLASSIFICATION.index - 1] || '',
            project_code: rejectedTransaction[columns.PROJECT_CODE.index - 1] || '',
            project: rejectedTransaction[columns.PROJECT_NAME.index - 1] || '',
            item: rejectedTransaction[columns.ITEM.index - 1] || '',
            details: rejectedTransaction[columns.DETAILS.index - 1] || '',
            party: rejectedTransaction[columns.PARTY_NAME.index - 1] || '',
            amount: rejectedTransaction[columns.AMOUNT.index - 1] || 0,
            currency: rejectedTransaction[columns.CURRENCY.index - 1] || 'USD',
            exchangeRate: rejectedTransaction[columns.EXCHANGE_RATE.index - 1] || 1,
            payment_method: rejectedTransaction[columns.PAYMENT_METHOD.index - 1] || '',
            payment_term: rejectedTransaction[columns.PAYMENT_TERM_TYPE.index - 1] || '',
            due_date: rejectedTransaction[columns.DUE_DATE.index - 1] || '',
            originalText: rejectedTransaction[columns.NOTES.index - 1] || ''
        };

        const rejectionReason = rejectedTransaction[columns.REVIEW_NOTES.index - 1] || 'غير محدد';

        // حفظ في الجلسة
        session.transaction = transactionData;
        session.validation = { enriched: transactionData };
        session.rejectedRowIndex = rejectedRowIndex;
        session.isEditingRejected = true;
        session.state = AI_CONFIG.AI_CONVERSATION_STATES.IDLE;
        saveAIUserSession(chatId, session);

        // عرض ملخص الحركة المرفوضة مع خيارات التعديل
        let message = '✏️ *تعديل الحركة المرفوضة*\n';
        message += '━━━━━━━━━━━━━━━━━━━━\n\n';
        message += `❌ *سبب الرفض:* ${rejectionReason}\n\n`;

        // ⭐ عرض النص الأصلي الذي أرسله المستخدم
        if (transactionData.originalText) {
            message += '📝 *النص الأصلي الذي أرسلته:*\n';
            message += `"${transactionData.originalText}"\n\n`;
        }

        message += '📋 *بيانات الحركة:*\n';
        message += `• الطبيعة: ${transactionData.nature}\n`;
        message += `• التصنيف: ${transactionData.classification}\n`;
        message += `• المشروع: ${transactionData.project || 'غير محدد'}\n`;
        message += `• الطرف: ${transactionData.party}\n`;
        message += `• المبلغ: ${transactionData.amount} ${transactionData.currency}\n`;
        message += `• طريقة الدفع: ${transactionData.payment_method}\n\n`;
        message += '━━━━━━━━━━━━━━━━━━━━\n';
        message += '💡 *أرسل الحركة المعدلة كنص جديد*\n';

        // ⭐ مثال ديناميكي بناءً على بيانات الحركة الفعلية
        const exampleParty = transactionData.party || 'الطرف';
        const exampleAmount = transactionData.amount || '500';
        const exampleCurrency = transactionData.currency || 'دولار';
        message += `مثال: "دفعت لـ${exampleParty} ${exampleAmount} ${exampleCurrency}"`;

        sendAIMessage(chatId, message, { parse_mode: 'Markdown' });

        Logger.log('✅ Rejected transaction loaded for editing');

    } catch (error) {
        Logger.log('❌ Error in handleEditRejectedTransaction: ' + error.message);
        sendAIMessage(chatId, '❌ حدث خطأ أثناء تحميل الحركة. يرجى إرسالها من جديد.');
    }
}

/**
 * ⭐ حذف حركة مرفوضة
 * @param {string} transactionId - رقم الحركة المحددة (اختياري)
 */
function handleDeleteRejectedTransaction(chatId, transactionId) {
    try {
        Logger.log('🗑️ handleDeleteRejectedTransaction called for chatId: ' + chatId + ', transactionId: ' + transactionId);

        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const sheet = ss.getSheetByName(CONFIG.SHEETS.BOT_TRANSACTIONS);

        if (!sheet) {
            sendAIMessage(chatId, '❌ لم يتم العثور على شيت الحركات');
            return;
        }

        const data = sheet.getDataRange().getValues();
        const columns = BOT_CONFIG.BOT_TRANSACTIONS_COLUMNS;

        // البحث عن الحركة المرفوضة
        for (let i = data.length - 1; i >= 1; i--) {
            const row = data[i];
            const rowTransactionId = String(row[columns.TRANSACTION_ID.index - 1] || '');
            const rowChatId = String(row[columns.TELEGRAM_CHAT_ID.index - 1] || '');
            const status = row[columns.REVIEW_STATUS.index - 1];

            let isMatch = false;
            if (transactionId) {
                // البحث برقم الحركة المحدد
                isMatch = rowTransactionId === transactionId && status === CONFIG.TELEGRAM_BOT.REVIEW_STATUS.REJECTED;
            } else {
                // البحث عن آخر حركة مرفوضة للمستخدم (للتوافق مع الإصدارات القديمة)
                isMatch = rowChatId === String(chatId) && status === CONFIG.TELEGRAM_BOT.REVIEW_STATUS.REJECTED;
            }

            if (isMatch) {
                // حذف الصف
                sheet.deleteRow(i + 1);
                sendAIMessage(chatId, '🗑️ تم حذف الحركة المرفوضة بنجاح.\n\nيمكنك إرسال حركة جديدة الآن.');
                Logger.log('✅ Rejected transaction deleted at row: ' + (i + 1));
                return;
            }
        }

        sendAIMessage(chatId, '❌ لم يتم العثور على حركة مرفوضة للحذف.');

    } catch (error) {
        Logger.log('❌ Error in handleDeleteRejectedTransaction: ' + error.message);
        sendAIMessage(chatId, '❌ حدث خطأ أثناء حذف الحركة.');
    }
}

/**
 * الرد على الـ callback query
 */
function answerAICallback(callbackQueryId, text) {
    const token = getAIBotToken();
    const url = `https://api.telegram.org/bot${token}/answerCallbackQuery`;

    const payload = { callback_query_id: callbackQueryId };

    // ⭐ إضافة رسالة toast إذا وُجدت
    if (text) {
        payload.text = text;
        payload.show_alert = false; // رسالة صغيرة في الأعلى
    }

    UrlFetchApp.fetch(url, {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
    });
}


// ==================== التأكيد والحفظ ====================

/**
 * معالجة تأكيد الحركة
 */
function handleAIConfirmation(chatId, session, user) {
    try {
        Logger.log('AI Confirmation started for chatId: ' + chatId);

        // التحقق من وجود بيانات الحركة
        if (!session.transaction) {
            sendAIMessage(chatId, '❌ عذراً، لم أجد بيانات الحركة لتأكيدها. يرجى إعادة المحاولة.');
            return;
        }

        // حفظ الحركة
        const result = saveAITransaction(session.transaction, user, chatId);

        if (result.success) {
            const successMsg = AI_CONFIG.AI_MESSAGES.SUCCESS.replace('#{id}', result.transactionId);
            sendAIMessage(chatId, successMsg, { parse_mode: 'Markdown' });

            // إرسال إشعار للمراجعين (اختياري)
            notifyReviewers(result.transactionId, session.transaction);
        } else {
            // إرسال رسالة الخطأ المحددة
            sendAIMessage(chatId, '❌ فشل حفظ الحركة:\n' + result.error);
        }

        resetAIUserSession(chatId);

    } catch (error) {
        Logger.log('Confirmation Error: ' + error.message);
        sendAIMessage(chatId, '❌ خطأ غير متوقع عند التأكيد:\n' + error.message);
    }
}

/**
 * حفظ الحركة في شيت حركات البوت
 */
/**
 * حفظ الحركة مباشرة في شيت الحركات الرئيسي
 * ✅ البنية الجديدة: الحفظ مباشرة بدون مرحلة المراجعة
 */
function saveAITransaction(transaction, user, chatId) {
    Logger.log('🚀 بدء عملية حفظ الحركة (البوت الذكي)...');
    try {
        const userName = `${user.first_name || ''} ${user.last_name || ''}`.trim();

        // ✅ إذا كان طرف جديد، أضفه مباشرة لشيت الأطراف الرئيسي
        if (transaction.isNewParty) {
            try {
                Logger.log('👤 إضافة طرف جديد مباشرة...');
                const partyType = transaction.partyType || inferPartyType(transaction.nature, transaction.classification);
                const partyResult = addPartyDirectly({
                    name: transaction.party,
                    type: partyType,
                    notes: `(مضاف من البوت الذكي بواسطة ${userName})`
                });
                if (partyResult.success) {
                    Logger.log('✅ تم إضافة الطرف الجديد.');
                } else if (partyResult.alreadyExists) {
                    Logger.log('⚠️ الطرف موجود مسبقاً.');
                } else {
                    Logger.log('⚠️ فشل إضافة الطرف: ' + partyResult.error);
                }
            } catch (e) {
                Logger.log('⚠️ تحذير: فشل إضافة الطرف الجديد: ' + e.message);
            }
        }

        // تجهيز بيانات الحركة
        const transactionDate = transaction.due_date && transaction.due_date !== 'TODAY'
            ? transaction.due_date
            : new Date();

        // ⭐ تجهيز التفاصيل مع تاريخ استحقاق السلفة إذا وجد
        let details = transaction.details || '';
        if (transaction.loan_due_date) {
            const loanDueDateNote = `[تاريخ السداد: ${transaction.loan_due_date}]`;
            details = details ? `${details} ${loanDueDateNote}` : loanDueDateNote;
        }

        const transactionData = {
            date: transactionDate,
            nature: transaction.nature,
            classification: transaction.classification,
            projectCode: transaction.project_code || '',
            projectName: transaction.project || '',
            item: transaction.item || '',
            details: details,
            partyName: transaction.party,
            amount: transaction.amount,
            currency: transaction.currency,
            exchangeRate: transaction.exchangeRate || 1,
            paymentMethod: transaction.payment_method || 'تحويل بنكي',
            paymentTermType: transaction.payment_term || 'فوري',
            weeks: transaction.payment_term_weeks || '',
            customDate: transaction.payment_term_date || transaction.loan_due_date || '',
            telegramUser: userName,
            chatId: chatId,
            attachmentUrl: '',
            isNewParty: transaction.isNewParty,
            unitCount: transaction.unit_count || transaction.unitCount || '',
            notes: transaction.originalText ? `النص الأصلي: ${transaction.originalText}` : ''
        };

        // ✅ حفظ الحركة مباشرة في شيت الحركات الرئيسي
        const result = addTransactionDirectly(transactionData, '🤖 بوت ذكي');

        if (result.success) {
            Logger.log('✅ تم حفظ الحركة بنجاح - رقم: ' + result.transactionId);
        } else {
            Logger.log('❌ فشل حفظ الحركة: ' + result.error);
        }

        return result;

    } catch (error) {
        Logger.log('❌ Save Transaction Error: ' + error.message);
        Logger.log(error.stack);
        return {
            success: false,
            error: error.message
        };
    }
}

/**
 * إضافة طرف جديد من البوت الذكي
 */
function addNewPartyFromAI(transaction, user, chatId) {
    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const sheet = ss.getSheetByName(CONFIG.SHEETS.BOT_PARTIES);

        if (!sheet) return;

        const timestamp = Utilities.formatDate(new Date(), 'Asia/Istanbul', 'yyyy-MM-dd HH:mm:ss');

        // تحديد نوع الطرف
        const partyType = transaction.partyType || inferPartyType(transaction.nature, transaction.classification);

        const rowData = [
            transaction.party,          // اسم الطرف
            partyType,                  // نوع الطرف
            '',                         // التخصص
            '',                         // رقم الهاتف
            '',                         // البريد
            '',                         // المدينة
            '',                         // طريقة الدفع
            '',                         // بيانات البنك
            'تمت إضافته من البوت الذكي', // ملاحظات
            CONFIG.TELEGRAM_BOT.REVIEW_STATUS.PENDING,
            `${user.first_name || ''} ${user.last_name || ''}`.trim(),
            chatId,
            timestamp,
            '',
            '',
            ''
        ];

        sheet.appendRow(rowData);

    } catch (error) {
        Logger.log('Add New Party Error: ' + error.message);
    }
}

/**
 * إنشاء رقم حركة فريد
 */
function generateTransactionId() {
    const now = new Date();
    const timestamp = Utilities.formatDate(now, 'Asia/Istanbul', 'yyMMddHHmmss');
    const random = Math.floor(Math.random() * 100).toString().padStart(2, '0');
    return `AI${timestamp}${random}`;
}


// ==================== التعديل ====================

/**
 * معالجة طلب التعديل
 */
function handleEditRequest(chatId, data, session, messageId) {
    if (data === 'ai_edit') {
        // عرض قائمة الحقول للتعديل
        sendAIMessage(chatId, '✏️ *اختر الحقل الذي تريد تعديله:*', {
            parse_mode: 'Markdown',
            reply_markup: JSON.stringify(AI_CONFIG.AI_KEYBOARDS.EDIT_FIELDS)
        });
        session.state = AI_CONFIG.AI_CONVERSATION_STATES.WAITING_EDIT;
        saveAIUserSession(chatId, session);
        return;
    }

    // معالجة تعديل حقل محدد
    const field = data.replace('ai_edit_', '');

    if (field === 'done') {
        showTransactionConfirmation(chatId, session);
        return;
    }

    session.editingField = field;
    session.state = AI_CONFIG.AI_CONVERSATION_STATES.WAITING_EDIT;
    saveAIUserSession(chatId, session);

    // ⭐ عرض لوحات مفاتيح للحقول التي تحتاج اختيار
    if (field === 'currency') {
        // عرض لوحة مفاتيح العملات
        sendAIMessage(chatId, '💱 اختر العملة الجديدة:', {
            parse_mode: 'Markdown',
            reply_markup: JSON.stringify(AI_CONFIG.AI_KEYBOARDS.CURRENCY)
        });
        return;
    }

    if (field === 'nature') {
        // عرض لوحة مفاتيح طبيعة الحركة
        const natureKeyboard = {
            inline_keyboard: [
                [
                    { text: '📤 استحقاق مصروف', callback_data: 'ai_editnature_استحقاق مصروف' },
                    { text: '💸 دفعة مصروف', callback_data: 'ai_editnature_دفعة مصروف' }
                ],
                [
                    { text: '📥 استحقاق إيراد', callback_data: 'ai_editnature_استحقاق إيراد' },
                    { text: '💰 تحصيل إيراد', callback_data: 'ai_editnature_تحصيل إيراد' }
                ],
                [
                    { text: '🏦 تمويل (دخول قرض)', callback_data: 'ai_editnature_تمويل (دخول قرض)' },
                    { text: '💳 سداد تمويل', callback_data: 'ai_editnature_سداد تمويل' }
                ],
                [
                    { text: '🔒 تأمين مدفوع', callback_data: 'ai_editnature_تأمين مدفوع للقناة' },
                    { text: '🔓 استرداد تأمين', callback_data: 'ai_editnature_استرداد تأمين من القناة' }
                ],
                [
                    { text: '❌ إلغاء', callback_data: 'ai_edit_cancel' }
                ]
            ]
        };
        sendAIMessage(chatId, '📤 اختر طبيعة الحركة الجديدة:', {
            parse_mode: 'Markdown',
            reply_markup: JSON.stringify(natureKeyboard)
        });
        return;
    }

    if (field === 'classification') {
        // عرض التصنيفات بناءً على طبيعة الحركة الحالية
        const nature = session.transaction.nature || '';
        let classificationKeyboard;

        if (nature.includes('مصروف')) {
            classificationKeyboard = {
                inline_keyboard: [
                    [{ text: '🎬 إنتاجي', callback_data: 'ai_editclass_إنتاجي' }],
                    [{ text: '🏢 إداري وعمومي', callback_data: 'ai_editclass_إداري وعمومي' }],
                    [{ text: '📢 تسويقي', callback_data: 'ai_editclass_تسويقي' }],
                    [{ text: '❌ إلغاء', callback_data: 'ai_edit_cancel' }]
                ]
            };
        } else if (nature.includes('إيراد')) {
            classificationKeyboard = {
                inline_keyboard: [
                    [{ text: '🎬 إيراد إنتاجي', callback_data: 'ai_editclass_إيراد إنتاجي' }],
                    [{ text: '📺 إيراد قناة', callback_data: 'ai_editclass_إيراد قناة' }],
                    [{ text: '💼 إيراد آخر', callback_data: 'ai_editclass_إيراد آخر' }],
                    [{ text: '❌ إلغاء', callback_data: 'ai_edit_cancel' }]
                ]
            };
        } else {
            classificationKeyboard = {
                inline_keyboard: [
                    [{ text: '🎬 إنتاجي', callback_data: 'ai_editclass_إنتاجي' }],
                    [{ text: '🏢 إداري وعمومي', callback_data: 'ai_editclass_إداري وعمومي' }],
                    [{ text: '📢 تسويقي', callback_data: 'ai_editclass_تسويقي' }],
                    [{ text: '📺 إيراد قناة', callback_data: 'ai_editclass_إيراد قناة' }],
                    [{ text: '❌ إلغاء', callback_data: 'ai_edit_cancel' }]
                ]
            };
        }

        sendAIMessage(chatId, '📊 اختر التصنيف الجديد:', {
            parse_mode: 'Markdown',
            reply_markup: JSON.stringify(classificationKeyboard)
        });
        return;
    }

    // باقي الحقول تحتاج إدخال نصي
    const fieldMessages = {
        'project': '🎬 اكتب اسم المشروع:',
        'item': '📁 اكتب اسم البند:',
        'party': '👤 اكتب اسم الطرف:',
        'amount': '💰 اكتب المبلغ الجديد:',
        'date': '📅 اكتب التاريخ (مثال: 15/01/2025):',
        'details': '📝 اكتب التفاصيل:'
    };

    sendAIMessage(chatId, fieldMessages[field] || 'اكتب القيمة الجديدة:', {
        parse_mode: 'Markdown'
    });
}

/**
 * معالجة إدخال التعديل
 */
function handleEditInput(chatId, text, session) {
    const field = session.editingField;

    // تحديث الحقل
    switch (field) {
        case 'amount':
            // ⭐ تحويل الأرقام العربية إلى إنجليزية
            const arabicToEnglish = {
                '٠': '0', '١': '1', '٢': '2', '٣': '3', '٤': '4',
                '٥': '5', '٦': '6', '٧': '7', '٨': '8', '٩': '9'
            };
            let normalizedText = text;
            for (const [arabic, english] of Object.entries(arabicToEnglish)) {
                normalizedText = normalizedText.replace(new RegExp(arabic, 'g'), english);
            }
            const amount = parseFloat(normalizedText.replace(/[^0-9.]/g, ''));
            if (isNaN(amount) || amount <= 0) {
                sendAIMessage(chatId, '❌ المبلغ غير صحيح. اكتب رقماً صحيحاً (مثال: 2000 أو ٢٠٠٠):');
                return;
            }
            session.transaction.amount = amount;
            break;

        case 'date':
        case 'due_date':
            session.transaction.due_date = parseArabicDate(text);
            break;

        case 'item':
            // البحث الذكي عن البند
            const context = loadAIContext();
            const itemMatch = matchItem(text, context.items);

            if (itemMatch.found && itemMatch.score >= 0.8) {
                // تطابق عالي - نقبله مباشرة
                session.transaction.item = itemMatch.match;
                saveAIUserSession(chatId, session);
                sendAIMessage(chatId, '✅ تم التحديث!\n\nهل تريد تعديل حقل آخر؟', {
                    parse_mode: 'Markdown',
                    reply_markup: JSON.stringify(AI_CONFIG.AI_KEYBOARDS.EDIT_FIELDS)
                });
            } else if (itemMatch.found && itemMatch.score >= 0.4) {
                // تطابق متوسط - نعرض اقتراحات
                const keyboard = {
                    inline_keyboard: [
                        [{ text: `✅ ${itemMatch.match}`, callback_data: `ai_edit_item_${itemMatch.match}` }]
                    ]
                };

                // إضافة البدائل
                if (itemMatch.alternatives && itemMatch.alternatives.length > 0) {
                    itemMatch.alternatives.slice(0, 2).forEach(alt => {
                        keyboard.inline_keyboard.push([
                            { text: `📁 ${alt}`, callback_data: `ai_edit_item_${alt}` }
                        ]);
                    });
                }

                // إضافة خيار الإلغاء
                keyboard.inline_keyboard.push([
                    { text: '❌ إلغاء', callback_data: 'ai_edit_cancel' }
                ]);

                sendAIMessage(chatId, `🔍 هل تقصد أحد هذه البنود؟`, {
                    parse_mode: 'Markdown',
                    reply_markup: JSON.stringify(keyboard)
                });
            } else {
                // لم نجد تطابق - نقبل كما هو مع تحذير
                session.transaction.item = text;
                saveAIUserSession(chatId, session);
                sendAIMessage(chatId, `⚠️ تم حفظ البند: "${text}"\n(لم نجد تطابق في قاعدة البيانات)\n\nهل تريد تعديل حقل آخر؟`, {
                    parse_mode: 'Markdown',
                    reply_markup: JSON.stringify(AI_CONFIG.AI_KEYBOARDS.EDIT_FIELDS)
                });
            }
            return; // نخرج هنا لأننا عالجنا الأمر

        case 'project':
            // البحث الذكي عن المشروع
            const ctx = loadAIContext();
            const projMatch = matchProject(text, ctx.projects);

            if (projMatch.found && projMatch.score >= 0.8) {
                session.transaction.project = projMatch.match;
                session.transaction.project_code = projMatch.code || '';
                saveAIUserSession(chatId, session);
                sendAIMessage(chatId, '✅ تم التحديث!\n\nهل تريد تعديل حقل آخر؟', {
                    parse_mode: 'Markdown',
                    reply_markup: JSON.stringify(AI_CONFIG.AI_KEYBOARDS.EDIT_FIELDS)
                });
            } else if (projMatch.found && projMatch.score >= 0.4) {
                const keyboard = {
                    inline_keyboard: [
                        [{ text: `✅ ${projMatch.match}`, callback_data: `ai_edit_project_${projMatch.match}` }]
                    ]
                };

                if (projMatch.alternatives && projMatch.alternatives.length > 0) {
                    projMatch.alternatives.slice(0, 2).forEach(alt => {
                        keyboard.inline_keyboard.push([
                            { text: `🎬 ${alt}`, callback_data: `ai_edit_project_${alt}` }
                        ]);
                    });
                }

                keyboard.inline_keyboard.push([
                    { text: '❌ إلغاء', callback_data: 'ai_edit_cancel' }
                ]);

                sendAIMessage(chatId, `🔍 هل تقصد أحد هذه المشاريع؟`, {
                    parse_mode: 'Markdown',
                    reply_markup: JSON.stringify(keyboard)
                });
            } else {
                session.transaction.project = text;
                session.transaction.project_code = '';
                saveAIUserSession(chatId, session);
                sendAIMessage(chatId, `⚠️ تم حفظ المشروع: "${text}"\n(لم نجد تطابق في قاعدة البيانات)\n\nهل تريد تعديل حقل آخر؟`, {
                    parse_mode: 'Markdown',
                    reply_markup: JSON.stringify(AI_CONFIG.AI_KEYBOARDS.EDIT_FIELDS)
                });
            }
            return;

        default:
            session.transaction[field] = text;
    }

    // ✅ حفظ الجلسة بعد التعديل
    saveAIUserSession(chatId, session);

    // العودة لعرض قائمة التعديل
    sendAIMessage(chatId, '✅ تم التحديث!\n\nهل تريد تعديل حقل آخر؟', {
        parse_mode: 'Markdown',
        reply_markup: JSON.stringify(AI_CONFIG.AI_KEYBOARDS.EDIT_FIELDS)
    });
}


// ==================== اختيار المشروع والطرف ====================

/**
 * معالجة اختيار المشروع
 */
function handleAIProjectSelection(chatId, text, session) {
    // البحث عن المشروع
    const context = loadAIContext();
    const match = matchProject(text, context.projects);

    Logger.log('🔍 Project search for: "' + text + '"');
    Logger.log('🔍 Match result: ' + JSON.stringify(match));

    if (match.found && match.score >= 0.9) {
        // ⭐ تطابق عالي - نقبله مباشرة
        Logger.log('✅ High score match: ' + match.match);
        session.transaction.project = match.match;
        session.transaction.project_code = match.code || '';

        // إذا كنا في وضع اختيار المشروع الاختياري
        if (session.validation && session.validation.needsProjectSelection) {
            session.validation.needsProjectSelection = false;
            saveAIUserSession(chatId, session);
            sendAIMessage(chatId, `✅ تم اختيار المشروع: ${match.match}`);
            continueValidation(chatId, session);
        } else {
            saveAIUserSession(chatId, session);
            moveToNextMissingField(chatId, session);
        }
    } else if (match.found && match.score >= 0.5) {
        // ⭐ تطابق متوسط - نسأل للتأكيد مع بدائل
        Logger.log('🔄 Medium score match, showing suggestions');

        const keyboard = {
            inline_keyboard: [
                [{ text: `✅ نعم: ${match.match}`, callback_data: `ai_project_${match.match}` }]
            ]
        };

        // إضافة البدائل
        if (match.alternatives && match.alternatives.length > 0) {
            match.alternatives.slice(0, 3).forEach(alt => {
                const altName = typeof alt === 'object' ? alt.name : alt;
                keyboard.inline_keyboard.push([
                    { text: `🎬 ${altName}`, callback_data: `ai_project_${altName}` }
                ]);
            });
        }

        keyboard.inline_keyboard.push([{ text: '⏭️ تخطي - بدون مشروع', callback_data: 'ai_skip_project' }]);
        keyboard.inline_keyboard.push([{ text: '❌ إلغاء', callback_data: 'ai_cancel' }]);

        sendAIMessage(chatId, `🔍 هل تقصد *"${match.match}"*؟\n\nأو اختر من القائمة:`, {
            parse_mode: 'Markdown',
            reply_markup: JSON.stringify(keyboard)
        });
    } else {
        // ⭐ لم نجد تطابق - نعرض أقرب المشاريع
        Logger.log('❌ No match found, showing closest projects');

        // البحث عن أقرب المشاريع بناءً على الكلمات المشتركة
        const suggestions = findClosestProjects(text, context.projects, 5);

        if (suggestions.length > 0) {
            const keyboard = {
                inline_keyboard: []
            };

            suggestions.forEach(proj => {
                const name = typeof proj === 'object' ? proj.name : proj;
                keyboard.inline_keyboard.push([
                    { text: `🎬 ${name}`, callback_data: `ai_project_${name}` }
                ]);
            });

            keyboard.inline_keyboard.push([{ text: '⏭️ تخطي - بدون مشروع', callback_data: 'ai_skip_project' }]);
            keyboard.inline_keyboard.push([{ text: '❌ إلغاء', callback_data: 'ai_cancel' }]);

            sendAIMessage(chatId, `🔍 لم أجد *"${text}"* بالضبط.\n\nهل تقصد أحد هذه المشاريع؟`, {
                parse_mode: 'Markdown',
                reply_markup: JSON.stringify(keyboard)
            });
        } else {
            // لا توجد اقتراحات
            sendAIMessage(chatId, '❌ لم أجد هذا المشروع. يمكنك:\n• كتابة جزء من اسم المشروع\n• الضغط على تخطي للمتابعة بدون مشروع', {
                parse_mode: 'Markdown',
                reply_markup: JSON.stringify({
                    inline_keyboard: [
                        [{ text: '⏭️ تخطي - بدون مشروع', callback_data: 'ai_skip_project' }],
                        [{ text: '❌ إلغاء', callback_data: 'ai_cancel' }]
                    ]
                })
            });
        }
    }
}

/**
 * ⭐ البحث عن أقرب المشاريع بناءً على الكلمات المشتركة
 */
function findClosestProjects(searchText, projectsList, limit) {
    if (!searchText || !projectsList || projectsList.length === 0) {
        return [];
    }

    const normalizedSearch = normalizeArabicText(searchText).toLowerCase();
    const searchWords = normalizedSearch.split(/\s+/);

    const scored = projectsList.map(proj => {
        const name = typeof proj === 'object' ? proj.name : proj;
        const normalizedName = normalizeArabicText(name).toLowerCase();

        let score = 0;

        // التحقق من الاحتواء الكامل
        if (normalizedName.includes(normalizedSearch)) {
            score += 10;
        }

        // التحقق من الكلمات المشتركة
        searchWords.forEach(word => {
            if (word.length >= 2 && normalizedName.includes(word)) {
                score += 5;
            }
        });

        // التحقق من بداية الكلمات
        const nameWords = normalizedName.split(/\s+/);
        searchWords.forEach(searchWord => {
            nameWords.forEach(nameWord => {
                if (nameWord.startsWith(searchWord) || searchWord.startsWith(nameWord)) {
                    score += 3;
                }
            });
        });

        return { project: proj, score };
    });

    // ترتيب وإرجاع الأعلى درجة
    return scored
        .filter(s => s.score > 0)
        .sort((a, b) => b.score - a.score)
        .slice(0, limit || 5)
        .map(s => s.project);
}

/**
 * معالجة callback المشروع
 */
function handleProjectCallback(chatId, project, session) {
    session.transaction.project = project;

    // ⭐ جلب كود المشروع من قاعدة البيانات
    const context = loadAIContext();
    const projectData = context.projects.find(p => {
        const name = typeof p === 'object' ? p.name : p;
        return name === project;
    });

    if (projectData && typeof projectData === 'object' && projectData.code) {
        session.transaction.project_code = projectData.code;
        Logger.log('✅ Project code found: ' + projectData.code);
    } else {
        session.transaction.project_code = '';
        Logger.log('⚠️ No project code found for: ' + project);
    }

    // ⭐ إذا كنا في وضع اختيار المشروع الاختياري
    if (session.validation && session.validation.needsProjectSelection) {
        session.validation.needsProjectSelection = false;
        saveAIUserSession(chatId, session);
        sendAIMessage(chatId, `✅ تم اختيار: ${project}` + (session.transaction.project_code ? ` (${session.transaction.project_code})` : ''));
        continueValidation(chatId, session);
    } else {
        saveAIUserSession(chatId, session);
        moveToNextMissingField(chatId, session);
    }
}

/**
 * ⭐ تخطي اختيار المشروع (بدون مشروع)
 */
function handleSkipProject(chatId, session) {
    Logger.log('📋 handleSkipProject called');

    // لا نحدد مشروع
    session.transaction.project = null;
    session.transaction.project_code = null;

    // إلغاء الحاجة لاختيار المشروع
    if (session.validation) {
        session.validation.needsProjectSelection = false;
    }

    saveAIUserSession(chatId, session);

    sendAIMessage(chatId, '✅ تم - بدون مشروع');

    // الانتقال للخطوة التالية
    continueValidation(chatId, session);
}

/**
 * معالجة اختيار الطرف
 */
function handleAIPartySelection(chatId, text, session) {
    const context = loadAIContext();
    const match = matchParty(text, context.parties);

    if (match.found && match.score > 0.9) {
        // طرف موجود
        session.transaction.party = match.match.name;
        session.transaction.partyType = match.match.type;
        session.transaction.isNewParty = false;
        session.validation.needsPartyConfirmation = false;
        saveAIUserSession(chatId, session);

        // إذا كنا في مرحلة تأكيد طرف جديد، انتقل لتأكيد الحركة
        if (session.newPartyName) {
            delete session.newPartyName;
            delete session.newPartyType;
            session.state = AI_CONFIG.AI_CONVERSATION_STATES.CONFIRM_WAIT;
            saveAIUserSession(chatId, session);
            showTransactionConfirmation(chatId, session);
        } else {
            moveToNextMissingField(chatId, session);
        }
    } else {
        // طرف جديد - تحديث الاسم وطلب التأكيد
        session.transaction.party = text;
        session.newPartyName = text;
        session.transaction.isNewParty = true;
        saveAIUserSession(chatId, session);

        // طلب تأكيد إضافة الطرف الجديد
        askNewPartyConfirmation(chatId, session);
    }
}

/**
 * معالجة callback الطرف
 */
function handlePartyCallback(chatId, party, session) {
    const context = loadAIContext();
    const partyData = context.parties.find(p => p.name === party);

    session.transaction.party = party;
    session.transaction.partyType = partyData ? partyData.type : 'مورد';
    session.transaction.isNewParty = false;

    moveToNextMissingField(chatId, session);
}

/**
 * عرض اختيار نوع الطرف الجديد
 */
function showNewPartyTypeSelection(chatId, session) {
    sendAIMessage(chatId, `👤 الطرف "${session.transaction.party}" جديد.\n\nاختر نوع الطرف:`, {
        parse_mode: 'Markdown',
        reply_markup: JSON.stringify(AI_CONFIG.AI_KEYBOARDS.NEW_PARTY_TYPE)
    });
}

/**
 * معالجة اختيار نوع الطرف الجديد
 */
function handleAINewPartyType(chatId, partyType, session) {
    session.transaction.partyType = partyType;
    session.transaction.isNewParty = true;
    moveToNextMissingField(chatId, session);
}

/**
 * الانتقال للحقل الناقص التالي أو عرض التأكيد
 */
function moveToNextMissingField(chatId, session) {
    session.currentMissingIndex++;

    saveAIUserSession(chatId, session);

    if (session.missingFields && session.currentMissingIndex < session.missingFields.length) {
        askForMissingField(chatId, session);
    } else {
        session.state = AI_CONFIG.AI_CONVERSATION_STATES.CONFIRM_WAIT;
        showTransactionConfirmation(chatId, session);
    }
}


// ==================== بناء لوحات المفاتيح ====================

/**
 * بناء لوحة المشاريع مع خيار التخطي
 */
function buildProjectsKeyboard(includeSkip = false) {
    const context = loadAIContext();
    const projects = context.projects.slice(0, 10); // أول 10 مشاريع

    const keyboard = {
        inline_keyboard: []
    };

    // صفين في كل سطر
    for (let i = 0; i < projects.length; i += 2) {
        const row = [];
        // التعامل مع المشاريع ككائنات {code, name} أو نصوص
        const p1 = projects[i];
        const name1 = typeof p1 === 'object' ? p1.name : p1;
        row.push({ text: name1, callback_data: `ai_project_${name1}` });

        if (projects[i + 1]) {
            const p2 = projects[i + 1];
            const name2 = typeof p2 === 'object' ? p2.name : p2;
            row.push({ text: name2, callback_data: `ai_project_${name2}` });
        }
        keyboard.inline_keyboard.push(row);
    }

    // ⭐ زر التخطي (بدون مشروع)
    if (includeSkip) {
        keyboard.inline_keyboard.push([{ text: '⏭️ تخطي - بدون مشروع', callback_data: 'ai_skip_project' }]);
    }

    keyboard.inline_keyboard.push([{ text: '❌ إلغاء', callback_data: 'ai_cancel' }]);

    return keyboard;
}

/**
 * بناء لوحة اقتراحات المشاريع
 */
function buildProjectSuggestionsKeyboard(mainMatch, alternatives) {
    const keyboard = {
        inline_keyboard: [
            [{ text: `✅ ${mainMatch}`, callback_data: `ai_project_${mainMatch}` }]
        ]
    };

    alternatives.forEach(alt => {
        keyboard.inline_keyboard.push([
            { text: alt, callback_data: `ai_project_${alt}` }
        ]);
    });

    keyboard.inline_keyboard.push([{ text: '❌ إلغاء', callback_data: 'ai_cancel' }]);

    return keyboard;
}


// ==================== إدارة الجلسة (Persistent Session) ====================

/**
 * الحصول على جلسة المستخدم
 */
function getAIUserSession(chatId) {
    const cache = CacheService.getScriptCache();
    const key = `AI_SESSION_${chatId}`;
    const cachedData = cache.get(key);

    if (cachedData) {
        return JSON.parse(cachedData);
    }

    // جلسة جديدة افتراضية
    return {
        state: AI_CONFIG.AI_CONVERSATION_STATES.IDLE,
        transaction: null,
        validation: null,
        missingFields: [],
        currentMissingIndex: 0,
        originalText: ''
    };
}

/**
 * حفظ جلسة المستخدم
 */
function saveAIUserSession(chatId, session) {
    const cache = CacheService.getScriptCache();
    const key = `AI_SESSION_${chatId}`;
    // حفظ لمدة 6 ساعات (21600 ثانية)
    cache.put(key, JSON.stringify(session), 21600);
}

/**
 * إعادة تعيين الجلسة
 */
function resetAIUserSession(chatId) {
    const cache = CacheService.getScriptCache();
    const key = `AI_SESSION_${chatId}`;
    cache.remove(key);
}


// ==================== التحقق من الصلاحيات ====================

/**
 * التحقق من صلاحيات المستخدم للبوت الذكي
 * يستخدم الشيت الموحد مع نظام Checkboxes
 */
function checkAIUserPermission(chatId, user) {
    try {
        // استخدام الدالة الموحدة من BotSheets.js
        const username = user.username || '';
        const result = checkUserAuthorization(null, chatId, username, 'ai_bot');

        if (result.authorized) {
            return {
                authorized: true,
                userName: result.name,
                permissions: result.permissions
            };
        }

        return { authorized: false, reason: 'المستخدم غير مصرح' };

    } catch (error) {
        Logger.log('Permission Check Error: ' + error.message);
        return { authorized: false, reason: error.message };
    }
}


// ==================== إرسال الرسائل ====================

/**
 * إرسال رسالة عبر البوت الذكي
 */
function sendAIMessage(chatId, text, options = {}) {
    try {
        const token = getAIBotToken();
        const url = `https://api.telegram.org/bot${token}/sendMessage`;

        const payload = {
            chat_id: chatId,
            text: text,
            parse_mode: options.parse_mode || 'Markdown'
        };

        if (options.reply_markup) {
            payload.reply_markup = options.reply_markup;
        }

        const response = UrlFetchApp.fetch(url, {
            method: 'post',
            contentType: 'application/json',
            payload: JSON.stringify(payload),
            muteHttpExceptions: true
        });

        return JSON.parse(response.getContentText());

    } catch (error) {
        Logger.log('Send Message Error: ' + error.message);

        // محاولة إعادة الإرسال بدون تنسيق (Plain Text) في حال فشل الـ Markdown
        if (options.parse_mode && error.message.includes('Bad Request')) {
            try {
                Logger.log('Retrying with plain text...');
                const payload = {
                    chat_id: chatId,
                    text: text
                };
                if (options.reply_markup) {
                    payload.reply_markup = options.reply_markup;
                }
                const response = UrlFetchApp.fetch(url, {
                    method: 'post',
                    contentType: 'application/json',
                    payload: JSON.stringify(payload),
                    muteHttpExceptions: true
                });
                return JSON.parse(response.getContentText());
            } catch (retryError) {
                Logger.log('Retry Failed: ' + retryError.message);
            }
        }

        return null;
    }
}

/**
 * تعديل رسالة موجودة
 */
function editAIMessage(chatId, messageId, text, options = {}) {
    try {
        const token = getAIBotToken();
        const url = `https://api.telegram.org/bot${token}/editMessageText`;

        const payload = {
            chat_id: chatId,
            message_id: messageId,
            text: text,
            parse_mode: options.parse_mode || 'Markdown'
        };

        if (options.reply_markup) {
            payload.reply_markup = options.reply_markup;
        }

        UrlFetchApp.fetch(url, {
            method: 'post',
            contentType: 'application/json',
            payload: JSON.stringify(payload),
            muteHttpExceptions: true
        });

    } catch (error) {
        Logger.log('Edit Message Error: ' + error.message);
    }
}


// ==================== دوال مساعدة ====================

/**
 * عرض حالة حركات المستخدم
 */
function showUserTransactionStatus(chatId, user) {
    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const sheet = ss.getSheetByName(CONFIG.SHEETS.BOT_TRANSACTIONS);

        if (!sheet) {
            sendAIMessage(chatId, '❌ لا توجد حركات');
            return;
        }

        const data = sheet.getDataRange().getValues();
        const userTransactions = data.filter((row, index) =>
            index > 0 && row[27] && row[27].toString() === chatId.toString()
        );

        if (userTransactions.length === 0) {
            sendAIMessage(chatId, '📭 لا توجد حركات مسجلة لك');
            return;
        }

        // آخر 5 حركات
        const recent = userTransactions.slice(-5).reverse();

        let message = '📊 *آخر حركاتك:*\n━━━━━━━━━━━━━━━━\n\n';

        recent.forEach((row, index) => {
            const status = row[25];
            const statusEmoji = status === 'معتمد' ? '✅' : status === 'مرفوض' ? '❌' : '⏳';

            message += `${index + 1}. ${statusEmoji} ${row[2]}\n`;
            message += `   💰 ${row[9]} ${row[10]}\n`;
            message += `   👤 ${row[8]}\n`;
            message += `   📅 ${row[1]}\n\n`;
        });

        sendAIMessage(chatId, message, { parse_mode: 'Markdown' });

    } catch (error) {
        Logger.log('Status Error: ' + error.message);
        sendAIMessage(chatId, '❌ حدث خطأ في جلب الحالة');
    }
}

/**
 * إرسال إشعار للمراجعين
 */
function notifyReviewers(transactionId, transaction) {
    // يمكن إضافة منطق إرسال إشعارات للمراجعين هنا
    Logger.log(`New AI Transaction: ${transactionId}`);
}

/**
 * تحويل التاريخ العربي
 */
function parseArabicDate(dateStr) {
    try {
        // ⭐ تحويل الأرقام العربية للإنجليزية أولاً
        const arabicNumerals = '٠١٢٣٤٥٦٧٨٩';
        const englishNumerals = '0123456789';
        let convertedStr = dateStr;
        for (let i = 0; i < arabicNumerals.length; i++) {
            convertedStr = convertedStr.replace(new RegExp(arabicNumerals[i], 'g'), englishNumerals[i]);
        }

        // محاولة تحويل صيغ مختلفة
        const formats = [
            /(\d{1,2})\/(\d{1,2})\/(\d{4})/,  // dd/mm/yyyy
            /(\d{1,2})-(\d{1,2})-(\d{4})/,     // dd-mm-yyyy
            /(\d{4})\/(\d{1,2})\/(\d{1,2})/,   // yyyy/mm/dd
            /(\d{4})-(\d{1,2})-(\d{1,2})/      // yyyy-mm-dd
        ];

        for (const format of formats) {
            const match = convertedStr.match(format);
            if (match) {
                if (match[1].length === 4) {
                    // yyyy-mm-dd
                    return `${match[1]}-${match[2].padStart(2, '0')}-${match[3].padStart(2, '0')}`;
                } else {
                    // dd/mm/yyyy
                    return `${match[3]}-${match[2].padStart(2, '0')}-${match[1].padStart(2, '0')}`;
                }
            }
        }

        return convertedStr;
    } catch (error) {
        return Utilities.formatDate(new Date(), 'Asia/Istanbul', 'yyyy-MM-dd');
    }
}


// ==================== إعداد البوت الذكي ====================

/**
 * التحقق من وجود شيت المستخدمين الموحد أو إنشاؤه
 * ملاحظة: تم توحيد شيت المستخدمين - يُستخدم CONFIG.SHEETS.BOT_USERS
 */
function setupAIBotUsersSheet() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(CONFIG.SHEETS.BOT_USERS);

    if (sheet) {
        Logger.log('شيت المستخدمين الموحد موجود بالفعل');
        return sheet;
    }

    // إنشاء الشيت باستخدام الدالة من BotSheets.js
    Logger.log('جاري إنشاء شيت المستخدمين الموحد...');
    return createBotUsersSheet();
}

/**
 * إعداد Trigger للبوت الذكي
 */
function setupAIBotTrigger() {
    // حذف أي triggers قديمة للبوت الذكي
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(trigger => {
        if (trigger.getHandlerFunction() === 'processAIBotUpdates') {
            ScriptApp.deleteTrigger(trigger);
        }
    });

    // إنشاء trigger جديد (كل دقيقة)
    ScriptApp.newTrigger('processAIBotUpdates')
        .timeBased()
        .everyMinutes(1)
        .create();

    Logger.log('تم إعداد Trigger للبوت الذكي - سيعمل كل دقيقة');
}

/**
 * إيقاف البوت الذكي
 */
function stopAIBot() {
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(trigger => {
        if (trigger.getHandlerFunction() === 'processAIBotUpdates') {
            ScriptApp.deleteTrigger(trigger);
        }
    });
    Logger.log('تم إيقاف البوت الذكي');
}

/**
 * حذف Webhook إذا كان موجوداً - مطلوب لعمل Long Polling
 * شغّل هذه الدالة مرة واحدة إذا كان البوت لا يستجيب
 */
function deleteAIBotWebhook() {
    const token = PropertiesService.getScriptProperties().getProperty('AI_BOT_TOKEN');
    if (!token) {
        Logger.log('❌ لم يتم تعيين AI_BOT_TOKEN');
        return;
    }

    // أولاً: فحص الـ Webhook الحالي
    const infoUrl = `https://api.telegram.org/bot${token}/getWebhookInfo`;
    const infoResponse = UrlFetchApp.fetch(infoUrl);
    const infoData = JSON.parse(infoResponse.getContentText());

    Logger.log('📋 معلومات Webhook الحالية:');
    Logger.log(JSON.stringify(infoData, null, 2));

    if (infoData.result && infoData.result.url && infoData.result.url !== '') {
        Logger.log('⚠️ يوجد Webhook مُفعّل: ' + infoData.result.url);
        Logger.log('🗑️ جاري حذف الـ Webhook...');

        // حذف الـ Webhook
        const deleteUrl = `https://api.telegram.org/bot${token}/deleteWebhook`;
        const deleteResponse = UrlFetchApp.fetch(deleteUrl);
        const deleteData = JSON.parse(deleteResponse.getContentText());

        if (deleteData.ok) {
            Logger.log('✅ تم حذف الـ Webhook بنجاح!');
            Logger.log('🔄 الآن البوت سيعمل بـ Long Polling');
        } else {
            Logger.log('❌ فشل حذف الـ Webhook: ' + JSON.stringify(deleteData));
        }
    } else {
        Logger.log('✅ لا يوجد Webhook مُفعّل - البوت يعمل بـ Long Polling');
    }

    return infoData;
}

// ==================== دوال الأوردر المشترك (Shared Order) ====================

/**
 * ⭐ معالجة الأوردر المشترك
 * @param {number} chatId - معرّف المحادثة
 * @param {Object} transaction - بيانات الأوردر من الذكاء الاصطناعي
 * @param {Object} user - بيانات المستخدم
 */
function handleSharedOrder(chatId, transaction, user) {
    Logger.log('═══════════════════════════════════════');
    Logger.log('📦 handleSharedOrder STARTED');
    Logger.log('📦 Projects: ' + JSON.stringify(transaction.projects));

    // حفظ الأوردر في الجلسة
    const session = getAIUserSession(chatId);
    session.sharedOrder = transaction;
    session.user = user;
    session.state = AI_CONFIG.AI_CONVERSATION_STATES.WAITING_SHARED_ORDER_CONFIRMATION;
    saveAIUserSession(chatId, session);

    // عرض ملخص الأوردر المشترك
    showSharedOrderConfirmation(chatId, transaction);
}

/**
 * ⭐ عرض ملخص الأوردر المشترك للتأكيد (نسخة ذكية محسنة)
 */
function showSharedOrderConfirmation(chatId, order) {
    // ⭐ تحميل قاعدة بيانات المشاريع للبحث الذكي
    const context = loadAIContext();
    const projectsList = context.projects || [];

    const projects = order.projects || [];
    const items = order.items || [{ item: order.item, amount: order.total_amount || order.amount }];
    const totalGuests = order.total_guests || projects.reduce((sum, p) => sum + (p.guests || 0), 0);
    const totalAppearances = order.total_appearances || projects.length; // ⭐ عدد الظهورات = عدد المشاريع
    const currency = order.currency || 'USD';

    // ⭐ حساب التوزيع بناءً على الظهورات وليس الضيوف
    let distributionText = '';
    projects.forEach((project, index) => {
        const prefix = index === projects.length - 1 ? '└─' : '├─';
        const guestNames = project.guest_names ? project.guest_names.join('، ') : '';
        const guests = project.guests || 1;

        // ⭐ البحث عن المشروع في قاعدة البيانات
        let projectName = project.name;
        let projectCode = project.code || '';
        const projectMatch = matchProject(project.name, projectsList);
        if (projectMatch.found) {
            projectName = projectMatch.match;
            projectCode = projectMatch.code || '';
            // تحديث البيانات في order.projects للحفظ لاحقاً
            project.name = projectName;
            project.code = projectCode;
        }

        // حساب المبلغ لكل بند
        let amountsText = '';
        items.forEach((itemObj, i) => {
            const itemAmount = itemObj.amount || 0;
            const projectAmount = Math.round((itemAmount / totalAppearances) * 100) / 100;
            if (i > 0) amountsText += ' + ';
            amountsText += `${projectAmount.toLocaleString()} ${currency}`;
        });

        let projectDisplay = projectName;
        if (projectCode) projectDisplay += ` (${projectCode})`;
        distributionText += `${prefix} *${projectDisplay}*: ${guestNames || guests + ' ضيف'}`;
        distributionText += ` → ${amountsText}\n`;
    });

    // ⭐ عرض البنود المتعددة
    let itemsText = '';
    let grandTotal = 0;
    if (items.length > 1) {
        items.forEach(itemObj => {
            itemsText += `   • ${itemObj.item}: ${(itemObj.amount || 0).toLocaleString()} ${currency}\n`;
            grandTotal += itemObj.amount || 0;
        });
    } else {
        grandTotal = items[0].amount || 0;
    }

    // بناء رسالة الملخص
    let summary = `
📦 *أوردر مشترك*
━━━━━━━━━━━━━━━━
📌 *النوع:* ${order.nature || 'استحقاق مصروف'}
📁 *التصنيف:* ${order.classification || '-'}`;

    if (items.length > 1) {
        summary += `\n📂 *البنود:*\n${itemsText}`;
    } else {
        summary += `\n📂 *البند:* ${items[0].item || '-'}`;
    }

    summary += `
👤 *الطرف:* ${order.party || '-'}
💰 *المبلغ الإجمالي:* ${grandTotal.toLocaleString()} ${currency}
📊 *عدد الوحدات:* ${order.unit_count || totalAppearances}
💳 *طريقة الدفع:* ${order.payment_method || 'تحويل بنكي'}
📅 *شرط الدفع:* ${order.payment_term || 'فوري'}`;

    // عرض تاريخ الاستحقاق إذا كان تاريخ مخصص
    if (order.payment_term === 'تاريخ مخصص' && order.payment_term_date) {
        summary += `\n📆 *تاريخ الاستحقاق:* ${order.payment_term_date}`;
    }

    summary += `

🎬 *التوزيع على المشاريع (${totalGuests} ضيف):*
${distributionText}
━━━━━━━━━━━━━━━━
`;

    // أزرار التأكيد
    const keyboard = {
        inline_keyboard: [
            [
                { text: '✅ تأكيد وحفظ', callback_data: 'shared_confirm' },
                { text: '❌ إلغاء', callback_data: 'shared_cancel' }
            ]
        ]
    };

    sendAIMessage(chatId, summary, {
        parse_mode: 'Markdown',
        reply_markup: JSON.stringify(keyboard)
    });
}

/**
 * ⭐ حفظ الأوردر المشترك (عدة حركات) - يستخدم addTransactionDirectly مثل الحركات العادية
 */
function saveSharedOrderFromAI(chatId, session) {
    Logger.log('═══════════════════════════════════════');
    Logger.log('💾 saveSharedOrderFromAI STARTED');

    const order = session.sharedOrder;
    const user = session.user;
    const userName = `${user.first_name || ''} ${user.last_name || ''}`.trim();

    if (!order || !order.projects || order.projects.length === 0) {
        sendAIMessage(chatId, '❌ خطأ: بيانات الأوردر غير مكتملة', { parse_mode: 'Markdown' });
        return { success: false };
    }

    try {
        // ⭐ تحميل قاعدة بيانات المشاريع للبحث الذكي
        const context = loadAIContext();
        const projectsList = context.projects || [];
        Logger.log('📦 Loaded ' + projectsList.length + ' projects from database');

        const projects = order.projects;
        const items = order.items || [{ item: order.item, amount: order.total_amount || order.amount }];
        const totalAppearances = order.total_appearances || projects.length;
        const totalGuests = order.total_guests || projects.reduce((sum, p) => sum + (p.guests || 0), 0);

        // تاريخ الحركة
        const transactionDate = order.due_date && order.due_date !== 'TODAY'
            ? order.due_date
            : new Date();

        // إنشاء رقم الأوردر المشترك
        const sharedOrderId = 'SO-' + Utilities.formatDate(new Date(), 'Asia/Istanbul', 'yyyyMMdd-HHmmss');

        const savedTransactions = [];
        let transactionCounter = 0;

        // ⭐ حفظ حركة لكل مشروع ولكل بند باستخدام addTransactionDirectly
        for (const itemObj of items) {
            const itemName = itemObj.item || order.item || '';
            const itemTotalAmount = itemObj.amount || 0;
            const amountPerProject = Math.round((itemTotalAmount / totalAppearances) * 100) / 100;

            for (const project of projects) {
                transactionCounter++;
                const guestNames = project.guest_names ? project.guest_names.join('، ') : '';
                const guests = project.guests || 1;

                // ⭐ البحث عن المشروع في قاعدة البيانات
                let projectName = project.name;
                let projectCode = project.code || '';

                const projectMatch = matchProject(project.name, projectsList);
                if (projectMatch.found) {
                    projectName = projectMatch.match;
                    projectCode = projectMatch.code || '';
                    Logger.log(`✅ Project matched: "${project.name}" → "${projectName}" (${projectCode})`);
                } else {
                    Logger.log(`⚠️ Project not found in DB: "${project.name}" - using as-is`);
                }

                // ⭐ تفاصيل تتضمن أسماء الضيوف ورقم الأوردر
                let details = `أوردر مشترك [${sharedOrderId}]`;
                if (guestNames) {
                    details += `: ${guestNames}`;
                } else {
                    details += `: ${guests}/${totalGuests} ضيف`;
                }
                if (order.details) {
                    details += ` - ${order.details}`;
                }

                // ⭐ تجهيز البيانات بنفس صيغة الحركات العادية
                const transactionData = {
                    date: transactionDate,
                    nature: order.nature || 'استحقاق مصروف',
                    classification: order.classification || '',
                    projectCode: projectCode,
                    projectName: projectName,
                    item: itemName,
                    details: details,
                    partyName: order.party || '',
                    amount: amountPerProject,
                    currency: order.currency || 'USD',
                    exchangeRate: order.exchange_rate || 1,
                    paymentMethod: order.payment_method || 'تحويل بنكي',
                    paymentTermType: order.payment_term || 'فوري',
                    weeks: order.payment_term_weeks || '',
                    customDate: order.payment_term_date || '',
                    telegramUser: userName,
                    chatId: chatId,
                    unitCount: guests,
                    notes: `أوردر مشترك: ${sharedOrderId}`
                };

                // ⭐ استخدام addTransactionDirectly للحفظ في دفتر الحركات مباشرة
                const result = addTransactionDirectly(transactionData, '🤖 بوت ذكي - أوردر مشترك');

                if (result.success) {
                    savedTransactions.push({
                        id: result.transactionId,
                        project: projectName,
                        code: projectCode,
                        item: itemName,
                        amount: amountPerProject,
                        guests: guestNames || guests
                    });
                    Logger.log(`✅ Saved: ${projectName} (${projectCode}) - ${itemName} - ${amountPerProject} - Row: ${result.rowNumber}`);
                } else {
                    Logger.log(`❌ Failed to save: ${projectName} - ${result.error}`);
                }
            }
        }

        // إرسال رسالة النجاح
        let successMessage = `✅ *تم حفظ الأوردر المشترك بنجاح!*\n\n`;
        successMessage += `📦 *رقم الأوردر:* \`${sharedOrderId}\`\n`;
        successMessage += `📊 *عدد الحركات:* ${savedTransactions.length}\n\n`;
        successMessage += `*التفاصيل:*\n`;

        savedTransactions.forEach(t => {
            successMessage += `• ${t.project}`;
            if (t.code) successMessage += ` (${t.code})`;
            if (items.length > 1) successMessage += ` - ${t.item}`;
            successMessage += `: ${t.amount.toLocaleString()} ${order.currency || 'USD'}\n`;
        });

        successMessage += `\n📋 *تم الحفظ في دفتر الحركات مباشرة.*`;

        sendAIMessage(chatId, successMessage, { parse_mode: 'Markdown' });

        // مسح الجلسة
        resetAIUserSession(chatId);

        return { success: true, orderId: sharedOrderId, count: savedTransactions.length };

    } catch (error) {
        Logger.log('❌ Error saving shared order: ' + error.message);
        Logger.log(error.stack);
        sendAIMessage(chatId, '❌ حدث خطأ أثناء حفظ الأوردر: ' + error.message, { parse_mode: 'Markdown' });
        return { success: false, error: error.message };
    }
}

/**
 * ⭐ معالجة ضغط أزرار الأوردر المشترك
 */
function handleSharedOrderCallback(chatId, messageId, data, session) {
    Logger.log('📦 handleSharedOrderCallback: ' + data);

    if (data === 'shared_confirm') {
        // تأكيد وحفظ
        editAIMessage(chatId, messageId, '⏳ *جاري حفظ الأوردر المشترك...*');
        saveSharedOrderFromAI(chatId, session);
    } else if (data === 'shared_cancel') {
        // إلغاء
        resetAIUserSession(chatId);
        editAIMessage(chatId, messageId, '❌ تم إلغاء الأوردر المشترك.');
    }
}

/**
 * إعداد كامل للبوت الذكي
 */
function setupAIBot() {
    Logger.log('=== بدء إعداد البوت الذكي ===');

    // 1. التحقق من المفاتيح
    const setup = checkAIBotSetup();
    if (!setup.ready) {
        Logger.log('❌ يرجى إعداد المفاتيح أولاً باستخدام setupAIBotCredentials(botToken, geminiKey)');
        return false;
    }

    // 2. إنشاء شيت المستخدمين
    setupAIBotUsersSheet();

    // 3. إعداد Trigger
    setupAIBotTrigger();

    Logger.log('=== تم إعداد البوت الذكي بنجاح! ===');
    return true;
}
