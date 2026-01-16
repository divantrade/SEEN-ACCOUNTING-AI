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

    // التحقق من الصلاحيات
    const permission = checkAIUserPermission(chatId, user);
    if (!permission.authorized) {
        sendAIMessage(chatId, AI_CONFIG.AI_MESSAGES.UNAUTHORIZED);
        return;
    }

    // معالجة الأوامر
    if (text && text.startsWith('/')) {
        handleAICommand(chatId, text, user);
        return;
    }

    // الحصول على جلسة المستخدم
    const session = getAIUserSession(chatId);

    // معالجة حسب حالة المحادثة
    switch (session.state) {
        case AI_CONFIG.AI_CONVERSATION_STATES.WAITING_MISSING_FIELD:
            handleMissingFieldInput(chatId, text, session);
            break;

        case AI_CONFIG.AI_CONVERSATION_STATES.WAITING_PROJECT_SELECTION:
            handleProjectSelection(chatId, text, session);
            break;

        case AI_CONFIG.AI_CONVERSATION_STATES.WAITING_PARTY_SELECTION:
            handlePartySelection(chatId, text, session);
            break;

        case AI_CONFIG.AI_CONVERSATION_STATES.WAITING_EDIT:
            handleEditInput(chatId, text, session);
            break;

        case AI_CONFIG.AI_CONVERSATION_STATES.WAITING_EXCHANGE_RATE:
            // ⭐ معالجة إدخال سعر الصرف
            handleExchangeRateInput(chatId, text, session);
            break;

        case AI_CONFIG.AI_CONVERSATION_STATES.WAITING_PAYMENT_TERM:
            // ⭐ معالجة إدخال شرط الدفع (أسابيع أو تاريخ)
            handlePaymentTermInput(chatId, text, session);
            break;

        default:
            // تحليل النص كحركة مالية جديدة
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
    // إرسال رسالة "جاري التحليل"
    const loadingMsg = sendAIMessage(chatId, AI_CONFIG.AI_MESSAGES.ANALYZING, { parse_mode: 'Markdown' });

    try {
        // تحليل النص
        const result = analyzeTransaction(text);

        if (!result.success) {
            sendAIMessage(chatId, result.error || AI_CONFIG.AI_MESSAGES.ERROR_PARSE, { parse_mode: 'Markdown' });
            return;
        }

        // حفظ الحركة في الجلسة
        const session = getAIUserSession(chatId);
        session.transaction = result.transaction;
        session.validation = result.validation;
        session.originalText = text;

        // حفظ التغييرات في الكاش
        saveAIUserSession(chatId, session);

        // التحقق من الحقول الناقصة
        if (result.needsInput && result.missingFields.length > 0) {
            handleMissingFields(chatId, result.missingFields, session);
            return;
        }

        // التحقق من طرف جديد يحتاج تأكيد
        if (result.validation && result.validation.needsPartyConfirmation) {
            askNewPartyConfirmation(chatId, session);
            return;
        }

        // ⭐ التحقق من طريقة الدفع
        if (result.validation && result.validation.needsPaymentMethod) {
            askPaymentMethod(chatId, session);
            return;
        }

        // ⭐ التحقق من العملة
        if (result.validation && result.validation.needsCurrency) {
            askCurrency(chatId, session);
            return;
        }

        // ⭐ التحقق من سعر الصرف (إذا العملة غير دولار)
        if (result.validation && result.validation.needsExchangeRate) {
            askExchangeRate(chatId, session);
            return;
        }

        // ⭐ التحقق من شرط الدفع (للاستحقاقات فقط)
        if (result.validation && result.validation.needsPaymentTerm) {
            askPaymentTerm(chatId, session);
            return;
        }

        // ⭐ التحقق من عدد الأسابيع (لشرط بعد التسليم)
        if (result.validation && result.validation.needsPaymentTermWeeks) {
            askPaymentTermWeeks(chatId, session);
            return;
        }

        // ⭐ التحقق من تاريخ الدفع المخصص
        if (result.validation && result.validation.needsPaymentTermDate) {
            askPaymentTermDate(chatId, session);
            return;
        }

        // عرض ملخص للتأكيد
        showTransactionConfirmation(chatId, session);

    } catch (error) {
        Logger.log('Process Transaction Error: ' + error.message);
        Logger.log('Stack: ' + error.stack);
        sendAIMessage(chatId, `❌ *حدث خطأ غير متوقع:*\n${error.message}\n\nيرجى إعادة المحاولة أو التواصل مع الدعم التقني.`);
    }
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
function handlePaymentMethodSelection(chatId, method, session) {
    session.transaction.payment_method = method;
    session.validation.enriched.payment_method = method;
    session.validation.needsPaymentMethod = false;
    saveAIUserSession(chatId, session);

    sendAIMessage(chatId, `✅ تم تحديد طريقة الدفع: *${method}*`, { parse_mode: 'Markdown' });

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
function handleCurrencySelection(chatId, currency, session) {
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
    session.state = AI_CONFIG.AI_CONVERSATION_STATES.WAITING_EXCHANGE_RATE;
    saveAIUserSession(chatId, session);

    const currency = session.transaction.currency || session.validation.enriched.currency;
    const currencyNames = { 'TRY': 'الليرة التركية', 'EGP': 'الجنيه المصري' };
    const currencyName = currencyNames[currency] || currency;

    const message = AI_CONFIG.AI_MESSAGES.ASK_EXCHANGE_RATE.replace('{currency}', currencyName);
    sendAIMessage(chatId, message, { parse_mode: 'Markdown' });
}

/**
 * ⭐ معالجة إدخال سعر الصرف
 */
function handleExchangeRateInput(chatId, text, session) {
    // استخراج الرقم من النص
    const rate = parseFloat(text.replace(/[^0-9.]/g, ''));

    if (isNaN(rate) || rate <= 0) {
        sendAIMessage(chatId, '❌ سعر الصرف غير صحيح. يرجى إدخال رقم صحيح (مثال: 32.5):');
        return;
    }

    session.transaction.exchangeRate = rate;
    session.transaction.exchange_rate = rate;
    session.validation.enriched.exchangeRate = rate;
    session.validation.needsExchangeRate = false;
    saveAIUserSession(chatId, session);

    sendAIMessage(chatId, `✅ تم تحديد سعر الصرف: *${rate}*`, { parse_mode: 'Markdown' });

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
function handlePaymentTermSelection(chatId, term, session) {
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
 * ⭐ معالجة إدخال بيانات شرط الدفع (أسابيع أو تاريخ)
 */
function handlePaymentTermInput(chatId, text, session) {
    if (session.waitingFor === 'weeks') {
        // إدخال عدد الأسابيع
        const weeks = parseInt(text.replace(/[^0-9]/g, ''));
        if (isNaN(weeks) || weeks <= 0) {
            sendAIMessage(chatId, '❌ يرجى إدخال رقم صحيح (مثال: 2):');
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
        const parsedDate = parseArabicDate(text);
        if (!parsedDate) {
            sendAIMessage(chatId, '❌ تنسيق التاريخ غير صحيح. يرجى الإدخال بصيغة: 15/2/2026');
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

    // الرد على الـ callback
    answerAICallback(callbackQuery.id);

    // التحقق من الصلاحيات
    const permission = checkAIUserPermission(chatId, user);
    if (!permission.authorized) {
        return;
    }

    const session = getAIUserSession(chatId);

    // ⭐ فحص وجود الجلسة والـ validation
    if (!session || !session.validation) {
        Logger.log('⚠️ Session or validation missing for callback: ' + data);
        sendAIMessage(chatId, '⚠️ انتهت الجلسة. يرجى إعادة إرسال الحركة.');
        resetAIUserSession(chatId);
        return;
    }

    // معالجة حسب نوع الـ callback
    if (data.startsWith('ai_confirm')) {
        handleAIConfirmation(chatId, session, user);
    } else if (data.startsWith('ai_edit')) {
        handleEditRequest(chatId, data, session, messageId);
    } else if (data.startsWith('ai_cancel')) {
        sendAIMessage(chatId, AI_CONFIG.AI_MESSAGES.CANCELLED);
        resetAIUserSession(chatId);
    } else if (data.startsWith('ai_project_')) {
        const project = data.replace('ai_project_', '');
        handleProjectCallback(chatId, project, session);
    } else if (data.startsWith('ai_select_party_')) {
        // ⭐ معالجة اختيار طرف من القائمة
        const index = parseInt(data.replace('ai_select_party_', ''));
        handleSelectPartyFromSuggestions(chatId, index, session);
    } else if (data.startsWith('ai_payment_')) {
        // ⭐ معالجة اختيار طريقة الدفع
        const method = data.replace('ai_payment_', '');
        handlePaymentMethodSelection(chatId, method, session);
    } else if (data.startsWith('ai_currency_')) {
        // ⭐ معالجة اختيار العملة
        const currency = data.replace('ai_currency_', '');
        handleCurrencySelection(chatId, currency, session);
    } else if (data.startsWith('ai_term_')) {
        // ⭐ معالجة اختيار شرط الدفع
        const term = data.replace('ai_term_', '');
        handlePaymentTermSelection(chatId, term, session);
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
        handleNewPartyType(chatId, partyType, session);
    } else if (data === 'ai_add_party') {
        showNewPartyTypeSelection(chatId, session);
    } else if (data === 'ai_edit_done') {
        showTransactionConfirmation(chatId, session);
    }
}

/**
 * الرد على الـ callback query
 */
function answerAICallback(callbackQueryId) {
    const token = getAIBotToken();
    const url = `https://api.telegram.org/bot${token}/answerCallbackQuery`;

    UrlFetchApp.fetch(url, {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify({ callback_query_id: callbackQueryId }),
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
 * حفظ الحركة في شيت حركات البوت
 */
function saveAITransaction(transaction, user, chatId) {
    Logger.log('🚀 بدء عملية حفظ الحركة...');
    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const sheetName = CONFIG.SHEETS.BOT_TRANSACTIONS;
        Logger.log(`📂 البحث عن الشيت: "${sheetName}"`);

        let sheet = ss.getSheetByName(sheetName);

        // إذا كان الشيت غير موجود، نقوم بإنشائه (اختياري أو إرسال خطأ واضح)
        if (!sheet) {
            Logger.log(`⚠️ الشيت "${sheetName}" غير موجود. جاري إنشاؤه...`);
            try {
                sheet = ss.insertSheet(sheetName);
                // إعداد الهيدر إذا لزم الأمر
                const headers = Object.values(BOT_CONFIG.BOT_TRANSACTIONS_COLUMNS).map(col => col.name);
                sheet.appendRow(headers);
                sheet.getRange(1, 1, 1, headers.length).setBackground('#f3f3f3').setFontWeight('bold');
                Logger.log('✅ تم إنشاء الشيت والهيدر بنجاح.');
            } catch (e) {
                Logger.log(`❌ فشل إنشاء الشيت: ${e.message}`);
                throw new Error(`شيت "${sheetName}" غير موجود وفشل إنشاؤه تلقائياً. تأكد من الصلاحيات.`);
            }
        } else {
            Logger.log('✅ الشيت موجود.');
        }

        // إنشاء رقم الحركة
        const transactionId = generateTransactionId();
        Logger.log(`🆔 رقم الحركة الجديد: ${transactionId}`);

        const now = new Date();
        const timestamp = Utilities.formatDate(now, 'Asia/Istanbul', 'yyyy-MM-dd HH:mm:ss');
        const month = Utilities.formatDate(now, 'Asia/Istanbul', 'yyyy-MM');

        // حساب القيمة بالدولار
        let amountUSD = 0;
        try {
            amountUSD = calculateUSDAmount(
                transaction.amount,
                transaction.currency,
                transaction.exchangeRate
            );
        } catch (e) {
            Logger.log('⚠️ خطأ في حساب الدولار: ' + e.message);
            amountUSD = transaction.amount; // Fallback
        }

        // تحديد نوع الحركة
        const movementType = inferMovementType(transaction.nature);

        // بناء صف البيانات (مطابق لهيكل BOT_TRANSACTIONS_COLUMNS)
        const rowData = [
            transactionId,                                          // رقم الحركة
            transaction.due_date && transaction.due_date !== 'TODAY' ? transaction.due_date : timestamp.split(' ')[0], // التاريخ
            transaction.nature,                                     // طبيعة الحركة
            transaction.classification,                             // تصنيف الحركة
            transaction.project_code || '',                         // كود المشروع
            transaction.project || '',                              // اسم المشروع
            transaction.item || '',                                 // البند
            transaction.details || '',                              // التفاصيل
            transaction.party,                                      // اسم المورد/الجهة
            transaction.amount,                                     // المبلغ بالعملة الأصلية
            transaction.currency,                                   // العملة
            transaction.exchangeRate || 1,                          // سعر الصرف
            amountUSD,                                              // القيمة بالدولار
            movementType,                                           // نوع الحركة
            '',                                                     // الرصيد
            '',                                                     // رقم مرجعي
            transaction.payment_method || 'تحويل بنكي',            // طريقة الدفع
            transaction.payment_term || 'فوري',                     // نوع شرط الدفع
            transaction.payment_term_weeks || '',                   // عدد الأسابيع
            transaction.payment_term_date || '',                    // تاريخ مخصص
            transaction.due_date && transaction.due_date !== 'TODAY' ? transaction.due_date : timestamp.split(' ')[0], // تاريخ الاستحقاق
            'معلق',                                                 // حالة السداد
            month,                                                  // الشهر
            transaction.originalText || '',                         // ملاحظات (النص الأصلي)
            '',                                                     // كشف
            CONFIG.TELEGRAM_BOT.REVIEW_STATUS.PENDING,             // حالة المراجعة
            `${user.first_name || ''} ${user.last_name || ''}`.trim(), // المُدخل
            chatId,                                                 // معرّف المحادثة
            timestamp,                                              // تاريخ الإدخال
            '',                                                     // المُراجع
            '',                                                     // تاريخ المراجعة
            '',                                                     // ملاحظات المراجعة
            '',                                                     // رابط المرفق
            transaction.isNewParty ? 'نعم' : 'لا',                 // طرف جديد؟
            'بوت ذكي'                                               // مصدر الإدخال
        ];

        Logger.log('📝 تجهيز البيانات للحفظ: ' + JSON.stringify(rowData));

        // إضافة الصف
        try {
            sheet.appendRow(rowData);
            Logger.log('✅ تم إضافة الصف بنجاح!');
        } catch (appendError) {
            Logger.log('❌ خطأ أثناء appendRow: ' + appendError.message);
            throw new Error('فشل في الكتابة في الشيت: ' + appendError.message);
        }

        // إذا كان طرف جديد، أضفه لشيت أطراف البوت
        if (transaction.isNewParty) {
            try {
                Logger.log('👤 إضافة طرف جديد...');
                addNewPartyFromAI(transaction, user, chatId);
                Logger.log('✅ تم إضافة الطرف الجديد.');
            } catch (e) {
                Logger.log('⚠️ تحذير: فشل إضافة الطرف الجديد: ' + e.message);
                // لا نوقف العملية إذا فشل إضافة الطرف فقط
            }
        }

        return {
            success: true,
            transactionId: transactionId
        };

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

    const fieldMessages = {
        'nature': '📤 اختر طبيعة الحركة الجديدة:',
        'classification': '📊 اختر التصنيف الجديد:',
        'project': '🎬 اكتب اسم المشروع:',
        'item': '📁 اكتب اسم البند:',
        'party': '👤 اكتب اسم الطرف:',
        'amount': '💰 اكتب المبلغ الجديد:',
        'currency': '💱 اختر العملة:',
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
            const amount = parseFloat(text.replace(/[^0-9.]/g, ''));
            if (isNaN(amount)) {
                sendAIMessage(chatId, '❌ المبلغ غير صحيح. اكتب رقماً صحيحاً:');
                return;
            }
            session.transaction.amount = amount;
            break;

        case 'date':
        case 'due_date':
            session.transaction.due_date = parseArabicDate(text);
            break;

        default:
            session.transaction[field] = text;
    }

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
function handleProjectSelection(chatId, text, session) {
    // البحث عن المشروع
    const context = loadAIContext();
    const match = matchProject(text, context.projects);

    if (match.found && match.score > 0.7) {
        session.transaction.project = match.match;
        moveToNextMissingField(chatId, session);
    } else if (match.found && match.alternatives) {
        // عرض اقتراحات
        const keyboard = buildProjectSuggestionsKeyboard(match.match, match.alternatives);
        sendAIMessage(chatId, `🎬 هل تقصد "${match.match}"؟`, {
            parse_mode: 'Markdown',
            reply_markup: JSON.stringify(keyboard)
        });
    } else {
        sendAIMessage(chatId, '❌ لم أجد هذا المشروع. حاول مرة أخرى أو اكتب اسماً مختلفاً:');
    }
}

/**
 * معالجة callback المشروع
 */
function handleProjectCallback(chatId, project, session) {
    session.transaction.project = project;
    moveToNextMissingField(chatId, session);
}

/**
 * معالجة اختيار الطرف
 */
function handlePartySelection(chatId, text, session) {
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
function handleNewPartyType(chatId, partyType, session) {
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
 * بناء لوحة المشاريع
 */
function buildProjectsKeyboard() {
    const context = loadAIContext();
    const projects = context.projects.slice(0, 10); // أول 10 مشاريع

    const keyboard = {
        inline_keyboard: []
    };

    // صفين في كل سطر
    for (let i = 0; i < projects.length; i += 2) {
        const row = [];
        row.push({ text: projects[i], callback_data: `ai_project_${projects[i]}` });
        if (projects[i + 1]) {
            row.push({ text: projects[i + 1], callback_data: `ai_project_${projects[i + 1]}` });
        }
        keyboard.inline_keyboard.push(row);
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
        // محاولة تحويل صيغ مختلفة
        const formats = [
            /(\d{1,2})\/(\d{1,2})\/(\d{4})/,  // dd/mm/yyyy
            /(\d{1,2})-(\d{1,2})-(\d{4})/,     // dd-mm-yyyy
            /(\d{4})\/(\d{1,2})\/(\d{1,2})/,   // yyyy/mm/dd
            /(\d{4})-(\d{1,2})-(\d{1,2})/      // yyyy-mm-dd
        ];

        for (const format of formats) {
            const match = dateStr.match(format);
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

        return dateStr;
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
