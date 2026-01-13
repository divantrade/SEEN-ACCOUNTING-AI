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
    try {
        // التحقق من إعداد البوت
        const setup = checkAIBotSetup();
        if (!setup.ready) {
            Logger.log('البوت الذكي غير جاهز - يرجى إعداد المفاتيح أولاً');
            return;
        }

        const token = getAIBotToken();
        const lastUpdateId = getAILastUpdateId();

        // جلب التحديثات
        const url = `https://api.telegram.org/bot${token}/getUpdates?offset=${lastUpdateId + 1}&timeout=50`;

        const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
        const data = JSON.parse(response.getContentText());

        if (!data.ok) {
            Logger.log('AI Bot Error: ' + JSON.stringify(data));
            return;
        }

        const updates = data.result;

        if (updates.length === 0) {
            return;
        }

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

    } catch (error) {
        Logger.log('AI Bot Main Error: ' + error.message);
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

        // التحقق من الحقول الناقصة
        if (result.needsInput && result.missingFields.length > 0) {
            handleMissingFields(chatId, result.missingFields, session);
            return;
        }

        // عرض ملخص للتأكيد
        showTransactionConfirmation(chatId, session);

    } catch (error) {
        Logger.log('Process Transaction Error: ' + error.message);
        sendAIMessage(chatId, AI_CONFIG.AI_MESSAGES.ERROR_API);
    }
}

/**
 * معالجة الحقول الناقصة
 */
function handleMissingFields(chatId, missingFields, session) {
    session.missingFields = missingFields;
    session.currentMissingIndex = 0;
    session.state = AI_CONFIG.AI_CONVERSATION_STATES.WAITING_MISSING_FIELD;

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
            break;

        case 'party':
            message = AI_CONFIG.AI_MESSAGES.ASK_PARTY;
            session.state = AI_CONFIG.AI_CONVERSATION_STATES.WAITING_PARTY_SELECTION;
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

    if (session.currentMissingIndex < session.missingFields.length) {
        askForMissingField(chatId, session);
    } else {
        // اكتملت الحقول - عرض التأكيد
        session.state = AI_CONFIG.AI_CONVERSATION_STATES.IDLE;
        showTransactionConfirmation(chatId, session);
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

    // معالجة حسب نوع الـ callback
    if (data.startsWith('ai_confirm')) {
        handleConfirmation(chatId, session, user);
    } else if (data.startsWith('ai_edit')) {
        handleEditRequest(chatId, data, session, messageId);
    } else if (data.startsWith('ai_cancel')) {
        sendAIMessage(chatId, AI_CONFIG.AI_MESSAGES.CANCELLED);
        resetAIUserSession(chatId);
    } else if (data.startsWith('ai_project_')) {
        const project = data.replace('ai_project_', '');
        handleProjectCallback(chatId, project, session);
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
function handleConfirmation(chatId, session, user) {
    try {
        // حفظ الحركة
        const result = saveAITransaction(session.transaction, user, chatId);

        if (result.success) {
            const successMsg = AI_CONFIG.AI_MESSAGES.SUCCESS.replace('#{id}', result.transactionId);
            sendAIMessage(chatId, successMsg, { parse_mode: 'Markdown' });

            // إرسال إشعار للمراجعين (اختياري)
            notifyReviewers(result.transactionId, session.transaction);
        } else {
            sendAIMessage(chatId, AI_CONFIG.AI_MESSAGES.ERROR_API);
        }

        resetAIUserSession(chatId);

    } catch (error) {
        Logger.log('Confirmation Error: ' + error.message);
        sendAIMessage(chatId, AI_CONFIG.AI_MESSAGES.ERROR_API);
    }
}

/**
 * حفظ الحركة في شيت حركات البوت
 */
function saveAITransaction(transaction, user, chatId) {
    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const sheet = ss.getSheetByName(CONFIG.SHEETS.BOT_TRANSACTIONS);

        if (!sheet) {
            throw new Error('شيت حركات البوت غير موجود');
        }

        // إنشاء رقم الحركة
        const transactionId = generateTransactionId();
        const timestamp = Utilities.formatDate(new Date(), 'Asia/Istanbul', 'yyyy-MM-dd HH:mm:ss');
        const month = Utilities.formatDate(new Date(), 'Asia/Istanbul', 'yyyy-MM');

        // حساب القيمة بالدولار
        const amountUSD = calculateUSDAmount(
            transaction.amount,
            transaction.currency,
            transaction.exchangeRate
        );

        // تحديد نوع الحركة
        const movementType = inferMovementType(transaction.nature);

        // بناء صف البيانات (مطابق لهيكل BOT_TRANSACTIONS_COLUMNS)
        const rowData = [
            transactionId,                                          // رقم الحركة
            transaction.due_date || timestamp.split(' ')[0],        // التاريخ
            transaction.nature,                                     // طبيعة الحركة
            transaction.classification,                             // تصنيف الحركة
            '',                                                     // كود المشروع
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
            'فوري',                                                 // نوع شرط الدفع
            '',                                                     // عدد الأسابيع
            '',                                                     // تاريخ مخصص
            transaction.due_date || timestamp.split(' ')[0],        // تاريخ الاستحقاق
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
            transaction.isNewParty ? 'نعم' : 'لا'                  // طرف جديد؟
        ];

        // إضافة الصف
        sheet.appendRow(rowData);

        // إذا كان طرف جديد، أضفه لشيت أطراف البوت
        if (transaction.isNewParty) {
            addNewPartyFromAI(transaction, user, chatId);
        }

        return {
            success: true,
            transactionId: transactionId
        };

    } catch (error) {
        Logger.log('Save Transaction Error: ' + error.message);
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

    if (match.found && match.score > 0.7) {
        session.transaction.party = match.match.name;
        session.transaction.partyType = match.match.type;
        session.transaction.isNewParty = false;
        moveToNextMissingField(chatId, session);
    } else {
        // طرف جديد
        session.transaction.party = text;
        session.transaction.isNewParty = true;
        showNewPartyTypeSelection(chatId, session);
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

    if (session.missingFields && session.currentMissingIndex < session.missingFields.length) {
        askForMissingField(chatId, session);
    } else {
        session.state = AI_CONFIG.AI_CONVERSATION_STATES.IDLE;
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


// ==================== إدارة الجلسات ====================

/**
 * الحصول على جلسة المستخدم
 */
function getAIUserSession(chatId) {
    if (!aiUserSessions[chatId]) {
        aiUserSessions[chatId] = {
            state: AI_CONFIG.AI_CONVERSATION_STATES.IDLE,
            transaction: {},
            missingFields: [],
            currentMissingIndex: 0
        };
    }
    return aiUserSessions[chatId];
}

/**
 * إعادة تعيين جلسة المستخدم
 */
function resetAIUserSession(chatId) {
    aiUserSessions[chatId] = {
        state: AI_CONFIG.AI_CONVERSATION_STATES.IDLE,
        transaction: {},
        missingFields: [],
        currentMissingIndex: 0
    };
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
