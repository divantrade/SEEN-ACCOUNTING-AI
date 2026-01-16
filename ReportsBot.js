/**
 * ═══════════════════════════════════════════════════════════════════════════
 *                          REPORTS BOT
 *                  معالجة أوامر التقارير في البوت
 * ═══════════════════════════════════════════════════════════════════════════
 */

/**
 * ═══════════════════════════════════════════════════════════════════════════
 *                         معالجة الأوامر
 * ═══════════════════════════════════════════════════════════════════════════
 */

/**
 * معالجة أمر /reports أو /تقارير
 * @param {string} chatId - معرف المحادثة
 */
function handleReportsCommand(chatId) {
    try {
        Logger.log('📊 Reports command received from chatId: ' + chatId);

        // إرسال القائمة الرئيسية للتقارير
        sendAIMessage(chatId, REPORTS_CONFIG.MESSAGES.MAIN_MENU, {
            parse_mode: 'Markdown',
            reply_markup: JSON.stringify(REPORTS_CONFIG.KEYBOARDS.MAIN_MENU)
        });

        // تحديث حالة الجلسة
        const session = getAIUserSession(chatId);
        session.reportState = REPORTS_CONFIG.STATES.IDLE;
        session.reportData = {};
        saveAIUserSession(chatId, session);

    } catch (error) {
        Logger.log('❌ Error in handleReportsCommand: ' + error.message);
        sendAIMessage(chatId, REPORTS_CONFIG.MESSAGES.ERROR);
    }
}

/**
 * ═══════════════════════════════════════════════════════════════════════════
 *                      معالجة Callbacks التقارير
 * ═══════════════════════════════════════════════════════════════════════════
 */

/**
 * معالجة callback التقارير
 * @param {string} chatId - معرف المحادثة
 * @param {string} data - بيانات الـ callback
 * @param {Object} session - جلسة المستخدم
 */
function handleReportCallback(chatId, data, session) {
    try {
        Logger.log('📊 Report callback: ' + data);

        // التأكد من وجود reportData
        if (!session.reportData) {
            session.reportData = {};
        }

        // ═══════════════════════════════════════════════════════════════
        //                     القائمة الرئيسية
        // ═══════════════════════════════════════════════════════════════

        if (data === 'report_statement') {
            // كشف حساب - اختيار نوع الطرف
            session.reportState = REPORTS_CONFIG.STATES.WAITING_PARTY_TYPE;
            session.reportData.type = 'statement';
            saveAIUserSession(chatId, session);

            sendAIMessage(chatId, REPORTS_CONFIG.MESSAGES.SELECT_PARTY_TYPE, {
                parse_mode: 'Markdown',
                reply_markup: JSON.stringify(REPORTS_CONFIG.KEYBOARDS.PARTY_TYPE)
            });

        } else if (data === 'report_alerts') {
            // تنبيهات الاستحقاق - اختيار الفترة
            session.reportState = REPORTS_CONFIG.STATES.WAITING_ALERTS_PERIOD;
            session.reportData.type = 'alerts';
            saveAIUserSession(chatId, session);

            sendAIMessage(chatId, '⏰ *اختر فترة التنبيهات:*', {
                parse_mode: 'Markdown',
                reply_markup: JSON.stringify(REPORTS_CONFIG.KEYBOARDS.ALERTS_PERIOD)
            });

        } else if (data === 'report_balances') {
            // تقرير الأرصدة - مباشر
            handleGenerateBalancesReport(chatId, session);

        } else if (data === 'report_profitability') {
            // ربحية المشاريع - مباشر
            handleGenerateProfitabilityReport(chatId, session);

        } else if (data === 'report_expenses') {
            // تقرير المصروفات - مباشر
            handleGenerateExpensesReport(chatId, session);

        } else if (data === 'report_revenues') {
            // تقرير الإيرادات - مباشر
            handleGenerateRevenuesReport(chatId, session);

        }

        // ═══════════════════════════════════════════════════════════════
        //                     اختيار نوع الطرف
        // ═══════════════════════════════════════════════════════════════

        else if (data.startsWith('report_party_')) {
            const partyType = data.replace('report_party_', '');
            session.reportData.partyType = partyType;
            session.reportState = REPORTS_CONFIG.STATES.WAITING_PARTY_NAME;
            saveAIUserSession(chatId, session);

            sendAIMessage(chatId, REPORTS_CONFIG.MESSAGES.ENTER_PARTY_NAME, {
                parse_mode: 'Markdown'
            });

        }

        // ═══════════════════════════════════════════════════════════════
        //                     اختيار الطرف من القائمة
        // ═══════════════════════════════════════════════════════════════

        else if (data.startsWith('report_select_party_')) {
            const index = parseInt(data.replace('report_select_party_', ''));
            const parties = session.reportData.foundParties || [];

            if (index >= 0 && index < parties.length) {
                const selectedParty = parties[index];
                session.reportData.partyName = selectedParty.name;
                session.reportData.partyCode = selectedParty.code;

                // توليد كشف الحساب
                handleGenerateStatementReport(chatId, session);
            } else {
                sendAIMessage(chatId, '❌ اختيار غير صالح');
            }

        }

        // ═══════════════════════════════════════════════════════════════
        //                     فترة التنبيهات
        // ═══════════════════════════════════════════════════════════════

        else if (data.startsWith('report_alerts_')) {
            const days = parseInt(data.replace('report_alerts_', ''));
            session.reportData.alertsDays = days;
            handleGenerateAlertsReport(chatId, session, days);

        }

        // ═══════════════════════════════════════════════════════════════
        //                     التنقل والإلغاء
        // ═══════════════════════════════════════════════════════════════

        else if (data === 'report_back_main') {
            // الرجوع للقائمة الرئيسية
            handleReportsCommand(chatId);

        } else if (data === 'report_back_party') {
            // الرجوع لاختيار نوع الطرف
            session.reportState = REPORTS_CONFIG.STATES.WAITING_PARTY_TYPE;
            saveAIUserSession(chatId, session);

            sendAIMessage(chatId, REPORTS_CONFIG.MESSAGES.SELECT_PARTY_TYPE, {
                parse_mode: 'Markdown',
                reply_markup: JSON.stringify(REPORTS_CONFIG.KEYBOARDS.PARTY_TYPE)
            });

        } else if (data === 'report_cancel') {
            // إلغاء
            session.reportState = null;
            session.reportData = {};
            saveAIUserSession(chatId, session);

            sendAIMessage(chatId, '❌ تم إلغاء طلب التقرير.\n\nيمكنك إرسال /reports لطلب تقرير جديد.');

        }

    } catch (error) {
        Logger.log('❌ Error in handleReportCallback: ' + error.message);
        sendAIMessage(chatId, REPORTS_CONFIG.MESSAGES.ERROR);
    }
}

/**
 * ═══════════════════════════════════════════════════════════════════════════
 *                     معالجة إدخال اسم الطرف
 * ═══════════════════════════════════════════════════════════════════════════
 */

/**
 * معالجة النص المُدخل للبحث عن الطرف
 * @param {string} chatId - معرف المحادثة
 * @param {string} text - النص المُدخل
 * @param {Object} session - جلسة المستخدم
 */
function handleReportPartySearch(chatId, text, session) {
    try {
        Logger.log('🔍 Searching for party: ' + text);

        const partyType = session.reportData.partyType;
        const parties = searchPartiesByName(text, partyType);

        if (parties.length === 0) {
            sendAIMessage(chatId, REPORTS_CONFIG.MESSAGES.NO_PARTIES_FOUND);
            return;
        }

        if (parties.length === 1) {
            // طرف واحد فقط - استخدامه مباشرة
            session.reportData.partyName = parties[0].name;
            session.reportData.partyCode = parties[0].code;
            handleGenerateStatementReport(chatId, session);
            return;
        }

        // عدة أطراف - عرض قائمة للاختيار
        session.reportData.foundParties = parties;
        session.reportState = REPORTS_CONFIG.STATES.WAITING_PARTY_SELECTION;
        saveAIUserSession(chatId, session);

        // بناء لوحة مفاتيح الاختيار
        const keyboard = buildPartySelectionKeyboard(parties);

        sendAIMessage(chatId, REPORTS_CONFIG.MESSAGES.SELECT_PARTY, {
            parse_mode: 'Markdown',
            reply_markup: JSON.stringify(keyboard)
        });

    } catch (error) {
        Logger.log('❌ Error in handleReportPartySearch: ' + error.message);
        sendAIMessage(chatId, REPORTS_CONFIG.MESSAGES.ERROR);
    }
}

/**
 * البحث عن الأطراف بالاسم
 * @param {string} searchText - نص البحث
 * @param {string} partyType - نوع الطرف
 * @returns {Array} - قائمة الأطراف المطابقة
 */
function searchPartiesByName(searchText, partyType) {
    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const results = [];
        const addedNames = new Set();

        // ⭐ 1. البحث في شيت الأطراف الرئيسي
        const mainSheet = ss.getSheetByName(CONFIG.SHEETS.PARTIES);
        if (mainSheet) {
            const data = mainSheet.getDataRange().getValues();
            for (let i = 1; i < data.length; i++) {
                const row = data[i];
                const name = String(row[0] || '').trim(); // العمود A - الاسم
                const type = String(row[1] || '').trim(); // العمود B - النوع

                // التحقق من النوع والاسم
                if (type === partyType && name && name.toLowerCase().includes(searchText.toLowerCase())) {
                    if (!addedNames.has(name.toLowerCase())) {
                        results.push({
                            name: name,
                            type: type,
                            code: '' // لا يوجد كود في الشيت الرئيسي
                        });
                        addedNames.add(name.toLowerCase());
                    }
                }
            }
        }

        // ⭐ 2. البحث في شيت أطراف البوت
        const botSheet = ss.getSheetByName(CONFIG.SHEETS.BOT_PARTIES);
        if (botSheet) {
            const data = botSheet.getDataRange().getValues();
            for (let i = 1; i < data.length; i++) {
                const row = data[i];
                const name = String(row[0] || '').trim(); // العمود A - الاسم
                const type = String(row[1] || '').trim(); // العمود B - النوع

                if (type === partyType && name && name.toLowerCase().includes(searchText.toLowerCase())) {
                    if (!addedNames.has(name.toLowerCase())) {
                        results.push({
                            name: name,
                            type: type,
                            code: ''
                        });
                        addedNames.add(name.toLowerCase());
                    }
                }
            }
        }

        Logger.log('🔍 Found ' + results.length + ' parties matching "' + searchText + '" of type ' + partyType);

        // ترتيب النتائج حسب التطابق
        results.sort((a, b) => {
            const aExact = a.name.toLowerCase() === searchText.toLowerCase();
            const bExact = b.name.toLowerCase() === searchText.toLowerCase();
            if (aExact && !bExact) return -1;
            if (!aExact && bExact) return 1;
            return a.name.localeCompare(b.name, 'ar');
        });

        return results.slice(0, 10); // أقصى 10 نتائج

    } catch (error) {
        Logger.log('❌ Error searching parties: ' + error.message);
        return [];
    }
}

/**
 * بناء لوحة مفاتيح اختيار الطرف
 */
function buildPartySelectionKeyboard(parties) {
    const keyboard = { inline_keyboard: [] };

    parties.forEach((party, index) => {
        keyboard.inline_keyboard.push([{
            text: party.name + (party.code ? ' (' + party.code + ')' : ''),
            callback_data: 'report_select_party_' + index
        }]);
    });

    // زر الرجوع والإلغاء
    keyboard.inline_keyboard.push([
        { text: '🔙 رجوع', callback_data: 'report_back_party' },
        { text: '❌ إلغاء', callback_data: 'report_cancel' }
    ]);

    return keyboard;
}

/**
 * ═══════════════════════════════════════════════════════════════════════════
 *                        توليد التقارير
 * ═══════════════════════════════════════════════════════════════════════════
 */

/**
 * توليد كشف الحساب
 */
function handleGenerateStatementReport(chatId, session) {
    try {
        const partyName = session.reportData.partyName;
        const partyType = session.reportData.partyType;

        // إرسال رسالة الانتظار
        sendAIMessage(chatId, REPORTS_CONFIG.MESSAGES.GENERATING_REPORT, {
            parse_mode: 'Markdown'
        });

        // توليد وإرسال التقرير
        const result = generateStatementPDF(chatId, partyName, partyType);

        if (result.success) {
            // إعادة تعيين الجلسة
            session.reportState = null;
            session.reportData = {};
            saveAIUserSession(chatId, session);
        } else {
            sendAIMessage(chatId, REPORTS_CONFIG.MESSAGES.ERROR + '\n\n' + (result.error || ''));
        }

    } catch (error) {
        Logger.log('❌ Error generating statement: ' + error.message);
        sendAIMessage(chatId, REPORTS_CONFIG.MESSAGES.ERROR);
    }
}

/**
 * توليد تقرير التنبيهات
 */
function handleGenerateAlertsReport(chatId, session, daysAhead) {
    try {
        // إرسال رسالة الانتظار
        sendAIMessage(chatId, REPORTS_CONFIG.MESSAGES.GENERATING_REPORT, {
            parse_mode: 'Markdown'
        });

        // أولاً: إرسال ملخص نصي سريع
        sendAlertsTextSummary(chatId, daysAhead);

        // ثانياً: توليد وإرسال PDF
        const result = generateAlertsPDF(chatId, daysAhead);

        // إعادة تعيين الجلسة
        session.reportState = null;
        session.reportData = {};
        saveAIUserSession(chatId, session);

    } catch (error) {
        Logger.log('❌ Error generating alerts: ' + error.message);
        sendAIMessage(chatId, REPORTS_CONFIG.MESSAGES.ERROR);
    }
}

/**
 * إرسال ملخص نصي للتنبيهات
 */
function sendAlertsTextSummary(chatId, daysAhead) {
    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const alertsSheet = ss.getSheetByName(CONFIG.SHEETS.ALERTS);

        if (!alertsSheet) {
            return;
        }

        const data = alertsSheet.getDataRange().getValues();
        const today = new Date();
        const targetDate = new Date(today.getTime() + (daysAhead * 24 * 60 * 60 * 1000));

        let urgentAlerts = [];
        let importantAlerts = [];

        // تحليل التنبيهات
        for (let i = 1; i < data.length; i++) {
            const row = data[i];
            const priority = row[1]; // الأولوية
            const party = row[4];    // الطرف
            const amount = row[5];   // المبلغ
            const dueDate = row[7];  // تاريخ الاستحقاق

            if (priority === '🔴 عاجل') {
                urgentAlerts.push({ party, amount, dueDate });
            } else if (priority === '🟠 مهم') {
                importantAlerts.push({ party, amount, dueDate });
            }
        }

        // بناء الرسالة
        let message = REPORTS_CONFIG.MESSAGES.ALERTS_HEADER;

        if (urgentAlerts.length === 0 && importantAlerts.length === 0) {
            message += REPORTS_CONFIG.MESSAGES.NO_ALERTS;
        } else {
            if (urgentAlerts.length > 0) {
                message += '🔴 *عاجل (' + urgentAlerts.length + '):*\n';
                urgentAlerts.slice(0, 5).forEach(alert => {
                    message += `• ${alert.party}: ${alert.amount} USD\n`;
                });
                message += '\n';
            }

            if (importantAlerts.length > 0) {
                message += '🟠 *مهم (' + importantAlerts.length + '):*\n';
                importantAlerts.slice(0, 5).forEach(alert => {
                    message += `• ${alert.party}: ${alert.amount} USD\n`;
                });
            }

            message += '\n_التفاصيل الكاملة في الملف المرفق_';
        }

        sendAIMessage(chatId, message, { parse_mode: 'Markdown' });

    } catch (error) {
        Logger.log('⚠️ Error sending alerts summary: ' + error.message);
    }
}

/**
 * توليد تقرير الأرصدة
 */
function handleGenerateBalancesReport(chatId, session) {
    try {
        sendAIMessage(chatId, REPORTS_CONFIG.MESSAGES.GENERATING_REPORT, {
            parse_mode: 'Markdown'
        });

        const result = generateBalancesPDF(chatId);

        session.reportState = null;
        session.reportData = {};
        saveAIUserSession(chatId, session);

        if (!result.success) {
            sendAIMessage(chatId, REPORTS_CONFIG.MESSAGES.ERROR);
        }

    } catch (error) {
        Logger.log('❌ Error generating balances: ' + error.message);
        sendAIMessage(chatId, REPORTS_CONFIG.MESSAGES.ERROR);
    }
}

/**
 * توليد تقرير الربحية
 */
function handleGenerateProfitabilityReport(chatId, session) {
    try {
        sendAIMessage(chatId, REPORTS_CONFIG.MESSAGES.GENERATING_REPORT, {
            parse_mode: 'Markdown'
        });

        const result = generateProfitabilityPDF(chatId);

        session.reportState = null;
        session.reportData = {};
        saveAIUserSession(chatId, session);

        if (!result.success) {
            sendAIMessage(chatId, REPORTS_CONFIG.MESSAGES.ERROR);
        }

    } catch (error) {
        Logger.log('❌ Error generating profitability: ' + error.message);
        sendAIMessage(chatId, REPORTS_CONFIG.MESSAGES.ERROR);
    }
}

/**
 * توليد تقرير المصروفات
 */
function handleGenerateExpensesReport(chatId, session) {
    try {
        sendAIMessage(chatId, REPORTS_CONFIG.MESSAGES.GENERATING_REPORT, {
            parse_mode: 'Markdown'
        });

        const result = generateExpensesPDF(chatId);

        session.reportState = null;
        session.reportData = {};
        saveAIUserSession(chatId, session);

        if (!result.success) {
            sendAIMessage(chatId, REPORTS_CONFIG.MESSAGES.ERROR);
        }

    } catch (error) {
        Logger.log('❌ Error generating expenses: ' + error.message);
        sendAIMessage(chatId, REPORTS_CONFIG.MESSAGES.ERROR);
    }
}

/**
 * توليد تقرير الإيرادات
 */
function handleGenerateRevenuesReport(chatId, session) {
    try {
        sendAIMessage(chatId, REPORTS_CONFIG.MESSAGES.GENERATING_REPORT, {
            parse_mode: 'Markdown'
        });

        const result = generateRevenuesPDF(chatId);

        session.reportState = null;
        session.reportData = {};
        saveAIUserSession(chatId, session);

        if (!result.success) {
            sendAIMessage(chatId, REPORTS_CONFIG.MESSAGES.ERROR);
        }

    } catch (error) {
        Logger.log('❌ Error generating revenues: ' + error.message);
        sendAIMessage(chatId, REPORTS_CONFIG.MESSAGES.ERROR);
    }
}

/**
 * ═══════════════════════════════════════════════════════════════════════════
 *                    التحقق من حالة التقارير
 * ═══════════════════════════════════════════════════════════════════════════
 */

/**
 * التحقق هل المستخدم في وضع طلب تقرير
 * @param {Object} session - جلسة المستخدم
 * @returns {boolean}
 */
function isInReportMode(session) {
    return session.reportState &&
        session.reportState !== REPORTS_CONFIG.STATES.IDLE;
}

/**
 * التحقق هل الـ callback خاص بالتقارير
 * @param {string} data - بيانات الـ callback
 * @returns {boolean}
 */
function isReportCallback(data) {
    return data && data.startsWith('report_');
}

/**
 * ═══════════════════════════════════════════════════════════════════════════
 *                    إعداد قائمة أوامر البوت
 * ═══════════════════════════════════════════════════════════════════════════
 */

/**
 * إعداد قائمة الأوامر في تليجرام (تظهر بجوار حقل الكتابة)
 * يجب تشغيل هذه الدالة مرة واحدة فقط لإعداد القائمة
 */
function setupBotCommands() {
    try {
        const token = CONFIG.TELEGRAM_BOT.AI_BOT_TOKEN;
        const url = 'https://api.telegram.org/bot' + token + '/setMyCommands';

        const commands = [
            { command: 'start', description: '🚀 بدء استخدام البوت' },
            { command: 'help', description: '❓ المساعدة والدليل' },
            { command: 'reports', description: '📊 التقارير وكشوف الحساب' },
            { command: 'status', description: '📋 حالة حركاتي' },
            { command: 'cancel', description: '❌ إلغاء العملية الحالية' }
        ];

        const response = UrlFetchApp.fetch(url, {
            method: 'post',
            contentType: 'application/json',
            payload: JSON.stringify({ commands: commands }),
            muteHttpExceptions: true
        });

        const result = JSON.parse(response.getContentText());

        if (result.ok) {
            Logger.log('✅ Bot commands menu set up successfully');
            return { success: true };
        } else {
            Logger.log('❌ Failed to set up bot commands: ' + result.description);
            return { success: false, error: result.description };
        }

    } catch (error) {
        Logger.log('❌ Error setting up bot commands: ' + error.message);
        return { success: false, error: error.message };
    }
}
