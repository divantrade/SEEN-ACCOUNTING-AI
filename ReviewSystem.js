// ==================== نظام مراجعة واعتماد حركات البوت ====================
/**
 * ملف نظام المراجعة والاعتماد
 * يدير عملية مراجعة الحركات والأطراف القادمة من البوت
 */

// ==================== واجهة المراجعة ====================

/**
 * عرض Sidebar للمراجعة السريعة
 */
function showBotReviewSidebar() {
    const html = HtmlService.createHtmlOutput(getBotReviewSidebarHtml())
        .setTitle('📋 مراجعة حركات البوت')
        .setWidth(400);

    SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * إنشاء HTML للـ Sidebar
 */
function getBotReviewSidebarHtml() {
    const pending = getPendingBotTransactions();
    const pendingParties = getPendingBotParties();

    let html = `
    <!DOCTYPE html>
    <html dir="rtl">
    <head>
        <base target="_top">
        <style>
            * { box-sizing: border-box; }
            body {
                font-family: 'Segoe UI', Tahoma, Arial, sans-serif;
                padding: 16px;
                background: #f5f5f5;
                margin: 0;
            }
            .header {
                background: linear-gradient(135deg, #0088cc, #005f8f);
                color: white;
                padding: 16px;
                border-radius: 8px;
                margin-bottom: 16px;
                text-align: center;
            }
            .header h2 { margin: 0 0 8px 0; }
            .stats {
                display: flex;
                justify-content: space-around;
                margin-top: 12px;
            }
            .stat {
                text-align: center;
            }
            .stat-number {
                font-size: 24px;
                font-weight: bold;
            }
            .stat-label {
                font-size: 12px;
                opacity: 0.9;
            }
            .section {
                background: white;
                border-radius: 8px;
                padding: 12px;
                margin-bottom: 12px;
                box-shadow: 0 2px 4px rgba(0,0,0,0.1);
            }
            .section-title {
                font-weight: bold;
                color: #333;
                margin-bottom: 12px;
                padding-bottom: 8px;
                border-bottom: 2px solid #0088cc;
            }
            .transaction-card {
                background: #fff8e1;
                border-right: 4px solid #ffc107;
                padding: 12px;
                margin-bottom: 10px;
                border-radius: 4px;
            }
            .transaction-card.new-party {
                background: #e3f2fd;
                border-right-color: #2196f3;
            }
            .card-header {
                display: flex;
                justify-content: space-between;
                margin-bottom: 8px;
            }
            .card-id {
                font-weight: bold;
                color: #0088cc;
            }
            .card-amount {
                font-weight: bold;
                color: #2e7d32;
            }
            .card-details {
                font-size: 13px;
                color: #666;
                line-height: 1.6;
            }
            .card-details strong {
                color: #333;
            }
            .card-actions {
                display: flex;
                gap: 8px;
                margin-top: 12px;
            }
            .btn {
                flex: 1;
                padding: 8px 12px;
                border: none;
                border-radius: 4px;
                cursor: pointer;
                font-size: 13px;
                font-weight: bold;
            }
            .btn-approve {
                background: #4caf50;
                color: white;
            }
            .btn-reject {
                background: #f44336;
                color: white;
            }
            .btn-edit {
                background: #ff9800;
                color: white;
            }
            .btn:hover {
                opacity: 0.9;
            }
            .empty-state {
                text-align: center;
                padding: 40px;
                color: #999;
            }
            .empty-state-icon {
                font-size: 48px;
                margin-bottom: 16px;
            }
            .new-party-badge {
                background: #2196f3;
                color: white;
                padding: 2px 8px;
                border-radius: 12px;
                font-size: 11px;
                margin-right: 8px;
            }
            .refresh-btn {
                width: 100%;
                padding: 12px;
                background: #0088cc;
                color: white;
                border: none;
                border-radius: 8px;
                cursor: pointer;
                font-size: 14px;
                margin-bottom: 16px;
            }
            .refresh-btn:hover {
                background: #006699;
            }
            #loading {
                display: none;
                text-align: center;
                padding: 20px;
            }
            .spinner {
                border: 3px solid #f3f3f3;
                border-top: 3px solid #0088cc;
                border-radius: 50%;
                width: 30px;
                height: 30px;
                animation: spin 1s linear infinite;
                margin: 0 auto;
            }
            @keyframes spin {
                0% { transform: rotate(0deg); }
                100% { transform: rotate(360deg); }
            }
            .toast {
                position: fixed;
                bottom: 20px;
                left: 50%;
                transform: translateX(-50%);
                background: #333;
                color: white;
                padding: 12px 24px;
                border-radius: 8px;
                z-index: 1000;
                transition: opacity 0.3s ease;
                box-shadow: 0 4px 12px rgba(0,0,0,0.3);
            }
            .btn:disabled {
                cursor: not-allowed;
            }
        </style>
    </head>
    <body>
        <div class="header">
            <h2>🤖 مراجعة حركات البوت</h2>
            <div class="stats">
                <div class="stat">
                    <div class="stat-number" id="pending-count">${pending.length}</div>
                    <div class="stat-label">حركات معلقة</div>
                </div>
                <div class="stat">
                    <div class="stat-number" id="parties-count">${pendingParties.length}</div>
                    <div class="stat-label">أطراف جديدة</div>
                </div>
            </div>
        </div>

        <button class="refresh-btn" onclick="refresh()">🔄 تحديث</button>

        <div id="loading">
            <div class="spinner"></div>
            <p>جاري التحميل...</p>
        </div>

        <div id="content">
    `;

    if (pending.length === 0 && pendingParties.length === 0) {
        html += `
            <div class="empty-state">
                <div class="empty-state-icon">✅</div>
                <p>لا توجد حركات معلقة للمراجعة</p>
                <p style="font-size: 12px; color: #666; margin-top: 8px;">
                    💡 الحركات تُسجَّل مباشرة في دفتر الحركات الآن.<br>
                    لرفض حركة: حدد الصف ➡️ القائمة ➡️ "❌ رفض الحركة المحددة"
                </p>
            </div>
        `;
    } else {
        // عرض الحركات المعلقة
        if (pending.length > 0) {
            html += `<div class="section" id="transactions-section"><div class="section-title">📋 الحركات المعلقة (<span id="pending-section-count">${pending.length}</span>)</div>`;

            pending.forEach(tx => {
                const isNewParty = tx.isNewParty;
                html += `
                    <div class="transaction-card ${isNewParty ? 'new-party' : ''}" id="card-${tx.rowNumber}">
                        <div class="card-header">
                            <span class="card-id">#${tx.transactionId}</span>
                            <span class="card-amount">${tx.amount} ${tx.currency}</span>
                        </div>
                        <div class="card-details">
                            <strong>النوع:</strong> ${tx.nature}<br>
                            <strong>المشروع:</strong> ${tx.projectName || '-'}<br>
                            <strong>الطرف:</strong> ${isNewParty ? '<span class="new-party-badge">جديد</span>' : ''} ${tx.partyName}<br>
                            <strong>التفاصيل:</strong> ${tx.details || '-'}<br>
                            <strong>المُدخل:</strong> ${tx.telegramUser}
                        </div>
                        <div class="card-actions">
                            <button class="btn btn-approve" onclick="approve(${tx.rowNumber})">✅ اعتماد</button>
                            <button class="btn btn-reject" onclick="reject(${tx.rowNumber})">❌ رفض</button>
                        </div>
                    </div>
                `;
            });

            html += `</div>`;
        }
    }

    html += `
        </div>

        <script>
            let pendingCount = ${pending.length};
            let partiesCount = ${pendingParties.length};

            function updateStats() {
                document.getElementById('pending-count').textContent = pendingCount;
                document.getElementById('parties-count').textContent = partiesCount;
                document.getElementById('pending-section-count').textContent = pendingCount;
            }

            function showLoading() {
                document.getElementById('loading').style.display = 'block';
                document.getElementById('content').style.display = 'none';
            }

            function hideLoading() {
                document.getElementById('loading').style.display = 'none';
                document.getElementById('content').style.display = 'block';
            }

            function removeCard(rowNumber) {
                const card = document.getElementById('card-' + rowNumber);
                if (card) {
                    card.style.transition = 'all 0.3s ease';
                    card.style.opacity = '0';
                    card.style.transform = 'translateX(100px)';
                    setTimeout(() => {
                        card.remove();
                        pendingCount--;
                        updateStats();

                        // إذا لم تبق حركات، أظهر رسالة الفراغ
                        if (pendingCount === 0 && partiesCount === 0) {
                            document.getElementById('transactions-section').innerHTML =
                                '<div class="empty-state"><div class="empty-state-icon">✅</div><p>تم مراجعة جميع الحركات!</p></div>';
                        }
                    }, 300);
                }
            }

            function disableCardButtons(rowNumber) {
                const card = document.getElementById('card-' + rowNumber);
                if (card) {
                    const buttons = card.querySelectorAll('.btn');
                    buttons.forEach(btn => {
                        btn.disabled = true;
                        btn.style.opacity = '0.5';
                    });
                }
            }

            function refresh() {
                showLoading();
                google.script.run
                    .withSuccessHandler(function() {
                        google.script.host.close();
                        google.script.run.showBotReviewSidebar();
                    })
                    .withFailureHandler(function(error) {
                        alert('خطأ: ' + error.message);
                        hideLoading();
                    })
                    .refreshBotReviewData();
            }

            function approve(rowNumber) {
                if (!confirm('هل تريد اعتماد هذه الحركة؟')) return;

                disableCardButtons(rowNumber);

                google.script.run
                    .withSuccessHandler(function(result) {
                        if (result.success) {
                            removeCard(rowNumber);
                            // رسالة نجاح صغيرة بدون alert
                            showToast('✅ تم الاعتماد بنجاح');
                        } else {
                            alert('❌ خطأ: ' + result.error);
                            hideLoading();
                        }
                    })
                    .withFailureHandler(function(error) {
                        alert('خطأ: ' + error.message);
                        hideLoading();
                    })
                    .approveTransaction(rowNumber);
            }

            function reject(rowNumber) {
                var reason = prompt('يرجى إدخال سبب الرفض:');
                if (!reason) return;

                disableCardButtons(rowNumber);

                google.script.run
                    .withSuccessHandler(function(result) {
                        if (result.success) {
                            removeCard(rowNumber);
                            showToast('✅ تم الرفض');
                        } else {
                            alert('❌ خطأ: ' + result.error);
                            hideLoading();
                        }
                    })
                    .withFailureHandler(function(error) {
                        alert('خطأ: ' + error.message);
                        hideLoading();
                    })
                    .rejectBotTransaction(rowNumber, reason);
            }

            function showToast(message) {
                const toast = document.createElement('div');
                toast.className = 'toast';
                toast.textContent = message;
                document.body.appendChild(toast);
                setTimeout(() => {
                    toast.style.opacity = '0';
                    setTimeout(() => toast.remove(), 300);
                }, 2000);
            }
        </script>
    </body>
    </html>
    `;

    return html;
}

/**
 * تحديث بيانات المراجعة (للـ Sidebar)
 */
function refreshBotReviewData() {
    return true;
}

// ==================== دوال الاعتماد والرفض ====================

/**
 * اعتماد حركة
 */
function approveTransaction(rowNumber) {
    try {
        Logger.log('=== بدء اعتماد الحركة ===');
        Logger.log('Row Number: ' + rowNumber);

        const botSheet = getBotTransactionsSheet();
        const mainSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.TRANSACTIONS);
        const columns = BOT_CONFIG.BOT_TRANSACTIONS_COLUMNS;

        Logger.log('Main Sheet Name: ' + CONFIG.SHEETS.TRANSACTIONS);
        Logger.log('Main Sheet Found: ' + (mainSheet ? 'YES' : 'NO'));

        if (!mainSheet) {
            Logger.log('❌ لم يتم العثور على دفتر الحركات');
            return { success: false, error: 'لم يتم العثور على دفتر الحركات' };
        }

        // قراءة بيانات الحركة
        const rowData = botSheet.getRange(rowNumber, 1, 1, Object.keys(columns).length).getValues()[0];

        // التحقق من الحالة
        const currentStatus = rowData[columns.REVIEW_STATUS.index - 1];
        if (String(currentStatus).trim() !== CONFIG.TELEGRAM_BOT.REVIEW_STATUS.PENDING) {
            console.log(`Transaction at row ${rowNumber} skipped. Status is '${currentStatus}'`);
            return { success: false, error: 'الحركة ليست في حالة انتظار (الحالة الحالية: ' + currentStatus + ')' };
        }

        // إذا كان هناك طرف جديد، أضفه لقاعدة البيانات أولاً
        const isNewParty = rowData[columns.IS_NEW_PARTY.index - 1] === 'نعم';
        if (isNewParty) {
            const partyName = rowData[columns.PARTY_NAME.index - 1];
            const addPartyResult = approveAndAddNewParty(partyName, rowData[columns.TELEGRAM_CHAT_ID.index - 1]);
            if (!addPartyResult.success) {
                return { success: false, error: 'فشل إضافة الطرف الجديد: ' + addPartyResult.error };
            }
        }

        // نسخ البيانات لدفتر الحركات الرئيسي
        const mainLastRow = mainSheet.getLastRow();
        const newRow = mainLastRow + 1;

        Logger.log('Main Sheet Last Row: ' + mainLastRow);
        Logger.log('New Row Number: ' + newRow);

        // إعداد بيانات الصف الجديد (28 عمود - مع أعمدة Y, Z, AA, AB)
        const mainRowData = [
            newRow - 1, // رقم الحركة
            rowData[columns.DATE.index - 1],
            rowData[columns.NATURE.index - 1],
            rowData[columns.CLASSIFICATION.index - 1],
            rowData[columns.PROJECT_CODE.index - 1],
            rowData[columns.PROJECT_NAME.index - 1],
            rowData[columns.ITEM.index - 1],
            rowData[columns.DETAILS.index - 1],
            rowData[columns.PARTY_NAME.index - 1],
            rowData[columns.AMOUNT.index - 1],
            rowData[columns.CURRENCY.index - 1],
            rowData[columns.EXCHANGE_RATE.index - 1],
            rowData[columns.AMOUNT_USD.index - 1],
            rowData[columns.MOVEMENT_TYPE.index - 1],
            '', // الرصيد - سيُحسب بالصيغة
            '', // رقم مرجعي - سيُحسب بالصيغة
            rowData[columns.PAYMENT_METHOD.index - 1],
            rowData[columns.PAYMENT_TERM_TYPE.index - 1],
            rowData[columns.WEEKS.index - 1],
            rowData[columns.CUSTOM_DATE.index - 1],
            '', // تاريخ الاستحقاق - سيُحسب بالصيغة
            '', // حالة السداد - سيُحسب بالصيغة
            '', // الشهر - سيُحسب بالصيغة
            rowData[columns.NOTES.index - 1] || `(من البوت: ${rowData[columns.TELEGRAM_USER.index - 1]})`,
            '',                                          // Y: كشف
            '',                                          // Z: رقم الأوردر
            rowData[columns.UNIT_COUNT.index - 1] || '', // AA: عدد الوحدات
            `(من البوت: ${rowData[columns.TELEGRAM_USER.index - 1] || ''})` // AB: مصدر الإدخال
        ];

        // إضافة الصف
        Logger.log('📝 جاري كتابة البيانات في الصف ' + newRow);
        mainSheet.getRange(newRow, 1, 1, 28).setValues([mainRowData]);
        Logger.log('✅ تم كتابة البيانات بنجاح!');

        // Force flush to ensure data is written
        SpreadsheetApp.flush();
        Logger.log('✅ تم حفظ التغييرات (flush)');

        // ⭐ حساب الأعمدة التلقائية (M, U, O, V) - لأن setValues لا يُفعّل onEdit
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        try {
            calculateUsdValue_(mainSheet, newRow);           // M: القيمة بالدولار
            calculateDueDate_(ss, mainSheet, newRow);        // U: تاريخ الاستحقاق
            recalculatePartyBalance_(mainSheet, newRow);     // O: الرصيد + V: حالة السداد
            Logger.log('✅ تم حساب الأعمدة التلقائية (M, U, O, V)');
        } catch (calcError) {
            Logger.log('⚠️ خطأ في حساب الأعمدة التلقائية: ' + calcError.message);
        }

        // تحديث حالة الحركة في شيت البوت
        let reviewerEmail = 'Unknown';
        try {
            reviewerEmail = Session.getActiveUser().getEmail();
        } catch (e) {
            console.log('Could not get active user email: ' + e.message);
        }
        botSheet.getRange(rowNumber, columns.REVIEW_STATUS.index).setValue(CONFIG.TELEGRAM_BOT.REVIEW_STATUS.APPROVED);
        botSheet.getRange(rowNumber, columns.REVIEWER.index).setValue(reviewerEmail);
        botSheet.getRange(rowNumber, columns.REVIEW_TIMESTAMP.index).setValue(new Date());

        // إرسال إشعار للمستخدم
        const chatId = rowData[columns.TELEGRAM_CHAT_ID.index - 1];
        if (chatId) {
            notifyUserApproval(chatId, {
                transactionId: rowData[columns.TRANSACTION_ID.index - 1],
                date: Utilities.formatDate(new Date(rowData[columns.DATE.index - 1]), 'Asia/Istanbul', 'dd/MM/yyyy'),
                amount: rowData[columns.AMOUNT.index - 1],
                currency: rowData[columns.CURRENCY.index - 1],
                partyName: rowData[columns.PARTY_NAME.index - 1]
            });
        }

        // تسجيل النشاط
        logActivity('اعتماد حركة بوت', CONFIG.SHEETS.BOT_TRANSACTIONS, rowNumber,
            rowData[columns.TRANSACTION_ID.index - 1], 'تم نقل الحركة لدفتر الحركات');

        return { success: true, newRowNumber: newRow };

    } catch (error) {
        Logger.log('Error approving transaction: ' + error.message);
        return { success: false, error: error.message };
    }
}

/**
 * رفض حركة من شيت حركات البوت (النظام القديم)
 * ⚠️ ملاحظة: هذه الدالة للنظام القديم. الحركات الآن تُسجَّل مباشرة.
 * لرفض حركة من دفتر الحركات الرئيسي، استخدم rejectTransaction في BotSheets.js
 */
function rejectBotTransaction(rowNumber, reason) {
    try {
        const botSheet = getBotTransactionsSheet();
        const columns = BOT_CONFIG.BOT_TRANSACTIONS_COLUMNS;

        // التحقق من وجود الشيت
        if (!botSheet) {
            return { success: false, error: 'شيت حركات البوت غير موجود.\nالحركات تُسجَّل مباشرة الآن.\nلرفض حركة: حدد الصف ← القائمة ← "رفض الحركة المحددة"' };
        }

        // التحقق من الحالة
        const currentStatus = botSheet.getRange(rowNumber, columns.REVIEW_STATUS.index).getValue();
        if (currentStatus !== CONFIG.TELEGRAM_BOT.REVIEW_STATUS.PENDING) {
            return { success: false, error: 'الحركة ليست في حالة انتظار.\nربما تم اعتمادها أو رفضها مسبقاً، أو أنها حركة مباشرة.' };
        }

        // تحديث حالة الحركة
        let reviewerEmail = 'Unknown';
        try {
            reviewerEmail = Session.getActiveUser().getEmail();
        } catch (e) {
            console.log('Could not get active user email: ' + e.message);
        }
        botSheet.getRange(rowNumber, columns.REVIEW_STATUS.index).setValue(CONFIG.TELEGRAM_BOT.REVIEW_STATUS.REJECTED);
        botSheet.getRange(rowNumber, columns.REVIEWER.index).setValue(reviewerEmail);
        botSheet.getRange(rowNumber, columns.REVIEW_TIMESTAMP.index).setValue(new Date());
        botSheet.getRange(rowNumber, columns.REVIEW_NOTES.index).setValue(reason);

        // قراءة بيانات الحركة للإشعار
        const rowData = botSheet.getRange(rowNumber, 1, 1, Object.keys(columns).length).getValues()[0];

        // إرسال إشعار للمستخدم
        const chatId = rowData[columns.TELEGRAM_CHAT_ID.index - 1];
        if (chatId) {
            notifyUserRejection(chatId, {
                transactionId: rowData[columns.TRANSACTION_ID.index - 1],
                date: Utilities.formatDate(new Date(rowData[columns.DATE.index - 1]), 'Asia/Istanbul', 'dd/MM/yyyy'),
                amount: rowData[columns.AMOUNT.index - 1],
                currency: rowData[columns.CURRENCY.index - 1],
                partyName: rowData[columns.PARTY_NAME.index - 1]
            }, reason);
        }

        // تسجيل النشاط
        logActivity('رفض حركة بوت', CONFIG.SHEETS.BOT_TRANSACTIONS, rowNumber,
            rowData[columns.TRANSACTION_ID.index - 1], 'سبب الرفض: ' + reason);

        return { success: true };

    } catch (error) {
        Logger.log('Error rejecting transaction: ' + error.message);
        return { success: false, error: error.message };
    }
}

/**
 * اعتماد وإضافة طرف جديد
 */
function approveAndAddNewParty(partyName, linkedChatId) {
    try {
        const botPartiesSheet = getBotPartiesSheet();
        const mainPartiesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.PARTIES);
        const botColumns = BOT_CONFIG.BOT_PARTIES_COLUMNS;

        if (!mainPartiesSheet) {
            return { success: false, error: 'لم يتم العثور على قاعدة بيانات الأطراف' };
        }

        // البحث عن الطرف في شيت أطراف البوت
        const data = botPartiesSheet.getDataRange().getValues();
        let partyRow = null;
        let rowIndex = -1;

        for (let i = 1; i < data.length; i++) {
            if (data[i][botColumns.PARTY_NAME.index - 1] === partyName &&
                data[i][botColumns.REVIEW_STATUS.index - 1] === CONFIG.TELEGRAM_BOT.REVIEW_STATUS.PENDING) {
                partyRow = data[i];
                rowIndex = i + 1;
                break;
            }
        }

        if (!partyRow) {
            // الطرف غير موجود في شيت البوت - أضفه مباشرة
            const newPartyRow = [
                partyName,
                'مورد', // نوع افتراضي
                '', '', '', '', '', '', '' // باقي الأعمدة فارغة
            ];
            mainPartiesSheet.appendRow(newPartyRow);
            return { success: true };
        }

        // نسخ الطرف لقاعدة البيانات الرئيسية
        const mainPartyRow = [
            partyRow[botColumns.PARTY_NAME.index - 1],
            partyRow[botColumns.PARTY_TYPE.index - 1],
            partyRow[botColumns.SPECIALIZATION.index - 1] || '',
            partyRow[botColumns.PHONE.index - 1] || '',
            partyRow[botColumns.EMAIL.index - 1] || '',
            partyRow[botColumns.LOCATION.index - 1] || '',
            partyRow[botColumns.PAYMENT_METHOD.index - 1] || '',
            partyRow[botColumns.BANK_DETAILS.index - 1] || '',
            `(أُضيف من البوت بواسطة: ${partyRow[botColumns.TELEGRAM_USER.index - 1] || ''})`
        ];

        mainPartiesSheet.appendRow(mainPartyRow);

        // تحديث حالة الطرف في شيت البوت
        if (rowIndex > 0) {
            botPartiesSheet.getRange(rowIndex, botColumns.REVIEW_STATUS.index)
                .setValue(CONFIG.TELEGRAM_BOT.REVIEW_STATUS.APPROVED);
            botPartiesSheet.getRange(rowIndex, botColumns.REVIEWER.index)
                .setValue(Session.getActiveUser().getEmail());
            botPartiesSheet.getRange(rowIndex, botColumns.REVIEW_TIMESTAMP.index)
                .setValue(new Date());
        }

        return { success: true };

    } catch (error) {
        Logger.log('Error adding new party: ' + error.message);
        return { success: false, error: error.message };
    }
}

/**
 * الحصول على الأطراف المعلقة
 */
function getPendingBotParties() {
    const sheet = getBotPartiesSheet();
    const columns = BOT_CONFIG.BOT_PARTIES_COLUMNS;
    const data = sheet.getDataRange().getValues();
    const pending = [];

    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const status = row[columns.REVIEW_STATUS.index - 1];

        if (status === CONFIG.TELEGRAM_BOT.REVIEW_STATUS.PENDING) {
            pending.push({
                rowNumber: i + 1,
                name: row[columns.PARTY_NAME.index - 1],
                type: row[columns.PARTY_TYPE.index - 1],
                telegramUser: row[columns.TELEGRAM_USER.index - 1],
                linkedTransactions: row[columns.LINKED_TRANSACTIONS.index - 1]
            });
        }
    }

    return pending;
}

// ==================== دوال الاعتماد الجماعي ====================

/**
 * اعتماد جميع الحركات المعلقة
 */
function approveAllPendingTransactions() {
    const ui = SpreadsheetApp.getUi();
    const pending = getPendingBotTransactions();

    if (pending.length === 0) {
        ui.alert('📋 لا توجد حركات معلقة');
        return;
    }

    const result = ui.alert(
        '🔄 اعتماد جماعي',
        `هل تريد اعتماد جميع الحركات المعلقة؟\n\nعدد الحركات: ${pending.length}`,
        ui.ButtonSet.YES_NO
    );

    if (result !== ui.Button.YES) return;

    let successCount = 0;
    let errorCount = 0;

    pending.forEach(tx => {
        const approvalResult = approveTransaction(tx.rowNumber);
        if (approvalResult.success) {
            successCount++;
        } else {
            errorCount++;
            console.log(`Failed to approve transaction ${tx.transactionId}: ${approvalResult.error}`);
        }
    });

    ui.alert(
        '✅ تم الاعتماد الجماعي',
        `تم اعتماد: ${successCount} حركة\nفشل: ${errorCount} حركة`,
        ui.ButtonSet.OK
    );
}

/**
 * عرض تقرير الحركات المعلقة
 */
function showPendingTransactionsReport() {
    const pending = getPendingBotTransactions();
    const pendingParties = getPendingBotParties();

    let message = '📊 تقرير الحركات المعلقة\n';
    message += '═'.repeat(30) + '\n\n';

    message += `📋 حركات معلقة: ${pending.length}\n`;
    message += `👤 أطراف جديدة: ${pendingParties.length}\n\n`;

    if (pending.length > 0) {
        message += '📋 تفاصيل الحركات:\n';
        message += '─'.repeat(25) + '\n';

        pending.slice(0, 10).forEach((tx, i) => {
            message += `${i + 1}. ${tx.transactionId}\n`;
            message += `   ${tx.amount} ${tx.currency} - ${tx.partyName}\n`;
            message += `   ${tx.nature}\n\n`;
        });

        if (pending.length > 10) {
            message += `... و${pending.length - 10} حركات أخرى\n`;
        }
    }

    SpreadsheetApp.getUi().alert('📊 تقرير المعلقات', message, SpreadsheetApp.getUi().ButtonSet.OK);
}

// ==================== دوال مساعدة ====================
// ملاحظة: دالة logActivity موجودة في Main.js وتُستخدم من هناك

/**
 * الحصول على إحصائيات البوت
 */
function getBotStatistics() {
    const botSheet = getBotTransactionsSheet();
    const columns = BOT_CONFIG.BOT_TRANSACTIONS_COLUMNS;
    const data = botSheet.getDataRange().getValues();

    let stats = {
        total: 0,
        pending: 0,
        approved: 0,
        rejected: 0,
        totalAmount: 0
    };

    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        if (!row[columns.TRANSACTION_ID.index - 1]) continue;

        stats.total++;
        const status = row[columns.REVIEW_STATUS.index - 1];
        const amount = parseFloat(row[columns.AMOUNT_USD.index - 1]) || 0;

        switch (status) {
            case CONFIG.TELEGRAM_BOT.REVIEW_STATUS.PENDING:
                stats.pending++;
                break;
            case CONFIG.TELEGRAM_BOT.REVIEW_STATUS.APPROVED:
                stats.approved++;
                stats.totalAmount += amount;
                break;
            case CONFIG.TELEGRAM_BOT.REVIEW_STATUS.REJECTED:
                stats.rejected++;
                break;
        }
    }

    return stats;
}

/**
 * عرض إحصائيات البوت
 */
function showBotStatistics() {
    const stats = getBotStatistics();

    const message = `
📊 إحصائيات بوت تليجرام
═══════════════════

📋 إجمالي الحركات: ${stats.total}

⏳ قيد الانتظار: ${stats.pending}
✅ معتمدة: ${stats.approved}
❌ مرفوضة: ${stats.rejected}

💵 إجمالي المعتمد: $${stats.totalAmount.toFixed(2)}
    `;

    SpreadsheetApp.getUi().alert('📊 إحصائيات البوت', message, SpreadsheetApp.getUi().ButtonSet.OK);
}

// ==================== دوال التشخيص ====================

/**
 * تشخيص مشكلة الاعتماد - شغّل هذه الدالة لفهم المشكلة
 */
function diagnoseApprovalIssue() {
    const ui = SpreadsheetApp.getUi();
    let report = '🔍 تقرير التشخيص\n';
    report += '═'.repeat(30) + '\n\n';

    try {
        // 1. فحص شيت حركات البوت
        const botSheet = getBotTransactionsSheet();
        if (!botSheet) {
            report += '❌ شيت حركات البوت غير موجود!\n';
            ui.alert('تقرير التشخيص', report, ui.ButtonSet.OK);
            return;
        }
        report += '✅ شيت حركات البوت: موجود\n';
        report += `   الاسم: ${botSheet.getName()}\n`;
        report += `   عدد الصفوف: ${botSheet.getLastRow()}\n`;
        report += `   عدد الأعمدة: ${botSheet.getLastColumn()}\n\n`;

        // 2. فحص شيت دفتر الحركات
        const mainSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.TRANSACTIONS);
        if (!mainSheet) {
            report += '❌ شيت دفتر الحركات المالية غير موجود!\n';
            report += `   الاسم المطلوب: "${CONFIG.SHEETS.TRANSACTIONS}"\n`;
            ui.alert('تقرير التشخيص', report, ui.ButtonSet.OK);
            return;
        }
        report += '✅ شيت دفتر الحركات: موجود\n';
        report += `   الاسم: ${mainSheet.getName()}\n`;
        report += `   عدد الصفوف: ${mainSheet.getLastRow()}\n`;
        report += `   عدد الأعمدة: ${mainSheet.getLastColumn()}\n\n`;

        // 3. فحص الحركات المعلقة
        const columns = BOT_CONFIG.BOT_TRANSACTIONS_COLUMNS;
        const pendingValue = CONFIG.TELEGRAM_BOT.REVIEW_STATUS.PENDING;
        report += `📋 قيمة "قيد الانتظار" المتوقعة: "${pendingValue}"\n\n`;

        // 4. فحص أول حركة معلقة
        const data = botSheet.getDataRange().getValues();
        let foundPending = false;

        for (let i = 1; i < data.length; i++) {
            const row = data[i];
            const statusIndex = columns.REVIEW_STATUS.index - 1;
            const actualStatus = row[statusIndex];

            if (i === 1) {
                report += `📊 عمود حالة المراجعة:\n`;
                report += `   الفهرس المتوقع: ${columns.REVIEW_STATUS.index}\n`;
                report += `   القيمة في الصف 2: "${actualStatus}"\n\n`;
            }

            if (actualStatus === pendingValue) {
                foundPending = true;
                report += `✅ وجدت حركة معلقة في الصف ${i + 1}\n`;
                report += `   رقم الحركة: ${row[columns.TRANSACTION_ID.index - 1]}\n`;
                report += `   المبلغ: ${row[columns.AMOUNT.index - 1]}\n`;
                report += `   الحالة: "${actualStatus}"\n\n`;

                // محاولة الاعتماد التجريبي
                report += '🧪 محاولة اعتماد تجريبية...\n';
                const result = approveTransaction(i + 1);
                if (result.success) {
                    report += `✅ نجح الاعتماد! الصف الجديد: ${result.newRowNumber}\n`;
                } else {
                    report += `❌ فشل الاعتماد: ${result.error}\n`;
                }
                break;
            }
        }

        if (!foundPending) {
            report += '⚠️ لا توجد حركات بحالة "قيد الانتظار"\n';
            report += '\nالحالات الموجودة في عمود المراجعة:\n';
            const statuses = new Set();
            for (let i = 1; i < Math.min(data.length, 10); i++) {
                const status = data[i][columns.REVIEW_STATUS.index - 1];
                if (status) statuses.add(status);
            }
            statuses.forEach(s => report += `   - "${s}"\n`);
        }

    } catch (error) {
        report += `\n🔥 خطأ: ${error.message}\n`;
        report += `Stack: ${error.stack}\n`;
    }

    Logger.log(report);
    ui.alert('🔍 تقرير التشخيص', report, ui.ButtonSet.OK);
}

/**
 * اعتماد يدوي مع تفاصيل - للتجربة
 */
function manualApproveWithDetails() {
    const ui = SpreadsheetApp.getUi();

    const response = ui.prompt(
        '🔧 اعتماد يدوي',
        'أدخل رقم الصف في شيت حركات البوت:',
        ui.ButtonSet.OK_CANCEL
    );

    if (response.getSelectedButton() !== ui.Button.OK) return;

    const rowNumber = parseInt(response.getResponseText().trim());
    if (isNaN(rowNumber) || rowNumber < 2) {
        ui.alert('❌ رقم صف غير صالح');
        return;
    }

    Logger.log('=== بدء الاعتماد اليدوي للصف ' + rowNumber + ' ===');

    try {
        const botSheet = getBotTransactionsSheet();
        const mainSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.TRANSACTIONS);
        const columns = BOT_CONFIG.BOT_TRANSACTIONS_COLUMNS;

        Logger.log('عدد أعمدة التعريف: ' + Object.keys(columns).length);
        Logger.log('عدد أعمدة شيت البوت الفعلية: ' + botSheet.getLastColumn());

        // قراءة البيانات
        const rowData = botSheet.getRange(rowNumber, 1, 1, botSheet.getLastColumn()).getValues()[0];
        Logger.log('بيانات الصف: ' + JSON.stringify(rowData));

        const currentStatus = rowData[columns.REVIEW_STATUS.index - 1];
        Logger.log('الحالة الحالية: "' + currentStatus + '"');
        Logger.log('الحالة المتوقعة: "' + CONFIG.TELEGRAM_BOT.REVIEW_STATUS.PENDING + '"');

        // إجبار الاعتماد
        const mainLastRow = mainSheet.getLastRow();
        const newRow = mainLastRow + 1;

        const mainRowData = [
            newRow - 1,
            rowData[columns.DATE.index - 1],
            rowData[columns.NATURE.index - 1],
            rowData[columns.CLASSIFICATION.index - 1] || '',
            rowData[columns.PROJECT_CODE.index - 1] || '',
            rowData[columns.PROJECT_NAME.index - 1] || '',
            rowData[columns.ITEM.index - 1] || '',
            rowData[columns.DETAILS.index - 1] || '',
            rowData[columns.PARTY_NAME.index - 1] || '',
            rowData[columns.AMOUNT.index - 1] || 0,
            rowData[columns.CURRENCY.index - 1] || 'USD',
            rowData[columns.EXCHANGE_RATE.index - 1] || 1,
            rowData[columns.AMOUNT_USD.index - 1] || 0,
            rowData[columns.MOVEMENT_TYPE.index - 1] || '',
            '', '', // الرصيد، رقم مرجعي
            rowData[columns.PAYMENT_METHOD.index - 1] || '',
            rowData[columns.PAYMENT_TERM_TYPE.index - 1] || '',
            rowData[columns.WEEKS.index - 1] || 0,
            rowData[columns.CUSTOM_DATE.index - 1] || '',
            '', '', '', // تاريخ استحقاق، حالة سداد، شهر
            rowData[columns.NOTES.index - 1] || `(من البوت)`,
            '',                                          // Y: كشف
            '',                                          // Z: رقم الأوردر
            rowData[columns.UNIT_COUNT.index - 1] || '', // AA: عدد الوحدات
            `(من البوت: ${rowData[columns.TELEGRAM_USER.index - 1] || ''})` // AB: مصدر الإدخال
        ];

        Logger.log('البيانات للإدخال: ' + JSON.stringify(mainRowData));
        Logger.log('عدد الأعمدة: ' + mainRowData.length);

        // الإدخال
        mainSheet.getRange(newRow, 1, 1, 28).setValues([mainRowData]);

        // ⭐ حساب الأعمدة التلقائية (M, U, O, V) - لأن setValues لا يُفعّل onEdit
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        try {
            calculateUsdValue_(mainSheet, newRow);           // M: القيمة بالدولار
            calculateDueDate_(ss, mainSheet, newRow);        // U: تاريخ الاستحقاق
            recalculatePartyBalance_(mainSheet, newRow);     // O: الرصيد + V: حالة السداد
            Logger.log('✅ تم حساب الأعمدة التلقائية (M, U, O, V)');
        } catch (calcError) {
            Logger.log('⚠️ خطأ في حساب الأعمدة التلقائية: ' + calcError.message);
        }

        // تحديث حالة البوت
        botSheet.getRange(rowNumber, columns.REVIEW_STATUS.index).setValue(CONFIG.TELEGRAM_BOT.REVIEW_STATUS.APPROVED);
        botSheet.getRange(rowNumber, columns.REVIEWER.index).setValue(Session.getActiveUser().getEmail() || 'manual');
        botSheet.getRange(rowNumber, columns.REVIEW_TIMESTAMP.index).setValue(new Date());

        ui.alert('✅ نجاح', `تم نقل الحركة للصف ${newRow} في دفتر الحركات`, ui.ButtonSet.OK);

    } catch (error) {
        Logger.log('خطأ: ' + error.message);
        Logger.log('Stack: ' + error.stack);
        ui.alert('❌ خطأ', error.message, ui.ButtonSet.OK);
    }
}
