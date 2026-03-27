/**
 * ═══════════════════════════════════════════════════════════════════════════
 *                          PDF GENERATOR
 *                   توليد وإرسال ملفات PDF للتقارير
 * ═══════════════════════════════════════════════════════════════════════════
 */

/**
 * ═══════════════════════════════════════════════════════════════════════════
 *                         تصدير شيت كـ PDF
 * ═══════════════════════════════════════════════════════════════════════════
 */

/**
 * تصدير شيت محدد كملف PDF
 * @param {Sheet} sheet - الشيت المراد تصديره
 * @param {string} fileName - اسم الملف (بدون .pdf)
 * @param {Object} options - خيارات إضافية
 * @returns {Blob} - ملف PDF كـ Blob
 */
function exportSheetAsPDF(sheet, fileName, options) {
    try {
        const ss = sheet.getParent();
        const sheetId = sheet.getSheetId();

        // الخيارات الافتراضية
        const defaultOptions = {
            size: 'A4',           // حجم الورقة
            portrait: true,       // طولي (Portrait)
            fitWidth: true,       // ملائمة العرض
            gridlines: false,     // بدون خطوط الشبكة
            printtitle: false,    // بدون العنوان في الأعلى
            sheetnames: false,    // بدون اسم الشيت
            pagenumbers: false,   // بدون أرقام الصفحات
            fzr: false,           // بدون تكرار الصفوف المجمدة
            fzc: false            // بدون تكرار الأعمدة المجمدة
        };

        const opts = { ...defaultOptions, ...options };

        // بناء رابط التصدير
        const exportUrl = ss.getUrl().replace(/\/edit.*$/, '') +
            '/export?format=pdf' +
            '&gid=' + sheetId +
            '&size=' + opts.size +
            '&portrait=' + opts.portrait +
            '&fitw=' + opts.fitWidth +
            '&gridlines=' + opts.gridlines +
            '&printtitle=' + opts.printtitle +
            '&sheetnames=' + opts.sheetnames +
            '&pagenumbers=' + opts.pagenumbers +
            '&fzr=' + opts.fzr +
            '&fzc=' + opts.fzc;

        Logger.log('📄 Exporting PDF from URL: ' + exportUrl);

        // جلب الملف
        const token = ScriptApp.getOAuthToken();
        const response = UrlFetchApp.fetch(exportUrl, {
            headers: {
                'Authorization': 'Bearer ' + token
            },
            muteHttpExceptions: true
        });

        if (response.getResponseCode() !== 200) {
            throw new Error('Failed to export PDF: ' + response.getContentText());
        }

        const pdfBlob = response.getBlob().setName(fileName + '.pdf');
        Logger.log('✅ PDF exported successfully: ' + fileName + '.pdf');

        return pdfBlob;

    } catch (error) {
        Logger.log('❌ Error exporting PDF: ' + error.message);
        throw error;
    }
}

/**
 * ═══════════════════════════════════════════════════════════════════════════
 *                         حفظ PDF في Drive
 * ═══════════════════════════════════════════════════════════════════════════
 */

/**
 * الحصول على مجلد التقارير أو إنشاؤه
 * @returns {Folder} - مجلد التقارير
 */
function getOrCreateReportsFolder() {
    try {
        const folderName = REPORTS_CONFIG.STORAGE.REPORTS_FOLDER_NAME;
        const folders = DriveApp.getFoldersByName(folderName);

        if (folders.hasNext()) {
            return folders.next();
        }

        // إنشاء المجلد الرئيسي
        const mainFolder = DriveApp.createFolder(folderName);
        Logger.log('📁 Created reports folder: ' + folderName);

        return mainFolder;

    } catch (error) {
        Logger.log('❌ Error getting/creating reports folder: ' + error.message);
        throw error;
    }
}

/**
 * الحصول على مجلد الشهر الحالي أو إنشاؤه
 * @returns {Folder} - مجلد الشهر
 */
function getOrCreateMonthFolder() {
    try {
        const mainFolder = getOrCreateReportsFolder();
        const now = new Date();
        const monthFolderName = Utilities.formatDate(now, CONFIG.COMPANY.TIMEZONE, 'yyyy-MM');

        // البحث عن مجلد الشهر
        const folders = mainFolder.getFoldersByName(monthFolderName);

        if (folders.hasNext()) {
            return folders.next();
        }

        // إنشاء مجلد الشهر
        const monthFolder = mainFolder.createFolder(monthFolderName);
        Logger.log('📁 Created month folder: ' + monthFolderName);

        return monthFolder;

    } catch (error) {
        Logger.log('❌ Error getting/creating month folder: ' + error.message);
        throw error;
    }
}

/**
 * حفظ PDF في أرشيف التقارير
 * @param {Blob} pdfBlob - ملف PDF
 * @param {string} reportType - نوع التقرير
 * @param {string} partyName - اسم الطرف (اختياري)
 * @returns {File} - الملف المحفوظ
 */
function savePDFToArchive(pdfBlob, reportType, partyName) {
    try {
        const folder = getOrCreateMonthFolder();
        const now = new Date();
        const dateStr = Utilities.formatDate(now, CONFIG.COMPANY.TIMEZONE, 'yyyy-MM-dd');

        // بناء اسم الملف
        let fileName = reportType;
        if (partyName) {
            fileName += ' - ' + partyName;
        }
        fileName += ' - ' + dateStr + '.pdf';

        // حفظ الملف
        const file = folder.createFile(pdfBlob.setName(fileName));
        Logger.log('💾 PDF saved to archive: ' + fileName);

        return file;

    } catch (error) {
        Logger.log('❌ Error saving PDF to archive: ' + error.message);
        throw error;
    }
}

/**
 * ═══════════════════════════════════════════════════════════════════════════
 *                      إرسال PDF عبر تليجرام
 * ═══════════════════════════════════════════════════════════════════════════
 */

/**
 * إرسال ملف PDF للمستخدم عبر تليجرام
 * @param {string} chatId - معرف المحادثة
 * @param {Blob} pdfBlob - ملف PDF
 * @param {string} caption - التعليق على الملف
 * @returns {boolean} - نجاح الإرسال
 */
function sendPDFToTelegram(chatId, pdfBlob, caption) {
    try {
        // ⭐ الحصول على التوكن باستخدام نفس الدالة المستخدمة في sendAIMessage
        const token = getAIBotToken();

        // ⭐ التأكد من أن الـ Blob له نوع محتوى صحيح واسم ملف
        pdfBlob.setContentType('application/pdf');
        const fileName = pdfBlob.getName() || 'report.pdf';
        pdfBlob.setName(fileName);

        Logger.log('📤 Sending PDF to chat_id: ' + chatId);
        Logger.log('📄 PDF: ' + fileName + ', size: ' + pdfBlob.getBytes().length + ' bytes');

        const url = 'https://api.telegram.org/bot' + token + '/sendDocument';

        // ⭐ استخدام contentType: 'multipart/form-data' صراحةً
        // لا نحدده لأن UrlFetchApp يحدده تلقائياً عند وجود Blob
        const formData = {
            method: 'post',
            payload: {
                chat_id: String(chatId),
                document: pdfBlob,
                caption: caption || '',
                parse_mode: 'Markdown'
            },
            muteHttpExceptions: true
        };

        const response = UrlFetchApp.fetch(url, formData);
        const responseText = response.getContentText();

        Logger.log('📥 Telegram response: ' + responseText);

        const result = JSON.parse(responseText);

        if (result.ok) {
            Logger.log('✅ PDF sent to Telegram successfully');
            return true;
        } else {
            Logger.log('❌ Telegram API error: ' + result.description);
            Logger.log('❌ Error code: ' + result.error_code);

            // محاولة بديلة إذا فشلت الطريقة الأولى
            Logger.log('⚠️ Trying Drive URL method...');
            return sendPDFViaURL(chatId, pdfBlob, caption, token);
        }

    } catch (error) {
        Logger.log('❌ Error sending PDF to Telegram: ' + error.message);
        Logger.log('❌ Stack: ' + error.stack);
        return false;
    }
}

/**
 * طريقة بديلة لإرسال PDF - تحميل للـ Drive ثم إرسال الرابط العام
 * هذه الطريقة تعمل عندما تفشل الطريقة الأولى
 */
function sendPDFViaURL(chatId, pdfBlob, caption, token) {
    try {
        Logger.log('📤 Alternative: Uploading to Drive and sending URL...');

        // حفظ الملف مؤقتاً في Drive
        const tempFile = DriveApp.createFile(pdfBlob);
        tempFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

        // الحصول على رابط التحميل المباشر
        const fileId = tempFile.getId();
        const directUrl = 'https://drive.google.com/uc?export=download&id=' + fileId;
        Logger.log('📁 Direct download URL: ' + directUrl);

        // إرسال عبر URL - Telegram يدعم إرسال ملفات عبر URL
        const url = 'https://api.telegram.org/bot' + token + '/sendDocument';

        const response = UrlFetchApp.fetch(url, {
            method: 'post',
            contentType: 'application/json',
            payload: JSON.stringify({
                chat_id: String(chatId),
                document: directUrl,
                caption: caption || '',
                parse_mode: 'Markdown'
            }),
            muteHttpExceptions: true
        });

        const responseText = response.getContentText();
        Logger.log('📥 Drive URL response: ' + responseText);

        const result = JSON.parse(responseText);

        if (result.ok) {
            Logger.log('✅ PDF sent via Drive URL method');
            // حذف الملف المؤقت بعد النجاح
            Utilities.sleep(3000);
            tempFile.setTrashed(true);
            Logger.log('🗑️ Temp file deleted');
            return true;
        } else {
            Logger.log('❌ Drive URL method failed: ' + result.description);

            // إذا فشل أيضاً، نرسل رسالة للمستخدم مع رابط التحميل
            // لا نحذف الملف حتى يتمكن المستخدم من تحميله
            sendPDFDownloadLink(chatId, tempFile, caption);
            return true;
        }

    } catch (error) {
        Logger.log('❌ Drive URL method error: ' + error.message);
        return false;
    }
}

/**
 * إرسال رابط تحميل الملف كرسالة نصية (خطة أخيرة)
 */
function sendPDFDownloadLink(chatId, driveFile, caption) {
    try {
        const fileUrl = driveFile.getUrl();
        const message = (caption || '📄 *التقرير جاهز*') +
            '\n\n📥 لم نتمكن من إرسال الملف مباشرة.\nيمكنك تحميله من هنا:\n' + fileUrl;

        sendAIMessage(chatId, message, { parse_mode: 'Markdown' });

    } catch (error) {
        Logger.log('❌ Error sending download link: ' + error.message);
    }
}

/**
 * ═══════════════════════════════════════════════════════════════════════════
 *                    الدالة الرئيسية لتوليد وإرسال التقرير
 * ═══════════════════════════════════════════════════════════════════════════
 */

/**
 * توليد تقرير PDF وإرساله للمستخدم
 * @param {string} chatId - معرف المحادثة
 * @param {string} reportType - نوع التقرير
 * @param {Sheet} sheet - الشيت المراد تصديره
 * @param {string} partyName - اسم الطرف (اختياري)
 * @param {boolean} saveToArchive - حفظ في الأرشيف
 * @returns {Object} - نتيجة العملية
 */
function generateAndSendReport(chatId, reportType, sheet, partyName, saveToArchive) {
    try {
        Logger.log('📊 Generating report: ' + reportType + ' for ' + (partyName || 'N/A'));

        // 1. تصدير الشيت كـ PDF
        const fileName = buildReportFileName(reportType, partyName);
        const pdfBlob = exportSheetAsPDF(sheet, fileName, {
            portrait: true,
            fitWidth: true
        });

        // 2. حفظ في الأرشيف إذا مطلوب
        let savedFile = null;
        if (saveToArchive) {
            savedFile = savePDFToArchive(pdfBlob, reportType, partyName);
        }

        // 3. بناء التعليق
        const caption = buildReportCaption(reportType, partyName, saveToArchive);

        // 4. إرسال للتليجرام
        const sent = sendPDFToTelegram(chatId, pdfBlob, caption);

        return {
            success: sent,
            fileName: fileName + '.pdf',
            savedToArchive: saveToArchive,
            archiveFile: savedFile
        };

    } catch (error) {
        Logger.log('❌ Error in generateAndSendReport: ' + error.message);
        return {
            success: false,
            error: error.message
        };
    }
}

/**
 * بناء اسم ملف التقرير
 */
// ⭐ خريطة أسماء التقارير الموحدة (بدلاً من التكرار في كل دالة)
var REPORT_DISPLAY_NAMES_ = {
    'statement': { name: 'كشف حساب', emoji: '📄' },
    'alerts': { name: 'تنبيهات الاستحقاق', emoji: '⏰' },
    'balances': { name: 'تقرير الأرصدة', emoji: '💰' },
    'profitability': { name: 'ربحية المشاريع', emoji: '📈' },
    'expenses': { name: 'تقرير المصروفات', emoji: '📊' },
    'revenues': { name: 'تقرير الإيرادات', emoji: '💵' },
    'dynamic_expenses': { name: 'تحليل المصروفات', emoji: '📊' },
    'cashflow': { name: 'التدفقات النقدية', emoji: '💸' },
    'income_statement': { name: 'قائمة الدخل', emoji: '📋' },
    'balance_sheet': { name: 'المركز المالي', emoji: '🏦' },
    'funders': { name: 'تقرير الممولين', emoji: '💼' },
    'projects': { name: 'تقرير المشروعات', emoji: '🎬' }
};

function buildReportFileName(reportType, partyName) {
    var info = REPORT_DISPLAY_NAMES_[reportType];
    let name = info ? info.name : reportType;
    if (partyName) {
        name += ' - ' + partyName;
    }

    const now = new Date();
    const dateStr = Utilities.formatDate(now, CONFIG.COMPANY.TIMEZONE, 'yyyy-MM-dd');
    name += ' - ' + dateStr;

    return name;
}

/**
 * بناء تعليق التقرير
 */
function buildReportCaption(reportType, partyName, savedToArchive) {
    var info = REPORT_DISPLAY_NAMES_[reportType];
    var displayName = info ? (info.emoji + ' ' + info.name) : reportType;

    let caption = '*' + displayName + '*';

    if (partyName) {
        caption += '\n👤 ' + partyName;
    }

    const now = new Date();
    const dateStr = Utilities.formatDate(now, CONFIG.COMPANY.TIMEZONE, 'dd/MM/yyyy HH:mm');
    caption += '\n📅 ' + dateStr;

    if (savedToArchive) {
        caption += '\n\n💾 _تم حفظ نسخة في الأرشيف_';
    }

    return caption;
}

/**
 * ═══════════════════════════════════════════════════════════════════════════
 *                         حذف الشيت المؤقت
 * ═══════════════════════════════════════════════════════════════════════════
 */

/**
 * حذف شيت مؤقت بعد التصدير
 * @param {Sheet} sheet - الشيت المراد حذفه
 * @param {boolean} force - فرض الحذف
 */
function deleteTemporarySheet(sheet, force) {
    try {
        const sheetName = sheet.getName();

        // التحقق من أنه شيت تقرير مؤقت
        const isTemporary = sheetName.startsWith('كشف حساب') ||
            sheetName.startsWith('تقرير');

        if (isTemporary || force) {
            const ss = sheet.getParent();
            ss.deleteSheet(sheet);
            Logger.log('🗑️ Temporary sheet deleted: ' + sheetName);
        }

    } catch (error) {
        Logger.log('⚠️ Could not delete temporary sheet: ' + error.message);
    }
}

/**
 * ═══════════════════════════════════════════════════════════════════════════
 *                    دوال مساعدة لأنواع التقارير المختلفة
 * ═══════════════════════════════════════════════════════════════════════════
 */

/**
 * توليد كشف حساب وإرساله
 * @param {string} chatId - معرف المحادثة
 * @param {string} partyName - اسم الطرف
 * @param {string} partyType - نوع الطرف (مورد/عميل/ممول)
 * @returns {Object} - نتيجة العملية
 */
function generateStatementPDF(chatId, partyName, partyType) {
    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();

        // ⭐ استخدام الدالة الخاصة بالبوت (بدون UI)
        const sheet = generateStatementForBot_(ss, partyName, partyType);

        if (!sheet) {
            throw new Error('لم يتم إنشاء شيت كشف الحساب');
        }

        // تصدير وإرسال (مع حفظ في الأرشيف)
        const result = generateAndSendReport(chatId, 'statement', sheet, partyName, true);

        // حذف الشيت المؤقت بعد الإرسال
        if (result.success) {
            deleteTemporarySheet(sheet, true);
        }

        return result;

    } catch (error) {
        Logger.log('❌ Error generating statement PDF: ' + error.message);
        return { success: false, error: error.message };
    }
}

/**
 * ⭐ إنشاء كشف حساب للبوت (بدون UI) - نسخة نظيفة للـ PDF
 * يستخدم الدالة الموحدة renderStatementSheet_ من Utilities.js
 * بدون إيموجي وبدون عمود الأوردر لضمان التوافق مع PDF export
 */
function generateStatementForBot_(ss, partyName, partyType) {
    return renderStatementSheet_({
        ss: ss,
        partyName: partyName,
        partyType: partyType,
        useEmoji: false,
        includeOrderCol: false
    });
}

// ═══════════════════════════════════════════════════════════
// ⭐ دالة موحدة لتوليد التقارير (بدلاً من 5 دوال متكررة)
// ═══════════════════════════════════════════════════════════

/**
 * خريطة إعدادات التقارير
 * كل تقرير: updateFn (دالة التحديث) + sheetName + saveToArchive
 */
var PDF_REPORT_REGISTRY_ = {
    alerts:        { updateFn: function() { updateAlerts(true); },                              sheetName: CONFIG.SHEETS.ALERTS,          save: false, errorMsg: 'لم يتم العثور على شيت التنبيهات' },
    balances:      { updateFn: function() { rebuildVendorSummaryReport(true); },                sheetName: CONFIG.SHEETS.VENDORS_REPORT,  save: false, errorMsg: 'لم يتم العثور على شيت تقرير الأرصدة' },
    profitability: { updateFn: function() { generateAllProjectsProfitabilityReport(true); },    sheetName: 'تقارير ربحية المشاريع',       save: true,  errorMsg: 'لم يتم إنشاء تقرير الربحية' },
    expenses:      { updateFn: function() { rebuildExpenseSummaryReport(true); },               sheetName: CONFIG.SHEETS.EXPENSES_REPORT, save: true,  errorMsg: 'لم يتم العثور على تقرير المصروفات' },
    revenues:      { updateFn: function() { rebuildRevenueSummaryReport(true); },               sheetName: CONFIG.SHEETS.REVENUE_REPORT,  save: true,  errorMsg: 'لم يتم العثور على تقرير الإيرادات' },
    dynamic_expenses: { updateFn: function() { generateDynamicExpenseReport(true); },           sheetName: CONFIG.SHEETS.DYNAMIC_EXPENSES, save: true, errorMsg: 'لم يتم العثور على تقرير تحليل المصروفات' },
    cashflow:         { updateFn: function() { rebuildCashFlowReport(true); },                 sheetName: CONFIG.SHEETS.CASHFLOW,         save: true, errorMsg: 'لم يتم العثور على تقرير التدفقات النقدية' },
    income_statement: { updateFn: function() { rebuildIncomeStatement(true); },                sheetName: CONFIG.SHEETS.INCOME_STATEMENT, save: true, errorMsg: 'لم يتم العثور على قائمة الدخل' },
    balance_sheet:    { updateFn: function() { rebuildBalanceSheet(true); },                   sheetName: CONFIG.SHEETS.BALANCE_SHEET,    save: true, errorMsg: 'لم يتم العثور على المركز المالي' },
    funders:          { updateFn: function() { rebuildFunderSummaryReport(true); },            sheetName: CONFIG.SHEETS.FUNDERS_REPORT,   save: true, errorMsg: 'لم يتم العثور على تقرير الممولين' },
    projects:         { updateFn: function() { rebuildProjectDetailReport(true); },            sheetName: CONFIG.SHEETS.PROJECT_REPORT,   save: true, errorMsg: 'لم يتم العثور على تقرير المشروعات' }
};

/**
 * ⭐ دالة موحدة لتوليد أي تقرير PDF وإرساله
 * @param {string} chatId - معرف المحادثة
 * @param {string} reportType - نوع التقرير (alerts, balances, profitability, expenses, revenues)
 * @returns {Object} - نتيجة العملية
 */
function generateReportPDF_(chatId, reportType) {
    try {
        var config = PDF_REPORT_REGISTRY_[reportType];
        if (!config) {
            return { success: false, error: 'نوع تقرير غير معروف: ' + reportType };
        }

        var ss = SpreadsheetApp.getActiveSpreadsheet();

        // تحديث البيانات
        config.updateFn();

        // البحث عن الشيت
        var sheet = ss.getSheetByName(config.sheetName);
        if (!sheet) {
            throw new Error(config.errorMsg);
        }

        // تصدير وإرسال
        return generateAndSendReport(chatId, reportType, sheet, null, config.save);

    } catch (error) {
        Logger.log('❌ Error generating ' + reportType + ' PDF: ' + error.message);
        return { success: false, error: error.message };
    }
}

// ⭐ الدوال العامة تستدعي الدالة الموحدة (للتوافق مع الكود الحالي)
function generateAlertsPDF(chatId)        { return generateReportPDF_(chatId, 'alerts'); }
function generateBalancesPDF(chatId)      { return generateReportPDF_(chatId, 'balances'); }
function generateProfitabilityPDF(chatId) { return generateReportPDF_(chatId, 'profitability'); }
function generateExpensesPDF(chatId)      { return generateReportPDF_(chatId, 'expenses'); }
function generateRevenuesPDF(chatId)      { return generateReportPDF_(chatId, 'revenues'); }
function generateDynamicExpensesPDF(chatId) { return generateReportPDF_(chatId, 'dynamic_expenses'); }
function generateCashflowPDF(chatId)        { return generateReportPDF_(chatId, 'cashflow'); }
function generateIncomeStatementPDF(chatId)  { return generateReportPDF_(chatId, 'income_statement'); }
function generateBalanceSheetPDF(chatId)     { return generateReportPDF_(chatId, 'balance_sheet'); }
function generateFundersPDF(chatId)          { return generateReportPDF_(chatId, 'funders'); }
function generateProjectsPDF(chatId)         { return generateReportPDF_(chatId, 'projects'); }
