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
            portrait: false,      // أفقي
            fitWidth: true,       // ملائمة العرض
            gridlines: false,     // بدون خطوط الشبكة
            printtitle: true,     // طباعة العنوان
            sheetnames: false,    // بدون اسم الشيت
            pagenumbers: true,    // أرقام الصفحات
            fzr: true,            // تكرار الصفوف المجمدة
            fzc: true             // تكرار الأعمدة المجمدة
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
        const monthFolderName = Utilities.formatDate(now, 'Asia/Istanbul', 'yyyy-MM');

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
        const dateStr = Utilities.formatDate(now, 'Asia/Istanbul', 'yyyy-MM-dd');

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
        const token = CONFIG.TELEGRAM_BOT.AI_BOT_TOKEN;
        const url = 'https://api.telegram.org/bot' + token + '/sendDocument';

        // إعداد البيانات للإرسال
        const formData = {
            'method': 'post',
            'payload': {
                'chat_id': chatId,
                'document': pdfBlob,
                'caption': caption || '',
                'parse_mode': 'Markdown'
            },
            'muteHttpExceptions': true
        };

        const response = UrlFetchApp.fetch(url, formData);
        const result = JSON.parse(response.getContentText());

        if (result.ok) {
            Logger.log('✅ PDF sent to Telegram successfully');
            return true;
        } else {
            Logger.log('❌ Telegram API error: ' + result.description);
            return false;
        }

    } catch (error) {
        Logger.log('❌ Error sending PDF to Telegram: ' + error.message);
        return false;
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
            portrait: false,
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
function buildReportFileName(reportType, partyName) {
    const reportNames = {
        'statement': 'كشف حساب',
        'alerts': 'تنبيهات الاستحقاق',
        'balances': 'تقرير الأرصدة',
        'profitability': 'ربحية المشاريع',
        'expenses': 'تقرير المصروفات',
        'revenues': 'تقرير الإيرادات'
    };

    let name = reportNames[reportType] || reportType;
    if (partyName) {
        name += ' - ' + partyName;
    }

    const now = new Date();
    const dateStr = Utilities.formatDate(now, 'Asia/Istanbul', 'yyyy-MM-dd');
    name += ' - ' + dateStr;

    return name;
}

/**
 * بناء تعليق التقرير
 */
function buildReportCaption(reportType, partyName, savedToArchive) {
    const reportNames = {
        'statement': '📄 كشف حساب',
        'alerts': '⏰ تنبيهات الاستحقاق',
        'balances': '💰 تقرير الأرصدة',
        'profitability': '📈 ربحية المشاريع',
        'expenses': '📊 تقرير المصروفات',
        'revenues': '💵 تقرير الإيرادات'
    };

    let caption = '*' + (reportNames[reportType] || reportType) + '*';

    if (partyName) {
        caption += '\n👤 ' + partyName;
    }

    const now = new Date();
    const dateStr = Utilities.formatDate(now, 'Asia/Istanbul', 'dd/MM/yyyy HH:mm');
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

        // استخدام الدالة الموجودة لتوليد كشف الحساب
        generateUnifiedStatement_(ss, partyName, partyType);

        // البحث عن الشيت المُنشأ
        const sheetName = 'كشف حساب - ' + partyName;
        const sheet = ss.getSheetByName(sheetName);

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
 * توليد تقرير التنبيهات وإرساله
 * @param {string} chatId - معرف المحادثة
 * @param {number} daysAhead - عدد الأيام للتنبيهات
 * @returns {Object} - نتيجة العملية
 */
function generateAlertsPDF(chatId, daysAhead) {
    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();

        // تحديث شيت التنبيهات
        updateAlerts();

        // البحث عن شيت التنبيهات
        const sheet = ss.getSheetByName(CONFIG.SHEETS.ALERTS);

        if (!sheet) {
            throw new Error('لم يتم العثور على شيت التنبيهات');
        }

        // تصدير وإرسال (بدون حفظ في الأرشيف)
        const result = generateAndSendReport(chatId, 'alerts', sheet, null, false);

        return result;

    } catch (error) {
        Logger.log('❌ Error generating alerts PDF: ' + error.message);
        return { success: false, error: error.message };
    }
}

/**
 * توليد تقرير الأرصدة وإرساله
 * @param {string} chatId - معرف المحادثة
 * @returns {Object} - نتيجة العملية
 */
function generateBalancesPDF(chatId) {
    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();

        // تحديث تقرير الموردين
        rebuildVendorSummaryReport(true);

        // البحث عن شيت تقرير الموردين
        const sheet = ss.getSheetByName(CONFIG.SHEETS.VENDOR_REPORT);

        if (!sheet) {
            throw new Error('لم يتم العثور على شيت تقرير الأرصدة');
        }

        // تصدير وإرسال (بدون حفظ في الأرشيف)
        const result = generateAndSendReport(chatId, 'balances', sheet, null, false);

        return result;

    } catch (error) {
        Logger.log('❌ Error generating balances PDF: ' + error.message);
        return { success: false, error: error.message };
    }
}

/**
 * توليد تقرير ربحية المشاريع وإرساله
 * @param {string} chatId - معرف المحادثة
 * @returns {Object} - نتيجة العملية
 */
function generateProfitabilityPDF(chatId) {
    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();

        // توليد تقرير الربحية
        generateAllProjectsProfitabilityReport();

        // البحث عن الشيت
        const sheet = ss.getSheetByName('تقرير ربحية المشاريع');

        if (!sheet) {
            throw new Error('لم يتم إنشاء تقرير الربحية');
        }

        // تصدير وإرسال (مع حفظ في الأرشيف)
        const result = generateAndSendReport(chatId, 'profitability', sheet, null, true);

        return result;

    } catch (error) {
        Logger.log('❌ Error generating profitability PDF: ' + error.message);
        return { success: false, error: error.message };
    }
}

/**
 * توليد تقرير المصروفات وإرساله
 * @param {string} chatId - معرف المحادثة
 * @returns {Object} - نتيجة العملية
 */
function generateExpensesPDF(chatId) {
    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();

        // تحديث تقرير المصروفات
        rebuildExpenseSummaryReport(true);

        // البحث عن الشيت
        const sheet = ss.getSheetByName(CONFIG.SHEETS.EXPENSE_REPORT);

        if (!sheet) {
            throw new Error('لم يتم العثور على تقرير المصروفات');
        }

        // تصدير وإرسال (مع حفظ في الأرشيف)
        const result = generateAndSendReport(chatId, 'expenses', sheet, null, true);

        return result;

    } catch (error) {
        Logger.log('❌ Error generating expenses PDF: ' + error.message);
        return { success: false, error: error.message };
    }
}

/**
 * توليد تقرير الإيرادات وإرساله
 * @param {string} chatId - معرف المحادثة
 * @returns {Object} - نتيجة العملية
 */
function generateRevenuesPDF(chatId) {
    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();

        // تحديث تقرير الإيرادات
        rebuildRevenueSummaryReport(true);

        // البحث عن الشيت
        const sheet = ss.getSheetByName(CONFIG.SHEETS.REVENUE_REPORT);

        if (!sheet) {
            throw new Error('لم يتم العثور على تقرير الإيرادات');
        }

        // تصدير وإرسال (مع حفظ في الأرشيف)
        const result = generateAndSendReport(chatId, 'revenues', sheet, null, true);

        return result;

    } catch (error) {
        Logger.log('❌ Error generating revenues PDF: ' + error.message);
        return { success: false, error: error.message };
    }
}
