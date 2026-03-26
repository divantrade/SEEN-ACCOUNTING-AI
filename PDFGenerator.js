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
    'revenues': { name: 'تقرير الإيرادات', emoji: '💵' }
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
 * ═══════════════════════════════════════════════════════════════════════════
 * بدون إيموجي أو لوجو لضمان التوافق مع PDF export
 * ═══════════════════════════════════════════════════════════════════════════
 */
function generateStatementForBot_(ss, partyName, partyType) {
    const transSheet = ss.getSheetByName(CONFIG.SHEETS.TRANSACTIONS);

    if (!transSheet) {
        throw new Error('شيت دفتر الحركات المالية غير موجود!');
    }

    // ═══════════════════════════════════════════════════════════
    // جلب بيانات الطرف
    // ═══════════════════════════════════════════════════════════
    const partyData = getPartyData_(ss, partyName, partyType);

    // ═══════════════════════════════════════════════════════════
    // جلب بيانات لوجو الشركة من قاعدة بيانات البنود (D2)
    // ═══════════════════════════════════════════════════════════
    let logoBlob = null;
    let logoFileId = '';
    let logoOriginalUrl = '';
    let hasCellImage = false;
    let logoSourceRange = null;
    try {
        const itemsSheet = ss.getSheetByName(CONFIG.SHEETS.ITEMS || 'قاعدة بيانات البنود');
        if (itemsSheet) {
            const d2Range = itemsSheet.getRange('D2');
            const d2Value = d2Range.getValue();
            const d2Type = typeof d2Value;
            let logoUrl = '';

            // الحالة 1: CellImage (صورة مدرجة داخل الخلية)
            if (d2Value && d2Type === 'object') {
                hasCellImage = true;
                logoSourceRange = d2Range;
                try {
                    if (typeof d2Value.getContentUrl === 'function') {
                        logoUrl = d2Value.getContentUrl() || '';
                    }
                    if (!logoUrl && typeof d2Value.getUrl === 'function') {
                        logoUrl = d2Value.getUrl() || '';
                    }
                } catch (imgErr) {
                    Logger.log('🖼️ CellImage read error: ' + imgErr.message);
                }
            }

            // الحالة 2: نص عادي (URL مباشر)
            if (!logoUrl && d2Value && d2Type === 'string') {
                logoUrl = String(d2Value).trim();
            }

            // الحالة 3: معادلة IMAGE
            if (!logoUrl) {
                const formula = d2Range.getFormula() || '';
                const formulaMatch = formula.match(/IMAGE\s*\(\s*"([^"]+)"/i);
                if (formulaMatch) {
                    logoUrl = formulaMatch[1];
                }
            }

            // الحالة 4: صورة عائمة فوق الخلايا (OverGridImage)
            if (!logoUrl && !logoBlob) {
                try {
                    const images = itemsSheet.getImages();
                    if (images.length > 0) {
                        logoBlob = images[0].getBlob();
                    }
                } catch (imgErr) {
                    Logger.log('🖼️ getImages error: ' + imgErr.message);
                }
            }

            // الحالة 5: البحث عن رابط اللوجو في الخلايا المجاورة (احتياطي)
            if (!logoUrl && !logoBlob) {
                try {
                    const lastCol = Math.min(itemsSheet.getLastColumn(), 10);
                    if (lastCol >= 5) {
                        const searchRange = itemsSheet.getRange(1, 5, Math.min(itemsSheet.getLastRow(), 5), lastCol - 4);
                        const searchValues = searchRange.getValues();
                        for (let r = 0; r < searchValues.length && !logoUrl; r++) {
                            for (let c = 0; c < searchValues[r].length && !logoUrl; c++) {
                                const val = String(searchValues[r][c] || '').trim();
                                if (val && (val.includes('drive.google.com/file/d/') || val.includes('googleusercontent.com/d/'))) {
                                    logoUrl = val;
                                }
                            }
                        }
                    }
                } catch (searchErr) {
                    Logger.log('🖼️ Nearby cell search error: ' + searchErr.message);
                }
            }

            logoOriginalUrl = logoUrl;

            // استخراج File ID من URL
            if (logoUrl) {
                const m = logoUrl.match(/\/file\/d\/([a-zA-Z0-9_-]+)/) ||
                          logoUrl.match(/[?&]id=([a-zA-Z0-9_-]+)/) ||
                          logoUrl.match(/googleusercontent\.com\/d\/([a-zA-Z0-9_-]+)/) ||
                          logoUrl.match(/\/d\/([a-zA-Z0-9_-]{10,})/);
                if (m) logoFileId = m[1];
            }
        }
    } catch (e) {
        Logger.log('⚠️ Logo extraction error: ' + e.message);
    }

    // ═══════════════════════════════════════════════════════════
    // تحديد عنوان الكشف ولون التبويب حسب نوع الطرف
    // ═══════════════════════════════════════════════════════════
    let titlePrefix = 'كشف حساب';
    let tabColor = '#4a86e8';

    if (partyType === 'مورد') {
        titlePrefix = 'كشف مورد';
        tabColor = CONFIG.COLORS.TAB.VENDOR_STATEMENT || '#00897b';
    } else if (partyType === 'عميل') {
        titlePrefix = 'كشف عميل';
        tabColor = CONFIG.COLORS.TAB.CLIENT_STATEMENT || '#4caf50';
    } else if (partyType === 'ممول') {
        titlePrefix = 'كشف ممول';
        tabColor = CONFIG.COLORS.TAB.FUNDER_STATEMENT || '#ff9800';
    }

    // ═══════════════════════════════════════════════════════════
    // إنشاء شيت جديد (حذف القديم إن وجد)
    // ═══════════════════════════════════════════════════════════
    const sheetName = titlePrefix + ' - ' + partyName;
    let sheet = ss.getSheetByName(sheetName);

    if (sheet) {
        ss.deleteSheet(sheet);
    }

    sheet = ss.insertSheet(sheetName);
    sheet.setTabColor(tabColor);
    sheet.setRightToLeft(true);

    // ═══════════════════════════════════════════════════════════
    // عرض الأعمدة (8 أعمدة - مع تاريخ الاستحقاق + عمود اللوجو)
    // ═══════════════════════════════════════════════════════════
    sheet.setColumnWidth(1, 100);  // تاريخ التسجيل
    sheet.setColumnWidth(2, 100);  // تاريخ الاستحقاق
    sheet.setColumnWidth(3, 140);  // المشروع
    sheet.setColumnWidth(4, 210);  // التفاصيل
    sheet.setColumnWidth(5, 110);  // مدين
    sheet.setColumnWidth(6, 110);  // دائن
    sheet.setColumnWidth(7, 110);  // الرصيد
    sheet.setColumnWidth(8, 80);   // عمود اللوجو

    // ═══════════════════════════════════════════════════════════
    // Header الشركة في أعلى التقرير
    // ═══════════════════════════════════════════════════════════
    // ارتفاع الصفوف لمنطقة اللوجو
    sheet.setRowHeight(1, 45);
    sheet.setRowHeight(2, 35);
    sheet.setRowHeight(3, 30);

    // خلفية رمادية فاتحة للترويسة
    sheet.getRange('A1:H3').setBackground('#f8f9fa');

    // صف 1: اسم الشركة (بولد، خط كبير)
    sheet.getRange('A1:F1').merge();
    sheet.getRange('A1')
        .setValue(CONFIG.COMPANY.NAME)
        .setFontWeight('bold')
        .setFontSize(14)
        .setFontColor('#1a237e')
        .setHorizontalAlignment('center')
        .setVerticalAlignment('bottom');

    // صف 2: العنوان
    sheet.getRange('A2:F2').merge();
    sheet.getRange('A2')
        .setValue(CONFIG.COMPANY.ADDRESS)
        .setFontSize(10)
        .setFontColor('#555555')
        .setHorizontalAlignment('center')
        .setVerticalAlignment('middle');

    // صف 3: البريد والموقع
    sheet.getRange('A3:F3').merge();
    sheet.getRange('A3')
        .setValue(CONFIG.COMPANY.CONTACT)
        .setFontSize(10)
        .setFontColor('#555555')
        .setHorizontalAlignment('center')
        .setVerticalAlignment('top');

    // خط فاصل سفلي للترويسة
    sheet.getRange('A3:H3').setBorder(
        false, false, true, false, false, false,
        '#1a237e', SpreadsheetApp.BorderStyle.SOLID_MEDIUM
    );

    // ═══════════════════════════════════════════════════════════
    // إضافة اللوجو في G1:H3 مدمجة
    // ═══════════════════════════════════════════════════════════
    let logoInserted = false;

    function ensureLogoMerge_() {
        sheet.getRange('G1:H3').merge()
            .setBackground('#f8f9fa')
            .setHorizontalAlignment('center')
            .setVerticalAlignment('middle');
    }

    if (hasCellImage || logoBlob || logoFileId || logoOriginalUrl) {

        // الطريقة الأولى: نسخ CellImage مباشرة من خلية المصدر
        if (hasCellImage && logoSourceRange && !logoInserted) {
            try {
                logoSourceRange.copyTo(sheet.getRange('G1'), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
                ensureLogoMerge_();
                logoInserted = true;
            } catch (e) {
                Logger.log('⚠️ Method CellCopy FAILED: ' + e.message);
            }
        }

        // الطريقة 0: blob مباشر (من OverGridImage)
        if (logoBlob && !logoInserted) {
            try {
                ensureLogoMerge_();
                const image = sheet.insertImage(logoBlob, 7, 1);
                image.setWidth(CONFIG.COMPANY.LOGO.WIDTH);
                image.setHeight(CONFIG.COMPANY.LOGO.HEIGHT);
                logoInserted = true;
            } catch (e) {
                Logger.log('⚠️ Method 0 FAILED: ' + e.message);
            }
        }

        // الطريقة 1: DriveApp
        if (logoFileId && !logoInserted) {
            try {
                ensureLogoMerge_();
                const file = DriveApp.getFileById(logoFileId);
                const blob = file.getBlob();
                const image = sheet.insertImage(blob, 7, 1);
                image.setWidth(CONFIG.COMPANY.LOGO.WIDTH);
                image.setHeight(CONFIG.COMPANY.LOGO.HEIGHT);
                logoInserted = true;
            } catch (e) {
                Logger.log('⚠️ Method 1 FAILED: ' + e.message);
            }
        }

        // الطريقة 2a: UrlFetchApp + رابط lh3
        if (logoFileId && !logoInserted) {
            try {
                const lh3Url = 'https://lh3.googleusercontent.com/d/' + logoFileId;
                const response = UrlFetchApp.fetch(lh3Url, { muteHttpExceptions: true, followRedirects: true });
                if (response.getResponseCode() === 200) {
                    ensureLogoMerge_();
                    const image = sheet.insertImage(response.getBlob(), 7, 1);
                    image.setWidth(140);
                    image.setHeight(100);
                    logoInserted = true;
                }
            } catch (e) {
                Logger.log('⚠️ Method 2a FAILED: ' + e.message);
            }
        }

        // الطريقة 2b: UrlFetchApp + رابط uc?export=view
        if (logoFileId && !logoInserted) {
            try {
                const directUrl = 'https://drive.google.com/uc?export=view&id=' + logoFileId;
                const response = UrlFetchApp.fetch(directUrl, { muteHttpExceptions: true, followRedirects: true });
                if (response.getResponseCode() === 200) {
                    ensureLogoMerge_();
                    const image = sheet.insertImage(response.getBlob(), 7, 1);
                    image.setWidth(140);
                    image.setHeight(100);
                    logoInserted = true;
                }
            } catch (e) {
                Logger.log('⚠️ Method 2b FAILED: ' + e.message);
            }
        }

        // الطريقة 3: IMAGE formula
        if (!logoInserted) {
            try {
                const imgUrl = logoFileId
                    ? 'https://lh3.googleusercontent.com/d/' + logoFileId
                    : logoOriginalUrl;
                if (imgUrl) {
                    ensureLogoMerge_();
                    sheet.getRange('G1').setFormula('=IMAGE("' + imgUrl + '", 2)');
                    logoInserted = true;
                }
            } catch (e) {
                Logger.log('⚠️ Method 3 FAILED: ' + e.message);
            }
        }
    }

    // ضمان دمج G1:H3 حتى لو مفيش لوجو
    if (!logoInserted) {
        ensureLogoMerge_();
    }

    // صف 4: فاصل
    sheet.setRowHeight(4, 10);

    // ═══════════════════════════════════════════════════════════
    // العنوان الرئيسي (كشف مورد/عميل/ممول)
    // ═══════════════════════════════════════════════════════════
    sheet.getRange('A5:H5').merge();
    sheet.getRange('A5')
        .setValue(titlePrefix)
        .setBackground(CONFIG.COLORS.HEADER.DASHBOARD)
        .setFontColor(CONFIG.COLORS.TEXT.WHITE)
        .setFontWeight('bold')
        .setFontSize(15)
        .setHorizontalAlignment('center')
        .setVerticalAlignment('middle');

    // ═══════════════════════════════════════════════════════════
    // كارت بيانات الطرف
    // ═══════════════════════════════════════════════════════════
    const cardHeaderRow = 7;
    const cardDataStartRow = cardHeaderRow + 1;

    sheet.getRange('A' + cardHeaderRow + ':H' + cardHeaderRow).merge()
        .setValue('بيانات ' + partyType)
        .setBackground(CONFIG.COLORS.HEADER.SUMMARY)
        .setFontColor(CONFIG.COLORS.TEXT.WHITE)
        .setFontWeight('bold')
        .setHorizontalAlignment('center');

    sheet.getRange('A' + cardDataStartRow + ':H' + (cardDataStartRow + 3)).setBackground(CONFIG.COLORS.BG.LIGHT_BLUE);

    sheet.getRange('A' + cardDataStartRow).setValue('الاسم:').setFontWeight('bold');
    sheet.getRange('B' + cardDataStartRow + ':C' + cardDataStartRow).merge().setValue(partyName);

    sheet.getRange('E' + cardDataStartRow).setValue('التخصص:').setFontWeight('bold');
    sheet.getRange('F' + cardDataStartRow + ':H' + cardDataStartRow).merge().setValue(partyData.specialization || '');

    sheet.getRange('A' + (cardDataStartRow + 1)).setValue('الهاتف:').setFontWeight('bold');
    sheet.getRange('B' + (cardDataStartRow + 1) + ':C' + (cardDataStartRow + 1)).merge().setValue(partyData.phone || '');

    sheet.getRange('E' + (cardDataStartRow + 1)).setValue('البريد:').setFontWeight('bold');
    sheet.getRange('F' + (cardDataStartRow + 1) + ':H' + (cardDataStartRow + 1)).merge().setValue(partyData.email || '');

    sheet.getRange('A' + (cardDataStartRow + 2)).setValue('البنك:').setFontWeight('bold');
    sheet.getRange('B' + (cardDataStartRow + 2) + ':H' + (cardDataStartRow + 2)).merge().setValue(partyData.bankInfo || '');

    sheet.getRange('A' + (cardDataStartRow + 3)).setValue('ملاحظات:').setFontWeight('bold');
    sheet.getRange('B' + (cardDataStartRow + 3) + ':H' + (cardDataStartRow + 3)).merge().setValue(partyData.notes || '').setWrap(true);

    sheet.getRange('A' + cardDataStartRow + ':H' + (cardDataStartRow + 3)).setBorder(
        true, true, true, true, true, true,
        '#1565c0', SpreadsheetApp.BorderStyle.SOLID
    );

    // ═══════════════════════════════════════════════════════════
    // استخراج حركات الطرف (بدون فلتر طبيعة الحركة)
    // ═══════════════════════════════════════════════════════════
    const data = transSheet.getDataRange().getValues();
    const rows = [];

    // ⭐ أرقام الأعمدة كثوابت (بدلاً من magic numbers)
    var COL = { DATE: 1, PROJECT: 5, DETAILS: 7, PARTY: 8, AMOUNT_USD: 12, MOVEMENT: 13, DUE_DATE: 20 };

    let totalDebit = 0, totalCredit = 0, balance = 0;

    for (let i = 1; i < data.length; i++) {
        const row = data[i];

        // الفلتر الوحيد: اسم الطرف
        if (row[COL.PARTY] !== partyName) continue;

        const movementKind = String(row[COL.MOVEMENT] || '');
        const amountUsd = Number(row[COL.AMOUNT_USD]) || 0;

        // تجاهل الحركات بدون مبلغ
        if (!amountUsd) continue;

        const date = row[COL.DATE];
        const dueDate = row[COL.DUE_DATE];
        const project = row[COL.PROJECT];
        const details = row[COL.DETAILS];

        let debit = 0, credit = 0;

        // استخدام includes للتعامل مع الإيموجي
        if (movementKind.includes(CONFIG.MOVEMENT.DEBIT) || movementKind.includes('مدين')) {
            debit = amountUsd;
            totalDebit += debit;
        } else if (movementKind.includes(CONFIG.MOVEMENT.CREDIT) || movementKind.includes('دائن')) {
            credit = amountUsd;
            totalCredit += credit;
        }

        rows.push([
            date,
            (debit > 0 && dueDate) ? dueDate : '',
            project || '',
            details || '',
            debit || '',
            credit || '',
            0  // الرصيد يُحسب بعد الترتيب
        ]);
    }

    // ترتيب زمني
    rows.sort((a, b) => {
        const dateA = a[0] instanceof Date ? a[0].getTime() : new Date(a[0]).getTime();
        const dateB = b[0] instanceof Date ? b[0].getTime() : new Date(b[0]).getTime();
        return dateA - dateB;
    });

    // ⭐ حساب الرصيد مرة واحدة فقط (بعد الترتيب)
    balance = 0;
    for (let i = 0; i < rows.length; i++) {
        const debit = rows[i][4] || 0;
        const credit = rows[i][5] || 0;
        balance += debit - credit;
        rows[i][6] = Math.round(balance * 100) / 100;
    }

    // ═══════════════════════════════════════════════════════════
    // الملخص المالي - ⭐ استخدام ألوان CONFIG الموحدة
    // ═══════════════════════════════════════════════════════════
    const summaryHeaderRow = cardDataStartRow + 5;
    const summaryDataStartRow = summaryHeaderRow + 1;

    sheet.getRange('A' + summaryHeaderRow + ':H' + summaryHeaderRow).merge()
        .setValue('الملخص المالي')
        .setBackground(CONFIG.COLORS.HEADER.SUMMARY)
        .setFontColor(CONFIG.COLORS.TEXT.WHITE)
        .setFontWeight('bold')
        .setHorizontalAlignment('center');

    sheet.getRange('A' + summaryDataStartRow + ':H' + (summaryDataStartRow + 1)).setBackground(CONFIG.COLORS.BG.LIGHT_BLUE);

    sheet.getRange('A' + summaryDataStartRow).setValue('إجمالي المدين:').setFontWeight('bold');
    sheet.getRange('B' + summaryDataStartRow).setValue(totalDebit).setNumberFormat('$#,##0.00');

    sheet.getRange('E' + summaryDataStartRow).setValue('إجمالي الدائن:').setFontWeight('bold');
    sheet.getRange('F' + summaryDataStartRow).setValue(totalCredit).setNumberFormat('$#,##0.00');

    sheet.getRange('A' + (summaryDataStartRow + 1)).setValue('الرصيد الحالي:').setFontWeight('bold');
    sheet.getRange('B' + (summaryDataStartRow + 1)).setValue(balance).setNumberFormat('$#,##0.00')
        .setFontWeight('bold')
        .setBackground(balance > 0 ? '#ffcdd2' : '#c8e6c9');

    sheet.getRange('E' + (summaryDataStartRow + 1)).setValue('عدد الحركات:').setFontWeight('bold');
    sheet.getRange('F' + (summaryDataStartRow + 1)).setValue(rows.length);

    sheet.getRange('A' + summaryDataStartRow + ':H' + (summaryDataStartRow + 1)).setBorder(
        true, true, true, true, true, true,
        '#1565c0', SpreadsheetApp.BorderStyle.SOLID
    );

    // ═══════════════════════════════════════════════════════════
    // رأس جدول الحركات (بدون إيموجي للتوافق مع PDF)
    // ═══════════════════════════════════════════════════════════
    const tableHeaderRow = summaryDataStartRow + 3;
    const headers = [
        'تاريخ التسجيل',
        'تاريخ الاستحقاق',
        'المشروع',
        'التفاصيل',
        'مدين (USD)',
        'دائن (USD)',
        'الرصيد (USD)'
    ];

    sheet.getRange(tableHeaderRow, 1, 1, headers.length)
        .setValues([headers])
        .setBackground(CONFIG.COLORS.HEADER.DASHBOARD)
        .setFontColor(CONFIG.COLORS.TEXT.WHITE)
        .setFontWeight('bold')
        .setHorizontalAlignment('center');

    // ═══════════════════════════════════════════════════════════
    // بيانات الحركات - ⭐ نفس التلوين المتناوب
    // ═══════════════════════════════════════════════════════════
    const dataStartRow = tableHeaderRow + 1;

    if (rows.length > 0) {
        sheet.getRange(dataStartRow, 1, rows.length, headers.length).setValues(rows);
        sheet.getRange(dataStartRow, 1, rows.length, 1).setNumberFormat('dd/mm/yyyy');  // تاريخ التسجيل
        sheet.getRange(dataStartRow, 2, rows.length, 1).setNumberFormat('dd/mm/yyyy');  // تاريخ الاستحقاق
        sheet.getRange(dataStartRow, 5, rows.length, 3).setNumberFormat('$#,##0.00');

        // تلوين متناوب للصفوف (أبيض و أزرق فاتح) - نفس Main.js
        for (let i = 0; i < rows.length; i++) {
            const r = dataStartRow + i;
            const bg = i % 2 === 0 ? '#ffffff' : CONFIG.COLORS.BG.LIGHT_BLUE;
            sheet.getRange(r, 1, 1, headers.length).setBackground(bg);
        }

        // إطار الجدول
        sheet.getRange(tableHeaderRow, 1, rows.length + 1, headers.length)
            .setBorder(true, true, true, true, true, true, '#bdbdbd', SpreadsheetApp.BorderStyle.SOLID);
    } else {
        sheet.getRange(dataStartRow, 1).setValue('لا توجد حركات').setFontStyle('italic');
    }

    // ═══════════════════════════════════════════════════════════
    // تذييل التقرير
    // ═══════════════════════════════════════════════════════════
    const footerRow = dataStartRow + Math.max(rows.length, 1) + 2;
    const reportDateStr = Utilities.formatDate(new Date(), CONFIG.COMPANY.TIMEZONE, 'dd/MM/yyyy HH:mm');
    sheet.getRange('A' + footerRow + ':H' + footerRow).merge()
        .setValue('تاريخ التقرير: ' + reportDateStr)
        .setHorizontalAlignment('center')
        .setFontSize(9)
        .setFontColor('#757575');

    const creditRow = footerRow + 1;
    sheet.getRange('A' + creditRow + ':H' + creditRow).merge()
        .setValue(CONFIG.COMPANY.CREDITS)
        .setHorizontalAlignment('center')
        .setFontSize(9)
        .setFontColor('#9e9e9e');

    // ⭐ مهم جداً: flush لضمان كتابة البيانات قبل التصدير
    SpreadsheetApp.flush();
    Logger.log('✅ Statement sheet created for: ' + partyName + ' with ' + rows.length + ' rows');

    return sheet;
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
    revenues:      { updateFn: function() { rebuildRevenueSummaryReport(true); },               sheetName: CONFIG.SHEETS.REVENUE_REPORT,  save: true,  errorMsg: 'لم يتم العثور على تقرير الإيرادات' }
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
