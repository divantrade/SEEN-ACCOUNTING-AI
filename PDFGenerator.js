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
 * ⭐ إنشاء كشف حساب للبوت (بدون UI) - بتنسيق مطابق للأصلي تماماً
 * نسخة كاملة من generateUnifiedStatement_ تعمل بدون SpreadsheetApp.getUi()
 * ═══════════════════════════════════════════════════════════════════════════
 * تم توحيد الألوان والتنسيق مع الدالة الأصلية في Main.js
 * ═══════════════════════════════════════════════════════════════════════════
 */
function generateStatementForBot_(ss, partyName, partyType) {
    const transSheet = ss.getSheetByName(CONFIG.SHEETS.TRANSACTIONS);

    if (!transSheet) {
        throw new Error('شيت دفتر الحركات المالية غير موجود!');
    }

    // ═══════════════════════════════════════════════════════════
    // جلب لوجو الشركة من قاعدة بيانات البنود (D2)
    // يدعم روابط Google Drive العادية ويحولها لروابط مباشرة
    // ═══════════════════════════════════════════════════════════
    let companyLogo = '';
    try {
        const itemsSheet = ss.getSheetByName(CONFIG.SHEETS.ITEMS || 'قاعدة بيانات البنود');
        if (itemsSheet) {
            let logoUrl = itemsSheet.getRange('D2').getValue() || '';

            // تحويل رابط Google Drive العادي إلى رابط مباشر
            // مثال: https://drive.google.com/file/d/FILE_ID/view?usp=drive_link
            // يتحول إلى: https://drive.google.com/uc?id=FILE_ID
            if (logoUrl && logoUrl.includes('drive.google.com/file/d/')) {
                const match = logoUrl.match(/\/file\/d\/([^\/\?]+)/);
                if (match && match[1]) {
                    logoUrl = 'https://drive.google.com/uc?id=' + match[1];
                }
            }

            // التحقق من أن الرابط صالح (ليس مجلد)
            if (logoUrl && !logoUrl.includes('/folders/')) {
                companyLogo = logoUrl;
            }
            Logger.log('🖼️ Company logo URL: ' + (companyLogo ? companyLogo : 'Not found'));
        }
    } catch (e) {
        Logger.log('⚠️ Could not get company logo: ' + e.message);
    }

    // ═══════════════════════════════════════════════════════════
    // جلب بيانات الطرف
    // ═══════════════════════════════════════════════════════════
    const partyData = getPartyData_(ss, partyName, partyType);

    // ═══════════════════════════════════════════════════════════
    // تحديد عنوان الكشف ولون التبويب حسب نوع الطرف
    // ⭐ استخدام ألوان CONFIG الموحدة
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
    // عرض الأعمدة (6 أعمدة بدون طبيعة الحركة والبند)
    // ═══════════════════════════════════════════════════════════
    sheet.setColumnWidth(1, 110);  // التاريخ
    sheet.setColumnWidth(2, 160);  // المشروع
    sheet.setColumnWidth(3, 250);  // التفاصيل
    sheet.setColumnWidth(4, 130);  // مدين
    sheet.setColumnWidth(5, 130);  // دائن
    sheet.setColumnWidth(6, 130);  // الرصيد

    // ═══════════════════════════════════════════════════════════
    // العنوان الرئيسي - ⭐ استخدام ألوان CONFIG الموحدة
    // ═══════════════════════════════════════════════════════════
    sheet.getRange('A1:F1').merge();
    sheet.getRange('A1')
        .setValue('📊 ' + titlePrefix)
        .setBackground(CONFIG.COLORS.HEADER.DASHBOARD)
        .setFontColor(CONFIG.COLORS.TEXT.WHITE)
        .setFontWeight('bold')
        .setFontSize(15)
        .setHorizontalAlignment('center')
        .setVerticalAlignment('middle');

    // ═══════════════════════════════════════════════════════════
    // إضافة اللوجو من Google Drive (باستخدام insertImage)
    // ═══════════════════════════════════════════════════════════
    let logoRowOffset = 0;
    if (companyLogo) {
        try {
            // استخراج File ID من الرابط
            let fileId = '';
            if (companyLogo.includes('uc?id=')) {
                fileId = companyLogo.split('uc?id=')[1].split('&')[0];
            } else if (companyLogo.includes('/file/d/')) {
                const match = companyLogo.match(/\/file\/d\/([^\/\?]+)/);
                if (match) fileId = match[1];
            }

            if (fileId) {
                // جلب الصورة من Drive وإدراجها
                const file = DriveApp.getFileById(fileId);
                const blob = file.getBlob();

                // إدراج الصورة في الصف 2
                sheet.setRowHeight(2, 80);
                const image = sheet.insertImage(blob, 3, 2); // العمود C، الصف 2
                image.setWidth(70);
                image.setHeight(70);

                logoRowOffset = 1;
                Logger.log('✅ Logo inserted from Drive: ' + fileId);
            } else {
                Logger.log('⚠️ Could not extract file ID from logo URL');
            }
        } catch (e) {
            Logger.log('⚠️ Could not insert logo: ' + e.message);
            logoRowOffset = 0;
        }
    }

    // ═══════════════════════════════════════════════════════════
    // كارت بيانات الطرف - ⭐ استخدام ألوان CONFIG الموحدة
    // ═══════════════════════════════════════════════════════════
    const cardHeaderRow = 3 + logoRowOffset;
    const cardDataStartRow = cardHeaderRow + 1;

    sheet.getRange('A' + cardHeaderRow + ':F' + cardHeaderRow).merge()
        .setValue('بيانات ' + partyType)
        .setBackground(CONFIG.COLORS.HEADER.SUMMARY)
        .setFontColor(CONFIG.COLORS.TEXT.WHITE)
        .setFontWeight('bold')
        .setHorizontalAlignment('center');

    sheet.getRange('A' + cardDataStartRow + ':F' + (cardDataStartRow + 3)).setBackground(CONFIG.COLORS.BG.LIGHT_BLUE);

    sheet.getRange('A' + cardDataStartRow).setValue('الاسم:').setFontWeight('bold');
    sheet.getRange('B' + cardDataStartRow + ':C' + cardDataStartRow).merge().setValue(partyName);

    sheet.getRange('D' + cardDataStartRow).setValue('التخصص:').setFontWeight('bold');
    sheet.getRange('E' + cardDataStartRow + ':F' + cardDataStartRow).merge().setValue(partyData.specialization || '');

    sheet.getRange('A' + (cardDataStartRow + 1)).setValue('الهاتف:').setFontWeight('bold');
    sheet.getRange('B' + (cardDataStartRow + 1) + ':C' + (cardDataStartRow + 1)).merge().setValue(partyData.phone || '');

    sheet.getRange('D' + (cardDataStartRow + 1)).setValue('البريد:').setFontWeight('bold');
    sheet.getRange('E' + (cardDataStartRow + 1) + ':F' + (cardDataStartRow + 1)).merge().setValue(partyData.email || '');

    sheet.getRange('A' + (cardDataStartRow + 2)).setValue('البنك:').setFontWeight('bold');
    sheet.getRange('B' + (cardDataStartRow + 2) + ':F' + (cardDataStartRow + 2)).merge().setValue(partyData.bankInfo || '');

    sheet.getRange('A' + (cardDataStartRow + 3)).setValue('ملاحظات:').setFontWeight('bold');
    sheet.getRange('B' + (cardDataStartRow + 3) + ':F' + (cardDataStartRow + 3)).merge().setValue(partyData.notes || '').setWrap(true);

    sheet.getRange('A' + cardDataStartRow + ':F' + (cardDataStartRow + 3)).setBorder(
        true, true, true, true, true, true,
        '#1565c0', SpreadsheetApp.BorderStyle.SOLID
    );

    // ═══════════════════════════════════════════════════════════
    // استخراج حركات الطرف (بدون فلتر طبيعة الحركة)
    // ═══════════════════════════════════════════════════════════
    const data = transSheet.getDataRange().getValues();
    const rows = [];

    let totalDebit = 0, totalCredit = 0, balance = 0;

    for (let i = 1; i < data.length; i++) {
        const row = data[i];

        // الفلتر الوحيد: اسم الطرف
        if (row[8] !== partyName) continue;

        const movementKind = String(row[13] || '');  // N: نوع الحركة
        const amountUsd = Number(row[12]) || 0;      // M: القيمة بالدولار

        // تجاهل الحركات بدون مبلغ
        if (!amountUsd) continue;

        const date = row[1];       // B: التاريخ
        const project = row[5];    // F: اسم المشروع
        const details = row[7];    // H: التفاصيل

        let debit = 0, credit = 0;

        // استخدام includes للتعامل مع الإيموجي
        if (movementKind.includes(CONFIG.MOVEMENT.DEBIT) || movementKind.includes('مدين')) {
            debit = amountUsd;
            balance += debit;
            totalDebit += debit;
        } else if (movementKind.includes(CONFIG.MOVEMENT.CREDIT) || movementKind.includes('دائن')) {
            credit = amountUsd;
            balance -= credit;
            totalCredit += credit;
        }

        rows.push([
            date,
            project || '',
            details || '',
            debit || '',
            credit || '',
            Math.round(balance * 100) / 100
        ]);
    }

    // ترتيب زمني
    rows.sort((a, b) => {
        const dateA = a[0] instanceof Date ? a[0].getTime() : new Date(a[0]).getTime();
        const dateB = b[0] instanceof Date ? b[0].getTime() : new Date(b[0]).getTime();
        return dateA - dateB;
    });

    // إعادة حساب الرصيد بعد الترتيب
    balance = 0;
    for (let i = 0; i < rows.length; i++) {
        const debit = rows[i][3] || 0;
        const credit = rows[i][4] || 0;
        balance += debit - credit;
        rows[i][5] = Math.round(balance * 100) / 100;
    }

    // ═══════════════════════════════════════════════════════════
    // الملخص المالي - ⭐ استخدام ألوان CONFIG الموحدة
    // ═══════════════════════════════════════════════════════════
    const summaryHeaderRow = cardDataStartRow + 5;
    const summaryDataStartRow = summaryHeaderRow + 1;

    sheet.getRange('A' + summaryHeaderRow + ':F' + summaryHeaderRow).merge()
        .setValue('الملخص المالي')
        .setBackground(CONFIG.COLORS.HEADER.SUMMARY)
        .setFontColor(CONFIG.COLORS.TEXT.WHITE)
        .setFontWeight('bold')
        .setHorizontalAlignment('center');

    sheet.getRange('A' + summaryDataStartRow + ':F' + (summaryDataStartRow + 1)).setBackground(CONFIG.COLORS.BG.LIGHT_BLUE);

    sheet.getRange('A' + summaryDataStartRow).setValue('إجمالي المدين:').setFontWeight('bold');
    sheet.getRange('B' + summaryDataStartRow).setValue(totalDebit).setNumberFormat('$#,##0.00');

    sheet.getRange('D' + summaryDataStartRow).setValue('إجمالي الدائن:').setFontWeight('bold');
    sheet.getRange('E' + summaryDataStartRow).setValue(totalCredit).setNumberFormat('$#,##0.00');

    sheet.getRange('A' + (summaryDataStartRow + 1)).setValue('الرصيد الحالي:').setFontWeight('bold');
    sheet.getRange('B' + (summaryDataStartRow + 1)).setValue(balance).setNumberFormat('$#,##0.00')
        .setFontWeight('bold')
        .setBackground(balance > 0 ? '#ffcdd2' : '#c8e6c9');

    sheet.getRange('D' + (summaryDataStartRow + 1)).setValue('عدد الحركات:').setFontWeight('bold');
    sheet.getRange('E' + (summaryDataStartRow + 1)).setValue(rows.length);

    sheet.getRange('A' + summaryDataStartRow + ':F' + (summaryDataStartRow + 1)).setBorder(
        true, true, true, true, true, true,
        '#1565c0', SpreadsheetApp.BorderStyle.SOLID
    );

    // ═══════════════════════════════════════════════════════════
    // رأس جدول الحركات - ⭐ نفس الإيموجي والألوان
    // ═══════════════════════════════════════════════════════════
    const tableHeaderRow = summaryDataStartRow + 3;
    const headers = [
        '📅 التاريخ',
        '🎬 المشروع',
        '📝 التفاصيل',
        '💰 مدين (USD)',
        '💸 دائن (USD)',
        '📊 الرصيد (USD)'
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
        sheet.getRange(dataStartRow, 1, rows.length, 1).setNumberFormat('dd/mm/yyyy');
        sheet.getRange(dataStartRow, 4, rows.length, 3).setNumberFormat('$#,##0.00');

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

    // تجميد الصفوف حتى رأس الجدول
    sheet.setFrozenRows(tableHeaderRow);

    // ═══════════════════════════════════════════════════════════
    // تذييل التقرير
    // ═══════════════════════════════════════════════════════════
    const footerRow = dataStartRow + Math.max(rows.length, 1) + 2;
    sheet.getRange('A' + footerRow + ':F' + footerRow).merge()
        .setValue('تاريخ التقرير: ' + Utilities.formatDate(new Date(), 'Asia/Istanbul', 'dd/MM/yyyy HH:mm') + ' | تم إنشاؤه بواسطة AI Bot')
        .setHorizontalAlignment('center')
        .setFontSize(9)
        .setFontColor('#757575');

    // ⭐ مهم جداً: flush لضمان كتابة البيانات قبل التصدير
    SpreadsheetApp.flush();
    Logger.log('✅ Statement sheet created for: ' + partyName + ' with ' + rows.length + ' rows');

    return sheet;
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
