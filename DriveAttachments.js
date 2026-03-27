// ==================== إدارة المرفقات في Google Drive ====================
/**
 * ملف إدارة المرفقات
 * يتعامل مع حفظ الصور والملفات من تليجرام في Google Drive
 */

// ==================== دوال حفظ المرفقات ====================

/**
 * حفظ مرفق من تليجرام في Google Drive
 */
function saveAttachmentToDrive(fileId, fileName, session) {
    try {
        // الحصول على معلومات الملف من تليجرام
        const fileInfo = getFileInfo(fileId);

        if (!fileInfo.ok) {
            throw new Error('فشل الحصول على معلومات الملف');
        }

        const filePath = fileInfo.result.file_path;

        // تنزيل الملف
        const fileBlob = downloadTelegramFile(filePath);

        // إنشاء اسم الملف
        const timestamp = Utilities.formatDate(new Date(), CONFIG.COMPANY.TIMEZONE, 'yyyyMMdd_HHmmss');
        const extension = getFileExtension(fileName || filePath);
        const newFileName = `حركة_${timestamp}.${extension}`;

        // الحصول على/إنشاء مجلد الشهر
        const monthFolder = getOrCreateAttachmentsMonthFolder_();

        // حفظ الملف
        fileBlob.setName(newFileName);
        const savedFile = monthFolder.createFile(fileBlob);

        // إعداد صلاحيات المشاركة
        savedFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

        Logger.log('Saved attachment: ' + savedFile.getUrl());

        return savedFile.getUrl();

    } catch (error) {
        Logger.log('Error saving attachment: ' + error.message);
        throw error;
    }
}

/**
 * الحصول على أو إنشاء مجلد الشهر الحالي
 */
function getOrCreateAttachmentsMonthFolder_() {
    const parentFolderId = CONFIG.TELEGRAM_BOT.ATTACHMENTS_FOLDER_ID;
    const parentFolder = DriveApp.getFolderById(parentFolderId);

    // اسم مجلد الشهر
    const monthName = Utilities.formatDate(new Date(), CONFIG.COMPANY.TIMEZONE, 'yyyy-MM');

    // البحث عن المجلد
    const folders = parentFolder.getFoldersByName(monthName);

    if (folders.hasNext()) {
        return folders.next();
    }

    // إنشاء مجلد جديد
    const newFolder = parentFolder.createFolder(monthName);
    Logger.log('Created month folder: ' + monthName);

    return newFolder;
}

/**
 * استخراج امتداد الملف
 */
function getFileExtension(fileName) {
    if (!fileName) return 'jpg';

    const parts = fileName.split('.');
    if (parts.length > 1) {
        return parts[parts.length - 1].toLowerCase();
    }

    // تخمين الامتداد من نوع الملف
    if (fileName.includes('photo')) return 'jpg';
    if (fileName.includes('document')) return 'pdf';

    return 'jpg';
}

/**
 * التحقق من صلاحية نوع الملف
 */
function isValidFileType(mimeType) {
    const allowedTypes = BOT_CONFIG.ATTACHMENTS.ALLOWED_TYPES;
    return allowedTypes.includes(mimeType);
}

/**
 * التحقق من حجم الملف
 */
function isValidFileSize(fileSize) {
    const maxSize = BOT_CONFIG.ATTACHMENTS.MAX_SIZE_MB * 1024 * 1024;
    return fileSize <= maxSize;
}

// ==================== دوال إدارة المجلدات ====================

/**
 * إنشاء هيكل مجلدات المرفقات
 */
function setupAttachmentsFolders() {
    try {
        const parentFolderId = CONFIG.TELEGRAM_BOT.ATTACHMENTS_FOLDER_ID;
        const parentFolder = DriveApp.getFolderById(parentFolderId);

        // إنشاء مجلدات للأشهر الثلاثة القادمة
        const today = new Date();
        for (let i = 0; i <= 3; i++) {
            const date = new Date(today);
            date.setMonth(date.getMonth() + i);
            const monthName = Utilities.formatDate(date, CONFIG.COMPANY.TIMEZONE, 'yyyy-MM');

            const folders = parentFolder.getFoldersByName(monthName);
            if (!folders.hasNext()) {
                parentFolder.createFolder(monthName);
                Logger.log('Created folder: ' + monthName);
            }
        }

        SpreadsheetApp.getUi().alert(
            '✅ تم إعداد المجلدات',
            'تم إنشاء مجلدات المرفقات بنجاح',
            SpreadsheetApp.getUi().ButtonSet.OK
        );

    } catch (error) {
        Logger.log('Error setting up folders: ' + error.message);
        SpreadsheetApp.getUi().alert('❌ خطأ', error.message, SpreadsheetApp.getUi().ButtonSet.OK);
    }
}

/**
 * الحصول على قائمة المرفقات للحركة
 */
function getTransactionAttachments(transactionId) {
    const sheet = getBotTransactionsSheet();
    const columns = BOT_CONFIG.BOT_TRANSACTIONS_COLUMNS;
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
        if (data[i][columns.TRANSACTION_ID.index - 1] === transactionId) {
            const attachmentUrl = data[i][columns.ATTACHMENT_URL.index - 1];
            if (attachmentUrl) {
                return [{ url: attachmentUrl, type: 'attachment' }];
            }
        }
    }

    return [];
}

/**
 * حذف مرفق من Drive
 */
function deleteAttachment(fileUrl) {
    try {
        // استخراج معرف الملف من الرابط
        const fileId = extractFileIdFromUrl(fileUrl);
        if (fileId) {
            const file = DriveApp.getFileById(fileId);
            file.setTrashed(true);
            Logger.log('Deleted attachment: ' + fileId);
            return true;
        }
        return false;
    } catch (error) {
        Logger.log('Error deleting attachment: ' + error.message);
        return false;
    }
}

/**
 * استخراج معرف الملف من الرابط
 */
function extractFileIdFromUrl(url) {
    if (!url) return null;

    // نمط رابط Google Drive
    const patterns = [
        /\/d\/([a-zA-Z0-9_-]+)/,
        /id=([a-zA-Z0-9_-]+)/,
        /\/file\/d\/([a-zA-Z0-9_-]+)/
    ];

    for (const pattern of patterns) {
        const match = url.match(pattern);
        if (match && match[1]) {
            return match[1];
        }
    }

    return null;
}

// ==================== تقارير المرفقات ====================

/**
 * الحصول على إحصائيات المرفقات
 */
function getAttachmentsStatistics() {
    try {
        const parentFolderId = CONFIG.TELEGRAM_BOT.ATTACHMENTS_FOLDER_ID;
        const parentFolder = DriveApp.getFolderById(parentFolderId);

        let totalFiles = 0;
        let totalSize = 0;
        const monthStats = [];

        const subFolders = parentFolder.getFolders();
        while (subFolders.hasNext()) {
            const folder = subFolders.next();
            const files = folder.getFiles();
            let monthFiles = 0;
            let monthSize = 0;

            while (files.hasNext()) {
                const file = files.next();
                monthFiles++;
                monthSize += file.getSize();
            }

            monthStats.push({
                month: folder.getName(),
                files: monthFiles,
                size: monthSize
            });

            totalFiles += monthFiles;
            totalSize += monthSize;
        }

        return {
            totalFiles: totalFiles,
            totalSize: totalSize,
            sizeFormatted: formatFileSize(totalSize),
            months: monthStats.sort((a, b) => b.month.localeCompare(a.month))
        };

    } catch (error) {
        Logger.log('Error getting attachments statistics: ' + error.message);
        return null;
    }
}

/**
 * تنسيق حجم الملف
 */
function formatFileSize(bytes) {
    if (bytes === 0) return '0 Bytes';

    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));

    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}

/**
 * عرض تقرير المرفقات
 */
function showAttachmentsReport() {
    const stats = getAttachmentsStatistics();

    if (!stats) {
        SpreadsheetApp.getUi().alert('❌ خطأ', 'فشل تحميل الإحصائيات', SpreadsheetApp.getUi().ButtonSet.OK);
        return;
    }

    let message = '📎 تقرير المرفقات\n';
    message += '═'.repeat(25) + '\n\n';
    message += `📁 إجمالي الملفات: ${stats.totalFiles}\n`;
    message += `💾 الحجم الإجمالي: ${stats.sizeFormatted}\n\n`;

    if (stats.months.length > 0) {
        message += '📅 توزيع حسب الشهر:\n';
        message += '─'.repeat(20) + '\n';

        stats.months.slice(0, 6).forEach(m => {
            message += `${m.month}: ${m.files} ملفات (${formatFileSize(m.size)})\n`;
        });
    }

    SpreadsheetApp.getUi().alert('📎 تقرير المرفقات', message, SpreadsheetApp.getUi().ButtonSet.OK);
}

// ==================== تنظيف المرفقات القديمة ====================

/**
 * تنظيف المرفقات للحركات المرفوضة أو الملغاة
 */
function cleanupRejectedAttachments() {
    const ui = SpreadsheetApp.getUi();

    const result = ui.alert(
        '🗑️ تنظيف المرفقات',
        'سيتم حذف مرفقات الحركات المرفوضة.\n\nهل تريد المتابعة؟',
        ui.ButtonSet.YES_NO
    );

    if (result !== ui.Button.YES) return;

    try {
        const sheet = getBotTransactionsSheet();
        const columns = BOT_CONFIG.BOT_TRANSACTIONS_COLUMNS;
        const data = sheet.getDataRange().getValues();

        let deletedCount = 0;

        for (let i = 1; i < data.length; i++) {
            const status = data[i][columns.REVIEW_STATUS.index - 1];
            const attachmentUrl = data[i][columns.ATTACHMENT_URL.index - 1];

            if (status === CONFIG.TELEGRAM_BOT.REVIEW_STATUS.REJECTED && attachmentUrl) {
                if (deleteAttachment(attachmentUrl)) {
                    // مسح رابط المرفق من الشيت
                    sheet.getRange(i + 1, columns.ATTACHMENT_URL.index).setValue('');
                    deletedCount++;
                }
            }
        }

        ui.alert(
            '✅ تم التنظيف',
            `تم حذف ${deletedCount} مرفق`,
            ui.ButtonSet.OK
        );

    } catch (error) {
        ui.alert('❌ خطأ', error.message, ui.ButtonSet.OK);
    }
}

// ==================== اختبار المرفقات ====================

/**
 * اختبار الاتصال بمجلد المرفقات
 */
function testAttachmentsFolder() {
    try {
        const folderId = CONFIG.TELEGRAM_BOT.ATTACHMENTS_FOLDER_ID;
        const folder = DriveApp.getFolderById(folderId);

        const testBlob = Utilities.newBlob('Test content', 'text/plain', 'test.txt');
        const testFile = folder.createFile(testBlob);

        const url = testFile.getUrl();
        testFile.setTrashed(true);

        SpreadsheetApp.getUi().alert(
            '✅ اختبار ناجح',
            'تم الاتصال بمجلد المرفقات بنجاح!\n\nالمجلد: ' + folder.getName(),
            SpreadsheetApp.getUi().ButtonSet.OK
        );

        return true;

    } catch (error) {
        SpreadsheetApp.getUi().alert(
            '❌ فشل الاختبار',
            'لا يمكن الوصول لمجلد المرفقات:\n' + error.message,
            SpreadsheetApp.getUi().ButtonSet.OK
        );

        return false;
    }
}
