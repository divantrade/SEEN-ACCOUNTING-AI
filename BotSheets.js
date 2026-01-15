// ==================== إنشاء وإدارة شيتات البوت ====================
/**
 * ملف إنشاء وإدارة شيتات بوت تليجرام
 * يحتوي على دوال إنشاء الشيتات الثلاثة الجديدة
 */

/**
 * إنشاء جميع شيتات البوت
 * يتم استدعاؤها من القائمة الرئيسية
 */
function setupBotSheets() {
    const ui = SpreadsheetApp.getUi();

    const result = ui.alert(
        '🤖 إعداد شيتات بوت تليجرام',
        'سيتم إنشاء الشيتات التالية:\n\n' +
        '1. حركات البوت (للحركات المعلقة)\n' +
        '2. أطراف البوت (للأطراف الجديدة المعلقة)\n' +
        '3. المستخدمين المصرح لهم\n\n' +
        'هل تريد المتابعة؟',
        ui.ButtonSet.YES_NO
    );

    if (result !== ui.Button.YES) {
        return;
    }

    try {
        // إنشاء الشيتات
        createBotTransactionsSheet();
        createBotPartiesSheet();
        createBotUsersSheet();

        ui.alert(
            '✅ تم بنجاح',
            'تم إنشاء جميع شيتات البوت بنجاح!\n\n' +
            'الخطوة التالية: قم بإعداد بوت تليجرام وإضافة أرقام الهواتف المصرح لها.',
            ui.ButtonSet.OK
        );

    } catch (error) {
        ui.alert('❌ خطأ', 'حدث خطأ: ' + error.message, ui.ButtonSet.OK);
        Logger.log('Error in setupBotSheets: ' + error.message);
    }
}

/**
 * ⭐ واجهة تحديث Data Validation (تُستدعى من القائمة)
 */
function updateBotSheetValidationUI() {
    const ui = SpreadsheetApp.getUi();
    const result = updateBotSheetValidation();

    if (result.success) {
        ui.alert(
            '✅ تم التحديث',
            'تم تحديث القوائم المنسدلة في شيت حركات البوت:\n\n' +
            '• طبيعة الحركة ← من عمود B في شيت البنود\n' +
            '• تصنيف الحركة ← من عمود C في شيت البنود\n' +
            '• البند ← من عمود A في شيت البنود',
            ui.ButtonSet.OK
        );
    } else {
        ui.alert('❌ خطأ', result.error, ui.ButtonSet.OK);
    }
}

/**
 * ⭐ تحديث Data Validation لشيت حركات البوت من شيت البنود
 * يُستدعى لتحديث الشيت الموجود بالقوائم المنسدلة
 */
function updateBotSheetValidation() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const botSheet = ss.getSheetByName(CONFIG.SHEETS.BOT_TRANSACTIONS);
    const itemsSheet = ss.getSheetByName(CONFIG.SHEETS.ITEMS);
    const columns = BOT_CONFIG.BOT_TRANSACTIONS_COLUMNS;

    if (!botSheet) {
        Logger.log('❌ شيت حركات البوت غير موجود');
        return { success: false, error: 'شيت حركات البوت غير موجود' };
    }

    if (!itemsSheet) {
        Logger.log('❌ شيت البنود غير موجود');
        return { success: false, error: 'شيت البنود غير موجود' };
    }

    try {
        const lastItemsRow = Math.max(itemsSheet.getLastRow(), 2);
        const lastBotRow = Math.max(botSheet.getLastRow(), CONFIG.SHEET.DEFAULT_ROWS);

        // طبيعة الحركة (من عمود B في شيت البنود)
        const natureRange = itemsSheet.getRange('B2:B' + lastItemsRow);
        const natureRule = SpreadsheetApp.newDataValidation()
            .requireValueInRange(natureRange, true)
            .setAllowInvalid(true)
            .setHelpText('اختر طبيعة الحركة من "قاعدة بيانات البنود"')
            .build();
        botSheet.getRange(2, columns.NATURE.index, lastBotRow, 1)
            .setDataValidation(natureRule);

        // تصنيف الحركة (من عمود C في شيت البنود)
        const classRange = itemsSheet.getRange('C2:C' + lastItemsRow);
        const classRule = SpreadsheetApp.newDataValidation()
            .requireValueInRange(classRange, true)
            .setAllowInvalid(true)
            .setHelpText('اختر تصنيف الحركة من "قاعدة بيانات البنود"')
            .build();
        botSheet.getRange(2, columns.CLASSIFICATION.index, lastBotRow, 1)
            .setDataValidation(classRule);

        // البند (من عمود A في شيت البنود)
        const itemRange = itemsSheet.getRange('A2:A' + lastItemsRow);
        const itemRule = SpreadsheetApp.newDataValidation()
            .requireValueInRange(itemRange, true)
            .setAllowInvalid(true)
            .setHelpText('اختر البند من "قاعدة بيانات البنود"')
            .build();
        botSheet.getRange(2, columns.ITEM.index, lastBotRow, 1)
            .setDataValidation(itemRule);

        Logger.log('✅ تم تحديث Data Validation لشيت حركات البوت');
        return { success: true };

    } catch (error) {
        Logger.log('❌ خطأ في تحديث Data Validation: ' + error.message);
        return { success: false, error: error.message };
    }
}

/**
 * إنشاء شيت حركات البوت
 */
function createBotTransactionsSheet() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetName = CONFIG.SHEETS.BOT_TRANSACTIONS;

    // التحقق من وجود الشيت
    let sheet = ss.getSheetByName(sheetName);
    if (sheet) {
        Logger.log('Sheet already exists: ' + sheetName);
        return sheet;
    }

    // إنشاء شيت جديد
    sheet = ss.insertSheet(sheetName);

    // الأعمدة من BOT_CONFIG
    const columns = BOT_CONFIG.BOT_TRANSACTIONS_COLUMNS;
    const headers = [];
    const widths = [];

    // جمع العناوين والعروض
    Object.values(columns).forEach(col => {
        headers[col.index - 1] = col.name;
        widths[col.index - 1] = col.width;
    });

    // إضافة صف العناوين
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setValues([headers]);

    // تنسيق الهيدر
    headerRange
        .setBackground(CONFIG.COLORS.BOT.HEADER)
        .setFontColor(CONFIG.COLORS.TEXT.WHITE)
        .setFontWeight('bold')
        .setFontSize(CONFIG.FONT.NORMAL)
        .setHorizontalAlignment('center')
        .setVerticalAlignment('middle')
        .setWrap(true);

    // تعيين عرض الأعمدة
    widths.forEach((width, index) => {
        sheet.setColumnWidth(index + 1, width);
    });

    // تجميد الصف الأول
    sheet.setFrozenRows(1);

    // إضافة Data Validation لحالة المراجعة
    const reviewStatusCol = columns.REVIEW_STATUS.index;
    const reviewStatusRule = SpreadsheetApp.newDataValidation()
        .requireValueInList([
            CONFIG.TELEGRAM_BOT.REVIEW_STATUS.PENDING,
            CONFIG.TELEGRAM_BOT.REVIEW_STATUS.APPROVED,
            CONFIG.TELEGRAM_BOT.REVIEW_STATUS.REJECTED,
            CONFIG.TELEGRAM_BOT.REVIEW_STATUS.NEEDS_EDIT
        ])
        .setAllowInvalid(false)
        .build();

    sheet.getRange(2, reviewStatusCol, CONFIG.SHEET.DEFAULT_ROWS, 1)
        .setDataValidation(reviewStatusRule);

    // إضافة التنسيق الشرطي لحالة المراجعة
    applyBotReviewConditionalFormatting(sheet, reviewStatusCol);

    // ✅ إضافة Data Validation لطبيعة الحركة وتصنيف الحركة من شيت البنود
    const itemsSheet = ss.getSheetByName(CONFIG.SHEETS.ITEMS);
    if (itemsSheet) {
        const lastItemsRow = Math.max(itemsSheet.getLastRow(), 2);

        // طبيعة الحركة (من عمود B في شيت البنود)
        const natureCol = columns.NATURE.index;
        const natureRange = itemsSheet.getRange('B2:B' + lastItemsRow);
        const natureRule = SpreadsheetApp.newDataValidation()
            .requireValueInRange(natureRange, true)
            .setAllowInvalid(true)
            .setHelpText('اختر طبيعة الحركة من "قاعدة بيانات البنود"')
            .build();
        sheet.getRange(2, natureCol, CONFIG.SHEET.DEFAULT_ROWS, 1)
            .setDataValidation(natureRule);

        // تصنيف الحركة (من عمود C في شيت البنود)
        const classificationCol = columns.CLASSIFICATION.index;
        const classRange = itemsSheet.getRange('C2:C' + lastItemsRow);
        const classRule = SpreadsheetApp.newDataValidation()
            .requireValueInRange(classRange, true)
            .setAllowInvalid(true)
            .setHelpText('اختر تصنيف الحركة من "قاعدة بيانات البنود"')
            .build();
        sheet.getRange(2, classificationCol, CONFIG.SHEET.DEFAULT_ROWS, 1)
            .setDataValidation(classRule);

        // البند (من عمود A في شيت البنود)
        const itemCol = columns.ITEM.index;
        const itemRange = itemsSheet.getRange('A2:A' + lastItemsRow);
        const itemRule = SpreadsheetApp.newDataValidation()
            .requireValueInRange(itemRange, true)
            .setAllowInvalid(true)
            .setHelpText('اختر البند من "قاعدة بيانات البنود"')
            .build();
        sheet.getRange(2, itemCol, CONFIG.SHEET.DEFAULT_ROWS, 1)
            .setDataValidation(itemRule);

        Logger.log('✅ تم ربط Data Validation مع شيت البنود');
    } else {
        // Fallback: استخدام القائمة الثابتة إذا لم يكن شيت البنود موجوداً
        Logger.log('⚠️ شيت البنود غير موجود، استخدام القائمة الثابتة');
        const natureCol = columns.NATURE.index;
        const natureRule = SpreadsheetApp.newDataValidation()
            .requireValueInList(CONFIG.NATURE_TYPES)
            .setAllowInvalid(true)
            .build();
        sheet.getRange(2, natureCol, CONFIG.SHEET.DEFAULT_ROWS, 1)
            .setDataValidation(natureRule);
    }

    // إضافة Data Validation للعملة
    const currencyCol = columns.CURRENCY.index;
    const currencyRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(CONFIG.CURRENCIES.LIST)
        .setAllowInvalid(false)
        .build();

    sheet.getRange(2, currencyCol, CONFIG.SHEET.DEFAULT_ROWS, 1)
        .setDataValidation(currencyRule);

    // تنسيق أعمدة التاريخ
    sheet.getRange(2, columns.DATE.index, CONFIG.SHEET.DEFAULT_ROWS, 1)
        .setNumberFormat('dd/mm/yyyy');
    sheet.getRange(2, columns.INPUT_TIMESTAMP.index, CONFIG.SHEET.DEFAULT_ROWS, 1)
        .setNumberFormat('dd/mm/yyyy hh:mm:ss');
    sheet.getRange(2, columns.REVIEW_TIMESTAMP.index, CONFIG.SHEET.DEFAULT_ROWS, 1)
        .setNumberFormat('dd/mm/yyyy hh:mm:ss');

    // تنسيق أعمدة الأرقام
    sheet.getRange(2, columns.AMOUNT.index, CONFIG.SHEET.DEFAULT_ROWS, 1)
        .setNumberFormat(CONFIG.FORMATS.CURRENCY);
    sheet.getRange(2, columns.AMOUNT_USD.index, CONFIG.SHEET.DEFAULT_ROWS, 1)
        .setNumberFormat(CONFIG.FORMATS.CURRENCY);
    sheet.getRange(2, columns.EXCHANGE_RATE.index, CONFIG.SHEET.DEFAULT_ROWS, 1)
        .setNumberFormat(CONFIG.FORMATS.RATE);

    // تحديد لون التبويب
    sheet.setTabColor(CONFIG.COLORS.BOT.HEADER);

    Logger.log('Created sheet: ' + sheetName);
    return sheet;
}

/**
 * إنشاء شيت أطراف البوت
 */
function createBotPartiesSheet() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetName = CONFIG.SHEETS.BOT_PARTIES;

    // التحقق من وجود الشيت
    let sheet = ss.getSheetByName(sheetName);
    if (sheet) {
        Logger.log('Sheet already exists: ' + sheetName);
        return sheet;
    }

    // إنشاء شيت جديد
    sheet = ss.insertSheet(sheetName);

    // الأعمدة من BOT_CONFIG
    const columns = BOT_CONFIG.BOT_PARTIES_COLUMNS;
    const headers = [];
    const widths = [];

    // جمع العناوين والعروض
    Object.values(columns).forEach(col => {
        headers[col.index - 1] = col.name;
        widths[col.index - 1] = col.width;
    });

    // إضافة صف العناوين
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setValues([headers]);

    // تنسيق الهيدر
    headerRange
        .setBackground(CONFIG.COLORS.HEADER.PARTIES)
        .setFontColor(CONFIG.COLORS.TEXT.WHITE)
        .setFontWeight('bold')
        .setFontSize(CONFIG.FONT.NORMAL)
        .setHorizontalAlignment('center')
        .setVerticalAlignment('middle')
        .setWrap(true);

    // تعيين عرض الأعمدة
    widths.forEach((width, index) => {
        sheet.setColumnWidth(index + 1, width);
    });

    // تجميد الصف الأول
    sheet.setFrozenRows(1);

    // إضافة Data Validation لنوع الطرف
    const partyTypeCol = columns.PARTY_TYPE.index;
    const partyTypeRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(CONFIG.PARTY_TYPES.LIST)
        .setAllowInvalid(false)
        .build();

    sheet.getRange(2, partyTypeCol, CONFIG.SHEET.DEFAULT_ROWS, 1)
        .setDataValidation(partyTypeRule);

    // إضافة Data Validation لحالة المراجعة
    const reviewStatusCol = columns.REVIEW_STATUS.index;
    const reviewStatusRule = SpreadsheetApp.newDataValidation()
        .requireValueInList([
            CONFIG.TELEGRAM_BOT.REVIEW_STATUS.PENDING,
            CONFIG.TELEGRAM_BOT.REVIEW_STATUS.APPROVED,
            CONFIG.TELEGRAM_BOT.REVIEW_STATUS.REJECTED
        ])
        .setAllowInvalid(false)
        .build();

    sheet.getRange(2, reviewStatusCol, CONFIG.SHEET.DEFAULT_ROWS, 1)
        .setDataValidation(reviewStatusRule);

    // إضافة التنسيق الشرطي
    applyBotReviewConditionalFormatting(sheet, reviewStatusCol);

    // تنسيق أعمدة التاريخ
    sheet.getRange(2, columns.INPUT_TIMESTAMP.index, CONFIG.SHEET.DEFAULT_ROWS, 1)
        .setNumberFormat('dd/mm/yyyy hh:mm:ss');
    sheet.getRange(2, columns.REVIEW_TIMESTAMP.index, CONFIG.SHEET.DEFAULT_ROWS, 1)
        .setNumberFormat('dd/mm/yyyy hh:mm:ss');

    // تحديد لون التبويب
    sheet.setTabColor(CONFIG.COLORS.HEADER.PARTIES);

    Logger.log('Created sheet: ' + sheetName);
    return sheet;
}

/**
 * إنشاء شيت المستخدمين المصرح لهم (موحد مع Checkboxes)
 */
function createBotUsersSheet() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetName = CONFIG.SHEETS.BOT_USERS;

    // التحقق من وجود الشيت
    let sheet = ss.getSheetByName(sheetName);
    if (sheet) {
        Logger.log('Sheet already exists: ' + sheetName);
        return sheet;
    }

    // إنشاء شيت جديد
    sheet = ss.insertSheet(sheetName);

    // الأعمدة من BOT_CONFIG
    const columns = BOT_CONFIG.BOT_USERS_COLUMNS;
    const headers = [];
    const widths = [];

    // جمع العناوين والعروض
    Object.values(columns).forEach(col => {
        headers[col.index - 1] = col.name;
        widths[col.index - 1] = col.width;
    });

    // إضافة صف العناوين
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setValues([headers]);

    // تنسيق الهيدر
    headerRange
        .setBackground('#7b1fa2') // بنفسجي للمستخدمين
        .setFontColor(CONFIG.COLORS.TEXT.WHITE)
        .setFontWeight('bold')
        .setFontSize(CONFIG.FONT.NORMAL)
        .setHorizontalAlignment('center')
        .setVerticalAlignment('middle')
        .setWrap(true);

    // تعيين عرض الأعمدة
    widths.forEach((width, index) => {
        sheet.setColumnWidth(index + 1, width);
    });

    // تجميد الصف الأول
    sheet.setFrozenRows(1);

    // إضافة Checkboxes للصلاحيات
    const checkboxColumns = [
        columns.PERM_TRADITIONAL_BOT.index,
        columns.PERM_AI_BOT.index,
        columns.PERM_SHEET.index,
        columns.PERM_REVIEW.index,
        columns.IS_ACTIVE.index
    ];

    checkboxColumns.forEach(colIndex => {
        sheet.getRange(2, colIndex, CONFIG.SHEET.DEFAULT_ROWS, 1)
            .insertCheckboxes();
    });

    // تنسيق أعمدة التاريخ
    sheet.getRange(2, columns.ADDED_DATE.index, CONFIG.SHEET.DEFAULT_ROWS, 1)
        .setNumberFormat('dd/mm/yyyy');

    // تحديد لون التبويب
    sheet.setTabColor('#7b1fa2');

    // تنسيق أعمدة الـ Checkboxes (توسيط)
    checkboxColumns.forEach(colIndex => {
        sheet.getRange(2, colIndex, CONFIG.SHEET.DEFAULT_ROWS, 1)
            .setHorizontalAlignment('center');
    });

    // إضافة تنسيق شرطي للصف كامل إذا كان غير نشط
    const dataRange = sheet.getRange(2, 1, CONFIG.SHEET.DEFAULT_ROWS, headers.length);
    const inactiveRowRule = SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=$' + columnToLetter(columns.IS_ACTIVE.index) + '2=FALSE')
        .setBackground('#f5f5f5')
        .setFontColor('#9e9e9e')
        .setRanges([dataRange])
        .build();

    sheet.setConditionalFormatRules([inactiveRowRule]);

    Logger.log('Created sheet: ' + sheetName);
    return sheet;
}

/**
 * تحديث شيت المستخدمين القديم للهيكل الجديد
 * شغّل هذه الدالة مرة واحدة لتحويل الشيت القديم
 */
function upgradeUsersSheetToNewFormat() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetName = CONFIG.SHEETS.BOT_USERS;
    const sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
        Logger.log('شيت المستخدمين غير موجود - سيتم إنشاء شيت جديد');
        createBotUsersSheet();
        return;
    }

    // حذف الشيت القديم وإنشاء جديد
    const ui = SpreadsheetApp.getUi();
    const result = ui.alert(
        '⚠️ تحديث شيت المستخدمين',
        'سيتم حذف شيت المستخدمين القديم وإنشاء شيت جديد بالهيكل المحدث.\n\n' +
        '⚠️ تأكد من نسخ بيانات المستخدمين الحاليين قبل المتابعة!\n\n' +
        'هل تريد المتابعة؟',
        ui.ButtonSet.YES_NO
    );

    if (result !== ui.Button.YES) {
        return;
    }

    // حذف الشيت القديم
    ss.deleteSheet(sheet);

    // إنشاء الشيت الجديد
    createBotUsersSheet();

    ui.alert('✅ تم', 'تم تحديث شيت المستخدمين بنجاح!\n\nيرجى إعادة إضافة المستخدمين.', ui.ButtonSet.OK);
}

/**
 * تطبيق التنسيق الشرطي لحالة المراجعة
 */
function applyBotReviewConditionalFormatting(sheet, reviewStatusCol) {
    const range = sheet.getRange(2, 1, CONFIG.SHEET.DEFAULT_ROWS, sheet.getMaxColumns());

    const rules = [];

    // قيد الانتظار - برتقالي
    rules.push(SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=$' + columnToLetter(reviewStatusCol) + '2="' + CONFIG.TELEGRAM_BOT.REVIEW_STATUS.PENDING + '"')
        .setBackground(CONFIG.COLORS.BOT.PENDING)
        .setRanges([range])
        .build());

    // معتمد - أخضر
    rules.push(SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=$' + columnToLetter(reviewStatusCol) + '2="' + CONFIG.TELEGRAM_BOT.REVIEW_STATUS.APPROVED + '"')
        .setBackground(CONFIG.COLORS.BOT.APPROVED)
        .setRanges([range])
        .build());

    // مرفوض - أحمر
    rules.push(SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=$' + columnToLetter(reviewStatusCol) + '2="' + CONFIG.TELEGRAM_BOT.REVIEW_STATUS.REJECTED + '"')
        .setBackground(CONFIG.COLORS.BOT.REJECTED)
        .setRanges([range])
        .build());

    // يحتاج تعديل - أصفر
    rules.push(SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=$' + columnToLetter(reviewStatusCol) + '2="' + CONFIG.TELEGRAM_BOT.REVIEW_STATUS.NEEDS_EDIT + '"')
        .setBackground(CONFIG.COLORS.BOT.NEEDS_EDIT)
        .setRanges([range])
        .build());

    sheet.setConditionalFormatRules(rules);
}

/**
 * تحويل رقم العمود لحرف
 */
function columnToLetter(column) {
    let temp, letter = '';
    while (column > 0) {
        temp = (column - 1) % 26;
        letter = String.fromCharCode(temp + 65) + letter;
        column = (column - temp - 1) / 26;
    }
    return letter;
}

// ==================== دوال مساعدة للشيتات ====================

/**
 * الحصول على شيت حركات البوت
 */
function getBotTransactionsSheet() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(CONFIG.SHEETS.BOT_TRANSACTIONS);

    if (!sheet) {
        sheet = createBotTransactionsSheet();
    }

    return sheet;
}

/**
 * الحصول على شيت أطراف البوت
 */
function getBotPartiesSheet() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(CONFIG.SHEETS.BOT_PARTIES);

    if (!sheet) {
        sheet = createBotPartiesSheet();
    }

    return sheet;
}

/**
 * الحصول على شيت المستخدمين المصرح لهم
 */
function getBotUsersSheet() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(CONFIG.SHEETS.BOT_USERS);

    if (!sheet) {
        sheet = createBotUsersSheet();
    }

    return sheet;
}

/**
 * إضافة حركة جديدة لشيت حركات البوت
 */
function addBotTransaction(transactionData) {
    const sheet = getBotTransactionsSheet();
    const columns = BOT_CONFIG.BOT_TRANSACTIONS_COLUMNS;

    // إيجاد آخر صف
    const lastRow = sheet.getLastRow();
    const newRow = lastRow + 1;

    // تحديد رقم الحركة
    const transactionId = 'BOT-' + Utilities.formatDate(new Date(), 'Asia/Istanbul', 'yyyyMMdd-HHmmss');

    // إعداد البيانات
    const rowData = new Array(Object.keys(columns).length).fill('');

    rowData[columns.TRANSACTION_ID.index - 1] = transactionId;
    rowData[columns.DATE.index - 1] = transactionData.date || new Date();
    rowData[columns.NATURE.index - 1] = transactionData.nature;
    rowData[columns.CLASSIFICATION.index - 1] = transactionData.classification || '';
    rowData[columns.PROJECT_CODE.index - 1] = transactionData.projectCode || '';
    rowData[columns.PROJECT_NAME.index - 1] = transactionData.projectName || '';
    rowData[columns.ITEM.index - 1] = transactionData.item || '';
    rowData[columns.DETAILS.index - 1] = transactionData.details || '';
    rowData[columns.PARTY_NAME.index - 1] = transactionData.partyName || '';
    rowData[columns.AMOUNT.index - 1] = transactionData.amount || 0;
    rowData[columns.CURRENCY.index - 1] = transactionData.currency || 'USD';
    rowData[columns.EXCHANGE_RATE.index - 1] = transactionData.exchangeRate || 1;

    // حساب القيمة بالدولار
    const amountUSD = transactionData.currency === 'USD' || transactionData.currency === 'دولار'
        ? transactionData.amount
        : transactionData.amount / (transactionData.exchangeRate || 1);
    rowData[columns.AMOUNT_USD.index - 1] = amountUSD;

    // تحديد نوع الحركة
    const movementType = getMovementType(transactionData.nature);
    rowData[columns.MOVEMENT_TYPE.index - 1] = movementType;

    rowData[columns.PAYMENT_METHOD.index - 1] = transactionData.paymentMethod || '';
    rowData[columns.PAYMENT_TERM_TYPE.index - 1] = transactionData.paymentTermType || 'فوري';
    rowData[columns.WEEKS.index - 1] = transactionData.weeks || 0;
    rowData[columns.CUSTOM_DATE.index - 1] = transactionData.customDate || '';

    // أعمدة البوت
    rowData[columns.REVIEW_STATUS.index - 1] = CONFIG.TELEGRAM_BOT.REVIEW_STATUS.PENDING;
    rowData[columns.TELEGRAM_USER.index - 1] = transactionData.telegramUser || '';
    rowData[columns.TELEGRAM_CHAT_ID.index - 1] = transactionData.chatId || '';
    rowData[columns.INPUT_TIMESTAMP.index - 1] = new Date();
    rowData[columns.ATTACHMENT_URL.index - 1] = transactionData.attachmentUrl || '';
    rowData[columns.IS_NEW_PARTY.index - 1] = transactionData.isNewParty ? 'نعم' : 'لا';

    // إضافة الصف
    sheet.getRange(newRow, 1, 1, rowData.length).setValues([rowData]);

    return {
        success: true,
        transactionId: transactionId,
        rowNumber: newRow
    };
}

/**
 * تحديد نوع الحركة من طبيعتها
 */
function getMovementType(nature) {
    if (nature.includes('استحقاق')) {
        return CONFIG.MOVEMENT.DEBIT;
    } else if (nature.includes('دفعة') || nature.includes('تحصيل')) {
        return CONFIG.MOVEMENT.CREDIT;
    }
    return '';
}

/**
 * إضافة طرف جديد لشيت أطراف البوت
 */
function addBotParty(partyData) {
    const sheet = getBotPartiesSheet();
    const columns = BOT_CONFIG.BOT_PARTIES_COLUMNS;

    const lastRow = sheet.getLastRow();
    const newRow = lastRow + 1;

    const rowData = new Array(Object.keys(columns).length).fill('');

    rowData[columns.PARTY_NAME.index - 1] = partyData.name;
    rowData[columns.PARTY_TYPE.index - 1] = partyData.type;
    rowData[columns.REVIEW_STATUS.index - 1] = CONFIG.TELEGRAM_BOT.REVIEW_STATUS.PENDING;
    rowData[columns.TELEGRAM_USER.index - 1] = partyData.telegramUser || '';
    rowData[columns.TELEGRAM_CHAT_ID.index - 1] = partyData.chatId || '';
    rowData[columns.INPUT_TIMESTAMP.index - 1] = new Date();
    rowData[columns.LINKED_TRANSACTIONS.index - 1] = partyData.linkedTransactionId || '';

    sheet.getRange(newRow, 1, 1, rowData.length).setValues([rowData]);

    return {
        success: true,
        rowNumber: newRow
    };
}

/**
 * التحقق من صلاحية المستخدم (الهيكل الجديد مع Checkboxes)
 * يبحث بالهاتف أو اسم المستخدم أو معرّف المحادثة
 * @param {string} phoneNumber - رقم الهاتف
 * @param {string} chatId - معرّف المحادثة
 * @param {string} username - اسم المستخدم تليجرام
 * @param {string} permissionType - نوع الصلاحية المطلوبة: 'traditional_bot' | 'ai_bot' | 'sheet' | 'review'
 */
function checkUserAuthorization(phoneNumber, chatId, username, permissionType = 'traditional_bot') {
    const sheet = getBotUsersSheet();
    const columns = BOT_CONFIG.BOT_USERS_COLUMNS;

    const data = sheet.getDataRange().getValues();

    // تنظيف المدخلات
    const inputPhone = phoneNumber ? String(phoneNumber).replace(/\D/g, '') : '';
    const inputUsername = username ? String(username).toLowerCase().replace('@', '') : '';
    const inputChatId = chatId ? String(chatId) : '';

    Logger.log('Authorization check - Phone: ' + inputPhone + ', Username: ' + inputUsername + ', ChatId: ' + inputChatId + ', PermType: ' + permissionType);

    for (let i = 1; i < data.length; i++) {
        const row = data[i];

        // قراءة البيانات من الشيت
        const sheetPhone = String(row[columns.PHONE.index - 1] || '').replace(/\D/g, '');
        const sheetUsername = String(row[columns.TELEGRAM_USERNAME.index - 1] || '').toLowerCase().replace('@', '');
        const sheetChatId = String(row[columns.TELEGRAM_CHAT_ID.index - 1] || '');
        const isActive = row[columns.IS_ACTIVE.index - 1];

        // قراءة الصلاحيات (Checkboxes)
        const permTraditionalBot = row[columns.PERM_TRADITIONAL_BOT.index - 1] === true;
        const permAIBot = row[columns.PERM_AI_BOT.index - 1] === true;
        const permSheet = row[columns.PERM_SHEET.index - 1] === true;
        const permReview = row[columns.PERM_REVIEW.index - 1] === true;

        // التحقق من أن المستخدم نشط
        if (isActive !== true) {
            continue;
        }

        // التحقق من الصلاحية المطلوبة
        let hasPermission = false;
        switch (permissionType) {
            case 'traditional_bot':
                hasPermission = permTraditionalBot;
                break;
            case 'ai_bot':
                hasPermission = permAIBot;
                break;
            case 'sheet':
                hasPermission = permSheet;
                break;
            case 'review':
                hasPermission = permReview;
                break;
            default:
                hasPermission = permTraditionalBot; // افتراضي
        }

        if (!hasPermission) {
            continue;
        }

        // المطابقة بالهاتف (آخر 10 أرقام لتجاوز اختلاف الصيغ الدولية)
        let matched = false;

        if (inputPhone && sheetPhone) {
            const inputSuffix = inputPhone.slice(-10);
            const sheetSuffix = sheetPhone.slice(-10);

            if (inputSuffix === sheetSuffix) {
                matched = true;
                Logger.log('Matched by phone (Fuzzy)!');
            }
        }

        if (!matched && inputUsername && sheetUsername && inputUsername === sheetUsername) {
            matched = true;
            Logger.log('Matched by username!');
        } else if (!matched && inputChatId && sheetChatId && inputChatId === sheetChatId) {
            matched = true;
            Logger.log('Matched by chat ID!');
        }

        if (matched) {
            // تحديث Chat ID إذا لم يكن موجوداً
            if (!sheetChatId && chatId) {
                sheet.getRange(i + 1, columns.TELEGRAM_CHAT_ID.index).setValue(chatId);
            }

            // تحديث اسم المستخدم إذا لم يكن موجوداً
            if (!sheetUsername && username) {
                sheet.getRange(i + 1, columns.TELEGRAM_USERNAME.index).setValue(username);
            }

            return {
                authorized: true,
                name: row[columns.NAME.index - 1],
                permissions: {
                    traditionalBot: permTraditionalBot,
                    aiBot: permAIBot,
                    sheet: permSheet,
                    review: permReview
                }
            };
        }
    }

    Logger.log('No match found - User not authorized');
    return { authorized: false };
}

/**
 * التحقق من صلاحية المستخدم للبوت الذكي
 */
function checkAIBotAuthorization(chatId, username) {
    return checkUserAuthorization(null, chatId, username, 'ai_bot');
}

/**
 * التحقق من صلاحية المستخدم للبوت التقليدي
 */
function checkTraditionalBotAuthorization(phoneNumber, chatId, username) {
    return checkUserAuthorization(phoneNumber, chatId, username, 'traditional_bot');
}

/**
 * دالة اختبار التحقق من صلاحية المستخدم
 * شغّلها من Apps Script للتشخيص
 */
function testAuthorization() {
    // اختبار برقم الهاتف من الشيت
    const testPhone = "905530649846";

    Logger.log("═══════════════════════════════════════");
    Logger.log("=== بداية اختبار التصريح ===");
    Logger.log("═══════════════════════════════════════");
    Logger.log("Testing phone: " + testPhone);

    // عرض محتويات BOT_CONFIG.USER_TYPES
    Logger.log("BOT_CONFIG.USER_TYPES.BOT = '" + BOT_CONFIG.USER_TYPES.BOT + "'");
    Logger.log("BOT_CONFIG.USER_TYPES.BOTH = '" + BOT_CONFIG.USER_TYPES.BOTH + "'");

    const result = checkUserAuthorization(testPhone, null, null);

    Logger.log("═══════════════════════════════════════");
    Logger.log("=== النتيجة ===");
    Logger.log(JSON.stringify(result));

    if (result.authorized) {
        Logger.log("✅ المستخدم مصرح له!");
        Logger.log("الاسم: " + result.name);
        Logger.log("الصلاحية: " + result.permission);
    } else {
        Logger.log("❌ المستخدم غير مصرح له");
    }
    Logger.log("═══════════════════════════════════════");

    return result;
}

/**
 * دالة اختبار صلاحية البوت الذكي
 * شغّلها من Apps Script للتشخيص
 */
function testAIBotAuthorization() {
    const testChatId = "786700586"; // معرّف المحادثة الخاص بك
    const testUsername = "adelsolmn";

    Logger.log("═══════════════════════════════════════");
    Logger.log("=== اختبار صلاحية البوت الذكي ===");
    Logger.log("═══════════════════════════════════════");
    Logger.log("ChatId: " + testChatId);
    Logger.log("Username: " + testUsername);

    // قراءة الشيت للتحقق من البيانات
    const sheet = getBotUsersSheet();
    const data = sheet.getDataRange().getValues();
    Logger.log("عدد الصفوف في الشيت: " + data.length);

    // طباعة كل الصفوف للتحقق
    for (let i = 0; i < Math.min(data.length, 5); i++) {
        Logger.log("صف " + i + ": " + JSON.stringify(data[i]));
    }

    // اختبار الصلاحية
    const result = checkUserAuthorization(null, testChatId, testUsername, 'ai_bot');

    Logger.log("═══════════════════════════════════════");
    Logger.log("=== النتيجة ===");
    Logger.log(JSON.stringify(result));

    if (result.authorized) {
        Logger.log("✅ مصرح للبوت الذكي!");
        Logger.log("الاسم: " + result.name);
    } else {
        Logger.log("❌ غير مصرح للبوت الذكي");
    }
    Logger.log("═══════════════════════════════════════");

    return result;
}

/**
 * البحث عن المستخدم بالإيميل
 * تُستخدم لتسجيل النشاط مع اسم المستخدم
 * @param {string} email - البريد الإلكتروني للمستخدم
 * @returns {Object} بيانات المستخدم أو null
 */
function getUserByEmail(email) {
    try {
        if (!email) return null;

        const sheet = getBotUsersSheet();
        const columns = BOT_CONFIG.BOT_USERS_COLUMNS;
        const data = sheet.getDataRange().getValues();

        const inputEmail = String(email).toLowerCase().trim();

        for (let i = 1; i < data.length; i++) {
            const row = data[i];
            const sheetEmail = String(row[columns.EMAIL.index - 1] || '').toLowerCase().trim();

            if (sheetEmail && sheetEmail === inputEmail) {
                return {
                    found: true,
                    name: row[columns.NAME.index - 1] || '',
                    email: sheetEmail,
                    userType: row[columns.USER_TYPE.index - 1] || '',
                    permission: row[columns.PERMISSION.index - 1] || '',
                    isActive: row[columns.IS_ACTIVE.index - 1] === 'نعم'
                };
            }
        }

        return { found: false };
    } catch (error) {
        Logger.log('Error in getUserByEmail: ' + error.message);
        return { found: false };
    }
}

/**
 * الحصول على الحركات المعلقة للمراجعة
 */
function getPendingBotTransactions() {
    const sheet = getBotTransactionsSheet();
    const columns = BOT_CONFIG.BOT_TRANSACTIONS_COLUMNS;

    const data = sheet.getDataRange().getValues();
    const pending = [];

    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const status = row[columns.REVIEW_STATUS.index - 1];

        // تسجيل للتشخيص (مؤقت)
        if (i < 5) { // تسجيل أول 5 صفوف فقط لتجنب امتلاء السجل
            console.log(`Row ${i + 1}: Status='${status}' (Expected='${CONFIG.TELEGRAM_BOT.REVIEW_STATUS.PENDING}')`);
        }

        // استخدام String().trim() للتأكد من عدم تأثر المقارنة بالمسافات الزائدة
        if (String(status).trim() === CONFIG.TELEGRAM_BOT.REVIEW_STATUS.PENDING) {
            pending.push({
                rowNumber: i + 1,
                transactionId: row[columns.TRANSACTION_ID.index - 1],
                date: row[columns.DATE.index - 1],
                nature: row[columns.NATURE.index - 1],
                projectName: row[columns.PROJECT_NAME.index - 1],
                partyName: row[columns.PARTY_NAME.index - 1],
                amount: row[columns.AMOUNT.index - 1],
                currency: row[columns.CURRENCY.index - 1],
                details: row[columns.DETAILS.index - 1],
                telegramUser: row[columns.TELEGRAM_USER.index - 1],
                chatId: row[columns.TELEGRAM_CHAT_ID.index - 1],
                isNewParty: row[columns.IS_NEW_PARTY.index - 1] === 'نعم'
            });
        }
    }

    return pending;
}

/**
 * الحصول على عدد الحركات المعلقة
 */
function getPendingTransactionsCount() {
    return getPendingBotTransactions().length;
}
