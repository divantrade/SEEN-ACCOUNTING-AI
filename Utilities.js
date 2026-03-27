// ==================== دوال مساعدة (Utility Functions) ====================
/**
 * ⚡ تحسينات الأولوية المتوسطة:
 * - توحيد الأنماط المتكررة في دوال مساعدة
 * - تقليل تكرار الكود وتحسين الصيانة
 */

/**
 * تعيين عرض الأعمدة دفعة واحدة
 * @param {Sheet} sheet - الشيت المستهدف
 * @param {number[]} widths - مصفوفة عروض الأعمدة
 */
function setColumnWidths_(sheet, widths) {
    widths.forEach((width, i) => sheet.setColumnWidth(i + 1, width));
}

/**
 * إعداد هيدر الشيت مع التنسيق
 * @param {Sheet} sheet - الشيت المستهدف
 * @param {string[]} headers - مصفوفة عناوين الأعمدة
 * @param {string} bgColor - لون الخلفية
 * @param {Object} options - خيارات إضافية (fontSize, textColor)
 */
function setupSheetHeader_(sheet, headers, bgColor, options = {}) {
    const textColor = options.textColor || CONFIG.COLORS.TEXT.WHITE;
    const fontSize = options.fontSize || CONFIG.FONT.NORMAL;

    sheet.getRange(1, 1, 1, headers.length)
        .setValues([headers])
        .setBackground(bgColor)
        .setFontColor(textColor)
        .setFontWeight('bold')
        .setFontSize(fontSize);
}

/**
 * الحصول على شيت أو إنشاؤه مع مسح المحتوى
 * @param {Spreadsheet} ss - ملف الجدول
 * @param {string} sheetName - اسم الشيت
 * @param {boolean} deleteExisting - حذف الشيت الموجود (default: false)
 * @returns {Sheet} الشيت
 */
function getOrCreateSheet_(ss, sheetName, deleteExisting = false) {
    let sheet = ss.getSheetByName(sheetName);
    if (sheet && deleteExisting) {
        ss.deleteSheet(sheet);
        sheet = null;
    }
    if (!sheet) {
        sheet = ss.insertSheet(sheetName);
    } else {
        sheet.clear();
    }
    return sheet;
}

/**
 * إعداد شيت كامل (هيدر + عرض أعمدة + تجميد)
 * @param {Sheet} sheet - الشيت المستهدف
 * @param {string[]} headers - مصفوفة عناوين الأعمدة
 * @param {number[]} widths - مصفوفة عروض الأعمدة
 * @param {string} bgColor - لون خلفية الهيدر
 * @param {Object} options - خيارات إضافية
 */
function setupSheet_(sheet, headers, widths, bgColor, options = {}) {
    setupSheetHeader_(sheet, headers, bgColor, options);
    setColumnWidths_(sheet, widths);
    sheet.setFrozenRows(options.frozenRows || 1);
    if (options.frozenCols) sheet.setFrozenColumns(options.frozenCols);
}


// ==================== دالة مساعدة للحصول على بيانات الطرف ====================
/**
 * الحصول على بيانات الطرف من قاعدة البيانات الموحدة (PARTIES) مع fallback للقواعد القديمة
 * @param {Spreadsheet} ss - ملف الجدول
 * @param {string} partyName - اسم الطرف
 * @param {string} partyType - نوع الطرف ('مورد' / 'عميل' / 'ممول') - اختياري للتصفية
 * @returns {Object} بيانات الطرف {name, type, specialization, phone, email, city, paymentMethod, bankInfo, notes}
 */
function getPartyData_(ss, partyName, partyType) {
    // النتيجة الافتراضية
    const defaultResult = {
        name: partyName,
        type: partyType || '',
        specialization: '',
        phone: '',
        email: '',
        city: '',
        paymentMethod: '',
        bankInfo: '',
        notes: ''
    };

    if (!partyName) return defaultResult;

    // ✅ أولاً: البحث في قاعدة البيانات الموحدة (PARTIES)
    const partiesSheet = ss.getSheetByName(CONFIG.SHEETS.PARTIES);
    if (partiesSheet) {
        const partiesData = partiesSheet.getDataRange().getValues();
        for (let i = 1; i < partiesData.length; i++) {
            const row = partiesData[i];
            if (row[0] === partyName) {
                // تأكد من مطابقة نوع الطرف إذا تم تحديده
                if (partyType && row[1] !== partyType) continue;

                return {
                    name: row[0] || partyName,
                    type: row[1] || partyType || '',
                    specialization: row[2] || '',
                    phone: row[3] || '',
                    email: row[4] || '',
                    city: row[5] || '',
                    paymentMethod: row[6] || '',
                    bankInfo: row[7] || '',
                    notes: row[8] || ''
                };
            }
        }
    }

    // ✅ ثانياً: Fallback للقواعد القديمة حسب نوع الطرف
    if (!partyType || partyType === 'مورد') {
        const vendorsSheet = ss.getSheetByName(CONFIG.SHEETS.LEGACY_VENDORS);
        if (vendorsSheet) {
            const vData = vendorsSheet.getDataRange().getValues();
            for (let i = 1; i < vData.length; i++) {
                if (vData[i][0] === partyName) {
                    return {
                        name: partyName,
                        type: 'مورد',
                        specialization: vData[i][1] || '',
                        phone: vData[i][2] || '',
                        email: vData[i][3] || '',
                        city: vData[i][4] || '',
                        paymentMethod: '',
                        bankInfo: vData[i][5] || '',
                        notes: vData[i][6] || ''
                    };
                }
            }
        }
    }

    if (!partyType || partyType === 'عميل') {
        const clientsSheet = ss.getSheetByName(CONFIG.SHEETS.LEGACY_CLIENTS);
        if (clientsSheet) {
            const cData = clientsSheet.getDataRange().getValues();
            for (let i = 1; i < cData.length; i++) {
                if (cData[i][0] === partyName) {
                    return {
                        name: partyName,
                        type: 'عميل',
                        specialization: cData[i][1] || '', // نوع العميل
                        phone: cData[i][2] || '',
                        email: cData[i][3] || '',
                        city: cData[i][4] || '',
                        paymentMethod: cData[i][5] || '', // قناة التواصل
                        bankInfo: cData[i][6] || '', // الشخص المسئول
                        notes: cData[i][7] || ''
                    };
                }
            }
        }
    }

    if (!partyType || partyType === 'ممول') {
        const fundersSheet = ss.getSheetByName(CONFIG.SHEETS.LEGACY_FUNDERS);
        if (fundersSheet) {
            const fData = fundersSheet.getDataRange().getValues();
            for (let i = 1; i < fData.length; i++) {
                if (fData[i][0] === partyName) {
                    return {
                        name: partyName,
                        type: 'ممول',
                        specialization: fData[i][1] || '',
                        phone: fData[i][2] || '',
                        email: fData[i][3] || '',
                        city: fData[i][4] || '',
                        paymentMethod: '',
                        bankInfo: fData[i][5] || '',
                        notes: fData[i][6] || ''
                    };
                }
            }
        }
    }

    return defaultResult;
}

/**
 * الحصول على خريطة تخصصات جميع الأطراف من نوع معين
 * @param {Spreadsheet} ss - ملف الجدول
 * @param {string} partyType - نوع الطرف ('مورد' / 'عميل' / 'ممول')
 * @returns {Object} خريطة {اسم الطرف: التخصص}
 */
function getPartySpecializationMap_(ss, partyType) {
    const specialMap = {};

    // ✅ أولاً: من قاعدة البيانات الموحدة
    const partiesSheet = ss.getSheetByName(CONFIG.SHEETS.PARTIES);
    if (partiesSheet) {
        const partiesData = partiesSheet.getDataRange().getValues();
        for (let i = 1; i < partiesData.length; i++) {
            const row = partiesData[i];
            if (row[0] && (!partyType || row[1] === partyType)) {
                specialMap[row[0]] = row[2] || '';
            }
        }
    }

    // ✅ ثانياً: Fallback للقواعد القديمة
    if (!partyType || partyType === 'مورد') {
        const vendorsSheet = ss.getSheetByName(CONFIG.SHEETS.LEGACY_VENDORS);
        if (vendorsSheet) {
            const vData = vendorsSheet.getDataRange().getValues();
            for (let i = 1; i < vData.length; i++) {
                if (vData[i][0] && !specialMap[vData[i][0]]) {
                    specialMap[vData[i][0]] = vData[i][1] || '';
                }
            }
        }
    }

    if (!partyType || partyType === 'عميل') {
        const clientsSheet = ss.getSheetByName(CONFIG.SHEETS.LEGACY_CLIENTS);
        if (clientsSheet) {
            const cData = clientsSheet.getDataRange().getValues();
            for (let i = 1; i < cData.length; i++) {
                if (cData[i][0] && !specialMap[cData[i][0]]) {
                    specialMap[cData[i][0]] = cData[i][1] || '';
                }
            }
        }
    }

    return specialMap;
}

/**
 * الانتقال لآخر سطر في دفتر الحركات المالية
 */
function scrollToLastRow_() {
    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const sheet = ss.getSheetByName(CONFIG.SHEETS.TRANSACTIONS);
        if (!sheet) return;

        // البحث عن آخر صف فيه بيانات فعلية في عمود B (التاريخ)
        const colB = sheet.getRange('B:B').getValues();
        let lastDataRow = 1;
        for (let i = colB.length - 1; i >= 0; i--) {
            if (colB[i][0] !== '' && colB[i][0] !== null) {
                lastDataRow = i + 1;
                break;
            }
        }

        if (lastDataRow > 1) {
            // تفعيل الشيت والانتقال لآخر سطر فيه بيانات
            ss.setActiveSheet(sheet);
            sheet.getRange(lastDataRow, 1).activate();
        }
    } catch (e) {
        // تجاهل الأخطاء - لا نريد منع فتح الملف
    }
}

/**
 * تحديد نوع الحركة (مدين/دائن) بناءً على طبيعة الحركة
 * يدعم البحث داخل النص (مثل: "🏦 تمويل (دخول قرض)" يتطابق مع "تمويل")
 * @param {string} natureType - طبيعة الحركة مثل 'استحقاق مصروف' أو '🏦 تمويل (دخول قرض)'
 * @returns {string} نوع الحركة 'مدين استحقاق' أو 'دائن دفعة' أو 'دائن تسوية'
 */
function getMovementTypeFromNature_(natureType) {
    if (!natureType) return null;

    const nature = String(natureType).trim();

    // ═══════════════════════════════════════════════════════════
    // دائن تسوية (خصم/تسوية من استحقاق - بدون حركة نقدية)
    // ⚠️ يجب فحصه قبل isCredit لأن "تسوية استحقاق مصروف" يحتوي على "مصروف"
    // ═══════════════════════════════════════════════════════════
    const isSettlement =
        nature.indexOf('تسوية استحقاق مصروف') !== -1 ||
        nature.indexOf('تسوية استحقاق إيراد') !== -1;

    // ═══════════════════════════════════════════════════════════
    // دائن دفعة (دفع/تحصيل/حركة نقدية فعلية)
    // ═══════════════════════════════════════════════════════════
    const isCredit =
        nature.indexOf('دفعة مصروف') !== -1 ||
        nature.indexOf('تحصيل إيراد') !== -1 ||
        nature.indexOf('سداد تمويل') !== -1 ||           // دفع قسط للممول = خروج نقدية
        nature.indexOf('تأمين مدفوع') !== -1 ||          // دفع تأمين = خروج نقدية
        nature.indexOf('استلام تمويل') !== -1 ||         // دخول نقدية القرض
        nature.indexOf('دخول قرض') !== -1 ||             // تمويل (دخول قرض) = دخول نقدية فعلية
        nature.indexOf('استرداد تأمين') !== -1 ||        // استرداد نقدية
        nature.indexOf('تحويل داخلي') !== -1 ||          // تحويل بين حسابات = حركة نقدية
        nature.indexOf('تغيير عملة') !== -1 ||             // تبديل عملة = حركة نقدية
        nature.indexOf('مصاريف بنكية') !== -1;           // خصم بنكي = خروج نقدية

    // ═══════════════════════════════════════════════════════════
    // مدين استحقاق (دين/التزام ورقي بدون حركة نقدية فعلية)
    // ملاحظة: "تمويل (دخول قرض)" = نقدية فعلية ← دائن دفعة (أعلاه)
    // ═══════════════════════════════════════════════════════════
    const isDebit =
        nature.indexOf('استحقاق مصروف') !== -1 ||
        nature.indexOf('استحقاق إيراد') !== -1 ||
        (nature.indexOf('تمويل') !== -1 &&
         nature.indexOf('سداد تمويل') === -1 &&
         nature.indexOf('استلام تمويل') === -1 &&
         nature.indexOf('دخول قرض') === -1);             // تمويل عام فقط (بدون دخول قرض)

    // الترتيب مهم: نفحص isSettlement أولاً ثم isCredit لأن "سداد تمويل" يحتوي على "تمويل"
    if (isSettlement) {
        return CONFIG.MOVEMENT.SETTLEMENT; // 'دائن تسوية'
    } else if (isCredit) {
        return CONFIG.MOVEMENT.CREDIT; // 'دائن دفعة'
    } else if (isDebit) {
        return CONFIG.MOVEMENT.DEBIT; // 'مدين استحقاق'
    }
    return null;
}

/**
 * قراءة أنواع الحركة (طبيعة الحركة) ديناميكياً من قاعدة بيانات البنود
 * تُقرأ القيم الفريدة من العمود B وتُرتب
 * @returns {string[]} مصفوفة بأنواع الحركة الفريدة
 */
function getNatureTypesFromItemsDB_() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var itemsSheet = ss.getSheetByName(CONFIG.SHEETS.ITEMS);

    if (!itemsSheet) {
        return []; // سيستخدم fallback
    }

    var lastRow = itemsSheet.getLastRow();
    if (lastRow < 2) {
        return [];
    }

    // قراءة عمود B (طبيعة الحركة) من الصف 2
    var data = itemsSheet.getRange(2, 2, lastRow - 1, 1).getValues();

    // استخراج القيم الفريدة
    var uniqueTypes = [];
    var seen = {};

    for (var i = 0; i < data.length; i++) {
        var value = data[i][0];
        if (value && value.toString().trim() !== '' && !seen[value]) {
            seen[value] = true;
            uniqueTypes.push(value.toString().trim());
        }
    }

    return uniqueTypes;
}

/**
 * البحث عن آخر صف فيه بيانات في عمود معين
 */
function findLastDataRowInColumn_(sheet, colNum) {
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return 1;

    var data = sheet.getRange(2, colNum, lastRow - 1, 1).getValues();
    for (var i = data.length - 1; i >= 0; i--) {
        if (data[i][0] !== '' && data[i][0] !== null) {
            return i + 2; // +2 لأن البيانات تبدأ من الصف 2
        }
    }
    return 1;
}

/**
 * تحويل صيغة التاريخ من dd/MM/yyyy أو dd.MM.yyyy إلى dd/MM/yyyy
 */
function parseDateInput_(dateStr) {
    // يقبل / أو . أو - كفاصل
    const regex = /^(\d{1,2})[\/\.\-](\d{1,2})[\/\.\-](\d{4})$/;
    const match = dateStr.match(regex);

    if (!match) {
        return { success: false, error: 'صيغة غير صحيحة!\nالمطلوب: يوم/شهر/سنة\nمثال: 24/12/2025' };
    }

    const day = parseInt(match[1], 10);
    const month = parseInt(match[2], 10);
    const year = parseInt(match[3], 10);

    if (month < 1 || month > 12) return { success: false, error: 'الشهر يجب أن يكون 1-12' };
    if (day < 1 || day > 31) return { success: false, error: 'اليوم يجب أن يكون 1-31' };

    const dateObj = new Date(year, month - 1, day);
    if (dateObj.getDate() !== day || dateObj.getMonth() !== month - 1) {
        return { success: false, error: 'تاريخ غير موجود!' };
    }

    return {
        success: true,
        date: String(day).padStart(2, '0') + '/' + String(month).padStart(2, '0') + '/' + year,
        dateObj: dateObj
    };
}

// ═══════════════════════════════════════════════════════════════════════════
// ⭐ دوال كشف الحساب الموحدة
// تُستخدم من مسار البوت (PDFGenerator.js) ومسار الشيت (Main.js)
// ═══════════════════════════════════════════════════════════════════════════

/**
 * أرقام أعمدة شيت دفتر الحركات المالية (0-based index)
 */
var TRANS_COL_ = {
    DATE: 1,         // B: تاريخ التسجيل
    NATURE: 2,       // C: طبيعة الحركة
    PROJECT: 5,      // F: المشروع
    ITEM: 6,         // G: البند
    DETAILS: 7,      // H: التفاصيل
    PARTY: 8,        // I: اسم الطرف
    AMOUNT_USD: 12,  // M: القيمة بالدولار
    MOVEMENT: 13,    // N: نوع الحركة (مدين/دائن/تسوية)
    DUE_DATE: 20,    // U: تاريخ الاستحقاق
    ORDER_NUM: 25    // Z: رقم الأوردر
};

/**
 * استخراج وتجهيز بيانات كشف الحساب من شيت الحركات
 * هذه الدالة الموحدة تُستخدم من البوت والشيت
 *
 * @param {Sheet} transSheet - شيت دفتر الحركات المالية
 * @param {string} partyName - اسم الطرف
 * @returns {Object} - { rows: Array, totalDebit, totalCredit, balance }
 *   rows: مصفوفة كائنات { date, dueDate, project, orderNumber, details, debit, credit, balance }
 */
function buildStatementData_(transSheet, partyName) {
    var data = transSheet.getDataRange().getValues();
    var C = TRANS_COL_;

    var regularRows = [];
    var sharedOrders = {};
    var totalDebit = 0, totalCredit = 0;

    for (var i = 1; i < data.length; i++) {
        var row = data[i];
        if (row[C.PARTY] !== partyName) continue;

        var movementKind = String(row[C.MOVEMENT] || '');
        var amountUsd = Number(row[C.AMOUNT_USD]) || 0;
        var natureType = String(row[C.NATURE] || '');
        if (!amountUsd) continue;

        var date = row[C.DATE];
        var dueDate = row[C.DUE_DATE];
        var project = row[C.PROJECT];
        var item = row[C.ITEM];
        var details = row[C.DETAILS];
        var orderNumber = row[C.ORDER_NUM] || '';

        // تمويل (دخول قرض) = يُعامل كمدين (دين على الشركة للممول)
        var isFundingIn = natureType.indexOf('تمويل') !== -1 && natureType.indexOf('سداد تمويل') === -1;

        var debit = 0, credit = 0;

        if (movementKind.includes(CONFIG.MOVEMENT.DEBIT) || movementKind.includes('مدين') || isFundingIn) {
            debit = amountUsd;
            totalDebit += debit;
        } else if (movementKind.includes(CONFIG.MOVEMENT.SETTLEMENT) || movementKind.includes('تسوية')) {
            credit = amountUsd;
            totalCredit += credit;
        } else if (movementKind.includes(CONFIG.MOVEMENT.CREDIT) || movementKind.includes('دائن')) {
            credit = amountUsd;
            totalCredit += credit;
        }

        // تجميع الأوردرات المشتركة (SO-)
        if (orderNumber && String(orderNumber).startsWith('SO-')) {
            if (!sharedOrders[orderNumber]) {
                sharedOrders[orderNumber] = {
                    date: date,
                    dueDate: (debit > 0 && dueDate) ? dueDate : '',
                    items: [],
                    _itemSet: {},
                    details: [],
                    _detailSet: {},
                    debit: 0,
                    credit: 0,
                    orderNumber: orderNumber
                };
            }
            var so = sharedOrders[orderNumber];
            if (item && !so._itemSet[item]) { so.items.push(item); so._itemSet[item] = true; }
            if (details && !so._detailSet[details]) { so.details.push(details); so._detailSet[details] = true; }
            so.debit += debit;
            so.credit += credit;
            if (debit > 0 && dueDate && !so.dueDate) so.dueDate = dueDate;
        } else {
            regularRows.push({
                date: date,
                dueDate: (debit > 0 && dueDate) ? dueDate : '',
                project: project || '',
                orderNumber: orderNumber,
                details: details || '',
                debit: debit,
                credit: credit
            });
        }
    }

    // تحويل الأوردرات المشتركة إلى صفوف مدمجة
    var sharedRows = [];
    var orderKeys = Object.keys(sharedOrders);
    for (var k = 0; k < orderKeys.length; k++) {
        var order = sharedOrders[orderKeys[k]];
        var guestNames = [];
        var totalGuestsFromDetails = 0;

        for (var d = 0; d < order.details.length; d++) {
            var detail = order.details[d];
            if (!detail) continue;

            var namePart = '';
            var newFmt = detail.match(/^مقابلة\s+(.+?)\s*\(\d+\s*من\s*\d+\)/);
            if (newFmt) {
                namePart = newFmt[1].trim();
            } else {
                namePart = detail.split(' - ')[0].trim();
            }

            var gcMatch = detail.match(/(\d+)\s*ضي[وف]/);
            if (gcMatch && totalGuestsFromDetails === 0) {
                totalGuestsFromDetails = parseInt(gcMatch[1]);
            }

            if (namePart) {
                var names = namePart.split(/[،,]/).map(function(n) { return n.trim(); }).filter(function(n) { return n; });
                for (var n = 0; n < names.length; n++) {
                    if (names[n] && guestNames.indexOf(names[n]) === -1) {
                        guestNames.push(names[n]);
                    }
                }
            }
        }

        var itemsStr = order.items.join(' + ');
        var guestCount = totalGuestsFromDetails || guestNames.length || order.details.length;
        var detailsStr = itemsStr;
        if (guestCount > 0) {
            detailsStr += ' - ' + guestCount + ' ضيوف';
            if (guestNames.length > 0) {
                detailsStr += ': ' + guestNames.join('، ');
            }
        }

        sharedRows.push({
            date: order.date,
            dueDate: order.dueDate || '',
            project: '',
            orderNumber: order.orderNumber,
            details: detailsStr,
            debit: order.debit,
            credit: order.credit
        });
    }

    // دمج + ترتيب زمني
    var allRows = regularRows.concat(sharedRows);
    allRows.sort(function(a, b) {
        var dateA = a.date instanceof Date ? a.date.getTime() : new Date(a.date).getTime();
        var dateB = b.date instanceof Date ? b.date.getTime() : new Date(b.date).getTime();
        return dateA - dateB;
    });

    // حساب الرصيد التراكمي
    var balance = 0;
    for (var r = 0; r < allRows.length; r++) {
        balance += (allRows[r].debit || 0) - (allRows[r].credit || 0);
        allRows[r].balance = Math.round(balance * 100) / 100;
    }

    return {
        rows: allRows,
        totalDebit: totalDebit,
        totalCredit: totalCredit,
        balance: Math.round(balance * 100) / 100
    };
}

/**
 * استخراج بيانات اللوجو من شيت قاعدة بيانات البنود (D2)
 * يدعم: CellImage + IMAGE formula + URL نصي + OverGridImage + بحث مجاور
 *
 * @param {Spreadsheet} ss - الجدول
 * @returns {Object} - { logoBlob, logoFileId, logoOriginalUrl, hasCellImage, logoSourceRange }
 */
function extractLogoData_(ss) {
    var result = {
        logoBlob: null,
        logoFileId: '',
        logoOriginalUrl: '',
        hasCellImage: false,
        logoSourceRange: null
    };

    try {
        var itemsSheet = ss.getSheetByName(CONFIG.SHEETS.ITEMS || 'قاعدة بيانات البنود');
        if (!itemsSheet) return result;

        var d2Range = itemsSheet.getRange('D2');
        var d2Value = d2Range.getValue();
        var d2Type = typeof d2Value;
        var logoUrl = '';

        // الحالة 1: CellImage
        if (d2Value && d2Type === 'object') {
            result.hasCellImage = true;
            result.logoSourceRange = d2Range;
            try {
                if (typeof d2Value.getContentUrl === 'function') logoUrl = d2Value.getContentUrl() || '';
                if (!logoUrl && typeof d2Value.getUrl === 'function') logoUrl = d2Value.getUrl() || '';
            } catch (e) {
                Logger.log('extractLogoData_: CellImage read error: ' + e.message);
            }
        }

        // الحالة 2: نص عادي (URL)
        if (!logoUrl && d2Value && d2Type === 'string') {
            logoUrl = String(d2Value).trim();
        }

        // الحالة 3: معادلة IMAGE
        if (!logoUrl) {
            var formula = d2Range.getFormula() || '';
            var fMatch = formula.match(/IMAGE\s*\(\s*"([^"]+)"/i);
            if (fMatch) logoUrl = fMatch[1];
        }

        // الحالة 4: صورة عائمة (OverGridImage)
        if (!logoUrl && !result.logoBlob) {
            try {
                var images = itemsSheet.getImages();
                if (images.length > 0) result.logoBlob = images[0].getBlob();
            } catch (e) {
                Logger.log('extractLogoData_: getImages error: ' + e.message);
            }
        }

        // الحالة 5: بحث في الخلايا المجاورة
        if (!logoUrl && !result.logoBlob) {
            try {
                var lastCol = Math.min(itemsSheet.getLastColumn(), 10);
                if (lastCol >= 5) {
                    var searchRange = itemsSheet.getRange(1, 5, Math.min(itemsSheet.getLastRow(), 5), lastCol - 4);
                    var searchValues = searchRange.getValues();
                    for (var r = 0; r < searchValues.length && !logoUrl; r++) {
                        for (var c = 0; c < searchValues[r].length && !logoUrl; c++) {
                            var val = String(searchValues[r][c] || '').trim();
                            if (val && (val.includes('drive.google.com/file/d/') || val.includes('googleusercontent.com/d/'))) {
                                logoUrl = val;
                            }
                        }
                    }
                }
            } catch (e) {
                Logger.log('extractLogoData_: nearby cell search error: ' + e.message);
            }
        }

        result.logoOriginalUrl = logoUrl;

        // استخراج File ID
        if (logoUrl) {
            var m = logoUrl.match(/\/file\/d\/([a-zA-Z0-9_-]+)/) ||
                    logoUrl.match(/[?&]id=([a-zA-Z0-9_-]+)/) ||
                    logoUrl.match(/googleusercontent\.com\/d\/([a-zA-Z0-9_-]+)/) ||
                    logoUrl.match(/\/d\/([a-zA-Z0-9_-]{10,})/);
            if (m) result.logoFileId = m[1];
        }
    } catch (e) {
        Logger.log('extractLogoData_: error: ' + e.message);
    }

    return result;
}

/**
 * إدراج اللوجو في شيت كشف الحساب (G1:H3)
 * يجرب 5 طرق متتالية حتى ينجح
 *
 * @param {Sheet} sheet - الشيت المستهدف
 * @param {Object} logo - نتيجة extractLogoData_()
 */
function insertLogo_(sheet, logo) {
    var inserted = false;

    function ensureMerge_() {
        sheet.getRange('G1:H3').merge()
            .setBackground('#f8f9fa')
            .setHorizontalAlignment('center')
            .setVerticalAlignment('middle');
    }

    if (!logo.hasCellImage && !logo.logoBlob && !logo.logoFileId && !logo.logoOriginalUrl) {
        ensureMerge_();
        return;
    }

    // الطريقة 1: نسخ CellImage
    if (logo.hasCellImage && logo.logoSourceRange && !inserted) {
        try {
            logo.logoSourceRange.copyTo(sheet.getRange('G1'), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
            ensureMerge_();
            inserted = true;
        } catch (e) { Logger.log('insertLogo_ CellCopy FAILED: ' + e.message); }
    }

    // الطريقة 2: blob مباشر
    if (logo.logoBlob && !inserted) {
        try {
            ensureMerge_();
            var img = sheet.insertImage(logo.logoBlob, 7, 1);
            img.setWidth(CONFIG.COMPANY.LOGO.WIDTH);
            img.setHeight(CONFIG.COMPANY.LOGO.HEIGHT);
            inserted = true;
        } catch (e) { Logger.log('insertLogo_ blob FAILED: ' + e.message); }
    }

    // الطريقة 3: DriveApp
    if (logo.logoFileId && !inserted) {
        try {
            ensureMerge_();
            var file = DriveApp.getFileById(logo.logoFileId);
            var blob = file.getBlob();
            var img2 = sheet.insertImage(blob, 7, 1);
            img2.setWidth(CONFIG.COMPANY.LOGO.WIDTH);
            img2.setHeight(CONFIG.COMPANY.LOGO.HEIGHT);
            inserted = true;
        } catch (e) { Logger.log('insertLogo_ DriveApp FAILED: ' + e.message); }
    }

    // الطريقة 4a: UrlFetchApp + lh3
    if (logo.logoFileId && !inserted) {
        try {
            var lh3Url = 'https://lh3.googleusercontent.com/d/' + logo.logoFileId;
            var resp = UrlFetchApp.fetch(lh3Url, { muteHttpExceptions: true, followRedirects: true });
            if (resp.getResponseCode() === 200) {
                ensureMerge_();
                var img3 = sheet.insertImage(resp.getBlob(), 7, 1);
                img3.setWidth(CONFIG.COMPANY.LOGO.WIDTH);
                img3.setHeight(CONFIG.COMPANY.LOGO.HEIGHT);
                inserted = true;
            }
        } catch (e) { Logger.log('insertLogo_ lh3 FAILED: ' + e.message); }
    }

    // الطريقة 4b: UrlFetchApp + uc?export
    if (logo.logoFileId && !inserted) {
        try {
            var directUrl = 'https://drive.google.com/uc?export=view&id=' + logo.logoFileId;
            var resp2 = UrlFetchApp.fetch(directUrl, { muteHttpExceptions: true, followRedirects: true });
            if (resp2.getResponseCode() === 200) {
                ensureMerge_();
                var img4 = sheet.insertImage(resp2.getBlob(), 7, 1);
                img4.setWidth(CONFIG.COMPANY.LOGO.WIDTH);
                img4.setHeight(CONFIG.COMPANY.LOGO.HEIGHT);
                inserted = true;
            }
        } catch (e) { Logger.log('insertLogo_ uc FAILED: ' + e.message); }
    }

    // الطريقة 5: IMAGE formula
    if (!inserted) {
        try {
            var imgUrl = logo.logoFileId
                ? 'https://lh3.googleusercontent.com/d/' + logo.logoFileId
                : logo.logoOriginalUrl;
            if (imgUrl) {
                ensureMerge_();
                sheet.getRange('G1').setFormula('=IMAGE("' + imgUrl + '", 2)');
                inserted = true;
            }
        } catch (e) { Logger.log('insertLogo_ formula FAILED: ' + e.message); }
    }

    if (!inserted) ensureMerge_();
}

/**
 * رسم كشف الحساب على شيت
 * دالة موحدة تُستخدم من البوت والشيت
 *
 * @param {Object} opts
 * @param {Spreadsheet} opts.ss - الجدول
 * @param {string} opts.partyName - اسم الطرف
 * @param {string} opts.partyType - نوع الطرف (مورد/عميل/ممول)
 * @param {boolean} [opts.useEmoji=false] - إضافة إيموجي لعناوين الأعمدة
 * @param {boolean} [opts.includeOrderCol=false] - إضافة عمود رقم الأوردر
 * @returns {Sheet} - الشيت المُنشأ
 */
function renderStatementSheet_(opts) {
    var ss = opts.ss;
    var partyName = opts.partyName;
    var partyType = opts.partyType;
    var useEmoji = opts.useEmoji || false;
    var includeOrderCol = opts.includeOrderCol || false;

    var transSheet = ss.getSheetByName(CONFIG.SHEETS.TRANSACTIONS);
    if (!transSheet) throw new Error('شيت دفتر الحركات المالية غير موجود!');

    // === بيانات الطرف ===
    var partyData = getPartyData_(ss, partyName, partyType);

    // === اللوجو ===
    var logo = extractLogoData_(ss);

    // === بيانات الحركات (الدالة الموحدة) ===
    var stmtData = buildStatementData_(transSheet, partyName);

    // === عنوان ولون التبويب ===
    var titlePrefix = 'كشف حساب';
    var tabColor = '#4a86e8';

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

    // === إنشاء الشيت ===
    var sheetName = titlePrefix + ' - ' + partyName;
    var sheet = ss.getSheetByName(sheetName);
    if (sheet) ss.deleteSheet(sheet);
    sheet = ss.insertSheet(sheetName);
    sheet.setTabColor(tabColor);
    sheet.setRightToLeft(true);

    // === عدد الأعمدة وعرضها ===
    var totalCols = includeOrderCol ? 8 : 7;
    // col indices (1-based) for data columns shift based on includeOrderCol
    if (includeOrderCol) {
        sheet.setColumnWidth(1, 100);  // تاريخ التسجيل
        sheet.setColumnWidth(2, 100);  // تاريخ الاستحقاق
        sheet.setColumnWidth(3, 140);  // المشروع
        sheet.setColumnWidth(4, 120);  // رقم الأوردر
        sheet.setColumnWidth(5, 200);  // التفاصيل
        sheet.setColumnWidth(6, 110);  // مدين
        sheet.setColumnWidth(7, 110);  // دائن
        sheet.setColumnWidth(8, 150);  // الرصيد
    } else {
        sheet.setColumnWidth(1, 100);  // تاريخ التسجيل
        sheet.setColumnWidth(2, 100);  // تاريخ الاستحقاق
        sheet.setColumnWidth(3, 140);  // المشروع
        sheet.setColumnWidth(4, 210);  // التفاصيل
        sheet.setColumnWidth(5, 110);  // مدين
        sheet.setColumnWidth(6, 110);  // دائن
        sheet.setColumnWidth(7, 110);  // الرصيد
        sheet.setColumnWidth(8, 80);   // عمود اللوجو
    }

    var lastColLetter = includeOrderCol ? 'H' : 'H';

    // === الترويسة (Letterhead) ===
    sheet.setRowHeight(1, 45);
    sheet.setRowHeight(2, 35);
    sheet.setRowHeight(3, 30);
    sheet.getRange('A1:H3').setBackground('#f8f9fa');

    sheet.getRange('A1:F1').merge()
        .setValue(CONFIG.COMPANY.NAME)
        .setFontWeight('bold')
        .setFontSize(14)
        .setFontColor('#1a237e')
        .setHorizontalAlignment('center')
        .setVerticalAlignment('bottom');

    sheet.getRange('A2:F2').merge()
        .setValue(CONFIG.COMPANY.ADDRESS)
        .setFontSize(10)
        .setFontColor('#555555')
        .setHorizontalAlignment('center')
        .setVerticalAlignment('middle');

    sheet.getRange('A3:F3').merge()
        .setValue(CONFIG.COMPANY.CONTACT)
        .setFontSize(10)
        .setFontColor('#555555')
        .setHorizontalAlignment('center')
        .setVerticalAlignment('top');

    sheet.getRange('A3:H3').setBorder(
        false, false, true, false, false, false,
        '#1a237e', SpreadsheetApp.BorderStyle.SOLID_MEDIUM
    );

    // === اللوجو ===
    insertLogo_(sheet, logo);

    // === فاصل + عنوان ===
    sheet.setRowHeight(4, 10);
    sheet.getRange('A5:H5').merge();
    sheet.getRange('A5')
        .setValue((useEmoji ? '📊 ' : '') + titlePrefix)
        .setBackground(CONFIG.COLORS.HEADER.DASHBOARD)
        .setFontColor(CONFIG.COLORS.TEXT.WHITE)
        .setFontWeight('bold')
        .setFontSize(15)
        .setHorizontalAlignment('center')
        .setVerticalAlignment('middle');

    // === كارت بيانات الطرف ===
    var cardHeaderRow = useEmoji ? 6 : 7;
    var cardDataStartRow = cardHeaderRow + 1;

    sheet.getRange('A' + cardHeaderRow + ':H' + cardHeaderRow).merge()
        .setValue('بيانات ' + partyType)
        .setBackground(CONFIG.COLORS.HEADER.SUMMARY)
        .setFontColor(CONFIG.COLORS.TEXT.WHITE)
        .setFontWeight('bold')
        .setHorizontalAlignment('center');

    sheet.getRange('A' + cardDataStartRow + ':H' + (cardDataStartRow + 3)).setBackground(CONFIG.COLORS.BG.LIGHT_BLUE);

    // الصف الأول: اسم + تخصص
    sheet.getRange('A' + cardDataStartRow).setValue('الاسم:').setFontWeight('bold');
    if (includeOrderCol) {
        sheet.getRange('B' + cardDataStartRow + ':D' + cardDataStartRow).merge().setValue(partyName);
        sheet.getRange('F' + cardDataStartRow).setValue('التخصص:').setFontWeight('bold');
        sheet.getRange('G' + cardDataStartRow + ':H' + cardDataStartRow).merge().setValue(partyData.specialization || '');
    } else {
        sheet.getRange('B' + cardDataStartRow + ':C' + cardDataStartRow).merge().setValue(partyName);
        sheet.getRange('E' + cardDataStartRow).setValue('التخصص:').setFontWeight('bold');
        sheet.getRange('F' + cardDataStartRow + ':H' + cardDataStartRow).merge().setValue(partyData.specialization || '');
    }

    // الصف الثاني: هاتف + بريد
    sheet.getRange('A' + (cardDataStartRow + 1)).setValue('الهاتف:').setFontWeight('bold');
    if (includeOrderCol) {
        sheet.getRange('B' + (cardDataStartRow + 1) + ':D' + (cardDataStartRow + 1)).merge().setValue(partyData.phone || '');
        sheet.getRange('F' + (cardDataStartRow + 1)).setValue('البريد:').setFontWeight('bold');
        sheet.getRange('G' + (cardDataStartRow + 1) + ':H' + (cardDataStartRow + 1)).merge().setValue(partyData.email || '');
    } else {
        sheet.getRange('B' + (cardDataStartRow + 1) + ':C' + (cardDataStartRow + 1)).merge().setValue(partyData.phone || '');
        sheet.getRange('E' + (cardDataStartRow + 1)).setValue('البريد:').setFontWeight('bold');
        sheet.getRange('F' + (cardDataStartRow + 1) + ':H' + (cardDataStartRow + 1)).merge().setValue(partyData.email || '');
    }

    // الصف الثالث: بنك
    sheet.getRange('A' + (cardDataStartRow + 2)).setValue('البنك:').setFontWeight('bold');
    sheet.getRange('B' + (cardDataStartRow + 2) + ':H' + (cardDataStartRow + 2)).merge().setValue(partyData.bankInfo || '');

    // الصف الرابع: ملاحظات
    sheet.getRange('A' + (cardDataStartRow + 3)).setValue('ملاحظات:').setFontWeight('bold');
    sheet.getRange('B' + (cardDataStartRow + 3) + ':H' + (cardDataStartRow + 3)).merge().setValue(partyData.notes || '').setWrap(true);

    sheet.getRange('A' + cardDataStartRow + ':H' + (cardDataStartRow + 3)).setBorder(
        true, true, true, true, true, true,
        '#1565c0', SpreadsheetApp.BorderStyle.SOLID
    );

    // === الملخص المالي ===
    var summaryHeaderRow = cardDataStartRow + 5;
    var summaryDataStartRow = summaryHeaderRow + 1;

    sheet.getRange('A' + summaryHeaderRow + ':H' + summaryHeaderRow).merge()
        .setValue('الملخص المالي')
        .setBackground(CONFIG.COLORS.HEADER.SUMMARY)
        .setFontColor(CONFIG.COLORS.TEXT.WHITE)
        .setFontWeight('bold')
        .setHorizontalAlignment('center');

    sheet.getRange('A' + summaryDataStartRow + ':H' + (summaryDataStartRow + 1)).setBackground(CONFIG.COLORS.BG.LIGHT_BLUE);

    sheet.getRange('A' + summaryDataStartRow).setValue('إجمالي المدين:').setFontWeight('bold');
    sheet.getRange('B' + summaryDataStartRow).setValue(stmtData.totalDebit).setNumberFormat('$#,##0.00');

    sheet.getRange('E' + summaryDataStartRow).setValue('إجمالي الدائن:').setFontWeight('bold');
    sheet.getRange('F' + summaryDataStartRow).setValue(stmtData.totalCredit).setNumberFormat('$#,##0.00');

    sheet.getRange('A' + (summaryDataStartRow + 1)).setValue('الرصيد الحالي:').setFontWeight('bold');
    sheet.getRange('B' + (summaryDataStartRow + 1)).setValue(stmtData.balance).setNumberFormat('$#,##0.00')
        .setFontWeight('bold')
        .setBackground(stmtData.balance > 0 ? '#ffcdd2' : '#c8e6c9');

    sheet.getRange('E' + (summaryDataStartRow + 1)).setValue('عدد الحركات:').setFontWeight('bold');
    sheet.getRange('F' + (summaryDataStartRow + 1)).setValue(stmtData.rows.length);

    sheet.getRange('A' + summaryDataStartRow + ':H' + (summaryDataStartRow + 1)).setBorder(
        true, true, true, true, true, true,
        '#1565c0', SpreadsheetApp.BorderStyle.SOLID
    );

    // === رأس جدول الحركات ===
    var tableHeaderRow = summaryDataStartRow + 3;
    var headers;
    if (includeOrderCol) {
        if (useEmoji) {
            headers = ['📅 تاريخ التسجيل', '⏰ تاريخ الاستحقاق', '🎬 المشروع', '📦 رقم الأوردر', '📝 التفاصيل', '💰 مدين (USD)', '💸 دائن (USD)', '📊 الرصيد (USD)'];
        } else {
            headers = ['تاريخ التسجيل', 'تاريخ الاستحقاق', 'المشروع', 'رقم الأوردر', 'التفاصيل', 'مدين (USD)', 'دائن (USD)', 'الرصيد (USD)'];
        }
    } else {
        if (useEmoji) {
            headers = ['📅 تاريخ التسجيل', '⏰ تاريخ الاستحقاق', '🎬 المشروع', '📝 التفاصيل', '💰 مدين (USD)', '💸 دائن (USD)', '📊 الرصيد (USD)'];
        } else {
            headers = ['تاريخ التسجيل', 'تاريخ الاستحقاق', 'المشروع', 'التفاصيل', 'مدين (USD)', 'دائن (USD)', 'الرصيد (USD)'];
        }
    }

    sheet.getRange(tableHeaderRow, 1, 1, headers.length)
        .setValues([headers])
        .setBackground(CONFIG.COLORS.HEADER.DASHBOARD)
        .setFontColor(CONFIG.COLORS.TEXT.WHITE)
        .setFontWeight('bold')
        .setHorizontalAlignment('center');

    // === بيانات الحركات ===
    var dataStartRow = tableHeaderRow + 1;
    var debitCol = includeOrderCol ? 6 : 5;

    if (stmtData.rows.length > 0) {
        // تحويل كائنات الصفوف إلى مصفوفات
        var sheetRows = [];
        for (var i = 0; i < stmtData.rows.length; i++) {
            var r = stmtData.rows[i];
            if (includeOrderCol) {
                sheetRows.push([r.date, r.dueDate || '', r.project, r.orderNumber || '', r.details, r.debit || '', r.credit || '', r.balance]);
            } else {
                sheetRows.push([r.date, r.dueDate || '', r.project, r.details, r.debit || '', r.credit || '', r.balance]);
            }
        }

        sheet.getRange(dataStartRow, 1, sheetRows.length, headers.length).setValues(sheetRows);
        sheet.getRange(dataStartRow, 1, sheetRows.length, 1).setNumberFormat('dd/mm/yyyy');
        sheet.getRange(dataStartRow, 2, sheetRows.length, 1).setNumberFormat('dd/mm/yyyy');
        sheet.getRange(dataStartRow, debitCol, sheetRows.length, 3).setNumberFormat('$#,##0.00');

        // تلوين متناوب
        for (var j = 0; j < sheetRows.length; j++) {
            var rowNum = dataStartRow + j;
            var bg = j % 2 === 0 ? '#ffffff' : CONFIG.COLORS.BG.LIGHT_BLUE;
            sheet.getRange(rowNum, 1, 1, headers.length).setBackground(bg);
        }

        // إطار الجدول
        sheet.getRange(tableHeaderRow, 1, sheetRows.length + 1, headers.length)
            .setBorder(true, true, true, true, true, true, '#bdbdbd', SpreadsheetApp.BorderStyle.SOLID);
    } else {
        sheet.getRange(dataStartRow, 1).setValue('لا توجد حركات').setFontStyle('italic');
    }

    // === التذييل ===
    var footerRow = dataStartRow + Math.max(stmtData.rows.length, 1) + 2;
    var reportDateStr = Utilities.formatDate(new Date(), CONFIG.COMPANY.TIMEZONE, 'dd/MM/yyyy HH:mm');

    sheet.getRange('A' + footerRow + ':H' + footerRow).merge()
        .setValue('تاريخ التقرير: ' + reportDateStr)
        .setHorizontalAlignment('center')
        .setFontSize(9)
        .setFontColor('#757575');

    sheet.getRange('A' + (footerRow + 1) + ':H' + (footerRow + 1)).merge()
        .setValue(CONFIG.COMPANY.CREDITS)
        .setHorizontalAlignment('center')
        .setFontSize(9)
        .setFontColor('#9e9e9e');

    SpreadsheetApp.flush();
    Logger.log('renderStatementSheet_: ' + partyName + ' - ' + stmtData.rows.length + ' rows');

    // إرجاع الشيت + البيانات لتجنب حسابها مرتين
    sheet._stmtData = stmtData;
    return sheet;
}
