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
 * @returns {string} نوع الحركة 'مدين استحقاق' أو 'دائن دفعة'
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
        nature.indexOf('استرداد تأمين') !== -1;          // استرداد نقدية

    // ═══════════════════════════════════════════════════════════
    // مدين استحقاق (دين/التزام)
    // ملاحظة: "تمويل" = دين علينا، لكن يجب استبعاد "سداد تمويل" و "استلام تمويل"
    // ═══════════════════════════════════════════════════════════
    const isDebit =
        nature.indexOf('استحقاق مصروف') !== -1 ||
        nature.indexOf('استحقاق إيراد') !== -1 ||
        (nature.indexOf('تمويل') !== -1 &&
         nature.indexOf('سداد تمويل') === -1 &&
         nature.indexOf('استلام تمويل') === -1);         // تمويل فقط (دين علينا)

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
