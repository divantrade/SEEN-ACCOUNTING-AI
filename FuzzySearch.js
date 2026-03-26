// ==================== نظام البحث الذكي للنصوص العربية ====================
/**
 * ملف البحث الذكي (Fuzzy Search) للنصوص العربية
 * يدعم التطابق التقريبي والتعامل مع اختلافات الكتابة العربية
 */

// ==================== تطبيع النص العربي ====================

/**
 * تطبيع النص العربي للبحث
 * يوحد الأحرف المتشابهة ويزيل التشكيل
 */
function normalizeArabicText(text) {
    if (!text) return '';

    let normalized = String(text).trim();

    // إزالة التشكيل (الحركات) - regex أشمل من المصفوفة اليدوية
    normalized = normalized.replace(/[\u064B-\u065F\u0670\u0640]/g, '');

    // توحيد الهمزات
    normalized = normalized.replace(/[أإآٱ]/g, 'ا');
    normalized = normalized.replace(/[ؤ]/g, 'و');
    normalized = normalized.replace(/[ئ]/g, 'ي');

    // توحيد التاء المربوطة والهاء
    normalized = normalized.replace(/ة/g, 'ه');

    // توحيد الألف المقصورة والياء
    normalized = normalized.replace(/ى/g, 'ي');

    // إزالة المسافات الزائدة
    normalized = normalized.replace(/\s+/g, ' ').trim();

    // تحويل للحروف الصغيرة (للنصوص الإنجليزية المختلطة)
    normalized = normalized.toLowerCase();

    return normalized;
}

/**
 * حساب مسافة Levenshtein بين نصين
 * (عدد التعديلات المطلوبة لتحويل نص لآخر)
 */
function levenshteinDistance(str1, str2) {
    const m = str1.length;
    const n = str2.length;

    // إنشاء مصفوفة المسافات
    const dp = Array(m + 1).fill(null).map(() => Array(n + 1).fill(0));

    // تهيئة الصف والعمود الأول
    for (let i = 0; i <= m; i++) dp[i][0] = i;
    for (let j = 0; j <= n; j++) dp[0][j] = j;

    // حساب المسافات
    for (let i = 1; i <= m; i++) {
        for (let j = 1; j <= n; j++) {
            const cost = str1[i - 1] === str2[j - 1] ? 0 : 1;
            dp[i][j] = Math.min(
                dp[i - 1][j] + 1,      // حذف
                dp[i][j - 1] + 1,      // إضافة
                dp[i - 1][j - 1] + cost // استبدال
            );
        }
    }

    return dp[m][n];
}

/**
 * حساب نسبة التشابه بين نصين (0-1)
 */
function calculateSimilarity(str1, str2) {
    const s1 = normalizeArabicText(str1);
    const s2 = normalizeArabicText(str2);

    if (s1 === s2) return 1.0;
    if (!s1 || !s2) return 0.0;

    const distance = levenshteinDistance(s1, s2);
    const maxLength = Math.max(s1.length, s2.length);

    return 1 - (distance / maxLength);
}

/**
 * التحقق من احتواء النص على الكلمة المبحوث عنها
 */
function containsMatch(text, searchText) {
    const normalizedText = normalizeArabicText(text);
    const normalizedSearch = normalizeArabicText(searchText);

    return normalizedText.includes(normalizedSearch);
}

/**
 * فحص تطابق البحث الذكي للنصوص العربية
 * يستخدم 5 مستويات من التطابق (100/80/70/60/40)
 * @param {string} name - النص المراد مقارنته
 * @param {string} searchText - نص البحث
 * @returns {Object} - { match: boolean, score: number }
 */
function smartArabicMatch(name, searchText) {
    var normalizedName = normalizeArabicText(name);
    var normalizedSearch = normalizeArabicText(searchText);

    // 1️⃣ تطابق تام
    if (normalizedName === normalizedSearch) {
        return { match: true, score: 100 };
    }

    // 2️⃣ الاسم يحتوي على نص البحث كاملاً
    if (normalizedName.includes(normalizedSearch)) {
        return { match: true, score: 80 };
    }

    // 3️⃣ تطابق الاسم الأول
    var nameParts = normalizedName.split(' ');
    var searchParts = normalizedSearch.split(' ');
    if (nameParts[0] === searchParts[0]) {
        return { match: true, score: 70 };
    }

    // 4️⃣ أي جزء من الاسم يتطابق مع البحث
    for (var p = 0; p < nameParts.length; p++) {
        if (nameParts[p].includes(normalizedSearch) || normalizedSearch.includes(nameParts[p])) {
            return { match: true, score: 60 };
        }
    }

    // 5️⃣ تطابق جزئي (حرفين على الأقل)
    if (normalizedSearch.length >= 2) {
        for (var i = 0; i <= normalizedName.length - 2; i++) {
            var chunk = normalizedName.substring(i, i + Math.min(normalizedSearch.length, normalizedName.length - i));
            if (chunk.includes(normalizedSearch.substring(0, 2))) {
                return { match: true, score: 40 };
            }
        }
    }

    return { match: false, score: 0 };
}

/**
 * البحث الذكي في قائمة
 * يعيد النتائج مرتبة حسب نسبة التشابه
 */
function fuzzySearch(items, searchText, options = {}) {
    const {
        minScore = BOT_CONFIG.FUZZY_SEARCH.MIN_MATCH_SCORE,
        maxResults = BOT_CONFIG.FUZZY_SEARCH.MAX_RESULTS,
        keys = ['name'] // المفاتيح التي سيتم البحث فيها
    } = options;

    const normalizedSearch = normalizeArabicText(searchText);

    if (!normalizedSearch) return [];

    const results = [];

    items.forEach(item => {
        let bestScore = 0;
        let matchedValue = '';

        keys.forEach(key => {
            const value = typeof item === 'string' ? item : item[key];
            if (!value) return;

            // التحقق من التطابق الجزئي أولاً (أسرع)
            if (containsMatch(value, searchText)) {
                bestScore = Math.max(bestScore, 0.9);
                matchedValue = value;
            } else {
                // حساب التشابه
                const similarity = calculateSimilarity(value, searchText);
                if (similarity > bestScore) {
                    bestScore = similarity;
                    matchedValue = value;
                }
            }

            // التحقق من تطابق بداية الكلمة
            const normalizedValue = normalizeArabicText(value);
            if (normalizedValue.startsWith(normalizedSearch)) {
                bestScore = Math.max(bestScore, 0.95);
                matchedValue = value;
            }
        });

        if (bestScore >= minScore) {
            results.push({
                item: item,
                score: bestScore,
                matchedValue: matchedValue
            });
        }
    });

    // ترتيب حسب النتيجة
    results.sort((a, b) => b.score - a.score);

    // إرجاع أفضل النتائج
    return results.slice(0, maxResults).map(r => r.item);
}

// ==================== دوال البحث في قواعد البيانات ====================

/**
 * البحث في المشاريع
 */
function searchProjects(searchText) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.PROJECTS);

    if (!sheet) return [];

    const data = sheet.getDataRange().getValues();
    const projects = [];

    // تجاهل الصف الأول (العناوين)
    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const code = row[0]; // كود المشروع
        const name = row[1]; // اسم المشروع

        if (code && name) {
            projects.push({
                code: code,
                name: name
            });
        }
    }

    return fuzzySearch(projects, searchText, { keys: ['name', 'code'] });
}

/**
 * الحصول على مشروع بالكود
 */
function getProjectByCode(code) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.PROJECTS);

    if (!sheet) return null;

    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
        if (data[i][0] === code) {
            return {
                code: data[i][0],
                name: data[i][1]
            };
        }
    }

    return null;
}

/**
 * البحث في البنود
 */
function searchItems(searchText) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.ITEMS);

    if (!sheet) return [];

    const data = sheet.getDataRange().getValues();
    const items = [];

    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const name = row[0]; // اسم البند

        if (name) {
            items.push({
                name: name,
                nature: row[1] || '',       // طبيعة الحركة
                classification: row[2] || '' // تصنيف الحركة
            });
        }
    }

    return fuzzySearch(items, searchText, { keys: ['name'] });
}

/**
 * البحث في الأطراف (موردين/عملاء/ممولين)
 * يبحث في شيت الأطراف الرئيسي + شيت أطراف البوت
 * @param {string} searchText - نص البحث
 * @param {Object} [options] - خيارات البحث
 * @param {string} [options.partyType] - فلترة حسب نوع الطرف (مورد/عميل/ممول)
 * @param {number} [options.maxResults=10] - أقصى عدد نتائج
 * @param {boolean} [options.useSmartMatch=false] - استخدام smartArabicMatch بدل fuzzySearch
 * @returns {Array} - قائمة الأطراف المطابقة
 */
function searchParties(searchText, options) {
    var opts = options || {};
    var partyType = opts.partyType || null;
    var maxResults = opts.maxResults || 10;
    var useSmartMatch = opts.useSmartMatch || false;

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var addedNames = {};
    var parties = [];

    // جمع الأطراف من الشيتات
    var sheetNames = [CONFIG.SHEETS.PARTIES];
    if (CONFIG.SHEETS.BOT_PARTIES) sheetNames.push(CONFIG.SHEETS.BOT_PARTIES);

    for (var s = 0; s < sheetNames.length; s++) {
        var sheet = ss.getSheetByName(sheetNames[s]);
        if (!sheet) continue;

        var data = sheet.getDataRange().getValues();
        for (var i = 1; i < data.length; i++) {
            var row = data[i];
            var name = String(row[0] || '').trim();
            var type = String(row[1] || '').trim();

            if (!name) continue;
            if (partyType && type !== partyType) continue;

            // تجنب التكرار
            var normalizedKey = normalizeArabicText(name);
            if (addedNames[normalizedKey]) continue;
            addedNames[normalizedKey] = true;

            parties.push({
                name: name,
                type: type || 'مورد'
            });
        }
    }

    if (useSmartMatch) {
        // استخدام smartArabicMatch (للبوت - نتائج أوسع)
        var results = [];
        for (var j = 0; j < parties.length; j++) {
            var matchResult = smartArabicMatch(parties[j].name, searchText);
            if (matchResult.match) {
                results.push({
                    name: parties[j].name,
                    type: parties[j].type,
                    score: matchResult.score
                });
            }
        }
        results.sort(function(a, b) {
            if (b.score !== a.score) return b.score - a.score;
            return a.name.localeCompare(b.name, 'ar');
        });
        return results.slice(0, maxResults);
    }

    // استخدام fuzzySearch (الافتراضي)
    return fuzzySearch(parties, searchText, { keys: ['name'], maxResults: maxResults });
}

/**
 * التحقق من وجود طرف بالاسم
 */
function partyExists(partyName) {
    const results = searchParties(partyName);
    if (results.length === 0) return false;

    const normalizedSearch = normalizeArabicText(partyName);
    const normalizedResult = normalizeArabicText(results[0].name);

    return normalizedSearch === normalizedResult;
}

/**
 * الحصول على جميع المشاريع
 */
function getAllProjects() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.PROJECTS);

    if (!sheet) return [];

    const data = sheet.getDataRange().getValues();
    const projects = [];

    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        if (row[0] && row[1]) {
            projects.push({
                code: row[0],
                name: row[1]
            });
        }
    }

    return projects;
}

/**
 * الحصول على جميع البنود
 */
function getAllItems() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.ITEMS);

    if (!sheet) return [];

    const data = sheet.getDataRange().getValues();
    const items = [];

    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        if (row[0]) {
            items.push({
                name: row[0],
                nature: row[1] || '',
                classification: row[2] || ''
            });
        }
    }

    return items;
}

/**
 * الحصول على جميع الأطراف
 */
function getAllParties() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.PARTIES);

    if (!sheet) return [];

    const data = sheet.getDataRange().getValues();
    const parties = [];

    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        if (row[0]) {
            parties.push({
                name: row[0],
                type: row[1] || 'مورد'
            });
        }
    }

    return parties;
}

// ==================== دوال مساعدة للبوت ====================

/**
 * إنشاء أزرار Inline Keyboard من نتائج البحث
 */
function createSearchResultsKeyboard(results, callbackPrefix, displayKey) {
    const buttons = results.map(item => {
        const displayText = typeof item === 'string' ? item : item[displayKey || 'name'];
        const callbackData = typeof item === 'string' ? item : (item.code || item.name);

        return [{
            text: displayText,
            callback_data: `${callbackPrefix}_${callbackData}`
        }];
    });

    buttons.push([{ text: '❌ إلغاء', callback_data: 'cancel' }]);

    return { inline_keyboard: buttons };
}

/**
 * اقتراحات البحث الذكي
 */
function getSearchSuggestions(searchText, type) {
    switch (type) {
        case 'project':
            return searchProjects(searchText);
        case 'item':
            return searchItems(searchText);
        case 'party':
            return searchParties(searchText);
        default:
            return [];
    }
}

// ==================== اختبار البحث ====================

/**
 * دالة اختبار البحث
 */
function testFuzzySearch() {
    const testCases = [
        { text: 'احمد', expected: 'أحمد' },
        { text: 'شركه', expected: 'شركة' },
        { text: 'الانتاج', expected: 'الإنتاج' },
        { text: 'مونتاج', expected: 'مونتاج' }
    ];

    Logger.log('=== اختبار البحث الذكي ===');

    // اختبار تطبيع النص
    Logger.log('\n--- تطبيع النص ---');
    testCases.forEach(tc => {
        const normalized = normalizeArabicText(tc.text);
        const expectedNormalized = normalizeArabicText(tc.expected);
        const match = normalized === expectedNormalized;
        Logger.log(`${tc.text} -> ${normalized} (${match ? '✓' : '✗'})`);
    });

    // اختبار البحث في المشاريع
    Logger.log('\n--- البحث في المشاريع ---');
    const projectResults = searchProjects('فيلم');
    Logger.log('البحث عن "فيلم": ' + JSON.stringify(projectResults));

    // اختبار البحث في البنود
    Logger.log('\n--- البحث في البنود ---');
    const itemResults = searchItems('مونتاج');
    Logger.log('البحث عن "مونتاج": ' + JSON.stringify(itemResults));

    // اختبار البحث في الأطراف
    Logger.log('\n--- البحث في الأطراف ---');
    const partyResults = searchParties('شركة');
    Logger.log('البحث عن "شركة": ' + JSON.stringify(partyResults));
}
