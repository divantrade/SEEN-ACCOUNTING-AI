// ==================== مجموعات الاختبارات ====================
/**
 * جميع اختبارات الوحدة للنظام
 * يتم استدعاؤها من TestRunner.js عبر runAllTests()
 *
 * ⚠️ هذا الملف لا يعدّل أي كود موجود - فقط يختبر الدوال الحالية
 */

// ==================== 1. اختبارات البحث الذكي (FuzzySearch) ====================

function testFuzzySearchSuite_() {
    // --- تطبيع النص العربي ---
    describe_('تطبيع النص العربي - normalizeArabicText');

    // توحيد الهمزات
    assertEqual_(normalizeArabicText('أحمد'), normalizeArabicText('احمد'), 'توحيد الهمزة: أحمد = احمد');
    assertEqual_(normalizeArabicText('إبراهيم'), normalizeArabicText('ابراهيم'), 'توحيد الهمزة: إبراهيم = ابراهيم');
    assertEqual_(normalizeArabicText('آمال'), normalizeArabicText('امال'), 'توحيد الهمزة: آمال = امال');

    // توحيد التاء المربوطة
    assertEqual_(normalizeArabicText('شركة'), normalizeArabicText('شركه'), 'التاء المربوطة: شركة = شركه');
    assertEqual_(normalizeArabicText('مؤسسة'), normalizeArabicText('مؤسسه'), 'التاء المربوطة: مؤسسة = مؤسسه');

    // توحيد الألف المقصورة
    assertEqual_(normalizeArabicText('مصطفى'), normalizeArabicText('مصطفي'), 'الألف المقصورة: مصطفى = مصطفي');

    // إزالة التشكيل
    assertEqual_(normalizeArabicText('مُحَمَّد'), normalizeArabicText('محمد'), 'إزالة التشكيل: مُحَمَّد = محمد');

    // إزالة المسافات الزائدة
    assertEqual_(normalizeArabicText('  أحمد   نايل  '), 'احمد نايل', 'إزالة المسافات الزائدة');

    // تحويل الإنجليزي لحروف صغيرة
    assertEqual_(normalizeArabicText('Ahmed'), 'ahmed', 'تحويل الإنجليزية لحروف صغيرة');

    // الحالات الحدية
    assertEqual_(normalizeArabicText(''), '', 'نص فارغ يعيد نص فارغ');
    assertEqual_(normalizeArabicText(null), '', 'null يعيد نص فارغ');
    assertEqual_(normalizeArabicText(undefined), '', 'undefined يعيد نص فارغ');
    assertEqual_(normalizeArabicText(123), '123', 'رقم يتحول لنص');

    // --- مسافة Levenshtein ---
    describe_('مسافة Levenshtein');

    assertEqual_(levenshteinDistance('abc', 'abc'), 0, 'نصان متطابقان = مسافة 0');
    assertEqual_(levenshteinDistance('abc', 'abd'), 1, 'حرف مختلف واحد = مسافة 1');
    assertEqual_(levenshteinDistance('abc', 'ab'), 1, 'حذف حرف واحد = مسافة 1');
    assertEqual_(levenshteinDistance('abc', 'abcd'), 1, 'إضافة حرف واحد = مسافة 1');
    assertEqual_(levenshteinDistance('', ''), 0, 'نصان فارغان = مسافة 0');
    assertEqual_(levenshteinDistance('abc', ''), 3, 'نص مقابل فارغ = طول النص');
    assertEqual_(levenshteinDistance('', 'xyz'), 3, 'فارغ مقابل نص = طول النص');
    assertEqual_(levenshteinDistance('kitten', 'sitting'), 3, 'kitten → sitting = 3');

    // --- نسبة التشابه ---
    describe_('نسبة التشابه - calculateSimilarity');

    assertAlmostEqual_(calculateSimilarity('أحمد', 'أحمد'), 1.0, 0.01, 'نفس النص = تشابه 100%');
    assertAlmostEqual_(calculateSimilarity('أحمد', 'احمد'), 1.0, 0.01, 'همزة مختلفة = تشابه 100% بعد التطبيع');
    assertAlmostEqual_(calculateSimilarity('شركة', 'شركه'), 1.0, 0.01, 'تاء مربوطة = تشابه 100% بعد التطبيع');
    assertEqual_(calculateSimilarity('', ''), 0.0, 'نصان فارغان = تشابه 0%');
    assertEqual_(calculateSimilarity('أحمد', ''), 0.0, 'نص مقابل فارغ = تشابه 0%');
    assertTrue_(calculateSimilarity('محمد', 'محمود') > 0.5, 'محمد ومحمود: تشابه أكبر من 50%');

    // --- التطابق الجزئي ---
    describe_('التطابق الجزئي - containsMatch');

    assertTrue_(containsMatch('شركة الإنتاج الفني', 'إنتاج'), 'البحث عن جزء من النص');
    assertTrue_(containsMatch('شركة الإنتاج الفني', 'الانتاج'), 'البحث مع همزة مختلفة');
    assertFalse_(containsMatch('شركة الإنتاج', 'مونتاج'), 'نص غير موجود');
    assertFalse_(containsMatch('', 'بحث'), 'بحث في نص فارغ');

    // --- البحث الذكي ---
    describe_('البحث الذكي - fuzzySearch');

    var testItems = [
        { name: 'أحمد نايل' },
        { name: 'محمد علي' },
        { name: 'شركة الإنتاج الفني' },
        { name: 'مؤسسة الرؤية' }
    ];

    var results = fuzzySearch(testItems, 'احمد', { minScore: 0.5, maxResults: 5, keys: ['name'] });
    assertTrue_(results.length > 0, 'البحث عن "احمد" يجد نتائج');
    assertEqual_(results[0].name, 'أحمد نايل', 'أول نتيجة هي "أحمد نايل"');

    results = fuzzySearch(testItems, 'شركه الانتاج', { minScore: 0.5, maxResults: 5, keys: ['name'] });
    assertTrue_(results.length > 0, 'البحث عن "شركه الانتاج" يجد نتائج');

    results = fuzzySearch(testItems, '', { minScore: 0.5, maxResults: 5, keys: ['name'] });
    assertEqual_(results.length, 0, 'البحث الفارغ لا يعيد نتائج');

    // البحث في نصوص مباشرة (strings بدلاً من objects)
    var stringItems = ['تصوير', 'مونتاج', 'مكساج', 'جرافيك'];
    results = fuzzySearch(stringItems, 'مونتاج', { minScore: 0.5, maxResults: 5, keys: ['name'] });
    assertTrue_(results.length > 0, 'البحث في مصفوفة نصوص يجد نتائج');
}

// ==================== 2. اختبارات تحليل JSON من AI ====================

function testAIParserSuite_() {
    describe_('تحليل رد AI - parseGeminiResponse');

    // JSON صحيح
    var result = parseGeminiResponse('{"type": "expense", "amount": 500}');
    assertNotNull_(result, 'تحليل JSON بسيط يعيد نتيجة');
    assertEqual_(result.type, 'expense', 'استخراج type صحيح');
    assertEqual_(result.amount, 500, 'استخراج amount صحيح');

    // JSON داخل markdown code block
    result = parseGeminiResponse('```json\n{"type": "expense", "amount": 300}\n```');
    assertNotNull_(result, 'استخراج JSON من markdown code block');
    assertEqual_(result.amount, 300, 'القيمة صحيحة من code block');

    // JSON مع نص إضافي قبله وبعده
    result = parseGeminiResponse('Here is the result: {"success": true, "data": "test"} end of response');
    assertNotNull_(result, 'استخراج JSON من وسط نص');
    assertEqual_(result.success, true, 'القيمة صحيحة من وسط نص');

    // رد فارغ
    result = parseGeminiResponse('');
    assertTrue_(result.success === false, 'رد فارغ يعطي success: false');

    result = parseGeminiResponse(null);
    assertTrue_(result.success === false, 'null يعطي success: false');

    // نص بدون JSON
    result = parseGeminiResponse('لا يوجد JSON هنا');
    assertTrue_(result.success === false, 'نص بدون JSON يعطي success: false');

    // --- تنظيف JSON ---
    describe_('تنظيف JSON - cleanJsonString_');

    // فاصلة زائدة قبل }
    assertEqual_(cleanJsonString_('{"a": 1,}'), '{"a": 1}', 'إزالة فاصلة زائدة قبل }');

    // فاصلة زائدة قبل ]
    assertEqual_(cleanJsonString_('[1, 2, 3,]'), '[1, 2, 3]', 'إزالة فاصلة زائدة قبل ]');

    // تعليقات JavaScript
    var withComments = '// comment\n{"a": 1}';
    assertContains_(cleanJsonString_(withComments), '"a"', 'إزالة تعليقات السطر الواحد');

    // أحرف تحكم
    assertEqual_(cleanJsonString_('{"a":\x01"b"}'), '{"a":"b"}', 'إزالة أحرف التحكم');

    // null input
    assertNull_(cleanJsonString_(null), 'null يعيد null');

    // --- إصلاح النصوص غير المغلقة ---
    describe_('إصلاح النصوص - fixUnclosedStrings_');

    assertEqual_(fixUnclosedStrings_('{"a": "hello}'), '{"a": "hello}"', 'إغلاق نص مفتوح');
    assertEqual_(fixUnclosedStrings_('{"a": "ok"}'), '{"a": "ok"}', 'نص مغلق لا يتغير');
    assertNull_(fixUnclosedStrings_(null), 'null يعيد null');

    // --- موازنة الأقواس ---
    describe_('موازنة الأقواس - balanceBrackets_');

    assertEqual_(balanceBrackets_('{"a": [1, 2'), '{"a": [1, 2]}', 'إغلاق قوس مربع وحاصرة ناقصين');
    assertEqual_(balanceBrackets_('{"a": 1'), '{"a": 1}', 'إغلاق حاصرة ناقصة');
    assertEqual_(balanceBrackets_('{"a": 1}'), '{"a": 1}', 'أقواس متوازنة لا تتغير');
    assertNull_(balanceBrackets_(null), 'null يعيد null');
}

// ==================== 3. اختبارات حسابات العملات ====================

function testCurrencySuite_() {
    describe_('حساب القيمة بالدولار - calculateUSDAmount');

    // الدولار يبقى كما هو
    assertEqual_(calculateUSDAmount(100, 'USD', 1), 100, '100 USD = 100 USD');
    assertEqual_(calculateUSDAmount(500, 'USD', null), 500, '500 USD بدون سعر صرف = 500 USD');

    // تحويل الليرة التركية
    assertAlmostEqual_(calculateUSDAmount(380, 'TRY', 38), 10, 0.01, '380 TRY / 38 = 10 USD');
    assertAlmostEqual_(calculateUSDAmount(1000, 'TRY', 38), 26.32, 0.01, '1000 TRY / 38 = 26.32 USD');

    // تحويل الجنيه المصري
    assertAlmostEqual_(calculateUSDAmount(500, 'EGP', 50), 10, 0.01, '500 EGP / 50 = 10 USD');

    // سعر صرف افتراضي
    assertAlmostEqual_(calculateUSDAmount(380, 'TRY', null), 10, 0.01, 'TRY بسعر افتراضي 38');

    describe_('سعر الصرف الافتراضي - getDefaultExchangeRate');

    assertEqual_(getDefaultExchangeRate('USD'), 1.0, 'USD = 1.0');
    assertEqual_(getDefaultExchangeRate('TRY'), 38.0, 'TRY = 38.0');
    assertEqual_(getDefaultExchangeRate('EGP'), 50.0, 'EGP = 50.0');
    assertEqual_(getDefaultExchangeRate('XYZ'), 1.0, 'عملة غير معروفة = 1.0');
}

// ==================== 4. اختبارات التحقق من المدخلات ====================

function testValidationSuite_() {
    describe_('كشف المعاملات الجديدة - looksLikeNewTransaction_');

    // حالات يجب أن تكون true (نص مالي + رقم)
    assertTrue_(looksLikeNewTransaction_('استحقاق مصروف 500 دولار'), 'استحقاق مصروف مع مبلغ');
    assertTrue_(looksLikeNewTransaction_('دفعة مصروف 1000 ليرة'), 'دفعة مصروف مع مبلغ');
    assertTrue_(looksLikeNewTransaction_('تحصيل إيراد 200 دولار'), 'تحصيل إيراد مع مبلغ');
    assertTrue_(looksLikeNewTransaction_('مصاريف بنكية 50 دولار'), 'مصاريف بنكية مع مبلغ');
    assertTrue_(looksLikeNewTransaction_('تحويل داخلي 300 USD'), 'تحويل داخلي مع مبلغ');
    assertTrue_(looksLikeNewTransaction_('تغيير عملة 100 دولار'), 'تغيير عملة مع مبلغ');
    assertTrue_(looksLikeNewTransaction_('سداد تمويل 500'), 'سداد تمويل مع رقم');

    // حالات يجب أن تكون false
    assertFalse_(looksLikeNewTransaction_('أحمد نايل'), 'اسم شخص بدون رقم');
    assertFalse_(looksLikeNewTransaction_('شركة الإنتاج'), 'اسم شركة بدون رقم');
    assertFalse_(looksLikeNewTransaction_(''), 'نص فارغ');
    assertFalse_(looksLikeNewTransaction_(null), 'null');
    assertFalse_(looksLikeNewTransaction_('مرحبا 123'), 'رقم بدون كلمة مالية');

    // --- تنظيف Markdown ---
    describe_('تنظيف Markdown - sanitizeMarkdown_');

    // نجمات متوازنة لا تتغير
    assertEqual_(sanitizeMarkdown_('*bold*'), '*bold*', 'نجمات متوازنة لا تتغير');

    // نجمة غير متوازنة تُزال
    var result = sanitizeMarkdown_('hello * world');
    var starCount = (result.match(/\*/g) || []).length;
    assertTrue_(starCount % 2 === 0, 'نجمة وحيدة تُزال لتصبح متوازنة');

    // شرطات سفلية متوازنة
    assertEqual_(sanitizeMarkdown_('_italic_'), '_italic_', 'شرطات سفلية متوازنة لا تتغير');

    // backticks متوازنة
    assertEqual_(sanitizeMarkdown_('`code`'), '`code`', 'backticks متوازنة لا تتغير');

    // null/empty
    assertNull_(sanitizeMarkdown_(null), 'null يعيد null');
    assertEqual_(sanitizeMarkdown_(''), '', 'فارغ يعيد فارغ');
}

// ==================== 5. اختبارات تنسيق الأرقام والنصوص ====================

function testFormattingSuite_() {
    describe_('تنسيق الأرقام - formatNumber');

    assertEqual_(formatNumber(1000), '1,000.00', '1000 → 1,000.00');
    assertEqual_(formatNumber(0), '0', 'صفر يعيد "0"');
    assertEqual_(formatNumber(null), '0', 'null يعيد "0"');
    assertEqual_(formatNumber(99.5), '99.50', '99.5 → 99.50');
    assertEqual_(formatNumber(1234567.89), '1,234,567.89', 'رقم كبير مع فواصل');

    describe_('إيموجي نوع الحركة - getTransactionEmoji');

    assertEqual_(getTransactionEmoji('استحقاق مصروف'), '📤', 'استحقاق مصروف = 📤');
    assertEqual_(getTransactionEmoji('دفعة مصروف'), '💸', 'دفعة مصروف = 💸');
    assertEqual_(getTransactionEmoji('استحقاق إيراد'), '📥', 'استحقاق إيراد = 📥');
    assertEqual_(getTransactionEmoji('تحصيل إيراد'), '💰', 'تحصيل إيراد = 💰');
    assertEqual_(getTransactionEmoji('تمويل'), '🏦', 'تمويل = 🏦');
    assertEqual_(getTransactionEmoji('سداد تمويل'), '💳', 'سداد تمويل = 💳');
    assertEqual_(getTransactionEmoji('نوع غير معروف'), '📋', 'نوع غير معروف = 📋 (افتراضي)');

    describe_('عنوان نوع الحركة - getTypeLabel');

    assertEqual_(getTypeLabel('استحقاق مصروف'), 'استحقاق مصروف', 'يعيد نفس النص');
    assertEqual_(getTypeLabel(''), 'حركة مالية', 'نص فارغ = حركة مالية');
    assertEqual_(getTypeLabel(null), 'حركة مالية', 'null = حركة مالية');

    describe_('بناء ملخص المعاملة - buildTransactionSummary');

    var testTransaction = {
        nature: 'استحقاق مصروف',
        project: 'فيلم وثائقي',
        project_code: 'P001',
        classification: 'مصروفات إنتاج',
        item: 'تصوير',
        party: 'أحمد نايل',
        amount: 500,
        currency: 'USD',
        due_date: '15/01/2026',
        details: 'تصوير مقابلات'
    };

    var summary = buildTransactionSummary(testTransaction);
    assertContains_(summary, 'فيلم وثائقي', 'الملخص يحتوي على اسم المشروع');
    assertContains_(summary, 'P001', 'الملخص يحتوي على كود المشروع');
    assertContains_(summary, 'أحمد نايل', 'الملخص يحتوي على اسم الطرف');
    assertContains_(summary, '500.00', 'الملخص يحتوي على المبلغ');
    assertContains_(summary, 'تصوير مقابلات', 'الملخص يحتوي على التفاصيل');

    // معاملة بعملة غير USD
    var tryTransaction = {
        nature: 'دفعة مصروف',
        classification: 'مصروفات',
        amount: 3800,
        currency: 'TRY',
        exchangeRate: 38,
        due_date: '20/01/2026'
    };

    summary = buildTransactionSummary(tryTransaction);
    assertContains_(summary, 'TRY', 'الملخص يحتوي على العملة');
    assertContains_(summary, 'USD', 'الملخص يحتوي على القيمة بالدولار');
}

// ==================== 6. اختبارات التاريخ ====================

function testDateParsingSuite_() {
    describe_('تحويل التاريخ العربي - parseArabicDate');

    // تاريخ بصيغة dd/mm/yyyy
    assertEqual_(parseArabicDate('15/01/2026'), '2026-01-15', 'dd/mm/yyyy → yyyy-mm-dd');
    assertEqual_(parseArabicDate('1/2/2025'), '2025-02-01', 'd/m/yyyy مع padding');

    // تاريخ بصيغة dd-mm-yyyy
    assertEqual_(parseArabicDate('15-01-2026'), '2026-01-15', 'dd-mm-yyyy → yyyy-mm-dd');

    // تاريخ بصيغة yyyy-mm-dd
    assertEqual_(parseArabicDate('2026-01-15'), '2026-01-15', 'yyyy-mm-dd يبقى كما هو');

    // تاريخ بصيغة yyyy/mm/dd
    assertEqual_(parseArabicDate('2026/01/15'), '2026-01-15', 'yyyy/mm/dd → yyyy-mm-dd');

    // أرقام عربية
    assertEqual_(parseArabicDate('١٥/٠١/٢٠٢٦'), '2026-01-15', 'أرقام عربية تتحول لإنجليزية');
    assertEqual_(parseArabicDate('٢٠٢٦-٠١-١٥'), '2026-01-15', 'أرقام عربية بشرطة');

    describe_('تحليل مدخلات التاريخ - parseDateInput_');

    // تاريخ صحيح
    var result = parseDateInput_('24/12/2025');
    assertTrue_(result.success, 'تاريخ صحيح: 24/12/2025');
    assertEqual_(result.date, '24/12/2025', 'التاريخ المنسق صحيح');

    // فواصل مختلفة
    result = parseDateInput_('24.12.2025');
    assertTrue_(result.success, 'فاصل نقطة: 24.12.2025');

    result = parseDateInput_('24-12-2025');
    assertTrue_(result.success, 'فاصل شرطة: 24-12-2025');

    // يوم بدون padding
    result = parseDateInput_('1/2/2025');
    assertTrue_(result.success, 'يوم/شهر بدون صفر: 1/2/2025');
    assertEqual_(result.date, '01/02/2025', 'padding تلقائي');

    // شهر خاطئ
    result = parseDateInput_('15/13/2025');
    assertFalse_(result.success, 'شهر 13 غير صحيح');

    // يوم خاطئ
    result = parseDateInput_('32/01/2025');
    assertFalse_(result.success, 'يوم 32 غير صحيح');

    // تاريخ غير موجود (30 فبراير)
    result = parseDateInput_('30/02/2025');
    assertFalse_(result.success, '30 فبراير غير موجود');

    // صيغة خاطئة
    result = parseDateInput_('not a date');
    assertFalse_(result.success, 'نص عشوائي = صيغة غير صحيحة');
}

// ==================== 7. اختبارات نوع الحركة المالية ====================

function testMovementTypeSuite_() {
    describe_('تحديد نوع الحركة - getMovementTypeFromNature_');

    // مدين استحقاق
    assertEqual_(getMovementTypeFromNature_('استحقاق مصروف'), CONFIG.MOVEMENT.DEBIT,
        'استحقاق مصروف = مدين استحقاق');
    assertEqual_(getMovementTypeFromNature_('استحقاق إيراد'), CONFIG.MOVEMENT.DEBIT,
        'استحقاق إيراد = مدين استحقاق');

    // دائن دفعة
    assertEqual_(getMovementTypeFromNature_('دفعة مصروف'), CONFIG.MOVEMENT.CREDIT,
        'دفعة مصروف = دائن دفعة');
    assertEqual_(getMovementTypeFromNature_('تحصيل إيراد'), CONFIG.MOVEMENT.CREDIT,
        'تحصيل إيراد = دائن دفعة');
    assertEqual_(getMovementTypeFromNature_('سداد تمويل'), CONFIG.MOVEMENT.CREDIT,
        'سداد تمويل = دائن دفعة');
    assertEqual_(getMovementTypeFromNature_('تأمين مدفوع'), CONFIG.MOVEMENT.CREDIT,
        'تأمين مدفوع = دائن دفعة');
    assertEqual_(getMovementTypeFromNature_('استلام تمويل'), CONFIG.MOVEMENT.CREDIT,
        'استلام تمويل = دائن دفعة');
    assertEqual_(getMovementTypeFromNature_('تحويل داخلي'), CONFIG.MOVEMENT.CREDIT,
        'تحويل داخلي = دائن دفعة');
    assertEqual_(getMovementTypeFromNature_('تغيير عملة'), CONFIG.MOVEMENT.CREDIT,
        'تغيير عملة = دائن دفعة');
    assertEqual_(getMovementTypeFromNature_('مصاريف بنكية'), CONFIG.MOVEMENT.CREDIT,
        'مصاريف بنكية = دائن دفعة');
    assertEqual_(getMovementTypeFromNature_('استرداد تأمين'), CONFIG.MOVEMENT.CREDIT,
        'استرداد تأمين = دائن دفعة');

    // دائن تسوية
    assertEqual_(getMovementTypeFromNature_('تسوية استحقاق مصروف'), CONFIG.MOVEMENT.SETTLEMENT,
        'تسوية استحقاق مصروف = دائن تسوية');
    assertEqual_(getMovementTypeFromNature_('تسوية استحقاق إيراد'), CONFIG.MOVEMENT.SETTLEMENT,
        'تسوية استحقاق إيراد = دائن تسوية');

    // تمويل عام (بدون سداد/استلام/دخول قرض) = مدين
    assertEqual_(getMovementTypeFromNature_('تمويل'), CONFIG.MOVEMENT.DEBIT,
        'تمويل عام = مدين استحقاق');

    // 🏦 تمويل (دخول قرض) = دائن دفعة (نقدية فعلية)
    assertEqual_(getMovementTypeFromNature_('🏦 تمويل (دخول قرض)'), CONFIG.MOVEMENT.CREDIT,
        'تمويل (دخول قرض) = دائن دفعة');

    // حالات حدية
    assertNull_(getMovementTypeFromNature_(''), 'نص فارغ = null');
    assertNull_(getMovementTypeFromNature_(null), 'null = null');
    assertNull_(getMovementTypeFromNature_('نوع غير معروف'), 'نوع غير معروف = null');

    // الترتيب مهم: "تسوية استحقاق مصروف" يحتوي على "مصروف" لكن يجب أن يكون تسوية
    assertEqual_(getMovementTypeFromNature_('تسوية استحقاق مصروف'), CONFIG.MOVEMENT.SETTLEMENT,
        'أولوية التسوية على الدفعة (تحتوي على "مصروف")');
}

// ==================== 8. اختبارات الإعدادات ====================

function testConfigSuite_() {
    describe_('كائن الإعدادات - CONFIG');

    // التحقق من وجود الشيتات الأساسية
    assertNotNull_(CONFIG.SHEETS.TRANSACTIONS, 'شيت الحركات المالية موجود');
    assertNotNull_(CONFIG.SHEETS.PROJECTS, 'شيت المشاريع موجود');
    assertNotNull_(CONFIG.SHEETS.PARTIES, 'شيت الأطراف موجود');
    assertNotNull_(CONFIG.SHEETS.ITEMS, 'شيت البنود موجود');

    // التحقق من العملات
    assertTrue_(CONFIG.CURRENCIES.LIST.length > 0, 'قائمة العملات غير فارغة');
    assertEqual_(CONFIG.CURRENCIES.DEFAULT, 'USD', 'العملة الافتراضية USD');
    assertNotNull_(CONFIG.CURRENCIES.SYMBOLS.USD, 'رمز الدولار موجود');
    assertNotNull_(CONFIG.CURRENCIES.SYMBOLS.TRY, 'رمز الليرة موجود');
    assertNotNull_(CONFIG.CURRENCIES.SYMBOLS.EGP, 'رمز الجنيه موجود');

    // التحقق من أنواع الحركات
    assertEqual_(CONFIG.MOVEMENT.TYPES.length, 3, '3 أنواع حركات');
    assertContains_(CONFIG.MOVEMENT.TYPES.join(','), 'مدين استحقاق', 'يحتوي على مدين استحقاق');
    assertContains_(CONFIG.MOVEMENT.TYPES.join(','), 'دائن دفعة', 'يحتوي على دائن دفعة');
    assertContains_(CONFIG.MOVEMENT.TYPES.join(','), 'دائن تسوية', 'يحتوي على دائن تسوية');

    // التحقق من أنواع الأطراف
    assertEqual_(CONFIG.PARTY_TYPES.LIST.length, 3, '3 أنواع أطراف');

    // التحقق من طبيعة الحركة
    assertTrue_(CONFIG.NATURE_TYPES.length >= 10, 'أكثر من 10 أنواع طبيعة حركة');

    // التحقق من وحدات الإنتاج
    assertEqual_(CONFIG.getUnitType('تصوير'), 'مقابلة', 'وحدة التصوير = مقابلة');
    assertEqual_(CONFIG.getUnitType('مونتاج'), 'دقيقة', 'وحدة المونتاج = دقيقة');
    assertNull_(CONFIG.getUnitType('بند غير موجود'), 'بند غير موجود = null');
    assertTrue_(CONFIG.getItemsWithUnits().length > 0, 'قائمة البنود مع وحدات غير فارغة');

    // التحقق من طرق الدفع
    assertTrue_(CONFIG.PAYMENT_METHODS.length >= 4, 'أكثر من 4 طرق دفع');

    // التحقق من الألوان
    assertNotNull_(CONFIG.COLORS.HEADER.TRANSACTIONS, 'لون هيدر الحركات موجود');
    assertTrue_(CONFIG.COLORS.HEADER.TRANSACTIONS.startsWith('#'), 'اللون يبدأ بـ #');
}
