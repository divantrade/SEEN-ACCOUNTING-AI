// ==================== ملف الاختبارات الشاملة ====================
/**
 * جميع مجموعات اختبارات نظام SEEN المحاسبي
 * يعتمد على TestRunner.js للتأكيدات والتقارير
 *
 * التشغيل: شغّل runAllTests() من محرر Apps Script
 */

// ==================== 1. اختبارات البحث الذكي (FuzzySearch) ====================

function testFuzzySearchSuite_() {
    describe_('البحث الذكي - تطبيع النص العربي');

    // تطبيع الهمزات
    assertEqual_(normalizeArabicText('أحمد'), normalizeArabicText('احمد'), 'توحيد همزة الألف (أ → ا)');
    assertEqual_(normalizeArabicText('إبراهيم'), normalizeArabicText('ابراهيم'), 'توحيد همزة الألف السفلية (إ → ا)');
    assertEqual_(normalizeArabicText('آمنة'), normalizeArabicText('امنه'), 'توحيد الألف المدّة (آ → ا)');

    // التاء المربوطة والهاء
    assertEqual_(normalizeArabicText('شركة'), normalizeArabicText('شركه'), 'توحيد التاء المربوطة (ة → ه)');
    assertEqual_(normalizeArabicText('مؤسسة الإنتاج'), normalizeArabicText('موسسه الانتاج'), 'توحيد مركّب');

    // الألف المقصورة والياء
    assertEqual_(normalizeArabicText('مستشفى'), normalizeArabicText('مستشفي'), 'توحيد الألف المقصورة (ى → ي)');

    // إزالة التشكيل
    assertEqual_(normalizeArabicText('مُحَمَّد'), normalizeArabicText('محمد'), 'إزالة التشكيل');

    // إزالة المسافات الزائدة
    assertEqual_(normalizeArabicText('أحمد   نايل'), normalizeArabicText('احمد نايل'), 'إزالة المسافات الزائدة');

    // حالات حدية
    assertEqual_(normalizeArabicText(''), '', 'نص فارغ');
    assertEqual_(normalizeArabicText(null), '', 'قيمة null');
    assertEqual_(normalizeArabicText(undefined), '', 'قيمة undefined');
    assertEqual_(normalizeArabicText(123), '123', 'رقم بدلاً من نص');

    // نصوص إنجليزية مختلطة
    assertEqual_(normalizeArabicText('Ahmed Test'), 'ahmed test', 'تحويل الإنجليزي لأحرف صغيرة');

    // --- اختبارات Levenshtein ---
    describe_('البحث الذكي - مسافة Levenshtein');

    assertEqual_(levenshteinDistance('', ''), 0, 'نصان فارغان = 0');
    assertEqual_(levenshteinDistance('abc', 'abc'), 0, 'نصان متطابقان = 0');
    assertEqual_(levenshteinDistance('abc', 'abd'), 1, 'حرف واحد مختلف = 1');
    assertEqual_(levenshteinDistance('abc', 'ab'), 1, 'حذف حرف = 1');
    assertEqual_(levenshteinDistance('abc', 'abcd'), 1, 'إضافة حرف = 1');
    assertEqual_(levenshteinDistance('kitten', 'sitting'), 3, 'kitten → sitting = 3');
    assertEqual_(levenshteinDistance('', 'abc'), 3, 'نص فارغ ونص من 3 = 3');

    // --- اختبارات calculateSimilarity ---
    describe_('البحث الذكي - نسبة التشابه');

    assertAlmostEqual_(calculateSimilarity('أحمد', 'أحمد'), 1.0, 0.01, 'نفس الاسم = 1.0');
    assertAlmostEqual_(calculateSimilarity('أحمد', 'احمد'), 1.0, 0.01, 'نفس الاسم بعد التطبيع = 1.0');
    assertAlmostEqual_(calculateSimilarity('', ''), 0.0, 0.01, 'نصان فارغان = 0.0');
    assertTrue_(calculateSimilarity('مونتاج', 'مونتاچ') > 0.7, 'كلمات متشابهة > 0.7');
    assertTrue_(calculateSimilarity('تصوير', 'برمجة') < 0.5, 'كلمات مختلفة < 0.5');

    // --- اختبارات containsMatch ---
    describe_('البحث الذكي - التطابق الجزئي');

    assertTrue_(containsMatch('شركة الإنتاج الإعلامي', 'انتاج'), 'يحتوي على الكلمة');
    assertTrue_(containsMatch('أحمد نايل', 'احمد'), 'يحتوي على الاسم الأول');
    assertFalse_(containsMatch('محمود', 'أحمد'), 'لا يحتوي على اسم مختلف');
    assertFalse_(containsMatch('', 'بحث'), 'نص فارغ لا يحتوي شيء');
}

// ==================== 2. اختبارات تحليل JSON من AI ====================

function testAIParserSuite_() {
    describe_('تحليل JSON - تنظيف النصوص');

    // cleanJsonString_
    assertEqual_(
        cleanJsonString_('{"a": 1,}'),
        '{"a": 1}',
        'إزالة فاصلة زائدة قبل }'
    );
    assertEqual_(
        cleanJsonString_('[1, 2, 3,]'),
        '[1, 2, 3]',
        'إزالة فاصلة زائدة قبل ]'
    );
    assertEqual_(
        cleanJsonString_('{"a": 1} // تعليق'),
        '{"a": 1} ',
        'إزالة تعليق سطري'
    );
    assertEqual_(
        cleanJsonString_(null),
        null,
        'null يعود null'
    );

    // --- parseGeminiResponse ---
    describe_('تحليل JSON - استخراج من رد Gemini');

    var jsonResult1 = parseGeminiResponse('```json\n{"success": true, "amount": 500}\n```');
    assertTrue_(jsonResult1.success === true, 'استخراج JSON من code block');
    assertEqual_(jsonResult1.amount, 500, 'استخراج المبلغ من code block');

    var jsonResult2 = parseGeminiResponse('هذا نص عادي {"nature": "دفعة مصروف"} ونص آخر');
    assertEqual_(jsonResult2.nature, 'دفعة مصروف', 'استخراج JSON من نص مختلط');

    var jsonResult3 = parseGeminiResponse('');
    assertTrue_(jsonResult3.success === false, 'نص فارغ يعيد خطأ');

    var jsonResult4 = parseGeminiResponse('لا يوجد JSON هنا');
    assertTrue_(jsonResult4.success === false, 'نص بدون JSON يعيد خطأ');

    // --- fixUnclosedStrings_ ---
    describe_('تحليل JSON - إصلاح النصوص المفتوحة');

    assertEqual_(
        fixUnclosedStrings_('{"name": "أحمد'),
        '{"name": "أحمد"',
        'إغلاق نص غير مغلق'
    );
    assertEqual_(
        fixUnclosedStrings_('{"a": "b"}'),
        '{"a": "b"}',
        'نص مغلق بالفعل لا يتغير'
    );

    // --- balanceBrackets_ ---
    describe_('تحليل JSON - موازنة الأقواس');

    assertEqual_(
        balanceBrackets_('{"a": [1, 2'),
        '{"a": [1, 2]}',
        'إغلاق ] و } ناقصتين'
    );
    assertEqual_(
        balanceBrackets_('{"a": 1}'),
        '{"a": 1}',
        'أقواس متوازنة لا تتغير'
    );
}

// ==================== 3. اختبارات حسابات العملات ====================

function testCurrencySuite_() {
    describe_('العملات - حساب القيمة بالدولار');

    assertEqual_(calculateUSDAmount(100, 'USD', 1), 100, 'دولار بدون تحويل');
    assertAlmostEqual_(calculateUSDAmount(3800, 'TRY', 38), 100, 0.01, '3800 ليرة / 38 = 100 دولار');
    assertAlmostEqual_(calculateUSDAmount(5000, 'EGP', 50), 100, 0.01, '5000 جنيه / 50 = 100 دولار');
    assertEqual_(calculateUSDAmount(500, 'USD', null), 500, 'دولار مع سعر صرف null');

    // --- أسعار الصرف الافتراضية ---
    describe_('العملات - أسعار الصرف الافتراضية');

    assertEqual_(getDefaultExchangeRate('USD'), 1.0, 'سعر الدولار = 1');
    assertEqual_(getDefaultExchangeRate('TRY'), 38.0, 'سعر الليرة الافتراضي = 38');
    assertEqual_(getDefaultExchangeRate('EGP'), 50.0, 'سعر الجنيه الافتراضي = 50');
    assertEqual_(getDefaultExchangeRate('XXX'), 1.0, 'عملة غير معروفة = 1');

    // --- تنسيق الأرقام ---
    describe_('العملات - تنسيق الأرقام');

    assertEqual_(formatNumber(1234.56), '1,234.56', 'تنسيق بفواصل');
    assertEqual_(formatNumber(0), '0', 'صفر');
    assertEqual_(formatNumber(null), '0', 'null = 0');
    assertEqual_(formatNumber(1000), '1,000.00', 'عدد صحيح مع خانتين عشريتين');
}

// ==================== 4. اختبارات التحقق من المدخلات ====================

function testValidationSuite_() {
    describe_('التحقق - كشف الحركات الجديدة');

    // looksLikeNewTransaction_
    assertTrue_(looksLikeNewTransaction_('دفعت 500 دولار لأحمد'), 'رقم + كلمة مفتاحية = حركة');
    assertTrue_(looksLikeNewTransaction_('استحقاق 1000 ليرة'), 'استحقاق + رقم = حركة');
    assertTrue_(looksLikeNewTransaction_('مصاريف بنكية 15 دولار'), 'مصاريف بنكية = حركة');
    assertFalse_(looksLikeNewTransaction_('أحمد نايل'), 'اسم فقط ≠ حركة');
    assertFalse_(looksLikeNewTransaction_('مرحبا'), 'تحية ≠ حركة');
    assertFalse_(looksLikeNewTransaction_(''), 'نص فارغ ≠ حركة');
    assertFalse_(looksLikeNewTransaction_(null), 'null ≠ حركة');

    // --- sanitizeMarkdown_ ---
    describe_('التحقق - تنظيف Markdown');

    assertEqual_(sanitizeMarkdown_('*نص عادي*'), '*نص عادي*', 'نجمتان متوازنتان لا تتغير');
    assertEqual_(sanitizeMarkdown_('نص *غير مغلق'), 'نص غير مغلق', 'نجمة واحدة تُزال');
    assertEqual_(sanitizeMarkdown_('`كود`'), '`كود`', 'backticks متوازنة لا تتغير');
    assertEqual_(sanitizeMarkdown_('نص `غير مغلق'), 'نص غير مغلق', 'backtick واحدة تُزال');
    assertEqual_(sanitizeMarkdown_(null), null, 'null يعود null');
    assertEqual_(sanitizeMarkdown_(''), '', 'فارغ يعود فارغ');
}

// ==================== 5. اختبارات التنسيق ====================

function testFormattingSuite_() {
    describe_('التنسيق - إيموجي نوع الحركة');

    assertEqual_(getTransactionEmoji('استحقاق مصروف'), '📤', 'إيموجي استحقاق مصروف');
    assertEqual_(getTransactionEmoji('دفعة مصروف'), '💸', 'إيموجي دفعة مصروف');
    assertEqual_(getTransactionEmoji('استحقاق إيراد'), '📥', 'إيموجي استحقاق إيراد');
    assertEqual_(getTransactionEmoji('تحصيل إيراد'), '💰', 'إيموجي تحصيل إيراد');
    assertEqual_(getTransactionEmoji('تمويل'), '🏦', 'إيموجي تمويل');
    assertEqual_(getTransactionEmoji('نوع غير معروف'), '📋', 'إيموجي افتراضي');

    // --- getTypeLabel ---
    describe_('التنسيق - عنوان نوع الحركة');

    assertEqual_(getTypeLabel('دفعة مصروف'), 'دفعة مصروف', 'طبيعة موجودة');
    assertEqual_(getTypeLabel(null), 'حركة مالية', 'null = حركة مالية');
    assertEqual_(getTypeLabel(''), 'حركة مالية', 'فارغ = حركة مالية');

    // --- buildReportFileName ---
    describe_('التنسيق - أسماء ملفات التقارير');

    var fileName = buildReportFileName('statement', 'أحمد نايل');
    assertContains_(fileName, 'كشف حساب', 'اسم الملف يحتوي نوع التقرير');
    assertContains_(fileName, 'أحمد نايل', 'اسم الملف يحتوي اسم الطرف');

    var fileNameNoParty = buildReportFileName('expenses', null);
    assertContains_(fileNameNoParty, 'تقرير المصروفات', 'تقرير بدون طرف');
}

// ==================== 6. اختبارات التاريخ ====================

function testDateParsingSuite_() {
    describe_('التاريخ - تحليل المدخلات');

    // parseDateInput_ (from Utilities.js)
    var result1 = parseDateInput_('24/12/2025');
    assertTrue_(result1.success, 'صيغة dd/mm/yyyy صحيحة');
    assertEqual_(result1.date, '24/12/2025', 'التاريخ المنسق');

    var result2 = parseDateInput_('24.12.2025');
    assertTrue_(result2.success, 'صيغة dd.mm.yyyy صحيحة');

    var result3 = parseDateInput_('24-12-2025');
    assertTrue_(result3.success, 'صيغة dd-mm-yyyy صحيحة');

    var result4 = parseDateInput_('31/02/2025');
    assertFalse_(result4.success, 'تاريخ غير موجود (31 فبراير)');

    var result5 = parseDateInput_('abc');
    assertFalse_(result5.success, 'نص غير صالح');

    var result6 = parseDateInput_('01/13/2025');
    assertFalse_(result6.success, 'شهر 13 غير صالح');

    // --- parseArabicDate ---
    describe_('التاريخ - تحويل الأرقام العربية');

    assertEqual_(parseArabicDate('٢٤/١٢/٢٠٢٥'), '2025-12-24', 'أرقام عربية → إنجليزية dd/mm/yyyy');
    assertEqual_(parseArabicDate('2025-03-15'), '2025-03-15', 'صيغة ISO تبقى كما هي');
    assertEqual_(parseArabicDate('15/03/2025'), '2025-03-15', 'تحويل dd/mm/yyyy → ISO');
}

// ==================== 7. اختبارات نوع الحركة المالية ====================

function testMovementTypeSuite_() {
    describe_('نوع الحركة - استحقاق (مدين)');

    assertEqual_(getMovementTypeFromNature_('استحقاق مصروف'), CONFIG.MOVEMENT.DEBIT, 'استحقاق مصروف = مدين');
    assertEqual_(getMovementTypeFromNature_('استحقاق إيراد'), CONFIG.MOVEMENT.DEBIT, 'استحقاق إيراد = مدين');

    describe_('نوع الحركة - دفعة (دائن)');

    assertEqual_(getMovementTypeFromNature_('دفعة مصروف'), CONFIG.MOVEMENT.CREDIT, 'دفعة مصروف = دائن');
    assertEqual_(getMovementTypeFromNature_('تحصيل إيراد'), CONFIG.MOVEMENT.CREDIT, 'تحصيل إيراد = دائن');
    assertEqual_(getMovementTypeFromNature_('سداد تمويل'), CONFIG.MOVEMENT.CREDIT, 'سداد تمويل = دائن');
    assertEqual_(getMovementTypeFromNature_('استلام تمويل'), CONFIG.MOVEMENT.CREDIT, 'استلام تمويل = دائن');
    assertEqual_(getMovementTypeFromNature_('تحويل داخلي'), CONFIG.MOVEMENT.CREDIT, 'تحويل داخلي = دائن');
    assertEqual_(getMovementTypeFromNature_('تغيير عملة'), CONFIG.MOVEMENT.CREDIT, 'تغيير عملة = دائن');
    assertEqual_(getMovementTypeFromNature_('مصاريف بنكية'), CONFIG.MOVEMENT.CREDIT, 'مصاريف بنكية = دائن');

    describe_('نوع الحركة - تسوية');

    assertEqual_(getMovementTypeFromNature_('تسوية استحقاق مصروف'), CONFIG.MOVEMENT.SETTLEMENT, 'تسوية مصروف = تسوية');
    assertEqual_(getMovementTypeFromNature_('تسوية استحقاق إيراد'), CONFIG.MOVEMENT.SETTLEMENT, 'تسوية إيراد = تسوية');

    describe_('نوع الحركة - حالات خاصة');

    assertEqual_(getMovementTypeFromNature_('تمويل'), CONFIG.MOVEMENT.DEBIT, 'تمويل عام = مدين');
    // ⭐ مهم: "تمويل (دخول قرض)" = دائن وليس مدين لأنه يحتوي "دخول قرض"
    assertEqual_(getMovementTypeFromNature_('🏦 تمويل (دخول قرض)'), CONFIG.MOVEMENT.CREDIT, 'تمويل دخول قرض = دائن');
    assertNull_(getMovementTypeFromNature_(''), 'نص فارغ = null');
    assertNull_(getMovementTypeFromNature_(null), 'null = null');
}

// ==================== 8. اختبارات الإعدادات ====================

function testConfigSuite_() {
    describe_('الإعدادات - CONFIG موجود ومكتمل');

    assertNotNull_(CONFIG, 'CONFIG موجود');
    assertNotNull_(CONFIG.SHEETS, 'CONFIG.SHEETS موجود');
    assertNotNull_(CONFIG.COLORS, 'CONFIG.COLORS موجود');
    assertNotNull_(CONFIG.CURRENCIES, 'CONFIG.CURRENCIES موجود');
    assertNotNull_(CONFIG.MOVEMENT, 'CONFIG.MOVEMENT موجود');
    assertNotNull_(CONFIG.COMPANY, 'CONFIG.COMPANY موجود');

    describe_('الإعدادات - بيانات الشركة');

    assertNotNull_(CONFIG.COMPANY.NAME, 'اسم الشركة موجود');
    assertNotNull_(CONFIG.COMPANY.ADDRESS, 'عنوان الشركة موجود');
    assertNotNull_(CONFIG.COMPANY.TIMEZONE, 'المنطقة الزمنية موجودة');
    assertEqual_(CONFIG.COMPANY.TIMEZONE, 'Asia/Istanbul', 'المنطقة الزمنية صحيحة');

    describe_('الإعدادات - أنواع الحركات');

    assertEqual_(CONFIG.MOVEMENT.DEBIT, 'مدين استحقاق', 'نوع المدين');
    assertEqual_(CONFIG.MOVEMENT.CREDIT, 'دائن دفعة', 'نوع الدائن');
    assertEqual_(CONFIG.MOVEMENT.SETTLEMENT, 'دائن تسوية', 'نوع التسوية');
    assertEqual_(CONFIG.MOVEMENT.TYPES.length, 3, '3 أنواع حركات');

    describe_('الإعدادات - العملات');

    assertEqual_(CONFIG.CURRENCIES.DEFAULT, 'USD', 'العملة الافتراضية دولار');
    assertTrue_(CONFIG.CURRENCIES.LIST.length >= 3, '3 عملات على الأقل');

    describe_('الإعدادات - الشيتات');

    assertNotNull_(CONFIG.SHEETS.TRANSACTIONS, 'شيت الحركات موجود');
    assertNotNull_(CONFIG.SHEETS.PROJECTS, 'شيت المشاريع موجود');
    assertNotNull_(CONFIG.SHEETS.PARTIES, 'شيت الأطراف موجود');
    assertNotNull_(CONFIG.SHEETS.ITEMS, 'شيت البنود موجود');

    describe_('الإعدادات - وحدات الإنتاج');

    assertEqual_(CONFIG.getUnitType('تصوير'), 'مقابلة', 'وحدة التصوير = مقابلة');
    assertEqual_(CONFIG.getUnitType('مونتاج'), 'دقيقة', 'وحدة المونتاج = دقيقة');
    assertEqual_(CONFIG.getUnitType('جرافيك - رسم'), 'رسمة', 'وحدة الجرافيك = رسمة');
    assertNull_(CONFIG.getUnitType('بند غير موجود'), 'بند بدون وحدة = null');
}

// ==================== تشغيل مجموعات فردية ====================
// دوال التشغيل الفردية (runFuzzySearchTests, runAIParserTests, etc.)
// موجودة في TestRunner.js
