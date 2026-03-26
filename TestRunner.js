// ==================== إطار عمل الاختبارات التلقائية ====================
/**
 * إطار اختبار خفيف لبيئة Google Apps Script
 * يدعم التأكيدات (assertions)، تجميع الاختبارات، وتقارير النتائج
 *
 * الاستخدام:
 *   شغّل دالة runAllTests() من محرر Apps Script
 *   أو شغّل مجموعة محددة مثل runFuzzySearchTests()
 */

// ==================== محرك الاختبارات ====================

/**
 * كائن تسجيل نتائج الاختبارات
 */
var TestResults_ = {
    total: 0,
    passed: 0,
    failed: 0,
    errors: [],
    currentSuite: ''
};

/**
 * إعادة تعيين النتائج
 */
function resetTestResults_() {
    TestResults_.total = 0;
    TestResults_.passed = 0;
    TestResults_.failed = 0;
    TestResults_.errors = [];
    TestResults_.currentSuite = '';
}

/**
 * تعيين اسم مجموعة الاختبارات الحالية
 */
function describe_(suiteName) {
    TestResults_.currentSuite = suiteName;
    Logger.log('\n📦 ' + suiteName);
    Logger.log('─'.repeat(40));
}

// ==================== دوال التأكيد (Assertions) ====================

/**
 * التحقق من تساوي قيمتين
 */
function assertEqual_(actual, expected, message) {
    TestResults_.total++;
    if (actual === expected) {
        TestResults_.passed++;
        Logger.log('  ✅ ' + message);
    } else {
        TestResults_.failed++;
        var error = '  ❌ ' + message + ' | متوقع: ' + JSON.stringify(expected) + ' | فعلي: ' + JSON.stringify(actual);
        Logger.log(error);
        TestResults_.errors.push(TestResults_.currentSuite + ' → ' + message);
    }
}

/**
 * التحقق من أن القيمة true
 */
function assertTrue_(value, message) {
    assertEqual_(!!value, true, message);
}

/**
 * التحقق من أن القيمة false
 */
function assertFalse_(value, message) {
    assertEqual_(!!value, false, message);
}

/**
 * التحقق من تساوي تقريبي للأرقام العشرية
 */
function assertAlmostEqual_(actual, expected, tolerance, message) {
    TestResults_.total++;
    if (Math.abs(actual - expected) <= (tolerance || 0.01)) {
        TestResults_.passed++;
        Logger.log('  ✅ ' + message);
    } else {
        TestResults_.failed++;
        var error = '  ❌ ' + message + ' | متوقع: ~' + expected + ' | فعلي: ' + actual;
        Logger.log(error);
        TestResults_.errors.push(TestResults_.currentSuite + ' → ' + message);
    }
}

/**
 * التحقق من أن القيمة null أو undefined
 */
function assertNull_(value, message) {
    TestResults_.total++;
    if (value === null || value === undefined) {
        TestResults_.passed++;
        Logger.log('  ✅ ' + message);
    } else {
        TestResults_.failed++;
        var error = '  ❌ ' + message + ' | متوقع: null | فعلي: ' + JSON.stringify(value);
        Logger.log(error);
        TestResults_.errors.push(TestResults_.currentSuite + ' → ' + message);
    }
}

/**
 * التحقق من أن القيمة ليست null أو undefined
 */
function assertNotNull_(value, message) {
    TestResults_.total++;
    if (value !== null && value !== undefined) {
        TestResults_.passed++;
        Logger.log('  ✅ ' + message);
    } else {
        TestResults_.failed++;
        var error = '  ❌ ' + message + ' | القيمة null/undefined بشكل غير متوقع';
        Logger.log(error);
        TestResults_.errors.push(TestResults_.currentSuite + ' → ' + message);
    }
}

/**
 * التحقق من أن القيمة تحتوي على نص معين
 */
function assertContains_(haystack, needle, message) {
    TestResults_.total++;
    if (String(haystack).indexOf(needle) !== -1) {
        TestResults_.passed++;
        Logger.log('  ✅ ' + message);
    } else {
        TestResults_.failed++;
        var error = '  ❌ ' + message + ' | "' + needle + '" غير موجود في النص';
        Logger.log(error);
        TestResults_.errors.push(TestResults_.currentSuite + ' → ' + message);
    }
}

/**
 * التحقق من التساوي العميق للكائنات والمصفوفات
 */
function assertDeepEqual_(actual, expected, message) {
    TestResults_.total++;
    if (JSON.stringify(actual) === JSON.stringify(expected)) {
        TestResults_.passed++;
        Logger.log('  ✅ ' + message);
    } else {
        TestResults_.failed++;
        var error = '  ❌ ' + message + ' | متوقع: ' + JSON.stringify(expected) + ' | فعلي: ' + JSON.stringify(actual);
        Logger.log(error);
        TestResults_.errors.push(TestResults_.currentSuite + ' → ' + message);
    }
}

/**
 * التحقق من أن دالة ترمي خطأ
 */
function assertThrows_(fn, message) {
    TestResults_.total++;
    try {
        fn();
        TestResults_.failed++;
        Logger.log('  ❌ ' + message + ' | لم يتم رمي خطأ');
        TestResults_.errors.push(TestResults_.currentSuite + ' → ' + message);
    } catch (e) {
        TestResults_.passed++;
        Logger.log('  ✅ ' + message);
    }
}

// ==================== طباعة التقرير ====================

/**
 * طباعة ملخص نتائج الاختبارات
 */
function printTestReport_() {
    Logger.log('\n' + '═'.repeat(50));
    Logger.log('📊 ملخص نتائج الاختبارات');
    Logger.log('═'.repeat(50));
    Logger.log('  المجموع: ' + TestResults_.total);
    Logger.log('  ✅ نجح:  ' + TestResults_.passed);
    Logger.log('  ❌ فشل:  ' + TestResults_.failed);

    if (TestResults_.failed > 0) {
        Logger.log('\n⚠️ الاختبارات الفاشلة:');
        TestResults_.errors.forEach(function(err) {
            Logger.log('  • ' + err);
        });
    }

    var status = TestResults_.failed === 0 ? '✅ جميع الاختبارات نجحت!' : '❌ بعض الاختبارات فشلت';
    Logger.log('\n' + status);
    Logger.log('═'.repeat(50));

    return {
        total: TestResults_.total,
        passed: TestResults_.passed,
        failed: TestResults_.failed,
        success: TestResults_.failed === 0
    };
}

// ==================== تشغيل جميع الاختبارات ====================

/**
 * تشغيل جميع مجموعات الاختبارات
 * شغّل هذه الدالة من محرر Apps Script لتشغيل كل الاختبارات
 */
function runAllTests() {
    resetTestResults_();

    Logger.log('🚀 بدء تشغيل جميع الاختبارات...');
    Logger.log('📅 ' + new Date().toLocaleString('ar-EG'));
    Logger.log('═'.repeat(50));

    // 1. اختبارات البحث الذكي
    testFuzzySearchSuite_();

    // 2. اختبارات تحليل JSON من AI
    testAIParserSuite_();

    // 3. اختبارات حسابات العملات
    testCurrencySuite_();

    // 4. اختبارات التحقق من المدخلات
    testValidationSuite_();

    // 5. اختبارات تنسيق الأرقام والنصوص
    testFormattingSuite_();

    // 6. اختبارات تاريخ ومدخلات التاريخ
    testDateParsingSuite_();

    // 7. اختبارات نوع الحركة المالية
    testMovementTypeSuite_();

    // 8. اختبارات الإعدادات
    testConfigSuite_();

    // طباعة التقرير النهائي
    return printTestReport_();
}
