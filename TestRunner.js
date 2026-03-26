// ==================== إطار عمل الاختبارات - SEEN Test Runner ====================
// يوفر دوال التحقق (assertions) وتتبع النتائج وطباعة التقارير
// يُستخدم مع Tests.js الذي يحتوي على مجموعات الاختبارات الفعلية

/**
 * كائن تتبع نتائج الاختبارات
 */
var TestResults_ = {
  total: 0,
  passed: 0,
  failed: 0,
  errors: [],
  currentSuite: '',
  currentTest: ''
};

/**
 * إعادة تهيئة نتائج الاختبارات
 */
function resetTestResults_() {
  TestResults_.total = 0;
  TestResults_.passed = 0;
  TestResults_.failed = 0;
  TestResults_.errors = [];
  TestResults_.currentSuite = '';
  TestResults_.currentTest = '';
}

/**
 * تعريف مجموعة اختبارات
 * @param {string} suiteName - اسم المجموعة
 * @param {Function} fn - دالة تحتوي على الاختبارات
 */
function describe_(suiteName, fn) {
  TestResults_.currentSuite = suiteName;
  Logger.log('\n📋 ' + suiteName);
  Logger.log('─'.repeat(50));
  try {
    fn();
  } catch (e) {
    TestResults_.total++;
    TestResults_.failed++;
    TestResults_.errors.push({
      suite: suiteName,
      test: 'Suite Error',
      message: e.message || String(e)
    });
    Logger.log('  ❌ Suite error: ' + (e.message || String(e)));
  }
}

/**
 * التحقق من تساوي قيمتين
 * @param {*} actual - القيمة الفعلية
 * @param {*} expected - القيمة المتوقعة
 * @param {string} testName - اسم الاختبار
 */
function assertEqual_(actual, expected, testName) {
  TestResults_.total++;
  TestResults_.currentTest = testName;
  if (actual === expected) {
    TestResults_.passed++;
    Logger.log('  ✅ ' + testName);
  } else {
    TestResults_.failed++;
    var msg = 'Expected "' + expected + '" but got "' + actual + '"';
    TestResults_.errors.push({
      suite: TestResults_.currentSuite,
      test: testName,
      message: msg
    });
    Logger.log('  ❌ ' + testName + ' → ' + msg);
  }
}

/**
 * التحقق من أن القيمة true
 * @param {*} value - القيمة
 * @param {string} testName - اسم الاختبار
 */
function assertTrue_(value, testName) {
  TestResults_.total++;
  TestResults_.currentTest = testName;
  if (value === true) {
    TestResults_.passed++;
    Logger.log('  ✅ ' + testName);
  } else {
    TestResults_.failed++;
    var msg = 'Expected true but got "' + value + '"';
    TestResults_.errors.push({
      suite: TestResults_.currentSuite,
      test: testName,
      message: msg
    });
    Logger.log('  ❌ ' + testName + ' → ' + msg);
  }
}

/**
 * التحقق من أن القيمة false
 * @param {*} value - القيمة
 * @param {string} testName - اسم الاختبار
 */
function assertFalse_(value, testName) {
  TestResults_.total++;
  TestResults_.currentTest = testName;
  if (value === false) {
    TestResults_.passed++;
    Logger.log('  ✅ ' + testName);
  } else {
    TestResults_.failed++;
    var msg = 'Expected false but got "' + value + '"';
    TestResults_.errors.push({
      suite: TestResults_.currentSuite,
      test: testName,
      message: msg
    });
    Logger.log('  ❌ ' + testName + ' → ' + msg);
  }
}

/**
 * التحقق من تساوي تقريبي (للأرقام العشرية)
 * @param {number} actual - القيمة الفعلية
 * @param {number} expected - القيمة المتوقعة
 * @param {string} testName - اسم الاختبار
 * @param {number} [tolerance=0.01] - هامش الخطأ المسموح
 */
function assertAlmostEqual_(actual, expected, testName, tolerance) {
  tolerance = tolerance || 0.01;
  TestResults_.total++;
  TestResults_.currentTest = testName;
  if (Math.abs(actual - expected) <= tolerance) {
    TestResults_.passed++;
    Logger.log('  ✅ ' + testName);
  } else {
    TestResults_.failed++;
    var msg = 'Expected ~' + expected + ' (±' + tolerance + ') but got ' + actual;
    TestResults_.errors.push({
      suite: TestResults_.currentSuite,
      test: testName,
      message: msg
    });
    Logger.log('  ❌ ' + testName + ' → ' + msg);
  }
}

/**
 * التحقق من أن القيمة null
 * @param {*} value - القيمة
 * @param {string} testName - اسم الاختبار
 */
function assertNull_(value, testName) {
  TestResults_.total++;
  TestResults_.currentTest = testName;
  if (value === null || value === undefined) {
    TestResults_.passed++;
    Logger.log('  ✅ ' + testName);
  } else {
    TestResults_.failed++;
    var msg = 'Expected null/undefined but got "' + value + '"';
    TestResults_.errors.push({
      suite: TestResults_.currentSuite,
      test: testName,
      message: msg
    });
    Logger.log('  ❌ ' + testName + ' → ' + msg);
  }
}

/**
 * التحقق من أن القيمة ليست null
 * @param {*} value - القيمة
 * @param {string} testName - اسم الاختبار
 */
function assertNotNull_(value, testName) {
  TestResults_.total++;
  TestResults_.currentTest = testName;
  if (value !== null && value !== undefined) {
    TestResults_.passed++;
    Logger.log('  ✅ ' + testName);
  } else {
    TestResults_.failed++;
    var msg = 'Expected non-null value but got ' + value;
    TestResults_.errors.push({
      suite: TestResults_.currentSuite,
      test: testName,
      message: msg
    });
    Logger.log('  ❌ ' + testName + ' → ' + msg);
  }
}

/**
 * التحقق من أن النص يحتوي على جزء معين
 * @param {string} haystack - النص الكامل
 * @param {string} needle - الجزء المطلوب
 * @param {string} testName - اسم الاختبار
 */
function assertContains_(haystack, needle, testName) {
  TestResults_.total++;
  TestResults_.currentTest = testName;
  if (typeof haystack === 'string' && haystack.indexOf(needle) !== -1) {
    TestResults_.passed++;
    Logger.log('  ✅ ' + testName);
  } else {
    TestResults_.failed++;
    var msg = 'Expected string to contain "' + needle + '"';
    TestResults_.errors.push({
      suite: TestResults_.currentSuite,
      test: testName,
      message: msg
    });
    Logger.log('  ❌ ' + testName + ' → ' + msg);
  }
}

/**
 * التحقق من تساوي كائنين (مقارنة عميقة)
 * @param {*} actual - القيمة الفعلية
 * @param {*} expected - القيمة المتوقعة
 * @param {string} testName - اسم الاختبار
 */
function assertDeepEqual_(actual, expected, testName) {
  TestResults_.total++;
  TestResults_.currentTest = testName;
  var actualStr = JSON.stringify(actual);
  var expectedStr = JSON.stringify(expected);
  if (actualStr === expectedStr) {
    TestResults_.passed++;
    Logger.log('  ✅ ' + testName);
  } else {
    TestResults_.failed++;
    var msg = 'Deep equal failed.\n    Expected: ' + expectedStr + '\n    Actual:   ' + actualStr;
    TestResults_.errors.push({
      suite: TestResults_.currentSuite,
      test: testName,
      message: msg
    });
    Logger.log('  ❌ ' + testName + ' → ' + msg);
  }
}

/**
 * التحقق من أن الدالة ترمي خطأ
 * @param {Function} fn - الدالة المراد اختبارها
 * @param {string} testName - اسم الاختبار
 */
function assertThrows_(fn, testName) {
  TestResults_.total++;
  TestResults_.currentTest = testName;
  try {
    fn();
    TestResults_.failed++;
    var msg = 'Expected function to throw an error but it did not';
    TestResults_.errors.push({
      suite: TestResults_.currentSuite,
      test: testName,
      message: msg
    });
    Logger.log('  ❌ ' + testName + ' → ' + msg);
  } catch (e) {
    TestResults_.passed++;
    Logger.log('  ✅ ' + testName);
  }
}

/**
 * طباعة تقرير نتائج الاختبارات
 */
function printTestReport_() {
  Logger.log('\n' + '═'.repeat(60));
  Logger.log('📊 تقرير نتائج الاختبارات');
  Logger.log('═'.repeat(60));
  Logger.log('  إجمالي الاختبارات: ' + TestResults_.total);
  Logger.log('  ✅ ناجح: ' + TestResults_.passed);
  Logger.log('  ❌ فاشل: ' + TestResults_.failed);

  if (TestResults_.total > 0) {
    var pct = Math.round((TestResults_.passed / TestResults_.total) * 100);
    Logger.log('  📈 نسبة النجاح: ' + pct + '%');
  }

  if (TestResults_.errors.length > 0) {
    Logger.log('\n⚠️ تفاصيل الأخطاء:');
    Logger.log('─'.repeat(50));
    for (var i = 0; i < TestResults_.errors.length; i++) {
      var err = TestResults_.errors[i];
      Logger.log('  [' + err.suite + '] ' + err.test);
      Logger.log('    → ' + err.message);
    }
  }

  Logger.log('═'.repeat(60));
  if (TestResults_.failed === 0) {
    Logger.log('🎉 جميع الاختبارات ناجحة!');
  } else {
    Logger.log('⚠️ يوجد ' + TestResults_.failed + ' اختبار فاشل يحتاج مراجعة');
  }
}

/**
 * تشغيل جميع مجموعات الاختبارات
 * هذه الدالة الرئيسية التي يمكن تشغيلها من محرر Apps Script
 */
function runAllTests() {
  resetTestResults_();

  Logger.log('🚀 بدء تشغيل جميع الاختبارات...');
  Logger.log('التوقيت: ' + new Date().toISOString());
  Logger.log('');

  // تشغيل كل مجموعات الاختبارات
  testFuzzySearchSuite_();
  testAIParserSuite_();
  testCurrencySuite_();
  testValidationSuite_();
  testFormattingSuite_();
  testDateParsingSuite_();
  testMovementTypeSuite_();
  testConfigSuite_();

  // طباعة التقرير النهائي
  printTestReport_();

  return TestResults_;
}

/**
 * تشغيل اختبارات البحث الذكي فقط
 */
function runFuzzySearchTests() {
  resetTestResults_();
  testFuzzySearchSuite_();
  printTestReport_();
  return TestResults_;
}

/**
 * تشغيل اختبارات محلل الذكاء الاصطناعي فقط
 */
function runAIParserTests() {
  resetTestResults_();
  testAIParserSuite_();
  printTestReport_();
  return TestResults_;
}

/**
 * تشغيل اختبارات العملات فقط
 */
function runCurrencyTests() {
  resetTestResults_();
  testCurrencySuite_();
  printTestReport_();
  return TestResults_;
}

/**
 * تشغيل اختبارات أنواع الحركات فقط
 */
function runMovementTypeTests() {
  resetTestResults_();
  testMovementTypeSuite_();
  printTestReport_();
  return TestResults_;
}
