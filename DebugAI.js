// ==================== دوال تشخيص البوت الذكي ====================
/**
 * ملف تشخيص مشاكل البوت الذكي
 * ملاحظة: دالة listAvailableModels موجودة في AIAgent.js
 */

/**
 * اختبار جميع الموديلات المحتملة لمعرفة أيها يعمل (تشخيص عميق)
 */
function testAllModels() {
    Logger.log('═══════════════════════════════════════');
    Logger.log('=== 🕵️‍♂️ تشخيص اتصال Gemini العميق ===');
    Logger.log('═══════════════════════════════════════');

    // 1. فحص المفتاح
    let apiKey = '';
    try {
        apiKey = getGeminiApiKey();
        if (!apiKey) throw new Error('المفتاح فارغ');
        Logger.log(`🔑 حالة المفتاح: ✅ موجود (ينتهي بـ ...${apiKey.slice(-4)})`);
    } catch (e) {
        Logger.log(`⛔ خطا حرج: لم يتم العثور على مفتاح API! السبب: ${e.message}`);
        return;
    }

    const modelsToTest = [
        'gemini-1.5-flash',
        'gemini-1.5-flash-latest',
        'gemini-1.5-flash-001',
        'gemini-1.5-flash-002',
        'gemini-1.5-pro',
        'gemini-pro'
    ];

    const payload = {
        contents: [{ parts: [{ text: "Hello" }] }]
    };

    // 2. تجربة الموديلات
    modelsToTest.forEach(modelName => {
        const url = `https://generativelanguage.googleapis.com/v1beta/models/${modelName}:generateContent?key=${apiKey}`;

        Logger.log(`🔄 تجربة: ${modelName}`);

        try {
            const response = UrlFetchApp.fetch(url, {
                method: 'post',
                contentType: 'application/json',
                payload: JSON.stringify(payload),
                muteHttpExceptions: true
            });

            const code = response.getResponseCode();
            const text = response.getContentText();

            if (code === 200) {
                Logger.log(`✅ نـجــاح! (${modelName}) يعمل.`);
            } else {
                Logger.log(`❌ فشل (${code}):`);
                // محاولة استخراج رسالة الخطأ من JSON
                try {
                    const json = JSON.parse(text);
                    if (json.error) {
                        Logger.log(`   الرسالة: ${json.error.message}`);
                        Logger.log(`   الحالة: ${json.error.status}`);
                    } else {
                        Logger.log(`   الرد الخام: ${text.substring(0, 200)}`);
                    }
                } catch (e) {
                    Logger.log(`   الرد الخام: ${text.substring(0, 200)}`);
                }
            }
        } catch (e) {
            Logger.log(`💥 خطأ تنفيذ: ${e.message}`);
        }
        Logger.log('-----------------------------------');
    });

    Logger.log('📝 انتهى التشخيص. انسخ هذا السجل وأرسله للمطور.');
}

/**
 * اختبار الكتابة المباشرة في الشيت (تشخيص الذاكرة)
 */
function testSaveToSheet() {
    Logger.log('═══════════════════════════════════════');
    Logger.log('=== 💾 اختبار الكتابة في الشيت ===');
    Logger.log('═══════════════════════════════════════');

    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const sheetName = CONFIG.SHEETS.BOT_TRANSACTIONS; // 'حركات البوت'
        Logger.log(`📂 البحث عن الشيت: "${sheetName}"`);

        let sheet = ss.getSheetByName(sheetName);

        if (!sheet) {
            Logger.log(`❌ الشيت غير موجود! هل الاسم صحيح؟`);
            Logger.log(`الشيتات الموجودة في الملف:`);
            ss.getSheets().forEach(s => Logger.log(`- ${s.getName()}`));
            return;
        }

        Logger.log(`✅ الشيت موجود. محاولة إضافة صف تجريبي...`);

        // صف تجريبي
        const debugRow = [
            'TEST-001', '2025-01-01', 'تجربة', 'تجربة', '', 'مشروع تجريبي',
            'بند تجريبي', 'تفاصيل تجريبية من التشخيص', 'طرف تجريبي',
            100, 'USD', 1, 100, 'مدين', '', '', 'نقدي', 'فوري',
            '', '', '2025-01-01', 'معلق', '2025-01', 'Test Note', '',
            'قيد الانتظار', 'Admin', '123456', '2025-01-01 12:00:00',
            '', '', '', '', 'لا', 'تجربة يدوية'
        ];

        sheet.appendRow(debugRow);
        Logger.log(`✅ تم تنفيذ appendRow بنجاح!`);
        Logger.log(`🎉 اذهب للشيت وتأكد من ظهور صف جديد يبدأ بـ TEST-001`);

    } catch (error) {
        Logger.log(`❌ فشل ذريع أثناء الكتابة: ${error.message}`);
        Logger.log(error.stack);
    }
}
