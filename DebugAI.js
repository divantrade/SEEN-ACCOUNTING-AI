
/**
 * سرد الموديلات المتاحة من Gemini API
 * استخدم هذه الدالة لمعرفة الموديلات الصحيحة عند ظهور خطأ 404
 */
function listAvailableModels() {
    Logger.log('═══════════════════════════════════════');
    Logger.log('=== قائمة موديلات Gemini المتاحة ===');
    Logger.log('═══════════════════════════════════════');

    try {
        const apiKey = getGeminiApiKey();
        // استخدام endpoint v1beta لسرد الموديلات
        const url = `https://generativelanguage.googleapis.com/v1beta/models?key=${apiKey}`;

        const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
        const data = JSON.parse(response.getContentText());

        if (data.models) {
            Logger.log('✅ تم جلب القائمة بنجاح:');
            data.models.forEach(model => {
                // تصفية وعرض الموديلات القابلة للتوليد (generateContent)
                if (model.supportedGenerationMethods && model.supportedGenerationMethods.includes('generateContent')) {
                    Logger.log(`📌 الموديل: ${model.name}`);
                    Logger.log(`   الوصف: ${model.displayName}`);
                    Logger.log(`   الإصدار: ${model.version}`);
                    Logger.log('-----------------------------------');
                }
            });
        } else {
            Logger.log('❌ لم يتم العثور على موديلات أو حدث خطأ:');
            Logger.log(JSON.stringify(data, null, 2));
        }

    } catch (error) {
        Logger.log('❌ خطأ في جلب الموديلات: ' + error.message);
    }
}

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
