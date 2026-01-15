
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
 * اختبار جميع الموديلات المحتملة لمعرفة أيها يعمل
 */
function testAllModels() {
    Logger.log('═══════════════════════════════════════');
    Logger.log('=== اختبار اتصال الموديلات (Live Test) ===');
    Logger.log('═══════════════════════════════════════');

    const modelsToTest = [
        'gemini-1.5-flash',
        'gemini-1.5-flash-latest',
        'gemini-1.5-flash-001',
        'gemini-1.5-flash-002',
        'gemini-1.5-flash-8b',
        'gemini-1.5-pro',
        'gemini-1.0-pro',
        'gemini-pro'
    ];

    const apiKey = getGeminiApiKey();
    const payload = {
        contents: [{ parts: [{ text: "Hello" }] }]
    };

    modelsToTest.forEach(modelName => {
        const url = `https://generativelanguage.googleapis.com/v1beta/models/${modelName}:generateContent?key=${apiKey}`;

        Logger.log(`🔄 تجربة الموديل: ${modelName}...`);

        try {
            const response = UrlFetchApp.fetch(url, {
                method: 'post',
                contentType: 'application/json',
                payload: JSON.stringify(payload),
                muteHttpExceptions: true
            });

            const code = response.getResponseCode();
            if (code === 200) {
                Logger.log(`✅ نـجــاح! الموديل ${modelName} يعمل.`);
            } else {
                Logger.log(`❌ فشل (${code}): ${response.getContentText().substring(0, 100)}...`);
            }
        } catch (e) {
            Logger.log(`❌ خطأ في التنفيذ: ${e.message}`);
        }
        Logger.log('-----------------------------------');
    });
}
