
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
