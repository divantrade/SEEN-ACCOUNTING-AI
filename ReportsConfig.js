/**
 * ═══════════════════════════════════════════════════════════════════════════
 *                          REPORTS CONFIGURATION
 *                      إعدادات التقارير وكشوف الحساب
 * ═══════════════════════════════════════════════════════════════════════════
 */

const REPORTS_CONFIG = {

    // ═══════════════════════════════════════════════════════════════════════
    //                           أنواع التقارير
    // ═══════════════════════════════════════════════════════════════════════

    REPORT_TYPES: {
        STATEMENT: 'statement',           // كشف حساب
        ALERTS: 'alerts',                 // تنبيهات الاستحقاق
        BALANCES: 'balances',             // تقرير الأرصدة
        PROFITABILITY: 'profitability',   // ربحية المشاريع
        EXPENSES: 'expenses',             // تقرير المصروفات
        REVENUES: 'revenues'              // تقرير الإيرادات
    },

    // ═══════════════════════════════════════════════════════════════════════
    //                           أنواع الأطراف
    // ═══════════════════════════════════════════════════════════════════════

    PARTY_TYPES: {
        VENDOR: 'مورد',
        CLIENT: 'عميل',
        FUNDER: 'ممول'
    },

    // ═══════════════════════════════════════════════════════════════════════
    //                         فترات التقارير
    // ═══════════════════════════════════════════════════════════════════════

    DATE_RANGES: {
        LAST_MONTH: 'last_month',
        LAST_3_MONTHS: 'last_3_months',
        LAST_6_MONTHS: 'last_6_months',
        THIS_YEAR: 'this_year',
        ALL_TIME: 'all_time',
        CUSTOM: 'custom'
    },

    // ═══════════════════════════════════════════════════════════════════════
    //                         إعدادات الحفظ
    // ═══════════════════════════════════════════════════════════════════════

    STORAGE: {
        // مجلد حفظ التقارير في Google Drive
        REPORTS_FOLDER_NAME: 'تقارير البوت',

        // التقارير التي يتم حفظها بشكل دائم
        SAVE_PERMANENTLY: ['statement', 'profitability', 'expenses', 'revenues'],

        // التقارير التي تُرسل مباشرة بدون حفظ
        SEND_DIRECTLY: ['alerts', 'balances'],

        // تنسيق اسم الملف
        FILE_NAME_FORMAT: '{report_type} - {party_name} - {date}.pdf'
    },

    // ═══════════════════════════════════════════════════════════════════════
    //                         إعدادات PDF
    // ═══════════════════════════════════════════════════════════════════════

    PDF_OPTIONS: {
        SIZE: 'A4',
        PORTRAIT: false,          // أفقي للتقارير العريضة
        FIT_WIDTH: true,
        GRIDLINES: false,
        PRINT_TITLE: true
    },

    // ═══════════════════════════════════════════════════════════════════════
    //                           الرسائل
    // ═══════════════════════════════════════════════════════════════════════

    MESSAGES: {
        MAIN_MENU: '📊 *التقارير وكشوف الحساب*\n\nاختر نوع التقرير المطلوب:',

        SELECT_PARTY_TYPE: '👤 *اختر نوع الطرف:*',

        ENTER_PARTY_NAME: '✍️ *اكتب اسم الطرف أو جزء منه:*\n\nمثال: أحمد أو الشركة',

        SELECT_PARTY: '📋 *اختر الطرف من القائمة:*',

        SELECT_DATE_RANGE: '📅 *اختر الفترة الزمنية:*',

        GENERATING_REPORT: '⏳ *جاري إعداد التقرير...*\n\nقد يستغرق هذا بضع ثوان.',

        // أسماء التقارير للعرض
        REPORT_NAMES: {
            'statement': 'كشف الحساب',
            'alerts': 'تقرير الاستحقاقات',
            'balances': 'تقرير الأرصدة',
            'profitability': 'تقرير ربحية المشاريع',
            'expenses': 'تقرير المصروفات',
            'revenues': 'تقرير الإيرادات'
        },

        REPORT_READY: '✅ *تم إعداد التقرير بنجاح!*',

        REPORT_SENT: '📄 تم إرسال التقرير.',

        REPORT_SAVED: '💾 تم حفظ نسخة في أرشيف التقارير.',

        NO_DATA: '⚠️ لا توجد بيانات للفترة المحددة.',

        NO_PARTIES_FOUND: '❌ لم يتم العثور على طرف بهذا الاسم.\n\nيرجى المحاولة باسم آخر.',

        ERROR: '❌ حدث خطأ أثناء إعداد التقرير.\n\nيرجى المحاولة لاحقاً.',

        NO_ALERTS: '✅ *لا توجد استحقاقات قريبة*\n\nجميع المستحقات في وضع جيد.',

        ALERTS_HEADER: '⏰ *تنبيهات الاستحقاق*\n━━━━━━━━━━━━━━━━━━━━\n\n'
    },

    // ═══════════════════════════════════════════════════════════════════════
    //                         لوحات المفاتيح
    // ═══════════════════════════════════════════════════════════════════════

    KEYBOARDS: {

        // القائمة الرئيسية للتقارير
        MAIN_MENU: {
            inline_keyboard: [
                [
                    { text: '📄 كشف حساب', callback_data: 'report_statement' },
                    { text: '💰 الأرصدة', callback_data: 'report_balances' }
                ],
                [
                    { text: '⏰ الاستحقاقات', callback_data: 'report_alerts' },
                    { text: '📈 ربحية المشاريع', callback_data: 'report_profitability' }
                ],
                [
                    { text: '📊 المصروفات', callback_data: 'report_expenses' },
                    { text: '💵 الإيرادات', callback_data: 'report_revenues' }
                ],
                [
                    { text: '❌ إلغاء', callback_data: 'report_cancel' }
                ]
            ]
        },

        // اختيار نوع الطرف
        PARTY_TYPE: {
            inline_keyboard: [
                [
                    { text: '🏪 مورد', callback_data: 'report_party_مورد' },
                    { text: '👤 عميل', callback_data: 'report_party_عميل' }
                ],
                [
                    { text: '💼 ممول', callback_data: 'report_party_ممول' }
                ],
                [
                    { text: '🔙 رجوع', callback_data: 'report_back_main' },
                    { text: '❌ إلغاء', callback_data: 'report_cancel' }
                ]
            ]
        },

        // اختيار الفترة الزمنية
        DATE_RANGE: {
            inline_keyboard: [
                [
                    { text: '📅 آخر شهر', callback_data: 'report_date_last_month' },
                    { text: '📅 آخر 3 شهور', callback_data: 'report_date_last_3_months' }
                ],
                [
                    { text: '📅 آخر 6 شهور', callback_data: 'report_date_last_6_months' },
                    { text: '📅 السنة الحالية', callback_data: 'report_date_this_year' }
                ],
                [
                    { text: '📅 كل الفترات', callback_data: 'report_date_all_time' }
                ],
                [
                    { text: '🔙 رجوع', callback_data: 'report_back_party' },
                    { text: '❌ إلغاء', callback_data: 'report_cancel' }
                ]
            ]
        },

        // فترة التنبيهات
        ALERTS_PERIOD: {
            inline_keyboard: [
                [
                    { text: '📅 اليوم', callback_data: 'report_alerts_0' },
                    { text: '📅 خلال 3 أيام', callback_data: 'report_alerts_3' }
                ],
                [
                    { text: '📅 خلال أسبوع', callback_data: 'report_alerts_7' },
                    { text: '📅 خلال شهر', callback_data: 'report_alerts_30' }
                ],
                [
                    { text: '🔙 رجوع', callback_data: 'report_back_main' },
                    { text: '❌ إلغاء', callback_data: 'report_cancel' }
                ]
            ]
        }
    },

    // ═══════════════════════════════════════════════════════════════════════
    //                         حالات المحادثة
    // ═══════════════════════════════════════════════════════════════════════

    STATES: {
        IDLE: 'report_idle',
        WAITING_PARTY_TYPE: 'report_waiting_party_type',
        WAITING_PARTY_NAME: 'report_waiting_party_name',
        WAITING_PARTY_SELECTION: 'report_waiting_party_selection',
        WAITING_DATE_RANGE: 'report_waiting_date_range',
        WAITING_ALERTS_PERIOD: 'report_waiting_alerts_period',
        GENERATING: 'report_generating'
    }
};

// للوصول من ملفات أخرى
if (typeof module !== 'undefined') {
    module.exports = REPORTS_CONFIG;
}
