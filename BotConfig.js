// ==================== إعدادات بوت تليجرام التفصيلية ====================
/**
 * ملف إعدادات بوت تليجرام لنظام SEEN المحاسبي
 * يحتوي على هيكل الشيتات الجديدة وإعدادات المحادثة
 */

const BOT_CONFIG = {

    // ==================== هيكل شيت حركات البوت ====================
    BOT_TRANSACTIONS_COLUMNS: {
        // الأعمدة الأساسية (مطابقة لدفتر الحركات)
        TRANSACTION_ID: { index: 1, name: 'رقم الحركة', width: 80 },
        DATE: { index: 2, name: 'التاريخ', width: 100 },
        NATURE: { index: 3, name: 'طبيعة الحركة', width: 120 },
        CLASSIFICATION: { index: 4, name: 'تصنيف الحركة', width: 120 },
        PROJECT_CODE: { index: 5, name: 'كود المشروع', width: 100 },
        PROJECT_NAME: { index: 6, name: 'اسم المشروع', width: 150 },
        ITEM: { index: 7, name: 'البند', width: 120 },
        DETAILS: { index: 8, name: 'التفاصيل', width: 200 },
        PARTY_NAME: { index: 9, name: 'اسم المورد/الجهة', width: 150 },
        AMOUNT: { index: 10, name: 'المبلغ بالعملة الأصلية', width: 120 },
        CURRENCY: { index: 11, name: 'العملة', width: 80 },
        EXCHANGE_RATE: { index: 12, name: 'سعر الصرف', width: 100 },
        AMOUNT_USD: { index: 13, name: 'القيمة بالدولار', width: 120 },
        MOVEMENT_TYPE: { index: 14, name: 'نوع الحركة', width: 100 },
        BALANCE: { index: 15, name: 'الرصيد', width: 100 },
        REFERENCE: { index: 16, name: 'رقم مرجعي', width: 120 },
        PAYMENT_METHOD: { index: 17, name: 'طريقة الدفع', width: 100 },
        PAYMENT_TERM_TYPE: { index: 18, name: 'نوع شرط الدفع', width: 100 },
        WEEKS: { index: 19, name: 'عدد الأسابيع', width: 80 },
        CUSTOM_DATE: { index: 20, name: 'تاريخ مخصص', width: 100 },
        DUE_DATE: { index: 21, name: 'تاريخ الاستحقاق', width: 100 },
        PAYMENT_STATUS: { index: 22, name: 'حالة السداد', width: 100 },
        MONTH: { index: 23, name: 'الشهر', width: 80 },
        NOTES: { index: 24, name: 'ملاحظات', width: 150 },
        STATEMENT_LINK: { index: 25, name: '📄 كشف', width: 80 },

        // الأعمدة الإضافية للبوت
        REVIEW_STATUS: { index: 26, name: 'حالة المراجعة', width: 100 },
        TELEGRAM_USER: { index: 27, name: 'المُدخل (تليجرام)', width: 150 },
        TELEGRAM_CHAT_ID: { index: 28, name: 'معرّف المحادثة', width: 120 },
        INPUT_TIMESTAMP: { index: 29, name: 'تاريخ الإدخال', width: 150 },
        REVIEWER: { index: 30, name: 'المُراجع', width: 120 },
        REVIEW_TIMESTAMP: { index: 31, name: 'تاريخ المراجعة', width: 150 },
        REVIEW_NOTES: { index: 32, name: 'ملاحظات المراجعة', width: 200 },
        ATTACHMENT_URL: { index: 33, name: 'رابط المرفق', width: 200 },
        IS_NEW_PARTY: { index: 34, name: 'طرف جديد؟', width: 80 },
        INPUT_SOURCE: { index: 35, name: 'مصدر الإدخال', width: 120 }
    },

    // ==================== هيكل شيت أطراف البوت ====================
    BOT_PARTIES_COLUMNS: {
        PARTY_NAME: { index: 1, name: 'اسم الطرف', width: 200 },
        PARTY_TYPE: { index: 2, name: 'نوع الطرف', width: 100 },
        SPECIALIZATION: { index: 3, name: 'التخصص / الفئة', width: 150 },
        PHONE: { index: 4, name: 'رقم الهاتف', width: 120 },
        EMAIL: { index: 5, name: 'البريد الإلكتروني', width: 180 },
        LOCATION: { index: 6, name: 'المدينة / الدولة', width: 120 },
        PAYMENT_METHOD: { index: 7, name: 'طريقة الدفع المفضلة', width: 130 },
        BANK_DETAILS: { index: 8, name: 'بيانات الحساب البنكي', width: 200 },
        NOTES: { index: 9, name: 'ملاحظات', width: 200 },

        // أعمدة إضافية للبوت
        REVIEW_STATUS: { index: 10, name: 'حالة المراجعة', width: 100 },
        TELEGRAM_USER: { index: 11, name: 'المُدخل (تليجرام)', width: 150 },
        TELEGRAM_CHAT_ID: { index: 12, name: 'معرّف المحادثة', width: 120 },
        INPUT_TIMESTAMP: { index: 13, name: 'تاريخ الإدخال', width: 150 },
        REVIEWER: { index: 14, name: 'المُراجع', width: 120 },
        REVIEW_TIMESTAMP: { index: 15, name: 'تاريخ المراجعة', width: 150 },
        LINKED_TRANSACTIONS: { index: 16, name: 'الحركات المرتبطة', width: 150 }
    },

    // ==================== هيكل شيت المستخدمين المصرح لهم (موحد) ====================
    BOT_USERS_COLUMNS: {
        NAME: { index: 1, name: 'الاسم', width: 150 },
        EMAIL: { index: 2, name: 'الإيميل', width: 200 },
        PHONE: { index: 3, name: 'رقم الهاتف', width: 150 },
        TELEGRAM_USERNAME: { index: 4, name: 'اسم المستخدم تليجرام', width: 150 },
        TELEGRAM_CHAT_ID: { index: 5, name: 'معرّف المحادثة', width: 120 },
        PERM_TRADITIONAL_BOT: { index: 6, name: 'بوت تقليدي', width: 100 },
        PERM_AI_BOT: { index: 7, name: 'بوت ذكي', width: 100 },
        PERM_SHEET: { index: 8, name: 'شيت', width: 80 },
        PERM_REVIEW: { index: 9, name: 'مراجعة', width: 80 },
        IS_ACTIVE: { index: 10, name: 'نشط', width: 80 },
        ADDED_DATE: { index: 11, name: 'تاريخ الإضافة', width: 120 },
        NOTES: { index: 12, name: 'ملاحظات', width: 200 }
    },

    // أنواع الصلاحيات (للتوافق مع الكود القديم)
    USER_TYPES: {
        BOT: 'بوت',
        SHEET: 'شيت',
        BOTH: 'كلاهما'
    },

    // ==================== أوامر البوت ====================
    COMMANDS: {
        START: '/start',
        EXPENSE: '/مصروف',
        REVENUE: '/ايراد',
        FINANCING: '/تمويل',
        INSURANCE: '/تأمين',
        TRANSFER: '/تحويل',
        STATUS: '/حالة',
        HELP: '/مساعدة',
        CANCEL: '/الغاء'
    },

    // قائمة الأوامر للظهور في زر القائمة (Menu Button)
    // ملاحظة: Telegram يتطلب أسماء الأوامر بالإنجليزية فقط
    COMMAND_LIST: [
        { command: 'start', description: '👋 بدء استخدام البوت' },
        { command: 'expense', description: '📤 تسجيل مصروف جديد' },
        { command: 'revenue', description: '📥 تسجيل إيراد جديد' },
        { command: 'financing', description: '🏦 تسجيل تمويل (قرض/سداد)' },
        { command: 'insurance', description: '🔐 تسجيل تأمين (دفع/استرداد)' },
        { command: 'transfer', description: '🔄 تسجيل تحويل داخلي' },
        { command: 'status', description: '📊 عرض حالة الحركات' },
        { command: 'help', description: '❓ عرض المساعدة' },
        { command: 'cancel', description: '❌ إلغاء العملية' }
    ],

    // ==================== حالات المحادثة ====================
    CONVERSATION_STATES: {
        IDLE: 'idle',
        WAITING_NATURE: 'waiting_nature',
        WAITING_CLASSIFICATION: 'waiting_classification',
        WAITING_PROJECT: 'waiting_project',
        WAITING_ITEM: 'waiting_item',
        WAITING_PARTY: 'waiting_party',
        WAITING_NEW_PARTY_TYPE: 'waiting_new_party_type',
        WAITING_AMOUNT: 'waiting_amount',
        WAITING_EXCHANGE_RATE: 'waiting_exchange_rate',
        WAITING_DETAILS: 'waiting_details',
        WAITING_PAYMENT_METHOD: 'waiting_payment_method',
        WAITING_PAYMENT_TERM: 'waiting_payment_term',
        WAITING_WEEKS: 'waiting_weeks',
        WAITING_CUSTOM_DATE: 'waiting_custom_date',
        WAITING_ATTACHMENT: 'waiting_attachment',
        WAITING_CONFIRMATION: 'waiting_confirmation',
        WAITING_EDIT_FIELD: 'waiting_edit_field',
        WAITING_EDIT_VALUE: 'waiting_edit_value',
        WAITING_SEQUENTIAL_EDIT: 'waiting_sequential_edit'
    },

    // ==================== أنماط Inline Keyboard ====================
    KEYBOARDS: {
        // لوحة اختيار طبيعة الحركة (محدّثة)
        TRANSACTION_TYPE: {
            inline_keyboard: [
                [
                    { text: '📤 استحقاق مصروف', callback_data: 'nature_استحقاق مصروف' },
                    { text: '💸 دفعة مصروف', callback_data: 'nature_دفعة مصروف' }
                ],
                [
                    { text: '📥 استحقاق إيراد', callback_data: 'nature_استحقاق إيراد' },
                    { text: '💰 تحصيل إيراد', callback_data: 'nature_تحصيل إيراد' }
                ],
                [
                    { text: '🏦 تمويل (دخول قرض)', callback_data: 'nature_تمويل (دخول قرض)' },
                    { text: '💳 سداد تمويل', callback_data: 'nature_سداد تمويل' }
                ],
                [
                    { text: '🔒 تأمين مدفوع للقناة', callback_data: 'nature_تأمين مدفوع للقناة' },
                    { text: '🔓 استرداد تأمين من القناة', callback_data: 'nature_استرداد تأمين من القناة' }
                ],
                [
                    { text: '🔄 تحويل داخلي', callback_data: 'nature_تحويل داخلي' }
                ],
                [
                    { text: '❌ إلغاء', callback_data: 'cancel' }
                ]
            ]
        },

        // لوحة اختيار تصنيف الحركة - العامة (للتوافق)
        CLASSIFICATION: {
            inline_keyboard: [
                [
                    { text: '📊 مصروفات مباشرة', callback_data: 'class_مصروفات مباشرة' },
                    { text: '🏢 مصروفات عمومية', callback_data: 'class_مصروفات عمومية' }
                ],
                [
                    { text: '💵 ايراد', callback_data: 'class_ايراد' }
                ],
                [
                    { text: '🏦 تمويل طويل الأجل', callback_data: 'class_تمويل طويل الأجل' },
                    { text: '💳 سداد تمويل طويل الأجل', callback_data: 'class_سداد تمويل طويل الأجل' }
                ],
                [
                    { text: '💸 سلفة قصيرة', callback_data: 'class_سلفة قصيرة' },
                    { text: '💰 سداد سلفة قصيرة', callback_data: 'class_سداد سلفة قصيرة' }
                ],
                [
                    { text: '🔒 تأمين', callback_data: 'class_تأمين' },
                    { text: '🔓 استرداد تأمين', callback_data: 'class_استرداد تأمين' }
                ],
                [
                    { text: '🏦 تحويل للبنك', callback_data: 'class_تحويل للبنك' },
                    { text: '💵 تحويل للخزنة', callback_data: 'class_تحويل للخزنة' }
                ],
                [
                    { text: '❌ إلغاء', callback_data: 'cancel' }
                ]
            ]
        },

        // لوحات التصنيف حسب نوع الحركة
        CLASSIFICATION_EXPENSE: {
            inline_keyboard: [
                [
                    { text: '📊 مصروفات مباشرة', callback_data: 'class_مصروفات مباشرة' },
                    { text: '🏢 مصروفات عمومية', callback_data: 'class_مصروفات عمومية' }
                ],
                [
                    { text: '❌ إلغاء', callback_data: 'cancel' }
                ]
            ]
        },

        CLASSIFICATION_REVENUE: {
            inline_keyboard: [
                [
                    { text: '💵 ايراد', callback_data: 'class_ايراد' }
                ],
                [
                    { text: '❌ إلغاء', callback_data: 'cancel' }
                ]
            ]
        },

        CLASSIFICATION_FINANCE: {
            inline_keyboard: [
                [
                    { text: '🏦 تمويل طويل الأجل', callback_data: 'class_تمويل طويل الأجل' },
                    { text: '💳 سداد تمويل طويل الأجل', callback_data: 'class_سداد تمويل طويل الأجل' }
                ],
                [
                    { text: '💸 سلفة قصيرة', callback_data: 'class_سلفة قصيرة' },
                    { text: '💰 سداد سلفة قصيرة', callback_data: 'class_سداد سلفة قصيرة' }
                ],
                [
                    { text: '❌ إلغاء', callback_data: 'cancel' }
                ]
            ]
        },

        CLASSIFICATION_INSURANCE: {
            inline_keyboard: [
                [
                    { text: '🔒 تأمين', callback_data: 'class_تأمين' },
                    { text: '🔓 استرداد تأمين', callback_data: 'class_استرداد تأمين' }
                ],
                [
                    { text: '❌ إلغاء', callback_data: 'cancel' }
                ]
            ]
        },

        CLASSIFICATION_TRANSFER: {
            inline_keyboard: [
                [
                    { text: '🏦 تحويل للبنك', callback_data: 'class_تحويل للبنك' },
                    { text: '💵 تحويل للخزنة', callback_data: 'class_تحويل للخزنة' }
                ],
                [
                    { text: '❌ إلغاء', callback_data: 'cancel' }
                ]
            ]
        },

        // لوحة اختيار الحقل للتعديل (جديد)
        EDIT_FIELD_SELECT: {
            inline_keyboard: [
                [
                    { text: '📤 طبيعة الحركة', callback_data: 'editfield_nature' },
                    { text: '📊 تصنيف الحركة', callback_data: 'editfield_classification' }
                ],
                [
                    { text: '🎬 المشروع', callback_data: 'editfield_project' },
                    { text: '📁 البند', callback_data: 'editfield_item' }
                ],
                [
                    { text: '👤 الطرف', callback_data: 'editfield_party' },
                    { text: '💰 المبلغ', callback_data: 'editfield_amount' }
                ],
                [
                    { text: '💱 العملة', callback_data: 'editfield_currency' },
                    { text: '📝 التفاصيل', callback_data: 'editfield_details' }
                ],
                [
                    { text: '✅ إرسال بدون تعديل', callback_data: 'editfield_submit' }
                ],
                [
                    { text: '❌ إلغاء', callback_data: 'cancel' }
                ]
            ]
        },

        // لوحة اختيار العملة
        CURRENCY: {
            inline_keyboard: [
                [
                    { text: '💵 دولار USD', callback_data: 'currency_USD' },
                    { text: '🇹🇷 ليرة TRY', callback_data: 'currency_TRY' },
                    { text: '🇪🇬 جنيه EGP', callback_data: 'currency_EGP' }
                ],
                [
                    { text: '❌ إلغاء', callback_data: 'cancel' }
                ]
            ]
        },

        // لوحة اختيار طريقة الدفع
        PAYMENT_METHOD: {
            inline_keyboard: [
                [
                    { text: '💵 نقدي', callback_data: 'payment_نقدي' },
                    { text: '🏦 تحويل بنكي', callback_data: 'payment_تحويل بنكي' }
                ],
                [
                    { text: '📝 شيك', callback_data: 'payment_شيك' },
                    { text: '💳 بطاقة', callback_data: 'payment_بطاقة' }
                ],
                [
                    { text: '📦 أخرى', callback_data: 'payment_أخرى' }
                ],
                [
                    { text: '❌ إلغاء', callback_data: 'cancel' }
                ]
            ]
        },

        // لوحة شروط الدفع
        PAYMENT_TERMS: {
            inline_keyboard: [
                [
                    { text: '⚡ فوري', callback_data: 'term_فوري' }
                ],
                [
                    { text: '📦 بعد التسليم', callback_data: 'term_بعد التسليم' }
                ],
                [
                    { text: '📅 تاريخ مخصص', callback_data: 'term_تاريخ مخصص' }
                ],
                [
                    { text: '❌ إلغاء', callback_data: 'cancel' }
                ]
            ]
        },

        // لوحة نوع الطرف الجديد
        NEW_PARTY_TYPE: {
            inline_keyboard: [
                [
                    { text: '🏭 مورد', callback_data: 'partytype_مورد' },
                    { text: '👤 عميل', callback_data: 'partytype_عميل' },
                    { text: '💼 ممول', callback_data: 'partytype_ممول' }
                ],
                [
                    { text: '❌ إلغاء', callback_data: 'cancel' }
                ]
            ]
        },

        // لوحة التأكيد
        CONFIRMATION: {
            inline_keyboard: [
                [
                    { text: '✅ تأكيد وإرسال', callback_data: 'confirm_yes' }
                ],
                [
                    { text: '✏️ تعديل', callback_data: 'confirm_edit' },
                    { text: '❌ إلغاء', callback_data: 'confirm_cancel' }
                ]
            ]
        },

        // لوحة المرفقات
        ATTACHMENT: {
            inline_keyboard: [
                [
                    { text: '📎 إرفاق صورة/ملف', callback_data: 'attach_yes' }
                ],
                [
                    { text: '⏭️ تخطي بدون مرفق', callback_data: 'attach_skip' }
                ],
                [
                    { text: '❌ إلغاء', callback_data: 'cancel' }
                ]
            ]
        },

        // لوحة التعديل بعد الرفض
        EDIT_AFTER_REJECT: {
            inline_keyboard: [
                [
                    { text: '🔄 تعديل وإعادة إرسال', callback_data: 'edit_resend' }
                ],
                [
                    { text: '🗑️ حذف نهائي', callback_data: 'edit_delete' }
                ]
            ]
        },

        // لوحة تعديل أو تخطي الحقل
        EDIT_OR_SKIP: {
            inline_keyboard: [
                [
                    { text: '✏️ تعديل', callback_data: 'seq_edit' },
                    { text: '➡️ التالي كما هو', callback_data: 'seq_skip' }
                ],
                [
                    { text: '✅ إرسال الآن', callback_data: 'seq_submit' }
                ],
                [
                    { text: '❌ إلغاء', callback_data: 'cancel' }
                ]
            ]
        },

        // لوحة التأكيد النهائي للتعديل
        EDIT_FINAL_CONFIRM: {
            inline_keyboard: [
                [
                    { text: '✅ إرسال', callback_data: 'seq_submit' },
                    { text: '🔄 مراجعة من البداية', callback_data: 'seq_restart' }
                ],
                [
                    { text: '❌ إلغاء', callback_data: 'cancel' }
                ]
            ]
        }
    },

    // ==================== الرسائل التفاعلية ====================
    INTERACTIVE_MESSAGES: {
        WELCOME: `
👋 *مرحباً بك في نظام SEEN المحاسبي*

يمكنك تسجيل الحركات المالية من هنا مباشرة.

*الأوامر المتاحة:*
📤 /مصروف - تسجيل مصروف جديد
📥 /ايراد - تسجيل إيراد جديد
🏦 /تمويل - تسجيل تمويل (قرض/سداد)
🔐 /تأمين - تسجيل تأمين (دفع/استرداد)
🔄 /تحويل - تسجيل تحويل داخلي

📊 /حالة - عرض حالة حركاتك المعلقة
❓ /مساعدة - عرض المساعدة
❌ /الغاء - إلغاء العملية الحالية
        `,

        HELP: `
📖 *دليل الاستخدام*

*الأوامر الأساسية:*
📤 /مصروف - لتسجيل مصروف (استحقاق أو دفعة)
📥 /ايراد - لتسجيل إيراد (استحقاق أو تحصيل)
🏦 /تمويل - لتسجيل قرض أو سداد قرض
🔐 /تأمين - لتسجيل تأمين مدفوع أو مسترد
🔄 /تحويل - لتسجيل تحويل داخلي

*خطوات تسجيل حركة:*
1️⃣ اختر طبيعة الحركة
2️⃣ اختر تصنيف الحركة
3️⃣ اختر المشروع
4️⃣ اختر البند
5️⃣ اختر أو أضف الطرف
6️⃣ أدخل المبلغ والعملة
7️⃣ أدخل التفاصيل
8️⃣ حدد شروط الدفع
9️⃣ أرفق صورة (اختياري)
🔟 راجع وأكد

*بعد التأكيد:*
• تذهب الحركة للمراجعة
• ستصلك إشعار عند القبول أو الرفض
        `,

        SELECT_PROJECT: '🎬 *اختر المشروع:*\n\nاكتب اسم المشروع أو جزء منه للبحث:',
        SELECT_ITEM: '📋 *اختر البند:*\n\nاكتب اسم البند أو جزء منه للبحث:',
        SELECT_PARTY: '👤 *اختر الطرف:*\n\nاكتب اسم المورد/العميل أو جزء منه للبحث:',
        ENTER_AMOUNT: '💰 *أدخل المبلغ:*\n\nاكتب المبلغ بالأرقام فقط (مثال: 500)',
        SELECT_CURRENCY: '💱 *اختر العملة:*',
        ENTER_EXCHANGE_RATE: '📊 *أدخل سعر الصرف:*\n\n(سعر الدولار مقابل العملة المختارة)',
        ENTER_DETAILS: '📝 *أدخل التفاصيل:*\n\nوصف مختصر للحركة',
        SELECT_PAYMENT_METHOD: '💳 *اختر طريقة الدفع:*',
        SELECT_PAYMENT_TERM: '📅 *اختر شرط الدفع:*',
        ENTER_WEEKS: '📆 *أدخل عدد الأسابيع:*\n\n(بعد التسليم)',
        ENTER_CUSTOM_DATE: '📅 *أدخل التاريخ:*\n\n(بصيغة: يوم/شهر/سنة مثل 15/01/2024)',
        ASK_ATTACHMENT: '📎 *هل تريد إرفاق صورة الفاتورة؟*',
        SEND_ATTACHMENT: '📤 *أرسل صورة الفاتورة أو الإيصال:*',

        NO_RESULTS: '❌ لم يتم العثور على نتائج\n\nجرب كتابة اسم مختلف أو جزء آخر من الاسم',
        PARTY_NOT_FOUND: '❌ لم يتم العثور على الطرف\n\n*هل تريد إضافته كطرف جديد؟*',

        TRANSACTION_SUMMARY: `
📋 *ملخص الحركة*
─────────────────
📌 *النوع:* {nature}
🎬 *المشروع:* {project}
📂 *البند:* {item}
👤 *الطرف:* {party}
💵 *المبلغ:* {amount} {currency}
📝 *التفاصيل:* {details}
💳 *طريقة الدفع:* {payment_method}
📅 *شرط الدفع:* {payment_term}
📎 *مرفق:* {attachment}
─────────────────
        `,

        SUCCESS: '✅ *تم تسجيل الحركة بنجاح!*\n\nرقم الحركة: #{id}\n\n⏳ في انتظار مراجعة المحاسب',
        CANCELLED: '❌ *تم إلغاء العملية*',
        ERROR: '❌ *حدث خطأ*\n\nيرجى المحاولة مرة أخرى أو التواصل مع المسؤول',

        // إشعارات المراجعة
        APPROVED_NOTIFICATION: `
✅ *تم اعتماد الحركة!*
─────────────────
📌 رقم الحركة: #{id}
📅 التاريخ: {date}
💵 المبلغ: {amount}
👤 الطرف: {party}
─────────────────
✨ تم نقلها لدفتر الحركات الرئيسي
        `,

        REJECTED_NOTIFICATION: `
❌ *تم رفض الحركة*
─────────────────
📌 رقم الحركة: #{id}
📅 التاريخ: {date}
💵 المبلغ: {amount}
👤 الطرف: {party}
─────────────────
📝 *سبب الرفض:*
{reason}
─────────────────
        `
    },

    // ==================== إعدادات البحث الذكي ====================
    FUZZY_SEARCH: {
        // الحد الأدنى لنسبة التطابق (0.0 - 1.0)
        MIN_MATCH_SCORE: 0.6,

        // الحد الأقصى لعدد النتائج
        MAX_RESULTS: 5,

        // أحرف عربية متشابهة للمعالجة
        ARABIC_NORMALIZATIONS: {
            'أ': 'ا', 'إ': 'ا', 'آ': 'ا', 'ٱ': 'ا',
            'ة': 'ه',
            'ى': 'ي',
            'ؤ': 'و',
            'ئ': 'ي'
        },

        // أحرف تُحذف
        REMOVE_CHARS: ['ً', 'ٌ', 'ٍ', 'َ', 'ُ', 'ِ', 'ّ', 'ْ']
    },

    // ==================== إعدادات المرفقات ====================
    ATTACHMENTS: {
        // أنواع الملفات المسموح بها
        ALLOWED_TYPES: ['image/jpeg', 'image/png', 'image/gif', 'application/pdf'],

        // الحد الأقصى لحجم الملف (10 ميجا)
        MAX_SIZE_MB: 10,

        // تنسيق اسم الملف
        FILE_NAME_FORMAT: 'حركة_{id}_{timestamp}.{ext}'
    }
};
