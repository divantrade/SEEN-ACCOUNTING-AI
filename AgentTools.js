// ==================== أدوات الوكيل الذكي (Agent Tools) ====================
/**
 * تعريف الأدوات التي يستخدمها الـ AI Agent مع Gemini Function Calling
 * كل أداة = دالة يمكن لـ AI استدعاؤها لإنجاز مهمة محددة
 *
 * ⚠️ هذا الملف جزء من التحديث المعماري الجديد (Smart Agent)
 * الكود الأصلي في AIAgent.js و AIBot.js لا يزال يعمل بشكل مستقل
 */

// ==================== تعريف الأدوات لـ Gemini Function Calling ====================

/**
 * تعريفات الأدوات بصيغة Gemini Function Calling
 * هذه هي الأدوات التي يمكن لـ AI استدعاؤها
 */
var AGENT_TOOL_DECLARATIONS = [
    {
        name: 'search_projects',
        description: 'البحث في قائمة المشاريع المتاحة بالاسم أو الكود. استخدم هذه الأداة عندما تحتاج معرفة المشاريع الموجودة أو التأكد من اسم مشروع.',
        parameters: {
            type: 'OBJECT',
            properties: {
                query: {
                    type: 'STRING',
                    description: 'نص البحث - اسم المشروع أو جزء منه أو كوده'
                }
            },
            required: ['query']
        }
    },
    {
        name: 'search_parties',
        description: 'البحث في قائمة الأطراف (موردين/عملاء/ممولين). إلزامي عند ذكر اسم طرف. يُرجع: اسم الطرف + نوعه + تخصصه + المشاريع التي عمل فيها سابقاً. استخدم النتائج لاستنتاج التصنيف والبند والمشروع تلقائياً.',
        parameters: {
            type: 'OBJECT',
            properties: {
                query: {
                    type: 'STRING',
                    description: 'اسم الطرف أو جزء منه'
                }
            },
            required: ['query']
        }
    },
    {
        name: 'search_items',
        description: 'البحث في قائمة البنود المحاسبية (تصوير، مونتاج، جرافيك...). استخدم هذه الأداة لمطابقة بند من النص.',
        parameters: {
            type: 'OBJECT',
            properties: {
                query: {
                    type: 'STRING',
                    description: 'اسم البند أو وصفه'
                }
            },
            required: ['query']
        }
    },
    {
        name: 'get_account_balance',
        description: 'الحصول على رصيد حساب معين (بنك أو خزنة). استخدم هذه الأداة عند الحاجة للتحقق من الرصيد المتاح قبل تسجيل دفعة.',
        parameters: {
            type: 'OBJECT',
            properties: {
                account_type: {
                    type: 'STRING',
                    description: 'نوع الحساب',
                    enum: ['bank_usd', 'bank_try', 'cash_usd', 'cash_try', 'card_try']
                }
            },
            required: ['account_type']
        }
    },
    {
        name: 'ask_user',
        description: 'سؤال المستخدم عن معلومة ناقصة أو غامضة. استخدم هذه الأداة فقط عندما لا تستطيع استنتاج المعلومة من النص. اجمع عدة أسئلة في سؤال واحد إن أمكن.',
        parameters: {
            type: 'OBJECT',
            properties: {
                question: {
                    type: 'STRING',
                    description: 'السؤال الموجه للمستخدم باللغة العربية'
                },
                options: {
                    type: 'ARRAY',
                    description: 'خيارات للمستخدم (اختياري - تظهر كأزرار)',
                    items: {
                        type: 'OBJECT',
                        properties: {
                            text: { type: 'STRING', description: 'نص الزر' },
                            value: { type: 'STRING', description: 'القيمة عند الضغط' }
                        },
                        required: ['text', 'value']
                    }
                },
                field_name: {
                    type: 'STRING',
                    description: 'اسم الحقل المطلوب (project, party, payment_method, currency, exchange_rate, item, payment_term)'
                }
            },
            required: ['question']
        }
    },
    {
        name: 'save_transaction',
        description: 'حفظ الحركة المالية في دفتر الحركات بعد التأكد من اكتمال جميع البيانات المطلوبة. لا تستدعِ هذه الأداة إلا بعد تأكيد المستخدم.',
        parameters: {
            type: 'OBJECT',
            properties: {
                nature: {
                    type: 'STRING',
                    description: 'طبيعة الحركة'
                },
                classification: {
                    type: 'STRING',
                    description: 'تصنيف الحركة'
                },
                project: {
                    type: 'STRING',
                    description: 'اسم المشروع (اختياري لبعض أنواع الحركات)'
                },
                project_code: {
                    type: 'STRING',
                    description: 'كود المشروع'
                },
                item: {
                    type: 'STRING',
                    description: 'البند المحاسبي'
                },
                party: {
                    type: 'STRING',
                    description: 'اسم الطرف (مورد/عميل/ممول)'
                },
                party_type: {
                    type: 'STRING',
                    description: 'نوع الطرف',
                    enum: ['مورد', 'عميل', 'ممول']
                },
                is_new_party: {
                    type: 'BOOLEAN',
                    description: 'هل هذا طرف جديد غير مسجل؟'
                },
                amount: {
                    type: 'NUMBER',
                    description: 'المبلغ'
                },
                currency: {
                    type: 'STRING',
                    description: 'العملة',
                    enum: ['USD', 'TRY', 'EGP']
                },
                exchange_rate: {
                    type: 'NUMBER',
                    description: 'سعر الصرف (مطلوب لغير الدولار)'
                },
                payment_method: {
                    type: 'STRING',
                    description: 'طريقة الدفع',
                    enum: ['نقدي', 'تحويل بنكي', 'شيك', 'بطاقة']
                },
                payment_term: {
                    type: 'STRING',
                    description: 'شرط الدفع',
                    enum: ['فوري', 'بعد التسليم', 'تاريخ مخصص']
                },
                payment_term_weeks: {
                    type: 'NUMBER',
                    description: 'عدد الأسابيع بعد التسليم'
                },
                payment_term_date: {
                    type: 'STRING',
                    description: 'تاريخ الدفع المخصص (YYYY-MM-DD)'
                },
                due_date: {
                    type: 'STRING',
                    description: 'تاريخ الحركة (YYYY-MM-DD)'
                },
                details: {
                    type: 'STRING',
                    description: 'تفاصيل إضافية'
                },
                unit_count: {
                    type: 'NUMBER',
                    description: 'عدد الوحدات (مقابلات، دقائق، رسومات...)'
                },
                loan_due_date: {
                    type: 'STRING',
                    description: 'تاريخ سداد القرض/السلفة (YYYY-MM-DD)'
                }
            },
            required: ['nature', 'classification', 'amount', 'currency', 'due_date']
        }
    },
    {
        name: 'show_confirmation',
        description: 'عرض ملخص الحركة للمستخدم وطلب تأكيده قبل الحفظ. استخدم هذه الأداة دائماً قبل save_transaction.',
        parameters: {
            type: 'OBJECT',
            properties: {
                transaction: {
                    type: 'OBJECT',
                    description: 'بيانات الحركة الكاملة للعرض',
                    properties: {
                        nature: { type: 'STRING' },
                        classification: { type: 'STRING' },
                        project: { type: 'STRING' },
                        project_code: { type: 'STRING' },
                        item: { type: 'STRING' },
                        party: { type: 'STRING' },
                        amount: { type: 'NUMBER' },
                        currency: { type: 'STRING' },
                        exchange_rate: { type: 'NUMBER' },
                        payment_method: { type: 'STRING' },
                        payment_term: { type: 'STRING' },
                        due_date: { type: 'STRING' },
                        details: { type: 'STRING' },
                        unit_count: { type: 'NUMBER' }
                    }
                }
            },
            required: ['transaction']
        }
    }
];

// ==================== تنفيذ الأدوات (Tool Execution) ====================

/**
 * تنفيذ أداة استُدعيت من قبل AI
 * @param {string} toolName - اسم الأداة
 * @param {Object} args - المعاملات
 * @returns {Object} - نتيجة التنفيذ
 */
function executeAgentTool_(toolName, args) {
    Logger.log('🔧 تنفيذ أداة: ' + toolName + ' | args: ' + JSON.stringify(args));

    switch (toolName) {
        case 'search_projects':
            return executeSearchProjects_(args);
        case 'search_parties':
            return executeSearchParties_(args);
        case 'search_items':
            return executeSearchItems_(args);
        case 'get_account_balance':
            return executeGetAccountBalance_(args);
        case 'ask_user':
            return executeAskUser_(args);
        case 'save_transaction':
            return executeSaveTransaction_(args);
        case 'show_confirmation':
            return executeShowConfirmation_(args);
        default:
            return { error: 'أداة غير معروفة: ' + toolName };
    }
}

// ==================== تنفيذ كل أداة ====================

/**
 * البحث في المشاريع
 */
function executeSearchProjects_(args) {
    try {
        var results = searchProjects(args.query);
        if (results.length === 0) {
            // إرجاع كل المشاريع كبديل
            var allProjects = getAllProjects();
            return {
                found: false,
                message: 'لم يتم العثور على مشروع مطابق لـ "' + args.query + '"',
                available_projects: allProjects.slice(0, 10).map(function(p) {
                    return { code: p.code, name: p.name };
                })
            };
        }
        return {
            found: true,
            results: results.map(function(p) {
                return { code: p.code, name: p.name };
            })
        };
    } catch (e) {
        return { error: e.message };
    }
}

/**
 * البحث في الأطراف
 */
function executeSearchParties_(args) {
    try {
        var results = searchParties(args.query);
        if (results.length === 0) {
            // محاولة ثانية بالمطابقة الذكية (أوسع)
            var smartResults = searchParties(args.query, { useSmartMatch: true });
            if (smartResults.length > 0) {
                return {
                    found: false,
                    partial_match: true,
                    message: 'لم يتم العثور على تطابق كامل لـ "' + args.query + '" لكن وُجدت أسماء مشابهة',
                    similar_parties: smartResults.map(function(p) {
                        return { name: p.name, type: p.type, specialization: p.specialization || '' };
                    }),
                    hint: 'اعرض الأسماء المشابهة على المستخدم ليختار منها، أو أضفه كطرف جديد'
                };
            }
            return {
                found: false,
                message: 'لم يتم العثور على طرف مطابق لـ "' + args.query + '"',
                suggestion: 'يمكن إضافته كطرف جديد'
            };
        }

        var topMatch = results[0];
        var response = {
            found: true,
            results: results.map(function(p) {
                return { name: p.name, type: p.type, specialization: p.specialization || '' };
            })
        };

        // البحث عن مشاريع الطرف من سجل الحركات
        try {
            var partyProjects = findProjectsByParty_(topMatch.name);
            if (partyProjects.length > 0) {
                response.party_projects = partyProjects;
                response.hint = 'هذا الطرف عمل سابقاً في المشاريع المذكورة في party_projects';
            }
        } catch (e) {
            Logger.log('⚠️ خطأ في البحث عن مشاريع الطرف: ' + e.message);
        }

        return response;
    } catch (e) {
        return { error: e.message };
    }
}

/**
 * البحث عن المشاريع التي عمل فيها طرف معين من سجل الحركات
 */
function findProjectsByParty_(partyName) {
    try {
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var sheets = [CONFIG.SHEETS.TRANSACTIONS, CONFIG.SHEETS.BOT_TRANSACTIONS];
        var projectSet = {};

        for (var s = 0; s < sheets.length; s++) {
            var sheet = ss.getSheetByName(sheets[s]);
            if (!sheet || sheet.getLastRow() < 2) continue;

            var data = sheet.getDataRange().getValues();
            var headers = data[0];

            // البحث عن أعمدة الطرف والمشروع
            var partyCol = -1, projectCol = -1;
            for (var h = 0; h < headers.length; h++) {
                var header = String(headers[h]).trim();
                if (header === 'الطرف' || header === 'اسم الطرف') partyCol = h;
                if (header === 'المشروع' || header === 'كود المشروع') projectCol = h;
            }
            if (partyCol === -1 || projectCol === -1) continue;

            for (var i = data.length - 1; i >= 1; i--) {
                var rowParty = String(data[i][partyCol] || '').trim();
                var rowProject = String(data[i][projectCol] || '').trim();
                if (rowParty === partyName && rowProject && !projectSet[rowProject]) {
                    projectSet[rowProject] = true;
                    if (Object.keys(projectSet).length >= 5) break;
                }
            }
            if (Object.keys(projectSet).length >= 5) break;
        }

        return Object.keys(projectSet);
    } catch (e) {
        Logger.log('⚠️ findProjectsByParty_ error: ' + e.message);
        return [];
    }
}

/**
 * البحث في البنود
 */
function executeSearchItems_(args) {
    try {
        var results = searchItems(args.query);
        if (results.length === 0) {
            var allItems = getAllItems();
            return {
                found: false,
                message: 'لم يتم العثور على بند مطابق',
                available_items: allItems.slice(0, 15).map(function(i) { return i.name; })
            };
        }
        return {
            found: true,
            results: results.map(function(i) {
                return { name: i.name, nature: i.nature, classification: i.classification };
            })
        };
    } catch (e) {
        return { error: e.message };
    }
}

/**
 * الحصول على رصيد حساب
 */
function executeGetAccountBalance_(args) {
    try {
        var sheetMap = {
            'bank_usd': CONFIG.SHEETS.BANK_USD,
            'bank_try': CONFIG.SHEETS.BANK_TRY,
            'cash_usd': CONFIG.SHEETS.CASH_USD,
            'cash_try': CONFIG.SHEETS.CASH_TRY,
            'card_try': CONFIG.SHEETS.CARD_TRY
        };

        var sheetName = sheetMap[args.account_type];
        if (!sheetName) {
            return { error: 'نوع حساب غير معروف: ' + args.account_type };
        }

        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var balance = getLastBalanceFromSheet_(ss, sheetName);

        return {
            account: args.account_type,
            balance: balance || 0,
            currency: args.account_type.indexOf('usd') !== -1 ? 'USD' : 'TRY'
        };
    } catch (e) {
        return { error: e.message };
    }
}

/**
 * سؤال المستخدم - يُعيد تعليمات للـ Agent بانتظار الرد
 */
function executeAskUser_(args) {
    // هذه الأداة خاصة - لا تنفذ مباشرة بل تُعيد تعليمات
    // الـ SmartAgent يتعامل معها بإرسال رسالة Telegram وانتظار الرد
    return {
        action: 'WAIT_FOR_USER',
        question: args.question,
        options: args.options || null,
        field_name: args.field_name || null
    };
}

/**
 * عرض تأكيد الحركة
 */
function executeShowConfirmation_(args) {
    // مثل ask_user - تُعيد تعليمات للـ SmartAgent
    return {
        action: 'SHOW_CONFIRMATION',
        transaction: args.transaction
    };
}

/**
 * حفظ الحركة المالية
 */
function executeSaveTransaction_(args) {
    try {
        // التحقق من الحقول المطلوبة
        if (!args.nature || !args.amount || !args.currency) {
            return { error: 'حقول ناقصة: nature, amount, currency مطلوبة' };
        }

        // تجهيز بيانات الحركة بنفس الصيغة المتوقعة من saveAITransaction
        var transaction = {
            nature: args.nature,
            classification: args.classification || '',
            project: args.project || '',
            project_code: args.project_code || '',
            item: args.item || '',
            party: args.party || '',
            partyType: args.party_type || 'مورد',
            isNewParty: args.is_new_party || false,
            amount: args.amount,
            currency: args.currency,
            exchangeRate: args.exchange_rate || (args.currency === 'USD' ? 1 : null),
            payment_method: args.payment_method || '',
            payment_term: args.payment_term || 'فوري',
            payment_term_weeks: args.payment_term_weeks || '',
            payment_term_date: args.payment_term_date || '',
            due_date: args.due_date || formatDateISO_(new Date()),
            details: args.details || '',
            unit_count: args.unit_count || '',
            loan_due_date: args.loan_due_date || ''
        };

        // استخدام دالة الحفظ الموجودة
        // CURRENT_AGENT_CHAT_ID_ يُعيّن في processSmartTransaction_ قبل كل تحليل
        var chatId = CURRENT_AGENT_CHAT_ID_;
        var user = { first_name: 'Smart', last_name: 'Agent' };
        var result = saveAITransaction(transaction, user, chatId);
        return result;
    } catch (e) {
        Logger.log('❌ خطأ في حفظ الحركة: ' + e.message);
        return { success: false, error: e.message };
    }
}

// ==================== دوال مساعدة ====================

/**
 * تنسيق التاريخ كـ ISO
 */
function formatDateISO_(date) {
    var d = date || new Date();
    var year = d.getFullYear();
    var month = String(d.getMonth() + 1).padStart(2, '0');
    var day = String(d.getDate()).padStart(2, '0');
    return year + '-' + month + '-' + day;
}
