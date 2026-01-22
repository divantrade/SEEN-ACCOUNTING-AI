// ==================== دالة القائمة الرئيسية ====================
/**
 * دالة تُنفذ تلقائياً عند فتح ملف Google Sheets
 * تُنشئ القائمة المخصصة لنظام المحاسبة
 *
 * الترتيب الجديد:
 * 1. العمليات اليومية (الأكثر استخداماً)
 * 2. الاستحقاقات والمتابعة
 * 3. كشوف الحسابات
 * 4. التقارير (مشاريع، تشغيلية، ملخصة، مالية)
 * 5. الميزانية
 * 6. البنك والنقدية
 * 7. التحديث والصيانة
 * 8. الإعدادات
 * 9. دليل الاستخدام
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('💼 نظام المحاسبة')

    // ═══════════════════════════════════════════════════════════
    // 1. العمليات اليومية (الأكثر استخداماً - في الأعلى)
    // ═══════════════════════════════════════════════════════════
    .addItem('📝 إضافة حركة جديدة (نموذج)', 'showTransactionForm')
    .addItem('📦 أوردر مشترك (تقسيم بين مشاريع)', 'showSharedOrderForm')
    .addItem('🧾 إنشاء فاتورة قناة', 'generateChannelInvoice')
    .addItem('🔄 إعادة طباعة فاتورة', 'regenerateChannelInvoice')
    .addSeparator()

    // ═══════════════════════════════════════════════════════════
    // 2. التقارير التشغيلية
    // ═══════════════════════════════════════════════════════════
    .addSubMenu(
      ui.createMenu('📋 التقارير التشغيلية')
        .addItem('⏰ عرض الاستحقاقات (نافذة)', 'showUpcomingPayments')
        .addItem('📊 تقرير الاستحقاقات الشامل', 'generateDueReport')
        .addItem('📋 دفتر الأستاذ المساعد', 'generateDetailedPayablesReport')
    )

    // ═══════════════════════════════════════════════════════════
    // 3. كشوف الحسابات
    // ═══════════════════════════════════════════════════════════
    .addSubMenu(
      ui.createMenu('👥 كشوف الحسابات')
        .addItem('📄 كشف حساب مورد', 'generateVendorStatementSheet')
        .addItem('📄 كشف حساب عميل', 'generateClientStatementSheet')
        .addItem('📄 كشف حساب ممول', 'generateFunderStatementSheet')
    )

    // ═══════════════════════════════════════════════════════════
    // 4. التقارير
    // ═══════════════════════════════════════════════════════════

    // تقارير المشاريع والربحية
    .addSubMenu(
      ui.createMenu('📈 تقارير المشاريع والربحية')
        .addItem('📋 تقرير المشروع التفصيلي', 'rebuildProjectDetailReport')
        .addItem('📊 تقرير ربحية مشروع (نافذة)', 'showProjectProfitability')
        .addItem('📊 تقرير ربحية كل المشاريع', 'generateAllProjectsProfitabilityReport')
        .addSeparator()
        .addItem('📋 تقرير ميزانية مشروع', 'generateProjectBudgetReport')
        .addItem('📊 تقرير ميزانية مجمع (عدة مشاريع)', 'showCombinedBudgetReportForm')
        .addSeparator()
        .addSubMenu(
          ui.createMenu('📊 تقرير تكلفة الوحدة')
            .addItem('● مشروع واحد', 'showUnitCostReportSingle')
            .addItem('● عدة مشاريع (مجمع)', 'showUnitCostReportMultiple')
            .addItem('● كل المشاريع', 'generateUnitCostReportAll')
        )
        .addSeparator()
        .addItem('💰 تقرير عمولات مدير مشروعات', 'showCommissionReportDialog')
        .addItem('➕ إدراج استحقاق عمولة (من التقرير)', 'insertCommissionFromReport')
    )


    // التقارير الملخصة
    .addSubMenu(
      ui.createMenu('📈 التقارير الملخصة')
        .addItem('🏢 تقرير الموردين الملخص', 'rebuildVendorSummaryReport')
        .addItem('💼 تقرير الممولين الملخص', 'rebuildFunderSummaryReport')
        .addItem('💸 تقرير المصروفات الملخص', 'rebuildExpenseSummaryReport')
        .addItem('💰 تقرير الإيرادات الملخص', 'rebuildRevenueSummaryReport')
        .addItem('💵 تقرير التدفقات النقدية', 'rebuildCashFlowReport')
        .addSeparator()
        .addItem('🔄 تحديث كل التقارير الملخصة', 'rebuildAllSummaryReports')
    )

    // القوائم المالية
    .addSubMenu(
      ui.createMenu('📊 القوائم المالية')
        .addItem('📈 قائمة الدخل', 'rebuildIncomeStatement')
        .addItem('📋 المركز المالي', 'rebuildBalanceSheet')
        .addSeparator()
        .addItem('🌳 شجرة الحسابات', 'rebuildChartOfAccounts')
        .addItem('📒 دفتر الأستاذ العام', 'showGeneralLedger')
        .addItem('⚖️ ميزان المراجعة', 'rebuildTrialBalance')
        .addItem('📝 قيود اليومية', 'rebuildJournalEntries')
        .addSeparator()
        .addItem('🔒 الإقفال السنوي', 'performYearEndClosing')
    )
    .addSeparator()

    // ═══════════════════════════════════════════════════════════
    // 5. الميزانية
    // ═══════════════════════════════════════════════════════════
    .addItem('📝 إضافة ميزانية', 'addBudgetForm')
    .addItem('📊 مقارنة الميزانية', 'compareBudget')
    .addSeparator()

    // ═══════════════════════════════════════════════════════════
    // 6. البنك والنقدية
    // ═══════════════════════════════════════════════════════════
    .addSubMenu(
      ui.createMenu('🏦 البنك وخزنة العهدة')
        .addItem('🔄 تحديث شيتات البنك والخزنة والبطاقة', 'rebuildBankAndCashFromTransactions')
    )

    .addSubMenu(
      ui.createMenu('🔍 مطابقة الحساب البنكي / الكارت')
        .addItem('📝 إنشاء شيت مطابقة دولار', 'createBankReconciliationUsdSheet')
        .addItem('📝 إنشاء شيت مطابقة ليرة', 'createBankReconciliationTrySheet')
        .addItem('📝 إنشاء شيت مطابقة الكارت', 'createCardReconciliationSheet')
        .addSeparator()
        .addItem('✅ مطابقة حساب البنك - دولار', 'reconcileBankUsd')
        .addItem('✅ مطابقة حساب البنك - ليرة', 'reconcileBankTry')
        .addItem('✅ مطابقة الكارت', 'reconcileCard')
    )
    .addSeparator()

    // ═══════════════════════════════════════════════════════════
    // 7. التحديث والصيانة
    // ═══════════════════════════════════════════════════════════
    .addSubMenu(
      ui.createMenu('🔄 التحديث والصيانة')
        .addItem('📊 تحديث لوحة التحكم', 'refreshDashboard')
        .addItem('🔄 تحديث كل التقارير', 'rebuildAllSummaryReports')
        .addItem('🔔 تحديث التنبيهات', 'updateAlerts')
        .addItem('🔽 تحديث القوائم المنسدلة', 'refreshDropdowns')
        .addSeparator()
        .addItem('🔃 ترتيب الحركات حسب التاريخ', 'sortTransactionsByDate')
    )

    // ═══════════════════════════════════════════════════════════
    // 8. الإعدادات
    // ═══════════════════════════════════════════════════════════
    .addItem('🔍 تفعيل/إلغاء الفلتر', 'toggleFilter')

    .addSubMenu(
      ui.createMenu('👁️ إخفاء/إظهار الشيتات')
        .addItem('📊 إخفاء/إظهار التقارير', 'toggleReportsVisibility')
        .addItem('📋 إخفاء/إظهار التقارير التشغيلية', 'toggleOperationalReportsVisibility')
        .addItem('🏦 إخفاء/إظهار حسابات البنوك', 'toggleBankAccountsVisibility')
        .addItem('📁 إخفاء/إظهار قواعد البيانات', 'toggleDatabasesVisibility')
        .addItem('📒 إخفاء/إظهار الدفاتر والقوائم المحاسبية', 'toggleAccountingVisibility')
        .addItem('📋 إخفاء/إظهار سجل النشاط', 'toggleActivityLogVisibility')
    )

    // ═══════════════════════════════════════════════════════════
    // 8.5. بوت تليجرام
    // ═══════════════════════════════════════════════════════════
    .addSubMenu(
      ui.createMenu('🤖 بوت تليجرام')
        .addItem('📋 مراجعة حركات البوت', 'showBotReviewSidebar')
        .addItem('✅ اعتماد جميع الحركات المعلقة', 'approveAllPendingTransactions')
        .addItem('📊 تقرير الحركات المعلقة', 'showPendingTransactionsReport')
        .addSeparator()
        .addItem('📈 إحصائيات البوت', 'showBotStatistics')
        .addItem('📎 تقرير المرفقات', 'showAttachmentsReport')
        .addItem('🗑️ تنظيف مرفقات المرفوضة', 'cleanupRejectedAttachments')
    )

    .addSubMenu(
      ui.createMenu('⚙️ إعدادات متقدمة')
        .addItem('📅 تطبيع التواريخ', 'normalizeDateColumns')
        .addItem('📋 إصلاح القوائم المنسدلة', 'fixAllDropdowns')
        .addItem('🔗 مراجعة وإصلاح نوع الحركة', 'reviewAndFixMovementTypes')
        .addItem('⚖️ فحص الاستحقاقات والدفعات (سريع)', 'checkAccrualPaymentBalance')
        .addItem('⚖️ تقرير الاستحقاقات والدفعات (شيت)', 'generateAccrualPaymentReport')
        .addItem('🎨 إعادة تطبيق التلوين الشرطي', 'refreshTransactionsFormatting')
        .addItem('💵 تحديث شامل (M, O, U, V)', 'refreshValueAndBalanceFormulas')
        .addSeparator()
        .addItem('🔔 تفعيل التسجيل التلقائي', 'installActivityTriggers')
        .addItem('🔕 إيقاف التسجيل التلقائي', 'uninstallActivityTriggers')
        .addSeparator()
        .addItem('💾 إنشاء نسخة احتياطية للشيت', 'backupSpreadsheet')
    )

    // ═══════════════════════════════════════════════════════════
    // 9. تعريف المستخدم
    // ═══════════════════════════════════════════════════════════
    .addSeparator()
    .addSubMenu(
      ui.createMenu('👤 تعريف المستخدم')
        .addItem('🔑 تعريف نفسي الآن', 'showUserIdentificationDialog')
        .addSeparator()
        .addItem('🔔 تفعيل النافذة التلقائية', 'installUserIdentificationTrigger')
        .addItem('🔕 إلغاء النافذة التلقائية', 'uninstallUserIdentificationTrigger')
    )

    // ═══════════════════════════════════════════════════════════
    // 10. دليل الاستخدام
    // ═══════════════════════════════════════════════════════════
    .addSeparator()
    .addItem('📖 دليل الاستخدام', 'showGuide')
    .addToUi();

  // ═══════════════════════════════════════════════════════════
  // تنبيه المستخدم بتعريف نفسه إذا لم يكن معرّفاً
  // ملاحظة: Simple Trigger لا يمكنه عرض Modal Dialog مباشرة
  // لذلك نستخدم toast notification
  // ═══════════════════════════════════════════════════════════
  try {
    if (!isUserIdentified()) {
      // عرض تنبيه toast للمستخدم
      SpreadsheetApp.getActiveSpreadsheet().toast(
        '👤 يرجى تعريف نفسك من القائمة:\nنظام المحاسبة ← تعريف المستخدم',
        '⚠️ تنبيه',
        15  // يبقى 15 ثانية
      );
    }
  } catch (e) {
    console.log('تعذر عرض تنبيه تعريف المستخدم:', e.message);
  }
}


// ==================== تسجيل التعديلات التلقائي ====================

/**
 * تثبيت الـ Triggers للتسجيل التلقائي
 * يجب تشغيل هذه الدالة مرة واحدة فقط
 */
function installActivityTriggers() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // حذف الـ triggers القديمة أولاً
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    const funcName = trigger.getHandlerFunction();
    if (funcName === 'onEditHandler' || funcName === 'onChangeHandler') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // إنشاء trigger للتعديلات
  ScriptApp.newTrigger('onEditHandler')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onEdit()
    .create();

  // إنشاء trigger للتغييرات الهيكلية
  ScriptApp.newTrigger('onChangeHandler')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onChange()
    .create();

  // ═══════════════════════════════════════════════════════════════
  // تهيئة كاش البيانات للشيتات المتتبعة (لاكتشاف الحذف)
  // نحفظ أرقام الحركات/المعرفات للمقارنة لاحقاً
  // ═══════════════════════════════════════════════════════════════
  const trackedSheets = [
    CONFIG.SHEETS.TRANSACTIONS,
    CONFIG.SHEETS.PROJECTS,
    CONFIG.SHEETS.PARTIES,
    CONFIG.SHEETS.ITEMS,
    CONFIG.SHEETS.BUDGETS
  ];

  const props = PropertiesService.getScriptProperties();
  trackedSheets.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (sheet) {
      const lastRow = sheet.getLastRow();
      let rowData = [];

      if (lastRow > 1) {
        const data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
        rowData = data.map((row, index) => ({
          id: String(row[0] || ''),
          row: index + 2
        })).filter(item => item.id !== '');
      }

      const cacheKey = 'rowData_' + sheetName.replace(/\s/g, '_');
      props.setProperty(cacheKey, JSON.stringify(rowData));
    }
  });

  ui.alert(
    '✅ تم تثبيت التسجيل التلقائي',
    'سيتم الآن تسجيل جميع التعديلات تلقائياً:\n\n' +
    '• تعديل القيم\n' +
    '• إضافة صفوف\n' +
    '• حذف صفوف\n' +
    '• الإدخال اليدوي\n\n' +
    'يمكنك مراجعة السجل من:\n' +
    'نظام المحاسبة ← إخفاء/إظهار الشيتات ← سجل النشاط',
    ui.ButtonSet.OK
  );
}


/**
 * إزالة الـ Triggers للتسجيل التلقائي
 */
function uninstallActivityTriggers() {
  const ui = SpreadsheetApp.getUi();

  const response = ui.alert(
    '⚠️ إزالة التسجيل التلقائي',
    'هل تريد إيقاف التسجيل التلقائي للتعديلات؟',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  const triggers = ScriptApp.getProjectTriggers();
  let removed = 0;

  triggers.forEach(trigger => {
    const funcName = trigger.getHandlerFunction();
    if (funcName === 'onEditHandler' || funcName === 'onChangeHandler') {
      ScriptApp.deleteTrigger(trigger);
      removed++;
    }
  });

  ui.alert('✅ تم', `تم إزالة ${removed} trigger(s).`, ui.ButtonSet.OK);
}


/**
 * دالة تُنفذ تلقائياً عند أي تعديل في الشيت (Installable Trigger)
 * تسجل التعديلات اليدوية في سجل النشاط
 */
function onEditHandler(e) {
  try {
    // التحقق من وجود الحدث
    if (!e || !e.range) return;

    const sheet = e.range.getSheet();
    const sheetName = sheet.getName();

    // تسجيل التعديلات في الشيتات المهمة فقط
    const trackedSheets = [
      CONFIG.SHEETS.TRANSACTIONS,
      CONFIG.SHEETS.PROJECTS,
      CONFIG.SHEETS.PARTIES,
      CONFIG.SHEETS.ITEMS,
      CONFIG.SHEETS.BUDGETS
    ];

    if (!trackedSheets.includes(sheetName)) return;

    // جلب إيميل المستخدم من الحدث (e.user) - يعمل مع المستخدمين الآخرين
    let userEmail = '';
    try {
      if (e.user && e.user.getEmail) {
        userEmail = e.user.getEmail();
      } else if (e.user && e.user.email) {
        userEmail = e.user.email;
      }
    } catch (ue) {
      // تجاهل - سيتم استخدام الطرق البديلة في logActivity
    }

    // ═══════════════════════════════════════════════════════════════
    // اكتشاف حذف الصفوف بمقارنة أرقام الحركات/البيانات
    // ═══════════════════════════════════════════════════════════════
    try {
      detectDeletedRows(sheet, sheetName, userEmail);
    } catch (propErr) {
      // تجاهل أخطاء تتبع الصفوف
    }

    // تجاهل تعديلات الهيدر (الصف الأول)
    const row = e.range.getRow();
    if (row === 1) return;

    // جمع معلومات التعديل
    const col = e.range.getColumn();
    const oldValue = e.oldValue !== undefined ? e.oldValue : '';
    const newValue = e.value !== undefined ? e.value : '';

    // تجاهل إذا لم يتغير شيء
    if (oldValue === newValue) return;

    // جلب اسم العمود
    const columnHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const columnName = columnHeaders[col - 1] || `عمود ${col}`;

    // جلب رقم الحركة (العمود A) إذا كان شيت الحركات
    let transNum = '';
    if (sheetName === CONFIG.SHEETS.TRANSACTIONS) {
      transNum = sheet.getRange(row, 1).getValue() || '';
    }

    // تحديد نوع العملية
    let actionType = 'تعديل';
    if (oldValue === '' && newValue !== '') {
      actionType = 'إضافة قيمة';
    } else if (oldValue !== '' && newValue === '') {
      actionType = 'حذف قيمة';
    }

    // تسجيل النشاط مع تمرير إيميل المستخدم
    logActivity(
      actionType,
      sheetName,
      row,
      transNum,
      `${columnName}: "${oldValue}" → "${newValue}"`,
      {
        column: col,
        columnName: columnName,
        oldValue: oldValue,
        newValue: newValue
      },
      userEmail
    );

  } catch (err) {
    // لا نوقف التعديل في حالة خطأ التسجيل
    console.error('خطأ في تسجيل التعديل:', err.message);
  }
}


/**
 * دالة تُنفذ عند تغيير هيكل الشيت (إضافة/حذف صفوف أو أعمدة)
 */
function onChangeHandler(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const activeSheet = ss.getActiveSheet();
    const sheetName = activeSheet.getName();

    // الشيتات المتتبعة
    const trackedSheets = [
      CONFIG.SHEETS.TRANSACTIONS,
      CONFIG.SHEETS.PROJECTS,
      CONFIG.SHEETS.PARTIES,
      CONFIG.SHEETS.ITEMS,
      CONFIG.SHEETS.BUDGETS
    ];

    if (!trackedSheets.includes(sheetName)) {
      return;
    }

    const changeType = e ? e.changeType : 'UNKNOWN';

    // جلب إيميل المستخدم من الحدث (e.user) - يعمل مع المستخدمين الآخرين
    let userEmail = '';
    try {
      if (e && e.user && e.user.getEmail) {
        userEmail = e.user.getEmail();
      } else if (e && e.user && e.user.email) {
        userEmail = e.user.email;
      }
    } catch (ue) {
      // تجاهل - سيتم استخدام الطرق البديلة في logActivity
    }

    // ═══════════════════════════════════════════════════════════════
    // اكتشاف حذف الصفوف بمقارنة البيانات المخزنة
    // ═══════════════════════════════════════════════════════════════
    try {
      detectDeletedRows(activeSheet, sheetName, userEmail);
    } catch (propErr) {
      console.log('خطأ في اكتشاف الحذف:', propErr.message);
    }

    // ═══════════════════════════════════════════════════════════════
    // تسجيل إضافة الصفوف والأعمدة
    // ═══════════════════════════════════════════════════════════════
    if (changeType === 'INSERT_ROW') {
      logActivity(
        'إضافة صف',
        sheetName,
        null,
        null,
        `تم إضافة صف في ${sheetName}`,
        { changeType: changeType },
        userEmail
      );
    } else if (changeType === 'INSERT_COLUMN' || changeType === 'REMOVE_COLUMN') {
      const actionType = changeType === 'INSERT_COLUMN' ? 'إضافة عمود' : 'حذف عمود';
      logActivity(
        actionType,
        sheetName,
        null,
        null,
        `تم ${actionType} في ${sheetName}`,
        { changeType: changeType },
        userEmail
      );
    }

  } catch (err) {
    console.error('خطأ في تسجيل التغيير:', err.message, err.stack);
  }
}


/**
 * اكتشاف الصفوف المحذوفة بمقارنة البيانات المخزنة
 * يحفظ أرقام الحركات/المعرفات ويقارنها لاكتشاف المحذوف
 * @param {string} userEmail - إيميل المستخدم من e.user (اختياري)
 */
function detectDeletedRows(sheet, sheetName, userEmail) {
  const props = PropertiesService.getScriptProperties();
  const cacheKey = 'rowData_' + sheetName.replace(/\s/g, '_');

  // جلب البيانات الحالية (العمود الأول - المعرف/رقم الحركة)
  const lastRow = sheet.getLastRow();
  let currentIds = [];

  if (lastRow > 1) {
    const data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    currentIds = data.map((row, index) => ({
      id: String(row[0] || ''),
      row: index + 2
    })).filter(item => item.id !== '');
  }

  // جلب البيانات المخزنة سابقاً
  const cachedDataStr = props.getProperty(cacheKey) || '[]';
  let cachedData = [];
  try {
    cachedData = JSON.parse(cachedDataStr);
  } catch (e) {
    cachedData = [];
  }

  // إنشاء مجموعة من المعرفات الحالية للمقارنة السريعة
  const currentIdSet = new Set(currentIds.map(item => item.id));

  // البحث عن المعرفات المحذوفة
  const deletedItems = cachedData.filter(item => !currentIdSet.has(item.id));

  // تسجيل كل عنصر محذوف
  if (deletedItems.length > 0) {
    deletedItems.forEach(item => {
      logActivity(
        'حذف صف',
        sheetName,
        item.row,
        item.id,
        `تم حذف الصف ${item.row} (${sheetName === CONFIG.SHEETS.TRANSACTIONS ? 'حركة رقم ' + item.id : 'معرف: ' + item.id})`,
        {
          deletedId: item.id,
          deletedRow: item.row,
          totalDeleted: deletedItems.length
        },
        userEmail
      );
    });
  }

  // تحديث الكاش بالبيانات الحالية
  props.setProperty(cacheKey, JSON.stringify(currentIds));
}


// ==================== دوال ترتيب الحركات ====================
/**
 * ترتيب الحركات في دفتر الحركات المالية حسب التاريخ
 * الأقدم في الأعلى (صف 2) والأحدث في الأسفل (آخر صف)
 *
 * ⚠️ ملاحظة: هذه الدالة تحافظ على جميع المعادلات في الأعمدة المحسوبة:
 * A (رقم الحركة), M (القيمة بالدولار), O (الرصيد), P (رقم مرجعي),
 * U (تاريخ الاستحقاق), V (حالة السداد), W (الشهر)
 */
function sortTransactionsByDate() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.TRANSACTIONS);
  const projectsSheet = ss.getSheetByName(CONFIG.SHEETS.PROJECTS);

  if (!sheet) {
    ui.alert('❌ خطأ', 'شيت دفتر الحركات المالية غير موجود!', ui.ButtonSet.OK);
    return;
  }

  // تأكيد من المستخدم
  const response = ui.alert(
    '🔃 ترتيب الحركات',
    'سيتم ترتيب جميع الحركات حسب التاريخ:\n' +
    '• الأقدم في الأعلى (صف 2)\n' +
    '• الأحدث في الأسفل (آخر صف)\n' +
    '• ستتم إزالة الصفوف الفارغة (بدون تاريخ)\n' +
    '• سيتم إعادة بناء جميع المعادلات\n\n' +
    'هل تريد المتابعة؟',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    return;
  }

  const lastRow = sheet.getLastRow();
  const lastCol = Math.max(sheet.getLastColumn(), 24); // على الأقل 24 عمود (A-X)

  if (lastRow <= 1) {
    ui.alert('ℹ️ تنبيه', 'لا توجد حركات للترتيب!', ui.ButtonSet.OK);
    return;
  }

  // ═══════════════════════════════════════════════════════════
  // 1. قراءة كل البيانات
  // ═══════════════════════════════════════════════════════════
  const dataRange = sheet.getRange(2, 1, lastRow - 1, lastCol);
  const allData = dataRange.getValues();

  // ═══════════════════════════════════════════════════════════
  // 2. فلترة الصفوف الفارغة (صفوف بدون تاريخ صحيح في عمود B)
  // ═══════════════════════════════════════════════════════════
  const validRows = allData.filter(function (row) {
    const dateVal = row[1]; // B = index 1
    // تاريخ صحيح = كائن Date أو نص يمكن تحويله لتاريخ
    if (dateVal instanceof Date && !isNaN(dateVal.getTime())) {
      return true;
    }
    if (typeof dateVal === 'string' && dateVal.trim() !== '') {
      const parsed = new Date(dateVal);
      return !isNaN(parsed.getTime());
    }
    return false;
  });

  if (validRows.length === 0) {
    ui.alert('⚠️ تنبيه', 'لا توجد حركات بتواريخ صحيحة للترتيب!', ui.ButtonSet.OK);
    return;
  }

  // ═══════════════════════════════════════════════════════════
  // 3. ترتيب الصفوف حسب التاريخ (تصاعدي: الأقدم أولاً)
  // ═══════════════════════════════════════════════════════════

  // دالة مساعدة لتحويل التاريخ لـ timestamp بشكل صحيح
  function getDateTimestamp(dateVal) {
    // إذا كان Date object من Google Sheets
    if (dateVal instanceof Date && !isNaN(dateVal.getTime())) {
      return dateVal.getTime();
    }
    // إذا كان نصاً بصيغة dd/mm/yyyy
    if (typeof dateVal === 'string') {
      // إزالة أي شرطات مائلة مزدوجة أولاً
      dateVal = dateVal.replace(/\/+/g, '/').trim();
      const parts = dateVal.split('/');
      if (parts.length === 3) {
        const day = parseInt(parts[0], 10);
        const month = parseInt(parts[1], 10) - 1; // الشهر يبدأ من 0
        const year = parseInt(parts[2], 10);
        if (!isNaN(day) && !isNaN(month) && !isNaN(year)) {
          return new Date(year, month, day).getTime();
        }
      }
    }
    // محاولة أخيرة
    const parsed = new Date(dateVal);
    return isNaN(parsed.getTime()) ? 0 : parsed.getTime();
  }

  validRows.sort(function (a, b) {
    const dateA = getDateTimestamp(a[1]);
    const dateB = getDateTimestamp(b[1]);
    return dateA - dateB; // تصاعدي: الأقدم في الأعلى
  });

  // ═══════════════════════════════════════════════════════════
  // 4. مسح كل البيانات القديمة
  // ═══════════════════════════════════════════════════════════
  dataRange.clearContent();

  // ═══════════════════════════════════════════════════════════
  // 5. كتابة البيانات المرتبة (أعمدة البيانات فقط، بدون أعمدة المعادلات)
  // أعمدة البيانات: B, C, D, E, F, G, H, I, J, K, L, N, Q, R, S, T, X
  // أعمدة المعادلات/المحسوبة: A, M, O, P, U, V, W
  // ═══════════════════════════════════════════════════════════
  const numRows = validRows.length;

  // دالة مساعدة لتحويل التاريخ لـ Date object بشكل صحيح
  function ensureDateObject(dateVal) {
    if (dateVal instanceof Date) {
      return dateVal;
    }
    if (typeof dateVal === 'string') {
      // إزالة أي شرطات مائلة مزدوجة
      dateVal = dateVal.replace(/\/+/g, '/');
      const parts = dateVal.split('/');
      if (parts.length === 3) {
        const day = parseInt(parts[0], 10);
        const month = parseInt(parts[1], 10) - 1;
        const year = parseInt(parts[2], 10);
        return new Date(year, month, day);
      }
    }
    return new Date(dateVal);
  }

  // استخراج أعمدة البيانات فقط وكتابتها
  // B-L (indexes 1-11, columns 2-12)
  // مع التأكد من أن التاريخ (B) هو Date object صحيح
  const dataBtoL = validRows.map(function (row) {
    const rowData = row.slice(1, 12); // B to L
    // تحويل التاريخ (أول عنصر = B) لـ Date object
    rowData[0] = ensureDateObject(rowData[0]);
    return rowData;
  });
  sheet.getRange(2, 2, numRows, 11).setValues(dataBtoL);

  // N (index 13, column 14)
  const dataN = validRows.map(function (row) { return [row[13]]; });
  sheet.getRange(2, 14, numRows, 1).setValues(dataN);

  // Q-T (indexes 16-19, columns 17-20)
  // T (index 19) هو تاريخ مخصص - يجب التأكد من أنه Date object
  const dataQtoT = validRows.map(function (row) {
    const rowData = row.slice(16, 20); // Q to T
    // T = العنصر الرابع (index 3) - تحويله لـ Date إذا كان تاريخاً
    if (rowData[3]) {
      rowData[3] = ensureDateObject(rowData[3]);
    }
    return rowData;
  });
  sheet.getRange(2, 17, numRows, 4).setValues(dataQtoT);

  // X (index 23, column 24)
  const dataX = validRows.map(function (row) { return [row[23] || '']; });
  sheet.getRange(2, 24, numRows, 1).setValues(dataX);

  // ═══════════════════════════════════════════════════════════
  // 6. إعادة بناء معادلات الأعمدة: A, P, W
  // ═══════════════════════════════════════════════════════════
  const formulasA = [];
  const formulasP = [];
  const formulasW = [];

  for (let row = 2; row <= numRows + 1; row++) {
    // A: رقم الحركة
    formulasA.push([`=IF(B${row}="","",ROW()-1)`]);

    // P: رقم مرجعي (للحركات المدينة فقط)
    formulasP.push([
      `=IF(AND(N${row}="مدين استحقاق",B${row}<>""),` +
      `"REF-"&TEXT(B${row},"YYYYMMDD")&"-"&ROW(),"")`
    ]);

    // W: الشهر
    formulasW.push([`=IF(B${row}="","",TEXT(B${row},"YYYY-MM"))`]);
  }

  sheet.getRange(2, 1, numRows, 1).setFormulas(formulasA);   // A
  sheet.getRange(2, 16, numRows, 1).setFormulas(formulasP);  // P
  sheet.getRange(2, 23, numRows, 1).setFormulas(formulasW);  // W

  // ═══════════════════════════════════════════════════════════
  // 7. حساب القيم للأعمدة: M, O, U, V
  // (نفس منطق refreshValueAndBalanceFormulas ولكن للصفوف المرتبة)
  // ═══════════════════════════════════════════════════════════

  // جلب بيانات المشاريع (للحصول على تواريخ التسليم)
  const projectDeliveryDates = {};
  if (projectsSheet) {
    const projectData = projectsSheet.getRange('A2:K200').getValues();
    for (let i = 0; i < projectData.length; i++) {
      const code = projectData[i][0];
      const deliveryDate = projectData[i][10]; // K: تاريخ التسليم المتوقع
      if (code && deliveryDate instanceof Date) {
        projectDeliveryDates[code] = deliveryDate;
      }
    }
  }

  const valuesM = [];  // القيمة بالدولار
  const valuesO = [];  // الرصيد
  const valuesU = [];  // تاريخ الاستحقاق
  const valuesV = [];  // حالة السداد

  // لتتبع الأرصدة التراكمية لكل طرف
  const partyBalances = {};

  for (let i = 0; i < validRows.length; i++) {
    const row = validRows[i];
    const dateVal = row[1];                                   // B
    const projectCode = row[4];                               // E
    const party = String(row[8] || '').trim();                // I
    const amount = Number(row[9]) || 0;                       // J
    const currency = String(row[10] || '').trim().toUpperCase(); // K
    const exchangeRate = Number(row[11]) || 0;                // L
    const movementKind = String(row[13] || '').trim();        // N
    const paymentTermType = String(row[17] || '').trim();     // R
    const weeks = Number(row[18]) || 0;                       // S
    const customDate = row[19];                               // T

    // ─────────────────────────────────────────────────────────
    // M: حساب القيمة بالدولار
    // ─────────────────────────────────────────────────────────
    let amountUsd = 0;
    let hasValidConversion = true;
    if (amount > 0) {
      if (currency === 'USD' || currency === 'دولار' || currency === '') {
        amountUsd = amount;
      } else if (exchangeRate > 0) {
        amountUsd = amount / exchangeRate;
      } else {
        hasValidConversion = false;
      }
    }
    valuesM.push([hasValidConversion && amountUsd > 0 ? Math.round(amountUsd * 100) / 100 : '']);

    // ─────────────────────────────────────────────────────────
    // O, V: حساب الرصيد وحالة السداد
    // ✅ معالجة خاصة للتمويل: دائن دفعة لكن يزيد الرصيد (التزام للممول)
    // ─────────────────────────────────────────────────────────
    let balance = '';
    let status = '';
    const natureType = String(row[2] || '').trim(); // C: طبيعة الحركة
    const isFundingIn = natureType.indexOf('تمويل') !== -1 && natureType.indexOf('سداد تمويل') === -1;

    if (party && amountUsd > 0) {
      if (!partyBalances[party]) {
        partyBalances[party] = 0;
      }

      if (movementKind === 'مدين استحقاق') {
        partyBalances[party] += amountUsd;
      } else if (movementKind === 'دائن دفعة') {
        // ✅ تمويل = دائن دفعة لكن يزيد رصيد الممول (التزام علينا)
        if (isFundingIn) {
          partyBalances[party] += amountUsd;
        } else {
          partyBalances[party] -= amountUsd;
        }
      }

      balance = Math.round(partyBalances[party] * 100) / 100;

      if (movementKind === 'دائن دفعة') {
        status = CONFIG.PAYMENT_STATUS.OPERATION;
      } else if (balance > 0.01) {
        status = CONFIG.PAYMENT_STATUS.PENDING;
      } else {
        status = CONFIG.PAYMENT_STATUS.PAID;
      }
    }
    valuesO.push([balance]);
    valuesV.push([status]);

    // ─────────────────────────────────────────────────────────
    // U: حساب تاريخ الاستحقاق
    // ─────────────────────────────────────────────────────────
    let dueDate = '';

    if (movementKind === 'مدين استحقاق' && paymentTermType) {
      if (paymentTermType === 'فوري') {
        dueDate = dateVal;
      } else if (paymentTermType === 'بعد التسليم' && projectCode) {
        const deliveryDate = projectDeliveryDates[projectCode];
        if (deliveryDate) {
          const newDate = new Date(deliveryDate);
          newDate.setDate(newDate.getDate() + (weeks * 7));
          dueDate = newDate;
        }
      } else if (paymentTermType === 'تاريخ مخصص' && customDate) {
        dueDate = customDate;
      }
    }
    valuesU.push([dueDate]);
  }

  // كتابة القيم المحسوبة
  sheet.getRange(2, 13, numRows, 1).setValues(valuesM);  // M
  sheet.getRange(2, 15, numRows, 1).setValues(valuesO);  // O
  sheet.getRange(2, 21, numRows, 1).setValues(valuesU);  // U
  sheet.getRange(2, 22, numRows, 1).setValues(valuesV);  // V

  // ═══════════════════════════════════════════════════════════
  // 8. تنسيقات
  // ═══════════════════════════════════════════════════════════
  sheet.getRange(2, 2, numRows, 1).setNumberFormat('dd/mm/yyyy');   // B: التاريخ
  sheet.getRange(2, 10, numRows, 1).setNumberFormat('#,##0.00');    // J: المبلغ
  sheet.getRange(2, 12, numRows, 1).setNumberFormat('#,##0.0000');  // L: سعر الصرف
  sheet.getRange(2, 13, numRows, 1).setNumberFormat('#,##0.00');    // M: القيمة بالدولار
  sheet.getRange(2, 15, numRows, 1).setNumberFormat('#,##0.00');    // O: الرصيد
  sheet.getRange(2, 21, numRows, 1).setNumberFormat('dd/mm/yyyy');  // U: تاريخ الاستحقاق

  // ═══════════════════════════════════════════════════════════
  // 9. رسالة نجاح
  // ═══════════════════════════════════════════════════════════
  const removedRows = allData.length - validRows.length;
  let message = 'تم ترتيب ' + validRows.length + ' حركة حسب التاريخ.\n\n' +
    '• الأقدم في الأعلى (صف 2)\n' +
    '• الأحدث في الأسفل (آخر صف)\n' +
    '• تم إعادة بناء جميع المعادلات';

  if (removedRows > 0) {
    message += '\n• تم إزالة ' + removedRows + ' صف فارغ';
  }

  ui.alert('✅ تم الترتيب', message, ui.ButtonSet.OK);

  SpreadsheetApp.getActiveSpreadsheet().toast('تم ترتيب الحركات بنجاح!', '✅ تم', 3);
}


// ==================== إنشاء النظام - الجزء 1 ====================
function confirmReset() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    '⚠️ تحذير خطير',
    'هذا الإجراء سيحذف كل شيتات النظام ويعيد إنشائها من الصفر.\n\nلو حضرتك متأكد 100% اكتب كلمة: DELETE',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) {
    ui.alert('❌ تم إلغاء العملية.');
    return false;
  }

  if (response.getResponseText() !== 'DELETE') {
    ui.alert('❌ تم إلغاء العملية — كلمة السر غير صحيحة.');
    return false;
  }

  return true;
}

function setupPart1() {
  if (!confirmReset()) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // نحذف كل الشيتات القديمة
  const sheets = ss.getSheets();
  sheets.forEach(sheet => {
    try {
      ss.deleteSheet(sheet);
    } catch (e) { }
  });

  // إنشاء الشيتات الأساسية
  createTransactionsSheet(ss);   // دفتر الحركات (بنظام العملات)
  createProjectsSheet(ss);       // المشاريع
  createPartiesSheet(ss);        // 🆕 قاعدة بيانات الأطراف الموحدة
  createItemsSheet(ss);          // 🆕 قاعدة بيانات البنود (مبسطة)
  createBudgetsSheet(ss);        // الميزانيات
  createAlertsSheet(ss);         // التنبيهات
  createActivityLogSheet(ss);    // 🆕 سجل النشاط

  // 🆕 شيتات البنك وخزنة العهدة (دولار / ليرة)
  createBankAndCashSheets(ss);

  ui.alert(
    '✅ تم إنشاء الجزء 1 بنجاح!\n\n' +
    '🆕 التحديثات:\n' +
    '• نظام حركة مالية جديد (عملة أصلية + سعر صرف + قيمة بالدولار + نوع الحركة)\n' +
    '• قاعدة بيانات أطراف موحدة (مورد / عميل / ممول)\n' +
    '• قاعدة بيانات البنود\n' +
    '• شيتات البنك وخزنة العهدة بالدولار والليرة\n' +
    '• سجل النشاط (تتبع العمليات)\n' +
    '• التلوين حسب نوع الحركة فقط (استحقاق / دفعة)\n' +
    '• العملة الأساسية: USD\n\n' +
    'الآن اختر: 🔧 إنشاء النظام - الجزء 2 (لو موجود عندك في ملف آخر).'
  );
}


// ==================== 1. دفتر الحركات المالية (مع العملات + نوع الحركة) ====================
function createTransactionsSheet(ss) {
  let oldSheet = ss.getSheetByName(CONFIG.SHEETS.TRANSACTIONS);
  if (oldSheet) ss.deleteSheet(oldSheet);

  let sheet = ss.insertSheet(CONFIG.SHEETS.TRANSACTIONS);
  sheet.setTabColor(CONFIG.COLORS.TAB.TRANSACTIONS);   // أخضر لدفتر الحركات المالية

  const headers = [
    'رقم الحركة',          // 1 - A
    'التاريخ',             // 2 - B
    'طبيعة الحركة',        // 3 - C
    'تصنيف الحركة',        // 4 - D
    'كود المشروع',         // 5 - E
    'اسم المشروع',         // 6 - F
    'البند',               // 7 - G
    'التفاصيل',            // 8 - H
    'اسم المورد/الجهة',    // 9 - I

    // 💰 قلب الحركة المالية يبدأ من هنا:
    'المبلغ بالعملة الأصلية', // 10 - J
    'العملة',              // 11 - K
    'سعر الصرف',           // 12 - L
    'القيمة بالدولار',      // 13 - M
    'نوع الحركة',           // 14 - N (مدين استحقاق / دائن دفعة)

    'الرصيد',              // 15 - O
    'رقم مرجعي',           // 16 - P
    'طريقة الدفع',         // 17 - Q
    'نوع شرط الدفع',       // 18 - R
    'عدد الأسابيع',        // 19 - S
    'تاريخ مخصص',          // 20 - T
    'تاريخ الاستحقاق',     // 21 - U
    'حالة السداد',         // 22 - V
    'الشهر',               // 23 - W
    'ملاحظات',             // 24 - X
    '📄 كشف',              // 25 - Y (عمود روابط كشف الحساب)
    'رقم الأوردر',         // 26 - Z (لربط الحركات المشتركة)
    'عدد الوحدات',         // 27 - AA (لحساب تكلفة الوحدة)
    'مصدر الإدخال'         // 28 - AB (🤖 بوت / 📝 نموذج / ✍️ يدوي)
  ];

  sheet.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setBackground(CONFIG.COLORS.HEADER.TRANSACTIONS)
    .setFontColor(CONFIG.COLORS.TEXT.WHITE)
    .setFontWeight('bold')
    .setFontSize(11)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  const widths = [
    80,   // A
    100,  // B
    170,  // C
    170,  // D
    110,  // E
    180,  // F
    180,  // G
    220,  // H
    150,  // I
    130,  // J
    90,   // K
    110,  // L
    130,  // M
    130,  // N
    120,  // O
    120,  // P
    120,  // Q
    130,  // R
    100,  // S
    120,  // T
    130,  // U
    130,  // V
    90,   // W
    250,  // X
    60,   // Y (كشف)
    120,  // Z (رقم الأوردر)
    100,  // AA (عدد الوحدات)
    120   // AB (مصدر الإدخال)
  ];
  widths.forEach((width, i) => sheet.setColumnWidth(i + 1, width));

  const lastRow = 500;

  const projectsSheet = ss.getSheetByName(CONFIG.SHEETS.PROJECTS);
  const partiesSheet = ss.getSheetByName(CONFIG.SHEETS.PARTIES);
  const itemsSheet = ss.getSheetByName(CONFIG.SHEETS.ITEMS);

  // ✅ طبيعة الحركة من "قاعدة بيانات البنود" عمود B
  if (itemsSheet) {
    const movementRange = itemsSheet.getRange('B2:B200');
    const movementValidation = SpreadsheetApp.newDataValidation()
      .requireValueInRange(movementRange, true)
      .setAllowInvalid(true)
      .setHelpText('اختر طبيعة الحركة من "قاعدة بيانات البنود"')
      .build();
    sheet.getRange(2, 3, lastRow, 1) // C
      .setDataValidation(movementValidation)
      .setHorizontalAlignment('center');
  }

  // ✅ تصنيف الحركة من "قاعدة بيانات البنود" عمود C
  if (itemsSheet) {
    const classRange = itemsSheet.getRange('C2:C200');
    const classValidation = SpreadsheetApp.newDataValidation()
      .requireValueInRange(classRange, true)
      .setAllowInvalid(true)
      .setHelpText('اختر تصنيف الحركة من "قاعدة بيانات البنود"')
      .build();
    sheet.getRange(2, 4, lastRow, 1) // D
      .setDataValidation(classValidation)
      .setHorizontalAlignment('center');
  }

  // كود المشروع (E)
  if (projectsSheet) {
    const projectRange = projectsSheet.getRange('A2:A200');
    const projectValidation = SpreadsheetApp.newDataValidation()
      .requireValueInRange(projectRange, true)
      .setAllowInvalid(true)
      .setHelpText('اختر كود المشروع من القائمة أو اكتب يدوياً')
      .build();
    sheet.getRange(2, 5, lastRow, 1)
      .setDataValidation(projectValidation);

    // 🆕 اسم المشروع (F) - dropdown مرتبط بأسماء المشاريع
    const projectNameRange = projectsSheet.getRange('B2:B200');
    const projectNameValidation = SpreadsheetApp.newDataValidation()
      .requireValueInRange(projectNameRange, true)
      .setAllowInvalid(true)
      .setHelpText('اختر اسم المشروع - سيتم ملء كود المشروع تلقائياً')
      .build();
    sheet.getRange(2, 6, lastRow, 1) // F
      .setDataValidation(projectNameValidation);
  }

  // اسم المورد/الجهة (I) من قاعدة بيانات الأطراف
  if (partiesSheet) {
    const partyRange = partiesSheet.getRange('A2:A200');
    const partyValidation = SpreadsheetApp.newDataValidation()
      .requireValueInRange(partyRange, true)
      .setAllowInvalid(true)
      .setHelpText('اختر اسم الطرف من "قاعدة بيانات الأطراف" أو اكتب يدوياً')
      .build();
    sheet.getRange(2, 9, lastRow, 1) // I
      .setDataValidation(partyValidation);
  }

  // ✅ البند من "قاعدة بيانات البنود" عمود A (G)
  if (itemsSheet) {
    const itemsRange = itemsSheet.getRange('A2:A200');
    const itemValidation = SpreadsheetApp.newDataValidation()
      .requireValueInRange(itemsRange, true)
      .setAllowInvalid(true)
      .setHelpText('اختر البند من "قاعدة بيانات البنود" أو اكتب يدوياً')
      .build();
    sheet.getRange(2, 7, lastRow, 1) // G
      .setDataValidation(itemValidation);
  }

  // 🆕 دروب داون "نوع الحركة" (N)
  const movementTypeValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(CONFIG.MOVEMENT.TYPES, true)
    .setAllowInvalid(true)
    .setHelpText('اختر نوع الحركة: مدين استحقاق أو دائن دفعة')
    .build();
  sheet.getRange(2, 14, lastRow, 1) // N
    .setDataValidation(movementTypeValidation)
    .setHorizontalAlignment('center');

  // 🆕 دروب داون العملة (K)
  const currencyValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(CONFIG.CURRENCIES.LIST, true)
    .setAllowInvalid(true)
    .setHelpText('اختر العملة (USD / TRY / EGP)')
    .build();
  sheet.getRange(2, 11, lastRow, 1).setDataValidation(currencyValidation); // K

  // طريقة الدفع (Q = 17)
  const payMethodValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(['نقدي', 'تحويل بنكي', 'شيك', 'بطاقة', 'أخرى'])
    .setAllowInvalid(true)
    .setHelpText('اختر طريقة الدفع')
    .build();
  sheet.getRange(2, 17, lastRow, 1) // Q
    .setDataValidation(payMethodValidation);

  // نوع شرط الدفع (R = 18)
  const termValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(CONFIG.PAYMENT_TERMS.LIST)
    .setAllowInvalid(true)
    .setHelpText('اختر شرط الدفع')
    .build();
  sheet.getRange(2, 18, lastRow, 1) // R
    .setDataValidation(termValidation);

  // عدد الأسابيع (S = 19) - validation للأرقام فقط 0-52
  const weeksValidation = SpreadsheetApp.newDataValidation()
    .requireNumberBetween(0, 52)
    .setAllowInvalid(false)
    .setHelpText('أدخل عدد الأسابيع (0-52) - يُستخدم مع شرط "بعد التسليم"')
    .build();
  sheet.getRange(2, 19, lastRow, 1) // S
    .setDataValidation(weeksValidation)
    .setValue(0);  // قيمة افتراضية = 0

  // المعادلات لكل صف - باستخدام Batch Operations للأداء الأمثل
  // بدلاً من 4000 طلب API (7 معادلات × 500 صف) = 7 طلبات فقط
  // ملاحظة: عمود F (اسم المشروع) يُملأ عبر onEdit للمزامنة الثنائية مع E
  const numRows = lastRow - 1;

  const formulasA = [];  // رقم الحركة (A)
  const formulasM = [];  // القيمة بالدولار (M)
  const formulasO = [];  // الرصيد (O)
  const formulasP = [];  // رقم مرجعي (P)
  const formulasU = [];  // تاريخ الاستحقاق (U)
  const formulasV = [];  // حالة السداد (V)
  const formulasW = [];  // الشهر (W)

  for (let row = 2; row <= lastRow; row++) {
    // رقم الحركة (A)
    formulasA.push([`=IF(B${row}="","",ROW()-1)`]);

    // القيمة بالدولار (M)
    // إذا العملة دولار (USD أو دولار) → نفس قيمة J
    // إذا العملة أخرى (TRY/EGP/ليرة/جنيه) → J ÷ سعر الصرف (L)
    // ⚠️ إذا العملة أخرى ولا يوجد سعر صرف → ترك فارغ (يحتاج إدخال سعر الصرف)
    formulasM.push([
      `=IF(J${row}="","",` +
      `IF(OR(K${row}="USD",K${row}="دولار",K${row}=""),J${row},` +
      `IF(OR(L${row}="",L${row}=0),"",J${row}/L${row})))`
    ]);

    // الرصيد O = مجموع (مدين استحقاق - دائن دفعة) لنفس الطرف حتى هذا الصف
    formulasO.push([
      `=IF(OR(I${row}="",M${row}=""),"",` +
      `SUMIFS($M$2:M${row},$I$2:I${row},I${row},$N$2:N${row},"مدين استحقاق")-` +
      `SUMIFS($M$2:M${row},$I$2:I${row},I${row},$N$2:N${row},"دائن دفعة"))`
    ]);

    // رقم مرجعي P (16) للحركات المدينة
    formulasP.push([
      `=IF(AND(N${row}="مدين استحقاق",B${row}<>""),` +
      `"REF-"&TEXT(B${row},"YYYYMMDD")&"-"&ROW(),"")`
    ]);

    // تاريخ الاستحقاق U (21) - محسّن للتعامل مع القيم الفارغة
    // فوري = تاريخ الحركة
    // بعد التسليم = تاريخ التسليم المتوقع + (عدد الأسابيع × 7)
    // تاريخ مخصص = التاريخ المُدخل يدوياً
    // ملاحظة: التحقق من وجود تاريخ التسليم قبل الحساب لتجنب 30/12/1899
    formulasU.push([
      `=IF(OR(N${row}<>"مدين استحقاق",R${row}=""),"",` +
      `IF(R${row}="فوري",B${row},` +
      `IF(R${row}="بعد التسليم",` +
      `IF(OR(E${row}="",NOT(ISNUMBER(VLOOKUP(E${row},'قاعدة بيانات المشاريع'!A2:K200,11,FALSE)))),"",` +
      `VLOOKUP(E${row},'قاعدة بيانات المشاريع'!A2:K200,11,FALSE)+IF(OR(S${row}="",S${row}=0),0,S${row})*7),` +
      `IF(AND(R${row}="تاريخ مخصص",T${row}<>""),T${row},""))))`
    ]);

    // حالة السداد V (22)
    formulasV.push([
      `=IF(N${row}="مدين استحقاق",` +
      `IF(O${row}<=0,"مدفوع بالكامل","معلق"),` +
      `IF(N${row}="دائن دفعة","عملية دفع/تحصيل",""))`
    ]);

    // الشهر W (23)
    formulasW.push([`=IF(B${row}="","",TEXT(B${row},"YYYY-MM"))`]);
  }

  // تطبيق كل المعادلات دفعة واحدة (7 طلبات بدلاً من 3500)
  sheet.getRange(2, 1, numRows, 1).setFormulas(formulasA);   // A: رقم الحركة
  sheet.getRange(2, 13, numRows, 1).setFormulas(formulasM);  // M: القيمة بالدولار
  sheet.getRange(2, 15, numRows, 1).setFormulas(formulasO);  // O: الرصيد
  sheet.getRange(2, 16, numRows, 1).setFormulas(formulasP);  // P: رقم مرجعي
  sheet.getRange(2, 21, numRows, 1).setFormulas(formulasU);  // U: تاريخ الاستحقاق
  sheet.getRange(2, 22, numRows, 1).setFormulas(formulasV);  // V: حالة السداد
  sheet.getRange(2, 23, numRows, 1).setFormulas(formulasW);  // W: الشهر

  // تنسيقات الأرقام والتواريخ
  sheet.getRange(2, 10, lastRow, 1).setNumberFormat('#,##0.00');   // J
  sheet.getRange(2, 12, lastRow, 1).setNumberFormat('#,##0.0000'); // L
  sheet.getRange(2, 13, lastRow, 1).setNumberFormat('#,##0.00');   // M
  sheet.getRange(2, 15, lastRow, 1).setNumberFormat('#,##0.00');   // O

  sheet.getRange(2, 2, lastRow, 1).setNumberFormat('dd/mm/yyyy'); // B - التاريخ
  sheet.getRange(2, 20, lastRow, 1).setNumberFormat('dd/mm/yyyy'); // T - تاريخ مخصص
  sheet.getRange(2, 21, lastRow, 1).setNumberFormat('dd/mm/yyyy'); // U - تاريخ الاستحقاق

  // 🎨 تلوين شرطي حسب نوع الحركة فقط
  applyConditionalFormatting(sheet, lastRow);

  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(2);

  sheet.getRange('N1').setNote(
    'نوع الحركة:\n' +
    '• مدين استحقاق = فاتورة/استحقاق على الطرف\n' +
    '• دائن دفعة = دفعة/تحصيل تقلل الرصيد'
  );
}

// ==================== التلوين الشرطي (حسب نوع الحركة فقط) ====================
function applyConditionalFormatting(sheet, lastRow) {
  // مسح القواعد القديمة أولاً
  sheet.clearConditionalFormatRules();

  const rules = [];
  // استخدام نطاق أكبر لضمان شمول الصفوف الجديدة
  const maxRows = Math.max(lastRow, 1000);
  const dataRange = sheet.getRange(2, 1, maxRows, 24); // من A إلى X

  // ═══════════════════════════════════════════════════════════
  // 1. تمويل (دخول قرض) = أخضر مميز - الأولوية الأعلى
  // ═══════════════════════════════════════════════════════════
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND($C2<>"",ISNUMBER(SEARCH("تمويل",$C2)),ISERROR(SEARCH("سداد",$C2)),ISERROR(SEARCH("استلام",$C2)))')
      .setBackground('#a5d6a7')  // أخضر مميز
      .setRanges([dataRange])
      .build()
  );

  // ═══════════════════════════════════════════════════════════
  // 2. استحقاق = برتقالي فاتح
  // ═══════════════════════════════════════════════════════════
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND($N2<>"",$N2="مدين استحقاق")')
      .setBackground(CONFIG.COLORS.BG.LIGHT_ORANGE)
      .setRanges([dataRange])
      .build()
  );

  // ═══════════════════════════════════════════════════════════
  // 3. دفعة = أزرق فاتح
  // ═══════════════════════════════════════════════════════════
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND($N2<>"",$N2="دائن دفعة")')
      .setBackground(CONFIG.COLORS.BG.LIGHT_BLUE)
      .setRanges([dataRange])
      .build()
  );

  sheet.setConditionalFormatRules(rules);
}

/**
 * إعادة تطبيق التلوين الشرطي على دفتر الحركات المالية
 * يُستدعى من القائمة لإصلاح التلوين على الصفوف الموجودة والجديدة
 */
function refreshTransactionsFormatting() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.TRANSACTIONS);

  if (!sheet) {
    SpreadsheetApp.getUi().alert('❌ لم يتم العثور على شيت "دفتر الحركات المالية"');
    return;
  }

  const lastRow = Math.max(sheet.getLastRow(), 500);
  applyConditionalFormatting(sheet, lastRow);

  SpreadsheetApp.getUi().alert(
    '✅ تم تحديث التلوين الشرطي',
    'تم إعادة تطبيق التلوين الشرطي على دفتر الحركات المالية.\n\n' +
    '• 🏦 تمويل (دخول قرض) = أخضر فاتح 🟩\n' +
    '• مدين استحقاق = برتقالي فاتح 🟧\n' +
    '• دائن دفعة = أزرق فاتح 🟦\n\n' +
    'النطاق: ' + lastRow + ' صف',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * تحديث معادلة تاريخ الاستحقاق (U) على جميع الصفوف
 * المنطق:
 * - إذا N = "مدين استحقاق" و R = "فوري" → U = تاريخ الحركة (B)
 * - إذا N = "مدين استحقاق" و R = "بعد التسليم" → U = تاريخ التسليم من المشاريع + S أسابيع
 * - إذا N = "مدين استحقاق" و R = "تاريخ مخصص" → U = T (التاريخ المُدخل يدوياً)
 */
function refreshDueDateFormulas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.TRANSACTIONS);
  const ui = SpreadsheetApp.getUi();

  if (!sheet) {
    ui.alert('❌ لم يتم العثور على شيت "دفتر الحركات المالية"');
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    ui.alert('⚠️ لا توجد بيانات في الشيت');
    return;
  }

  // بناء المعادلات لكل صف
  // التحقق من وجود تاريخ التسليم قبل الحساب لتجنب 30/12/1899
  const formulas = [];
  for (let row = 2; row <= lastRow; row++) {
    formulas.push([
      `=IF(OR(N${row}<>"مدين استحقاق",R${row}=""),"",` +
      `IF(R${row}="فوري",B${row},` +
      `IF(R${row}="بعد التسليم",` +
      `IF(OR(E${row}="",NOT(ISNUMBER(VLOOKUP(E${row},'قاعدة بيانات المشاريع'!A2:K200,11,FALSE)))),"",` +
      `VLOOKUP(E${row},'قاعدة بيانات المشاريع'!A2:K200,11,FALSE)+IF(OR(S${row}="",S${row}=0),0,S${row})*7),` +
      `IF(AND(R${row}="تاريخ مخصص",T${row}<>""),T${row},""))))`
    ]);
  }

  // تطبيق المعادلات على العمود U
  sheet.getRange(2, 21, lastRow - 1, 1).setFormulas(formulas);

  // تنسيق العمود كتاريخ
  sheet.getRange(2, 21, lastRow - 1, 1).setNumberFormat('dd/mm/yyyy');

  ui.alert(
    '✅ تم تحديث معادلة تاريخ الاستحقاق',
    'تم تطبيق المعادلة على العمود U لجميع الصفوف.\n\n' +
    '📋 المنطق:\n' +
    '• فوري → تاريخ الحركة\n' +
    '• بعد التسليم → تاريخ التسليم + أسابيع\n' +
    '• تاريخ مخصص → العمود T\n\n' +
    '📊 عدد الصفوف: ' + (lastRow - 1),
    ui.ButtonSet.OK
  );
}

/**
 * تحديث شامل للأعمدة المحسوبة: M (القيمة بالدولار), O (الرصيد), U (تاريخ الاستحقاق), V (حالة السداد)
 * هذه الدالة تحسب القيم وتكتبها مباشرة (بدون معادلات) لحماية البيانات من أخطاء المستخدمين
 */
function refreshValueAndBalanceFormulas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.TRANSACTIONS);
  const projectsSheet = ss.getSheetByName(CONFIG.SHEETS.PROJECTS);
  const ui = SpreadsheetApp.getUi();

  if (!sheet) {
    ui.alert('❌ لم يتم العثور على شيت "دفتر الحركات المالية"');
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    ui.alert('⚠️ لا توجد بيانات في الشيت');
    return;
  }

  // جلب بيانات المشاريع (للحصول على تواريخ التسليم)
  const projectDeliveryDates = {};
  if (projectsSheet) {
    const projectData = projectsSheet.getRange('A2:K200').getValues();
    for (let i = 0; i < projectData.length; i++) {
      const code = projectData[i][0];
      const deliveryDate = projectData[i][10]; // K: تاريخ التسليم المتوقع
      if (code && deliveryDate instanceof Date) {
        projectDeliveryDates[code] = deliveryDate;
      }
    }
  }

  // قراءة كل البيانات المطلوبة مرة واحدة
  // A=1, B=2, E=5, I=9, J=10, K=11, L=12, N=14, R=18, S=19, T=20
  const dataRange = sheet.getRange(2, 1, lastRow - 1, 20); // A to T
  const data = dataRange.getValues();

  const valuesM = [];  // القيمة بالدولار (M) - column 13
  const valuesO = [];  // الرصيد (O) - column 15
  const valuesU = [];  // تاريخ الاستحقاق (U) - column 21
  const valuesV = [];  // حالة السداد (V) - column 22

  // لتتبع الأرصدة التراكمية لكل طرف
  const partyBalances = {};

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const dateVal = row[1];                    // B: تاريخ الحركة (index 1)
    const projectCode = row[4];                // E: كود المشروع (index 4)
    const party = String(row[8] || '').trim(); // I: الطرف (index 8)
    const amount = Number(row[9]) || 0;        // J: المبلغ (index 9)
    const currency = String(row[10] || '').trim().toUpperCase(); // K: العملة (index 10)
    const exchangeRate = Number(row[11]) || 0; // L: سعر الصرف (index 11)
    const movementKind = String(row[13] || '').trim(); // N: نوع الحركة (index 13)
    const paymentTermType = String(row[17] || '').trim(); // R: نوع شرط الدفع (index 17)
    const weeks = Number(row[18]) || 0;        // S: عدد الأسابيع (index 18)
    const customDate = row[19];                // T: تاريخ مخصص (index 19)

    // ═══════════════════════════════════════════════════════════
    // 1. حساب القيمة بالدولار (M)
    // ⚠️ إذا كانت العملة غير دولار ولا يوجد سعر صرف = ترك الخلية فارغة
    // ═══════════════════════════════════════════════════════════
    let amountUsd = 0;
    let hasValidConversion = true;
    if (amount > 0) {
      // حالة 1: العملة دولار أو فارغة (افتراضي دولار)
      if (currency === 'USD' || currency === 'دولار' || currency === '') {
        amountUsd = amount;
      }
      // حالة 2: عملة أخرى مع سعر صرف صحيح
      else if (exchangeRate > 0) {
        amountUsd = amount / exchangeRate;
      }
      // حالة 3: عملة أخرى بدون سعر صرف = ترك فارغ (⚠️ يحتاج إدخال سعر الصرف)
      else {
        hasValidConversion = false;
      }
    }
    valuesM.push([hasValidConversion && amountUsd > 0 ? Math.round(amountUsd * 100) / 100 : '']);

    // ═══════════════════════════════════════════════════════════
    // 2. حساب الرصيد (O) وحالة السداد (V)
    // ✅ معالجة خاصة للتمويل: دائن دفعة لكن يزيد الرصيد (التزام للممول)
    // ═══════════════════════════════════════════════════════════
    let balance = '';
    let status = '';
    const natureType = String(row[2] || '').trim(); // C: طبيعة الحركة
    const isFundingIn = natureType.indexOf('تمويل') !== -1 && natureType.indexOf('سداد تمويل') === -1;

    if (party && amountUsd > 0) {
      if (!partyBalances[party]) {
        partyBalances[party] = 0;
      }

      if (movementKind === 'مدين استحقاق') {
        partyBalances[party] += amountUsd;
      } else if (movementKind === 'دائن دفعة') {
        // ✅ تمويل = دائن دفعة لكن يزيد رصيد الممول (التزام علينا)
        if (isFundingIn) {
          partyBalances[party] += amountUsd;
        } else {
          partyBalances[party] -= amountUsd;
        }
      }

      balance = Math.round(partyBalances[party] * 100) / 100;

      // حساب حالة السداد (باستخدام CONFIG.PAYMENT_STATUS للتوحيد)
      if (movementKind === 'دائن دفعة') {
        status = CONFIG.PAYMENT_STATUS.OPERATION; // 'عملية دفع/تحصيل'
      } else if (balance > 0.01) {
        status = CONFIG.PAYMENT_STATUS.PENDING; // 'معلق'
      } else {
        status = CONFIG.PAYMENT_STATUS.PAID; // 'مدفوع بالكامل'
      }
    }
    valuesO.push([balance]);
    valuesV.push([status]);

    // ═══════════════════════════════════════════════════════════
    // 3. حساب تاريخ الاستحقاق (U)
    // ═══════════════════════════════════════════════════════════
    let dueDate = '';

    if (movementKind === 'مدين استحقاق' && paymentTermType) {
      if (paymentTermType === 'فوري') {
        dueDate = dateVal;
      } else if (paymentTermType === 'بعد التسليم' && projectCode) {
        const deliveryDate = projectDeliveryDates[projectCode];
        if (deliveryDate) {
          const newDate = new Date(deliveryDate);
          newDate.setDate(newDate.getDate() + (weeks * 7));
          dueDate = newDate;
        }
      } else if (paymentTermType === 'تاريخ مخصص' && customDate) {
        dueDate = customDate;
      }
    }
    valuesU.push([dueDate]);
  }

  const numRows = lastRow - 1;

  // كتابة كل القيم دفعة واحدة (بدون معادلات)
  sheet.getRange(2, 13, numRows, 1).setValues(valuesM);  // M: القيمة بالدولار
  sheet.getRange(2, 15, numRows, 1).setValues(valuesO);  // O: الرصيد
  sheet.getRange(2, 21, numRows, 1).setValues(valuesU);  // U: تاريخ الاستحقاق
  sheet.getRange(2, 22, numRows, 1).setValues(valuesV);  // V: حالة السداد

  // تنسيقات
  sheet.getRange(2, 13, numRows, 1).setNumberFormat('#,##0.00');  // M
  sheet.getRange(2, 15, numRows, 1).setNumberFormat('#,##0.00');  // O
  sheet.getRange(2, 21, numRows, 1).setNumberFormat('dd/mm/yyyy'); // U

  ui.alert(
    '✅ تم التحديث الشامل للأعمدة المحسوبة',
    'تم حساب وكتابة القيم (بدون معادلات) في:\n\n' +
    '• M - القيمة بالدولار: المبلغ ÷ سعر الصرف (أو نفسه للدولار)\n' +
    '   ⚠️ إذا كانت العملة غير دولار ولا يوجد سعر صرف = ترك فارغ\n' +
    '• O - الرصيد: مدين استحقاق - دائن دفعة لكل طرف\n' +
    '• U - تاريخ الاستحقاق: حسب نوع شرط الدفع\n' +
    '• V - حالة السداد: معلق / مدفوع بالكامل / عملية دفع/تحصيل\n\n' +
    '⚡ الحسابات تتم تلقائياً عند تعديل البيانات (onEdit)\n\n' +
    '📊 عدد الصفوف المحدثة: ' + numRows,
    ui.ButtonSet.OK
  );
}


// ==================== 2. قاعدة بيانات المشاريع ====================
function createProjectsSheet(ss) {
  let oldSheet = ss.getSheetByName(CONFIG.SHEETS.PROJECTS);
  if (oldSheet) ss.deleteSheet(oldSheet);

  let sheet = ss.insertSheet(CONFIG.SHEETS.PROJECTS);

  const headers = [
    'كود المشروع', 'اسم المشروع', 'نوع المشروع', 'القناة/الجهة',
    'اسم البرنامج', 'سنة الإنتاج', 'نوع التمويل', 'قيمة التمويل',
    'قيمة العقد مع القناة', 'تاريخ البدء', 'تاريخ التسليم المتوقع',
    'تاريخ التسليم الفعلي', 'المدة (أسابيع)', '🆕 مدة المشروع (أشهر)',
    'حالة المشروع', 'ملاحظات'
  ];

  sheet.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setBackground(CONFIG.COLORS.HEADER.PROJECTS)
    .setFontColor(CONFIG.COLORS.TEXT.WHITE)
    .setFontWeight('bold')
    .setFontSize(11)
    .setHorizontalAlignment('center');

  const widths = [150, 200, 130, 150, 150, 100, 130, 130, 150, 120, 150, 150, 120, 150, 130, 250];
  widths.forEach((width, i) => sheet.setColumnWidth(i + 1, width));

  // القوائم
  sheet.getRange(2, 3, 200, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(['وثائقي قصير', 'وثائقي طويل', 'سلسلة وثائقية', 'تقرير إخباري', 'فيلم روائي', 'برومو'])
      .build()
  );

  sheet.getRange(2, 7, 200, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(['خارجي', 'ذاتي', 'مشترك', 'لا يوجد'])
      .build()
  );

  const years = [];
  for (let y = 2020; y <= 2030; y++) years.push(y.toString());
  sheet.getRange(2, 6, 200, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(years)
      .build()
  );

  sheet.getRange(2, 15, 200, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(['تخطيط', 'جاري التنفيذ', 'تصوير', 'مونتاج', 'مراجعة', 'مكتمل', 'مسلّم', 'ملغي', 'معلق'])
      .build()
  );

  /**
   * ⚡ تحسينات الأداء:
   * - Batch Operations: 2 API calls بدلاً من 198 (99×2)
   */
  const numRows = 99;
  const formulasA = [];  // كود المشروع
  const formulasM = [];  // المدة (أسابيع)

  for (let row = 2; row <= 100; row++) {
    // كود المشروع (A)
    formulasA.push([
      `=IF(OR(D${row}="",E${row}="",F${row}=""),"",` +
      `UPPER(LEFT(D${row},2))&"-"&UPPER(LEFT(E${row},2))&"-"&` +
      `RIGHT(F${row},2)&"-"&TEXT(COUNTIFS($D$2:D${row},D${row},$E$2:E${row},E${row},$F$2:F${row},F${row}),"000"))`
    ]);
    // المدة بالأسابيع (M - column 13)
    formulasM.push([
      `=IF(OR(J${row}="",K${row}=""),"",ROUND((K${row}-J${row})/7,1))`
    ]);
  }

  // Batch apply formulas
  sheet.getRange(2, 1, numRows, 1).setFormulas(formulasA);
  sheet.getRange(2, 13, numRows, 1).setFormulas(formulasM);

  // تنسيق
  sheet.getRange(2, 8, 200, 2).setNumberFormat('$#,##0.00');
  sheet.getRange(2, 10, 200, 1).setNumberFormat('dd/mm/yyyy'); // J - تاريخ البدء
  sheet.getRange(2, 11, 200, 1).setNumberFormat('dd/mm/yyyy'); // K - تاريخ التسليم المتوقع
  sheet.getRange(2, 12, 200, 1).setNumberFormat('dd/mm/yyyy'); // L - تاريخ التسليم الفعلي
  sheet.getRange(2, 14, 200, 1).setNumberFormat('0');

  // تلوين حالة المشروع
  const rules = [];
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('مكتمل')
      .setBackground(CONFIG.COLORS.BG.LIGHT_GREEN_3)
      .setRanges([sheet.getRange(2, 15, 200, 1)])
      .build()
  );
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('جاري التنفيذ')
      .setBackground(CONFIG.COLORS.BG.LIGHT_YELLOW)
      .setRanges([sheet.getRange(2, 15, 200, 1)])
      .build()
  );
  // تظليل اسم المشروع عند وجود رقم فاتورة في العمود Q
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND($Q2<>"",$B2<>"")')
      .setBackground('#a8e6cf')  // أخضر فاتح مميز
      .setRanges([sheet.getRange(2, 2, 200, 1)])  // عمود B فقط
      .build()
  );

  sheet.setConditionalFormatRules(rules);

  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(2);  // تثبيت عمود كود المشروع واسم المشروع

  const protection = sheet.getRange(2, 1, 200, 1).protect();
  protection.setDescription('كود المشروع محسوب تلقائياً');
  protection.setWarningOnly(true);

  sheet.getRange('N1').setNote('🆕 مدة المشروع بالأشهر\nيُستخدم لحساب المصروفات العمومية 30% في تقرير الربحية');
}

/**
 * تطبيق إعدادات تثبيت الأعمدة وتظليل الفواتير على شيت المشاريع الموجود
 */
function applyProjectsSheetEnhancements() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.PROJECTS);

  if (!sheet) {
    SpreadsheetApp.getUi().alert('⚠️ لم يتم العثور على شيت "قاعدة بيانات المشاريع"');
    return;
  }

  // تثبيت العمودين الأولين
  sheet.setFrozenColumns(2);

  // الحصول على القواعد الموجودة
  const existingRules = sheet.getConditionalFormatRules();

  // إضافة قاعدة تظليل الفاتورة إذا لم تكن موجودة
  const invoiceRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($Q2<>"",$B2<>"")')
    .setBackground('#a8e6cf')  // أخضر فاتح مميز
    .setRanges([sheet.getRange(2, 2, 500, 1)])  // عمود B
    .build();

  existingRules.push(invoiceRule);
  sheet.setConditionalFormatRules(existingRules);

  SpreadsheetApp.getUi().alert('✅ تم تطبيق التحسينات:\n• تثبيت عمودي كود المشروع واسم المشروع\n• تظليل اسم المشروع عند وجود رقم فاتورة');
}


// ==================== 3. قاعدة بيانات الأطراف (مورد / عميل / ممول) ====================
function createPartiesSheet(ss) {
  const sheet = getOrCreateSheet_(ss, CONFIG.SHEETS.PARTIES, true);

  const headers = [
    'اسم الطرف',      // A
    'نوع الطرف',      // B (مورد / عميل / ممول)
    'التخصص / الفئة', // C
    'رقم الهاتف',     // D
    'البريد الإلكتروني', // E
    'المدينة / الدولة', // F
    'طريقة الدفع المفضلة', // G
    'بيانات الحساب البنكي / شروط خاصة', // H
    'ملاحظات'         // I
  ];
  const widths = [200, 140, 160, 140, 220, 160, 170, 260, 260];

  setupSheet_(sheet, headers, widths, CONFIG.COLORS.HEADER.PARTIES);

  // نوع الطرف
  sheet.getRange(2, 2, 500, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(['مورد', 'عميل', 'ممول'], true)
      .build()
  );

  // طريقة الدفع المفضلة
  sheet.getRange(2, 7, 500, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(['نقدي', 'تحويل بنكي', 'شيك', 'بطاقة', 'أخرى'], true)
      .build()
  );

  sheet.getRange('A1').setNote(
    'قاعدة موحدة لكل الأطراف (موردين / عملاء / ممولين)\n' +
    'يتم الربط مع دفتر الحركات من عمود "اسم المورد/الجهة".'
  );
}


// ==================== 4. قاعدة بيانات البنود (مدمجة) ====================
function createItemsSheet(ss) {
  const sheet = getOrCreateSheet_(ss, CONFIG.SHEETS.ITEMS, true);

  const headers = [
    'البند',           // A
    'طبيعة الحركة',    // B
    'تصنيف الحركة',    // C
    'ملاحظات'          // D
  ];
  const widths = [200, 180, 180, 250];

  setupSheet_(sheet, headers, widths, CONFIG.COLORS.HEADER.ITEMS);

  // البيانات التجريبية
  const sampleData = [
    ['مونتاج', 'استحقاق مصروف', 'مصروفات مباشرة', ''],
    ['تصوير', 'استحقاق مصروف', 'مصروفات مباشرة', ''],
    ['صوت', 'استحقاق مصروف', 'مصروفات مباشرة', ''],
    ['معدات', 'استحقاق مصروف', 'مصروفات مباشرة', ''],
    ['🏢 إيجار مكتب', 'استحقاق مصروف', 'مصروفات عمومية', ''],
    ['👥 مرتبات إدارية', 'استحقاق مصروف', 'مصروفات عمومية', ''],
    ['⚡ مرافق', 'استحقاق مصروف', 'مصروفات عمومية', ''],
    ['🧾 أخرى', 'استحقاق مصروف', 'مصروفات أخرى', '']
  ];
  sheet.getRange(2, 1, sampleData.length, sampleData[0].length).setValues(sampleData);

  // ملاحظات الأعمدة
  sheet.getRange('B1').setNote(
    'طبيعة الحركة (مثال):\n' +
    'استحقاق مصروف / دفعة مصروف / استحقاق إيراد / تحصيل إيراد / تمويل / سداد تمويل'
  );

  sheet.getRange('C1').setNote(
    'تصنيف الحركة (مثال):\n' +
    'مصروفات مباشرة / مصروفات عمومية / تحصيل فواتير / استلام قرض / سداد قرض'
  );

  // إرجاع الشيت
  return sheet;
}


// ==================== 5. شيت الميزانيات ====================
function createBudgetsSheet(ss) {
  let oldSheet = ss.getSheetByName(CONFIG.SHEETS.BUDGETS);
  if (oldSheet) ss.deleteSheet(oldSheet);

  let sheet = ss.insertSheet(CONFIG.SHEETS.BUDGETS);

  const headers = [
    'كود المشروع', 'اسم المشروع', 'البند', 'المبلغ المخطط',
    'المبلغ الفعلي', 'الفرق', 'نسبة التنفيذ %', 'ملاحظات'
  ];

  sheet.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setBackground(CONFIG.COLORS.HEADER.BUDGETS)
    .setFontColor(CONFIG.COLORS.TEXT.WHITE)
    .setFontWeight('bold')
    .setFontSize(11);

  const widths = [120, 180, 150, 120, 120, 120, 130, 250];
  widths.forEach((width, i) => sheet.setColumnWidth(i + 1, width));

  const projectsSheet = ss.getSheetByName(CONFIG.SHEETS.PROJECTS);
  const itemsSheet = ss.getSheetByName(CONFIG.SHEETS.ITEMS);

  // كود المشروع (A)
  if (projectsSheet) {
    const projectRange = projectsSheet.getRange('A2:A200');
    const projectValidation = SpreadsheetApp.newDataValidation()
      .requireValueInRange(projectRange, true)
      .setAllowInvalid(true)
      .setHelpText('اختر كود المشروع من القائمة - سيتم ملء اسم المشروع تلقائياً')
      .build();
    sheet.getRange(2, 1, 100, 1).setDataValidation(projectValidation);

    // 🆕 اسم المشروع (B) - dropdown مرتبط بأسماء المشاريع
    const projectNameRange = projectsSheet.getRange('B2:B200');
    const projectNameValidation = SpreadsheetApp.newDataValidation()
      .requireValueInRange(projectNameRange, true)
      .setAllowInvalid(true)
      .setHelpText('اختر اسم المشروع - سيتم ملء كود المشروع تلقائياً')
      .build();
    sheet.getRange(2, 2, 100, 1).setDataValidation(projectNameValidation);
  }

  // البند من قاعدة بيانات البنود
  if (itemsSheet) {
    const itemsRange = itemsSheet.getRange('A2:A200');
    const itemValidation = SpreadsheetApp.newDataValidation()
      .requireValueInRange(itemsRange, true)
      .setAllowInvalid(true)
      .setHelpText('اختر البند من "قاعدة بيانات البنود"')
      .build();
    sheet.getRange(2, 3, 100, 1).setDataValidation(itemValidation);
  }

  /**
   * ⚡ تحسينات الأداء:
   * - Batch Operations: 3 API calls بدلاً من 297 (99×3)
   * - نطاقات محددة بدل أعمدة كاملة (M2:M1000 بدل M:M)
   * - عمود B (اسم المشروع) يُملأ عبر onEdit للمزامنة الثنائية مع A
   */
  const numRows = 99;
  const formulasE = [];  // المبلغ الفعلي
  const formulasF = [];  // الفرق
  const formulasG = [];  // نسبة التنفيذ

  for (let row = 2; row <= 100; row++) {
    // المبلغ الفعلي = مجموع القيمة بالدولار من دفتر الحركات (مدين استحقاق فقط) (E)
    formulasE.push([
      `=SUMIFS('دفتر الحركات المالية'!M2:M1000,` +
      `'دفتر الحركات المالية'!E2:E1000,A${row},` +
      `'دفتر الحركات المالية'!G2:G1000,C${row},` +
      `'دفتر الحركات المالية'!N2:N1000,"مدين استحقاق")`
    ]);
    // الفرق (F)
    formulasF.push([`=IF(D${row}="","",D${row}-E${row})`]);
    // نسبة التنفيذ (G)
    formulasG.push([`=IF(D${row}=0,"",E${row}/D${row})`]);
  }

  // Batch apply formulas
  sheet.getRange(2, 5, numRows, 1).setFormulas(formulasE);
  sheet.getRange(2, 6, numRows, 1).setFormulas(formulasF);
  sheet.getRange(2, 7, numRows, 1).setFormulas(formulasG);

  // تنسيق الأرقام
  sheet.getRange(2, 4, 100, 2).setNumberFormat('$#,##0.00'); // المخطط + الفعلي
  sheet.getRange(2, 7, 100, 1).setNumberFormat('0.0%');
  sheet.setFrozenRows(1);
}

/**
 * 🆕 تطبيق التحسينات على شيت الموازنات المخططة الموجود
 * - إضافة dropdown لأسماء المشاريع في عمود B
 * - تحويل المعادلات في عمود B إلى قيم فعلية
 * - التناغم الثنائي يعمل تلقائياً عبر onEdit
 *
 * ⚠️ هذه الدالة لا تحذف البيانات - فقط تُحدّث الإعدادات
 */
function applyBudgetsSheetEnhancements() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // البحث عن الشيت بالاسم الجديد أو القديم
  let sheet = ss.getSheetByName(CONFIG.SHEETS.BUDGETS);
  if (!sheet) {
    sheet = ss.getSheetByName('الميزانيات المخططة');
    if (sheet) {
      // إعادة تسمية الشيت للاسم الجديد
      sheet.setName(CONFIG.SHEETS.BUDGETS);
    }
  }

  if (!sheet) {
    ui.alert('⚠️ شيت الموازنات المخططة غير موجود!');
    return;
  }

  const projectsSheet = ss.getSheetByName(CONFIG.SHEETS.PROJECTS);
  if (!projectsSheet) {
    ui.alert('⚠️ شيت قاعدة بيانات المشاريع غير موجود!');
    return;
  }

  const lastRow = Math.max(sheet.getLastRow(), 2);
  const dataRows = lastRow - 1;

  // 1. تحويل المعادلات في عمود B إلى قيم فعلية (للحفاظ على البيانات)
  const colBRange = sheet.getRange(2, 2, dataRows, 1);
  const colBValues = colBRange.getValues();
  colBRange.setValues(colBValues); // هذا يحول المعادلات إلى قيم

  // 2. إضافة dropdown لكود المشروع (A)
  const projectCodeRange = projectsSheet.getRange('A2:A200');
  const projectCodeValidation = SpreadsheetApp.newDataValidation()
    .requireValueInRange(projectCodeRange, true)
    .setAllowInvalid(true)
    .setHelpText('اختر كود المشروع - سيتم ملء اسم المشروع تلقائياً')
    .build();
  sheet.getRange(2, 1, 100, 1).setDataValidation(projectCodeValidation);

  // 3. إضافة dropdown لاسم المشروع (B)
  const projectNameRange = projectsSheet.getRange('B2:B200');
  const projectNameValidation = SpreadsheetApp.newDataValidation()
    .requireValueInRange(projectNameRange, true)
    .setAllowInvalid(true)
    .setHelpText('اختر اسم المشروع - سيتم ملء كود المشروع تلقائياً')
    .build();
  sheet.getRange(2, 2, 100, 1).setDataValidation(projectNameValidation);

  ui.alert(
    '✅ تم تطبيق التحسينات بنجاح!',
    '• تم إضافة قائمة منسدلة لأسماء المشاريع (عمود B)\n' +
    '• تم تحويل المعادلات إلى قيم فعلية\n' +
    '• التناغم الثنائي يعمل الآن:\n' +
    '   - اختيار كود المشروع ← يملأ اسم المشروع\n' +
    '   - اختيار اسم المشروع ← يملأ كود المشروع',
    ui.ButtonSet.OK
  );
}


// ==================== 6. التنبيهات ====================
function createAlertsSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.SHEETS.ALERTS);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEETS.ALERTS);
  }
  sheet.clear();

  const headers = [
    'نوع التنبيه', 'الأولوية', 'المشروع', 'المورد', 'المبلغ',
    'تاريخ الاستحقاق', 'الأيام المتبقية', 'الحالة', 'الإجراء المطلوب'
  ];

  sheet.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setBackground(CONFIG.COLORS.HEADER.ALERTS)
    .setFontColor(CONFIG.COLORS.TEXT.WHITE)
    .setFontWeight('bold')
    .setFontSize(11);

  const widths = [150, 100, 180, 150, 120, 130, 120, 120, 250];
  widths.forEach((width, i) => sheet.setColumnWidth(i + 1, width));

  sheet.setFrozenRows(1);
}


// ==================== 6.5. شيت سجل النشاط ====================
/**
 * إنشاء شيت سجل النشاط لتتبع جميع العمليات
 */
function createActivityLogSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.SHEETS.ACTIVITY_LOG);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEETS.ACTIVITY_LOG);
  }
  sheet.clear();

  const headers = [
    'الوقت',              // A: تاريخ ووقت العملية
    'المستخدم',           // B: البريد الإلكتروني
    'نوع العملية',        // C: إضافة / تعديل / حذف
    'الشيت',              // D: اسم الشيت المتأثر
    'رقم الصف',           // E: رقم الصف المتأثر
    'رقم الحركة',         // F: رقم الحركة (إن وجد)
    'ملخص العملية',       // G: وصف مختصر
    'التفاصيل'            // H: تفاصيل إضافية (JSON)
  ];

  sheet.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setBackground('#37474f')
    .setFontColor(CONFIG.COLORS.TEXT.WHITE)
    .setFontWeight('bold')
    .setFontSize(11);

  const widths = [160, 200, 120, 180, 80, 100, 300, 400];
  widths.forEach((width, i) => sheet.setColumnWidth(i + 1, width));

  sheet.setFrozenRows(1);

  // تنسيق عمود الوقت
  sheet.getRange('A:A').setNumberFormat('yyyy-mm-dd hh:mm:ss');
}


/**
 * تسجيل نشاط في سجل النشاط
 * @param {string} actionType - نوع العملية (إضافة حركة، تعديل، حذف، إلخ)
 * @param {string} sheetName - اسم الشيت المتأثر
 * @param {number} rowNum - رقم الصف المتأثر
 * @param {string|number} transNum - رقم الحركة (اختياري)
 * @param {string} summary - ملخص العملية
 * @param {Object} details - تفاصيل إضافية (اختياري)
 * @param {string} userEmailParam - إيميل المستخدم من e.user (اختياري)
 */
function logActivity(actionType, sheetName, rowNum, transNum, summary, details, userEmailParam) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let logSheet = ss.getSheetByName(CONFIG.SHEETS.ACTIVITY_LOG);

    // إنشاء الشيت إذا لم يكن موجوداً
    if (!logSheet) {
      createActivityLogSheet(ss);
      logSheet = ss.getSheetByName(CONFIG.SHEETS.ACTIVITY_LOG);
    }

    // جلب البريد الإلكتروني للمستخدم الحالي
    let userEmail = '';

    // ═══════════════════════════════════════════════════════════════
    // الأولوية 1: استخدام هوية المستخدم المحفوظة (من نافذة تعريف المستخدم)
    // هذا هو الحل الموثوق - المستخدم عرّف نفسه يدوياً
    // ═══════════════════════════════════════════════════════════════
    try {
      const userProps = PropertiesService.getUserProperties();
      const savedName = userProps.getProperty('currentUserName');
      const savedEmail = userProps.getProperty('currentUserEmail');

      if (savedName) {
        // استخدام الإيميل المحفوظ أو الاسم
        userEmail = savedEmail || savedName;
      }
    } catch (pe) { /* تجاهل */ }

    // ═══════════════════════════════════════════════════════════════
    // الأولوية 2: استخدام الإيميل الممرر (من النموذج فقط - ليس من e.user)
    // نستخدمه فقط إذا لم يكن المستخدم قد عرّف نفسه
    // ═══════════════════════════════════════════════════════════════
    if (!userEmail && userEmailParam) {
      userEmail = userEmailParam;
    }

    // ═══════════════════════════════════════════════════════════════
    // الأولوية 3: محاولة من Session (للمستخدمين من نفس الدومين)
    // ═══════════════════════════════════════════════════════════════
    if (!userEmail) {
      try {
        userEmail = Session.getActiveUser().getEmail();
        if (!userEmail) {
          userEmail = Session.getEffectiveUser().getEmail();
        }
      } catch (e) { /* تجاهل */ }
    }

    // ═══════════════════════════════════════════════════════════════
    // الأولوية 4: ScriptProperties كاحتياطي أخير
    // ═══════════════════════════════════════════════════════════════
    if (!userEmail) {
      try {
        userEmail = PropertiesService.getScriptProperties().getProperty('lastUserEmail') || '';
      } catch (pe) { /* تجاهل */ }
    }

    // إذا لم نحصل على شيء
    if (!userEmail) {
      userEmail = 'غير معروف';
    }

    // حفظ الإيميل في ScriptProperties للاستخدام المستقبلي
    if (userEmail && userEmail !== 'غير معروف') {
      try {
        PropertiesService.getScriptProperties().setProperty('lastUserEmail', userEmail);
      } catch (pe) { /* تجاهل */ }
    }

    // محاولة الحصول على اسم المستخدم من شيت المستخدمين
    let displayName = userEmail;
    try {
      if (userEmail && userEmail !== 'غير معروف' && userEmail !== 'غير متاح') {
        const userInfo = getUserByEmail(userEmail);
        if (userInfo && userInfo.found && userInfo.name) {
          displayName = userInfo.name + ' (' + userEmail + ')';
        }
      }
    } catch (ue) {
      // تجاهل - نستخدم الإيميل فقط
    }

    // تحويل التفاصيل لـ JSON إذا كانت كائن
    let detailsStr = '';
    if (details) {
      try {
        detailsStr = typeof details === 'string' ? details : JSON.stringify(details, null, 0);
      } catch (e) {
        detailsStr = String(details);
      }
    }

    // إضافة السجل الجديد في الصف الثاني (بعد الهيدر) - الأحدث في الأعلى
    const timestamp = new Date();

    // إدراج صف جديد بعد الهيدر
    logSheet.insertRowAfter(1);

    // الحصول على نطاق الصف الجديد
    const newRowRange = logSheet.getRange(2, 1, 1, 8);

    // إزالة التنسيق الموروث من الهيدر
    newRowRange.clearFormat();

    // كتابة البيانات في الصف الثاني
    newRowRange.setValues([[
      timestamp,
      displayName,
      actionType,
      sheetName,
      rowNum || '',
      transNum || '',
      summary,
      detailsStr
    ]]);

    // تطبيق التنسيق الصحيح للبيانات
    newRowRange
      .setBackground('#ffffff')
      .setFontColor('#000000')
      .setFontWeight('normal')
      .setFontSize(10)
      .setVerticalAlignment('middle');

    // تنسيق عمود الوقت
    logSheet.getRange(2, 1).setNumberFormat('yyyy-mm-dd hh:mm:ss');

  } catch (e) {
    // في حالة فشل التسجيل، لا نوقف العملية الأصلية
    console.error('فشل تسجيل النشاط:', e.message);
  }
}


/**
 * عرض شيت سجل النشاط (إنشاؤه إذا لم يكن موجوداً)
 */
function showActivityLog() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let logSheet = ss.getSheetByName(CONFIG.SHEETS.ACTIVITY_LOG);

  if (!logSheet) {
    createActivityLogSheet(ss);
    logSheet = ss.getSheetByName(CONFIG.SHEETS.ACTIVITY_LOG);
  }

  ss.setActiveSheet(logSheet);
  SpreadsheetApp.getUi().alert(
    '📋 سجل النشاط',
    'تم فتح شيت سجل النشاط.\n\n' +
    'يتم تسجيل جميع العمليات تلقائياً:\n' +
    '• إضافة الحركات\n' +
    '• التعديلات\n' +
    '• العمليات الأخرى\n\n' +
    'عدد السجلات الحالي: ' + Math.max(0, logSheet.getLastRow() - 1),
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}


/**
 * إخفاء/إظهار شيت سجل النشاط
 */
function toggleActivityLogVisibility() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let logSheet = ss.getSheetByName(CONFIG.SHEETS.ACTIVITY_LOG);

  // إنشاء الشيت إذا لم يكن موجوداً
  if (!logSheet) {
    createActivityLogSheet(ss);
    logSheet = ss.getSheetByName(CONFIG.SHEETS.ACTIVITY_LOG);
    SpreadsheetApp.getUi().alert('✅ تم', 'تم إنشاء شيت سجل النشاط وهو مرئي الآن.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  // تبديل الإظهار/الإخفاء
  if (logSheet.isSheetHidden()) {
    logSheet.showSheet();
    ss.setActiveSheet(logSheet);
    SpreadsheetApp.getUi().alert('👁️ تم', 'تم إظهار شيت سجل النشاط.', SpreadsheetApp.getUi().ButtonSet.OK);
  } else {
    logSheet.hideSheet();
    SpreadsheetApp.getUi().alert('🙈 تم', 'تم إخفاء شيت سجل النشاط.', SpreadsheetApp.getUi().ButtonSet.OK);
  }
}


// ==================== 7. الوظائف اليومية الأساسية ====================

// استحقاق جديد (مدين)
function addNewExpense() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.TRANSACTIONS);

  if (!sheet) {
    ui.alert('⚠️ شيت "دفتر الحركات المالية" غير موجود!');
    return;
  }

  const lastRow = sheet.getLastRow() + 1;

  ui.alert(
    '📝 تسجيل استحقاق جديد',
    'سيتم إضافة سطر جديد في الصف ' + lastRow + '\n\n' +
    'املأ البيانات التالية:\n' +
    '• التاريخ (B)\n' +
    '• طبيعة الحركة (C)\n' +
    '• تصنيف الحركة (D)\n' +
    '• كود المشروع (E)\n' +
    '• البند (G)\n' +
    '• اسم المورد/الجهة (I)\n' +
    '• المبلغ بالعملة الأصلية (J)\n' +
    '• العملة (K)\n' +
    '• سعر الصرف (L) إن وجد\n' +
    '• نوع الحركة = "مدين استحقاق" في (N)\n' +
    '• نوع شرط الدفع (R)\n\n' +
    'القيمة بالدولار (M) والرصيد (O) تتحسب تلقائياً.',
    ui.ButtonSet.OK
  );

  sheet.setActiveRange(sheet.getRange(lastRow, 2));
}

// دفعة (دائن)
function addPayment() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.TRANSACTIONS);

  if (!sheet) {
    ui.alert('⚠️ شيت "دفتر الحركات المالية" غير موجود!');
    return;
  }

  const vendorResponse = ui.prompt(
    '💵 تسجيل دفعة',
    'أدخل اسم المورد/الجهة كما هو في العمود I:',
    ui.ButtonSet.OK_CANCEL
  );

  if (vendorResponse.getSelectedButton() !== ui.Button.OK) return;
  const vendorName = vendorResponse.getResponseText().trim();

  if (!vendorName) {
    ui.alert('⚠️ يجب إدخال اسم المورد/الجهة!');
    return;
  }

  const data = sheet.getDataRange().getValues();
  let vendorBalance = 0;
  let vendorFound = false;

  // I = index 8, O = index 14
  for (let i = 1; i < data.length; i++) {
    if (data[i][8] === vendorName) {
      vendorBalance = data[i][14];
      vendorFound = true;
    }
  }

  if (!vendorFound) {
    ui.alert('⚠️ لم يتم العثور على أي حركة للطرف: ' + vendorName);
    return;
  }

  if (vendorBalance <= 0) {
    ui.alert('✅ رصيد ' + vendorName + ' = صفر أو أقل\n\nلا توجد مستحقات مفتوحة!');
    return;
  }

  const amountResponse = ui.prompt(
    '💵 تسجيل دفعة لـ ' + vendorName,
    'الرصيد الحالي (تقريبي بالدولار): $' + vendorBalance.toLocaleString() + '\n\n' +
    'أدخل مبلغ الدفعة (بالدولار):',
    ui.ButtonSet.OK_CANCEL
  );

  if (amountResponse.getSelectedButton() !== ui.Button.OK) return;
  const amountUsd = parseFloat(amountResponse.getResponseText());

  if (isNaN(amountUsd) || amountUsd <= 0) {
    ui.alert('⚠️ مبلغ غير صحيح!');
    return;
  }

  if (amountUsd > vendorBalance) {
    ui.alert('⚠️ المبلغ المدخل أكبر من الرصيد!\n\nالرصيد: $' + vendorBalance.toLocaleString());
    return;
  }

  const paymentResponse = ui.prompt(
    '💵 تسجيل دفعة لـ ' + vendorName,
    'المبلغ: $' + amountUsd.toLocaleString() + '\n\n' +
    'اختر طريقة الدفع:\n' +
    '1 = نقدي\n' +
    '2 = تحويل بنكي\n' +
    '3 = شيك',
    ui.ButtonSet.OK_CANCEL
  );

  if (paymentResponse.getSelectedButton() !== ui.Button.OK) return;
  const paymentChoice = paymentResponse.getResponseText().trim();

  let paymentMethod;
  switch (paymentChoice) {
    case '1': paymentMethod = 'نقدي'; break;
    case '2': paymentMethod = 'تحويل بنكي'; break;
    case '3': paymentMethod = 'شيك'; break;
    default:
      ui.alert('⚠️ اختيار غير صحيح!');
      return;
  }

  const lastRow = sheet.getLastRow() + 1;
  const today = new Date();

  sheet.getRange(lastRow, 2).setValue(today);             // B التاريخ
  sheet.getRange(lastRow, 3).setValue('دفعة مصروف');  // C طبيعة الحركة
  sheet.getRange(lastRow, 4).setValue('مصروفات مباشرة'); // D
  sheet.getRange(lastRow, 9).setValue(vendorName);        // I

  sheet.getRange(lastRow, 10).setValue(amountUsd);        // J المبلغ الأصلي
  sheet.getRange(lastRow, 11).setValue('USD');            // K
  sheet.getRange(lastRow, 12).setValue(1);                // L

  sheet.getRange(lastRow, 14).setValue('دائن دفعة');     // N
  sheet.getRange(lastRow, 17).setValue(paymentMethod);    // Q
  sheet.getRange(lastRow, 24).setValue('دفعة مسجلة تلقائياً'); // X

  // ⭐ حساب الأعمدة التلقائية (M, O, V) - لأن setValue لا يُفعّل onEdit
  calculateUsdValue_(sheet, lastRow);
  recalculatePartyBalance_(sheet, lastRow);

  ui.alert(
    '✅ تم تسجيل الدفعة بنجاح!\n\n' +
    'الطرف: ' + vendorName + '\n' +
    'المبلغ (بالدولار): $' + amountUsd.toLocaleString() + '\n' +
    'الطريقة: ' + paymentMethod + '\n\n' +
    'الرصيد التقريبي الجديد (على مستوى الطرف): $' + (vendorBalance - amountUsd).toLocaleString()
  );
}

// إيراد (تحصيل من عميل/قناة)
function addRevenue() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.TRANSACTIONS);

  if (!sheet) {
    ui.alert('⚠️ شيت "دفتر الحركات المالية" غير موجود!');
    return;
  }

  const lastRow = sheet.getLastRow() + 1;

  ui.alert(
    '💰 تسجيل إيراد جديد',
    'سيتم إضافة سطر جديد في الصف ' + lastRow + '\n\n' +
    'املأ البيانات التالية:\n' +
    '• التاريخ (B)\n' +
    '• طبيعة الحركة = "تحصيل إيراد" (C)\n' +
    '• تصنيف الحركة = "تحصيل فواتير" (D)\n' +
    '• كود المشروع (E)\n' +
    '• اسم العميل/القناة في "اسم المورد/الجهة" (I)\n' +
    '• المبلغ بالعملة الأصلية (J) + العملة (K) + سعر الصرف (L)\n' +
    '• نوع الحركة = "دائن دفعة" في (N)\n\n' +
    'القيمة بالدولار (M) والرصيد (O) تتحسب تلقائياً.',
    ui.ButtonSet.OK
  );

  sheet.setActiveRange(sheet.getRange(lastRow, 2));
}

// إضافة ميزانية يدوية
function addBudgetForm() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.BUDGETS);

  if (!sheet) {
    ui.alert('⚠️ شيت "الموازنات المخططة" غير موجود!');
    return;
  }

  const lastRow = sheet.getLastRow() + 1;

  ui.alert(
    '💰 إضافة ميزانية جديدة',
    'سيتم إضافة سطر جديد في الصف ' + lastRow + '\n\n' +
    'املأ البيانات التالية:\n' +
    '• كود المشروع (A)\n' +
    '• البند (C)\n' +
    '• المبلغ المخطط (D)\n\n' +
    'الباقي سيُحسب تلقائياً!',
    ui.ButtonSet.OK
  );

  sheet.setActiveRange(sheet.getRange(lastRow, 1));
}

// مقارنة الميزانية
function compareBudget() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const response = ui.prompt(
    '📊 مقارنة الميزانية',
    'أدخل كود المشروع:',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) return;

  const projectCode = response.getResponseText().trim();
  if (!projectCode) {
    ui.alert('⚠️ يجب إدخال كود المشروع!');
    return;
  }

  const budgetSheet = ss.getSheetByName(CONFIG.SHEETS.BUDGETS);
  if (!budgetSheet) {
    ui.alert('⚠️ شيت الميزانيات غير موجود!');
    return;
  }

  const data = budgetSheet.getDataRange().getValues();
  let report = '📊 مقارنة الميزانية - ' + projectCode + '\n\n';
  let found = false;
  let totalPlanned = 0;
  let totalActual = 0;

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === projectCode) { // كود المشروع في A
      found = true;
      const item = data[i][2];                     // C
      const planned = Number(data[i][3]) || 0;     // D
      const actual = Number(data[i][4]) || 0;      // E
      const diff = Number(data[i][5]) || 0;        // F
      const percent = Number(data[i][6]) || 0;     // G (0–1)

      report += `${item}:\n`;
      report += `  المخطط: $${planned.toLocaleString()}\n`;
      report += `  الفعلي: $${actual.toLocaleString()}\n`;
      report += `  الفرق: $${diff.toLocaleString()}\n`;
      report += `  النسبة: ${(percent * 100).toFixed(1)}%\n\n`;

      totalPlanned += planned;
      totalActual += actual;
    }
  }

  if (!found) {
    ui.alert('⚠️ لم يتم العثور على ميزانية للمشروع: ' + projectCode);
    return;
  }

  report += '─────────────────────\n';
  report += `الإجمالي المخطط: $${totalPlanned.toLocaleString()}\n`;
  report += `الإجمالي الفعلي: $${totalActual.toLocaleString()}\n`;
  report += `الفرق: $${(totalPlanned - totalActual).toLocaleString()}\n`;
  report += `نسبة التنفيذ: ${((totalActual / totalPlanned) * 100).toFixed(1)}%`;

  ui.alert(report);
}


// ==================== التنبيهات والاستحقاقات (محدث: مدين + دائن + أرصدة) ====================
function updateAlerts(silent) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const transSheet = ss.getSheetByName(CONFIG.SHEETS.TRANSACTIONS);
  const alertSheet = ss.getSheetByName(CONFIG.SHEETS.ALERTS);

  if (!transSheet || !alertSheet) {
    if (!silent) {
      try { SpreadsheetApp.getUi().alert('⚠️ شيت الحركات أو التنبيهات غير موجود!'); } catch(e) {}
    }
    return { success: false, error: 'الشيتات غير موجودة' };
  }

  alertSheet.clear();

  const headers = [
    'نوع التنبيه', 'الأولوية', 'المشروع', 'الطرف', 'المبلغ (USD)',
    'تاريخ الاستحقاق', 'الأيام المتبقية', 'الحالة', 'الإجراء المطلوب'
  ];

  alertSheet.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setBackground(CONFIG.COLORS.HEADER.ALERTS)
    .setFontColor(CONFIG.COLORS.TEXT.WHITE)
    .setFontWeight('bold');

  const data = transSheet.getDataRange().getValues();
  const today = new Date();
  const alerts = [];

  // تجميع أرصدة الأطراف لتنبيهات التحصيل
  const partyBalances = {};

  for (let i = 1; i < data.length; i++) {
    const movementKind = String(data[i][13] || ''); // N: نوع الحركة (مدين استحقاق / دائن دفعة)
    const project = data[i][5];  // F: اسم المشروع
    const party = data[i][8];  // I: الطرف (مورد/عميل/ممول)
    const amountUsd = Number(data[i][12]) || 0; // M: القيمة بالدولار
    const dueDate = data[i][20]; // U: تاريخ الاستحقاق
    const status = String(data[i][21] || ''); // V: حالة السداد
    const natureType = String(data[i][2] || '');  // C: طبيعة الحركة

    // استخدام includes للتعامل مع الإيموجي
    const isDebit = movementKind.includes(CONFIG.MOVEMENT.DEBIT) || movementKind.includes('مدين');
    const isCredit = movementKind.includes(CONFIG.MOVEMENT.CREDIT) || movementKind.includes('دائن');
    const isPaid = status.includes(CONFIG.PAYMENT_STATUS.PAID) || status.includes('مدفوع');

    // تجميع أرصدة الأطراف
    if (party && amountUsd > 0) {
      if (!partyBalances[party]) {
        partyBalances[party] = { debit: 0, credit: 0, nature: natureType, project: project };
      }
      if (isDebit) {
        partyBalances[party].debit += amountUsd;
      } else if (isCredit) {
        partyBalances[party].credit += amountUsd;
      }
    }

    // ═══════════════════════════════════════════════════════════
    // 1. تنبيهات الاستحقاقات المدينة (فواتير يجب سدادها)
    // ═══════════════════════════════════════════════════════════
    if (isDebit && amountUsd > 0 && dueDate && !isPaid) {
      const dueDateObj = new Date(dueDate);
      const daysLeft = Math.ceil((dueDateObj - today) / (1000 * 60 * 60 * 24));

      let priority, alertType, action;

      if (daysLeft < 0) {
        priority = '🔴 عاجل';
        alertType = '💸 استحقاق متأخر';
        action = 'سداد فوري';
      } else if (daysLeft <= 3) {
        priority = '🟠 مهم';
        alertType = '💸 استحقاق قريب';
        action = 'تجهيز المبلغ';
      } else if (daysLeft <= 7) {
        priority = '🟡 متوسط';
        alertType = '💸 استحقاق قادم';
        action = 'متابعة';
      } else {
        continue;
      }

      alerts.push([
        alertType,
        priority,
        project,
        party,
        amountUsd,
        Utilities.formatDate(dueDateObj, Session.getScriptTimeZone(), 'dd/MM/yyyy'),
        daysLeft + ' يوم',
        status || 'معلق',
        action
      ]);
    }
  }

  // ═══════════════════════════════════════════════════════════
  // 2. تنبيهات الأرصدة المستحقة التحصيل (إيرادات لم تُحصّل)
  // ═══════════════════════════════════════════════════════════
  for (const party in partyBalances) {
    const balance = partyBalances[party].debit - partyBalances[party].credit;

    // إذا كان الرصيد موجب (على الطرف لنا فلوس) وطبيعة الحركة إيرادية
    if (balance > 100 && partyBalances[party].nature &&
      (partyBalances[party].nature.includes('إيراد') || partyBalances[party].nature.includes('تحصيل'))) {
      alerts.push([
        '💰 تحصيل مستحق',
        '🟣 متابعة',
        partyBalances[party].project || '-',
        party,
        balance,
        '-',
        '-',
        'رصيد مستحق',
        'متابعة التحصيل'
      ]);
    }
  }

  if (alerts.length > 0) {
    // ترتيب: الاستحقاقات المتأخرة أولاً
    alerts.sort((a, b) => {
      // الأولوية: عاجل > مهم > متوسط > متابعة
      const priorityOrder = { '🔴 عاجل': 1, '🟠 مهم': 2, '🟡 متوسط': 3, '🟣 متابعة': 4 };
      return (priorityOrder[a[1]] || 5) - (priorityOrder[b[1]] || 5);
    });
    alertSheet.getRange(2, 1, alerts.length, headers.length).setValues(alerts);

    // تلوين الصفوف حسب الأولوية
    for (let i = 0; i < alerts.length; i++) {
      let bgColor = '#ffffff';
      if (alerts[i][1] === '🔴 عاجل') bgColor = '#ffcdd2';
      else if (alerts[i][1] === '🟠 مهم') bgColor = '#ffe0b2';
      else if (alerts[i][1] === '🟡 متوسط') bgColor = '#fff9c4';
      else if (alerts[i][1] === '🟣 متابعة') bgColor = '#e1bee7';

      alertSheet.getRange(i + 2, 1, 1, headers.length).setBackground(bgColor);
    }
  }

  // إحصائيات
  const urgentCount = alerts.filter(a => a[1] === '🔴 عاجل').length;
  const importantCount = alerts.filter(a => a[1] === '🟠 مهم').length;
  const collectCount = alerts.filter(a => a[0] === '💰 تحصيل مستحق').length;

  SpreadsheetApp.getUi().alert(
    '✅ تم تحديث التنبيهات!\n\n' +
    '📊 الإحصائيات:\n' +
    '• 🔴 عاجل: ' + urgentCount + '\n' +
    '• 🟠 مهم: ' + importantCount + '\n' +
    '• 💰 تحصيلات مستحقة: ' + collectCount + '\n\n' +
    '📝 إجمالي التنبيهات: ' + alerts.length
  );
}

// ==================== تقرير الاستحقاقات الشامل ====================
/**
 * إنشاء تقرير استحقاقات شامل في شيت منفصل يتضمن:
 * - الاستحقاقات المدينة (فواتير يجب سدادها) - بناءً على الرصيد الفعلي
 * - الإيرادات المستحقة التحصيل
 * - ملخص حسب الفترة الزمنية
 */
function generateDueReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const transSheet = ss.getSheetByName(CONFIG.SHEETS.TRANSACTIONS);
  const ui = SpreadsheetApp.getUi();

  if (!transSheet) {
    ui.alert('⚠️ شيت دفتر الحركات غير موجود!');
    return;
  }

  const data = transSheet.getDataRange().getValues();
  const today = new Date();

  // تجميع الأرصدة حسب الطرف (مدين استحقاق - دائن دفعة)
  const partyData = {};

  for (let i = 1; i < data.length; i++) {
    const movementKind = String(data[i][13] || ''); // N - نوع الحركة
    const party = String(data[i][8] || '').trim();  // I - الطرف
    const project = String(data[i][5] || '');       // F - المشروع
    const amountUsd = Number(data[i][12]) || 0;     // M - المبلغ
    const dueDate = data[i][20];                    // U - تاريخ الاستحقاق
    const natureType = String(data[i][2] || '');    // C - طبيعة الحركة

    if (!party || amountUsd <= 0) continue;

    // تحديد نوع الحركة
    const isDebitAccrual = movementKind.indexOf('مدين استحقاق') !== -1;
    const isCreditPayment = movementKind.indexOf('دائن دفعة') !== -1;

    if (!isDebitAccrual && !isCreditPayment) continue;

    if (!partyData[party]) {
      partyData[party] = {
        totalDebit: 0,
        totalCredit: 0,
        nature: natureType,
        project: project,
        earliestDueDate: null
      };
    }

    if (isDebitAccrual) {
      partyData[party].totalDebit += amountUsd;
      // حفظ أقرب تاريخ استحقاق
      if (dueDate) {
        const dueDateObj = new Date(dueDate);
        if (!partyData[party].earliestDueDate || dueDateObj < partyData[party].earliestDueDate) {
          partyData[party].earliestDueDate = dueDateObj;
        }
      }
    } else if (isCreditPayment) {
      partyData[party].totalCredit += amountUsd;
    }
  }

  // تصنيف الأطراف حسب الرصيد المتبقي وتاريخ الاستحقاق
  const overdue = [];      // متأخرة
  const thisWeek = [];     // هذا الأسبوع
  const thisMonth = [];    // هذا الشهر
  const later = [];        // لاحقاً (لها تاريخ بعد 30 يوم)
  const noDate = [];       // بدون تاريخ استحقاق
  const receivables = [];  // إيرادات مستحقة

  let totalOverdue = 0;
  let totalThisWeek = 0;
  let totalThisMonth = 0;
  let totalLater = 0;
  let totalNoDate = 0;
  let totalReceivables = 0;

  for (const party in partyData) {
    const pd = partyData[party];
    const balance = pd.totalDebit - pd.totalCredit;

    // تجاهل الأطراف الذين رصيدهم صفر أو سالب
    if (balance <= 0.01) continue;

    // تحديد إذا كان إيراد أو مصروف
    const isRevenue = pd.nature && (pd.nature.includes('إيراد') || pd.nature.includes('تحصيل'));

    if (isRevenue) {
      // إيرادات مستحقة التحصيل
      receivables.push({ party, amount: balance, project: pd.project, daysLeft: 0 });
      totalReceivables += balance;
    } else {
      // مستحقات علينا - تصنيف حسب تاريخ الاستحقاق
      const item = { party, project: pd.project, amount: balance, dueDate: pd.earliestDueDate, daysLeft: null };

      if (!pd.earliestDueDate) {
        // بدون تاريخ استحقاق
        noDate.push(item);
        totalNoDate += balance;
      } else {
        const daysLeft = Math.ceil((pd.earliestDueDate - today) / (1000 * 60 * 60 * 24));
        item.daysLeft = daysLeft;

        if (daysLeft < 0) {
          overdue.push(item);
          totalOverdue += balance;
        } else if (daysLeft <= 7) {
          thisWeek.push(item);
          totalThisWeek += balance;
        } else if (daysLeft <= 30) {
          thisMonth.push(item);
          totalThisMonth += balance;
        } else {
          later.push(item);
          totalLater += balance;
        }
      }
    }
  }

  // إنشاء أو إعادة استخدام الشيت
  const reportSheetName = 'تقرير الاستحقاقات';
  let reportSheet = ss.getSheetByName(reportSheetName);

  if (reportSheet) {
    reportSheet.clear();
  } else {
    reportSheet = ss.insertSheet(reportSheetName);
  }

  // === بناء التقرير في الشيت ===
  let currentRow = 1;
  const numCols = 5;

  // العنوان الرئيسي
  reportSheet.getRange(currentRow, 1, 1, numCols).merge();
  reportSheet.getRange(currentRow, 1)
    .setValue('📊 تقرير الاستحقاقات الشامل')
    .setFontSize(16)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setBackground('#4a86e8')
    .setFontColor('white');
  currentRow++;

  // تاريخ التقرير
  reportSheet.getRange(currentRow, 1, 1, numCols).merge();
  reportSheet.getRange(currentRow, 1)
    .setValue('📅 تاريخ التقرير: ' + Utilities.formatDate(today, 'Asia/Riyadh', 'yyyy-MM-dd HH:mm'))
    .setFontSize(10)
    .setHorizontalAlignment('center')
    .setBackground('#cfe2f3');
  currentRow += 2;

  // دالة مساعدة لإضافة قسم
  function addSection(title, items, total, bgColor, textColor, showDays) {
    // عنوان القسم
    reportSheet.getRange(currentRow, 1, 1, numCols).merge();
    reportSheet.getRange(currentRow, 1)
      .setValue(title + ' (' + items.length + ')')
      .setFontWeight('bold')
      .setFontSize(12)
      .setBackground(bgColor)
      .setFontColor(textColor);
    currentRow++;

    // هيدر الجدول
    const headers = showDays ? ['#', 'الطرف', 'المشروع', 'المبلغ ($)', 'الأيام'] : ['#', 'الطرف', 'المشروع', 'المبلغ ($)', ''];
    reportSheet.getRange(currentRow, 1, 1, numCols).setValues([headers]);
    reportSheet.getRange(currentRow, 1, 1, numCols)
      .setBackground('#e0e0e0')
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
    currentRow++;

    // البيانات
    if (items.length > 0) {
      items.sort((a, b) => showDays ? a.daysLeft - b.daysLeft : b.amount - a.amount);
      for (let i = 0; i < items.length; i++) {
        const item = items[i];
        const daysText = showDays ? (item.daysLeft < 0 ? 'متأخر ' + Math.abs(item.daysLeft) : item.daysLeft + ' يوم') : '';
        reportSheet.getRange(currentRow, 1, 1, numCols).setValues([[
          i + 1,
          item.party || '',
          item.project || '',
          item.amount,
          daysText
        ]]);
        if (i % 2 === 1) {
          reportSheet.getRange(currentRow, 1, 1, numCols).setBackground('#f5f5f5');
        }
        // تلوين الأيام المتأخرة
        if (showDays && item.daysLeft < 0) {
          reportSheet.getRange(currentRow, 5).setFontColor('#cc0000').setFontWeight('bold');
        }
        currentRow++;
      }
    } else {
      reportSheet.getRange(currentRow, 1, 1, numCols).merge();
      reportSheet.getRange(currentRow, 1)
        .setValue('✅ لا يوجد')
        .setHorizontalAlignment('center')
        .setFontColor('#2e7d32');
      currentRow++;
    }

    // إجمالي القسم
    reportSheet.getRange(currentRow, 1, 1, 3).merge();
    reportSheet.getRange(currentRow, 1)
      .setValue('💰 الإجمالي:')
      .setFontWeight('bold')
      .setHorizontalAlignment('left')
      .setBackground('#e8eaf6');
    reportSheet.getRange(currentRow, 4)
      .setValue(total)
      .setNumberFormat('$#,##0.00')
      .setFontWeight('bold')
      .setBackground('#e8eaf6');
    reportSheet.getRange(currentRow, 5).setBackground('#e8eaf6');
    currentRow += 2;
  }

  // 1. الاستحقاقات المتأخرة
  addSection('🔴 الاستحقاقات المتأخرة', overdue, totalOverdue, '#ffcdd2', '#b71c1c', true);

  // 2. استحقاقات هذا الأسبوع
  addSection('🟠 استحقاقات هذا الأسبوع', thisWeek, totalThisWeek, '#ffe0b2', '#e65100', true);

  // 3. استحقاقات هذا الشهر
  addSection('🟡 استحقاقات هذا الشهر', thisMonth, totalThisMonth, '#fff9c4', '#f57f17', true);

  // 4. استحقاقات لاحقة (لها تاريخ بعد 30 يوم)
  addSection('🟢 استحقاقات لاحقة', later, totalLater, '#c8e6c9', '#2e7d32', true);

  // 5. استحقاقات بدون تاريخ
  addSection('⚪ بدون تاريخ استحقاق', noDate, totalNoDate, '#e0e0e0', '#424242', false);

  // 6. الإيرادات المستحقة
  addSection('💰 إيرادات مستحقة التحصيل', receivables, totalReceivables, '#bbdefb', '#0d47a1', false);

  // === الملخص المالي ===
  currentRow++;
  reportSheet.getRange(currentRow, 1, 1, numCols).merge();
  reportSheet.getRange(currentRow, 1)
    .setValue('📈 الملخص المالي')
    .setFontWeight('bold')
    .setFontSize(14)
    .setBackground('#4a86e8')
    .setFontColor('white')
    .setHorizontalAlignment('center');
  currentRow++;

  const totalPayables = totalOverdue + totalThisWeek + totalThisMonth + totalLater + totalNoDate;
  const netPosition = totalReceivables - totalPayables;

  // إجمالي الاستحقاقات
  reportSheet.getRange(currentRow, 1, 1, 3).merge();
  reportSheet.getRange(currentRow, 1)
    .setValue('💸 إجمالي الاستحقاقات (علينا):')
    .setBackground('#ffcdd2');
  reportSheet.getRange(currentRow, 4, 1, 2).merge();
  reportSheet.getRange(currentRow, 4)
    .setValue(totalPayables)
    .setNumberFormat('$#,##0.00')
    .setFontWeight('bold')
    .setBackground('#ffcdd2')
    .setFontColor('#b71c1c');
  currentRow++;

  // إجمالي التحصيلات
  reportSheet.getRange(currentRow, 1, 1, 3).merge();
  reportSheet.getRange(currentRow, 1)
    .setValue('💰 إجمالي التحصيلات (لنا):')
    .setBackground('#c8e6c9');
  reportSheet.getRange(currentRow, 4, 1, 2).merge();
  reportSheet.getRange(currentRow, 4)
    .setValue(totalReceivables)
    .setNumberFormat('$#,##0.00')
    .setFontWeight('bold')
    .setBackground('#c8e6c9')
    .setFontColor('#2e7d32');
  currentRow++;

  // صافي الموقف
  reportSheet.getRange(currentRow, 1, 1, 3).merge();
  reportSheet.getRange(currentRow, 1)
    .setValue('📊 صافي الموقف:')
    .setFontWeight('bold')
    .setFontSize(12)
    .setBackground('#fff9c4');
  reportSheet.getRange(currentRow, 4, 1, 2).merge();
  reportSheet.getRange(currentRow, 4)
    .setValue(netPosition)
    .setNumberFormat('$#,##0.00')
    .setFontWeight('bold')
    .setFontSize(12)
    .setBackground('#fff9c4')
    .setFontColor(netPosition >= 0 ? '#2e7d32' : '#b71c1c');

  // تنسيق الأعمدة
  reportSheet.setColumnWidth(1, 40);   // #
  reportSheet.setColumnWidth(2, 180);  // الطرف
  reportSheet.setColumnWidth(3, 150);  // المشروع
  reportSheet.setColumnWidth(4, 120);  // المبلغ
  reportSheet.setColumnWidth(5, 100);  // الأيام

  // تجميد الصفوف العلوية
  reportSheet.setFrozenRows(2);

  // الانتقال للشيت
  ss.setActiveSheet(reportSheet);

  ui.alert('✅ تم إنشاء تقرير الاستحقاقات',
    'الملخص:\n\n' +
    '• متأخرة: $' + totalOverdue.toFixed(2) + ' (' + overdue.length + ')\n' +
    '• هذا الأسبوع: $' + totalThisWeek.toFixed(2) + ' (' + thisWeek.length + ')\n' +
    '• هذا الشهر: $' + totalThisMonth.toFixed(2) + ' (' + thisMonth.length + ')\n' +
    '• لاحقاً: $' + totalLater.toFixed(2) + ' (' + later.length + ')\n' +
    '• تحصيلات: $' + totalReceivables.toFixed(2) + ' (' + receivables.length + ')\n\n' +
    '📊 صافي الموقف: $' + netPosition.toFixed(2),
    ui.ButtonSet.OK);
}

// ==================== نافذة الاستحقاقات القادمة (30 يوم) ====================
function showUpcomingPayments() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const transSheet = ss.getSheetByName(CONFIG.SHEETS.TRANSACTIONS);

  if (!transSheet) {
    SpreadsheetApp.getUi().alert('⚠️ شيت دفتر الحركات غير موجود!');
    return;
  }

  const today = new Date();
  const next30Days = new Date(today.getTime() + 30 * 24 * 60 * 60 * 1000);

  const transData = transSheet.getDataRange().getValues();
  let upcomingPayments = [];

  for (let i = 1; i < transData.length; i++) {
    const movementKind = String(transData[i][13] || '');  // N: نوع الحركة
    const status = String(transData[i][21] || '');  // V: حالة السداد
    const dueDate = transData[i][20];  // U: تاريخ الاستحقاق
    const balance = Number(transData[i][14]) || 0; // O: الرصيد (بالدولار على مستوى الطرف)
    const party = transData[i][8];   // I: الطرف
    const project = transData[i][5];   // F: اسم المشروع

    // استخدام includes للتعامل مع الإيموجي
    const isDebit = movementKind.includes(CONFIG.MOVEMENT.DEBIT) || movementKind.includes('مدين');
    const isPaid = status.includes(CONFIG.PAYMENT_STATUS.PAID) || status.includes('مدفوع');
    if (isDebit && balance > 0 && dueDate && !isPaid) {
      const dueDateObj = new Date(dueDate);
      if (dueDateObj <= next30Days) {
        const daysLeft = Math.ceil((dueDateObj - today) / (1000 * 60 * 60 * 24));
        upcomingPayments.push({
          party: party,
          project: project,
          amount: balance, // رصيد الطرف بالدولار
          dueDate: Utilities.formatDate(dueDateObj, Session.getScriptTimeZone(), 'dd/MM/yyyy'),
          daysLeft: daysLeft,
          status: status
        });
      }
    }
  }

  upcomingPayments.sort((a, b) => a.daysLeft - b.daysLeft);

  let message = '🔔 الاستحقاقات خلال الـ 30 يوم القادمة:\n\n';

  if (upcomingPayments.length === 0) {
    message += '✅ لا توجد استحقاقات خلال الفترة القادمة';
  } else {
    let total = 0;
    upcomingPayments.forEach(payment => {
      const statusIcon = payment.daysLeft < 0
        ? '🔴 متأخر'
        : payment.daysLeft <= 3
          ? '🟠 عاجل'
          : '🟢 قريب';
      message += `${statusIcon} ${payment.party} - ${payment.project}\n`;
      message += `   المبلغ (USD): $${payment.amount.toLocaleString()} | التاريخ: ${payment.dueDate} | متبقي: ${payment.daysLeft} يوم\n\n`;
      total += payment.amount;
    });
    message += `\n💰 إجمالي المستحقات (تقريباً بالدولار): $${total.toLocaleString()}`;
  }

  SpreadsheetApp.getUi().alert(message);
}


// ==================== تقرير مورد تفصيلي (محدث بالعملات) ====================
function generateVendorDetailedReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const response = ui.prompt(
    '👤 تقرير طرف تفصيلي',
    'أدخل اسم الطرف (مورد/عميل/ممول) بالضبط كما في دفتر الحركات:',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) return;

  const partyName = response.getResponseText().trim();
  if (!partyName) {
    ui.alert('⚠️ يجب إدخال اسم الطرف!');
    return;
  }

  const transSheet = ss.getSheetByName(CONFIG.SHEETS.TRANSACTIONS);
  if (!transSheet) {
    ui.alert('⚠️ شيت الحركات غير موجود!');
    return;
  }

  const data = transSheet.getDataRange().getValues();
  const rows = [];

  for (let i = 1; i < data.length; i++) {
    if (data[i][8] === partyName) {  // I: اسم الطرف
      rows.push({
        date: data[i][1],   // B: التاريخ
        movementType: data[i][2],   // C: طبيعة الحركة
        classification: data[i][3],   // D: تصنيف الحركة
        project: data[i][5],   // F: اسم المشروع
        item: data[i][6],   // G: البند
        details: data[i][7],   // H: التفاصيل
        amountOriginal: Number(data[i][9]) || 0,  // J: المبلغ الأصلي
        currency: data[i][10] || '',        // K: العملة
        rate: Number(data[i][11]) || 0, // L: سعر الصرف
        amountUsd: Number(data[i][12]) || 0, // M: القيمة بالدولار
        movementKind: data[i][13],             // N: نوع الحركة
        balance: Number(data[i][14]) || 0, // O: الرصيد
        refNum: data[i][15],             // P: رقم مرجعي
        notes: data[i][23]              // X: ملاحظات
      });
    }
  }

  if (rows.length === 0) {
    ui.alert('⚠️ لم يتم العثور على حركات للطرف: ' + partyName);
    return;
  }

  // ترتيب زمني
  rows.sort((a, b) => new Date(a.date) - new Date(b.date));

  let totalDebitUsd = 0;
  let totalCreditUsd = 0;
  let paymentCount = 0;

  rows.forEach(row => {
    // استخدام includes للتعامل مع الإيموجي
    const kindStr = String(row.movementKind || '');
    if (kindStr.includes(CONFIG.MOVEMENT.DEBIT) || kindStr.includes('مدين')) {
      totalDebitUsd += row.amountUsd;
    } else if (kindStr.includes(CONFIG.MOVEMENT.CREDIT) || kindStr.includes('دائن')) {
      totalCreditUsd += row.amountUsd;
      if (row.amountUsd > 0) paymentCount++;
    }
  });

  const currentBalanceCalc = totalDebitUsd - totalCreditUsd;
  const lastBalance = rows[rows.length - 1].balance || currentBalanceCalc;

  let report = `📊 تقرير تفصيلي - ${partyName}\n`;
  report += '═'.repeat(50) + '\n\n';

  report += '💰 ملخص الحساب (بالدولار):\n';
  report += `• إجمالي الاستحقاقات (مدين استحقاق): $${totalDebitUsd.toLocaleString()}\n`;
  report += `• إجمالي الدفعات (دائن دفعة): $${totalCreditUsd.toLocaleString()}\n`;
  report += `• الرصيد الحالي التقريبي: $${lastBalance.toLocaleString()}\n`;
  report += `• عدد الدفعات: ${paymentCount}\n\n`;

  report += '📋 كشف الحساب التفصيلي:\n';
  report += '─'.repeat(50) + '\n';

  rows.forEach(row => {
    const dateStr = row.date
      ? Utilities.formatDate(new Date(row.date), Session.getScriptTimeZone(), 'dd/MM/yyyy')
      : '';

    report += `\n📅 ${dateStr} | ${row.movementType} (${row.classification})\n`;
    report += `   المشروع: ${row.project || '-'} - ${row.item || '-'}\n`;

    if (row.details) {
      report += `   التفاصيل: ${row.details}\n`;
    }

    // تنسيق مبلغ أصلي + بالدولار
    let originalPart = '';
    if (row.amountOriginal) {
      originalPart = `${row.amountOriginal.toLocaleString()} ${row.currency || ''}`.trim();
    }
    const usdPart = row.amountUsd ? `$${row.amountUsd.toLocaleString()}` : '';
    let amountText = usdPart;
    if (originalPart && usdPart) {
      amountText = `${originalPart} ≈ ${usdPart}`;
    } else if (originalPart) {
      amountText = originalPart;
    }

    // استخدام includes للتعامل مع الإيموجي
    const kindStr2 = String(row.movementKind || '');
    if (kindStr2.includes(CONFIG.MOVEMENT.DEBIT) || kindStr2.includes('مدين')) {
      report += `   مدين (استحقاق): ${amountText}\n`;
    } else if (kindStr2.includes(CONFIG.MOVEMENT.CREDIT) || kindStr2.includes('دائن')) {
      report += `   دائن (دفعة/تحصيل): ${amountText}\n`;
    }

    report += `   الرصيد (USD): $${row.balance.toLocaleString()}\n`;

    if (row.refNum) {
      report += `   رقم مرجعي: ${row.refNum}\n`;
    }
    if (row.notes) {
      report += `   📝 ${row.notes}\n`;
    }
  });

  report += '\n' + '═'.repeat(50) + '\n';
  report += `🔚 نهاية التقرير - الرصيد النهائي (تقريبي): $${lastBalance.toLocaleString()}`;

  ui.alert(report);
}


// ==================== كشف حساب طرف مختصر (محدث) ====================
function showVendorStatement() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const response = ui.prompt(
    '📋 كشف حساب طرف',
    'أدخل اسم الطرف (مورد/عميل/ممول):',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) return;

  const partyName = response.getResponseText().trim();
  if (!partyName) {
    ui.alert('⚠️ يجب إدخال اسم الطرف!');
    return;
  }

  const transSheet = ss.getSheetByName(CONFIG.SHEETS.TRANSACTIONS);
  if (!transSheet) {
    ui.alert('⚠️ شيت الحركات غير موجود!');
    return;
  }

  const data = transSheet.getDataRange().getValues();
  const rows = [];

  for (let i = 1; i < data.length; i++) {
    if (data[i][8] === partyName) { // I: الطرف
      rows.push({
        date: data[i][1],              // B
        movementType: data[i][2],              // C
        movementKind: data[i][13],             // N
        amountUsd: Number(data[i][12]) || 0,// M
        balance: Number(data[i][14]) || 0 // O
      });
    }
  }

  if (rows.length === 0) {
    ui.alert('⚠️ لم يتم العثور على حركات للطرف: ' + partyName);
    return;
  }

  rows.sort((a, b) => new Date(a.date) - new Date(b.date));

  let statement = `📋 كشف حساب: ${partyName}\n`;
  statement += '═'.repeat(40) + '\n\n';

  let currentBalance = 0;

  rows.forEach(row => {
    const dateStr = row.date
      ? Utilities.formatDate(new Date(row.date), Session.getScriptTimeZone(), 'dd/MM/yyyy')
      : '';

    statement += `${dateStr} | ${row.movementType}\n`;

    // استخدام includes للتعامل مع الإيموجي
    const kindStr = String(row.movementKind || '');
    if (kindStr.includes(CONFIG.MOVEMENT.DEBIT) || kindStr.includes('مدين')) {
      statement += `         مدين (استحقاق): $${row.amountUsd.toLocaleString()}\n`;
    } else if (kindStr.includes(CONFIG.MOVEMENT.CREDIT) || kindStr.includes('دائن')) {
      statement += `         دائن (دفعة/تحصيل): $${row.amountUsd.toLocaleString()}\n`;
    }

    currentBalance = row.balance;
    statement += `         رصيد (USD): $${row.balance.toLocaleString()}\n\n`;
  });

  statement += '═'.repeat(40) + '\n';
  statement += `الرصيد الحالي (تقريبي بالدولار): $${currentBalance.toLocaleString()}`;

  ui.alert(statement);
}


// ==================== تقرير ربحية المشروع (محدث باستخدام القيمة بالدولار) ====================
function showProjectProfitability() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const response = ui.prompt(
    '💹 تقرير ربحية مشروع',
    'أدخل كود المشروع:',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) return;

  const projectCode = response.getResponseText().trim().toUpperCase();
  if (!projectCode) {
    ui.alert('⚠️ يجب إدخال كود المشروع!');
    return;
  }

  const projectsSheet = ss.getSheetByName(CONFIG.SHEETS.PROJECTS);
  const transSheet = ss.getSheetByName(CONFIG.SHEETS.TRANSACTIONS);

  if (!projectsSheet || !transSheet) {
    ui.alert('⚠️ الشيتات المطلوبة غير موجودة!');
    return;
  }

  const projectsData = projectsSheet.getDataRange().getValues();
  let projectInfo = null;

  // كود المشروع (A)، اسم المشروع (B)، القناة (D)، نوع التمويل (G)، قيمة التمويل (H)، قيمة العقد (I)
  for (let i = 1; i < projectsData.length; i++) {
    if (String(projectsData[i][0]).trim().toUpperCase() === projectCode) {
      projectInfo = {
        code: projectsData[i][0],
        name: projectsData[i][1],
        channel: projectsData[i][3] || '',
        fundingType: projectsData[i][6],
        fundingValue: Number(projectsData[i][7]) || 0,
        contractValue: Number(projectsData[i][8]) || 0
      };
      break;
    }
  }

  if (!projectInfo) {
    ui.alert('⚠️ لم يتم العثور على المشروع: ' + projectCode);
    return;
  }

  const transData = transSheet.getDataRange().getValues();
  let directExpenses = 0;  // مصروفات مباشرة (مدين استحقاق)
  let revenues = 0;        // إيرادات (قيمة العقد أو التحصيل)

  for (let i = 1; i < transData.length; i++) {
    const rowProjCode = String(transData[i][4] || '').trim().toUpperCase();
    if (rowProjCode !== projectCode) continue;

    const movementType = String(transData[i][2] || '');  // C: طبيعة الحركة
    const classification = String(transData[i][3] || '');  // D: تصنيف الحركة
    const movementKind = String(transData[i][13] || ''); // N: نوع الحركة
    const amountUsd = Number(transData[i][12]) || 0;  // M: القيمة بالدولار

    const isDebit = movementKind.includes('مدين');

    // مصروفات مباشرة (استحقاق فقط)
    if (isDebit && amountUsd > 0) {
      directExpenses += amountUsd;
    }
  }

  // الإيرادات = قيمة العقد مع القناة
  revenues = projectInfo.contractValue;

  // ═══════════════════════════════════════════════════════════
  // الحسابات الجديدة
  // ═══════════════════════════════════════════════════════════
  const profitMargin = revenues - directExpenses;                    // هامش الربح
  const overheadExpenses = directExpenses * 0.35;                    // مصروفات عمومية 35%
  const netProfit = profitMargin - overheadExpenses;                 // صافي الربح

  const profitMarginPercent = revenues > 0 ? (profitMargin / revenues) * 100 : 0;
  const netProfitPercent = revenues > 0 ? (netProfit / revenues) * 100 : 0;

  // ═══════════════════════════════════════════════════════════
  // بناء التقرير
  // ═══════════════════════════════════════════════════════════
  let report = '═'.repeat(45) + '\n';
  report += `💹 تقرير ربحية المشروع: ${projectCode}\n`;
  report += `${projectInfo.name}\n`;
  report += '═'.repeat(45) + '\n\n';

  report += '📊 البيانات الأساسية:\n';
  report += `• القناة/الجهة: ${projectInfo.channel}\n`;
  report += `• قيمة العقد: $${revenues.toLocaleString()}\n`;
  report += `• نوع التمويل: ${projectInfo.fundingType || '-'}\n`;
  report += `• قيمة التمويل: $${projectInfo.fundingValue.toLocaleString()}\n`;
  report += '─'.repeat(45) + '\n\n';

  report += '💰 المصروفات:\n';
  report += `• المصروفات المباشرة: $${directExpenses.toLocaleString()}\n`;
  report += '─'.repeat(45) + '\n\n';

  const marginIcon = profitMargin >= 0 ? '✅' : '❌';
  report += `${marginIcon} هامش الربح: $${profitMargin.toLocaleString()} (${profitMarginPercent.toFixed(1)}%)\n`;
  report += '─'.repeat(45) + '\n\n';

  report += '🏢 المصروفات العمومية:\n';
  report += `• 35% من المصروفات المباشرة: $${overheadExpenses.toLocaleString()}\n`;
  report += '─'.repeat(45) + '\n\n';

  const netIcon = netProfit >= 0 ? '✅' : '❌';
  report += `${netIcon} صافي الربح: $${netProfit.toLocaleString()} (${netProfitPercent.toFixed(1)}%)\n`;
  report += '═'.repeat(45) + '\n';

  // ═══════════════════════════════════════════════════════════
  // عرض النافذة مع خيار إصدار شيت
  // ═══════════════════════════════════════════════════════════
  const alertResponse = ui.alert(
    report,
    'هل تريد إصدار التقرير في شيت منفصل؟',
    ui.ButtonSet.YES_NO
  );

  if (alertResponse === ui.Button.YES) {
    // إنشاء شيت التقرير
    createProfitabilityReportSheet_(ss, projectInfo, directExpenses, revenues, profitMargin, overheadExpenses, netProfit, profitMarginPercent, netProfitPercent);
  }
}

/**
 * إنشاء شيت تقرير الربحية
 */
function createProfitabilityReportSheet_(ss, projectInfo, directExpenses, revenues, profitMargin, overheadExpenses, netProfit, profitMarginPercent, netProfitPercent) {
  const reportSheetName = 'تقرير ربحية - ' + projectInfo.code;
  let reportSheet = ss.getSheetByName(reportSheetName);
  if (reportSheet) {
    ss.deleteSheet(reportSheet);
  }
  reportSheet = ss.insertSheet(reportSheetName);
  reportSheet.setRightToLeft(true);

  let row = 1;

  // العنوان
  reportSheet.getRange(row, 1, 1, 4).merge()
    .setValue('💹 تقرير ربحية المشروع: ' + projectInfo.code)
    .setBackground('#1a237e')
    .setFontColor('white')
    .setFontWeight('bold')
    .setFontSize(16)
    .setHorizontalAlignment('center');
  row++;

  reportSheet.getRange(row, 1, 1, 4).merge()
    .setValue(projectInfo.name)
    .setBackground('#283593')
    .setFontColor('white')
    .setFontSize(14)
    .setHorizontalAlignment('center');
  row += 2;

  // بيانات المشروع
  reportSheet.getRange(row, 1, 1, 4).merge()
    .setValue('📊 بيانات المشروع')
    .setBackground('#e8eaf6')
    .setFontWeight('bold');
  row++;

  const projectData = [
    ['القناة/الجهة', projectInfo.channel, 'قيمة العقد', revenues],
    ['نوع التمويل', projectInfo.fundingType || '-', 'قيمة التمويل', projectInfo.fundingValue]
  ];
  reportSheet.getRange(row, 1, 2, 4).setValues(projectData);
  reportSheet.getRange(row, 4, 2, 1).setNumberFormat('$#,##0.00');
  row += 3;

  // جدول الربحية
  reportSheet.getRange(row, 1, 1, 4).merge()
    .setValue('📈 تحليل الربحية')
    .setBackground('#e8eaf6')
    .setFontWeight('bold');
  row++;

  const headers = ['البند', 'المبلغ ($)', 'النسبة %', 'الحالة'];
  reportSheet.getRange(row, 1, 1, 4).setValues([headers])
    .setBackground('#3949ab')
    .setFontColor('white')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  row++;

  const profitData = [
    ['إجمالي الإيرادات (قيمة العقد)', revenues, '100%', ''],
    ['المصروفات المباشرة', directExpenses, ((directExpenses / revenues) * 100).toFixed(1) + '%', ''],
    ['', '', '', ''],
    ['هامش الربح', profitMargin, profitMarginPercent.toFixed(1) + '%', profitMargin >= 0 ? '✅' : '❌'],
    ['', '', '', ''],
    ['مصروفات عمومية (35%)', overheadExpenses, '35%', ''],
    ['', '', '', ''],
    ['صافي الربح', netProfit, netProfitPercent.toFixed(1) + '%', netProfit >= 0 ? '✅' : '❌']
  ];

  reportSheet.getRange(row, 1, profitData.length, 4).setValues(profitData);
  reportSheet.getRange(row, 2, profitData.length, 1).setNumberFormat('$#,##0.00');

  // تلوين هامش الربح
  const marginRow = row + 3;
  reportSheet.getRange(marginRow, 1, 1, 4)
    .setBackground(profitMargin >= 0 ? '#e8f5e9' : '#ffebee')
    .setFontWeight('bold');

  // تلوين صافي الربح
  const netRow = row + 7;
  reportSheet.getRange(netRow, 1, 1, 4)
    .setBackground(netProfit >= 0 ? '#c8e6c9' : '#ffcdd2')
    .setFontWeight('bold')
    .setFontSize(12);

  row += profitData.length + 2;

  // تاريخ التقرير
  reportSheet.getRange(row, 1, 1, 4).merge()
    .setValue('تاريخ التقرير: ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm'))
    .setFontSize(9)
    .setFontColor('#666666')
    .setHorizontalAlignment('center');

  // تنسيقات
  reportSheet.setColumnWidth(1, 200);
  reportSheet.setColumnWidth(2, 150);
  reportSheet.setColumnWidth(3, 100);
  reportSheet.setColumnWidth(4, 80);
  reportSheet.setFrozenRows(2);

  ss.setActiveSheet(reportSheet);
  ss.toast('✅ تم إنشاء شيت تقرير الربحية', 'نجاح', 3);
}

// ==================== تقرير ربحية كل المشاريع الشامل ====================

/**
 * إنشاء تقرير ربحية شامل لكل المشاريع
 * يعرض كل مشروع مع بنوده وهامش الربح والمصروفات العمومية وصافي الربح
 * @param {boolean} silent - إذا كان true لن يظهر رسائل للمستخدم (للاستخدام من البوت)
 */
function generateAllProjectsProfitabilityReport(silent) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // دالة مساعدة لعرض الرسائل (تتجاهل في الوضع الصامت)
  const showAlert = (msg) => {
    if (!silent) {
      try { SpreadsheetApp.getUi().alert(msg); } catch(e) {}
    }
  };

  const projectsSheet = ss.getSheetByName(CONFIG.SHEETS.PROJECTS);
  const transSheet = ss.getSheetByName(CONFIG.SHEETS.TRANSACTIONS);
  const budgetSheet = ss.getSheetByName(CONFIG.SHEETS.BUDGETS);

  if (!projectsSheet || !transSheet) {
    showAlert('⚠️ الشيتات المطلوبة غير موجودة!');
    return { success: false, error: 'الشيتات غير موجودة' };
  }

  // قراءة بيانات المشاريع
  const projectsData = projectsSheet.getDataRange().getValues();
  if (projectsData.length < 2) {
    showAlert('⚠️ لا توجد مشاريع في قاعدة البيانات');
    return { success: false, error: 'لا توجد مشاريع' };
  }

  // قراءة الميزانيات المخططة
  const budgetData = budgetSheet ? budgetSheet.getDataRange().getValues() : [];

  // قراءة الحركات
  const transData = transSheet.getDataRange().getValues();

  // إنشاء شيت التقرير
  const reportSheetName = 'تقارير ربحية المشاريع';
  let reportSheet = ss.getSheetByName(reportSheetName);
  if (reportSheet) {
    ss.deleteSheet(reportSheet);
  }
  reportSheet = ss.insertSheet(reportSheetName);
  reportSheet.setRightToLeft(true);

  let currentRow = 1;

  // العنوان الرئيسي
  reportSheet.getRange(currentRow, 1, 1, 7).merge()
    .setValue('📊 تقارير ربحية المشاريع الشاملة')
    .setBackground('#1a237e')
    .setFontColor('white')
    .setFontWeight('bold')
    .setFontSize(18)
    .setHorizontalAlignment('center');
  currentRow++;

  reportSheet.getRange(currentRow, 1, 1, 7).merge()
    .setValue('تاريخ التقرير: ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm'))
    .setBackground('#283593')
    .setFontColor('white')
    .setFontSize(11)
    .setHorizontalAlignment('center');
  currentRow += 2;

  // متغيرات للإجماليات
  let totalContracts = 0;
  let totalDirectExpenses = 0;
  let totalProfitMargin = 0;
  let totalOverhead = 0;
  let totalNetProfit = 0;
  let projectCount = 0;

  // معالجة كل مشروع
  for (let p = 1; p < projectsData.length; p++) {
    const projectCode = String(projectsData[p][0] || '').trim();
    const projectName = String(projectsData[p][1] || '').trim();
    const channel = String(projectsData[p][3] || '').trim();
    const contractValue = Number(projectsData[p][8]) || 0;

    if (!projectCode || contractValue === 0) continue;

    projectCount++;

    // جمع الميزانية المخططة للمشروع
    const plannedBudget = {};
    let totalPlanned = 0;
    for (let b = 1; b < budgetData.length; b++) {
      const budgetProjCode = String(budgetData[b][0] || '').trim().toUpperCase();
      if (budgetProjCode === projectCode.toUpperCase()) {
        const item = String(budgetData[b][2] || '').trim();
        const amount = Number(budgetData[b][3]) || 0;
        if (item) {
          plannedBudget[item] = (plannedBudget[item] || 0) + amount;
          totalPlanned += amount;
        }
      }
    }

    // جمع المصروفات الفعلية للمشروع (استحقاق مصروف فقط - استبعاد الإيرادات)
    const actualExpenses = {};
    let totalActual = 0;
    for (let t = 1; t < transData.length; t++) {
      const rowProjCode = String(transData[t][4] || '').trim().toUpperCase();
      if (rowProjCode !== projectCode.toUpperCase()) continue;

      const item = String(transData[t][6] || '').trim();
      const amountUsd = Number(transData[t][12]) || 0;
      const natureType = String(transData[t][2] || ''); // طبيعة الحركة (column C)

      // فقط استحقاق مصروف - استبعاد الإيرادات والدفعات
      if (natureType.includes('استحقاق مصروف') && amountUsd > 0) {
        if (!item) continue;
        actualExpenses[item] = (actualExpenses[item] || 0) + amountUsd;
        totalActual += amountUsd;
      }
    }

    // حسابات الربحية
    const profitMargin = contractValue - totalActual;
    const overheadExpenses = totalActual * 0.35;
    const netProfit = profitMargin - overheadExpenses;
    const profitMarginPercent = contractValue > 0 ? (profitMargin / contractValue) * 100 : 0;
    const netProfitPercent = contractValue > 0 ? (netProfit / contractValue) * 100 : 0;

    // تحديث الإجماليات
    totalContracts += contractValue;
    totalDirectExpenses += totalActual;
    totalProfitMargin += profitMargin;
    totalOverhead += overheadExpenses;
    totalNetProfit += netProfit;

    // ═══════════════════════════════════════════════════════════
    // عنوان المشروع
    // ═══════════════════════════════════════════════════════════
    reportSheet.getRange(currentRow, 1, 1, 7).merge()
      .setValue('🎬 المشروع: ' + projectCode + ' - ' + projectName)
      .setBackground('#3949ab')
      .setFontColor('white')
      .setFontWeight('bold')
      .setFontSize(12);
    currentRow++;

    reportSheet.getRange(currentRow, 1, 1, 7).merge()
      .setValue('القناة: ' + channel + ' | قيمة العقد: $' + contractValue.toLocaleString())
      .setBackground('#5c6bc0')
      .setFontColor('white')
      .setFontSize(10);
    currentRow++;

    // رؤوس أعمدة البنود
    const itemHeaders = ['البند', 'المخطط', 'الفعلي', 'الفرق', 'النسبة %', '', ''];
    reportSheet.getRange(currentRow, 1, 1, 7).setValues([itemHeaders])
      .setBackground('#e8eaf6')
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
    currentRow++;

    // جمع كل البنود
    const allItems = new Set([...Object.keys(plannedBudget), ...Object.keys(actualExpenses)]);
    const itemRows = [];

    allItems.forEach(item => {
      if (item.includes('عمولة مدير')) return;
      const planned = plannedBudget[item] || 0;
      const actual = actualExpenses[item] || 0;
      const diff = planned - actual;
      const percentage = planned > 0 ? Math.round((actual / planned) * 100) : (actual > 0 ? 999 : 0);
      itemRows.push([item, planned, actual, diff, percentage + '%', '', '']);
    });

    // ترتيب حسب الفعلي تنازلياً
    itemRows.sort((a, b) => b[2] - a[2]);

    if (itemRows.length > 0) {
      reportSheet.getRange(currentRow, 1, itemRows.length, 7).setValues(itemRows);
      reportSheet.getRange(currentRow, 2, itemRows.length, 3).setNumberFormat('$#,##0.00');

      // تلوين الفرق
      for (let i = 0; i < itemRows.length; i++) {
        const diffValue = itemRows[i][3];
        if (diffValue < 0) {
          reportSheet.getRange(currentRow + i, 4).setFontColor('#c62828');
        } else if (diffValue > 0) {
          reportSheet.getRange(currentRow + i, 4).setFontColor('#2e7d32');
        }
      }
      currentRow += itemRows.length;
    }

    // صف إجمالي المصروفات
    reportSheet.getRange(currentRow, 1, 1, 7)
      .setValues([['إجمالي المصروفات المباشرة', totalPlanned, totalActual, totalPlanned - totalActual,
        totalPlanned > 0 ? Math.round((totalActual / totalPlanned) * 100) + '%' : '-', '', '']])
      .setBackground('#e0e0e0')
      .setFontWeight('bold');
    reportSheet.getRange(currentRow, 2, 1, 3).setNumberFormat('$#,##0.00');
    currentRow++;

    // هامش الربح
    const marginIcon = profitMargin >= 0 ? '✅' : '❌';
    reportSheet.getRange(currentRow, 1, 1, 7).merge()
      .setValue(marginIcon + ' هامش الربح: $' + profitMargin.toLocaleString() + ' (' + profitMarginPercent.toFixed(1) + '%)')
      .setBackground(profitMargin >= 0 ? '#e8f5e9' : '#ffebee')
      .setFontWeight('bold')
      .setFontSize(11);
    currentRow++;

    // مصروفات عمومية
    reportSheet.getRange(currentRow, 1, 1, 7).merge()
      .setValue('🏢 مصروفات عمومية (35%): $' + overheadExpenses.toLocaleString())
      .setBackground('#fff3e0')
      .setFontSize(10);
    currentRow++;

    // صافي الربح
    const netIcon = netProfit >= 0 ? '✅' : '❌';
    reportSheet.getRange(currentRow, 1, 1, 7).merge()
      .setValue(netIcon + ' صافي الربح: $' + netProfit.toLocaleString() + ' (' + netProfitPercent.toFixed(1) + '%)')
      .setBackground(netProfit >= 0 ? '#c8e6c9' : '#ffcdd2')
      .setFontWeight('bold')
      .setFontSize(12);
    currentRow += 2;
  }

  // ═══════════════════════════════════════════════════════════
  // الملخص الإجمالي
  // ═══════════════════════════════════════════════════════════
  reportSheet.getRange(currentRow, 1, 1, 7).merge()
    .setValue('📊 الملخص الإجمالي لكل المشاريع')
    .setBackground('#1a237e')
    .setFontColor('white')
    .setFontWeight('bold')
    .setFontSize(14)
    .setHorizontalAlignment('center');
  currentRow++;

  const totalProfitMarginPercent = totalContracts > 0 ? (totalProfitMargin / totalContracts) * 100 : 0;
  const totalNetProfitPercent = totalContracts > 0 ? (totalNetProfit / totalContracts) * 100 : 0;

  const summaryData = [
    ['عدد المشاريع', projectCount, '', '', '', '', ''],
    ['إجمالي قيمة العقود', totalContracts, '', '', '', '', ''],
    ['إجمالي المصروفات المباشرة', totalDirectExpenses, '', '', '', '', ''],
    ['', '', '', '', '', '', ''],
    ['✅ إجمالي هامش الربح', totalProfitMargin, totalProfitMarginPercent.toFixed(1) + '%', '', '', '', ''],
    ['🏢 إجمالي المصروفات العمومية (35%)', totalOverhead, '35%', '', '', '', ''],
    ['', '', '', '', '', '', ''],
    ['💰 إجمالي صافي الربح', totalNetProfit, totalNetProfitPercent.toFixed(1) + '%', '', '', '', '']
  ];

  reportSheet.getRange(currentRow, 1, summaryData.length, 7).setValues(summaryData);
  reportSheet.getRange(currentRow + 1, 2, 2, 1).setNumberFormat('$#,##0.00');
  reportSheet.getRange(currentRow + 4, 2, 1, 1).setNumberFormat('$#,##0.00');
  reportSheet.getRange(currentRow + 5, 2, 1, 1).setNumberFormat('$#,##0.00');
  reportSheet.getRange(currentRow + 7, 2, 1, 1).setNumberFormat('$#,##0.00');

  // تنسيق صف هامش الربح الإجمالي
  reportSheet.getRange(currentRow + 4, 1, 1, 7)
    .setBackground(totalProfitMargin >= 0 ? '#e8f5e9' : '#ffebee')
    .setFontWeight('bold');

  // تنسيق صف المصروفات العمومية
  reportSheet.getRange(currentRow + 5, 1, 1, 7)
    .setBackground('#fff3e0');

  // تنسيق صف صافي الربح الإجمالي
  reportSheet.getRange(currentRow + 7, 1, 1, 7)
    .setBackground(totalNetProfit >= 0 ? '#c8e6c9' : '#ffcdd2')
    .setFontWeight('bold')
    .setFontSize(13);

  // تنسيقات عامة
  reportSheet.setColumnWidth(1, 220);
  reportSheet.setColumnWidth(2, 120);
  reportSheet.setColumnWidth(3, 120);
  reportSheet.setColumnWidth(4, 120);
  reportSheet.setColumnWidth(5, 100);
  reportSheet.setColumnWidth(6, 80);
  reportSheet.setColumnWidth(7, 80);
  reportSheet.setFrozenRows(2);

  if (!silent) {
    ss.setActiveSheet(reportSheet);
  }

  showAlert(
    '✅ تم إنشاء تقرير ربحية كل المشاريع\n\n' +
    '📁 عدد المشاريع: ' + projectCount + '\n' +
    '💵 إجمالي العقود: $' + totalContracts.toLocaleString() + '\n' +
    '💰 إجمالي المصروفات: $' + totalDirectExpenses.toLocaleString() + '\n\n' +
    '✅ إجمالي هامش الربح: $' + totalProfitMargin.toLocaleString() + '\n' +
    '🏢 إجمالي المصروفات العمومية: $' + totalOverhead.toLocaleString() + '\n' +
    '💰 إجمالي صافي الربح: $' + totalNetProfit.toLocaleString()
  );

  return { success: true, projectCount: projectCount };
}

// ==================== دليل الاستخدام (محدث لنظام العملات + نوع الحركة) ====================
function showGuide() {
  const ui = SpreadsheetApp.getUi();

  ui.alert(
    '📖 دليل الاستخدام - نظام العملات + نوع الحركة',
    '1️⃣ دفتر الحركات المالية:\n' +
    '   • J: المبلغ بالعملة الأصلية (TRY / USD / EGP ...)\n' +
    '   • K: العملة\n' +
    '   • L: سعر الصرف إلى دولار (لو فضلت فاضي = نفس العملة USD)\n' +
    '   • M: القيمة بالدولار (تحسب تلقائياً = J × L)\n' +
    '   • N: نوع الحركة = "مدين استحقاق" أو "دائن دفعة"\n' +
    '   • O: الرصيد بالدولار على مستوى الطرف (مجموع المدين - الدائن)\n\n' +
    '2️⃣ طبيعة الحركة (C) وتصنيف الحركة (D):\n' +
    '   • طبيعة الحركة: مثل استحقاق مصروف / دفعة مصروف / استحقاق إيراد / تحصيل إيراد\n' +
    '   • تصنيف الحركة: مصروفات مباشرة / مصروفات عمومية / تحصيل فواتير / ...\n\n' +
    '3️⃣ حالة السداد (V):\n' +
    '   • "معلق"      = استحقاق لم يُغلق بالكامل\n' +
    '   • "مدفوع بالكامل" = لا يوجد رصيد مفتوح على الطرف\n' +
    '   • "عملية دفع/تحصيل" = سطر دفعة/تحصيل فقط\n\n' +
    '4️⃣ التقارير:\n' +
    '   • تقرير ربحية المشروع يعتمد على القيمة بالدولار (M)\n' +
    '   • كشف حساب الطرف يوضح الاستحقاقات والدفعات بالدولار مع المحافظة على بيانات العملة الأصلية في الدفتر\n' +
    '   • التنبيهات والاستحقاقات تعتمد على نوع الحركة "مدين استحقاق" وتاريخ الاستحقاق (U).\n\n' +
    '5️⃣ قاعدة بيانات الأطراف:\n' +
    '   • شيت "قاعدة بيانات الأطراف" يحتوي على الموردين والعملاء والممولين في مكان واحد، والربط يتم من عمود "اسم المورد/الجهة" في دفتر الحركات.',
    ui.ButtonSet.OK
  );
}


// ==================== تحديث القوائم المنسدلة (موافق للهيكل الجديد) ====================
function refreshDropdowns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const transSheet = ss.getSheetByName(CONFIG.SHEETS.TRANSACTIONS);
  const projectsSheet = ss.getSheetByName(CONFIG.SHEETS.PROJECTS);
  const partiesSheet = ss.getSheetByName(CONFIG.SHEETS.PARTIES);
  const itemsSheet = ss.getSheetByName(CONFIG.SHEETS.ITEMS);
  const budgetSheet = ss.getSheetByName(CONFIG.SHEETS.BUDGETS);

  if (!transSheet) {
    ui.alert('⚠️ شيت "دفتر الحركات المالية" غير موجود!');
    return;
  }

  const lastRow = 500;

  // كود المشروع في دفتر الحركات (E)
  if (projectsSheet) {
    const projectRange = projectsSheet.getRange('A2:A200');
    const projectValidation = SpreadsheetApp.newDataValidation()
      .requireValueInRange(projectRange, true)
      .setAllowInvalid(true)
      .setHelpText('اختر كود المشروع من القائمة أو اكتب يدوياً')
      .build();
    transSheet.getRange(2, 5, lastRow, 1).setDataValidation(projectValidation); // E

    // 🆕 اسم المشروع في دفتر الحركات (F)
    // أولاً: تحويل المعادلات القديمة إلى قيم نصية (إزالة VLOOKUP formulas)
    const colF = transSheet.getRange(2, 6, lastRow, 1);
    const colFValues = colF.getValues();
    colF.setValues(colFValues); // تحويل المعادلات إلى قيم

    // ثانياً: إضافة الـ dropdown
    const projectNameRange = projectsSheet.getRange('B2:B200');
    const projectNameValidation = SpreadsheetApp.newDataValidation()
      .requireValueInRange(projectNameRange, true)
      .setAllowInvalid(true)
      .setHelpText('اختر اسم المشروع - سيتم ملء كود المشروع تلقائياً')
      .build();
    colF.setDataValidation(projectNameValidation); // F
  }

  // اسم الطرف (مورد/عميل/ممول) في دفتر الحركات (I)
  if (partiesSheet) {
    const partyRange = partiesSheet.getRange('A2:A500');
    const partyValidation = SpreadsheetApp.newDataValidation()
      .requireValueInRange(partyRange, true)
      .setAllowInvalid(true)
      .setHelpText('اختر اسم الطرف من "قاعدة بيانات الأطراف" أو اكتب يدوياً')
      .build();
    transSheet.getRange(2, 9, lastRow, 1).setDataValidation(partyValidation); // I
  }

  // البنود + طبيعة الحركة + تصنيف الحركة من "قاعدة بيانات البنود"
  if (itemsSheet) {
    const lastItemsRow = Math.max(itemsSheet.getLastRow() - 1, 1);

    // البند (G) من عمود A
    const itemsRange = itemsSheet.getRange(2, 1, lastItemsRow, 1); // A2:A
    const itemValidation = SpreadsheetApp.newDataValidation()
      .requireValueInRange(itemsRange, true)
      .setAllowInvalid(true)
      .setHelpText('اختر البند من "قاعدة بيانات البنود" أو اكتب يدوياً')
      .build();
    transSheet.getRange(2, 7, lastRow, 1).setDataValidation(itemValidation); // G

    // طبيعة الحركة (C) من عمود B
    const movementRange = itemsSheet.getRange(2, 2, lastItemsRow, 1); // B2:B
    const movementValidation = SpreadsheetApp.newDataValidation()
      .requireValueInRange(movementRange, true)
      .setAllowInvalid(true)
      .setHelpText('اختر طبيعة الحركة من "قاعدة بيانات البنود" (عمود B)')
      .build();
    transSheet.getRange(2, 3, lastRow, 1).setDataValidation(movementValidation); // C

    // تصنيف الحركة (D) من عمود C
    const classRange = itemsSheet.getRange(2, 3, lastItemsRow, 1); // C2:C
    const classValidation = SpreadsheetApp.newDataValidation()
      .requireValueInRange(classRange, true)
      .setAllowInvalid(true)
      .setHelpText('اختر تصنيف الحركة من "قاعدة بيانات البنود" (عمود C)')
      .build();
    transSheet.getRange(2, 4, lastRow, 1).setDataValidation(classValidation); // D
  }

  // كود المشروع في شيت الميزانيات (A) + البند (C)
  if (budgetSheet && projectsSheet) {
    const projectRange = projectsSheet.getRange('A2:A200');
    const projectValidation = SpreadsheetApp.newDataValidation()
      .requireValueInRange(projectRange, true)
      .setAllowInvalid(true)
      .setHelpText('اختر كود المشروع من القائمة أو اكتب يدوياً')
      .build();
    budgetSheet.getRange(2, 1, 100, 1).setDataValidation(projectValidation); // A
  }

  if (budgetSheet && itemsSheet) {
    const lastItemsRow = Math.max(itemsSheet.getLastRow() - 1, 1);
    const itemsRange = itemsSheet.getRange(2, 1, lastItemsRow, 1); // A2:A
    const itemValidation = SpreadsheetApp.newDataValidation()
      .requireValueInRange(itemsRange, true)
      .setAllowInvalid(true)
      .setHelpText('اختر البند من "قاعدة بيانات البنود"')
      .build();
    budgetSheet.getRange(2, 3, 100, 1).setDataValidation(itemValidation); // C
  }

  ui.alert(
    '✅ تم تحديث القوائم المنسدلة!\n\n' +
    '• كود المشروع في دفتر الحركات والميزانيات\n' +
    '• اسم الطرف من "قاعدة بيانات الأطراف"\n' +
    '• البنود + طبيعة الحركة + تصنيف الحركة من "قاعدة بيانات البنود"'
  );
}

/**
 * تنظيف الايقونات من عمود طبيعة الحركة (C) في البيانات الموجودة
 */
function cleanupNatureTypeEmojis() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.TRANSACTIONS);

  if (!sheet) {
    ui.alert('⚠️ شيت دفتر الحركات المالية غير موجود!');
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    ui.alert('لا توجد بيانات للتحديث');
    return;
  }

  // قراءة عمود C (طبيعة الحركة)
  const range = sheet.getRange(2, 3, lastRow - 1, 1);
  const values = range.getValues();

  // خريطة الاستبدال
  const emojiMap = {
    '💰 استحقاق مصروف': 'استحقاق مصروف',
    '💸 دفعة مصروف': 'دفعة مصروف',
    '📈 استحقاق إيراد': 'استحقاق إيراد',
    '✅ تحصيل إيراد': 'تحصيل إيراد',
    '🏦 تمويل': 'تمويل',
    '💳 سداد تمويل': 'سداد تمويل'
  };

  let updatedCount = 0;

  for (let i = 0; i < values.length; i++) {
    const oldValue = values[i][0];
    if (oldValue && emojiMap[oldValue]) {
      values[i][0] = emojiMap[oldValue];
      updatedCount++;
    }
  }

  if (updatedCount > 0) {
    range.setValues(values);
    ui.alert('✅ تم تحديث ' + updatedCount + ' خلية في عمود طبيعة الحركة');
  } else {
    ui.alert('لا توجد خلايا تحتاج تحديث');
  }
}

/**
 * تطبيع التواريخ في جميع الشيتات
 * - دفتر الحركات المالية: أعمدة B و T
 * - قاعدة بيانات المشاريع: أعمدة J و K
 * تحويل النصوص إلى Date objects وضبط التنسيق إلى dd/MM/yyyy
 */
function normalizeDateColumns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // تأكيد من المستخدم
  const response = ui.alert(
    '📅 تطبيع التواريخ',
    'سيتم تحويل جميع التواريخ إلى صيغة موحدة (dd/MM/yyyy)\n\n' +
    'الشيتات المشمولة:\n' +
    '• دفتر الحركات المالية: أعمدة B و T\n' +
    '• قاعدة بيانات المشاريع: أعمدة J و K\n\n' +
    'هذا سيصلح:\n' +
    '• التواريخ المكتوبة كنصوص\n' +
    '• التواريخ بفواصل مختلفة (/ . -)\n\n' +
    'هل تريد المتابعة؟',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  let results = [];

  // ═══════════════════════════════════════════════════════════
  // 1. دفتر الحركات المالية
  // ═══════════════════════════════════════════════════════════
  const transSheet = ss.getSheetByName(CONFIG.SHEETS.TRANSACTIONS);
  if (transSheet && transSheet.getLastRow() > 1) {
    const lastRow = transSheet.getLastRow();

    // عمود B (التاريخ)
    const updatedB = normalizeColumnDates_(transSheet, 2, lastRow);

    // عمود T (تاريخ مخصص)
    const updatedT = normalizeColumnDates_(transSheet, 20, lastRow);

    results.push('دفتر الحركات المالية:');
    results.push('  • عمود B (التاريخ): ' + updatedB + ' خلية');
    results.push('  • عمود T (تاريخ مخصص): ' + updatedT + ' خلية');
  }

  // ═══════════════════════════════════════════════════════════
  // 2. قاعدة بيانات المشاريع
  // ═══════════════════════════════════════════════════════════
  const projSheet = ss.getSheetByName(CONFIG.SHEETS.PROJECTS);
  if (projSheet && projSheet.getLastRow() > 1) {
    const lastRow = projSheet.getLastRow();

    // عمود J (تاريخ البدء)
    const updatedJ = normalizeColumnDates_(projSheet, 10, lastRow);

    // عمود K (تاريخ التسليم المتوقع)
    const updatedK = normalizeColumnDates_(projSheet, 11, lastRow);

    results.push('');
    results.push('قاعدة بيانات المشاريع:');
    results.push('  • عمود J (تاريخ البدء): ' + updatedJ + ' خلية');
    results.push('  • عمود K (تاريخ التسليم المتوقع): ' + updatedK + ' خلية');
  }

  ui.alert(
    '✅ تم تطبيع التواريخ!',
    results.join('\n') + '\n\nتم ضبط تنسيق جميع الأعمدة إلى dd/MM/yyyy',
    ui.ButtonSet.OK
  );
}

/**
 * تطبيع عمود تاريخ معين
 * @param {Sheet} sheet - الشيت
 * @param {number} col - رقم العمود
 * @param {number} lastRow - آخر صف
 * @returns {number} عدد الخلايا المحدثة
 */
function normalizeColumnDates_(sheet, col, lastRow) {
  const range = sheet.getRange(2, col, lastRow - 1, 1);
  const values = range.getValues();
  let updated = 0;

  for (let i = 0; i < values.length; i++) {
    const val = values[i][0];
    if (!val) continue;
    if (val instanceof Date) continue;

    if (typeof val === 'string') {
      const parseResult = parseDateInput_(val.trim());
      if (parseResult.success) {
        values[i][0] = parseResult.dateObj;
        updated++;
      }
    }
  }

  if (updated > 0) {
    range.setValues(values);
  }
  range.setNumberFormat('dd/mm/yyyy');

  return updated;
}

/**
 * إصلاح جميع الـ dropdowns في دفتر الحركات المالية
 * يُستخدم لإعادة تطبيق القوائم المنسدلة على الأعمدة
 */
function fixAllDropdowns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.TRANSACTIONS);

  if (!sheet) {
    ui.alert('⚠️ شيت دفتر الحركات المالية غير موجود!');
    return;
  }

  const lastRow = Math.max(sheet.getLastRow(), 500);

  // نوع الحركة (N = 14)
  const movementValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(CONFIG.MOVEMENT.TYPES, true)
    .setAllowInvalid(true)
    .build();
  sheet.getRange(2, 14, lastRow, 1).setDataValidation(movementValidation);

  // العملة (K = 11)
  const currencyValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(CONFIG.CURRENCIES.LIST, true)
    .setAllowInvalid(true)
    .build();
  sheet.getRange(2, 11, lastRow, 1).setDataValidation(currencyValidation);

  // طريقة الدفع (Q = 17)
  const payMethodValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(['نقدي', 'تحويل بنكي', 'شيك', 'بطاقة', 'أخرى'])
    .setAllowInvalid(true)
    .build();
  sheet.getRange(2, 17, lastRow, 1).setDataValidation(payMethodValidation);

  // نوع شرط الدفع (R = 18)
  const termValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(CONFIG.PAYMENT_TERMS.LIST)
    .setAllowInvalid(true)
    .build();
  sheet.getRange(2, 18, lastRow, 1).setDataValidation(termValidation);

  // عدد الأسابيع (S = 19) - validation للأرقام فقط 0-52
  const weeksValidation = SpreadsheetApp.newDataValidation()
    .requireNumberBetween(0, 52)
    .setAllowInvalid(false)
    .setHelpText('أدخل عدد الأسابيع (0-52) - يُستخدم مع شرط "بعد التسليم"')
    .build();
  sheet.getRange(2, 19, lastRow, 1).setDataValidation(weeksValidation);

  // تعيين القيمة الافتراضية 0 للخلايا الفارغة في عمود S
  const weeksRange = sheet.getRange(2, 19, lastRow, 1);
  const weeksValues = weeksRange.getValues();
  let fixedCount = 0;
  for (let i = 0; i < weeksValues.length; i++) {
    if (weeksValues[i][0] === '' || weeksValues[i][0] === null) {
      weeksValues[i][0] = 0;
      fixedCount++;
    }
  }
  if (fixedCount > 0) {
    weeksRange.setValues(weeksValues);
  }

  ui.alert(
    '✅ تم إصلاح القوائم المنسدلة!',
    'تم تطبيق الـ dropdowns والـ validations على:\n\n' +
    '• عمود N (نوع الحركة)\n' +
    '• عمود K (العملة)\n' +
    '• عمود Q (طريقة الدفع)\n' +
    '• عمود R (نوع شرط الدفع)\n' +
    '• عمود S (عدد الأسابيع) - أرقام 0-52\n\n' +
    'تم تصحيح ' + fixedCount + ' خلية فارغة في عمود S\n' +
    'عدد الصفوف: ' + lastRow,
    ui.ButtonSet.OK
  );

  // ═══════════════════════════════════════════════════════════
  // إصلاح قاعدة بيانات الأطراف
  // ═══════════════════════════════════════════════════════════
  const partiesSheet = ss.getSheetByName(CONFIG.SHEETS.PARTIES);
  if (partiesSheet) {
    const partiesLastRow = Math.max(partiesSheet.getLastRow(), 500);

    // نوع الطرف (B = 2) - مورد / عميل / ممول
    const partyTypeValidation = SpreadsheetApp.newDataValidation()
      .requireValueInList(['مورد', 'عميل', 'ممول'], true)
      .setAllowInvalid(true)
      .build();
    partiesSheet.getRange(2, 2, partiesLastRow, 1).setDataValidation(partyTypeValidation);

    // طريقة الدفع المفضلة (G = 7)
    const partyPayValidation = SpreadsheetApp.newDataValidation()
      .requireValueInList(['نقدي', 'تحويل بنكي', 'شيك', 'بطاقة', 'أخرى'], true)
      .setAllowInvalid(true)
      .build();
    partiesSheet.getRange(2, 7, partiesLastRow, 1).setDataValidation(partyPayValidation);

    ui.alert('✅ تم أيضاً إصلاح قاعدة بيانات الأطراف (مورد/عميل/ممول)');
  }
}


// ==================== إضافة عمود كشف الحساب للشيت الحالي ====================
/**
 * إضافة عمود "📄 كشف" (Y) لدفتر الحركات الحالي
 * يضيف العمود ويملأه بالرمز 📄 لكل صف فيه بيانات
 */
function addStatementLinkColumn() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.TRANSACTIONS);

  if (!sheet) {
    ui.alert('❌ خطأ', 'شيت دفتر الحركات المالية غير موجود!', ui.ButtonSet.OK);
    return;
  }

  // التحقق من وجود العمود مسبقاً
  const currentHeader = sheet.getRange(1, 25).getValue();
  if (currentHeader === '📄 كشف') {
    // العمود موجود، نسأل المستخدم إذا يريد إعادة ملء الرموز
    const response = ui.alert(
      '📄 عمود موجود',
      'عمود "📄 كشف" موجود بالفعل.\n\nهل تريد إعادة ملء الرموز 📄 في جميع الصفوف؟',
      ui.ButtonSet.YES_NO
    );
    if (response !== ui.Button.YES) return;
  } else {
    // إضافة العنوان
    sheet.getRange(1, 25)
      .setValue('📄 كشف')
      .setBackground('#4a86e8')
      .setFontColor('white')
      .setFontWeight('bold')
      .setHorizontalAlignment('center');

    // تعيين عرض العمود
    sheet.setColumnWidth(25, 60);
  }

  // ملء العمود بالرمز 📄 لكل صف فيه بيانات
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    ui.alert('✅ تم', 'تم إضافة عمود "📄 كشف".\n\nلا توجد بيانات لملء الرموز.', ui.ButtonSet.OK);
    return;
  }

  // قراءة عمود التاريخ (B) لمعرفة الصفوف التي فيها بيانات
  const dates = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
  const icons = [];

  for (let i = 0; i < dates.length; i++) {
    // إذا كان هناك تاريخ، نضع الرمز
    if (dates[i][0]) {
      icons.push(['📄']);
    } else {
      icons.push(['']);
    }
  }

  // كتابة الرموز دفعة واحدة
  sheet.getRange(2, 25, lastRow - 1, 1).setValues(icons);

  // تنسيق العمود
  sheet.getRange(2, 25, lastRow - 1, 1)
    .setHorizontalAlignment('center')
    .setFontSize(12);

  // إحصائية
  const filledCount = icons.filter(row => row[0] === '📄').length;

  ui.alert(
    '✅ تم بنجاح',
    'تم إضافة عمود "📄 كشف" (Y) لدفتر الحركات.\n\n' +
    '• عدد الصفوف التي تم ملؤها: ' + filledCount + '\n\n' +
    '📌 طريقة الاستخدام:\n' +
    'اضغط على خلية 📄 في أي صف → سيتم إنشاء كشف حساب للطرف تلقائياً',
    ui.ButtonSet.OK
  );
}

// ==================== إضافة عمود عدد الوحدات ====================
/**
 * إضافة عمود "عدد الوحدات" (AA - العمود 27) لدفتر الحركات
 * يُستخدم لحساب تكلفة الوحدة في التقارير
 */
function addUnitCountColumn() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.TRANSACTIONS);

  if (!sheet) {
    ui.alert('❌ خطأ', 'شيت دفتر الحركات المالية غير موجود!', ui.ButtonSet.OK);
    return;
  }

  // التحقق من وجود العمود مسبقاً
  const currentHeader = sheet.getRange(1, 27).getValue();
  if (currentHeader === 'عدد الوحدات') {
    ui.alert(
      '✅ موجود',
      'عمود "عدد الوحدات" موجود بالفعل في العمود AA.',
      ui.ButtonSet.OK
    );
    return;
  }

  // التحقق من أن العمود 26 هو "رقم الأوردر" (للتأكد من الترتيب الصحيح)
  const col26Header = sheet.getRange(1, 26).getValue();
  if (col26Header !== 'رقم الأوردر') {
    // ربما هيكل الشيت مختلف، نحذر المستخدم
    const response = ui.alert(
      '⚠️ تحذير',
      'العمود Z ليس "رقم الأوردر".\n' +
      'القيمة الحالية: "' + col26Header + '"\n\n' +
      'هل تريد المتابعة وإضافة عمود "عدد الوحدات" في العمود AA؟',
      ui.ButtonSet.YES_NO
    );
    if (response !== ui.Button.YES) return;
  }

  // إضافة العنوان
  sheet.getRange(1, 27)
    .setValue('عدد الوحدات')
    .setBackground(CONFIG.COLORS.HEADER.TRANSACTIONS)
    .setFontColor(CONFIG.COLORS.TEXT.WHITE)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setFontSize(11);

  // تعيين عرض العمود
  sheet.setColumnWidth(27, 100);

  // تنسيق خلايا البيانات
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 27, lastRow - 1, 1)
      .setHorizontalAlignment('center')
      .setNumberFormat('#,##0');
  }

  ui.alert(
    '✅ تم بنجاح',
    'تم إضافة عمود "عدد الوحدات" (AA) لدفتر الحركات.\n\n' +
    '📌 الاستخدام:\n' +
    '• عند إضافة حركة جديدة لبنود الإنتاج (تصوير، مونتاج، إلخ)،\n' +
    '  سيظهر حقل "عدد الوحدات" في النموذج.\n\n' +
    '• البنود المدعومة:\n' +
    Object.entries(CONFIG.UNIT_TYPES).map(([item, unit]) =>
      '  - ' + item + ' ← ' + unit
    ).join('\n'),
    ui.ButtonSet.OK
  );
}

// ==================== إضافة عمود كشف الحساب لتقرير الموردين الموجود ====================
/**
 * إضافة عمود "📄 كشف" لتقرير الموردين الموجود
 * يسمح بإنشاء كشف حساب للمورد بضغطة واحدة
 */
function addStatementColumnToVendorReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.VENDORS_REPORT);

  if (!sheet) {
    ui.alert('❌ خطأ', 'شيت "تقرير الموردين" غير موجود!', ui.ButtonSet.OK);
    return;
  }

  // التحقق من وجود العمود مسبقاً
  const currentHeader = sheet.getRange(1, 10).getValue();
  if (currentHeader === '📄 كشف') {
    // العمود موجود، نسأل المستخدم إذا يريد إعادة ملء الرموز
    const response = ui.alert(
      '📄 عمود موجود',
      'عمود "📄 كشف" موجود بالفعل.\n\nهل تريد إعادة ملء الرموز 📄 في جميع الصفوف؟',
      ui.ButtonSet.YES_NO
    );
    if (response !== ui.Button.YES) return;
  } else {
    // إضافة العنوان
    sheet.getRange(1, 10)
      .setValue('📄 كشف')
      .setBackground(CONFIG.COLORS.HEADER.VENDORS)
      .setFontColor('white')
      .setFontWeight('bold')
      .setHorizontalAlignment('center');

    // تعيين عرض العمود
    sheet.setColumnWidth(10, 60);

    // إضافة ملاحظة توضيحية
    sheet.getRange(1, 10).setNote(
      '📄 اضغط على أي خلية في هذا العمود لإنشاء كشف حساب للمورد'
    );
  }

  // ملء العمود بالرمز 📄 لكل صف فيه بيانات
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    ui.alert('✅ تم', 'تم إضافة عمود "📄 كشف".\n\nلا توجد بيانات لملء الرموز.', ui.ButtonSet.OK);
    return;
  }

  // قراءة عمود اسم المورد (A) لمعرفة الصفوف التي فيها بيانات
  const vendors = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  const icons = [];

  for (let i = 0; i < vendors.length; i++) {
    // إذا كان هناك اسم مورد، نضع الرمز
    if (vendors[i][0]) {
      icons.push(['📄']);
    } else {
      icons.push(['']);
    }
  }

  // كتابة الرموز دفعة واحدة
  sheet.getRange(2, 10, lastRow - 1, 1).setValues(icons);

  // تنسيق العمود
  sheet.getRange(2, 10, lastRow - 1, 1)
    .setHorizontalAlignment('center')
    .setFontSize(12);

  // إحصائية
  const filledCount = icons.filter(row => row[0] === '📄').length;

  ui.alert(
    '✅ تم بنجاح',
    'تم إضافة عمود "📄 كشف" (J) لتقرير الموردين.\n\n' +
    '• عدد الصفوف التي تم ملؤها: ' + filledCount + '\n\n' +
    '📌 طريقة الاستخدام:\n' +
    'اضغط على خلية 📄 في أي صف → سيتم إنشاء كشف حساب للمورد تلقائياً',
    ui.ButtonSet.OK
  );
}

// ==================== تصحيح عنوان عمود الملاحظات (مواكب للهيكل الجديد) ====================
function patchRenameNotesColumn() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.TRANSACTIONS);
  if (!sheet) return;
  // العمود 24 هو عمود الملاحظات (X) في الهيكل الجديد
  sheet.getRange(1, 24).setValue('ملاحظات');
}

// ==================== التقارير - الجزء 2 ====================
function setupPart2() {
  if (!confirmReset()) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  createProjectReportSheet(ss);
  createVendorReportSheet(ss);
  createFunderReportSheet(ss);
  createExpenseReportSheet(ss);
  createRevenueReportSheet(ss);
  createCashFlowSheet(ss);
  createDashboardSheet(ss);
  createInvoiceTemplateSheet(ss);   // 🆕 نموذج فاتورة القناة

  SpreadsheetApp.getUi().alert(
    '✅ تم إنشاء الجزء 2 بنجاح!\n\n' +
    'التقارير المتاحة:\n' +
    '• تقرير المشروع التفصيلي\n' +
    '• تقرير الموردين الملخص\n' +
    '• تقرير المصروفات\n' +
    '• تقرير الإيرادات\n' +
    '• التدفقات النقدية\n' +
    '• لوحة التحكم\n' +
    '• 🧾 نموذج فاتورة قناة\n\n' +
    '🎉 النظام جاهز!'
  );
}

// ==================== نموذج الفاتورة الإنجليزي ====================
function createInvoiceTemplateSheet(ss) {
  // نشتغل على نفس التاب اللي عندك في الصورة
  let sheet = ss.getSheetByName(CONFIG.SHEETS.INVOICE) || ss.getSheetByName('Invoice');
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEETS.INVOICE);
  }

  // نفرّغ ونبني من جديد
  sheet.clear();
  sheet.setRightToLeft(false);

  // عرض الأعمدة
  sheet.setColumnWidth(1, 220); // A
  sheet.setColumnWidth(2, 160); // B
  sheet.setColumnWidth(3, 120); // C
  sheet.setColumnWidth(4, 140); // D

  // ===== Company header =====
  sheet.getRange('A1:D1').merge()
    .setValue('START SCENE MEDIA PRODUKSIYON LIMITED')
    .setFontSize(13)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  sheet.getRange('A2:D2').merge()
    .setValue('212 My Office - Office No177 - Istanbul - Bagcilar')
    .setHorizontalAlignment('center');

  sheet.getRange('A3:D3').merge()
    .setValue('Finance@seenfilm.net  |   www.seenfilm.net')
    .setHorizontalAlignment('center');

  // ===== INVOICE title =====
  sheet.getRange('A5:D5').merge()
    .setValue('INVOICE')
    .setFontSize(18)
    .setFontWeight('bold')
    .setFontColor(CONFIG.COLORS.TEXT.WARNING)
    .setHorizontalAlignment('center');

  // ===== Invoice basic info =====
  sheet.getRange('A7').setValue('Invoice No:').setFontWeight('bold');
  sheet.getRange('B7').setValue(''); // سيتم ملؤه من الدالة

  sheet.getRange('A8').setValue('Invoice Date:').setFontWeight('bold');
  sheet.getRange('B8').setNumberFormat('yyyy-mm-dd'); // سيتم ملؤه من الدالة

  // ===== Client (TV Channel) =====
  sheet.getRange('A10').setValue('Client (TV Channel):').setFontWeight('bold');
  sheet.getRange('B10:D10').merge();   // Al Araby, Al Jazeera, ...

  sheet.getRange('A11').setValue('Client Email:').setFontWeight('bold');
  sheet.getRange('B11:D11').merge();

  // ===== Project name only =====
  sheet.getRange('A13').setValue('Project Name:').setFontWeight('bold');
  sheet.getRange('B13:D13').merge();

  // ===== Items table =====
  sheet.getRange('A15:D15')
    .setValues([['Description', 'Qty', 'Unit Price (USD)', 'Total (USD)']])
    .setBackground(CONFIG.COLORS.BG.GRAY)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  // سطر افتراضي للبند الرئيسي (هيتغيّر من الدالة)
  sheet.getRange('A16').setValue('Full project contract value');
  sheet.getRange('B16').setValue(1);
  sheet.getRange('C16').setNumberFormat('$#,##0.00');
  sheet.getRange('D16').setFormula('=B16*C16').setNumberFormat('$#,##0.00');

  // صفين إضافيين
  sheet.getRange('C17:C18').setNumberFormat('$#,##0.00');
  sheet.getRange('D17:D18').setFormulaR1C1('=RC[-1]*RC[-2]').setNumberFormat('$#,##0.00');

  // إجمالي
  sheet.getRange('C20').setValue('TOTAL:').setFontWeight('bold').setHorizontalAlignment('right');
  sheet.getRange('D20')
    .setFormula('=SUM(D16:D18)')
    .setNumberFormat('$#,##0.00')
    .setFontWeight('bold');

  // ===== Notes =====
  sheet.getRange('A22:D22').merge()
    .setValue('Notes:')
    .setFontWeight('bold')
    .setHorizontalAlignment('left');

  sheet.getRange('A23:D25').merge().setWrap(true);

  // ===== Bank details =====
  sheet.getRange('A27:D27').merge()
    .setValue('BANK ACCOUNT DETAILS')
    .setBackground(CONFIG.COLORS.BG.DARK_GRAY)
    .setFontColor(CONFIG.COLORS.TEXT.WHITE)
    .setFontWeight('bold')
    .setHorizontalAlignment('left');

  sheet.getRange('A28').setValue('Bank Name:').setFontWeight('bold');
  sheet.getRange('B28:D28').merge().setValue('KUVEYT TURK');

  sheet.getRange('A29').setValue('Account No:').setFontWeight('bold');
  sheet.getRange('B29:D29').merge().setValue('96160301');

  sheet.getRange('A30').setValue('IBAN:').setFontWeight('bold');
  sheet.getRange('B30:D30').merge().setValue('TR460020500009616030100101');

  sheet.getRange('A31').setValue('Account Name:').setFontWeight('bold');
  sheet.getRange('B31:D31').merge().setValue('Start Scene Media Produksiyon Limited');

  sheet.getRange('A32').setValue('Swift Code:').setFontWeight('bold');
  sheet.getRange('B32:D32').merge().setValue('KTEFTRIS');

  // خلي كل قيم بيانات البنك نص ومحاذاة موحدة لليسار
  sheet.getRange('B28:D32')
    .setNumberFormat('@')               // ← نص، وليس رقم
    .setHorizontalAlignment('left');    // ← نحيّة الشمال

  sheet.setFrozenRows(6);

  return sheet;
}

// ==================== إنشاء الفاتورة وملء البيانات ====================
function generateChannelInvoice() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const projectsSheet = ss.getSheetByName(CONFIG.SHEETS.PROJECTS);
  if (!projectsSheet) {
    ui.alert('⚠️ شيت "قاعدة بيانات المشاريع" غير موجود.');
    return;
  }

  // ١) نبني النموذج الإنجليزي في نفس التاب كل مرة
  const invoiceSheet = createInvoiceTemplateSheet(ss);

  // ٢) نطلب كود المشروع
  const response = ui.prompt(
    '🧾 Create invoice',
    'Enter the project code as in "قاعدة بيانات المشاريع":',
    ui.ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() !== ui.Button.OK) return;

  const projectCode = response.getResponseText().trim();
  if (!projectCode) {
    ui.alert('⚠️ No project code entered.');
    return;
  }

  // ٣) البحث عن المشروع
  const data = projectsSheet.getDataRange().getValues();
  const headers = data[0];        // صف العناوين
  let projectRow = null;
  let projectRowIndex = -1;       // رقم الصف (للحفظ لاحقًا)

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === projectCode) { // A = كود المشروع
      projectRow = data[i];
      projectRowIndex = i;
      break;
    }
  }

  if (!projectRow) {
    ui.alert('⚠️ Project not found: ' + projectCode);
    return;
  }

  const projectName = projectRow[1];              // اسم المشروع
  const projectType = projectRow[2];              // نوع المشروع
  const channelName = projectRow[3];              // القناة / الجهة
  const contractValue = Number(projectRow[8]) || 0; // قيمة العقد مع القناة

  if (!contractValue) {
    ui.alert('⚠️ Contract value is zero or missing (قيمة العقد مع القناة).');
    return;
  }

  // ٤) إيميل القناة (اختياري)
  const emailResp = ui.prompt(
    'Client email (optional)',
    'Enter client email (TV channel) or leave blank:',
    ui.ButtonSet.OK_CANCEL
  );
  const clientEmail = (emailResp.getSelectedButton() === ui.Button.OK)
    ? emailResp.getResponseText().trim()
    : '';

  // ٥) رقم الفاتورة والتاريخ
  const today = new Date();

  // استخراج أجزاء كود المشروع: OM-PA-25-002
  const codeParts = projectCode.split('-');
  const channelCode = codeParts[0] || '';  // OM
  const yearCode = codeParts[2] || '';     // 25
  const seqCode = codeParts[3] || '';      // 002

  // صيغة رقم الفاتورة: OM-25-002-1214
  const invoiceNumber = channelCode + '-' + yearCode + '-' + seqCode + '-' +
    Utilities.formatDate(today, Session.getScriptTimeZone(), 'MMdd');

  invoiceSheet.getRange('B7').setValue(invoiceNumber);
  invoiceSheet.getRange('B8').setValue(today).setNumberFormat('yyyy-mm-dd');

  // ٦) بيانات العميل والمشروع
  invoiceSheet.getRange('B10').setValue(channelName || '');
  invoiceSheet.getRange('B11').setValue(clientEmail || '');
  invoiceSheet.getRange('B13').setValue(projectName || '');

  // ٧) البند الرئيسي في الجدول — الوصف = نوع المشروع + اسم المشروع
  let descriptionText = '';
  if (projectType) descriptionText += projectType;
  if (projectType && projectName) descriptionText += ' - ';
  if (projectName) descriptionText += projectName;

  invoiceSheet.getRange('A16').setValue(descriptionText || projectName || projectType || '');
  invoiceSheet.getRange('B16').setValue(1);
  invoiceSheet.getRange('C16')
    .setValue(contractValue)
    .setNumberFormat('$#,##0.00');
  // D16 من القالب = B16*C16

  // ٨) حفظ رقم الفاتورة داخل "قاعدة بيانات المشاريع" في عمود جديد (رقم آخر فاتورة)
  let invoiceColIndex = headers.indexOf('رقم آخر فاتورة');
  if (invoiceColIndex === -1) {
    // لو العنوان مش موجود، نحطه في أول عمود فاضي بعد آخر عنوان
    invoiceColIndex = headers.length;
    projectsSheet.getRange(1, invoiceColIndex + 1).setValue('رقم آخر فاتورة');
  }
  // projectRowIndex هو إندكس الصف داخل data، فصف الشيت = projectRowIndex + 1
  projectsSheet.getRange(projectRowIndex + 1, invoiceColIndex + 1).setValue(invoiceNumber);

  // ٩) تسجيل الحركة في دفتر الحركات المالية
  const transSheet = ss.getSheetByName(CONFIG.SHEETS.TRANSACTIONS);
  if (!transSheet) {
    ui.alert('⚠️ خطأ', 'شيت "دفتر الحركات المالية" غير موجود!', ui.ButtonSet.OK);
    return;
  }

  try {
    // حساب تاريخ الاستحقاق (بعد شهر من تاريخ الفاتورة)
    const dueDate = new Date(today);
    dueDate.setMonth(dueDate.getMonth() + 1);

    // البحث عن آخر صف يحتوي على تاريخ فعلي في عمود B (التاريخ)
    // هذا يتجنب مشكلة الصفوف التي تحتوي على dropdowns فقط
    const dateColumn = transSheet.getRange('B:B').getValues();
    let lastDataRow = 1; // البداية من صف العناوين
    for (let i = dateColumn.length - 1; i >= 1; i--) {
      const cellValue = dateColumn[i][0];
      // التحقق من أن الخلية تحتوي على تاريخ فعلي
      if (cellValue && (cellValue instanceof Date || (typeof cellValue === 'string' && cellValue.trim() !== ''))) {
        lastDataRow = i + 1; // +1 لأن الفهرس يبدأ من 0
        break;
      }
    }
    const newRow = lastDataRow + 1;

    // البيانات حسب الأعمدة (A إلى T = 20 عمود):
    // A: # (تلقائي), B: التاريخ, C: طبيعة الحركة, D: تصنيف, E: كود المشروع
    // F: اسم المشروع, G: البند, H: التفاصيل, I: اسم الطرف
    // J: المبلغ, K: العملة, L: سعر الصرف, M: القيمة بالدولار, N: نوع الحركة
    // O: الرصيد (محسوب), P: رقم مرجعي, Q: طريقة الدفع
    // R: نوع شرط الدفع, S: عدد الأسابيع, T: تاريخ مخصص

    const rowData = [
      newRow - 1,                   // A: رقم الحركة (رقم الصف - 1)
      today,                        // B: التاريخ
      'استحقاق إيراد',              // C: طبيعة الحركة (بدون إيموجي)
      'ايراد',                      // D: تصنيف الحركة
      projectCode,                  // E: كود المشروع
      projectName,                  // F: اسم المشروع
      'ايراد',                      // G: البند
      descriptionText,              // H: التفاصيل
      channelName,                  // I: اسم الطرف
      contractValue,                // J: المبلغ
      'USD',                        // K: العملة
      1,                            // L: سعر الصرف
      contractValue,                // M: القيمة بالدولار
      'مدين استحقاق',               // N: نوع الحركة
      '',                           // O: الرصيد (سيُحسب تلقائياً)
      invoiceNumber,                // P: رقم مرجعي (رقم الفاتورة)
      'تحويل بنكي',                 // Q: طريقة الدفع
      'تاريخ مخصص',                 // R: نوع شرط الدفع
      '',                           // S: عدد الأسابيع (فارغ)
      dueDate                       // T: تاريخ مخصص (بعد شهر)
    ];

    // كتابة البيانات من A إلى T (20 عمود)
    transSheet.getRange(newRow, 1, 1, 20).setValues([rowData]);

    // تنسيق التاريخ والمبالغ
    transSheet.getRange(newRow, 2).setNumberFormat('dd/mm/yyyy');   // B: التاريخ
    transSheet.getRange(newRow, 10).setNumberFormat('#,##0.00');    // J: المبلغ
    transSheet.getRange(newRow, 13).setNumberFormat('$#,##0.00');   // M: القيمة بالدولار
    transSheet.getRange(newRow, 20).setNumberFormat('dd/mm/yyyy');  // T: تاريخ مخصص

    // تحديث الأعمدة المحسوبة (U, O, V) للصف الجديد
    SpreadsheetApp.flush(); // التأكد من حفظ البيانات أولاً

    // حساب تاريخ الاستحقاق (U) - سيأخذ القيمة من T (تاريخ مخصص)
    calculateDueDate_(ss, transSheet, newRow);

    // حساب الرصيد (O) وحالة السداد (V)
    recalculatePartyBalance_(transSheet, newRow);

    // حساب تاريخ الاستحقاق للعرض في الرسالة
    const dueDateFormatted = Utilities.formatDate(dueDate, Session.getScriptTimeZone(), 'dd/MM/yyyy');

    // ١٠) رسالة نجاح
    ui.alert(
      '✅ تم إنشاء الفاتورة بنجاح!\n\n' +
      '• رقم الفاتورة: ' + invoiceNumber + '\n' +
      '• القيمة: $' + contractValue.toLocaleString() + '\n' +
      '• تاريخ الاستحقاق: ' + dueDateFormatted + '\n' +
      '• تم حفظ رقم الفاتورة في "قاعدة بيانات المشاريع"\n' +
      '• تم تسجيل الحركة في "دفتر الحركات المالية" (صف ' + newRow + ')'
    );

  } catch (error) {
    ui.alert('⚠️ خطأ في تسجيل الحركة', 'حدث خطأ أثناء تسجيل الحركة:\n' + error.message, ui.ButtonSet.OK);
    console.error('خطأ في generateChannelInvoice:', error);
    return;
  }
}

// ==================== 🔄 إعادة طباعة فاتورة موجودة ====================
/**
 * إعادة طباعة فاتورة موجودة من كود المشروع
 * يبحث عن رقم الفاتورة المحفوظ في قاعدة بيانات المشاريع
 * ويعيد إنشاء شيت الفاتورة بنفس البيانات
 */
function regenerateChannelInvoice() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const projectsSheet = ss.getSheetByName(CONFIG.SHEETS.PROJECTS);
  if (!projectsSheet) {
    ui.alert('⚠️ شيت "قاعدة بيانات المشاريع" غير موجود.');
    return;
  }

  // طلب كود المشروع
  const response = ui.prompt(
    '🔄 إعادة طباعة فاتورة',
    'أدخل كود المشروع للبحث عن الفاتورة:',
    ui.ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() !== ui.Button.OK) return;

  const projectCode = response.getResponseText().trim();
  if (!projectCode) {
    ui.alert('⚠️ لم يتم إدخال كود المشروع.');
    return;
  }

  // البحث عن المشروع
  const data = projectsSheet.getDataRange().getValues();
  const headers = data[0];
  let projectRow = null;

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === projectCode) {
      projectRow = data[i];
      break;
    }
  }

  if (!projectRow) {
    ui.alert('⚠️ المشروع غير موجود: ' + projectCode);
    return;
  }

  // البحث عن عمود رقم الفاتورة
  const invoiceColIndex = headers.indexOf('رقم آخر فاتورة');
  if (invoiceColIndex === -1) {
    ui.alert('⚠️ لا يوجد عمود "رقم آخر فاتورة" في قاعدة بيانات المشاريع.\n\nيبدو أنه لم يتم إنشاء فاتورة لهذا المشروع من قبل.');
    return;
  }

  const invoiceNumber = projectRow[invoiceColIndex];
  if (!invoiceNumber) {
    ui.alert('⚠️ لا توجد فاتورة سابقة لهذا المشروع.\n\nاستخدم "إنشاء فاتورة قناة" لإنشاء فاتورة جديدة.');
    return;
  }

  // استخراج بيانات المشروع
  const projectName = projectRow[1];
  const projectType = projectRow[2];
  const channelName = projectRow[3];
  const contractValue = Number(projectRow[8]) || 0;

  if (!contractValue) {
    ui.alert('⚠️ قيمة العقد فارغة أو صفر.');
    return;
  }

  // البحث عن تاريخ الفاتورة من دفتر الحركات (إن وجد)
  let invoiceDate = new Date();
  const transSheet = ss.getSheetByName(CONFIG.SHEETS.TRANSACTIONS);
  if (transSheet) {
    const transData = transSheet.getDataRange().getValues();
    for (let i = 1; i < transData.length; i++) {
      // البحث في عمود P (رقم مرجعي) = 16
      if (transData[i][15] === invoiceNumber) {
        invoiceDate = transData[i][1] || new Date(); // B: التاريخ
        break;
      }
    }
  }

  // إنشاء شيت الفاتورة
  const invoiceSheet = createInvoiceTemplateSheet(ss);

  // ملء بيانات الفاتورة
  invoiceSheet.getRange('B7').setValue(invoiceNumber);
  invoiceSheet.getRange('B8').setValue(invoiceDate).setNumberFormat('yyyy-mm-dd');
  invoiceSheet.getRange('B10').setValue(channelName || '');
  invoiceSheet.getRange('B13').setValue(projectName || '');

  // الوصف
  let descriptionText = '';
  if (projectType) descriptionText += projectType;
  if (projectType && projectName) descriptionText += ' - ';
  if (projectName) descriptionText += projectName;

  invoiceSheet.getRange('A16').setValue(descriptionText || projectName || projectType || '');
  invoiceSheet.getRange('B16').setValue(1);
  invoiceSheet.getRange('C16')
    .setValue(contractValue)
    .setNumberFormat('$#,##0.00');

  // رسالة نجاح
  ui.alert(
    '✅ تم إعادة طباعة الفاتورة بنجاح!\n\n' +
    '• رقم الفاتورة: ' + invoiceNumber + '\n' +
    '• المشروع: ' + projectName + '\n' +
    '• القناة: ' + channelName + '\n' +
    '• القيمة: $' + contractValue.toLocaleString()
  );
}

// ==================== 📄 إنشاء كشف حساب من صف في دفتر الحركات ====================
/**
 * دالة مساعدة تُستدعى من onEdit عند التعديل على عمود "كشف" (Y)
 * تقرأ اسم الطرف من الصف وتنشئ كشف حساب له تلقائياً
 */
function generateStatementFromRow_(ss, sheet, row) {
  const ui = SpreadsheetApp.getUi();

  // قراءة اسم الطرف من عمود I (9)
  const partyName = sheet.getRange(row, 9).getValue();

  if (!partyName || String(partyName).trim() === '') {
    ui.alert('⚠️ تنبيه', 'لا يوجد اسم طرف في هذا الصف!', ui.ButtonSet.OK);
    // إعادة الرمز للخلية
    sheet.getRange(row, 25).setValue('📄');
    return;
  }

  // البحث عن نوع الطرف في قاعدة البيانات
  const partiesSheet = ss.getSheetByName(CONFIG.SHEETS.PARTIES);
  let partyType = null;

  if (partiesSheet) {
    const partiesData = partiesSheet.getRange('A2:B500').getValues();
    for (let i = 0; i < partiesData.length; i++) {
      if (partiesData[i][0] === partyName) {
        partyType = partiesData[i][1]; // B: نوع الطرف
        break;
      }
    }
  }

  // إعادة الرمز للخلية فوراً
  sheet.getRange(row, 25).setValue('📄');

  if (!partyType) {
    ui.alert('⚠️ تنبيه', 'الطرف "' + partyName + '" غير موجود في قاعدة بيانات الأطراف!', ui.ButtonSet.OK);
    return;
  }

  // استدعاء الدالة الموحدة
  generateUnifiedStatement_(ss, partyName, partyType);
}

/**
 * ═══════════════════════════════════════════════════════════════════════════
 * الدالة الموحدة لإنشاء كشف حساب (مورد/عميل/ممول)
 * تجمع بين التنسيق الجيد والبيانات الصحيحة
 * ═══════════════════════════════════════════════════════════════════════════
 */
function generateUnifiedStatement_(ss, partyName, partyType) {
  const ui = SpreadsheetApp.getUi();
  const transSheet = ss.getSheetByName(CONFIG.SHEETS.TRANSACTIONS);

  if (!transSheet) {
    ui.alert('❌ خطأ', 'شيت دفتر الحركات المالية غير موجود!', ui.ButtonSet.OK);
    return;
  }

  // ═══════════════════════════════════════════════════════════
  // جلب لوجو الشركة من قاعدة بيانات البنود (D2)
  // يدعم روابط Google Drive العادية ويحولها لروابط مباشرة
  // ═══════════════════════════════════════════════════════════
  let companyLogo = '';
  try {
    const itemsSheet = ss.getSheetByName(CONFIG.SHEETS.ITEMS || 'قاعدة بيانات البنود');
    if (itemsSheet) {
      let logoUrl = itemsSheet.getRange('D2').getValue() || '';

      // تحويل رابط Google Drive العادي إلى رابط مباشر
      // مثال: https://drive.google.com/file/d/FILE_ID/view?usp=drive_link
      // يتحول إلى: https://drive.google.com/uc?id=FILE_ID
      if (logoUrl && logoUrl.includes('drive.google.com/file/d/')) {
        const match = logoUrl.match(/\/file\/d\/([^\/\?]+)/);
        if (match && match[1]) {
          logoUrl = 'https://drive.google.com/uc?id=' + match[1];
        }
      }

      // التحقق من أن الرابط صالح (ليس مجلد)
      if (logoUrl && !logoUrl.includes('/folders/')) {
        companyLogo = logoUrl;
      }
      Logger.log('🖼️ Company logo URL: ' + (companyLogo ? companyLogo : 'Not found'));
    }
  } catch (e) {
    Logger.log('⚠️ Could not get company logo: ' + e.message);
  }

  // ═══════════════════════════════════════════════════════════
  // تحديد عنوان الكشف ولون التبويب حسب نوع الطرف
  // ═══════════════════════════════════════════════════════════
  let titlePrefix = 'كشف حساب';
  let tabColor = '#4a86e8';

  if (partyType === 'مورد') {
    titlePrefix = 'كشف مورد';
    tabColor = CONFIG.COLORS.TAB.VENDOR_STATEMENT || '#e91e63';
  } else if (partyType === 'عميل') {
    titlePrefix = 'كشف عميل';
    tabColor = CONFIG.COLORS.TAB.CLIENT_STATEMENT || '#4caf50';
  } else if (partyType === 'ممول') {
    titlePrefix = 'كشف ممول';
    tabColor = CONFIG.COLORS.TAB.FUNDER_STATEMENT || '#ff9800';
  }

  // ═══════════════════════════════════════════════════════════
  // إنشاء أو الحصول على شيت الكشف
  // ═══════════════════════════════════════════════════════════
  const sheetName = titlePrefix + ' - ' + partyName;
  let sheet = ss.getSheetByName(sheetName);

  if (sheet) {
    const confirm = ui.alert(
      '📋 كشف موجود',
      'يوجد كشف حساب لـ "' + partyName + '" بالفعل.\n\nهل تريد تحديثه؟',
      ui.ButtonSet.YES_NO
    );
    if (confirm !== ui.Button.YES) return;
    sheet.clear();
  } else {
    sheet = ss.insertSheet(sheetName);
  }

  sheet.setTabColor(tabColor);
  sheet.setRightToLeft(true);

  // ═══════════════════════════════════════════════════════════
  // عرض الأعمدة (7 أعمدة مع رقم الأوردر)
  // ═══════════════════════════════════════════════════════════
  sheet.setColumnWidth(1, 110);  // التاريخ
  sheet.setColumnWidth(2, 160);  // المشروع
  sheet.setColumnWidth(3, 130);  // رقم الأوردر
  sheet.setColumnWidth(4, 220);  // التفاصيل
  sheet.setColumnWidth(5, 120);  // مدين
  sheet.setColumnWidth(6, 120);  // دائن
  sheet.setColumnWidth(7, 120);  // الرصيد

  // ═══════════════════════════════════════════════════════════
  // بيانات الطرف من القاعدة
  // ═══════════════════════════════════════════════════════════
  const partyData = getPartyData_(ss, partyName, partyType);

  // ═══════════════════════════════════════════════════════════
  // العنوان الرئيسي
  // ═══════════════════════════════════════════════════════════
  sheet.getRange('A1:G1').merge();
  sheet.getRange('A1')
    .setValue('📊 ' + titlePrefix)
    .setBackground(CONFIG.COLORS.HEADER.DASHBOARD)
    .setFontColor(CONFIG.COLORS.TEXT.WHITE)
    .setFontWeight('bold')
    .setFontSize(15)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  // ═══════════════════════════════════════════════════════════
  // إضافة اللوجو من Google Drive (باستخدام insertImage)
  // ═══════════════════════════════════════════════════════════
  let logoRowOffset = 0;
  if (companyLogo) {
    try {
      // استخراج File ID من الرابط
      let fileId = '';
      if (companyLogo.includes('uc?id=')) {
        fileId = companyLogo.split('uc?id=')[1].split('&')[0];
      } else if (companyLogo.includes('/file/d/')) {
        const match = companyLogo.match(/\/file\/d\/([^\/\?]+)/);
        if (match) fileId = match[1];
      }

      if (fileId) {
        // جلب الصورة من Drive وإدراجها
        const file = DriveApp.getFileById(fileId);
        const blob = file.getBlob();

        // إدراج الصورة في الصف 2
        sheet.setRowHeight(2, 80);
        const image = sheet.insertImage(blob, 3, 2); // العمود C، الصف 2
        image.setWidth(70);
        image.setHeight(70);

        logoRowOffset = 1;
        Logger.log('✅ Logo inserted from Drive: ' + fileId);
      } else {
        Logger.log('⚠️ Could not extract file ID from logo URL');
      }
    } catch (e) {
      Logger.log('⚠️ Could not insert logo: ' + e.message);
      logoRowOffset = 0;
    }
  }

  // ═══════════════════════════════════════════════════════════
  // كارت بيانات الطرف (ديناميكي حسب وجود اللوجو)
  // ═══════════════════════════════════════════════════════════
  const cardHeaderRow = 3 + logoRowOffset;
  const cardDataStartRow = cardHeaderRow + 1;

  sheet.getRange('A' + cardHeaderRow + ':G' + cardHeaderRow).merge()
    .setValue('بيانات ' + partyType)
    .setBackground(CONFIG.COLORS.HEADER.SUMMARY)
    .setFontColor(CONFIG.COLORS.TEXT.WHITE)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  sheet.getRange('A' + cardDataStartRow + ':G' + (cardDataStartRow + 3)).setBackground(CONFIG.COLORS.BG.LIGHT_BLUE);

  sheet.getRange('A' + cardDataStartRow).setValue('الاسم:').setFontWeight('bold');
  sheet.getRange('B' + cardDataStartRow + ':D' + cardDataStartRow).merge().setValue(partyName);

  sheet.getRange('E' + cardDataStartRow).setValue('التخصص:').setFontWeight('bold');
  sheet.getRange('F' + cardDataStartRow + ':G' + cardDataStartRow).merge().setValue(partyData.specialization || '');

  sheet.getRange('A' + (cardDataStartRow + 1)).setValue('الهاتف:').setFontWeight('bold');
  sheet.getRange('B' + (cardDataStartRow + 1) + ':D' + (cardDataStartRow + 1)).merge().setValue(partyData.phone || '');

  sheet.getRange('E' + (cardDataStartRow + 1)).setValue('البريد:').setFontWeight('bold');
  sheet.getRange('F' + (cardDataStartRow + 1) + ':G' + (cardDataStartRow + 1)).merge().setValue(partyData.email || '');

  sheet.getRange('A' + (cardDataStartRow + 2)).setValue('البنك:').setFontWeight('bold');
  sheet.getRange('B' + (cardDataStartRow + 2) + ':G' + (cardDataStartRow + 2)).merge().setValue(partyData.bankInfo || '');

  sheet.getRange('A' + (cardDataStartRow + 3)).setValue('ملاحظات:').setFontWeight('bold');
  sheet.getRange('B' + (cardDataStartRow + 3) + ':G' + (cardDataStartRow + 3)).merge().setValue(partyData.notes || '').setWrap(true);

  sheet.getRange('A' + cardDataStartRow + ':G' + (cardDataStartRow + 3)).setBorder(
    true, true, true, true, true, true,
    '#1565c0', SpreadsheetApp.BorderStyle.SOLID
  );

  // ═══════════════════════════════════════════════════════════
  // استخراج حركات الطرف (بدون فلتر طبيعة الحركة)
  // ═══════════════════════════════════════════════════════════
  const data = transSheet.getDataRange().getValues();
  const rows = [];

  let totalDebit = 0, totalCredit = 0, balance = 0;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];

    // الفلتر الوحيد: اسم الطرف
    if (row[8] !== partyName) continue;

    const movementKind = String(row[13] || '');  // N: نوع الحركة
    const amountUsd = Number(row[12]) || 0;     // M: القيمة بالدولار

    // تجاهل الحركات بدون مبلغ
    if (!amountUsd) continue;

    const date = row[1];       // B: التاريخ
    const project = row[5];    // F: اسم المشروع
    const details = row[7];    // H: التفاصيل
    const orderNumber = row[25] || '';  // Z: رقم الأوردر

    let debit = 0, credit = 0;

    // استخدام includes للتعامل مع الإيموجي
    if (movementKind.includes(CONFIG.MOVEMENT.DEBIT) || movementKind.includes('مدين')) {
      debit = amountUsd;
      balance += debit;
      totalDebit += debit;
    } else if (movementKind.includes(CONFIG.MOVEMENT.CREDIT) || movementKind.includes('دائن')) {
      credit = amountUsd;
      balance -= credit;
      totalCredit += credit;
    }

    rows.push([
      date,
      project || '',
      orderNumber,
      details || '',
      debit || '',
      credit || '',
      Math.round(balance * 100) / 100
    ]);
  }

  // ترتيب زمني
  rows.sort((a, b) => {
    const dateA = a[0] instanceof Date ? a[0].getTime() : new Date(a[0]).getTime();
    const dateB = b[0] instanceof Date ? b[0].getTime() : new Date(b[0]).getTime();
    return dateA - dateB;
  });

  // إعادة حساب الرصيد بعد الترتيب
  balance = 0;
  for (let i = 0; i < rows.length; i++) {
    const debit = rows[i][4] || 0;
    const credit = rows[i][5] || 0;
    balance += debit - credit;
    rows[i][6] = Math.round(balance * 100) / 100;
  }

  // ═══════════════════════════════════════════════════════════
  // الملخص المالي (ديناميكي حسب وجود اللوجو)
  // ═══════════════════════════════════════════════════════════
  const summaryHeaderRow = cardDataStartRow + 5;
  const summaryDataStartRow = summaryHeaderRow + 1;

  sheet.getRange('A' + summaryHeaderRow + ':G' + summaryHeaderRow).merge()
    .setValue('الملخص المالي')
    .setBackground(CONFIG.COLORS.HEADER.SUMMARY)
    .setFontColor(CONFIG.COLORS.TEXT.WHITE)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  sheet.getRange('A' + summaryDataStartRow + ':G' + (summaryDataStartRow + 1)).setBackground(CONFIG.COLORS.BG.LIGHT_BLUE);

  sheet.getRange('A' + summaryDataStartRow).setValue('إجمالي المدين:').setFontWeight('bold');
  sheet.getRange('B' + summaryDataStartRow).setValue(totalDebit).setNumberFormat('$#,##0.00');

  sheet.getRange('D' + summaryDataStartRow).setValue('إجمالي الدائن:').setFontWeight('bold');
  sheet.getRange('E' + summaryDataStartRow).setValue(totalCredit).setNumberFormat('$#,##0.00');

  sheet.getRange('A' + (summaryDataStartRow + 1)).setValue('الرصيد الحالي:').setFontWeight('bold');
  sheet.getRange('B' + (summaryDataStartRow + 1)).setValue(balance).setNumberFormat('$#,##0.00')
    .setFontWeight('bold')
    .setBackground(balance > 0 ? '#ffcdd2' : '#c8e6c9');

  sheet.getRange('E' + (summaryDataStartRow + 1)).setValue('عدد الحركات:').setFontWeight('bold');
  sheet.getRange('F' + (summaryDataStartRow + 1)).setValue(rows.length);

  sheet.getRange('A' + summaryDataStartRow + ':G' + (summaryDataStartRow + 1)).setBorder(
    true, true, true, true, true, true,
    '#1565c0', SpreadsheetApp.BorderStyle.SOLID
  );

  // ═══════════════════════════════════════════════════════════
  // رأس جدول الحركات (ديناميكي حسب وجود اللوجو)
  // ═══════════════════════════════════════════════════════════
  const tableHeaderRow = summaryDataStartRow + 3;
  const headers = [
    '📅 التاريخ',
    '🎬 المشروع',
    '📦 رقم الأوردر',
    '📝 التفاصيل',
    '💰 مدين (USD)',
    '💸 دائن (USD)',
    '📊 الرصيد (USD)'
  ];

  sheet.getRange(tableHeaderRow, 1, 1, headers.length)
    .setValues([headers])
    .setBackground(CONFIG.COLORS.HEADER.DASHBOARD)
    .setFontColor(CONFIG.COLORS.TEXT.WHITE)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  // ═══════════════════════════════════════════════════════════
  // بيانات الحركات (ديناميكي حسب وجود اللوجو)
  // ═══════════════════════════════════════════════════════════
  const dataStartRow = tableHeaderRow + 1;

  if (rows.length > 0) {
    sheet.getRange(dataStartRow, 1, rows.length, headers.length).setValues(rows);
    sheet.getRange(dataStartRow, 1, rows.length, 1).setNumberFormat('dd/mm/yyyy');
    sheet.getRange(dataStartRow, 5, rows.length, 3).setNumberFormat('$#,##0.00');

    // تلوين متناوب للصفوف
    for (let i = 0; i < rows.length; i++) {
      const r = dataStartRow + i;
      const bg = i % 2 === 0 ? '#ffffff' : CONFIG.COLORS.BG.LIGHT_BLUE;
      sheet.getRange(r, 1, 1, headers.length).setBackground(bg);
    }

    // إطار الجدول
    sheet.getRange(tableHeaderRow, 1, rows.length + 1, headers.length)
      .setBorder(true, true, true, true, true, true, '#bdbdbd', SpreadsheetApp.BorderStyle.SOLID);
  } else {
    sheet.getRange(dataStartRow, 1).setValue('لا توجد حركات').setFontStyle('italic');
  }

  sheet.setFrozenRows(tableHeaderRow);

  // ═══════════════════════════════════════════════════════════
  // التذييل (ديناميكي حسب وجود اللوجو)
  // ═══════════════════════════════════════════════════════════
  const footerStart = dataStartRow + Math.max(rows.length, 1) + 3;

  sheet.getRange(footerStart, 1, 1, 7).merge()
    .setBackground(CONFIG.COLORS.HEADER.DASHBOARD);

  sheet.getRange(footerStart + 1, 1, 3, 7).merge()
    .setValue(
      "Seen Film\n" +
      "info@seenfilm.net | www.seenfilm.net\n" +
      "تاريخ الإنشاء: " + new Date().toLocaleDateString('ar-EG')
    )
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setFontSize(10)
    .setFontColor(CONFIG.COLORS.TEXT.DARK);

  // تفعيل الشيت
  ss.setActiveSheet(sheet);

  ui.alert(
    '✅ تم بنجاح',
    'تم إنشاء ' + titlePrefix + ' لـ "' + partyName + '"\n\n' +
    '• عدد الحركات: ' + rows.length + '\n' +
    '• الرصيد الحالي: $' + balance.toLocaleString(),
    ui.ButtonSet.OK
  );
}

// ==================== كشف حساب مورد - في شيت (يستخدم الدالة الموحدة) ====================
function generateVendorStatementSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // اسم المورد
  const response = ui.prompt(
    '📄 كشف حساب مورد',
    'اكتب اسم المورد كما هو مسجل:',
    ui.ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() !== ui.Button.OK) return;

  const vendorName = response.getResponseText().trim();
  if (!vendorName) {
    ui.alert('⚠️ لم يتم إدخال الاسم.');
    return;
  }

  // استدعاء الدالة الموحدة
  generateUnifiedStatement_(ss, vendorName, 'مورد');
}

// ==================== كشف حساب عميل - في شيت (يستخدم الدالة الموحدة) ====================
function generateClientStatementSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const response = ui.prompt(
    '📄 كشف حساب عميل',
    'اكتب اسم العميل كما هو مسجل:',
    ui.ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() !== ui.Button.OK) return;

  const clientName = response.getResponseText().trim();
  if (!clientName) {
    ui.alert('⚠️ لم يتم إدخال الاسم.');
    return;
  }

  // استدعاء الدالة الموحدة
  generateUnifiedStatement_(ss, clientName, 'عميل');
}

// ==================== كشف حساب ممول - في شيت (يستخدم الدالة الموحدة) ====================
function generateFunderStatementSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const response = ui.prompt(
    '📄 كشف حساب ممول',
    'اكتب اسم الممول كما هو مسجل:',
    ui.ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() !== ui.Button.OK) return;

  const funderName = response.getResponseText().trim();
  if (!funderName) {
    ui.alert('⚠️ لم يتم إدخال الاسم.');
    return;
  }

  // استدعاء الدالة الموحدة
  generateUnifiedStatement_(ss, funderName, 'ممول');
}
// ==================== إعادة بناء تقرير المشروع التفصيلي ====================

function rebuildProjectDetailReport(silent) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const transSheet = ss.getSheetByName(CONFIG.SHEETS.TRANSACTIONS);
  const reportSheet = ss.getSheetByName(CONFIG.SHEETS.PROJECT_REPORT);

  if (!transSheet || !reportSheet) {
    return silent ? { success: false, name: 'تقرير المشروع التفصيلي', error: 'الشيتات غير موجودة' } : undefined;
  }

  const data = transSheet.getDataRange().getValues();
  const map = {}; // key = projectCode|projectName|item|vendor

  for (let i = 1; i < data.length; i++) {
    const row = data[i];

    const projectCode = String(row[4] || '').trim();  // E: كود المشروع
    const projectName = String(row[5] || '').trim();  // F: اسم المشروع
    const item = String(row[6] || '').trim();  // G: البند
    const vendor = String(row[8] || '').trim();  // I: المورد / الجهة
    const type = String(row[2] || '').trim();  // C: طبيعة الحركة
    const amountUsd = Number(row[12]) || 0;         // M: القيمة بالدولار الموحد

    // لازم يكون في مشروع + جهة + نوع حركة + قيمة
    if (!projectCode || !vendor || !type || !amountUsd) continue;

    const key = [projectCode, projectName, item, vendor].join('||');

    if (!map[key]) {
      map[key] = {
        projectCode,
        projectName,
        item,
        vendor,
        totalDue: 0,     // إجمالي المستحق (مصروف + إيراد) بالدولار
        totalPaid: 0,    // المدفوع / المحصل بالدولار
        payments: 0      // عدد الدفعات / التحصيلات
      };
    }

    // 🔹 أي "استحقاق" (مصروف أو إيراد) يروح في إجمالي المستحق
    // استخدام includes للتعامل مع القيم التي تحتوي على إيموجي
    if (type.includes('استحقاق مصروف') || type.includes('استحقاق إيراد')) {
      map[key].totalDue += amountUsd;
    }

    // 🔹 أي "دفعة" أو "تحصيل" يروح في المدفوع
    if (type.includes('دفعة مصروف') || type.includes('تحصيل إيراد')) {
      map[key].totalPaid += amountUsd;
      if (amountUsd > 0) map[key].payments++;
    }
  }

  const rows = [];
  Object.keys(map).forEach(k => {
    const v = map[k];
    const remaining = v.totalDue - v.totalPaid;

    let status = 'لا يوجد استحقاق';
    if (v.totalDue > 0) {
      if (remaining === 0) {
        status = 'مسدد بالكامل';
      } else if (remaining > 0 && v.totalPaid > 0) {
        status = 'مسدد جزئياً';
      } else if (remaining > 0 && v.totalPaid === 0) {
        status = 'معلق';
      }
    }

    rows.push([
      v.projectCode,   // كود المشروع
      v.projectName,   // اسم المشروع
      v.item,          // البند
      v.vendor,        // الجهة (مورد / عميل / ممول)
      v.totalDue,      // إجمالي المستحق (USD)
      v.totalPaid,     // المدفوع / المحصل (USD)
      remaining,       // المتبقي (USD)
      v.payments,      // عدد الدفعات / التحصيلات
      status           // حالة السداد (محسوبة)
    ]);
  });

  // مسح التقرير القديم
  const lastCol = reportSheet.getLastColumn();
  if (reportSheet.getMaxRows() > 1) {
    reportSheet.getRange(2, 1, reportSheet.getMaxRows() - 1, lastCol).clearContent();
  }

  // كتابة التقرير الجديد
  if (rows.length) {
    rows.sort((a, b) => a[0].localeCompare(b[0]));
    reportSheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
    // المستحق + المدفوع + المتبقي
    reportSheet.getRange(2, 5, rows.length, 3).setNumberFormat('$#,##0.00');
  }

  return silent ? { success: true, name: 'تقرير المشروع التفصيلي' } : undefined;
}

// ==================== إعادة بناء تقرير الموردين الملخص ====================

function rebuildVendorSummaryReport(silent) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const transSheet = ss.getSheetByName(CONFIG.SHEETS.TRANSACTIONS);
  const reportSheet = ss.getSheetByName(CONFIG.SHEETS.VENDORS_REPORT);

  if (!transSheet || !reportSheet) {
    if (silent) return { success: false, name: 'تقرير الموردين', error: 'الشيتات غير موجودة' };
    SpreadsheetApp.getUi().alert('⚠️ تأكد من وجود "دفتر الحركات المالية" و "تقرير الموردين".');
    return;
  }

  // خريطة تخصص المورد (من القاعدة الموحدة مع fallback للقديمة)
  const specialMap = getPartySpecializationMap_(ss, 'مورد');

  const data = transSheet.getDataRange().getValues();
  const map = {};

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const vendor = row[8];               // I: اسم المورد/الجهة
    const type = row[2];               // C: طبيعة الحركة
    const movementKind = row[13];        // N: نوع الحركة (مدين استحقاق / دائن دفعة)
    const amountUsd = Number(row[12]) || 0; // M: القيمة بالدولار
    const project = row[4];              // E: كود المشروع
    const date = row[1];              // B: التاريخ

    if (!vendor || !amountUsd) continue;

    // استخدام includes للتعامل مع القيم التي تحتوي على إيموجي
    const typeStr = String(type || '');
    const movementStr = String(movementKind || '');

    // فلترة حركات الموردين فقط (استحقاق مصروف أو دفعة مصروف)
    if (!typeStr.includes('استحقاق مصروف') && !typeStr.includes('دفعة مصروف')) continue;

    if (!map[vendor]) {
      map[vendor] = {
        vendor,
        specialization: specialMap[vendor] || '',
        projects: new Set(),
        totalDebitUsd: 0,   // مدين (مستحق للمورد)
        totalCreditUsd: 0,  // دائن (مدفوع للمورد)
        payments: 0,
        lastDate: null
      };
    }

    const v = map[vendor];
    if (project) v.projects.add(project);

    // استخدام عمود N (نوع الحركة) للحساب - نفس منطق كشف الحساب
    if (movementStr.includes('مدين استحقاق') || movementStr.includes('مدين')) {
      v.totalDebitUsd += amountUsd;
    } else if (movementStr.includes('دائن دفعة') || movementStr.includes('دائن')) {
      v.totalCreditUsd += amountUsd;
      if (amountUsd > 0) v.payments++;
    }

    if (date) {
      const d = new Date(date);
      if (!v.lastDate || d > v.lastDate) {
        v.lastDate = d;
      }
    }
  }

  const rows = [];
  Object.keys(map).forEach(k => {
    const v = map[k];
    const projectsCount = v.projects.size;
    const currentBalance = v.totalDebitUsd - v.totalCreditUsd;  // مدين - دائن = الرصيد

    let status = 'مغلق';
    if (currentBalance > 0) status = 'له رصيد مستحق';
    else if (currentBalance < 0) status = 'صرف زائد';

    rows.push([
      v.vendor,
      v.specialization,
      projectsCount,
      v.totalDebitUsd,    // إجمالي المدين (المستحق)
      v.totalCreditUsd,   // إجمالي الدائن (المدفوع)
      currentBalance,
      v.payments,
      v.lastDate ? Utilities.formatDate(v.lastDate, Session.getScriptTimeZone(), 'dd/MM/yyyy') : '',
      status,
      '📄'  // عمود كشف الحساب
    ]);
  });

  const lastCol = reportSheet.getLastColumn();
  if (reportSheet.getMaxRows() > 1) {
    reportSheet.getRange(2, 1, reportSheet.getMaxRows() - 1, lastCol).clearContent();
  }

  if (rows.length) {
    rows.sort((a, b) => a[0].localeCompare(b[0]));
    reportSheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
    reportSheet.getRange(2, 4, rows.length, 3).setNumberFormat('$#,##0.00');
    // تنسيق عمود الكشف
    reportSheet.getRange(2, 10, rows.length, 1).setHorizontalAlignment('center');
  }

  if (silent) return { success: true, name: 'تقرير الموردين' };
  SpreadsheetApp.getUi().alert('✅ تم تحديث "تقرير الموردين" (بالدولار).');
}

// ==================== إعادة بناء تقرير المصروفات ====================

function rebuildExpenseSummaryReport(silent) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const transSheet = ss.getSheetByName(CONFIG.SHEETS.TRANSACTIONS);
  const reportSheet = ss.getSheetByName(CONFIG.SHEETS.EXPENSES_REPORT);
  if (!transSheet || !reportSheet) {
    if (silent) return { success: false, name: 'تقرير المصروفات', error: 'الشيتات غير موجودة' };
    SpreadsheetApp.getUi().alert('⚠️ تأكد من وجود "دفتر الحركات المالية" و "تقرير المصروفات".');
    return;
  }

  const data = transSheet.getDataRange().getValues();
  const map = {};

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const type = row[2];           // C: طبيعة الحركة
    const classification = row[3]; // D: تصنيف الحركة
    const item = row[6];           // G: البند
    const amountUsd = Number(row[12]) || 0; // M: القيمة بالدولار

    if (!item || !amountUsd) continue;
    // استخدام includes للتعامل مع القيم التي تحتوي على إيموجي
    const typeStr = String(type || '');
    if (!typeStr.includes('استحقاق مصروف') && !typeStr.includes('دفعة مصروف')) continue;

    const key = item + '||' + classification;
    if (!map[key]) {
      map[key] = {
        item,
        classification,
        totalAccrual: 0,
        totalPaid: 0,
        accrualCount: 0,
        paymentCount: 0
      };
    }
    const v = map[key];

    if (typeStr.includes('استحقاق مصروف')) {
      v.totalAccrual += amountUsd;
      v.accrualCount++;
    } else if (typeStr.includes('دفعة مصروف')) {
      v.totalPaid += amountUsd;
      v.paymentCount++;
    }
  }

  const rows = [];
  Object.keys(map).forEach(k => {
    const v = map[k];
    const remaining = v.totalAccrual - v.totalPaid;
    const percent = v.totalAccrual ? v.totalPaid / v.totalAccrual : 0;
    rows.push([
      v.item,
      v.classification,
      v.totalAccrual,
      v.totalPaid,
      remaining,
      v.accrualCount,
      v.paymentCount,
      v.totalAccrual ? percent : ''
    ]);
  });

  const lastCol = reportSheet.getLastColumn();
  if (reportSheet.getMaxRows() > 1) {
    reportSheet.getRange(2, 1, reportSheet.getMaxRows() - 1, lastCol).clearContent();
  }

  if (rows.length) {
    rows.sort((a, b) => a[0].localeCompare(b[0]));
    reportSheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
    reportSheet.getRange(2, 3, rows.length, 3).setNumberFormat('$#,##0.00');
    reportSheet.getRange(2, 8, rows.length, 1).setNumberFormat('0.0%');
  }

  if (silent) return { success: true, name: 'تقرير المصروفات' };
  SpreadsheetApp.getUi().alert('✅ تم تحديث "تقرير المصروفات" (بالدولار).');
}

// ==================== إعادة بناء تقرير الإيرادات ====================

function rebuildRevenueSummaryReport(silent) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const transSheet = ss.getSheetByName(CONFIG.SHEETS.TRANSACTIONS);
  const reportSheet = ss.getSheetByName(CONFIG.SHEETS.REVENUE_REPORT);
  const projectsSheet = ss.getSheetByName(CONFIG.SHEETS.PROJECTS);

  if (!transSheet || !reportSheet) {
    if (silent) return { success: false, name: 'تقرير الإيرادات', error: 'الشيتات غير موجودة' };
    SpreadsheetApp.getUi().alert('⚠️ تأكد من وجود "دفتر الحركات المالية" و "تقرير الإيرادات".');
    return;
  }

  // اسم المشروع والقناة من قاعدة المشاريع
  const projectMap = {};
  if (projectsSheet) {
    const pData = projectsSheet.getDataRange().getValues();
    for (let i = 1; i < pData.length; i++) {
      const code = pData[i][0];
      if (code) {
        projectMap[code] = {
          name: pData[i][1],
          channel: pData[i][3]
        };
      }
    }
  }

  const data = transSheet.getDataRange().getValues();
  const map = {}; // key = projectCode

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const type = row[2];       // C: طبيعة الحركة
    // استخدام includes للتعامل مع القيم التي تحتوي على إيموجي
    const typeStr = String(type || '');
    if (!typeStr.includes('استحقاق إيراد') && !typeStr.includes('تحصيل إيراد')) continue;

    const projectCode = row[4];              // E: كود المشروع
    const amountUsd = Number(row[12]) || 0;// M: القيمة بالدولار
    if (!projectCode || !amountUsd) continue;

    if (!map[projectCode]) {
      const info = projectMap[projectCode] || {};
      map[projectCode] = {
        projectCode,
        projectName: info.name || '',
        channel: info.channel || row[8] || '', // I: اسم العميل/القناة لو مش موجود في المشاريع
        expected: 0,
        received: 0,
        lastDate: null
      };
    }

    const v = map[projectCode];
    if (typeStr.includes('استحقاق إيراد')) {
      v.expected += amountUsd;
    }
    if (typeStr.includes('تحصيل إيراد')) {
      v.received += amountUsd;
      const date = row[1];
      if (date) {
        const d = new Date(date);
        if (!v.lastDate || d > v.lastDate) v.lastDate = d;
      }
    }
  }

  const rows = [];
  Object.keys(map).forEach(k => {
    const v = map[k];
    const remaining = v.expected - v.received;
    let status = 'لا يوجد بيانات';
    if (v.expected === 0 && v.received > 0) status = 'مقبوض بدون استحقاق';
    else if (v.expected > 0 && remaining === 0) status = 'مقبوض بالكامل';
    else if (v.expected > 0 && remaining > 0 && v.received > 0) status = 'مقبوض جزئياً';
    else if (v.expected > 0 && v.received === 0) status = 'لم يُقبض بعد';

    rows.push([
      v.projectName || v.projectCode,
      v.channel,
      'إيرادات عقد',
      v.expected,
      v.received,
      remaining,
      v.lastDate ? Utilities.formatDate(v.lastDate, Session.getScriptTimeZone(), 'dd/MM/yyyy') : '',
      status
    ]);
  });

  const lastCol = reportSheet.getLastColumn();
  if (reportSheet.getMaxRows() > 1) {
    reportSheet.getRange(2, 1, reportSheet.getMaxRows() - 1, lastCol).clearContent();
  }

  if (rows.length) {
    rows.sort((a, b) => a[0].localeCompare(b[0]));
    reportSheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
    reportSheet.getRange(2, 4, rows.length, 3).setNumberFormat('$#,##0.00');
  }

  if (silent) return { success: true, name: 'تقرير الإيرادات' };
  SpreadsheetApp.getUi().alert('✅ تم تحديث "تقرير الإيرادات" (بالدولار).');
}

// ==================== إعادة بناء التدفقات النقدية ====================

function rebuildCashFlowReport(silent) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const transSheet = ss.getSheetByName(CONFIG.SHEETS.TRANSACTIONS);
  const reportSheet = ss.getSheetByName(CONFIG.SHEETS.CASHFLOW);
  if (!transSheet || !reportSheet) {
    if (silent) return { success: false, name: 'التدفقات النقدية', error: 'الشيتات غير موجودة' };
    SpreadsheetApp.getUi().alert('⚠️ تأكد من وجود "دفتر الحركات المالية" و "التدفقات النقدية".');
    return;
  }

  const data = transSheet.getDataRange().getValues();
  const map = {}; // key = YYYY-MM

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const date = row[1];
    if (!date) continue;

    const type = row[2];               // C: طبيعة الحركة
    const amountUsd = Number(row[12]) || 0; // M: القيمة بالدولار
    if (!amountUsd) continue;

    // استخدام includes للتعامل مع القيم التي تحتوي على إيموجي
    const typeStr = String(type || '');

    const monthKey = Utilities.formatDate(new Date(date), Session.getScriptTimeZone(), 'yyyy-MM');
    if (!map[monthKey]) {
      map[monthKey] = { monthKey, accruals: 0, payments: 0, revenues: 0 };
    }

    if (typeStr.includes('استحقاق مصروف')) {
      map[monthKey].accruals += amountUsd;
    } else if (typeStr.includes('دفعة مصروف')) {
      map[monthKey].payments += amountUsd;
    } else if (typeStr.includes('تحصيل إيراد')) {
      map[monthKey].revenues += amountUsd;
    }
  }

  const months = Object.keys(map).sort();
  const rows = [];
  let cumulative = 0;
  months.forEach(m => {
    const v = map[m];
    const net = v.revenues - v.payments;
    cumulative += net;
    rows.push([
      m,
      v.accruals,
      v.payments,
      v.revenues,
      net,
      cumulative
    ]);
  });

  const lastCol = reportSheet.getLastColumn();
  if (reportSheet.getMaxRows() > 1) {
    reportSheet.getRange(2, 1, reportSheet.getMaxRows() - 1, lastCol).clearContent();
  }

  if (rows.length) {
    reportSheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
    reportSheet.getRange(2, 2, rows.length, 5).setNumberFormat('$#,##0.00');
  }

  if (silent) return { success: true, name: 'التدفقات النقدية' };
  SpreadsheetApp.getUi().alert('✅ تم تحديث "التدفقات النقدية" (بالدولار).');
}

/**
 * إعادة بناء كل التقارير الملخصة
 * @param {boolean} silent - إذا كان true، تُرجع مصفوفة بالنتائج بدلاً من إظهار رسائل
 * @returns {Array|undefined} - مصفوفة النتائج إذا كان silent = true
 */
function rebuildAllSummaryReports(silent) {
  const results = [];

  // 1️⃣ تحديث البنوك والخزنة والبطاقة أولاً
  results.push(rebuildBankAndCashFromTransactions(true));

  // 2️⃣ تحديث كل التقارير الملخصة
  results.push(rebuildProjectDetailReport(true));
  results.push(rebuildVendorSummaryReport(true));
  results.push(rebuildFunderSummaryReport(true));
  results.push(rebuildExpenseSummaryReport(true));
  results.push(rebuildRevenueSummaryReport(true));
  results.push(rebuildCashFlowReport(true));

  if (silent) return results;

  // إظهار رسالة واحدة شاملة
  const successList = results.filter(r => r && r.success).map(r => '✅ ' + r.name);
  const errorList = results.filter(r => r && !r.success).map(r => '❌ ' + r.name + ': ' + r.error);

  let message = '══════════════════════════════\n';
  message += '     تقرير تحديث البيانات\n';
  message += '══════════════════════════════\n\n';

  if (successList.length) {
    message += '✅ تم بنجاح:\n' + successList.join('\n') + '\n\n';
  }
  if (errorList.length) {
    message += '❌ فشل:\n' + errorList.join('\n');
  }

  SpreadsheetApp.getUi().alert(message);
}

// ==================== إنشاء شيتات التقارير (بدون تغيير كبير) ====================

function createProjectReportSheet(ss) {
  const sheet = getOrCreateSheet_(ss, CONFIG.SHEETS.PROJECT_REPORT);

  const headers = [
    'كود المشروع', 'اسم المشروع', 'البند', 'المورد',
    'إجمالي المستحق', 'المدفوع', 'المتبقي', 'عدد الدفعات', 'حالة السداد (يدوي)'
  ];
  const widths = [120, 180, 150, 150, 130, 130, 130, 100, 130];

  setupSheet_(sheet, headers, widths, CONFIG.COLORS.HEADER.REPORTS);
  sheet.getRange('A1').setNote(
    'هذا تقرير تفصيلي يمكن ملؤه عبر Pivot Table أو عبر نسخ بيانات من دفتر الحركات.'
  );
}

function createVendorReportSheet(ss) {
  const sheet = getOrCreateSheet_(ss, CONFIG.SHEETS.VENDORS_REPORT);

  const headers = [
    'اسم المورد', 'التخصص', 'عدد المشاريع', 'إجمالي المستحقات',
    'إجمالي المدفوع', 'الرصيد الحالي', 'عدد الدفعات', 'آخر تعامل', 'الحالة (يدوي)', '📄 كشف'
  ];
  const widths = [180, 120, 100, 140, 140, 130, 100, 120, 120, 60];

  setupSheet_(sheet, headers, widths, CONFIG.COLORS.HEADER.VENDORS);
  sheet.getRange('A1').setNote(
    'يمكنك إنشاء Pivot Table من "دفتر الحركات المالية" لتعبئة هذا التقرير تلقائياً.'
  );
  sheet.getRange('J1').setNote(
    '📄 اضغط على أي خلية في هذا العمود لإنشاء كشف حساب للمورد'
  );
}

// ========= تقرير الممولين =========
function createFunderReportSheet(ss) {
  const sheet = getOrCreateSheet_(ss, CONFIG.SHEETS.FUNDERS_REPORT);

  const headers = [
    'اسم الممول', 'نوع التمويل', 'عدد المشاريع', 'إجمالي التمويل',
    'إجمالي السداد', 'الرصيد المتبقي', 'عدد الدفعات', 'آخر تعامل', 'الحالة', '📄 كشف'
  ];
  const widths = [180, 120, 100, 140, 140, 130, 100, 120, 120, 60];

  setupSheet_(sheet, headers, widths, CONFIG.COLORS.HEADER.FUNDER);
  sheet.getRange('A1').setNote(
    'تقرير الممولين - يعرض حركات التمويل وسداد التمويل لكل ممول'
  );
  sheet.getRange('J1').setNote(
    '📄 اضغط على أي خلية في هذا العمود لإنشاء كشف حساب للممول'
  );
}

// ========= إعادة بناء تقرير الممولين =========
function rebuildFunderSummaryReport(silent) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const transSheet = ss.getSheetByName(CONFIG.SHEETS.TRANSACTIONS);

  if (!transSheet) {
    if (silent) return { success: false, name: 'تقرير الممولين', error: 'دفتر الحركات غير موجود' };
    SpreadsheetApp.getUi().alert('⚠️ تأكد من وجود "دفتر الحركات المالية".');
    return;
  }

  // إنشاء شيت تقرير الممولين إذا لم يكن موجوداً
  let reportSheet = ss.getSheetByName(CONFIG.SHEETS.FUNDERS_REPORT);
  if (!reportSheet) {
    createFunderReportSheet(ss);
    reportSheet = ss.getSheetByName(CONFIG.SHEETS.FUNDERS_REPORT);
  }

  const data = transSheet.getDataRange().getValues();
  const map = {};

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const funder = row[8];              // I: اسم الممول/الجهة
    const type = row[2];                // C: طبيعة الحركة
    const amountUsd = Number(row[12]) || 0; // M: القيمة بالدولار
    const project = row[4];             // E: كود المشروع
    const date = row[1];                // B: التاريخ
    const classification = row[3];      // D: تصنيف الحركة (نوع التمويل)

    if (!funder || !amountUsd) continue;

    // استخدام includes للتعامل مع القيم التي تحتوي على إيموجي
    const typeStr = String(type || '');
    if (!typeStr.includes('تمويل') && !typeStr.includes('سداد تمويل')) continue;

    if (!map[funder]) {
      map[funder] = {
        funder,
        fundingType: classification || '',
        projects: new Set(),
        totalFundingUsd: 0,
        totalRepaymentUsd: 0,
        payments: 0,
        lastDate: null
      };
    }

    const f = map[funder];
    if (project) f.projects.add(project);
    if (classification && !f.fundingType) f.fundingType = classification;

    if (typeStr.includes('سداد تمويل')) {
      f.totalRepaymentUsd += amountUsd;
      if (amountUsd > 0) f.payments++;
    } else if (typeStr.includes('تمويل')) {
      f.totalFundingUsd += amountUsd;
    }

    if (date) {
      const d = new Date(date);
      if (!f.lastDate || d > f.lastDate) {
        f.lastDate = d;
      }
    }
  }

  const rows = [];
  Object.keys(map).forEach(k => {
    const f = map[k];
    const projectsCount = f.projects.size;
    const balance = f.totalFundingUsd - f.totalRepaymentUsd;

    let status = 'مسدد بالكامل';
    if (balance > 0) status = 'رصيد متبقي';
    else if (balance < 0) status = 'سداد زائد';

    rows.push([
      f.funder,
      f.fundingType,
      projectsCount,
      f.totalFundingUsd,
      f.totalRepaymentUsd,
      balance,
      f.payments,
      f.lastDate ? Utilities.formatDate(f.lastDate, Session.getScriptTimeZone(), 'dd/MM/yyyy') : '',
      status,
      '📄'  // عمود كشف الحساب
    ]);
  });

  const lastCol = reportSheet.getLastColumn();
  if (reportSheet.getMaxRows() > 1) {
    reportSheet.getRange(2, 1, reportSheet.getMaxRows() - 1, lastCol).clearContent();
  }

  if (rows.length) {
    rows.sort((a, b) => a[0].localeCompare(b[0]));
    reportSheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
    reportSheet.getRange(2, 4, rows.length, 3).setNumberFormat('$#,##0.00');
    // تنسيق عمود الكشف
    reportSheet.getRange(2, 10, rows.length, 1).setHorizontalAlignment('center');
  }

  if (silent) return { success: true, name: 'تقرير الممولين' };
  SpreadsheetApp.getUi().alert('✅ تم تحديث "تقرير الممولين" (بالدولار).');
}

// ========= إضافة عمود كشف الحساب لتقرير الممولين الموجود =========
function addStatementColumnToFunderReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.FUNDERS_REPORT);

  if (!sheet) {
    ui.alert('⚠️ لم يتم العثور على شيت "تقرير الممولين"');
    return;
  }

  // التحقق من وجود العمود J
  const lastCol = sheet.getLastColumn();
  if (lastCol < 10) {
    // إضافة رأس العمود J
    sheet.getRange('J1')
      .setValue('📄 كشف')
      .setBackground(CONFIG.COLORS.HEADER.FUNDER)
      .setFontColor(CONFIG.COLORS.TEXT.WHITE)
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
    sheet.setColumnWidth(10, 60);
  }

  // إضافة ملاحظة
  sheet.getRange('J1').setNote(
    '📄 اضغط على أي خلية في هذا العمود لإنشاء كشف حساب للممول'
  );

  // ملء الأيقونات للصفوف الموجودة
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const icons = [];
    for (let i = 2; i <= lastRow; i++) {
      icons.push(['📄']);
    }
    sheet.getRange(2, 10, lastRow - 1, 1)
      .setValues(icons)
      .setHorizontalAlignment('center');
  }

  ui.alert('✅ تم إضافة عمود كشف الحساب لتقرير الممولين');
}

// ========= تقرير المصروفات (يتغذى مباشرة من دفتر الحركات) =========
/**
 * ⚡ تحسينات الأداء:
 * - Batch Operations: 7 API calls بدلاً من 693 (99×7)
 * - نطاقات محددة بدل أعمدة كاملة (G2:G1000 بدل G:G)
 */
function createExpenseReportSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.SHEETS.EXPENSES_REPORT);
  if (!sheet) sheet = ss.insertSheet(CONFIG.SHEETS.EXPENSES_REPORT);
  sheet.clear();

  const headers = [
    'البند', 'التصنيف', 'إجمالي المستحق', 'المدفوع فعلياً', 'المتبقي',
    'عدد الاستحقاقات', 'عدد الدفعات', 'النسبة % من إجمالي المستحقات'
  ];

  sheet.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setBackground(CONFIG.COLORS.HEADER.ITEMS)
    .setFontColor(CONFIG.COLORS.TEXT.WHITE)
    .setFontWeight('bold')
    .setFontSize(11);

  const widths = [180, 150, 150, 150, 130, 120, 120, 180];
  widths.forEach((width, i) => sheet.setColumnWidth(i + 1, width));

  sheet.setFrozenRows(1);

  // قائمة البنود من دفتر الحركات (UNIQUE)
  sheet.getRange('A2').setFormula(
    `=UNIQUE(FILTER('دفتر الحركات المالية'!G2:G1000,'دفتر الحركات المالية'!G2:G1000<>""))`
  );

  // ⚡ Batch Operations - بناء كل المعادلات مرة واحدة
  const numRows = 99;
  const formulas = {
    B: [],  // التصنيف
    C: [],  // إجمالي المستحق
    D: [],  // المدفوع
    E: [],  // المتبقي
    F: [],  // عدد الاستحقاقات
    G: [],  // عدد الدفعات
    H: []   // النسبة
  };

  for (let row = 2; row <= 100; row++) {
    // التصنيف - نطاق محدد بدل عمود كامل
    formulas.B.push([
      `=IF(A${row}="","",IFERROR(INDEX('دفتر الحركات المالية'!D2:D1000,MATCH(A${row},'دفتر الحركات المالية'!G2:G1000,0)),""))`
    ]);

    // إجمالي المستحق (استحقاق مصروف) - نطاقات محددة
    formulas.C.push([
      `=IF(A${row}="","",SUMIFS('دفتر الحركات المالية'!J2:J1000,'دفتر الحركات المالية'!G2:G1000,A${row},'دفتر الحركات المالية'!C2:C1000,"استحقاق مصروف"))`
    ]);

    // المدفوع (دفعة مصروف) - نطاقات محددة
    formulas.D.push([
      `=IF(A${row}="","",SUMIFS('دفتر الحركات المالية'!K2:K1000,'دفتر الحركات المالية'!G2:G1000,A${row},'دفتر الحركات المالية'!C2:C1000,"دفعة مصروف"))`
    ]);

    // المتبقي
    formulas.E.push([`=IF(A${row}="","",C${row}-D${row})`]);

    // عدد الاستحقاقات - نطاقات محددة
    formulas.F.push([
      `=IF(A${row}="","",COUNTIFS('دفتر الحركات المالية'!G2:G1000,A${row},'دفتر الحركات المالية'!C2:C1000,"استحقاق مصروف"))`
    ]);

    // عدد الدفعات - نطاقات محددة
    formulas.G.push([
      `=IF(A${row}="","",COUNTIFS('دفتر الحركات المالية'!G2:G1000,A${row},'دفتر الحركات المالية'!C2:C1000,"دفعة مصروف"))`
    ]);

    // النسبة من إجمالي المستحقات
    formulas.H.push([
      `=IF(A${row}="","",IF(SUM($C$2:$C$100)=0,"",C${row}/SUM($C$2:$C$100)))`
    ]);
  }

  // ⚡ تطبيق كل المعادلات دفعة واحدة (7 API calls بدلاً من 693)
  sheet.getRange(2, 2, numRows, 1).setFormulas(formulas.B);
  sheet.getRange(2, 3, numRows, 1).setFormulas(formulas.C);
  sheet.getRange(2, 4, numRows, 1).setFormulas(formulas.D);
  sheet.getRange(2, 5, numRows, 1).setFormulas(formulas.E);
  sheet.getRange(2, 6, numRows, 1).setFormulas(formulas.F);
  sheet.getRange(2, 7, numRows, 1).setFormulas(formulas.G);
  sheet.getRange(2, 8, numRows, 1).setFormulas(formulas.H);

  sheet.getRange(2, 3, numRows, 3).setNumberFormat('$#,##0.00');
  sheet.getRange(2, 8, numRows, 1).setNumberFormat('0.0%');
}

// ========= تقرير الإيرادات (قالب) =========

function createRevenueReportSheet(ss) {
  const sheet = getOrCreateSheet_(ss, CONFIG.SHEETS.REVENUE_REPORT);

  const headers = [
    'المشروع', 'القناة/الجهة', 'نوع الإيراد', 'المبلغ المستحق',
    'المستلم فعلياً', 'المتبقي', 'تاريخ الاستلام', 'الحالة (يدوي)'
  ];
  const widths = [180, 150, 130, 140, 140, 130, 130, 120];

  setupSheet_(sheet, headers, widths, CONFIG.COLORS.HEADER.REVENUE);
  sheet.getRange('A1').setNote(
    'يمكنك عمل Pivot Table من دفتر الحركات (طبيعة الحركة = استحقاق إيراد / تحصيل إيراد) لملء هذا التقرير.'
  );
}

// ==================== قائمة الدخل (Income Statement) ====================
/**
 * إنشاء شيت قائمة الدخل
 * قائمة الدخل = الإيرادات - المصروفات = صافي الربح
 */
function createIncomeStatementSheet(ss) {
  const sheet = getOrCreateSheet_(ss, CONFIG.SHEETS.INCOME_STATEMENT);

  // تحديد عرض الأعمدة
  sheet.setColumnWidth(1, 250);  // البيان
  sheet.setColumnWidth(2, 150);  // المبلغ
  sheet.setColumnWidth(3, 150);  // الإجمالي

  sheet.setFrozenRows(0);
  return sheet;
}

/**
 * إعادة بناء قائمة الدخل من دفتر الحركات المالية
 * @param {boolean} silent - إذا كان true لا يظهر رسالة تأكيد
 */
function rebuildIncomeStatement(silent) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const transSheet = ss.getSheetByName(CONFIG.SHEETS.TRANSACTIONS);

  if (!transSheet) {
    if (silent) return { success: false, name: 'قائمة الدخل', error: 'دفتر الحركات غير موجود' };
    SpreadsheetApp.getUi().alert('⚠️ تأكد من وجود "دفتر الحركات المالية".');
    return;
  }

  // إنشاء أو الحصول على الشيت
  let reportSheet = ss.getSheetByName(CONFIG.SHEETS.INCOME_STATEMENT);
  if (!reportSheet) {
    reportSheet = createIncomeStatementSheet(ss);
  } else {
    reportSheet.clear();
    // إعادة تعيين عرض الأعمدة
    reportSheet.setColumnWidth(1, 250);
    reportSheet.setColumnWidth(2, 150);
    reportSheet.setColumnWidth(3, 150);
  }

  // قراءة بيانات الحركات
  const data = transSheet.getDataRange().getValues();

  // تجميع الإيرادات والمصروفات
  const revenues = {};      // إيرادات حسب المشروع أو العميل
  const expenses = {};      // مصروفات حسب التصنيف
  let totalRevenue = 0;
  let totalExpense = 0;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const natureType = String(row[2] || '');  // C: طبيعة الحركة
    const classification = row[3] || 'غير مصنف';  // D: تصنيف الحركة
    const projectCode = row[4] || '';  // E: كود المشروع
    const clientName = row[8] || '';   // I: اسم الطرف
    const amountUsd = Number(row[12]) || 0;  // M: القيمة بالدولار

    if (!amountUsd) continue;

    // الإيرادات (استحقاق إيراد)
    if (natureType.includes('استحقاق إيراد')) {
      const key = projectCode || clientName || 'إيرادات أخرى';
      revenues[key] = (revenues[key] || 0) + amountUsd;
      totalRevenue += amountUsd;
    }

    // المصروفات (استحقاق مصروف)
    if (natureType.includes('استحقاق مصروف')) {
      const key = classification || 'مصروفات أخرى';
      expenses[key] = (expenses[key] || 0) + amountUsd;
      totalExpense += amountUsd;
    }
  }

  // صافي الربح
  const netProfit = totalRevenue - totalExpense;

  // بناء بيانات التقرير
  const rows = [];
  let currentRow = 1;

  // ===== عنوان التقرير =====
  rows.push(['قائمة الدخل', '', '']);
  rows.push([Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy'), '', '']);
  rows.push(['', '', '']);

  // ===== قسم الإيرادات =====
  rows.push(['الإيرادات', '', '']);

  // تفاصيل الإيرادات
  const revenueKeys = Object.keys(revenues).sort();
  revenueKeys.forEach(key => {
    rows.push(['    ' + key, revenues[key], '']);
  });

  // إجمالي الإيرادات
  rows.push(['إجمالي الإيرادات', '', totalRevenue]);
  rows.push(['', '', '']);

  // ===== قسم المصروفات =====
  rows.push(['المصروفات', '', '']);

  // تفاصيل المصروفات
  const expenseKeys = Object.keys(expenses).sort();
  expenseKeys.forEach(key => {
    rows.push(['    ' + key, expenses[key], '']);
  });

  // إجمالي المصروفات
  rows.push(['إجمالي المصروفات', '', totalExpense]);
  rows.push(['', '', '']);

  // ===== صافي الربح =====
  rows.push(['صافي الربح / (الخسارة)', '', netProfit]);

  // كتابة البيانات
  if (rows.length > 0) {
    reportSheet.getRange(1, 1, rows.length, 3).setValues(rows);
  }

  // ===== التنسيق =====
  const lastRow = rows.length;

  // تنسيق العنوان
  reportSheet.getRange(1, 1, 1, 3)
    .setFontSize(16)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setBackground(CONFIG.COLORS.HEADER.INCOME_STATEMENT)
    .setFontColor(CONFIG.COLORS.TEXT.WHITE);
  reportSheet.getRange(1, 1, 1, 3).merge();

  // تنسيق التاريخ
  reportSheet.getRange(2, 1, 1, 3)
    .setFontSize(11)
    .setHorizontalAlignment('center')
    .setFontColor(CONFIG.COLORS.TEXT.DARK);
  reportSheet.getRange(2, 1, 1, 3).merge();

  // تنسيق عناوين الأقسام (الإيرادات، المصروفات)
  const sectionRows = [4, 4 + revenueKeys.length + 3]; // صف "الإيرادات" و "المصروفات"
  sectionRows.forEach(row => {
    if (row <= lastRow) {
      reportSheet.getRange(row, 1, 1, 3)
        .setFontWeight('bold')
        .setFontSize(12)
        .setBackground(CONFIG.COLORS.BG.LIGHT_BLUE);
    }
  });

  // تنسيق إجمالي الإيرادات
  const totalRevenueRow = 4 + revenueKeys.length + 1;
  reportSheet.getRange(totalRevenueRow, 1, 1, 3)
    .setFontWeight('bold')
    .setBackground(CONFIG.COLORS.BG.LIGHT_GREEN_3);

  // تنسيق إجمالي المصروفات
  const totalExpenseRow = totalRevenueRow + 3 + expenseKeys.length + 1;
  reportSheet.getRange(totalExpenseRow, 1, 1, 3)
    .setFontWeight('bold')
    .setBackground(CONFIG.COLORS.BG.LIGHT_ORANGE);

  // تنسيق صافي الربح
  reportSheet.getRange(lastRow, 1, 1, 3)
    .setFontWeight('bold')
    .setFontSize(13)
    .setBackground(netProfit >= 0 ? CONFIG.COLORS.BG.LIGHT_GREEN_3 : '#ffcdd2')
    .setFontColor(netProfit >= 0 ? CONFIG.COLORS.TEXT.SUCCESS_DARK : CONFIG.COLORS.TEXT.DANGER);

  // تنسيق الأرقام
  reportSheet.getRange(1, 2, lastRow, 2).setNumberFormat('$#,##0.00');

  // محاذاة
  reportSheet.getRange(1, 2, lastRow, 2).setHorizontalAlignment('left');

  if (silent) return { success: true, name: 'قائمة الدخل' };
  SpreadsheetApp.getUi().alert('✅ تم تحديث "قائمة الدخل".\n\nإجمالي الإيرادات: $' + totalRevenue.toLocaleString() + '\nإجمالي المصروفات: $' + totalExpense.toLocaleString() + '\nصافي الربح: $' + netProfit.toLocaleString());
}

// ==================== المركز المالي (Balance Sheet) ====================
/**
 * إنشاء شيت المركز المالي المبسط
 * المركز المالي = الأصول - الخصوم = حقوق الملكية
 */
function createBalanceSheetSheet(ss) {
  const sheet = getOrCreateSheet_(ss, CONFIG.SHEETS.BALANCE_SHEET);

  // تحديد عرض الأعمدة
  sheet.setColumnWidth(1, 250);  // البيان
  sheet.setColumnWidth(2, 150);  // المبلغ
  sheet.setColumnWidth(3, 150);  // الإجمالي

  sheet.setFrozenRows(0);
  return sheet;
}

/**
 * إعادة بناء المركز المالي من دفتر الحركات المالية
 * @param {boolean} silent - إذا كان true لا يظهر رسالة تأكيد
 */
function rebuildBalanceSheet(silent) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const transSheet = ss.getSheetByName(CONFIG.SHEETS.TRANSACTIONS);

  if (!transSheet) {
    if (silent) return { success: false, name: 'المركز المالي', error: 'دفتر الحركات غير موجود' };
    SpreadsheetApp.getUi().alert('⚠️ تأكد من وجود "دفتر الحركات المالية".');
    return;
  }

  // إنشاء أو الحصول على الشيت
  let reportSheet = ss.getSheetByName(CONFIG.SHEETS.BALANCE_SHEET);
  if (!reportSheet) {
    reportSheet = createBalanceSheetSheet(ss);
  } else {
    reportSheet.clear();
    reportSheet.setColumnWidth(1, 250);
    reportSheet.setColumnWidth(2, 150);
    reportSheet.setColumnWidth(3, 150);
  }

  // قراءة بيانات الحركات
  const data = transSheet.getDataRange().getValues();

  // ===== حساب الأصول والخصوم =====

  // 1. النقدية: من شيتات البنك والخزنة
  let cashUsd = getLastBalanceFromSheet_(ss, CONFIG.SHEETS.BANK_USD);
  let cashTry = getLastBalanceFromSheet_(ss, CONFIG.SHEETS.BANK_TRY);
  let pettyUsd = getLastBalanceFromSheet_(ss, CONFIG.SHEETS.CASH_USD);
  let pettyTry = getLastBalanceFromSheet_(ss, CONFIG.SHEETS.CASH_TRY);
  let cardTry = getLastBalanceFromSheet_(ss, CONFIG.SHEETS.CARD_TRY);

  // 2. الذمم المدينة (مستحق من العملاء) = استحقاق إيراد - تحصيل إيراد
  let totalRevenueAccrual = 0;
  let totalRevenueCollection = 0;

  // 3. الذمم الدائنة (مستحق للموردين) = استحقاق مصروف - دفعة مصروف
  let totalExpenseAccrual = 0;
  let totalExpensePayment = 0;

  // 4. التمويل (القروض) = تمويل - سداد تمويل
  let totalFunding = 0;
  let totalFundingRepayment = 0;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const natureType = String(row[2] || '');  // C: طبيعة الحركة
    const amountUsd = Number(row[12]) || 0;   // M: القيمة بالدولار

    if (!amountUsd) continue;

    // إيرادات
    if (natureType.includes('استحقاق إيراد')) {
      totalRevenueAccrual += amountUsd;
    }
    if (natureType.includes('تحصيل إيراد')) {
      totalRevenueCollection += amountUsd;
    }

    // مصروفات
    if (natureType.includes('استحقاق مصروف')) {
      totalExpenseAccrual += amountUsd;
    }
    if (natureType.includes('دفعة مصروف')) {
      totalExpensePayment += amountUsd;
    }

    // تمويل
    if (natureType.includes('تمويل') &&
      !natureType.includes('سداد تمويل') &&
      !natureType.includes('استلام تمويل')) {
      totalFunding += amountUsd;
    }
    if (natureType.includes('سداد تمويل')) {
      totalFundingRepayment += amountUsd;
    }
  }

  // ===== حساب الإجماليات =====
  const receivables = totalRevenueAccrual - totalRevenueCollection;  // الذمم المدينة
  const payables = totalExpenseAccrual - totalExpensePayment;        // الذمم الدائنة
  const loansPayable = totalFunding - totalFundingRepayment;         // القروض

  const totalCash = cashUsd + pettyUsd;  // إجمالي النقدية بالدولار (TRY يحتاج تحويل)
  const totalAssets = totalCash + receivables;
  const totalLiabilities = payables + loansPayable;
  const equity = totalAssets - totalLiabilities;

  // بناء بيانات التقرير
  const rows = [];

  // ===== عنوان التقرير =====
  rows.push(['المركز المالي (قائمة مبسطة)', '', '']);
  rows.push([Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy'), '', '']);
  rows.push(['', '', '']);

  // ===== قسم الأصول =====
  rows.push(['الأصول', '', '']);
  rows.push(['الأصول المتداولة:', '', '']);
  rows.push(['    النقدية - البنك (دولار)', cashUsd, '']);
  rows.push(['    النقدية - خزنة العهدة (دولار)', pettyUsd, '']);
  if (cashTry !== 0) rows.push(['    البنك (ليرة) - للتحويل', cashTry, '']);
  if (pettyTry !== 0) rows.push(['    خزنة العهدة (ليرة) - للتحويل', pettyTry, '']);
  if (cardTry !== 0) rows.push(['    البطاقة (ليرة) - للتحويل', cardTry, '']);
  rows.push(['    الذمم المدينة (مستحق من العملاء)', receivables, '']);
  rows.push(['إجمالي الأصول', '', totalAssets]);
  rows.push(['', '', '']);

  // ===== قسم الخصوم =====
  rows.push(['الخصوم', '', '']);
  rows.push(['الخصوم المتداولة:', '', '']);
  rows.push(['    الذمم الدائنة (مستحق للموردين)', payables, '']);
  rows.push(['    القروض والتمويل', loansPayable, '']);
  rows.push(['إجمالي الخصوم', '', totalLiabilities]);
  rows.push(['', '', '']);

  // ===== حقوق الملكية =====
  rows.push(['حقوق الملكية', '', '']);
  rows.push(['    صافي الأصول (الأصول - الخصوم)', '', equity]);

  // كتابة البيانات
  if (rows.length > 0) {
    reportSheet.getRange(1, 1, rows.length, 3).setValues(rows);
  }

  // ===== التنسيق =====
  const lastRow = rows.length;

  // تنسيق العنوان
  reportSheet.getRange(1, 1, 1, 3)
    .setFontSize(16)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setBackground(CONFIG.COLORS.HEADER.BALANCE_SHEET)
    .setFontColor(CONFIG.COLORS.TEXT.WHITE);
  reportSheet.getRange(1, 1, 1, 3).merge();

  // تنسيق التاريخ
  reportSheet.getRange(2, 1, 1, 3)
    .setFontSize(11)
    .setHorizontalAlignment('center')
    .setFontColor(CONFIG.COLORS.TEXT.DARK);
  reportSheet.getRange(2, 1, 1, 3).merge();

  // تنسيق عناوين الأقسام (الأصول، الخصوم، حقوق الملكية)
  [4, 4 + 8 + (cashTry !== 0 ? 1 : 0) + (pettyTry !== 0 ? 1 : 0) + (cardTry !== 0 ? 1 : 0), lastRow - 1].forEach(row => {
    if (row <= lastRow && row > 0) {
      reportSheet.getRange(row, 1, 1, 3)
        .setFontWeight('bold')
        .setFontSize(12)
        .setBackground(CONFIG.COLORS.BG.LIGHT_BLUE);
    }
  });

  // تنسيق إجمالي الأصول والخصوم
  for (let r = 1; r <= lastRow; r++) {
    const cellValue = reportSheet.getRange(r, 1).getValue();
    if (cellValue === 'إجمالي الأصول') {
      reportSheet.getRange(r, 1, 1, 3)
        .setFontWeight('bold')
        .setBackground(CONFIG.COLORS.BG.LIGHT_GREEN_3);
    }
    if (cellValue === 'إجمالي الخصوم') {
      reportSheet.getRange(r, 1, 1, 3)
        .setFontWeight('bold')
        .setBackground(CONFIG.COLORS.BG.LIGHT_ORANGE);
    }
  }

  // تنسيق صافي الأصول
  reportSheet.getRange(lastRow, 1, 1, 3)
    .setFontWeight('bold')
    .setFontSize(13)
    .setBackground(equity >= 0 ? CONFIG.COLORS.BG.LIGHT_GREEN_3 : '#ffcdd2')
    .setFontColor(equity >= 0 ? CONFIG.COLORS.TEXT.SUCCESS_DARK : CONFIG.COLORS.TEXT.DANGER);

  // تنسيق الأرقام
  reportSheet.getRange(1, 2, lastRow, 2).setNumberFormat('$#,##0.00');

  // محاذاة
  reportSheet.getRange(1, 2, lastRow, 2).setHorizontalAlignment('left');

  if (silent) return { success: true, name: 'المركز المالي' };
  SpreadsheetApp.getUi().alert('✅ تم تحديث "المركز المالي".\n\nإجمالي الأصول: $' + totalAssets.toLocaleString() + '\nإجمالي الخصوم: $' + totalLiabilities.toLocaleString() + '\nصافي الأصول: $' + equity.toLocaleString());
}

// ==================== شجرة الحسابات (Chart of Accounts) ====================
/**
 * الحسابات الافتراضية لشجرة الحسابات المحاسبية
 * مبنية على المعايير المحاسبية مع تخصيص لطبيعة عمل الأفلام الوثائقية
 */
const DEFAULT_ACCOUNTS = [
  // الأصول (1xxx)
  { code: '1000', name: 'الأصول', type: 'أصول', parent: '', level: 0 },
  { code: '1100', name: 'الأصول المتداولة', type: 'أصول', parent: '1000', level: 1 },
  { code: '1110', name: 'النقدية وما في حكمها', type: 'أصول', parent: '1100', level: 2 },
  { code: '1111', name: 'البنك - دولار', type: 'أصول', parent: '1110', level: 3 },
  { code: '1112', name: 'البنك - ليرة', type: 'أصول', parent: '1110', level: 3 },
  { code: '1113', name: 'خزنة العهدة - دولار', type: 'أصول', parent: '1110', level: 3 },
  { code: '1114', name: 'خزنة العهدة - ليرة', type: 'أصول', parent: '1110', level: 3 },
  { code: '1115', name: 'البطاقة - ليرة', type: 'أصول', parent: '1110', level: 3 },
  { code: '1120', name: 'الذمم المدينة', type: 'أصول', parent: '1100', level: 2 },
  { code: '1121', name: 'ذمم العملاء', type: 'أصول', parent: '1120', level: 3 },
  { code: '1122', name: 'التأمينات المدفوعة', type: 'أصول', parent: '1120', level: 3 },

  // الخصوم (2xxx)
  { code: '2000', name: 'الخصوم', type: 'خصوم', parent: '', level: 0 },
  { code: '2100', name: 'الخصوم المتداولة', type: 'خصوم', parent: '2000', level: 1 },
  { code: '2110', name: 'الذمم الدائنة', type: 'خصوم', parent: '2100', level: 2 },
  { code: '2111', name: 'ذمم الموردين', type: 'خصوم', parent: '2110', level: 3 },
  { code: '2120', name: 'القروض والتمويل', type: 'خصوم', parent: '2100', level: 2 },
  { code: '2121', name: 'قروض الممولين', type: 'خصوم', parent: '2120', level: 3 },

  // حقوق الملكية (3xxx)
  { code: '3000', name: 'حقوق الملكية', type: 'حقوق ملكية', parent: '', level: 0 },
  { code: '3100', name: 'رأس المال', type: 'حقوق ملكية', parent: '3000', level: 1 },
  { code: '3200', name: 'الأرباح المحتجزة', type: 'حقوق ملكية', parent: '3000', level: 1 },

  // الإيرادات (4xxx)
  { code: '4000', name: 'الإيرادات', type: 'إيرادات', parent: '', level: 0 },
  { code: '4100', name: 'إيرادات المشاريع', type: 'إيرادات', parent: '4000', level: 1 },
  { code: '4110', name: 'إيرادات عقود الأفلام', type: 'إيرادات', parent: '4100', level: 2 },
  { code: '4120', name: 'إيرادات حقوق البث', type: 'إيرادات', parent: '4100', level: 2 },
  { code: '4200', name: 'إيرادات أخرى', type: 'إيرادات', parent: '4000', level: 1 },

  // المصروفات (5xxx)
  { code: '5000', name: 'المصروفات', type: 'مصروفات', parent: '', level: 0 },
  { code: '5100', name: 'مصروفات الإنتاج', type: 'مصروفات', parent: '5000', level: 1 },
  { code: '5110', name: 'أجور الفريق', type: 'مصروفات', parent: '5100', level: 2 },
  { code: '5120', name: 'تكاليف التصوير', type: 'مصروفات', parent: '5100', level: 2 },
  { code: '5130', name: 'تكاليف المونتاج', type: 'مصروفات', parent: '5100', level: 2 },
  { code: '5140', name: 'تكاليف السفر', type: 'مصروفات', parent: '5100', level: 2 },
  { code: '5200', name: 'مصروفات إدارية', type: 'مصروفات', parent: '5000', level: 1 },
  { code: '5210', name: 'إيجارات', type: 'مصروفات', parent: '5200', level: 2 },
  { code: '5220', name: 'مصروفات مكتبية', type: 'مصروفات', parent: '5200', level: 2 },
  { code: '5300', name: 'مصروفات تمويلية', type: 'مصروفات', parent: '5000', level: 1 },
  { code: '5310', name: 'عمولات بنكية', type: 'مصروفات', parent: '5300', level: 2 },
  { code: '5320', name: 'فوائد قروض', type: 'مصروفات', parent: '5300', level: 2 }
];

/**
 * إنشاء شيت شجرة الحسابات مع البيانات الافتراضية
 */
function createChartOfAccountsSheet(ss) {
  const sheet = getOrCreateSheet_(ss, CONFIG.SHEETS.CHART_OF_ACCOUNTS);

  const headers = [
    'رقم الحساب',    // A
    'اسم الحساب',    // B
    'نوع الحساب',    // C
    'الحساب الأب',   // D
    'المستوى',       // E
    'الرصيد الحالي', // F
    'ملاحظات'        // G
  ];
  const widths = [120, 200, 120, 120, 80, 150, 200];

  setupSheet_(sheet, headers, widths, CONFIG.COLORS.HEADER.CHART_OF_ACCOUNTS);

  return sheet;
}

/**
 * إعادة بناء شجرة الحسابات أو إنشائها إذا لم تكن موجودة
 * @param {boolean} silent - إذا كان true لا يظهر رسالة تأكيد
 */
function rebuildChartOfAccounts(silent) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // التحقق من وجود الشيت
  let sheet = ss.getSheetByName(CONFIG.SHEETS.CHART_OF_ACCOUNTS);
  const isNew = !sheet;

  if (isNew) {
    sheet = createChartOfAccountsSheet(ss);
  }

  // إذا كان الشيت جديد، نضيف الحسابات الافتراضية
  if (isNew) {
    const rows = DEFAULT_ACCOUNTS.map(acc => [
      acc.code,
      acc.name,
      acc.type,
      acc.parent,
      acc.level,
      0,  // الرصيد الحالي (سيُحسب لاحقاً)
      ''  // ملاحظات
    ]);

    if (rows.length > 0) {
      sheet.getRange(2, 1, rows.length, 7).setValues(rows);

      // تنسيق المسافات حسب المستوى
      for (let i = 0; i < rows.length; i++) {
        const level = DEFAULT_ACCOUNTS[i].level;
        const indent = '    '.repeat(level);
        sheet.getRange(i + 2, 2).setValue(indent + DEFAULT_ACCOUNTS[i].name);
      }

      // تنسيق الأرقام
      sheet.getRange(2, 6, rows.length, 1).setNumberFormat('$#,##0.00');

      // تلوين الصفوف حسب نوع الحساب
      for (let i = 0; i < rows.length; i++) {
        const rowNum = i + 2;
        const level = DEFAULT_ACCOUNTS[i].level;

        if (level === 0) {
          // الحسابات الرئيسية
          sheet.getRange(rowNum, 1, 1, 7)
            .setBackground(CONFIG.COLORS.BG.LIGHT_BLUE)
            .setFontWeight('bold');
        } else if (level === 1) {
          // الحسابات الفرعية المستوى الأول
          sheet.getRange(rowNum, 1, 1, 7)
            .setBackground(CONFIG.COLORS.BG.LIGHT_GREEN_2);
        }
      }
    }
  }

  // تحديث الأرصدة من دفتر الحركات
  updateAccountBalances_(ss, sheet);

  if (silent) return { success: true, name: 'شجرة الحسابات' };

  const msg = isNew ?
    '✅ تم إنشاء "شجرة الحسابات" مع الحسابات الافتراضية.' :
    '✅ تم تحديث أرصدة "شجرة الحسابات".';
  SpreadsheetApp.getUi().alert(msg);
}

/**
 * تحديث أرصدة الحسابات من دفتر الحركات المالية
 */
function updateAccountBalances_(ss, chartSheet) {
  const transSheet = ss.getSheetByName(CONFIG.SHEETS.TRANSACTIONS);
  if (!transSheet) return;

  const transData = transSheet.getDataRange().getValues();
  const chartData = chartSheet.getDataRange().getValues();

  // حساب الأرصدة
  const balances = {
    '1111': 0, // البنك - دولار
    '1112': 0, // البنك - ليرة
    '1113': 0, // خزنة العهدة - دولار
    '1114': 0, // خزنة العهدة - ليرة
    '1115': 0, // البطاقة - ليرة
    '1121': 0, // ذمم العملاء (مستحق من العملاء)
    '1122': 0, // التأمينات المدفوعة
    '2111': 0, // ذمم الموردين (مستحق للموردين)
    '2121': 0, // قروض الممولين
    '4110': 0, // إيرادات عقود الأفلام
    '5100': 0  // مصروفات الإنتاج
  };

  // حساب الأرصدة من شيتات البنك والخزنة
  balances['1111'] = getLastBalanceFromSheet_(ss, CONFIG.SHEETS.BANK_USD);
  balances['1112'] = getLastBalanceFromSheet_(ss, CONFIG.SHEETS.BANK_TRY);
  balances['1113'] = getLastBalanceFromSheet_(ss, CONFIG.SHEETS.CASH_USD);
  balances['1114'] = getLastBalanceFromSheet_(ss, CONFIG.SHEETS.CASH_TRY);
  balances['1115'] = getLastBalanceFromSheet_(ss, CONFIG.SHEETS.CARD_TRY);

  // حساب الذمم من دفتر الحركات
  for (let i = 1; i < transData.length; i++) {
    const row = transData[i];
    const natureType = String(row[2] || '');
    const amountUsd = Number(row[12]) || 0;

    if (!amountUsd) continue;

    // الإيرادات
    if (natureType.includes('استحقاق إيراد')) {
      balances['1121'] += amountUsd;  // ذمم العملاء (مدين)
      balances['4110'] += amountUsd;  // إيرادات (دائن)
    }
    if (natureType.includes('تحصيل إيراد')) {
      balances['1121'] -= amountUsd;  // تخفيض ذمم العملاء
    }

    // المصروفات
    if (natureType.includes('استحقاق مصروف')) {
      balances['2111'] += amountUsd;  // ذمم الموردين (دائن)
      balances['5100'] += amountUsd;  // مصروفات (مدين)
    }
    if (natureType.includes('دفعة مصروف')) {
      balances['2111'] -= amountUsd;  // تخفيض ذمم الموردين
    }

    // التمويل
    if (natureType.includes('تمويل') &&
      !natureType.includes('سداد تمويل') &&
      !natureType.includes('استلام تمويل')) {
      balances['2121'] += amountUsd;  // قروض الممولين (دائن)
    }
    if (natureType.includes('سداد تمويل')) {
      balances['2121'] -= amountUsd;  // تخفيض القروض
    }

    // التأمينات
    if (natureType.includes('تأمين مدفوع')) {
      balances['1122'] += amountUsd;  // تأمينات مدفوعة (مدين)
    }
    if (natureType.includes('استرداد تأمين')) {
      balances['1122'] -= amountUsd;  // تخفيض التأمينات
    }
  }

  // تحديث الأرصدة في الشيت
  for (let i = 1; i < chartData.length; i++) {
    const accountCode = String(chartData[i][0]);
    if (balances[accountCode] !== undefined) {
      chartSheet.getRange(i + 1, 6).setValue(balances[accountCode]);
    }
  }
}

// ==================== دفتر الأستاذ العام (General Ledger) ====================
/**
 * إنشاء شيت دفتر الأستاذ العام
 */
function createGeneralLedgerSheet(ss) {
  const sheet = getOrCreateSheet_(ss, CONFIG.SHEETS.GENERAL_LEDGER);

  const headers = [
    'التاريخ',        // A
    'رقم الحركة',     // B
    'البيان',         // C
    'رقم الحساب',     // D
    'اسم الحساب',     // E
    'مدين',           // F
    'دائن',           // G
    'الرصيد',         // H
    'المرجع'          // I
  ];
  const widths = [100, 100, 250, 100, 180, 120, 120, 130, 120];

  setupSheet_(sheet, headers, widths, CONFIG.COLORS.HEADER.GENERAL_LEDGER);

  return sheet;
}

/**
 * عرض دفتر الأستاذ لحساب معين
 * @param {string} accountCode - رقم الحساب (اختياري، إذا لم يُحدد يُظهر كل الحسابات)
 */
function showGeneralLedger(accountCode) {
  const ui = SpreadsheetApp.getUi();

  // إذا لم يُحدد رقم الحساب، نطلبه من المستخدم
  if (!accountCode) {
    const response = ui.prompt(
      '📒 دفتر الأستاذ العام',
      'أدخل رقم الحساب (مثال: 1111 للبنك دولار)\nأو اتركه فارغاً لعرض كل الحركات:',
      ui.ButtonSet.OK_CANCEL
    );

    if (response.getSelectedButton() !== ui.Button.OK) return;
    accountCode = response.getResponseText().trim();
  }

  rebuildGeneralLedger(false, accountCode);
}

/**
 * إعادة بناء دفتر الأستاذ العام
 * @param {boolean} silent - إذا كان true لا يظهر رسالة تأكيد
 * @param {string} filterAccount - رقم الحساب للتصفية (اختياري)
 */
function rebuildGeneralLedger(silent, filterAccount) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const transSheet = ss.getSheetByName(CONFIG.SHEETS.TRANSACTIONS);

  if (!transSheet) {
    if (silent) return { success: false, name: 'دفتر الأستاذ العام', error: 'دفتر الحركات غير موجود' };
    SpreadsheetApp.getUi().alert('⚠️ تأكد من وجود "دفتر الحركات المالية".');
    return;
  }

  // إنشاء أو الحصول على الشيت
  let ledgerSheet = ss.getSheetByName(CONFIG.SHEETS.GENERAL_LEDGER);
  if (!ledgerSheet) {
    ledgerSheet = createGeneralLedgerSheet(ss);
  } else {
    // مسح المحتوى مع الاحتفاظ بالهيدر
    if (ledgerSheet.getMaxRows() > 1) {
      ledgerSheet.getRange(2, 1, ledgerSheet.getMaxRows() - 1, 9).clearContent();
    }
  }

  // قراءة بيانات الحركات
  const transData = transSheet.getDataRange().getValues();

  // قراءة شجرة الحسابات للحصول على أسماء الحسابات
  const chartSheet = ss.getSheetByName(CONFIG.SHEETS.CHART_OF_ACCOUNTS);
  const accountNames = {};
  if (chartSheet) {
    const chartData = chartSheet.getDataRange().getValues();
    for (let i = 1; i < chartData.length; i++) {
      accountNames[chartData[i][0]] = String(chartData[i][1]).trim();
    }
  }

  // تحويل الحركات إلى قيود محاسبية
  const ledgerEntries = [];

  for (let i = 1; i < transData.length; i++) {
    const row = transData[i];
    const transNum = row[0];           // A: رقم الحركة
    const date = row[1];               // B: التاريخ
    const natureType = String(row[2] || '');  // C: طبيعة الحركة
    const classification = String(row[3] || '');  // D: تصنيف الحركة
    const description = row[7] || '';  // H: الوصف
    const partyName = row[8] || '';    // I: اسم الطرف
    const currency = String(row[10] || '').trim().toUpperCase();  // K: العملة
    const amountUsd = Number(row[12]) || 0;   // M: القيمة بالدولار
    const refNum = row[15] || '';      // P: رقم مرجعي

    if (!amountUsd || !date) continue;

    // تحديد حسابات البنك والخزنة حسب العملة
    const isTRY = currency === 'TRY' || currency === 'ليرة' || currency === 'LIRA';
    const bankAccount = isTRY ? '1112' : '1111';
    const bankName = isTRY ? 'البنك - ليرة' : 'البنك - دولار';
    const cashAccount = isTRY ? '1114' : '1113';
    const cashName = isTRY ? 'خزنة العهدة - ليرة' : 'خزنة العهدة - دولار';

    const fullDescription = partyName ? `${description} - ${partyName}` : (description || classification);
    const formattedDate = date instanceof Date ?
      Utilities.formatDate(date, Session.getScriptTimeZone(), 'dd/MM/yyyy') :
      date;

    // تحديد الحسابات المدينة والدائنة حسب طبيعة الحركة
    let entries = [];

    if (natureType.includes('استحقاق مصروف')) {
      // مصروف: مدين مصروفات، دائن ذمم الموردين
      entries.push({ account: '5100', name: 'مصروفات الإنتاج', debit: amountUsd, credit: 0 });
      entries.push({ account: '2111', name: 'ذمم الموردين', debit: 0, credit: amountUsd });
    }
    else if (natureType.includes('دفعة مصروف')) {
      // دفع للمورد: مدين ذمم الموردين، دائن النقدية
      entries.push({ account: '2111', name: 'ذمم الموردين', debit: amountUsd, credit: 0 });
      entries.push({ account: bankAccount, name: bankName, debit: 0, credit: amountUsd });
    }
    else if (natureType.includes('استحقاق إيراد')) {
      // إيراد: مدين ذمم العملاء، دائن الإيرادات
      entries.push({ account: '1121', name: 'ذمم العملاء', debit: amountUsd, credit: 0 });
      entries.push({ account: '4110', name: 'إيرادات عقود الأفلام', debit: 0, credit: amountUsd });
    }
    else if (natureType.includes('تحصيل إيراد')) {
      // تحصيل من عميل: مدين النقدية، دائن ذمم العملاء
      entries.push({ account: bankAccount, name: bankName, debit: amountUsd, credit: 0 });
      entries.push({ account: '1121', name: 'ذمم العملاء', debit: 0, credit: amountUsd });
    }
    else if (natureType.includes('تمويل') &&
      !natureType.includes('سداد تمويل') &&
      !natureType.includes('استلام تمويل')) {
      // تمويل (قرض): مدين النقدية، دائن القروض
      entries.push({ account: bankAccount, name: bankName, debit: amountUsd, credit: 0 });
      entries.push({ account: '2121', name: 'قروض الممولين', debit: 0, credit: amountUsd });
    }
    else if (natureType.includes('سداد تمويل')) {
      // سداد قرض: مدين القروض، دائن النقدية
      entries.push({ account: '2121', name: 'قروض الممولين', debit: amountUsd, credit: 0 });
      entries.push({ account: bankAccount, name: bankName, debit: 0, credit: amountUsd });
    }
    else if (natureType.includes('تأمين مدفوع')) {
      // تأمين مدفوع: مدين التأمينات، دائن النقدية
      entries.push({ account: '1122', name: 'التأمينات المدفوعة', debit: amountUsd, credit: 0 });
      entries.push({ account: bankAccount, name: bankName, debit: 0, credit: amountUsd });
    }
    else if (natureType.includes('استرداد تأمين')) {
      // استرداد تأمين: مدين النقدية، دائن التأمينات
      entries.push({ account: bankAccount, name: bankName, debit: amountUsd, credit: 0 });
      entries.push({ account: '1122', name: 'التأمينات المدفوعة', debit: 0, credit: amountUsd });
    }
    // ═══════════════════════════════════════════════════════════
    // 🔄 التحويلات الداخلية (بين البنك والخزنة)
    // ═══════════════════════════════════════════════════════════
    else if (natureType.includes('تحويل داخلي')) {
      const isTransferToCash = classification.includes('تحويل للخزنة') || classification.includes('تحويل للكاش');
      const isTransferToBank = classification.includes('تحويل للبنك');

      if (isTransferToCash) {
        // تحويل للخزنة = من البنك إلى الخزنة
        entries.push({ account: cashAccount, name: cashName, debit: amountUsd, credit: 0 });
        entries.push({ account: bankAccount, name: bankName, debit: 0, credit: amountUsd });
      } else if (isTransferToBank) {
        // تحويل للبنك = من الخزنة إلى البنك
        entries.push({ account: bankAccount, name: bankName, debit: amountUsd, credit: 0 });
        entries.push({ account: cashAccount, name: cashName, debit: 0, credit: amountUsd });
      } else {
        // تحويل داخلي غير محدد - افتراض تحويل للخزنة
        entries.push({ account: cashAccount, name: cashName, debit: amountUsd, credit: 0 });
        entries.push({ account: bankAccount, name: bankName, debit: 0, credit: amountUsd });
      }
    }

    // إضافة القيود
    entries.forEach(entry => {
      // التصفية حسب الحساب إذا تم تحديده
      if (filterAccount && entry.account !== filterAccount) return;

      ledgerEntries.push({
        date: formattedDate,
        transNum: transNum,
        description: fullDescription,
        accountCode: entry.account,
        accountName: accountNames[entry.account] || entry.name,
        debit: entry.debit,
        credit: entry.credit,
        ref: refNum
      });
    });
  }

  // ترتيب حسب التاريخ ثم رقم الحساب
  ledgerEntries.sort((a, b) => {
    if (a.accountCode !== b.accountCode) {
      return a.accountCode.localeCompare(b.accountCode);
    }
    return String(a.date).localeCompare(String(b.date));
  });

  // حساب الرصيد التراكمي لكل حساب
  const accountBalances = {};
  const rows = [];

  ledgerEntries.forEach(entry => {
    if (!accountBalances[entry.accountCode]) {
      accountBalances[entry.accountCode] = 0;
    }

    // الحسابات المدينة بطبيعتها (أصول، مصروفات): الرصيد = مدين - دائن
    // الحسابات الدائنة بطبيعتها (خصوم، إيرادات، حقوق ملكية): الرصيد = دائن - مدين
    const isDebitNature = entry.accountCode.startsWith('1') || entry.accountCode.startsWith('5');

    if (isDebitNature) {
      accountBalances[entry.accountCode] += entry.debit - entry.credit;
    } else {
      accountBalances[entry.accountCode] += entry.credit - entry.debit;
    }

    rows.push([
      entry.date,
      entry.transNum,
      entry.description,
      entry.accountCode,
      entry.accountName,
      entry.debit || '',
      entry.credit || '',
      accountBalances[entry.accountCode],
      entry.ref
    ]);
  });

  // كتابة البيانات
  if (rows.length > 0) {
    ledgerSheet.getRange(2, 1, rows.length, 9).setValues(rows);

    // تنسيق الأرقام
    ledgerSheet.getRange(2, 6, rows.length, 3).setNumberFormat('$#,##0.00');

    // تلوين بديل للصفوف
    let currentAccount = '';
    let colorToggle = false;

    for (let i = 0; i < rows.length; i++) {
      const rowNum = i + 2;
      const accountCode = rows[i][3];

      if (accountCode !== currentAccount) {
        currentAccount = accountCode;
        colorToggle = !colorToggle;
        // إضافة فاصل بصري عند تغيير الحساب
        ledgerSheet.getRange(rowNum, 1, 1, 9)
          .setBackground(colorToggle ? CONFIG.COLORS.BG.LIGHT_BLUE : CONFIG.COLORS.BG.WHITE);
      } else {
        ledgerSheet.getRange(rowNum, 1, 1, 9)
          .setBackground(colorToggle ? CONFIG.COLORS.BG.LIGHT_BLUE : CONFIG.COLORS.BG.WHITE);
      }
    }
  }

  if (silent) return { success: true, name: 'دفتر الأستاذ العام' };

  const filterMsg = filterAccount ? ` (حساب ${filterAccount})` : '';
  SpreadsheetApp.getUi().alert(`✅ تم تحديث "دفتر الأستاذ العام"${filterMsg}.\n\nعدد القيود: ${rows.length}`);
}

// ==================== ميزان المراجعة (Trial Balance) ====================
/**
 * إنشاء شيت ميزان المراجعة
 */
function createTrialBalanceSheet(ss) {
  const sheet = getOrCreateSheet_(ss, CONFIG.SHEETS.TRIAL_BALANCE);

  const headers = [
    'رقم الحساب',    // A
    'اسم الحساب',    // B
    'نوع الحساب',    // C
    'مدين',          // D
    'دائن',          // E
    'الرصيد'         // F
  ];
  const widths = [120, 220, 120, 140, 140, 140];

  setupSheet_(sheet, headers, widths, CONFIG.COLORS.HEADER.TRIAL_BALANCE);

  return sheet;
}

/**
 * إعادة بناء ميزان المراجعة
 * ميزان المراجعة يعرض أرصدة جميع الحسابات للتحقق من توازن المدين والدائن
 * @param {boolean} silent - إذا كان true لا يظهر رسالة تأكيد
 */
function rebuildTrialBalance(silent) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const transSheet = ss.getSheetByName(CONFIG.SHEETS.TRANSACTIONS);

  if (!transSheet) {
    if (silent) return { success: false, name: 'ميزان المراجعة', error: 'دفتر الحركات غير موجود' };
    SpreadsheetApp.getUi().alert('⚠️ تأكد من وجود "دفتر الحركات المالية".');
    return;
  }

  // إنشاء أو الحصول على الشيت
  let trialSheet = ss.getSheetByName(CONFIG.SHEETS.TRIAL_BALANCE);
  if (!trialSheet) {
    trialSheet = createTrialBalanceSheet(ss);
  } else {
    trialSheet.clear();
    // إعادة إنشاء الهيدر
    const headers = ['رقم الحساب', 'اسم الحساب', 'نوع الحساب', 'مدين', 'دائن', 'الرصيد'];
    const widths = [120, 220, 120, 140, 140, 140];
    setupSheet_(trialSheet, headers, widths, CONFIG.COLORS.HEADER.TRIAL_BALANCE);
  }

  // قراءة بيانات الحركات
  const transData = transSheet.getDataRange().getValues();

  // تجميع الأرصدة لكل حساب
  const accountBalances = {};

  // تهيئة الحسابات من شجرة الحسابات
  DEFAULT_ACCOUNTS.forEach(acc => {
    accountBalances[acc.code] = {
      code: acc.code,
      name: acc.name,
      type: acc.type,
      debit: 0,
      credit: 0
    };
  });

  // إضافة أرصدة النقدية من شيتات البنك والخزنة
  const cashBalances = {
    '1111': getLastBalanceFromSheet_(ss, CONFIG.SHEETS.BANK_USD),
    '1112': getLastBalanceFromSheet_(ss, CONFIG.SHEETS.BANK_TRY),
    '1113': getLastBalanceFromSheet_(ss, CONFIG.SHEETS.CASH_USD),
    '1114': getLastBalanceFromSheet_(ss, CONFIG.SHEETS.CASH_TRY),
    '1115': getLastBalanceFromSheet_(ss, CONFIG.SHEETS.CARD_TRY)
  };

  // إضافة أرصدة النقدية كمدين (لأنها أصول)
  Object.keys(cashBalances).forEach(code => {
    if (accountBalances[code] && cashBalances[code] > 0) {
      accountBalances[code].debit = cashBalances[code];
    }
  });

  // حساب الأرصدة من دفتر الحركات
  for (let i = 1; i < transData.length; i++) {
    const row = transData[i];
    const natureType = String(row[2] || '');
    const amountUsd = Number(row[12]) || 0;

    if (!amountUsd) continue;

    // استحقاق مصروف: مدين مصروفات، دائن ذمم الموردين
    if (natureType.includes('استحقاق مصروف')) {
      accountBalances['5100'].debit += amountUsd;
      accountBalances['2111'].credit += amountUsd;
    }
    // دفعة مصروف: مدين ذمم الموردين، دائن النقدية
    else if (natureType.includes('دفعة مصروف')) {
      accountBalances['2111'].debit += amountUsd;
      // النقدية تُخصم (لكن أرصدتها من شيتات البنك)
    }
    // استحقاق إيراد: مدين ذمم العملاء، دائن الإيرادات
    else if (natureType.includes('استحقاق إيراد')) {
      accountBalances['1121'].debit += amountUsd;
      accountBalances['4110'].credit += amountUsd;
    }
    // تحصيل إيراد: مدين النقدية، دائن ذمم العملاء
    else if (natureType.includes('تحصيل إيراد')) {
      accountBalances['1121'].credit += amountUsd;
    }
    // تمويل: مدين النقدية، دائن القروض
    else if (natureType.includes('تمويل') &&
      !natureType.includes('سداد تمويل') &&
      !natureType.includes('استلام تمويل')) {
      accountBalances['2121'].credit += amountUsd;
    }
    // سداد تمويل: مدين القروض، دائن النقدية
    else if (natureType.includes('سداد تمويل')) {
      accountBalances['2121'].debit += amountUsd;
    }
    // تأمين مدفوع
    else if (natureType.includes('تأمين مدفوع')) {
      accountBalances['1122'].debit += amountUsd;
    }
    // استرداد تأمين
    else if (natureType.includes('استرداد تأمين')) {
      accountBalances['1122'].credit += amountUsd;
    }
  }

  // بناء صفوف ميزان المراجعة
  const rows = [];
  let totalDebit = 0;
  let totalCredit = 0;

  // فقط الحسابات التي لها أرصدة
  Object.keys(accountBalances).sort().forEach(code => {
    const acc = accountBalances[code];
    const netDebit = acc.debit - acc.credit;
    const netCredit = acc.credit - acc.debit;

    // تخطي الحسابات بدون أرصدة
    if (acc.debit === 0 && acc.credit === 0) return;

    // تحديد الجانب الصحيح حسب طبيعة الحساب
    let displayDebit = 0;
    let displayCredit = 0;
    let balance = 0;

    // الأصول والمصروفات: طبيعتها مدينة
    if (code.startsWith('1') || code.startsWith('5')) {
      if (netDebit > 0) {
        displayDebit = netDebit;
        balance = netDebit;
      } else if (netCredit > 0) {
        displayCredit = netCredit;
        balance = -netCredit;
      }
    }
    // الخصوم والإيرادات وحقوق الملكية: طبيعتها دائنة
    else {
      if (netCredit > 0) {
        displayCredit = netCredit;
        balance = netCredit;
      } else if (netDebit > 0) {
        displayDebit = netDebit;
        balance = -netDebit;
      }
    }

    if (displayDebit > 0 || displayCredit > 0) {
      rows.push([
        acc.code,
        acc.name,
        acc.type,
        displayDebit || '',
        displayCredit || '',
        balance
      ]);

      totalDebit += displayDebit;
      totalCredit += displayCredit;
    }
  });

  // إضافة صف الإجماليات
  rows.push(['', '', '', '', '', '']);
  rows.push(['', 'الإجمالي', '', totalDebit, totalCredit, totalDebit - totalCredit]);

  // كتابة البيانات
  if (rows.length > 0) {
    trialSheet.getRange(2, 1, rows.length, 6).setValues(rows);

    // تنسيق الأرقام
    trialSheet.getRange(2, 4, rows.length, 3).setNumberFormat('$#,##0.00');

    // تنسيق صف الإجماليات
    const totalRow = rows.length + 1;
    trialSheet.getRange(totalRow, 1, 1, 6)
      .setFontWeight('bold')
      .setBackground(CONFIG.COLORS.BG.LIGHT_YELLOW);

    // التحقق من التوازن
    const isBalanced = Math.abs(totalDebit - totalCredit) < 0.01;
    trialSheet.getRange(totalRow, 6)
      .setBackground(isBalanced ? CONFIG.COLORS.BG.LIGHT_GREEN_3 : '#ffcdd2')
      .setFontColor(isBalanced ? CONFIG.COLORS.TEXT.SUCCESS_DARK : CONFIG.COLORS.TEXT.DANGER);
  }

  if (silent) return { success: true, name: 'ميزان المراجعة' };

  const isBalanced = Math.abs(totalDebit - totalCredit) < 0.01;
  const balanceStatus = isBalanced ? '✅ متوازن' : '⚠️ غير متوازن';
  SpreadsheetApp.getUi().alert(
    `✅ تم تحديث "ميزان المراجعة".\n\n` +
    `إجمالي المدين: $${totalDebit.toLocaleString()}\n` +
    `إجمالي الدائن: $${totalCredit.toLocaleString()}\n` +
    `الفرق: $${(totalDebit - totalCredit).toLocaleString()}\n\n` +
    `الحالة: ${balanceStatus}`
  );
}

// ==================== قيود اليومية (Journal Entries) ====================
/**
 * إنشاء شيت قيود اليومية
 */
function createJournalEntriesSheet(ss) {
  const sheet = getOrCreateSheet_(ss, CONFIG.SHEETS.JOURNAL_ENTRIES);

  const headers = [
    'رقم القيد',      // A
    'التاريخ',        // B
    'البيان',         // C
    'رقم الحساب',     // D
    'اسم الحساب',     // E
    'مدين',           // F
    'دائن',           // G
    'المرجع'          // H
  ];
  const widths = [80, 100, 250, 100, 180, 120, 120, 100];

  setupSheet_(sheet, headers, widths, CONFIG.COLORS.HEADER.JOURNAL_ENTRIES);

  return sheet;
}

/**
 * إعادة بناء قيود اليومية
 * يعرض كل الحركات كقيود محاسبية مزدوجة مرتبة بالتاريخ
 * @param {boolean} silent - إذا كان true لا يظهر رسالة تأكيد
 */
function rebuildJournalEntries(silent) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const transSheet = ss.getSheetByName(CONFIG.SHEETS.TRANSACTIONS);

  if (!transSheet) {
    if (silent) return { success: false, name: 'قيود اليومية', error: 'دفتر الحركات غير موجود' };
    SpreadsheetApp.getUi().alert('⚠️ تأكد من وجود "دفتر الحركات المالية".');
    return;
  }

  // إنشاء أو الحصول على الشيت
  let journalSheet = ss.getSheetByName(CONFIG.SHEETS.JOURNAL_ENTRIES);
  if (!journalSheet) {
    journalSheet = createJournalEntriesSheet(ss);
  } else {
    if (journalSheet.getMaxRows() > 1) {
      journalSheet.getRange(2, 1, journalSheet.getMaxRows() - 1, 8).clearContent();
    }
  }

  // قراءة شجرة الحسابات
  const chartSheet = ss.getSheetByName(CONFIG.SHEETS.CHART_OF_ACCOUNTS);
  const accountNames = {};
  if (chartSheet) {
    const chartData = chartSheet.getDataRange().getValues();
    for (let i = 1; i < chartData.length; i++) {
      accountNames[chartData[i][0]] = String(chartData[i][1]).trim();
    }
  }

  // قراءة بيانات الحركات
  const transData = transSheet.getDataRange().getValues();
  const journalEntries = [];
  let entryNumber = 1;

  for (let i = 1; i < transData.length; i++) {
    const row = transData[i];
    const transNum = row[0];
    const date = row[1];
    const natureType = String(row[2] || '');
    const classification = String(row[3] || '');  // D: تصنيف الحركة
    const description = row[7] || '';
    const partyName = row[8] || '';
    const currency = String(row[10] || '').trim().toUpperCase();  // K: العملة
    const amountUsd = Number(row[12]) || 0;
    const refNum = row[15] || '';

    if (!amountUsd || !date) continue;

    // تحديد حسابات البنك والخزنة حسب العملة
    const isTRY = currency === 'TRY' || currency === 'ليرة' || currency === 'LIRA';
    const bankAccount = isTRY ? '1112' : '1111';
    const bankName = isTRY ? 'البنك - ليرة' : 'البنك - دولار';
    const cashAccount = isTRY ? '1114' : '1113';
    const cashName = isTRY ? 'خزنة العهدة - ليرة' : 'خزنة العهدة - دولار';

    const fullDescription = partyName ? `${description} - ${partyName}` : (description || classification);
    const formattedDate = date instanceof Date ?
      Utilities.formatDate(date, Session.getScriptTimeZone(), 'dd/MM/yyyy') :
      date;

    // تحديد القيد حسب طبيعة الحركة
    let entries = [];

    if (natureType.includes('استحقاق مصروف')) {
      entries.push({ account: '5100', name: 'مصروفات الإنتاج', debit: amountUsd, credit: 0 });
      entries.push({ account: '2111', name: 'ذمم الموردين', debit: 0, credit: amountUsd });
    }
    else if (natureType.includes('دفعة مصروف')) {
      entries.push({ account: '2111', name: 'ذمم الموردين', debit: amountUsd, credit: 0 });
      entries.push({ account: bankAccount, name: bankName, debit: 0, credit: amountUsd });
    }
    else if (natureType.includes('استحقاق إيراد')) {
      entries.push({ account: '1121', name: 'ذمم العملاء', debit: amountUsd, credit: 0 });
      entries.push({ account: '4110', name: 'إيرادات عقود الأفلام', debit: 0, credit: amountUsd });
    }
    else if (natureType.includes('تحصيل إيراد')) {
      entries.push({ account: bankAccount, name: bankName, debit: amountUsd, credit: 0 });
      entries.push({ account: '1121', name: 'ذمم العملاء', debit: 0, credit: amountUsd });
    }
    else if (natureType.includes('تمويل') &&
      !natureType.includes('سداد تمويل') &&
      !natureType.includes('استلام تمويل')) {
      entries.push({ account: bankAccount, name: bankName, debit: amountUsd, credit: 0 });
      entries.push({ account: '2121', name: 'قروض الممولين', debit: 0, credit: amountUsd });
    }
    else if (natureType.includes('سداد تمويل')) {
      entries.push({ account: '2121', name: 'قروض الممولين', debit: amountUsd, credit: 0 });
      entries.push({ account: bankAccount, name: bankName, debit: 0, credit: amountUsd });
    }
    else if (natureType.includes('تأمين مدفوع')) {
      entries.push({ account: '1122', name: 'التأمينات المدفوعة', debit: amountUsd, credit: 0 });
      entries.push({ account: bankAccount, name: bankName, debit: 0, credit: amountUsd });
    }
    else if (natureType.includes('استرداد تأمين')) {
      entries.push({ account: bankAccount, name: bankName, debit: amountUsd, credit: 0 });
      entries.push({ account: '1122', name: 'التأمينات المدفوعة', debit: 0, credit: amountUsd });
    }
    // ═══════════════════════════════════════════════════════════
    // 🔄 التحويلات الداخلية (بين البنك والخزنة)
    // ═══════════════════════════════════════════════════════════
    else if (natureType.includes('تحويل داخلي')) {
      const isTransferToCash = classification.includes('تحويل للخزنة') || classification.includes('تحويل للكاش');
      const isTransferToBank = classification.includes('تحويل للبنك');

      if (isTransferToCash) {
        // تحويل للخزنة = من البنك إلى الخزنة
        entries.push({ account: cashAccount, name: cashName, debit: amountUsd, credit: 0 });
        entries.push({ account: bankAccount, name: bankName, debit: 0, credit: amountUsd });
      } else if (isTransferToBank) {
        // تحويل للبنك = من الخزنة إلى البنك
        entries.push({ account: bankAccount, name: bankName, debit: amountUsd, credit: 0 });
        entries.push({ account: cashAccount, name: cashName, debit: 0, credit: amountUsd });
      } else {
        // تحويل داخلي غير محدد - افتراض تحويل للخزنة
        entries.push({ account: cashAccount, name: cashName, debit: amountUsd, credit: 0 });
        entries.push({ account: bankAccount, name: bankName, debit: 0, credit: amountUsd });
      }
    }

    if (entries.length > 0) {
      entries.forEach((entry, idx) => {
        journalEntries.push({
          entryNum: idx === 0 ? entryNumber : '',
          date: idx === 0 ? formattedDate : '',
          description: idx === 0 ? fullDescription : '',
          accountCode: entry.account,
          accountName: accountNames[entry.account] || entry.name,
          debit: entry.debit,
          credit: entry.credit,
          ref: idx === 0 ? refNum : ''
        });
      });
      // إضافة صف فاصل بين القيود
      journalEntries.push({
        entryNum: '', date: '', description: '',
        accountCode: '', accountName: '', debit: '', credit: '', ref: ''
      });
      entryNumber++;
    }
  }

  // بناء الصفوف
  const rows = journalEntries.map(e => [
    e.entryNum,
    e.date,
    e.description,
    e.accountCode,
    e.accountName,
    e.debit || '',
    e.credit || '',
    e.ref
  ]);

  // كتابة البيانات
  if (rows.length > 0) {
    journalSheet.getRange(2, 1, rows.length, 8).setValues(rows);

    // تنسيق الأرقام
    journalSheet.getRange(2, 6, rows.length, 2).setNumberFormat('$#,##0.00');

    // تلوين صفوف القيود
    let colorToggle = false;
    for (let i = 0; i < rows.length; i++) {
      const rowNum = i + 2;
      if (rows[i][0] !== '') {
        colorToggle = !colorToggle;
      }
      if (rows[i][3] !== '') {
        journalSheet.getRange(rowNum, 1, 1, 8)
          .setBackground(colorToggle ? CONFIG.COLORS.BG.LIGHT_BLUE : CONFIG.COLORS.BG.WHITE);
      }
    }
  }

  if (silent) return { success: true, name: 'قيود اليومية' };

  SpreadsheetApp.getUi().alert(`✅ تم تحديث "قيود اليومية".\n\nعدد القيود: ${entryNumber - 1}`);
}

// ==================== الإقفال السنوي (Year-End Closing) ====================
/**
 * تنفيذ الإقفال السنوي
 * - إقفال حسابات الإيرادات والمصروفات
 * - ترحيل صافي الربح/الخسارة لحساب الأرباح المحتجزة
 */
function performYearEndClosing() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // تأكيد من المستخدم
  const response = ui.alert(
    '🔒 الإقفال السنوي',
    '⚠️ تحذير: الإقفال السنوي عملية مهمة!\n\n' +
    'سيتم:\n' +
    '1. حساب صافي الربح/الخسارة من الإيرادات والمصروفات\n' +
    '2. إنشاء قيد إقفال لترحيل النتيجة للأرباح المحتجزة\n' +
    '3. عرض تقرير بالنتائج\n\n' +
    'هل تريد المتابعة؟',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  // طلب السنة المراد إقفالها
  const yearResponse = ui.prompt(
    '📅 سنة الإقفال',
    'أدخل السنة المراد إقفالها (مثال: 2025):',
    ui.ButtonSet.OK_CANCEL
  );

  if (yearResponse.getSelectedButton() !== ui.Button.OK) return;

  const closingYear = parseInt(yearResponse.getResponseText().trim());
  if (isNaN(closingYear) || closingYear < 2000 || closingYear > 2100) {
    ui.alert('⚠️ خطأ', 'الرجاء إدخال سنة صحيحة (مثال: 2025)', ui.ButtonSet.OK);
    return;
  }

  // قراءة بيانات الحركات
  const transSheet = ss.getSheetByName(CONFIG.SHEETS.TRANSACTIONS);
  if (!transSheet) {
    ui.alert('⚠️ خطأ', 'دفتر الحركات المالية غير موجود!', ui.ButtonSet.OK);
    return;
  }

  const transData = transSheet.getDataRange().getValues();

  // حساب الإيرادات والمصروفات للسنة المحددة
  let totalRevenue = 0;
  let totalExpenses = 0;

  for (let i = 1; i < transData.length; i++) {
    const row = transData[i];
    const date = row[1];
    const natureType = String(row[2] || '');
    const amountUsd = Number(row[12]) || 0;

    if (!date || !amountUsd) continue;

    // التحقق من السنة
    const transDate = new Date(date);
    if (transDate.getFullYear() !== closingYear) continue;

    if (natureType.includes('استحقاق إيراد')) {
      totalRevenue += amountUsd;
    }
    if (natureType.includes('استحقاق مصروف')) {
      totalExpenses += amountUsd;
    }
  }

  const netProfit = totalRevenue - totalExpenses;
  const profitOrLoss = netProfit >= 0 ? 'ربح' : 'خسارة';

  // عرض النتائج
  const closingReport =
    `📊 تقرير الإقفال السنوي ${closingYear}\n` +
    `${'═'.repeat(35)}\n\n` +
    `إجمالي الإيرادات: $${totalRevenue.toLocaleString()}\n` +
    `إجمالي المصروفات: $${totalExpenses.toLocaleString()}\n` +
    `${'─'.repeat(35)}\n` +
    `صافي ال${profitOrLoss}: $${Math.abs(netProfit).toLocaleString()}\n\n` +
    `${'═'.repeat(35)}\n` +
    `قيد الإقفال المقترح:\n` +
    `${'─'.repeat(35)}\n`;

  let closingEntry = '';
  if (netProfit >= 0) {
    closingEntry =
      `من حـ/ ملخص الدخل    $${netProfit.toLocaleString()}\n` +
      `    إلى حـ/ الأرباح المحتجزة    $${netProfit.toLocaleString()}\n`;
  } else {
    closingEntry =
      `من حـ/ الأرباح المحتجزة    $${Math.abs(netProfit).toLocaleString()}\n` +
      `    إلى حـ/ ملخص الدخل    $${Math.abs(netProfit).toLocaleString()}\n`;
  }

  ui.alert('📋 نتائج الإقفال السنوي', closingReport + closingEntry, ui.ButtonSet.OK);

  // سؤال عن تحديث الأرباح المحتجزة في شجرة الحسابات
  const updateResponse = ui.alert(
    '💾 تحديث شجرة الحسابات',
    `هل تريد تحديث رصيد "الأرباح المحتجزة" في شجرة الحسابات؟\n\n` +
    `سيتم إضافة صافي ال${profitOrLoss} ($${Math.abs(netProfit).toLocaleString()}) للرصيد الحالي.`,
    ui.ButtonSet.YES_NO
  );

  if (updateResponse === ui.Button.YES) {
    // تحديث رصيد الأرباح المحتجزة
    const chartSheet = ss.getSheetByName(CONFIG.SHEETS.CHART_OF_ACCOUNTS);
    if (chartSheet) {
      const chartData = chartSheet.getDataRange().getValues();
      for (let i = 1; i < chartData.length; i++) {
        if (chartData[i][0] === '3200') { // حساب الأرباح المحتجزة
          const currentBalance = Number(chartData[i][5]) || 0;
          const newBalance = currentBalance + netProfit;
          chartSheet.getRange(i + 1, 6).setValue(newBalance);
          chartSheet.getRange(i + 1, 7).setValue(`إقفال سنة ${closingYear}`);
          break;
        }
      }
      ui.alert('✅ تم', `تم تحديث رصيد الأرباح المحتجزة.\n\nالرصيد الجديد: $${(netProfit).toLocaleString()}`, ui.ButtonSet.OK);
    }
  }
}

// ========= التدفقات النقدية (تلقائي مع ترتيب الأعمدة الجديد) =========
/**
 * ⚡ تحسينات الأداء:
 * - Batch Operations: 5 API calls بدلاً من 495 (99×5)
 * - نطاقات محددة بدل أعمدة كاملة (T2:T1000 بدل T:T)
 */
function createCashFlowSheet(ss) {
  const sheet = getOrCreateSheet_(ss, CONFIG.SHEETS.CASHFLOW);

  const headers = [
    'الشهر (YYYY-MM)',                // A
    'إجمالي الاستحقاقات (مصروفات)',  // B
    'إجمالي الدفعات (مصروفات)',      // C
    'إجمالي الإيرادات المحصلة',      // D
    'صافي التدفق النقدي',            // E
    'التدفق التراكمي'                 // F
  ];
  const widths = [130, 160, 180, 170, 170, 170];

  setupSheet_(sheet, headers, widths, CONFIG.COLORS.HEADER.CASHFLOW);

  // 🔹 قائمة الشهور من عمود "الشهر" الجديد = W (العمود 23)
  sheet.getRange('A2').setFormula(
    `=SORT(UNIQUE(FILTER('دفتر الحركات المالية'!W2:W1000,'دفتر الحركات المالية'!W2:W1000<>"")))`
  );

  // ⚡ Batch Operations - بناء كل المعادلات مرة واحدة
  const numRows = 99;
  const formulas = {
    B: [],  // إجمالي الاستحقاقات
    C: [],  // إجمالي الدفعات
    D: [],  // إجمالي الإيرادات
    E: [],  // صافي التدفق
    F: []   // التدفق التراكمي
  };

  for (let row = 2; row <= 100; row++) {
    // 🔹 إجمالي الاستحقاقات (مصروفات) في الشهر - نطاقات محددة
    formulas.B.push([
      `=IF(A${row}="","",SUMIFS('دفتر الحركات المالية'!J2:J1000,'دفتر الحركات المالية'!W2:W1000,A${row},'دفتر الحركات المالية'!C2:C1000,"استحقاق مصروف"))`
    ]);

    // 🔹 إجمالي الدفعات (مصروفات) - نطاقات محددة
    formulas.C.push([
      `=IF(A${row}="","",SUMIFS('دفتر الحركات المالية'!K2:K1000,'دفتر الحركات المالية'!W2:W1000,A${row},'دفتر الحركات المالية'!C2:C1000,"دفعة مصروف"))`
    ]);

    // 🔹 إجمالي الإيرادات المحصلة - نطاقات محددة
    formulas.D.push([
      `=IF(A${row}="","",SUMIFS('دفتر الحركات المالية'!K2:K1000,'دفتر الحركات المالية'!W2:W1000,A${row},'دفتر الحركات المالية'!C2:C1000,"تحصيل إيراد"))`
    ]);

    // 🔹 صافي التدفق = إيرادات - دفعات
    formulas.E.push([`=IF(A${row}="","",D${row}-C${row})`]);

    // 🔹 التدفق التراكمي
    formulas.F.push([`=IF(A${row}="","",SUM($E$2:E${row}))`]);
  }

  // ⚡ تطبيق كل المعادلات دفعة واحدة (5 API calls بدلاً من 495)
  sheet.getRange(2, 2, numRows, 1).setFormulas(formulas.B);
  sheet.getRange(2, 3, numRows, 1).setFormulas(formulas.C);
  sheet.getRange(2, 4, numRows, 1).setFormulas(formulas.D);
  sheet.getRange(2, 5, numRows, 1).setFormulas(formulas.E);
  sheet.getRange(2, 6, numRows, 1).setFormulas(formulas.F);

  sheet.getRange(2, 2, numRows, 5).setNumberFormat('$#,##0.00');
}

// ========= لوحة التحكم =========

/**
 * قراءة آخر رصيد من شيت حساب (بنك / خزنة / بطاقة)
 * @param {Spreadsheet} ss - الملف
 * @param {string} sheetName - اسم الشيت
 * @returns {number} - آخر رصيد أو 0 إذا لم يُعثر على بيانات
 */
function getLastBalanceFromSheet_(ss, sheetName) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return 0;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0; // لا توجد بيانات (فقط الهيدر)

  // عمود G = 7 = الرصيد (Balance)
  const balance = sheet.getRange(lastRow, 7).getValue();
  return Number(balance) || 0;
}

function createDashboardSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.SHEETS.DASHBOARD);
  if (!sheet) sheet = ss.insertSheet(CONFIG.SHEETS.DASHBOARD);
  sheet.clear();

  // إعداد الأعمدة
  sheet.setColumnWidth(1, 280);
  sheet.setColumnWidth(2, 180);
  sheet.setColumnWidth(3, 260);

  // العنوان الرئيسي
  sheet.getRange('A1:C1').merge();
  sheet.getRange('A1')
    .setValue('📊 لوحة التحكم')
    .setBackground(CONFIG.COLORS.HEADER.DASHBOARD)
    .setFontColor(CONFIG.COLORS.TEXT.WHITE)
    .setFontWeight('bold')
    .setFontSize(16)
    .setHorizontalAlignment('center');

  const metrics = [
    ['', '', ''],                                  // 3
    ['💰 المؤشرات المالية', '', ''],              // 4
    // إجمالي الاستحقاقات (مصروفات) من M (القيمة بالدولار)
    ['إجمالي الاستحقاقات (مصروفات)',
      '=SUMIFS(\'دفتر الحركات المالية\'!M2:M1000,\'دفتر الحركات المالية\'!C2:C1000,"*استحقاق مصروف*")',
      'USD'
    ],                                            // 5
    // إجمالي المدفوع (مصروفات) من M (القيمة بالدولار)
    ['إجمالي المدفوع (مصروفات)',
      '=SUMIFS(\'دفتر الحركات المالية\'!M2:M1000,\'دفتر الحركات المالية\'!C2:C1000,"*دفعة مصروف*")',
      'USD'
    ],                                            // 6
    ['🆕 المصروفات المباشرة',
      '=SUMIFS(\'دفتر الحركات المالية\'!M2:M1000,\'دفتر الحركات المالية\'!D2:D1000,"*مصروفات مباشرة*")',
      'USD'
    ],                                            // 7
    ['🆕 المصروفات العمومية',
      '=SUMIFS(\'دفتر الحركات المالية\'!M2:M1000,\'دفتر الحركات المالية\'!D2:D1000,"*مصروفات عمومية*")',
      'USD'
    ],                                            // 8
    // رصيد تقديري للموردين = استحقاقات مصروف - دفعات مصروف
    ['الرصيد الحالي (تقديري للموردين)',
      '=SUMIFS(\'دفتر الحركات المالية\'!M2:M1000,\'دفتر الحركات المالية\'!C2:C1000,"*استحقاق مصروف*")' +
      '-SUMIFS(\'دفتر الحركات المالية\'!M2:M1000,\'دفتر الحركات المالية\'!C2:C1000,"*دفعة مصروف*")',
      'USD'
    ],                                            // 9
    ['إجمالي الإيرادات المحصلة',
      '=SUMIFS(\'دفتر الحركات المالية\'!M2:M1000,\'دفتر الحركات المالية\'!C2:C1000,"*تحصيل إيراد*")',
      'USD'
    ],                                            // 10
    ['', '', ''],                                 // 11
    ['📈 إحصائيات', '', ''],                     // 12
    ['عدد المشاريع النشطة',
      '=COUNTIF(\'قاعدة بيانات المشاريع\'!O2:O200,"جاري التنفيذ")',
      'مشروع'
    ],                                            // 13
    ['عدد الموردين',
      '=COUNTIF(\'قاعدة بيانات الأطراف\'!B2:B500,"مورد")',
      'مورد'
    ],                                            // 14
    ['عدد العملاء',
      '=COUNTIF(\'قاعدة بيانات الأطراف\'!B2:B500,"عميل")',
      'عميل'
    ],                                            // 15
    ['عدد الممولين',
      '=COUNTIF(\'قاعدة بيانات الأطراف\'!B2:B500,"ممول")',
      'ممول'
    ],                                            // 16
    ['عدد الاستحقاقات المعلقة',
      '=COUNTIF(\'دفتر الحركات المالية\'!V2:V1000,"معلق")',
      'استحقاق'
    ],                                            // 17
    ['عدد الدفعات هذا الشهر',
      '=COUNTIFS(\'دفتر الحركات المالية\'!C2:C1000,"*دفعة مصروف*",\'دفتر الحركات المالية\'!W2:W1000,TEXT(TODAY(),"YYYY-MM"))',
      'دفعة'
    ],                                            // 18
    ['', '', ''],                                 // 19
    ['💵 السيولة المتاحة (بنك + خزنة + بطاقة)', '', ''], // 20
    // أرصدة البنك والخزنة والبطاقة (تُقرأ مباشرة من آخر صف في كل شيت)
    ['رصيد حساب البنك - دولار',
      getLastBalanceFromSheet_(ss, CONFIG.SHEETS.BANK_USD),
      'USD'
    ],                                            // 21
    ['رصيد حساب البنك - ليرة',
      getLastBalanceFromSheet_(ss, CONFIG.SHEETS.BANK_TRY),
      'TRY'
    ],                                            // 22
    ['رصيد خزنة العهدة - دولار',
      getLastBalanceFromSheet_(ss, CONFIG.SHEETS.CASH_USD),
      'USD'
    ],                                            // 23
    ['رصيد خزنة العهدة - ليرة',
      getLastBalanceFromSheet_(ss, CONFIG.SHEETS.CASH_TRY),
      'TRY'
    ],                                            // 24
    ['رصيد حساب البطاقة - ليرة',
      getLastBalanceFromSheet_(ss, CONFIG.SHEETS.CARD_TRY),
      'TRY'
    ],                                            // 25
    ['', '', ''],                                 // 26
    ['💱 سعر الدولار مقابل الليرة (أدخل يدوياً)', '', ''], // 27
    ['سعر الصرف (1 USD = ? TRY)',
      35,
      'أدخل: كم ليرة = 1 دولار'
    ],                                            // 28 (الخانة B28 = سعر الصرف)
    ['إجمالي السيولة المحسوبة بالدولار',
      '=B21 + B23 + (B22 / B28) + (B24 / B28) + (B25 / B28)',
      'USD (تقريبي)'
    ],                                            // 29
    ['', '', ''],                                 // 30
    ['📌 الديون (قروض + ذمم)', '', ''],          // 31
    // إجمالي القروض (تمويل الممولين) من M (القيمة بالدولار)
    ['إجمالي القروض المستلمة من الممولين',
      '=IFERROR(SUMIFS(\'دفتر الحركات المالية\'!M2:M1000,\'دفتر الحركات المالية\'!C2:C1000,"*تمويل*دخول*"),0)',
      'USD'
    ],                                            // 32
    // إجمالي سداد القروض من M (القيمة بالدولار)
    ['إجمالي سداد القروض',
      '=IFERROR(SUMIFS(\'دفتر الحركات المالية\'!M2:M1000,\'دفتر الحركات المالية\'!C2:C1000,"*سداد تمويل*"),0)',
      'USD'
    ],                                            // 33
    ['الرصيد القائم للقروض',
      '=B32-B33',
      'USD'
    ],                                            // 34
    // ذمم مدينة على العملاء = استحقاق إيراد - تحصيل إيراد (محسوب من دفتر الحركات مباشرة)
    ['إجمالي الذمم المدينة على العملاء (إيرادات لم تُحصَّل)',
      '=SUMIFS(\'دفتر الحركات المالية\'!M2:M1000,\'دفتر الحركات المالية\'!C2:C1000,"*استحقاق إيراد*")' +
      '-SUMIFS(\'دفتر الحركات المالية\'!M2:M1000,\'دفتر الحركات المالية\'!C2:C1000,"*تحصيل إيراد*")',
      'USD'
    ],                                            // 35
    // ذمم دائنة لصالح الموردين = نفس الرصيد الحالي للموردين (B9)
    ['إجمالي الذمم الدائنة لصالح الموردين',
      '=B9',
      'USD'
    ]                                             // 36
  ];

  // كتابة الجدول ابتداءً من الصف 3
  sheet.getRange(3, 1, metrics.length, 3).setValues(metrics);

  // تلوين عناوين الأقسام
  sheet.getRange('A4:C4')   // المؤشرات المالية
    .setBackground(CONFIG.COLORS.HEADER.REPORTS)
    .setFontColor(CONFIG.COLORS.TEXT.WHITE)
    .setFontWeight('bold');

  sheet.getRange('A12:C12') // إحصائيات
    .setBackground(CONFIG.COLORS.HEADER.REPORTS)
    .setFontColor(CONFIG.COLORS.TEXT.WHITE)
    .setFontWeight('bold');

  sheet.getRange('A20:C20') // السيولة
    .setBackground(CONFIG.COLORS.HEADER.REVENUE)
    .setFontColor(CONFIG.COLORS.TEXT.WHITE)
    .setFontWeight('bold');

  sheet.getRange('A31:C31') // الديون
    .setBackground('#6d4c41')
    .setFontColor(CONFIG.COLORS.TEXT.WHITE)
    .setFontWeight('bold');

  // تنسيقات الأرقام
  // مؤشرات مالية (صفوف 5–10 في العمود B)
  sheet.getRange(5, 2, 6, 1).setNumberFormat('$#,##0.00');
  // أرصدة بنك وخزنة وبطاقة (صفوف 21–25)
  sheet.getRange(21, 2, 5, 1).setNumberFormat('#,##0.00');
  // سعر الصرف
  sheet.getRange(28, 2).setNumberFormat('#,##0.0000');
  // إجمالي السيولة بالدولار
  sheet.getRange(29, 2).setNumberFormat('$#,##0.00');
  // القروض + الذمم (صفوف 32–36)
  sheet.getRange(32, 2, 5, 1).setNumberFormat('$#,##0.00');

  // إبراز سطر "إجمالي الإيرادات المحصلة" مثل ما كنا عاملين قبل كده
  sheet.getRange('A10:C10')
    .setBackground('#ffd54f')
    .setFontWeight('bold')
    .setFontSize(13);

  sheet.setFrozenRows(2);
}

/**
 * onEdit - معالجة التعديلات في الشيتات
 *
 * 1. تطبيع التواريخ:
 *    - دفتر الحركات المالية: أعمدة B (2) و T (20)
 *    - قاعدة بيانات المشاريع: أعمدة J (10) و K (11)
 *    - تحويل النصوص إلى Date objects
 *    - قبول فواصل متعددة (/ . -)
 *
 * 2. المزامنة الثنائية في دفتر الحركات (أعمدة E و F):
 *    - عند اختيار كود المشروع → يُملأ اسم المشروع تلقائياً
 *    - عند اختيار اسم المشروع → يُملأ كود المشروع تلقائياً
 *
 * 3. المزامنة الثنائية في الموازنات المخططة (أعمدة A و B):
 *    - عند اختيار كود المشروع → يُملأ اسم المشروع تلقائياً
 *    - عند اختيار اسم المشروع → يُملأ كود المشروع تلقائياً
 */
function onEdit(e) {
  if (!e || !e.range || !e.source) return;

  const sheet = e.source.getActiveSheet();
  const sheetName = sheet.getName();
  const row = e.range.getRow();
  const col = e.range.getColumn();

  // تجاهل الهيدر
  if (row <= 1) return;

  // ═══════════════════════════════════════════════════════════
  // تسجيل النشاط مع الإيميل الصحيح (Simple Trigger يمكنه جلب الإيميل)
  // ═══════════════════════════════════════════════════════════
  try {
    const trackedForLog = [
      CONFIG.SHEETS.TRANSACTIONS,
      CONFIG.SHEETS.PROJECTS,
      CONFIG.SHEETS.PARTIES,
      CONFIG.SHEETS.ITEMS,
      CONFIG.SHEETS.BUDGETS
    ];

    if (trackedForLog.includes(sheetName)) {
      // جلب الإيميل الفعلي للمستخدم (يعمل في Simple Trigger فقط)
      const realUserEmail = Session.getActiveUser().getEmail();

      if (realUserEmail) {
        const oldValue = e.oldValue !== undefined ? e.oldValue : '';
        const newValue = e.value !== undefined ? e.value : '';

        // تجاهل إذا لم يتغير شيء
        if (oldValue !== newValue) {
          const columnHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
          const columnName = columnHeaders[col - 1] || 'عمود ' + col;

          let transNum = '';
          if (sheetName === CONFIG.SHEETS.TRANSACTIONS) {
            transNum = sheet.getRange(row, 1).getValue() || '';
          }

          let actionType = 'تعديل';
          if (oldValue === '' && newValue !== '') {
            actionType = 'إضافة قيمة';
          } else if (oldValue !== '' && newValue === '') {
            actionType = 'حذف قيمة';
          }

          // تسجيل النشاط مع الإيميل الصحيح
          logActivity(
            actionType,
            sheetName,
            row,
            transNum,
            columnName + ': "' + oldValue + '" → "' + newValue + '"',
            {
              column: col,
              columnName: columnName,
              oldValue: oldValue,
              newValue: newValue
            },
            realUserEmail
          );
        }
      }
    }
  } catch (logErr) {
    // تجاهل أخطاء التسجيل - لا نوقف العمليات الأخرى
  }

  // ═══════════════════════════════════════════════════════════
  // قائمة الشيتات المعالجة - الخروج السريع إذا لم يكن الشيت منها
  // ═══════════════════════════════════════════════════════════
  const isTransactions = (sheetName === CONFIG.SHEETS.TRANSACTIONS);
  const isProjects = (sheetName === CONFIG.SHEETS.PROJECTS);
  const isBudgets = (sheetName === CONFIG.SHEETS.BUDGETS);
  const isVendorsReport = (sheetName === CONFIG.SHEETS.VENDORS_REPORT);
  const isFundersReport = (sheetName === CONFIG.SHEETS.FUNDERS_REPORT);
  const isCommissionReport = sheetName.indexOf('تقرير عمولة - ') === 0;

  // الخروج السريع إذا لم يكن الشيت يحتاج معالجة
  if (!isTransactions && !isProjects && !isBudgets && !isVendorsReport && !isFundersReport && !isCommissionReport) {
    return;
  }

  const value = e.value || e.range.getValue();

  // ═══════════════════════════════════════════════════════════
  // معالجة أعمدة التاريخ في قاعدة بيانات المشاريع: J (10) و K (11)
  // ═══════════════════════════════════════════════════════════
  if (isProjects) {
    if (col === 10 || col === 11) {
      if (value) normalizeDateCell_(e.range, value);
    }
    return;
  }

  // ═══════════════════════════════════════════════════════════
  // 🆕 معالجة الموازنات المخططة - المزامنة الثنائية (أعمدة A و B)
  // ═══════════════════════════════════════════════════════════
  if (isBudgets) {
    if ((col === 1 || col === 2) && value) {
      const projectsSheet = e.source.getSheetByName(CONFIG.SHEETS.PROJECTS);
      if (!projectsSheet) return;

      const projectData = projectsSheet.getRange('A2:B200').getValues();

      if (col === 1) {
        // تم اختيار كود المشروع (A) → ابحث عن الاسم (B)
        for (let i = 0; i < projectData.length; i++) {
          if (projectData[i][0] === value) {
            sheet.getRange(row, 2).setValue(projectData[i][1]);
            break;
          }
        }
      } else if (col === 2) {
        // تم اختيار اسم المشروع (B) → ابحث عن الكود (A)
        for (let i = 0; i < projectData.length; i++) {
          if (projectData[i][1] === value) {
            sheet.getRange(row, 1).setValue(projectData[i][0]);
            break;
          }
        }
      }
    }
    return;
  }

  // ═══════════════════════════════════════════════════════════
  // 📄 معالجة تقرير الموردين - عمود J (10) لإنشاء كشف حساب
  // ═══════════════════════════════════════════════════════════
  if (isVendorsReport) {
    if (col === 10) {
      const vendorName = sheet.getRange(row, 1).getValue();
      if (vendorName) {
        generateUnifiedStatement_(e.source, vendorName, 'مورد');
      }
    }
    return;
  }

  // ═══════════════════════════════════════════════════════════
  // معالجة تقرير الممولين - عمود كشف الحساب (J = 10)
  // ═══════════════════════════════════════════════════════════
  if (isFundersReport) {
    if (col === 10) {
      const funderName = sheet.getRange(row, 1).getValue();
      if (funderName) {
        generateUnifiedStatement_(e.source, funderName, 'ممول');
      }
    }
    return;
  }

  // ═══════════════════════════════════════════════════════════
  // 💰 معالجة تقرير العمولات - عمود الإدراج (H = 8)
  // ═══════════════════════════════════════════════════════════
  if (isCommissionReport) {
    if (col === 8 && value === true) {
      handleCommissionCheckbox(sheet, row, col);
    }
    return;
  }

  // ═══════════════════════════════════════════════════════════
  // معالجة دفتر الحركات المالية فقط
  // ═══════════════════════════════════════════════════════════
  if (!isTransactions) return;

  const ss = e.source;

  // ═══════════════════════════════════════════════════════════
  // معالجة أعمدة التاريخ: B (2) و T (20)
  // ═══════════════════════════════════════════════════════════
  if (col === 2 || col === 20) {
    if (value) normalizeDateCell_(e.range, value);
  }

  // ═══════════════════════════════════════════════════════════
  // 🔗 الربط التلقائي: طبيعة الحركة (C=3) ← نوع الحركة (N=14)
  // استحقاق = مدين استحقاق | تمويل/دفعة/تحصيل = دائن دفعة
  // ✅ تمويل = دائن دفعة (نقد داخل) مع قيد ضمني (التزام للممول)
  // ═══════════════════════════════════════════════════════════
  if (col === 3 && value) {
    const valueStr = String(value);
    // فقط الاستحقاقات = مدين استحقاق | كل شيء آخر (تمويل/دفعة/تحصيل) = دائن دفعة
    const isAccrual = valueStr.indexOf('استحقاق') !== -1;
    const movementType = isAccrual ? 'مدين استحقاق' : 'دائن دفعة';
    sheet.getRange(row, 14).setValue(movementType);
  }

  // ═══════════════════════════════════════════════════════════
  // 🆕 إدراج رقم الحركة والتاريخ تلقائياً عند بدء إدخال حركة جديدة
  // يعمل عند الكتابة في أي من الأعمدة الأساسية: B, C, D, E, F, G, H, I, J
  // ═══════════════════════════════════════════════════════════
  const isBasicColumn = (col >= 2 && col <= 10);
  if (isBasicColumn && value) {
    const cellA = sheet.getRange(row, 1);
    const cellB = sheet.getRange(row, 2);
    const valueA = cellA.getValue();
    const valueB = cellB.getValue();

    // إدراج رقم الحركة إذا كان فارغاً
    if (!valueA && valueA !== 0) {
      cellA.setValue(row - 1);
    }

    // إدراج التاريخ الحالي إذا كان فارغاً (عند الكتابة في أي عمود غير B)
    if (col !== 2 && !valueB) {
      cellB.setValue(new Date()).setNumberFormat('dd/mm/yyyy');
    }

    // إضافة أيقونة كشف الحساب في عمود Y (25) إذا كانت فارغة
    const cellY = sheet.getRange(row, 25);
    if (!cellY.getValue()) {
      cellY.setValue('📄');
    }
  }

  // ═══════════════════════════════════════════════════════════
  // حساب القيمة بالدولار (M) عند تغيير J أو K أو L
  // J=10 (المبلغ), K=11 (العملة), L=12 (سعر الصرف)
  // ═══════════════════════════════════════════════════════════
  if (col === 10 || col === 11 || col === 12) {
    calculateUsdValue_(sheet, row);
    // بعد حساب M، نحتاج تحديث الرصيد O لكل حركات نفس الطرف
    recalculatePartyBalance_(sheet, row);
  }

  // ═══════════════════════════════════════════════════════════
  // حساب تاريخ الاستحقاق (U) عند تغيير B أو E أو N أو R أو S أو T
  // B=2 (التاريخ), E=5 (كود المشروع), N=14 (نوع الحركة)
  // R=18 (نوع شرط الدفع), S=19 (عدد الأسابيع), T=20 (تاريخ مخصص)
  // ═══════════════════════════════════════════════════════════
  if (col === 2 || col === 5 || col === 14 || col === 18 || col === 19 || col === 20) {
    calculateDueDate_(ss, sheet, row);
  }

  // ═══════════════════════════════════════════════════════════
  // حساب حالة السداد (V) عند تغيير N أو I
  // N=14 (نوع الحركة), I=9 (الطرف)
  // ═══════════════════════════════════════════════════════════
  if (col === 14 || col === 9) {
    // تحديث الرصيد أولاً ثم حالة السداد
    recalculatePartyBalance_(sheet, row);
  }

  // ═══════════════════════════════════════════════════════════
  // 📄 إنشاء كشف حساب عند التعديل على عمود Y (25)
  // ═══════════════════════════════════════════════════════════
  if (col === 25) {
    generateStatementFromRow_(ss, sheet, row);
    return; // لا نحتاج معالجة إضافية
  }

  // ═══════════════════════════════════════════════════════════
  // معالجة أعمدة المشروع: E (5) و F (6)
  // ═══════════════════════════════════════════════════════════
  if ((col === 5 || col === 6) && value) {
    const projectsSheet = ss.getSheetByName(CONFIG.SHEETS.PROJECTS);
    if (!projectsSheet) return;

    const projectData = projectsSheet.getRange('A2:B200').getValues();

    if (col === 5) {
      // تم اختيار كود المشروع (E) → ابحث عن الاسم (F)
      for (let i = 0; i < projectData.length; i++) {
        if (projectData[i][0] === value) {
          sheet.getRange(row, 6).setValue(projectData[i][1]);
          break;
        }
      }
    } else if (col === 6) {
      // تم اختيار اسم المشروع (F) → ابحث عن الكود (E)
      for (let i = 0; i < projectData.length; i++) {
        if (projectData[i][1] === value) {
          sheet.getRange(row, 5).setValue(projectData[i][0]);
          break;
        }
      }
    }
  }
}

// ═══════════════════════════════════════════════════════════
// دوال الحساب التلقائي (تُستدعى من onEdit)
// ═══════════════════════════════════════════════════════════

/**
 * حساب القيمة بالدولار (M) لصف معين
 * المنطق: لو دولار = نفس القيمة، لو عملة أخرى = القيمة ÷ سعر الصرف
 * ⚠️ إذا كانت العملة غير دولار ولا يوجد سعر صرف = ترك الخلية فارغة (تحتاج إدخال سعر الصرف)
 */
function calculateUsdValue_(sheet, row) {
  const rowData = sheet.getRange(row, 10, 1, 3).getValues()[0]; // J, K, L
  const amount = Number(rowData[0]) || 0;      // J: المبلغ
  const currency = String(rowData[1] || '').trim().toUpperCase(); // K: العملة
  const exchangeRate = Number(rowData[2]) || 0; // L: سعر الصرف

  let amountUsd = '';
  if (amount > 0) {
    // حالة 1: العملة دولار أو فارغة (افتراضي دولار)
    if (currency === 'USD' || currency === 'دولار' || currency === '') {
      amountUsd = amount;
    }
    // حالة 2: عملة أخرى مع سعر صرف صحيح
    else if (exchangeRate > 0) {
      amountUsd = Math.round((amount / exchangeRate) * 100) / 100;
    }
    // حالة 3: عملة أخرى بدون سعر صرف = ترك فارغ (⚠️ يحتاج إدخال سعر الصرف)
    else {
      amountUsd = ''; // لا نفترض أن المبلغ بالدولار - هذا خطأ منطقي
    }
  }

  sheet.getRange(row, 13).setValue(amountUsd); // M
}

/**
 * حساب تاريخ الاستحقاق (U) لصف معين
 * المنطق: فوري=تاريخ الحركة، بعد التسليم=تاريخ التسليم+أسابيع، تاريخ مخصص=T
 */
function calculateDueDate_(ss, sheet, row) {
  // قراءة البيانات المطلوبة
  const dateVal = sheet.getRange(row, 2).getValue();      // B: تاريخ الحركة
  const projectCode = sheet.getRange(row, 5).getValue();  // E: كود المشروع
  const movementKind = sheet.getRange(row, 14).getValue(); // N: نوع الحركة
  const paymentTermType = sheet.getRange(row, 18).getValue(); // R: نوع شرط الدفع
  const weeks = Number(sheet.getRange(row, 19).getValue()) || 0; // S: عدد الأسابيع
  const customDate = sheet.getRange(row, 20).getValue();  // T: تاريخ مخصص

  let dueDate = '';

  // فقط للحركات من نوع "مدين استحقاق"
  if (movementKind !== 'مدين استحقاق' || !paymentTermType) {
    sheet.getRange(row, 21).setValue(''); // U
    return;
  }

  if (paymentTermType === 'فوري') {
    // تاريخ الاستحقاق = تاريخ الحركة
    dueDate = dateVal;
  } else if (paymentTermType === 'بعد التسليم') {
    // جلب تاريخ التسليم من قاعدة بيانات المشاريع
    if (projectCode) {
      const projectsSheet = ss.getSheetByName(CONFIG.SHEETS.PROJECTS);
      if (projectsSheet) {
        const projectData = projectsSheet.getRange('A2:K200').getValues();
        for (let i = 0; i < projectData.length; i++) {
          if (projectData[i][0] === projectCode) {
            const deliveryDate = projectData[i][10]; // K: تاريخ التسليم المتوقع
            if (deliveryDate instanceof Date) {
              const newDate = new Date(deliveryDate);
              newDate.setDate(newDate.getDate() + (weeks * 7));
              dueDate = newDate;
            }
            break;
          }
        }
      }
    }
  } else if (paymentTermType === 'تاريخ مخصص') {
    dueDate = customDate;
  }

  // كتابة القيمة
  sheet.getRange(row, 21).setValue(dueDate); // U
  if (dueDate) {
    sheet.getRange(row, 21).setNumberFormat('dd/mm/yyyy');
  }
}

/**
 * إعادة حساب الرصيد (O) وحالة السداد (V) لجميع حركات الطرف
 */
function recalculatePartyBalance_(sheet, editedRow) {
  const party = String(sheet.getRange(editedRow, 9).getValue() || '').trim(); // I: الطرف
  if (!party) return;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  // قراءة كل البيانات مرة واحدة (I, M, N)
  const allData = sheet.getRange(2, 9, lastRow - 1, 6).getValues(); // I to N (columns 9-14)
  // allData[i][0] = I (الطرف), index 9
  // allData[i][4] = M (القيمة بالدولار), index 13
  // allData[i][5] = N (نوع الحركة), index 14

  // حساب الأرصدة التراكمية لكل طرف
  const partyBalances = {};
  const balanceValues = [];
  const statusValues = [];

  for (let i = 0; i < allData.length; i++) {
    const rowParty = String(allData[i][0] || '').trim();
    const amountUsd = Number(allData[i][4]) || 0; // M at index 4 (relative to column 9)
    const movementKind = String(allData[i][5] || '').trim(); // N at index 5

    let balance = '';
    let status = '';

    if (rowParty && amountUsd > 0) {
      if (!partyBalances[rowParty]) {
        partyBalances[rowParty] = 0;
      }

      if (movementKind === 'مدين استحقاق') {
        partyBalances[rowParty] += amountUsd;
      } else if (movementKind === 'دائن دفعة') {
        partyBalances[rowParty] -= amountUsd;
      }

      balance = Math.round(partyBalances[rowParty] * 100) / 100;

      // حساب حالة السداد (باستخدام CONFIG.PAYMENT_STATUS للتوحيد)
      if (movementKind === 'دائن دفعة') {
        status = CONFIG.PAYMENT_STATUS.OPERATION; // 'عملية دفع/تحصيل'
      } else if (balance > 0.01) {
        status = CONFIG.PAYMENT_STATUS.PENDING; // 'معلق'
      } else {
        status = CONFIG.PAYMENT_STATUS.PAID; // 'مدفوع بالكامل'
      }
    }

    balanceValues.push([balance]);
    statusValues.push([status]);
  }

  // كتابة القيم دفعة واحدة
  sheet.getRange(2, 15, lastRow - 1, 1).setValues(balanceValues); // O: الرصيد
  sheet.getRange(2, 22, lastRow - 1, 1).setValues(statusValues);  // V: حالة السداد
}

/**
 * تطبيع خلية تاريخ - تحويل النص إلى Date object وضبط التنسيق
 * @param {Range} range - الخلية
 * @param {*} value - القيمة الحالية
 */
function normalizeDateCell_(range, value) {
  // تجاهل إذا كانت القيمة Date object بالفعل
  if (value instanceof Date) {
    // فقط تأكد من التنسيق الصحيح
    range.setNumberFormat('dd/mm/yyyy');
    return;
  }

  // تجاهل إذا كانت القيمة رقم (serial date من Sheets)
  if (typeof value === 'number') {
    range.setNumberFormat('dd/mm/yyyy');
    return;
  }

  // محاولة تحويل النص إلى تاريخ
  const dateStr = String(value).trim();
  if (!dateStr) return;

  const parseResult = parseDateInput_(dateStr);
  if (parseResult.success) {
    range.setValue(parseResult.dateObj);
    range.setNumberFormat('dd/mm/yyyy');
  }
}

function applyTransactionsDropdowns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.TRANSACTIONS);
  if (!sheet) return;

  const lastRow = sheet.getMaxRows();

  // نجيب العناوين علشان نشتغل بالاسم بدل رقم العمود
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  const colType = headers.indexOf('طبيعة الحركة') + 1;
  const colClass = headers.indexOf('تصنيف الحركة') + 1;
  const colItem = headers.indexOf('البند') + 1;

  const dvBuilder = SpreadsheetApp.newDataValidation;

  // 🔹 1) قائمة طبيعة الحركة
  if (colType > 0) {
    const typeRule = dvBuilder()
      .requireValueInList([
        'استحقاق مصروف',
        'دفعة مصروف',
        'استحقاق إيراد',
        'تحصيل إيراد',
        'تمويل',
        'سداد تمويل',
        'تأمين مدفوع للقناة',
        'استرداد تأمين من القناة'
      ], true)
      .setAllowInvalid(false)
      .build();

    sheet.getRange(2, colType, lastRow - 1, 1).setDataValidation(typeRule);
  }

  // 🔹 2) قائمة تصنيف الحركة
  if (colClass > 0) {
    const classRule = dvBuilder()
      .requireValueInList([
        'مصروفات مباشرة',
        'مصروفات عمومية',
        'مصروفات أخرى',
        'إيراد عقد',
        'تمويل'
      ], true)
      .setAllowInvalid(false)
      .build();

    sheet.getRange(2, colClass, lastRow - 1, 1).setDataValidation(classRule);
  }

  // 🔹 3) قائمة البنود من "قاعدة بيانات البنود"
  const itemsSheet = ss.getSheetByName(CONFIG.SHEETS.ITEMS);
  if (itemsSheet && colItem > 0) {
    const lastItemRow = itemsSheet.getLastRow();
    if (lastItemRow > 1) {
      const itemRange = itemsSheet.getRange(2, 1, lastItemRow - 1, 1); // A2:A

      const itemRule = dvBuilder()
        .requireValueInRange(itemRange, true)
        .setAllowInvalid(false)
        .build();

      sheet.getRange(2, colItem, lastRow - 1, 1).setDataValidation(itemRule);
    }
  }
}
// ==================== شيتات البنك وخزنة العهدة (دولار / ليرة) ====================

// دالة مساعدة صغيرة للبحث عن عمود بالاسم (أو أكثر من اسم محتمل)
function findHeaderIndex_(headers, names) {
  names = Array.isArray(names) ? names : [names];
  for (let i = 0; i < headers.length; i++) {
    const h = String(headers[i] || '').trim();
    for (let j = 0; j < names.length; j++) {
      if (h === String(names[j]).trim()) return i;
    }
  }
  return -1;
}

function createSingleAccountSheet(ss, sheetName, currency) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  sheet.clear();

  // نخلي اتجاه الشيت عادي (شمال → يمين) زي باقي النظام
  sheet.setRightToLeft(false);

  const headers = [
    'Date',
    'Statement',
    'Trans No',
    'Ref',
    'Debit (' + currency + ')',
    'Credit (' + currency + ')',
    'Balance (' + currency + ')',
    'Notes'
  ];

  sheet.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setBackground('#37474f')
    .setFontColor(CONFIG.COLORS.TEXT.WHITE)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  const widths = [110, 260, 110, 130, 130, 130, 130, 200];
  widths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));

  sheet.setFrozenRows(1);
  sheet.getRange(2, 5, sheet.getMaxRows() - 1, 3).setNumberFormat('#,##0.00');

  sheet.getRange('B2').setNote(
    'يتم تعبئة هذا الشيت تلقائياً من "دفتر الحركات المالية" عن طريق الدالة rebuildBankAndCashFromTransactions().'
  );

  return sheet;
}

function createBankAndCashSheets(ss) {
  createSingleAccountSheet(ss, 'حساب البنك - دولار', 'USD');
  createSingleAccountSheet(ss, 'حساب البنك - ليرة', 'TRY');
  createSingleAccountSheet(ss, 'خزنة العهدة - دولار', 'USD');
  createSingleAccountSheet(ss, 'خزنة العهدة - ليرة', 'TRY');
  // 🆕 شيت خاص بحركة البطاقة (عادة ليرة)
  createSingleAccountSheet(ss, 'حساب البطاقة - ليرة', 'TRY');
}

// ==================== بناء شيتات البنك والعهدة من دفتر الحركات (من غير أعمدة زيادة) ====================

function rebuildBankAndCashFromTransactions(silent) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const transSheet = ss.getSheetByName(CONFIG.SHEETS.TRANSACTIONS);
  if (!transSheet) {
    if (!silent) ui.alert('⚠️ شيت "دفتر الحركات المالية" غير موجود.');
    return { success: false, name: 'البنوك والخزنة', error: 'شيت الحركات غير موجود' };
  }

  // نتأكد من وجود شيتات البنك والعهدة والبطاقة
  createBankAndCashSheets(ss);

  const bankUsdSheet = ss.getSheetByName(CONFIG.SHEETS.BANK_USD);
  const bankTrySheet = ss.getSheetByName(CONFIG.SHEETS.BANK_TRY);
  const cashUsdSheet = ss.getSheetByName(CONFIG.SHEETS.CASH_USD);
  const cashTrySheet = ss.getSheetByName(CONFIG.SHEETS.CASH_TRY);
  const cardTrySheet = ss.getSheetByName(CONFIG.SHEETS.CARD_TRY);

  const data = transSheet.getDataRange().getValues();
  if (data.length < 2) {
    if (!silent) ui.alert('⚠️ لا توجد حركات في "دفتر الحركات المالية".');
    return { success: false, name: 'البنوك والخزنة', error: 'لا توجد حركات' };
  }

  const headers = data[0];

  // خريطة الأعمدة حسب ترتيبك الحالي
  const col = {
    transNo: findHeaderIndex_(headers, 'رقم الحركة'),
    date: findHeaderIndex_(headers, 'التاريخ'),
    type: findHeaderIndex_(headers, 'طبيعة الحركة'),
    classification: findHeaderIndex_(headers, 'تصنيف الحركة'),
    details: findHeaderIndex_(headers, 'التفاصيل'),
    party: findHeaderIndex_(headers, 'اسم المورد/الجهة'),
    amount: findHeaderIndex_(headers, 'المبلغ بالعملة الأصلية'),
    currency: findHeaderIndex_(headers, ['العملة', 'العملة الأصلية']),
    rate: findHeaderIndex_(headers, 'سعر الصرف'),
    amountUsd: findHeaderIndex_(headers, 'القيمة بالدولار'),
    refNo: findHeaderIndex_(headers, 'رقم مرجعي'),
    payMethod: findHeaderIndex_(headers, 'طريقة الدفع'),
    status: findHeaderIndex_(headers, 'حالة السداد'),
    notes: findHeaderIndex_(headers, 'ملاحظات')
  };

  // لو الأعمدة الأساسية مش موجودة نوقف
  if (col.currency === -1 || col.amount === -1 || col.payMethod === -1 || col.type === -1) {
    if (!silent) {
      ui.alert(
        '⚠️ لا يمكن تحديث شيتات البنك والعهدة.\n' +
        'تأكد من وجود الأعمدة التالية بالضبط في "دفتر الحركات المالية":\n' +
        'رقم الحركة، التاريخ، طبيعة الحركة، المبلغ بالعملة الأصلية، العملة، طريقة الدفع، ملاحظات (اختياري).'
      );
    }
    return { success: false, name: 'البنوك والخزنة', error: 'أعمدة ناقصة' };
  }

  // إعداد حاويات الحسابات
  const accounts = {
    bankUsd: { sheet: bankUsdSheet, rows: [], balance: 0 },
    bankTry: { sheet: bankTrySheet, rows: [], balance: 0 },
    cashUsd: { sheet: cashUsdSheet, rows: [], balance: 0 },
    cashTry: { sheet: cashTrySheet, rows: [], balance: 0 },
    cardTry: { sheet: cardTrySheet, rows: [], balance: 0 }
  };

  // 🔍 تحديد نوع الحساب (بنك / خزنة / بطاقة + العملة)
  function detectAccountKey(payMethodVal, currencyVal) {
    const pm = String(payMethodVal || '').toLowerCase();
    const cur = String(currencyVal || '').toLowerCase();

    const isCash =
      pm.indexOf('نقد') !== -1 ||
      pm.indexOf('كاش') !== -1 ||
      pm.indexOf('خزنة') !== -1 ||
      pm.indexOf('عهدة') !== -1 ||
      pm.indexOf('cash') !== -1;

    const isBank =
      pm.indexOf('تحويل') !== -1 ||
      pm.indexOf('بنكي') !== -1 ||
      pm.indexOf('bank') !== -1;

    const isCard =
      pm.indexOf('بطاقة') !== -1 ||
      pm.indexOf('كريدت') !== -1 ||
      pm.indexOf('credit') !== -1 ||
      pm.indexOf('visa') !== -1 ||
      pm.indexOf('ماستر') !== -1;

    const isUsd =
      cur.indexOf('usd') !== -1 ||
      cur.indexOf('دولار') !== -1 ||
      cur.indexOf('$') !== -1;

    const isTry =
      cur.indexOf('try') !== -1 ||
      cur.indexOf('tl') !== -1 ||
      cur.indexOf('ليرة') !== -1;

    const isEgp =
      cur.indexOf('egp') !== -1 ||
      cur.indexOf('جنيه') !== -1 ||
      cur.indexOf('ج.م') !== -1;

    if (isCard) return 'cardTry';             // البطاقة دائمًا ليرة
    if (isBank && isUsd) return 'bankUsd';
    if (isBank && isTry) return 'bankTry';
    if (isCash && isUsd) return 'cashUsd';
    if (isCash && isTry) return 'cashTry';
    if (isCash && isEgp) return 'cashUsd';             // الجنيه يتحول لدولار ويذهب لخزنة الدولار
    return null;
  }

  // نعدّي على كل الصفوف
  for (let i = 1; i < data.length; i++) {
    const row = data[i];

    const typeVal = String(row[col.type] || '').trim();
    const classVal = col.classification >= 0 ? String(row[col.classification] || '').trim() : '';
    const detailsVal = col.details >= 0 ? String(row[col.details] || '').trim() : '';
    const statusVal = col.status >= 0 ? String(row[col.status] || '').trim() : '';

    const payMethodVal = row[col.payMethod];
    const currencyVal = row[col.currency];

    // 1) لو مفيش طريقة دفع أو عملة ⇒ مش حركة نقدية أصلاً
    if (!payMethodVal || !currencyVal) continue;

    // 2) تحديد هل هي استحقاق؟
    const isAccrual =
      typeVal.indexOf('استحقاق') !== -1 ||   // طبيعة الحركة فيها "استحقاق"
      statusVal === 'معلق';                  // أو حالة السداد "معلق"

    // 3) تحديد هل هي تمويل (قصير/طويل/سلفة قصيرة الأجل)
    const isFinancing =
      // طبيعة الحركة = تمويل (بدون سداد تمويل)
      (typeVal.indexOf('تمويل') !== -1 && typeVal.indexOf('سداد تمويل') === -1) ||
      // أي نوع تمويل مذكور بالاسم في التصنيف أو البند
      classVal.indexOf('تمويل') !== -1 ||
      detailsVal.indexOf('تمويل') !== -1 ||

      // سلفة قصيرة الأجل (تُعامل كتمويل قصير الأجل)
      classVal.indexOf('سلفة قصيرة') !== -1 ||
      detailsVal.indexOf('سلفة قصيرة') !== -1;
    // 4) تحديد هل هي حركة مدفوعة فعليًا؟
    const isPaidMovement =
      statusVal === 'عملية دفع/تحصيل' ||
      statusVal === CONFIG.PAYMENT_STATUS.PAID ||
      statusVal === 'مدفوع جزئياً';

    // 5) تحديد هل هي تحويل داخلي؟
    const isInternalTransfer = typeVal.indexOf('تحويل داخلي') !== -1;

    // 🔴 استبعاد كل الاستحقاقات غير الممولة
    // 🔴 واستبعاد أي حركة غير مدفوعة فعليًا وليست استحقاق تمويل وليست تحويل داخلي
    if (!isPaidMovement && !(isAccrual && isFinancing) && !isInternalTransfer) {
      // يعني: ليست حركة مدفوعة، وليست استحقاق تمويل، وليست تحويل داخلي ⇒ مالهاش أثر نقدي
      continue;
    }

    // 5) تحديد الحساب المناسب
    const key = detectAccountKey(payMethodVal, currencyVal);
    if (!key || !accounts[key]) continue;

    const acc = accounts[key];

    const date = col.date >= 0 ? row[col.date] : '';
    const transNo = col.transNo >= 0 ? row[col.transNo] : '';
    const refNo = col.refNo >= 0 ? row[col.refNo] : '';
    const party = col.party >= 0 ? String(row[col.party] || '') : '';
    const notes = col.notes >= 0 ? row[col.notes] || '' : '';

    // 6) تحديد المبلغ:
    //    - USD / TRY → من "المبلغ بالعملة الأصلية"
    //    - EGP → من "القيمة بالدولار" ويروح لخزنة الدولار
    const cur = String(currencyVal).toLowerCase();
    let amount = 0;

    const isEgp =
      cur.indexOf('egp') !== -1 ||
      cur.indexOf('جنيه') !== -1 ||
      cur.indexOf('ج.م') !== -1;

    if (isEgp && col.amountUsd !== -1) {
      amount = Number(row[col.amountUsd]) || 0;
    } else {
      amount = Number(row[col.amount]) || 0;
    }

    if (!amount) continue;

    // 7) تحديد اتجاه الحركة (داخل / خارج الحساب)
    let debitAcc = 0;
    let creditAcc = 0;

    // ═══════════════════════════════════════════════════════════
    // 🔄 معالجة التحويل الداخلي (بين البنك والخزنة)
    // ═══════════════════════════════════════════════════════════
    if (isInternalTransfer) {
      const isTransferToBank = classVal.indexOf('تحويل للبنك') !== -1;
      const isTransferToCash = classVal.indexOf('تحويل للخزنة') !== -1 || classVal.indexOf('تحويل للكاش') !== -1;

      // تحديد العملة
      const isUsdCurrency = cur.indexOf('usd') !== -1 || cur.indexOf('دولار') !== -1 || cur.indexOf('$') !== -1;
      const isTryCurrency = cur.indexOf('try') !== -1 || cur.indexOf('tl') !== -1 || cur.indexOf('ليرة') !== -1;

      if (isTransferToBank) {
        // تحويل للبنك = خصم من الخزنة + إضافة للبنك
        const destKey = isUsdCurrency ? 'bankUsd' : (isTryCurrency ? 'bankTry' : null);
        const srcKey = isUsdCurrency ? 'cashUsd' : (isTryCurrency ? 'cashTry' : null);

        if (destKey && srcKey && accounts[destKey] && accounts[srcKey]) {
          // إضافة للبنك (الوجهة)
          accounts[destKey].balance += amount;
          accounts[destKey].rows.push([
            date, 'تحويل من الخزنة', transNo, refNo, amount, 0, accounts[destKey].balance, notes
          ]);

          // خصم من الخزنة (المصدر)
          accounts[srcKey].balance -= amount;
          accounts[srcKey].rows.push([
            date, 'تحويل إلى البنك', transNo, refNo, 0, amount, accounts[srcKey].balance, notes
          ]);
        }
      } else if (isTransferToCash) {
        // تحويل للخزنة = خصم من البنك + إضافة للخزنة
        const destKey = isUsdCurrency ? 'cashUsd' : (isTryCurrency ? 'cashTry' : null);
        const srcKey = isUsdCurrency ? 'bankUsd' : (isTryCurrency ? 'bankTry' : null);

        if (destKey && srcKey && accounts[destKey] && accounts[srcKey]) {
          // إضافة للخزنة (الوجهة)
          accounts[destKey].balance += amount;
          accounts[destKey].rows.push([
            date, 'تحويل من البنك', transNo, refNo, amount, 0, accounts[destKey].balance, notes
          ]);

          // خصم من البنك (المصدر)
          accounts[srcKey].balance -= amount;
          accounts[srcKey].rows.push([
            date, 'تحويل إلى الخزنة', transNo, refNo, 0, amount, accounts[srcKey].balance, notes
          ]);
        }
      }
      // التحويل الداخلي تم معالجته، ننتقل للصف التالي
      continue;
    }
    // ═══════════════════════════════════════════════════════════

    // فلوس داخلة الحساب (تحصيل / تمويل / استرداد…)
    // ملاحظة: "سداد تمويل" يحتوي على "تمويل" لكنه فلوس خارجة، لذا نستثنيه
    const isFundingIn = typeVal.indexOf('تمويل') !== -1 && typeVal.indexOf('سداد تمويل') === -1;

    if (
      typeVal.indexOf('تحصيل') !== -1 ||     // تحصيل إيراد
      isFundingIn ||                          // تمويل (قرض/دعم داخل الحساب) - بدون سداد تمويل
      typeVal.indexOf('استرداد') !== -1      // استرداد تأمين من القناة
    ) {
      debitAcc = amount;
    }
    // فلوس خارجة من الحساب (أي دفعة / سداد / تأمين مدفوع…)
    else {
      creditAcc = amount;
    }

    if (!debitAcc && !creditAcc) continue;

    // 8) تحديث رصيد الحساب
    acc.balance += debitAcc - creditAcc;

    // 9) وصف الحركة
    let statement = '';
    if (party && detailsVal) statement = party + ' - ' + detailsVal;
    else if (party) statement = party;
    else if (detailsVal) statement = detailsVal;
    else statement = typeVal;

    acc.rows.push([
      date,
      statement,
      transNo,
      refNo,
      debitAcc,
      creditAcc,
      acc.balance,
      notes
    ]);
  }

  // 10) تفريغ وكتابة الشيتات
  Object.values(accounts).forEach(acc => {
    const sheet = acc.sheet;
    if (!sheet) return;

    const maxRows = sheet.getMaxRows();
    if (maxRows > 1) {
      sheet.getRange(2, 1, maxRows - 1, 8).clearContent();
    }

    if (!acc.rows.length) return;

    sheet.getRange(2, 1, acc.rows.length, 8).setValues(acc.rows);
    sheet.getRange(2, 1, acc.rows.length, 1).setNumberFormat('dd/mm/yyyy');
    sheet.getRange(2, 5, acc.rows.length, 3).setNumberFormat('#,##0.00');
  });

  if (!silent) {
    ui.alert('✅ تم تحديث شيتات البنك وخزنة العهدة وحساب البطاقة من "دفتر الحركات المالية" بنجاح.');
  }
  return { success: true, name: 'البنوك والخزنة والبطاقة' };
}
// ==================== شيتات مطابقة البنك (دولار / ليرة) ====================

// إنشاء شيت مطابقة (نفس الشيت نلصق فيه كشف البنك وتظهر فيه النتيجة)
function createBankReconciliationSheet_(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }

  sheet.clear();
  sheet.setRightToLeft(false);

  const headers = ["Date", "Amount", "System Balance", "Bank Amount", "Match Status"];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground("#263238")
    .setFontColor("white")
    .setFontWeight("bold")
    .setHorizontalAlignment("center");

  sheet.setColumnWidth(1, 120); // Date
  sheet.setColumnWidth(2, 120); // Amount
  sheet.setColumnWidth(3, 140); // System Balance
  sheet.setColumnWidth(4, 140); // Bank Amount
  sheet.setColumnWidth(5, 200); // Status

  sheet.setFrozenRows(1);

  sheet.getRange("A2").setNote(
    "➡ الصق هنا كشف البنك الشهري (Date في العمود A / Amount في العمود B)\n" +
    "ثم من القائمة اختر أمر المطابقة للعملة المناسبة."
  );
}

// إنشاء شيت مطابقة البنك دولار
function createBankReconciliationUsdSheet() {
  createBankReconciliationSheet_("مطابقة البنك - دولار");
}

// إنشاء شيت مطابقة البنك ليرة
function createBankReconciliationTrySheet() {
  createBankReconciliationSheet_("مطابقة البنك - ليرة");
}

// توليد مفتاح موحّد من التاريخ + المبلغ (للاستخدام الداخلي فقط)
function makeReconcileKey_(date, amount) {
  if (!date || amount === "" || amount === null) return "";
  const tz = Session.getScriptTimeZone();
  const dStr = Utilities.formatDate(new Date(date), tz, "yyyy-MM-dd");
  const amt = Math.round((Number(amount) || 0) * 100) / 100; // تقريب لرقمين
  return dStr + "|" + amt.toFixed(2);
}

// قراءة بيانات حساب البنك من شيت النظام (حساب البنك - دولار / ليرة)
function getSystemBankMapForCurrency_(currency) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const sheetName = (currency === "USD")
    ? "حساب البنك - دولار"
    : "حساب البنك - ليرة";

  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return {};

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return {};

  // نتوقع هيكل شيت البنك:
  // A: Date, B: Statement, C: Trans No, D: Ref,
  // E: Debit, F: Credit, G: Balance, H: Notes
  const values = sheet.getRange(2, 1, lastRow - 1, 7).getValues();

  const map = {}; // key -> { balance, count }

  values.forEach(r => {
    const date = r[0];
    const debit = Number(r[4]) || 0;
    const credit = Number(r[5]) || 0;
    const balance = Number(r[6]) || 0;

    // نستخدم القيمة المطلقة للحركة (المبلغ الإيجابي)
    const amount = debit > 0 ? debit : (credit > 0 ? credit : 0);
    if (!date || !amount) return;

    const key = makeReconcileKey_(date, amount);
    if (!key) return;

    if (!map[key]) {
      map[key] = { balance: balance, count: 1 };
    } else {
      map[key].count++;
      map[key].balance = balance; // آخر رصيد لنفس الحركة
    }
  });

  return map;
}
// قراءة بيانات حساب البطاقة من شيت النظام "حساب البطاقة - ليرة"
function getSystemCardMap_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const sheetName = "حساب البطاقة - ليرة";
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return {};

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return {};

  // نتوقع نفس هيكل شيت البنك:
  // A: Date, B: Statement, C: Trans No, D: Ref,
  // E: Debit, F: Credit, G: Balance, H: Notes
  const values = sheet.getRange(2, 1, lastRow - 1, 7).getValues();

  const map = {}; // key -> { balance, count }

  values.forEach(r => {
    const date = r[0];
    const debit = Number(r[4]) || 0;
    const credit = Number(r[5]) || 0;
    const balance = Number(r[6]) || 0;

    // نستخدم القيمة المطلقة للحركة (المبلغ الإيجابي)
    const amount = debit > 0 ? debit : (credit > 0 ? credit : 0);
    if (!date || !amount) return;

    const key = makeReconcileKey_(date, amount);
    if (!key) return;

    if (!map[key]) {
      map[key] = { balance: balance, count: 1 };
    } else {
      map[key].count++;
      map[key].balance = balance; // آخر رصيد لنفس الحركة
    }
  });

  return map;
}
// دالة عامة للمطابقة لحساب البنك لعملة معينة
function bankReconcileForCurrency_(currency) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const sheetName = (currency === "USD")
    ? "مطابقة البنك - دولار"
    : "مطابقة البنك - ليرة";

  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    ui.alert("❗ شيت '" + sheetName + "' غير موجود.\nأنشئه أولاً (أو شغّل createBankReconciliationUsdSheet/Try من المحرّر).");
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    ui.alert(
      "⚠️ لا توجد بيانات بنك للمطابقة في '" + sheetName + "'.\n\n" +
      "رجاءً الصق كشف البنك الشهري (التاريخ في العمود A، المبلغ في العمود B) ثم أعد تشغيل المطابقة."
    );
    return;
  }

  // نقرأ بيانات البنك (اللي أنت لاصقها)
  const bankData = sheet.getRange(2, 1, lastRow - 1, 2).getValues(); // A:Date, B:Amount

  // نجيب خريطة النظام
  const sysMap = getSystemBankMapForCurrency_(currency);

  const headers = ["Date", "Amount", "System Balance", "Bank Amount", "Match Status"];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  const rowsOut = [];
  const statusColors = [];

  bankData.forEach(row => {
    const date = row[0];
    const amt = Number(row[1]) || 0;

    let systemBalance = "";
    let bankAmount = "";
    let status = "";

    if (!date || !amt) {
      // صف فاضي أو ناقص
      rowsOut.push([date || "", amt || "", "", "", "⚠️ بيانات ناقصة"]);
      statusColors.push("#ffcdd2"); // أحمر فاتح
      return;
    }

    const key = makeReconcileKey_(date, amt);
    const info = sysMap[key];

    bankAmount = amt;

    if (!info) {
      // موجود في كشف البنك بس مش موجود في النظام
      status = "❌ Bank only (غير مسجّل في النظام)";
      rowsOut.push([date, amt, "", bankAmount, status]);
      statusColors.push("#ffcdd2"); // أحمر فاتح
    } else if (info.count > 1) {
      // نفس التاريخ والمبلغ مكرر في النظام
      systemBalance = info.balance;
      status = "⚠️ Duplicate in system (مكرر في النظام)";
      rowsOut.push([date, amt, systemBalance, bankAmount, status]);
      statusColors.push("#fff9c4"); // أصفر فاتح
    } else {
      // مطابق 1:1
      systemBalance = info.balance;
      status = "✔ MATCH";
      rowsOut.push([date, amt, systemBalance, bankAmount, status]);
      statusColors.push("#c8e6c9"); // أخضر فاتح
    }
  });

  // نفضي المحتوى واللون القديم من A2:E
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, 5).clearContent().clearFormat();
  }

  const n = rowsOut.length;
  if (n > 0) {
    sheet.getRange(2, 1, n, 5).setValues(rowsOut);
    sheet.getRange(2, 1, n, 1).setNumberFormat("yyyy-MM-dd");
    sheet.getRange(2, 2, n, 3).setNumberFormat("#,##0.00");

    // تلوين حالة المطابقة
    const statusRange = sheet.getRange(2, 5, n, 1);
    const backgrounds = statusColors.map(c => [c || null]);
    statusRange.setBackgrounds(backgrounds);
  }

  ui.alert(
    "✅ انتهت المطابقة لحساب البنك " +
    (currency === "USD" ? "بالدولار" : "بالليرة") + ".\n\n" +
    "الأحمر = فروقات بين النظام وكشف البنك.\n" +
    "الأخضر = حركات متطابقة.\n" +
    "الأصفر = حركة مكررة في النظام."
  );
}

function reconcileCard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const sheetName = "مطابقة الكارت";
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    ui.alert("❗ شيت '" + sheetName + "' غير موجود.\nأنشئه أولاً عن طريق createCardReconciliationSheet من المحرّر.");
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    ui.alert(
      "⚠️ لا توجد بيانات كارت للمطابقة في '" + sheetName + "'.\n\n" +
      "رجاءً الصق كشف الكارت (التاريخ في العمود A، المبلغ في العمود B) ثم أعد تشغيل المطابقة."
    );
    return;
  }

  // نقرأ بيانات كشف الكارت (اللي أنت لاصقها)
  const cardData = sheet.getRange(2, 1, lastRow - 1, 2).getValues(); // A:Date, B:Amount

  // نجيب خريطة النظام من شيت "حساب البطاقة - ليرة"
  const sysMap = getSystemCardMap_();

  const headers = ["Date", "Amount", "System Balance", "Card Amount", "Match Status"];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground("#263238")
    .setFontColor("white")
    .setFontWeight("bold")
    .setHorizontalAlignment("center");

  const rowsOut = [];
  const statusColors = [];

  cardData.forEach(row => {
    const date = row[0];
    const amt = Number(row[1]) || 0;

    let systemBalance = "";
    let cardAmount = "";
    let status = "";

    if (!date || !amt) {
      // صف فاضي أو ناقص
      rowsOut.push([date || "", amt || "", "", "", "⚠️ بيانات ناقصة"]);
      statusColors.push("#ffcdd2"); // أحمر فاتح
      return;
    }

    const key = makeReconcileKey_(date, amt);
    const info = sysMap[key];

    cardAmount = amt;

    if (!info) {
      // موجود في كشف الكارت بس مش موجود في النظام
      status = "❌ Card only (غير مسجّل في النظام)";
      rowsOut.push([date, amt, "", cardAmount, status]);
      statusColors.push("#ffcdd2"); // أحمر فاتح
    } else if (info.count > 1) {
      // نفس التاريخ والمبلغ مكرر في النظام
      systemBalance = info.balance;
      status = "⚠️ Duplicate in system (مكرر في النظام)";
      rowsOut.push([date, amt, systemBalance, cardAmount, status]);
      statusColors.push("#fff9c4"); // أصفر فاتح
    } else {
      // مطابق 1:1
      systemBalance = info.balance;
      status = "✔ MATCH";
      rowsOut.push([date, amt, systemBalance, cardAmount, status]);
      statusColors.push("#c8e6c9"); // أخضر فاتح
    }
  });

  // نفضي المحتوى واللون القديم من A2:E
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, 5).clearContent().clearFormat();
  }

  const n = rowsOut.length;
  if (n > 0) {
    sheet.getRange(2, 1, n, 5).setValues(rowsOut);
    sheet.getRange(2, 1, n, 1).setNumberFormat("yyyy-MM-dd");
    sheet.getRange(2, 2, n, 3).setNumberFormat("#,##0.00");

    // تلوين حالة المطابقة
    const statusRange = sheet.getRange(2, 5, n, 1);
    const backgrounds = statusColors.map(c => [c || null]);
    statusRange.setBackgrounds(backgrounds);
  }

  ui.alert(
    "✅ انتهت المطابقة لحساب الكارت.\n\n" +
    "الأحمر = فروقات بين النظام وكشف الكارت.\n" +
    "الأخضر = حركات متطابقة.\n" +
    "الأصفر = حركة مكررة في النظام."
  );
}

// دوال مختصرة للمنيو (تتوافق مع onOpen الجديد)
function reconcileBankUsd() {
  bankReconcileForCurrency_("USD");
}

function reconcileBankTry() {
  bankReconcileForCurrency_("TRY");
}

// ==================== تحديث لوحة التحكم ====================

function refreshDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // 1) تحديث البنوك والخزنة + كل التقارير الملخصة (بوضع صامت)
  const reportResults = rebuildAllSummaryReports(true);

  // 2) إعادة بناء لوحة التحكم من جديد
  createDashboardSheet(ss);

  // 3) إظهار رسالة واحدة شاملة بالنتائج
  const successList = reportResults.filter(r => r && r.success).map(r => '  ✅ ' + r.name);
  const errorList = reportResults.filter(r => r && !r.success).map(r => '  ❌ ' + r.name + ': ' + r.error);

  let message = '════════════════════════════════════\n';
  message += '   📊 تحديث شامل للنظام\n';
  message += '════════════════════════════════════\n\n';

  // إضافة لوحة التحكم للقائمة
  successList.push('  ✅ لوحة التحكم');

  if (successList.length) {
    message += 'تم بنجاح:\n' + successList.join('\n') + '\n';
  }
  if (errorList.length) {
    message += '\nفشل:\n' + errorList.join('\n');
  }

  message += '\n════════════════════════════════════';

  ui.alert(message);
}

// ==================== 🔍 مراجعة وإصلاح نوع الحركة ====================

/**
 * مراجعة وإصلاح الربط بين طبيعة الحركة (C) ونوع الحركة (N)
 * القاعدة: استحقاق = مدين استحقاق | غير ذلك = دائن دفعة
 */
function reviewAndFixMovementTypes() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.TRANSACTIONS);

  if (!sheet) {
    ui.alert('⚠️ لم يتم العثور على دفتر الحركات المالية');
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    ui.alert('ℹ️ لا توجد بيانات للمراجعة');
    return;
  }

  // قراءة عمود A (رقم الحركة) و C (طبيعة الحركة) و N (نوع الحركة)
  // A = 1, C = 3, N = 14
  const dataA = sheet.getRange(2, 1, lastRow - 1, 1).getValues();   // رقم الحركة
  const dataC = sheet.getRange(2, 3, lastRow - 1, 1).getValues();   // طبيعة الحركة
  const dataN = sheet.getRange(2, 14, lastRow - 1, 1).getValues();  // نوع الحركة

  let errors = [];
  let fixes = [];

  for (let i = 0; i < dataC.length; i++) {
    const transNum = dataA[i][0] || '';  // رقم الحركة من عمود A
    const natureValue = String(dataC[i][0] || '').trim();
    const currentType = String(dataN[i][0] || '').trim();

    // تخطي الصفوف الفارغة
    if (!natureValue) continue;

    // تحديد النوع الصحيح باستخدام الدالة المركزية
    const correctType = getMovementTypeFromNature_(natureValue);

    // تخطي إذا لم يتم التعرف على طبيعة الحركة
    if (!correctType) continue;

    // التحقق من التطابق
    if (currentType !== correctType) {
      errors.push({
        transNum: transNum,  // رقم الحركة
        nature: natureValue,
        current: currentType || '(فارغ)',
        correct: correctType
      });
      fixes.push([correctType]);
    } else {
      fixes.push([currentType]); // لا تغيير
    }
  }

  // عرض النتائج
  if (errors.length === 0) {
    ui.alert('✅ مراجعة مكتملة\n\nكل البيانات صحيحة، لا توجد أخطاء للإصلاح.');
    return;
  }

  // إعداد رسالة التقرير
  let reportMsg = '🔍 تقرير المراجعة\n\n';
  reportMsg += 'تم العثور على ' + errors.length + ' خطأ:\n\n';

  // عرض أول 10 أخطاء فقط في الرسالة
  const showErrors = errors.slice(0, 10);
  showErrors.forEach((err, idx) => {
    reportMsg += (idx + 1) + '. حركة #' + err.transNum + ': ';
    reportMsg += err.nature.substring(0, 20) + '\n';
    reportMsg += '   الحالي: ' + err.current + ' ← الصحيح: ' + err.correct + '\n';
  });

  if (errors.length > 10) {
    reportMsg += '\n... و ' + (errors.length - 10) + ' أخطاء أخرى\n';
  }

  reportMsg += '\nهل تريد إصلاح كل الأخطاء؟';

  const response = ui.alert('مراجعة نوع الحركة', reportMsg, ui.ButtonSet.YES_NO);

  if (response === ui.Button.YES) {
    // تطبيق الإصلاحات
    sheet.getRange(2, 14, fixes.length, 1).setValues(fixes);
    ui.alert('✅ تم الإصلاح\n\nتم إصلاح ' + errors.length + ' خطأ بنجاح.');
  } else {
    ui.alert('ℹ️ تم إلغاء الإصلاح\n\nلم يتم تعديل أي بيانات.');
  }
}

/**
 * مراجعة فقط بدون إصلاح (تقرير)
 */
function reviewMovementTypesOnly() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.TRANSACTIONS);

  if (!sheet) {
    ui.alert('⚠️ لم يتم العثور على دفتر الحركات المالية');
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    ui.alert('ℹ️ لا توجد بيانات للمراجعة');
    return;
  }

  // قراءة عمود C (طبيعة الحركة) و N (نوع الحركة)
  const dataC = sheet.getRange(2, 3, lastRow - 1, 1).getValues();
  const dataN = sheet.getRange(2, 14, lastRow - 1, 1).getValues();

  let errorCount = 0;
  let correctCount = 0;
  let emptyCount = 0;

  for (let i = 0; i < dataC.length; i++) {
    const natureValue = String(dataC[i][0] || '').trim();
    const currentType = String(dataN[i][0] || '').trim();

    if (!natureValue) {
      emptyCount++;
      continue;
    }

    // تحديد النوع الصحيح باستخدام الدالة المركزية
    const correctType = getMovementTypeFromNature_(natureValue);

    // تخطي إذا لم يتم التعرف على طبيعة الحركة
    if (!correctType) {
      emptyCount++;
      continue;
    }

    if (currentType === correctType) {
      correctCount++;
    } else {
      errorCount++;
    }
  }

  let reportMsg = '📊 تقرير مراجعة نوع الحركة\n\n';
  reportMsg += '━━━━━━━━━━━━━━━━━━━━\n';
  reportMsg += '✅ صحيح: ' + correctCount + ' حركة\n';
  reportMsg += '❌ خطأ: ' + errorCount + ' حركة\n';
  reportMsg += '⬜ فارغ: ' + emptyCount + ' صف\n';
  reportMsg += '━━━━━━━━━━━━━━━━━━━━\n';
  reportMsg += '📝 الإجمالي: ' + (correctCount + errorCount) + ' حركة\n';

  if (errorCount > 0) {
    reportMsg += '\n💡 لإصلاح الأخطاء، استخدم:\nمراجعة وإصلاح نوع الحركة';
  }

  ui.alert('تقرير المراجعة', reportMsg, ui.ButtonSet.OK);
}

// ==================== 🔍 مراجعة الاستحقاقات والدفعات ====================

/**
 * فحص سريع: عرض المشاكل فقط (دفعات بدون استحقاق كافي)
 */
function checkAccrualPaymentBalance() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.TRANSACTIONS);

  if (!sheet) {
    ui.alert('⚠️ لم يتم العثور على دفتر الحركات المالية');
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    ui.alert('ℹ️ لا توجد بيانات للمراجعة');
    return;
  }

  // قراءة البيانات: I (الطرف), M (القيمة بالدولار), N (نوع الحركة)
  // I = 9, M = 13, N = 14
  const data = sheet.getRange(2, 1, lastRow - 1, 14).getValues();

  // تجميع حسب الطرف
  const parties = {};

  for (let i = 0; i < data.length; i++) {
    const partyName = String(data[i][8] || '').trim();  // عمود I (index 8)
    const amountUsd = Number(data[i][12]) || 0;         // عمود M (index 12)
    const movementType = String(data[i][13] || '');     // عمود N (index 13)

    if (!partyName) continue;

    if (!parties[partyName]) {
      parties[partyName] = { accruals: 0, payments: 0 };
    }

    if (movementType.indexOf('مدين استحقاق') !== -1) {
      parties[partyName].accruals += amountUsd;
    } else if (movementType.indexOf('دائن دفعة') !== -1) {
      parties[partyName].payments += amountUsd;
    }
  }

  // البحث عن المشاكل
  const problems = [];
  let healthyCount = 0;

  for (const partyName in parties) {
    const p = parties[partyName];
    const balance = p.accruals - p.payments;

    if (balance < -0.01) {  // سماح بفرق بسيط للتقريب
      problems.push({
        name: partyName,
        accruals: p.accruals,
        payments: p.payments,
        excess: Math.abs(balance)
      });
    } else {
      healthyCount++;
    }
  }

  // عرض النتائج
  if (problems.length === 0) {
    ui.alert('✅ فحص مكتمل',
      'كل الأطراف سليمين!\n\n' +
      '📊 تم فحص ' + (healthyCount) + ' طرف\n' +
      'لا توجد دفعات تتجاوز الاستحقاقات.',
      ui.ButtonSet.OK);
    return;
  }

  // إعداد تقرير المشاكل
  let reportMsg = '⚠️ تم العثور على ' + problems.length + ' مشكلة:\n\n';

  const showProblems = problems.slice(0, 8);
  showProblems.forEach((prob, idx) => {
    if (prob.accruals === 0) {
      reportMsg += (idx + 1) + '. ' + prob.name + '\n';
      reportMsg += '   ❌ دفعات $' + prob.payments.toFixed(2) + ' بدون أي استحقاق!\n\n';
    } else {
      reportMsg += (idx + 1) + '. ' + prob.name + '\n';
      reportMsg += '   استحقاق: $' + prob.accruals.toFixed(2);
      reportMsg += ' | دفعات: $' + prob.payments.toFixed(2) + '\n';
      reportMsg += '   ❌ زيادة: $' + prob.excess.toFixed(2) + '\n\n';
    }
  });

  if (problems.length > 8) {
    reportMsg += '... و ' + (problems.length - 8) + ' مشاكل أخرى\n\n';
  }

  reportMsg += '━━━━━━━━━━━━━━━━━━━━\n';
  reportMsg += '✅ سليم: ' + healthyCount + ' طرف\n';
  reportMsg += '❌ مشاكل: ' + problems.length + ' طرف\n';
  reportMsg += '\n💡 للتفاصيل الكاملة استخدم:\nتقرير الاستحقاقات والدفعات (شيت)';

  ui.alert('نتيجة الفحص', reportMsg, ui.ButtonSet.OK);
}

/**
 * تقرير شامل: إنشاء شيت بكل الأطراف واستحقاقاتهم ودفعاتهم
 */
function generateAccrualPaymentReport() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const transSheet = ss.getSheetByName(CONFIG.SHEETS.TRANSACTIONS);

  if (!transSheet) {
    ui.alert('⚠️ لم يتم العثور على دفتر الحركات المالية');
    return;
  }

  const lastRow = transSheet.getLastRow();
  if (lastRow < 2) {
    ui.alert('ℹ️ لا توجد بيانات للتقرير');
    return;
  }

  // قراءة البيانات
  const data = transSheet.getRange(2, 1, lastRow - 1, 14).getValues();

  // تجميع حسب الطرف
  const parties = {};

  for (let i = 0; i < data.length; i++) {
    const partyName = String(data[i][8] || '').trim();  // عمود I
    const amountUsd = Number(data[i][12]) || 0;         // عمود M
    const movementType = String(data[i][13] || '');     // عمود N

    if (!partyName) continue;

    if (!parties[partyName]) {
      parties[partyName] = { accruals: 0, payments: 0, transCount: 0 };
    }

    parties[partyName].transCount++;

    if (movementType.indexOf('مدين استحقاق') !== -1) {
      parties[partyName].accruals += amountUsd;
    } else if (movementType.indexOf('دائن دفعة') !== -1) {
      parties[partyName].payments += amountUsd;
    }
  }

  // إنشاء أو إعادة استخدام الشيت
  const reportSheetName = 'تقرير الاستحقاقات والدفعات';
  let reportSheet = ss.getSheetByName(reportSheetName);

  if (reportSheet) {
    reportSheet.clear();
  } else {
    reportSheet = ss.insertSheet(reportSheetName);
  }

  // إعداد الهيدر
  const headers = ['الطرف', 'الاستحقاقات ($)', 'الدفعات ($)', 'المتبقي ($)', 'الحالة', 'عدد الحركات'];
  reportSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  reportSheet.getRange(1, 1, 1, headers.length)
    .setBackground('#4a86e8')
    .setFontColor('white')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  // تحويل البيانات إلى مصفوفة وترتيبها
  const rows = [];
  let problemCount = 0;
  let healthyCount = 0;

  for (const partyName in parties) {
    const p = parties[partyName];
    const balance = p.accruals - p.payments;
    let status = '✅ سليم';

    if (balance < -0.01) {
      if (p.accruals === 0) {
        status = '❌ دفعة بدون استحقاق';
      } else {
        status = '❌ دفعات زائدة';
      }
      problemCount++;
    } else if (Math.abs(balance) < 0.01) {
      status = '✅ مسدد بالكامل';
      healthyCount++;
    } else {
      status = '✅ متبقي للسداد';
      healthyCount++;
    }

    rows.push([
      partyName,
      p.accruals,
      p.payments,
      balance,
      status,
      p.transCount
    ]);
  }

  // ترتيب: المشاكل أولاً، ثم حسب الاسم
  rows.sort((a, b) => {
    const aHasProblem = a[4].indexOf('❌') !== -1 ? 0 : 1;
    const bHasProblem = b[4].indexOf('❌') !== -1 ? 0 : 1;
    if (aHasProblem !== bHasProblem) return aHasProblem - bHasProblem;
    return a[0].localeCompare(b[0], 'ar');
  });

  // كتابة البيانات
  if (rows.length > 0) {
    reportSheet.getRange(2, 1, rows.length, headers.length).setValues(rows);

    // تنسيق الأرقام
    reportSheet.getRange(2, 2, rows.length, 3).setNumberFormat('$#,##0.00');

    // تلوين صفوف المشاكل
    for (let i = 0; i < rows.length; i++) {
      if (rows[i][4].indexOf('❌') !== -1) {
        reportSheet.getRange(i + 2, 1, 1, headers.length).setBackground('#ffcccc');
      }
    }
  }

  // إضافة صف الملخص
  const summaryRow = rows.length + 3;
  reportSheet.getRange(summaryRow, 1).setValue('📊 الملخص:');
  reportSheet.getRange(summaryRow, 1).setFontWeight('bold');
  reportSheet.getRange(summaryRow + 1, 1).setValue('✅ سليم: ' + healthyCount + ' طرف');
  reportSheet.getRange(summaryRow + 2, 1).setValue('❌ مشاكل: ' + problemCount + ' طرف');
  reportSheet.getRange(summaryRow + 3, 1).setValue('📝 الإجمالي: ' + rows.length + ' طرف');

  // تعديل عرض الأعمدة
  reportSheet.setColumnWidth(1, 200);  // الطرف
  reportSheet.setColumnWidth(2, 120);  // الاستحقاقات
  reportSheet.setColumnWidth(3, 120);  // الدفعات
  reportSheet.setColumnWidth(4, 120);  // المتبقي
  reportSheet.setColumnWidth(5, 150);  // الحالة
  reportSheet.setColumnWidth(6, 100);  // عدد الحركات

  reportSheet.setFrozenRows(1);

  // الانتقال للشيت
  ss.setActiveSheet(reportSheet);

  ui.alert('✅ تم إنشاء التقرير',
    'تم إنشاء تقرير الاستحقاقات والدفعات.\n\n' +
    '📊 الملخص:\n' +
    '• سليم: ' + healthyCount + ' طرف\n' +
    '• مشاكل: ' + problemCount + ' طرف\n' +
    '• الإجمالي: ' + rows.length + ' طرف',
    ui.ButtonSet.OK);
}

// ==================== 📊 تقرير ميزانية المشروع التفصيلي ====================

/**
 * إنشاء تقرير ميزانية مشروع تفصيلي
 * يعرض الميزانية المخططة vs الفعلية + عمولة مدير المشروعات
 */
function generateProjectBudgetReport() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1️⃣ طلب كود المشروع من المستخدم
  const response = ui.prompt(
    '📊 تقرير ميزانية المشروع',
    'أدخل كود المشروع (مثال: PRJ-001):',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) return;

  const projectCode = response.getResponseText().trim().toUpperCase();
  if (!projectCode) {
    ui.alert('⚠️ لم يتم إدخال كود المشروع');
    return;
  }

  // 2️⃣ قراءة بيانات المشروع
  const projectsSheet = ss.getSheetByName(CONFIG.SHEETS.PROJECTS);
  if (!projectsSheet) {
    ui.alert('⚠️ لم يتم العثور على قاعدة بيانات المشاريع');
    return;
  }

  const projectsData = projectsSheet.getDataRange().getValues();
  const projHeaders = projectsData[0];

  // البحث عن أعمدة المشروع
  const projCodeCol = 0; // A
  const projNameCol = 1; // B
  const channelCol = 3;  // D - القناة/الجهة
  const contractValueCol = 8; // I - قيمة العقد
  const fundingValueCol = 7;  // H - قيمة التمويل

  // البحث عن عمود مدير المشروعات ونسبة العمولة
  const managerColIdx = projHeaders.indexOf('مدير المشروعات');
  const commissionRateColIdx = projHeaders.indexOf('نسبة العمولة');

  let projectInfo = null;
  for (let i = 1; i < projectsData.length; i++) {
    if (String(projectsData[i][projCodeCol]).trim().toUpperCase() === projectCode) {
      projectInfo = {
        code: projectsData[i][projCodeCol],
        name: projectsData[i][projNameCol],
        channel: projectsData[i][channelCol] || '',
        contractValue: Number(projectsData[i][contractValueCol]) || 0,
        fundingValue: Number(projectsData[i][fundingValueCol]) || 0,
        manager: managerColIdx !== -1 ? (projectsData[i][managerColIdx] || '') : '',
        commissionRate: commissionRateColIdx !== -1 ? (Number(projectsData[i][commissionRateColIdx]) || 0) : 0
      };
      break;
    }
  }

  if (!projectInfo) {
    ui.alert('⚠️ لم يتم العثور على مشروع بكود: ' + projectCode);
    return;
  }

  // 3️⃣ قراءة الميزانية المخططة
  const budgetSheet = ss.getSheetByName(CONFIG.SHEETS.BUDGETS);
  const plannedBudget = {}; // { البند: المبلغ المخطط }
  let totalPlanned = 0;

  if (budgetSheet && budgetSheet.getLastRow() > 1) {
    const budgetData = budgetSheet.getDataRange().getValues();
    for (let i = 1; i < budgetData.length; i++) {
      const budgetProjCode = String(budgetData[i][0]).trim().toUpperCase();
      if (budgetProjCode === projectCode) {
        const item = String(budgetData[i][2] || '').trim();
        const amount = Number(budgetData[i][3]) || 0;
        if (item) {
          plannedBudget[item] = (plannedBudget[item] || 0) + amount;
          totalPlanned += amount;
        }
      }
    }
  }

  // 4️⃣ قراءة المصروفات الفعلية من دفتر الحركات
  const transSheet = ss.getSheetByName(CONFIG.SHEETS.TRANSACTIONS);
  if (!transSheet) {
    ui.alert('⚠️ لم يتم العثور على دفتر الحركات المالية');
    return;
  }

  const transData = transSheet.getDataRange().getValues();
  const transHeaders = transData[0];

  // تحديد أعمدة دفتر الحركات
  const colC = transHeaders.indexOf('طبيعة الحركة') !== -1 ? transHeaders.indexOf('طبيعة الحركة') : 2;
  const colE = transHeaders.indexOf('كود المشروع') !== -1 ? transHeaders.indexOf('كود المشروع') : 4;
  const colG = transHeaders.indexOf('البند') !== -1 ? transHeaders.indexOf('البند') : 6;
  const colI = transHeaders.indexOf('اسم المورد/الجهة') !== -1 ? transHeaders.indexOf('اسم المورد/الجهة') : 8;
  const colM = transHeaders.indexOf('القيمة بالدولار') !== -1 ? transHeaders.indexOf('القيمة بالدولار') : 12;

  const actualExpenses = {}; // { البند: { total: المبلغ, details: [{vendor, amount}] } }
  let totalActual = 0;
  let commissionAmount = 0;

  for (let i = 1; i < transData.length; i++) {
    const rowProjCode = String(transData[i][colE] || '').trim().toUpperCase();
    if (rowProjCode !== projectCode) continue;

    const item = String(transData[i][colG] || '').trim();
    const vendor = String(transData[i][colI] || '').trim();
    const amountUsd = Number(transData[i][colM]) || 0;
    const natureType = String(transData[i][colC] || '').trim();

    // فقط المصروفات (استحقاق مصروف) - استبعاد الإيرادات
    if (natureType.indexOf('استحقاق مصروف') === -1) continue;
    if (!item || amountUsd <= 0) continue;

    // تجميع حسب البند
    if (!actualExpenses[item]) {
      actualExpenses[item] = { total: 0, details: [] };
    }
    actualExpenses[item].total += amountUsd;
    actualExpenses[item].details.push({ vendor, amount: amountUsd });
    totalActual += amountUsd;

    // حساب عمولة مدير المشروعات
    if (item.indexOf('عمولة مدير') !== -1) {
      commissionAmount += amountUsd;
    }
  }

  // 5️⃣ حساب العمولة المتوقعة (إذا كان هناك مدير ونسبة)
  let expectedCommission = 0;
  if (projectInfo.manager && projectInfo.commissionRate > 0 && projectInfo.contractValue > 0) {
    expectedCommission = projectInfo.contractValue * (projectInfo.commissionRate / 100);
  }

  // 6️⃣ إنشاء شيت التقرير
  const reportSheetName = 'تقرير ميزانية - ' + projectCode;
  let reportSheet = ss.getSheetByName(reportSheetName);
  if (reportSheet) {
    ss.deleteSheet(reportSheet);
  }
  reportSheet = ss.insertSheet(reportSheetName);
  reportSheet.setRightToLeft(true);

  let currentRow = 1;

  // ═══════════════════════════════════════════════════════════
  // العنوان الرئيسي
  // ═══════════════════════════════════════════════════════════
  reportSheet.getRange(currentRow, 1, 1, 6).merge()
    .setValue('📊 تقرير ميزانية المشروع: ' + projectInfo.code)
    .setBackground('#1a237e')
    .setFontColor('white')
    .setFontWeight('bold')
    .setFontSize(16)
    .setHorizontalAlignment('center');
  currentRow++;

  reportSheet.getRange(currentRow, 1, 1, 6).merge()
    .setValue(projectInfo.name)
    .setBackground('#283593')
    .setFontColor('white')
    .setFontSize(14)
    .setHorizontalAlignment('center');
  currentRow += 2;

  // ═══════════════════════════════════════════════════════════
  // بيانات المشروع الأساسية
  // ═══════════════════════════════════════════════════════════
  reportSheet.getRange(currentRow, 1, 1, 6).merge()
    .setValue('📋 بيانات المشروع')
    .setBackground('#e8eaf6')
    .setFontWeight('bold')
    .setFontSize(12);
  currentRow++;

  const projectInfoData = [
    ['القناة/الجهة', projectInfo.channel, 'قيمة العقد', projectInfo.contractValue, 'قيمة التمويل', projectInfo.fundingValue]
  ];
  reportSheet.getRange(currentRow, 1, 1, 6).setValues(projectInfoData);
  reportSheet.getRange(currentRow, 4).setNumberFormat('$#,##0.00');
  reportSheet.getRange(currentRow, 6).setNumberFormat('$#,##0.00');
  currentRow += 2;

  // ═══════════════════════════════════════════════════════════
  // جدول المقارنة: الميزانية المخططة vs الفعلية
  // ═══════════════════════════════════════════════════════════
  reportSheet.getRange(currentRow, 1, 1, 6).merge()
    .setValue('📈 الميزانية المخططة vs الفعلية')
    .setBackground('#e8eaf6')
    .setFontWeight('bold')
    .setFontSize(12);
  currentRow++;

  // رؤوس الأعمدة
  const budgetHeaders = ['البند', 'المخطط', 'الفعلي', 'الفرق', 'النسبة %', 'الحالة'];
  reportSheet.getRange(currentRow, 1, 1, 6).setValues([budgetHeaders])
    .setBackground('#3949ab')
    .setFontColor('white')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  currentRow++;

  // جمع كل البنود (من المخطط والفعلي)
  const allItems = new Set([...Object.keys(plannedBudget), ...Object.keys(actualExpenses)]);
  const budgetRows = [];

  allItems.forEach(item => {
    // تجاهل عمولة مدير المشروعات (سيظهر في قسم منفصل)
    if (item.indexOf('عمولة مدير') !== -1) return;

    const planned = plannedBudget[item] || 0;
    const actual = actualExpenses[item] ? actualExpenses[item].total : 0;
    const diff = planned - actual;
    const percentage = planned > 0 ? Math.round((actual / planned) * 100) : (actual > 0 ? '∞' : 0);

    let status = '';
    if (planned === 0 && actual > 0) {
      status = '⚠️ غير مخطط';
    } else if (percentage === '∞' || percentage > 120) {
      status = '🔴 تجاوز';
    } else if (percentage > 100) {
      status = '🟡 تجاوز طفيف';
    } else if (percentage >= 80) {
      status = '🟢 ضمن الميزانية';
    } else if (actual === 0) {
      status = '⚪ لم يُصرف';
    } else {
      status = '🔵 وفر';
    }

    budgetRows.push([item, planned, actual, diff, percentage === '∞' ? '∞' : percentage + '%', status]);
  });

  // ترتيب حسب الفعلي تنازلياً
  budgetRows.sort((a, b) => b[2] - a[2]);

  if (budgetRows.length > 0) {
    reportSheet.getRange(currentRow, 1, budgetRows.length, 6).setValues(budgetRows);
    reportSheet.getRange(currentRow, 2, budgetRows.length, 3).setNumberFormat('$#,##0.00');

    // تلوين الفرق
    for (let i = 0; i < budgetRows.length; i++) {
      const diffCell = reportSheet.getRange(currentRow + i, 4);
      const diffValue = budgetRows[i][3];
      if (diffValue < 0) {
        diffCell.setFontColor('#c62828'); // أحمر للتجاوز
      } else if (diffValue > 0) {
        diffCell.setFontColor('#2e7d32'); // أخضر للوفر
      }
    }
    currentRow += budgetRows.length;

    // صف المجموع
    const totalDiff = totalPlanned - totalActual;
    const totalPercentage = totalPlanned > 0 ? Math.round((totalActual / totalPlanned) * 100) : (totalActual > 0 ? '∞' : 0);
    let totalStatus = '';
    if (totalPercentage === '∞' || totalPercentage > 120) {
      totalStatus = '🔴 تجاوز';
    } else if (totalPercentage > 100) {
      totalStatus = '🟡 تجاوز طفيف';
    } else if (totalPercentage >= 80) {
      totalStatus = '🟢 ضمن الميزانية';
    } else {
      totalStatus = '🔵 وفر';
    }

    const totalRow = ['📊 المجموع', totalPlanned, totalActual, totalDiff, totalPercentage === '∞' ? '∞' : totalPercentage + '%', totalStatus];
    reportSheet.getRange(currentRow, 1, 1, 6).setValues([totalRow])
      .setBackground('#e8eaf6')
      .setFontWeight('bold');
    reportSheet.getRange(currentRow, 2, 1, 3).setNumberFormat('$#,##0.00');

    // تلوين فرق المجموع
    const totalDiffCell = reportSheet.getRange(currentRow, 4);
    if (totalDiff < 0) {
      totalDiffCell.setFontColor('#c62828'); // أحمر للتجاوز
    } else if (totalDiff > 0) {
      totalDiffCell.setFontColor('#2e7d32'); // أخضر للوفر
    }
    currentRow++;
  } else {
    reportSheet.getRange(currentRow, 1, 1, 6).merge()
      .setValue('لا توجد بنود مسجلة')
      .setFontStyle('italic')
      .setHorizontalAlignment('center');
    currentRow++;
  }
  currentRow++;

  // ═══════════════════════════════════════════════════════════
  // عمولة مدير المشروعات
  // ═══════════════════════════════════════════════════════════
  reportSheet.getRange(currentRow, 1, 1, 6).merge()
    .setValue('💰 عمولة مدير المشروعات')
    .setBackground('#e8eaf6')
    .setFontWeight('bold')
    .setFontSize(12);
  currentRow++;

  if (projectInfo.manager) {
    const commissionHeaders = ['مدير المشروع', 'نسبة العمولة', 'العمولة المتوقعة', 'العمولة المسجلة', 'المتبقي', 'الحالة'];
    reportSheet.getRange(currentRow, 1, 1, 6).setValues([commissionHeaders])
      .setBackground('#7b1fa2')
      .setFontColor('white')
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
    currentRow++;

    const remaining = expectedCommission - commissionAmount;
    const commStatus = remaining <= 0 ? '✅ مكتمل' : '⏳ معلق';

    const commissionRow = [
      projectInfo.manager,
      projectInfo.commissionRate + '%',
      expectedCommission,
      commissionAmount,
      remaining > 0 ? remaining : 0,
      commStatus
    ];
    reportSheet.getRange(currentRow, 1, 1, 6).setValues([commissionRow]);
    reportSheet.getRange(currentRow, 3, 1, 3).setNumberFormat('$#,##0.00');
    currentRow++;
  } else {
    reportSheet.getRange(currentRow, 1, 1, 6).merge()
      .setValue('لا يوجد مدير مشروعات معين')
      .setFontStyle('italic')
      .setHorizontalAlignment('center');
    currentRow++;
  }
  currentRow++;

  // ═══════════════════════════════════════════════════════════
  // الملخص النهائي
  // ═══════════════════════════════════════════════════════════
  reportSheet.getRange(currentRow, 1, 1, 6).merge()
    .setValue('📊 الملخص النهائي')
    .setBackground('#e8eaf6')
    .setFontWeight('bold')
    .setFontSize(12);
  currentRow++;

  const totalWithCommission = totalActual; // العمولة مضمنة في المصروفات
  const budgetRemaining = totalPlanned - totalActual;

  const summaryData = [
    ['إجمالي الميزانية المخططة', totalPlanned, '', 'إجمالي المصروفات الفعلية', totalActual, ''],
    ['عمولة مدير المشروعات (مسجلة)', commissionAmount, '', 'عمولة مدير المشروعات (متوقعة)', expectedCommission, ''],
    ['', '', '', '', '', ''],
    ['الميزانية المتبقية', budgetRemaining, budgetRemaining >= 0 ? '✅' : '⚠️ تجاوز', 'نسبة الصرف', totalPlanned > 0 ? Math.round((totalActual / totalPlanned) * 100) + '%' : 'N/A', '']
  ];

  reportSheet.getRange(currentRow, 1, 4, 6).setValues(summaryData);
  reportSheet.getRange(currentRow, 2, 4, 1).setNumberFormat('$#,##0.00');
  reportSheet.getRange(currentRow, 5, 2, 1).setNumberFormat('$#,##0.00');

  // تلوين الميزانية المتبقية
  const remainingCell = reportSheet.getRange(currentRow + 3, 2);
  if (budgetRemaining < 0) {
    remainingCell.setFontColor('#c62828').setFontWeight('bold');
  } else {
    remainingCell.setFontColor('#2e7d32').setFontWeight('bold');
  }

  currentRow += 5;

  // ═══════════════════════════════════════════════════════════
  // تنسيقات عامة
  // ═══════════════════════════════════════════════════════════
  reportSheet.setColumnWidth(1, 180);
  reportSheet.setColumnWidth(2, 120);
  reportSheet.setColumnWidth(3, 120);
  reportSheet.setColumnWidth(4, 120);
  reportSheet.setColumnWidth(5, 100);
  reportSheet.setColumnWidth(6, 120);

  reportSheet.setFrozenRows(2);
  ss.setActiveSheet(reportSheet);

  // تاريخ التقرير
  reportSheet.getRange(currentRow, 1, 1, 6).merge()
    .setValue('تاريخ التقرير: ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm'))
    .setFontSize(9)
    .setFontColor('#666666')
    .setHorizontalAlignment('center');

  ui.alert(
    '✅ تم إنشاء تقرير ميزانية المشروع',
    'المشروع: ' + projectInfo.code + ' - ' + projectInfo.name + '\n\n' +
    '📋 الميزانية المخططة: $' + totalPlanned.toFixed(2) + '\n' +
    '💰 المصروفات الفعلية: $' + totalActual.toFixed(2) + '\n' +
    '📊 الفرق: $' + budgetRemaining.toFixed(2) + (budgetRemaining < 0 ? ' ⚠️' : ' ✅'),
    ui.ButtonSet.OK
  );
}

// ==================== 💰 نظام عمولات مديري المشروعات ====================

/**
 * إضافة أعمدة مدير المشروعات ونسبة العمولة لقاعدة بيانات المشاريع
 */
function addProjectManagerColumns() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const projectsSheet = ss.getSheetByName(CONFIG.SHEETS.PROJECTS);
  const partiesSheet = ss.getSheetByName(CONFIG.SHEETS.PARTIES);

  if (!projectsSheet) {
    ui.alert('⚠️ لم يتم العثور على قاعدة بيانات المشاريع');
    return;
  }

  if (!partiesSheet) {
    ui.alert('⚠️ لم يتم العثور على قاعدة بيانات الأطراف');
    return;
  }

  // التحقق من وجود الأعمدة مسبقاً (نبحث في الصف الأول)
  const headers = projectsSheet.getRange(1, 1, 1, 25).getValues()[0];
  const managerColIndex = headers.indexOf('مدير المشروعات');
  const commissionColIndex = headers.indexOf('نسبة العمولة');

  // تحديد العمود التالي المتاح (بعد عمود الفاتورة Q=17)
  let nextCol = 18; // R
  if (managerColIndex !== -1) {
    nextCol = managerColIndex + 1;
  }

  // إضافة عمود مدير المشروعات إذا لم يكن موجوداً
  if (managerColIndex === -1) {
    projectsSheet.getRange(1, nextCol)
      .setValue('مدير المشروعات')
      .setBackground('#4a86e8')
      .setFontColor('white')
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
    projectsSheet.setColumnWidth(nextCol, 150);

    // إنشاء قائمة منسدلة من أسماء الأطراف
    const partiesData = partiesSheet.getRange('A2:A200').getValues();
    const partyNames = partiesData.filter(row => row[0] !== '').map(row => row[0]);

    if (partyNames.length > 0) {
      projectsSheet.getRange(2, nextCol, 200, 1).setDataValidation(
        SpreadsheetApp.newDataValidation()
          .requireValueInList(partyNames, true)
          .setAllowInvalid(false)
          .build()
      );
    }
    nextCol++;
  }

  // إضافة عمود نسبة العمولة إذا لم يكن موجوداً
  if (commissionColIndex === -1) {
    projectsSheet.getRange(1, nextCol)
      .setValue('نسبة العمولة')
      .setBackground('#4a86e8')
      .setFontColor('white')
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
    projectsSheet.setColumnWidth(nextCol, 100);

    // تنسيق كنسبة مئوية
    projectsSheet.getRange(2, nextCol, 200, 1).setNumberFormat('0%');
  }

  ui.alert('✅ تم بنجاح',
    'تم إضافة أعمدة:\n• مدير المشروعات (قائمة منسدلة من الأطراف)\n• نسبة العمولة',
    ui.ButtonSet.OK);
}

/**
 * عرض نافذة اختيار مدير المشروعات لتقرير العمولة
 * نسخة مبسطة بدون HTML لتجنب مشاكل الصلاحيات
 */
function showCommissionReportDialog() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const projectsSheet = ss.getSheetByName(CONFIG.SHEETS.PROJECTS);

  if (!projectsSheet) {
    ui.alert('⚠️ لم يتم العثور على قاعدة بيانات المشاريع');
    return;
  }

  // البحث عن عمود مدير المشروعات
  const headers = projectsSheet.getRange(1, 1, 1, 25).getValues()[0];
  const managerColIndex = headers.indexOf('مدير المشروعات');

  if (managerColIndex === -1) {
    ui.alert('⚠️ عمود "مدير المشروعات" غير موجود.\n\nاستخدم أولاً: إضافة أعمدة العمولات للمشاريع');
    return;
  }

  // جمع أسماء مديري المشروعات الفريدة
  const lastRow = projectsSheet.getLastRow();
  if (lastRow < 2) {
    ui.alert('⚠️ لا توجد مشاريع في قاعدة البيانات');
    return;
  }

  const managersData = projectsSheet.getRange(2, managerColIndex + 1, lastRow - 1, 1).getValues();
  const uniqueManagers = [...new Set(managersData.filter(row => row[0] !== '').map(row => row[0]))];

  if (uniqueManagers.length === 0) {
    ui.alert('⚠️ لا يوجد مديرو مشروعات معينين.\n\nقم بتعيين مدير المشروعات في قاعدة بيانات المشاريع أولاً.');
    return;
  }

  // عرض قائمة المديرين للاختيار
  let managersList = 'مديرو المشروعات المتاحون:\n\n';
  uniqueManagers.forEach((m, i) => {
    managersList += (i + 1) + '. ' + m + '\n';
  });
  managersList += '\nأدخل رقم مدير المشروعات:';

  const response = ui.prompt('📊 تقرير العمولات', managersList, ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const selectedIndex = parseInt(response.getResponseText()) - 1;
  if (isNaN(selectedIndex) || selectedIndex < 0 || selectedIndex >= uniqueManagers.length) {
    ui.alert('⚠️ رقم غير صحيح. اختر رقماً من 1 إلى ' + uniqueManagers.length);
    return;
  }

  const selectedManager = uniqueManagers[selectedIndex];

  // تأكيد الاختيار
  const confirmResponse = ui.alert(
    '📊 تأكيد',
    'سيتم إنشاء تقرير عمولات لـ:\n\n' + selectedManager + '\n\nمتابعة؟',
    ui.ButtonSet.YES_NO
  );

  if (confirmResponse !== ui.Button.YES) {
    return;
  }

  // إنشاء التقرير
  generateManagerCommissionReport(selectedManager, '', '');
}

/**
 * إنشاء تقرير العمولات لمدير مشروعات معين
 */
function generateManagerCommissionReport(managerName, fromDateStr, toDateStr) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const projectsSheet = ss.getSheetByName(CONFIG.SHEETS.PROJECTS);
  const transSheet = ss.getSheetByName(CONFIG.SHEETS.TRANSACTIONS);

  if (!projectsSheet || !transSheet) {
    ui.alert('⚠️ خطأ في الوصول للشيتات المطلوبة');
    return;
  }

  // البحث عن أعمدة المشاريع
  const projHeaders = projectsSheet.getRange(1, 1, 1, 25).getValues()[0];
  const projCodeCol = 0;           // A - كود المشروع
  const projNameCol = 1;           // B - اسم المشروع
  const contractValueCol = 8;      // I - قيمة العقد
  const invoiceCol = projHeaders.indexOf('رقم الفاتورة') !== -1 ? projHeaders.indexOf('رقم الفاتورة') : 16; // Q
  const managerCol = projHeaders.indexOf('مدير المشروعات');
  const commissionCol = projHeaders.indexOf('نسبة العمولة');

  if (managerCol === -1 || commissionCol === -1) {
    ui.alert('⚠️ أعمدة العمولات غير موجودة');
    return;
  }

  // قراءة بيانات المشاريع
  const projLastRow = projectsSheet.getLastRow();
  const projData = projectsSheet.getRange(2, 1, projLastRow - 1, 25).getValues();

  // تصفية مشاريع المدير المحدد التي لها فاتورة
  const managerProjects = [];
  for (let i = 0; i < projData.length; i++) {
    const projectManager = projData[i][managerCol];
    const invoiceNum = projData[i][invoiceCol];
    const projectCode = projData[i][projCodeCol];

    if (projectManager === managerName && invoiceNum && projectCode) {
      // معالجة نسبة العمولة - قد تكون "3%" أو 0.03 أو 3
      let rateValue = projData[i][commissionCol];
      let commissionRate = 0;
      if (typeof rateValue === 'string') {
        // إذا كانت نص مثل "3%"
        commissionRate = parseFloat(rateValue.replace('%', '').replace('٪', '')) / 100;
      } else if (typeof rateValue === 'number') {
        // إذا كانت رقم - نتحقق إذا > 1 (مثل 3 تعني 3%) أو < 1 (مثل 0.03)
        commissionRate = rateValue > 1 ? rateValue / 100 : rateValue;
      }
      if (isNaN(commissionRate)) commissionRate = 0;

      managerProjects.push({
        code: projectCode,
        name: projData[i][projNameCol],
        contractValue: Number(projData[i][contractValueCol]) || 0,
        commissionRate: commissionRate,
        invoiceNum: invoiceNum
      });
    }
  }

  if (managerProjects.length === 0) {
    ui.alert('ℹ️ لا توجد مشاريع',
      'لا توجد مشاريع لـ "' + managerName + '" بها فواتير.\n\n' +
      'تأكد من:\n• تعيين مدير المشروعات\n• قطع فاتورة للمشروع',
      ui.ButtonSet.OK);
    return;
  }

  // قراءة الحركات المالية
  const transLastRow = transSheet.getLastRow();
  const transData = transSheet.getRange(2, 1, transLastRow - 1, 14).getValues();

  // تحويل التواريخ
  const fromDate = fromDateStr ? new Date(fromDateStr) : null;
  const toDate = toDateStr ? new Date(toDateStr) : null;

  // تجميع المصروفات والتحصيلات لكل مشروع
  const projectExpenses = {};
  const projectCollections = {};

  for (const proj of managerProjects) {
    projectExpenses[proj.code] = [];
    projectCollections[proj.code] = { total: 0, payments: [] };
  }

  for (let i = 0; i < transData.length; i++) {
    const transDate = transData[i][1];   // B - التاريخ
    const transType = String(transData[i][2] || '');  // C - طبيعة الحركة
    const classification = String(transData[i][3] || ''); // D - تصنيف الحركة
    const projectCode = String(transData[i][4] || '');    // E - كود المشروع
    const itemName = String(transData[i][6] || '');       // G - البند
    const partyName = String(transData[i][8] || '');      // I - الطرف
    const amountUsd = Number(transData[i][12]) || 0;      // M - القيمة بالدولار

    // تصفية حسب التاريخ
    if (fromDate && transDate && new Date(transDate) < fromDate) continue;
    if (toDate && transDate && new Date(transDate) > toDate) continue;

    // التحقق من أن المشروع ضمن مشاريع المدير
    if (!projectExpenses.hasOwnProperty(projectCode)) continue;

    // تصنيف الحركة
    if (classification.indexOf('مصروف') !== -1 || transType.indexOf('مصروف') !== -1) {
      // مصروفات
      projectExpenses[projectCode].push({
        item: itemName || classification,
        party: partyName,
        amount: amountUsd,
        date: transDate
      });
    } else if (transType.indexOf('تحصيل') !== -1) {
      // تحصيلات - نحفظ التفاصيل مع التاريخ
      projectCollections[projectCode].total += amountUsd;
      projectCollections[projectCode].payments.push({
        amount: amountUsd,
        date: transDate,
        party: partyName
      });
    }
  }

  // إنشاء شيت التقرير
  const reportSheetName = 'تقرير عمولة - ' + managerName;
  let reportSheet = ss.getSheetByName(reportSheetName);
  if (reportSheet) {
    reportSheet.clear();
  } else {
    reportSheet = ss.insertSheet(reportSheetName);
  }

  let currentRow = 1;

  // العنوان الرئيسي
  reportSheet.getRange(currentRow, 1, 1, 6).merge();
  reportSheet.getRange(currentRow, 1)
    .setValue('📊 تقرير عمولات: ' + managerName)
    .setFontSize(16)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setBackground('#4a86e8')
    .setFontColor('white');
  currentRow++;

  // الفترة
  let periodText = 'الفترة: ';
  if (fromDate || toDate) {
    periodText += (fromDateStr || 'البداية') + ' إلى ' + (toDateStr || 'الآن');
  } else {
    periodText += 'كل الفترات';
  }
  reportSheet.getRange(currentRow, 1, 1, 6).merge();
  reportSheet.getRange(currentRow, 1).setValue(periodText).setHorizontalAlignment('center');
  currentRow += 2;

  // تفاصيل كل مشروع
  let grandTotalExpenses = 0;
  let grandTotalCommission = 0;
  const projectSummaries = [];

  for (const proj of managerProjects) {
    const expenses = projectExpenses[proj.code];
    const totalExpenses = expenses.reduce((sum, e) => sum + e.amount, 0);
    const commission = totalExpenses * proj.commissionRate;
    const collectionData = projectCollections[proj.code];
    const collectedTotal = collectionData.total;
    const collectionRatio = proj.contractValue > 0 ? collectedTotal / proj.contractValue : 0;
    const dueCommission = commission * Math.min(collectionRatio, 1);
    // تاريخ آخر تحصيل
    let lastCollectionDate = null;
    if (collectionData.payments.length > 0) {
      collectionData.payments.sort((a, b) => new Date(b.date) - new Date(a.date));
      lastCollectionDate = collectionData.payments[0].date;
    }

    grandTotalExpenses += totalExpenses;
    grandTotalCommission += dueCommission;

    projectSummaries.push({
      code: proj.code,
      name: proj.name,
      expenses: totalExpenses,
      rate: proj.commissionRate,
      commission: commission,
      collectedTotal: collectedTotal,
      collectionPayments: collectionData.payments,
      lastCollectionDate: lastCollectionDate,
      contractValue: proj.contractValue,
      dueCommission: dueCommission
    });

    // عنوان المشروع
    reportSheet.getRange(currentRow, 1, 1, 6).merge();
    reportSheet.getRange(currentRow, 1)
      .setValue('📁 المشروع: ' + proj.code + ' - ' + proj.name)
      .setFontWeight('bold')
      .setBackground('#e8f0fe')
      .setFontSize(12);
    currentRow++;

    reportSheet.getRange(currentRow, 1)
      .setValue('   نسبة العمولة: ' + (proj.commissionRate * 100).toFixed(0) + '%');
    currentRow++;

    // هيدر تفاصيل المصروفات
    if (expenses.length > 0) {
      reportSheet.getRange(currentRow, 1, 1, 3)
        .setValues([['البند', 'المورد/الطرف', 'المبلغ ($)']])
        .setFontWeight('bold')
        .setBackground('#f3f3f3');
      currentRow++;

      // تفاصيل المصروفات
      for (const exp of expenses) {
        reportSheet.getRange(currentRow, 1, 1, 3)
          .setValues([[exp.item, exp.party, exp.amount]]);
        currentRow++;
      }

      // إجمالي مصروفات المشروع
      reportSheet.getRange(currentRow, 1, 1, 2).merge();
      reportSheet.getRange(currentRow, 1)
        .setValue('إجمالي مصروفات المشروع:')
        .setFontWeight('bold');
      reportSheet.getRange(currentRow, 3)
        .setValue(totalExpenses)
        .setFontWeight('bold')
        .setNumberFormat('$#,##0.00');
      currentRow++;
    } else {
      reportSheet.getRange(currentRow, 1).setValue('   لا توجد مصروفات مسجلة');
      currentRow++;
    }

    // قسم التحصيلات مع التواريخ
    currentRow++;
    if (collectionData.payments.length > 0) {
      reportSheet.getRange(currentRow, 1)
        .setValue('   💵 التحصيلات:')
        .setFontWeight('bold')
        .setFontColor('#0b5394');
      currentRow++;

      reportSheet.getRange(currentRow, 1, 1, 3)
        .setValues([['التاريخ', 'العميل', 'المبلغ ($)']])
        .setFontWeight('bold')
        .setBackground('#cfe2f3');
      currentRow++;

      for (const payment of collectionData.payments) {
        const paymentDate = payment.date ? Utilities.formatDate(new Date(payment.date), 'GMT+3', 'yyyy-MM-dd') : '-';
        reportSheet.getRange(currentRow, 1, 1, 3)
          .setValues([[paymentDate, payment.party, payment.amount]]);
        reportSheet.getRange(currentRow, 3).setNumberFormat('$#,##0.00');
        currentRow++;
      }

      reportSheet.getRange(currentRow, 1, 1, 2).merge();
      reportSheet.getRange(currentRow, 1)
        .setValue('إجمالي التحصيلات:')
        .setFontWeight('bold');
      reportSheet.getRange(currentRow, 3)
        .setValue(collectedTotal)
        .setFontWeight('bold')
        .setNumberFormat('$#,##0.00')
        .setBackground('#d0e0e3');
      currentRow++;
    } else {
      reportSheet.getRange(currentRow, 1)
        .setValue('   ⏳ لم يتم تحصيل أي مبالغ بعد')
        .setFontColor('#cc0000');
      currentRow++;
    }

    // ملخص العمولة للمشروع
    currentRow++;
    reportSheet.getRange(currentRow, 1)
      .setValue('   📈 حساب العمولة:')
      .setFontWeight('bold')
      .setFontColor('#38761d');
    currentRow++;

    const commissionCalc = totalExpenses + ' × ' + (proj.commissionRate * 100).toFixed(0) + '% = $' + commission.toFixed(2);
    reportSheet.getRange(currentRow, 1).setValue('      العمولة الإجمالية: ' + commissionCalc);
    currentRow++;

    const collectionPercent = (collectionRatio * 100).toFixed(1);
    reportSheet.getRange(currentRow, 1).setValue('      نسبة التحصيل: ' + collectionPercent + '% من قيمة العقد');
    currentRow++;

    reportSheet.getRange(currentRow, 1)
      .setValue('      ✅ العمولة المستحقة: $' + dueCommission.toFixed(2))
      .setFontWeight('bold')
      .setFontColor('#38761d');
    if (lastCollectionDate) {
      const formattedDate = Utilities.formatDate(new Date(lastCollectionDate), 'GMT+3', 'yyyy-MM-dd');
      reportSheet.getRange(currentRow, 3)
        .setValue('(تاريخ آخر تحصيل: ' + formattedDate + ')')
        .setFontColor('#666666');
    }
    currentRow++;

    currentRow++; // سطر فارغ بين المشاريع
  }

  // قسم الملخص
  currentRow++;
  reportSheet.getRange(currentRow, 1, 1, 6).merge();
  reportSheet.getRange(currentRow, 1)
    .setValue('📋 الملخص')
    .setFontSize(14)
    .setFontWeight('bold')
    .setBackground('#fce8b2')
    .setHorizontalAlignment('center');
  currentRow++;

  // هيدر الملخص
  reportSheet.getRange(currentRow, 1, 1, 8)
    .setValues([['المشروع', 'المصروفات', 'النسبة', 'العمولة الإجمالية', 'التحصيل', 'العمولة المستحقة', 'الحالة', 'إدراج ☑️']])
    .setFontWeight('bold')
    .setBackground('#fff2cc')
    .setHorizontalAlignment('center');
  currentRow++;

  // حفظ صف البداية للملخص لإضافة Checkbox
  const summaryStartRow = currentRow;

  // ملخص كل مشروع
  for (const summary of projectSummaries) {
    reportSheet.getRange(currentRow, 1).setValue(summary.code + ' - ' + summary.name);
    reportSheet.getRange(currentRow, 2).setValue(summary.expenses).setNumberFormat('$#,##0.00');
    reportSheet.getRange(currentRow, 3).setValue((summary.rate * 100).toFixed(0) + '%');
    reportSheet.getRange(currentRow, 4).setValue(summary.commission).setNumberFormat('$#,##0.00');
    reportSheet.getRange(currentRow, 5).setValue(summary.collectedTotal).setNumberFormat('$#,##0.00');
    reportSheet.getRange(currentRow, 6).setValue(summary.dueCommission).setNumberFormat('$#,##0.00');

    // التحقق من وجود استحقاق سابق
    const existing = checkExistingCommissionAccrual(summary.code, managerName);
    if (existing.exists) {
      if (Math.abs(existing.amount - summary.commission) < 0.01) {
        reportSheet.getRange(currentRow, 7).setValue('موجود ✅').setFontColor('#006400');
      } else if (existing.amount < summary.commission) {
        reportSheet.getRange(currentRow, 7).setValue('جزئي ⚠️').setFontColor('#b45f06');
      } else {
        reportSheet.getRange(currentRow, 7).setValue('موجود (أعلى) ⚠️').setFontColor('#cc0000');
      }
    } else {
      reportSheet.getRange(currentRow, 7).setValue('لم يُدرج').setFontColor('#999999');
    }

    // إضافة Checkbox
    reportSheet.getRange(currentRow, 8).insertCheckboxes();

    // تلوين حسب حالة التحصيل
    if (summary.collectedTotal > 0) {
      reportSheet.getRange(currentRow, 6).setBackground('#b6d7a8');
    } else {
      reportSheet.getRange(currentRow, 6).setBackground('#f4cccc');
    }
    currentRow++;
  }

  currentRow++;

  // الإجماليات
  reportSheet.getRange(currentRow, 1, 1, 3).merge();
  reportSheet.getRange(currentRow, 1)
    .setValue('📊 إجمالي المصروفات على جميع المشاريع:')
    .setFontWeight('bold');
  reportSheet.getRange(currentRow, 4)
    .setValue(grandTotalExpenses)
    .setFontWeight('bold')
    .setNumberFormat('$#,##0.00')
    .setBackground('#d9ead3');
  currentRow++;

  reportSheet.getRange(currentRow, 1, 1, 3).merge();
  reportSheet.getRange(currentRow, 1)
    .setValue('💰 إجمالي العمولة المستحقة:')
    .setFontWeight('bold')
    .setFontColor('#006400');
  reportSheet.getRange(currentRow, 4)
    .setValue(grandTotalCommission)
    .setFontWeight('bold')
    .setNumberFormat('$#,##0.00')
    .setBackground('#b6d7a8')
    .setFontColor('#006400');
  currentRow++;

  // تعديل عرض الأعمدة
  reportSheet.setColumnWidth(1, 200);
  reportSheet.setColumnWidth(2, 120);
  reportSheet.setColumnWidth(3, 80);
  reportSheet.setColumnWidth(4, 130);
  reportSheet.setColumnWidth(5, 100);
  reportSheet.setColumnWidth(6, 130);
  reportSheet.setColumnWidth(7, 100);
  reportSheet.setColumnWidth(8, 80);

  reportSheet.setFrozenRows(3);

  // الانتقال للشيت
  ss.setActiveSheet(reportSheet);

  SpreadsheetApp.getUi().alert('✅ تم إنشاء التقرير',
    'تقرير عمولات: ' + managerName + '\n\n' +
    '📁 عدد المشاريع: ' + managerProjects.length + '\n' +
    '💵 إجمالي المصروفات: $' + grandTotalExpenses.toFixed(2) + '\n' +
    '💰 العمولة المستحقة: $' + grandTotalCommission.toFixed(2),
    ui.ButtonSet.OK);
}

// ==================== 🎉 نهاية الكود ====================

/**
 * التحقق من وجود استحقاق عمولة سابق لمشروع معين
 * @param {string} projectCode - كود المشروع
 * @param {string} managerName - اسم مدير المشروع
 * @returns {object} - {exists: boolean, amount: number, row: number}
 */
function checkExistingCommissionAccrual(projectCode, managerName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const transSheet = ss.getSheetByName(CONFIG.SHEETS.TRANSACTIONS);

  if (!transSheet) return { exists: false, amount: 0, row: -1 };

  const lastRow = transSheet.getLastRow();
  if (lastRow < 2) return { exists: false, amount: 0, row: -1 };

  const data = transSheet.getRange(2, 1, lastRow - 1, 14).getValues();

  let totalExistingCommission = 0;
  let lastRow_found = -1;

  for (let i = 0; i < data.length; i++) {
    const rowProjectCode = String(data[i][4] || '');  // E - كود المشروع
    const rowItem = String(data[i][6] || '');         // G - البند
    const rowParty = String(data[i][8] || '');        // I - الطرف
    const rowAmount = Number(data[i][12]) || 0;       // M - المبلغ بالدولار
    const rowMovementType = String(data[i][13] || ''); // N - نوع الحركة

    // التحقق من أنها عمولة مدير انتاج لنفس المشروع ونفس المدير
    if (rowProjectCode === projectCode &&
      rowItem.indexOf('عمولة مدير') !== -1 &&
      rowParty === managerName &&
      rowMovementType.indexOf('مدين') !== -1) {
      totalExistingCommission += rowAmount;
      lastRow_found = i + 2; // +2 لأن البيانات تبدأ من السطر 2
    }
  }

  return {
    exists: totalExistingCommission > 0,
    amount: totalExistingCommission,
    row: lastRow_found
  };
}

/**
 * إدراج سطر استحقاق عمولة في شيت الحركات
 * @param {string} projectCode - كود المشروع
 * @param {string} managerName - اسم مدير المشروع
 * @param {number} commissionAmount - قيمة العمولة
 * @returns {boolean} - نجاح أو فشل
 */
function insertCommissionAccrual(projectCode, managerName, commissionAmount) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const transSheet = ss.getSheetByName(CONFIG.SHEETS.TRANSACTIONS);

  if (!transSheet || commissionAmount <= 0) return false;

  try {
    // البحث عن آخر رقم في العمود A وآخر صف به بيانات فعلية
    const colA = transSheet.getRange('A:A').getValues();
    const colB = transSheet.getRange('B:B').getValues();
    let maxNum = 0;
    let lastDataRow = 1;

    for (let i = 0; i < colA.length; i++) {
      const num = parseInt(colA[i][0]);
      if (!isNaN(num) && num > maxNum) maxNum = num;
      // البحث عن آخر صف به بيانات في عمود B (التاريخ)
      if (colB[i][0] !== '' && colB[i][0] !== null) {
        lastDataRow = i + 1;
      }
    }
    const newNum = maxNum + 1;
    const newRow = lastDataRow + 1;

    // تاريخ اليوم
    const today = new Date();

    // إعداد البيانات في مصفوفة واحدة (25 عمود = A إلى Y)
    // A=1, B=2, C=3, D=4, E=5, F=6, G=7, H=8, I=9, J=10, K=11, L=12, M=13, N=14, O=15, P=16, Q=17, R=18, S=19, T=20, U=21, V=22, W=23, X=24, Y=25
    const rowData = [
      newNum,                    // A - الرقم التسلسلي
      today,                     // B - التاريخ
      '💰 استحقاق مصروف',        // C - طبيعة الحركة
      'مصروفات مباشرة',          // D - تصنيف الحركة
      projectCode,               // E - كود المشروع
      '',                        // F - اسم المشروع (فارغ)
      'عمولة مدير انتاج',        // G - البند
      '',                        // H - التفاصيل (فارغ)
      managerName,               // I - المورد/الجهة
      commissionAmount,          // J - المبلغ بالعملة الأصلية
      'USD',                     // K - العملة
      '',                        // L - سعر الصرف (فارغ)
      commissionAmount,          // M - المبلغ بالدولار
      'مدين استحقاق',            // N - نوع الحركة
      '',                        // O - فارغ
      '',                        // P - فارغ
      'نقدي',                    // Q - طريقة الدفع
      'بعد التسليم',             // R - شرط الدفع
      3,                         // S - عدد الأسابيع
      '',                        // T - فارغ
      '',                        // U - فارغ
      '',                        // V - فارغ
      '',                        // W - فارغ
      '',                        // X - فارغ
      '📄'                       // Y - كشف
    ];

    // إدراج البيانات دفعة واحدة
    transSheet.getRange(newRow, 1, 1, 25).setValues([rowData]);

    return true;
  } catch (e) {
    Logger.log('خطأ في إدراج استحقاق العمولة: ' + e.message);
    return false;
  }
}

/**
 * معالجة النقر على Checkbox في تقرير العمولات
 * يُستدعى من onEdit
 */
function handleCommissionCheckbox(sheet, row, col) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // التحقق من أن العمود هو عمود الإدراج (العمود 8)
  if (col !== 8) return;

  // قراءة بيانات السطر
  const rowData = sheet.getRange(row, 1, 1, 8).getValues()[0];
  const projectInfo = String(rowData[0] || '');  // العمود 1 - المشروع
  const commissionAmount = Number(rowData[3]) || 0;  // العمود 4 - العمولة الإجمالية

  // استخراج كود المشروع من النص "PRJ-001 - اسم المشروع"
  const projectCode = projectInfo.split(' - ')[0].trim();

  if (!projectCode || commissionAmount <= 0) {
    sheet.getRange(row, col).setValue(false);
    ss.toast('⚠️ بيانات المشروع غير صحيحة', 'خطأ', 5);
    return;
  }

  // قراءة اسم المدير من عنوان التقرير
  const reportTitle = sheet.getRange(1, 1).getValue();
  const managerName = String(reportTitle).replace('📊 تقرير عمولات: ', '').trim();

  if (!managerName) {
    sheet.getRange(row, col).setValue(false);
    ss.toast('⚠️ لم يتم العثور على اسم المدير', 'خطأ', 5);
    return;
  }

  // التحقق من وجود استحقاق سابق
  const existing = checkExistingCommissionAccrual(projectCode, managerName);

  if (existing.exists) {
    const diff = commissionAmount - existing.amount;

    if (Math.abs(diff) < 0.01) {
      // السيناريو 2: موجود بنفس القيمة
      sheet.getRange(row, col).setValue(false);
      sheet.getRange(row, 7).setValue('موجود ✅');
      ss.toast('استحقاق العمولة موجود بالفعل بقيمة $' + existing.amount.toFixed(2), 'ℹ️ تنبيه', 5);
      return;
    } else if (diff > 0) {
      // السيناريو 3: موجود بقيمة مختلفة - إضافة الفرق تلقائياً
      const success = insertCommissionAccrual(projectCode, managerName, diff);
      if (success) {
        sheet.getRange(row, 7).setValue('تم إضافة الفرق ✅');
        sheet.getRange(row, col).setValue(false);
        ss.toast('تم إدراج فرق العمولة: $' + diff.toFixed(2), '✅ تم', 5);
      } else {
        sheet.getRange(row, col).setValue(false);
        ss.toast('فشل في إدراج استحقاق العمولة', '❌ خطأ', 5);
      }
      return;
    } else {
      // العمولة الجديدة أقل من السابقة
      sheet.getRange(row, col).setValue(false);
      sheet.getRange(row, 7).setValue('موجود (أعلى) ⚠️');
      ss.toast('يوجد استحقاق سابق بقيمة أعلى: $' + existing.amount.toFixed(2), '⚠️ تنبيه', 5);
      return;
    }
  }

  // السيناريو 1: لا يوجد استحقاق سابق - إدراج مباشر
  const success = insertCommissionAccrual(projectCode, managerName, commissionAmount);

  if (success) {
    sheet.getRange(row, 7).setValue('تم ✅');
    sheet.getRange(row, col).setValue(false);
    ss.toast('تم إدراج استحقاق العمولة: $' + commissionAmount.toFixed(2) + ' للمشروع ' + projectCode, '✅ تم بنجاح', 5);
  } else {
    sheet.getRange(row, col).setValue(false);
    ss.toast('فشل في إدراج استحقاق العمولة', '❌ خطأ', 5);
  }
}

/**
 * إدراج استحقاق عمولة من تقرير العمولات (يُستدعى من القائمة)
 * يقرأ السطر المحدد في تقرير العمولات ويدرج الاستحقاق
 */
function insertCommissionFromReport() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const sheetName = sheet.getName();

  // التحقق من أننا في تقرير عمولات
  if (sheetName.indexOf('تقرير عمولة - ') !== 0) {
    ui.alert('⚠️ خطأ', 'يجب أن تكون في شيت تقرير العمولات\n\nاسم الشيت يجب أن يبدأ بـ "تقرير عمولة - "', ui.ButtonSet.OK);
    return;
  }

  // الحصول على السطر المحدد
  const selection = ss.getActiveRange();
  const row = selection.getRow();

  // التحقق من أن السطر ليس الهيدر
  if (row <= 4) {
    ui.alert('⚠️ خطأ', 'يرجى تحديد سطر مشروع في جدول الملخص', ui.ButtonSet.OK);
    return;
  }

  // قراءة بيانات السطر
  const rowData = sheet.getRange(row, 1, 1, 8).getValues()[0];
  const projectInfo = String(rowData[0] || '');
  const commissionAmount = Number(rowData[3]) || 0;

  // استخراج كود المشروع
  const projectCode = projectInfo.split(' - ')[0].trim();

  if (!projectCode) {
    ui.alert('⚠️ خطأ', 'لم يتم العثور على كود المشروع في السطر المحدد\n\nتأكد من تحديد سطر في جدول الملخص', ui.ButtonSet.OK);
    return;
  }

  if (commissionAmount <= 0) {
    ui.alert('⚠️ خطأ', 'قيمة العمولة صفر أو غير صحيحة', ui.ButtonSet.OK);
    return;
  }

  // قراءة اسم المدير من عنوان التقرير
  const reportTitle = sheet.getRange(1, 1).getValue();
  const managerName = String(reportTitle).replace('📊 تقرير عمولات: ', '').trim();

  if (!managerName) {
    ui.alert('⚠️ خطأ', 'لم يتم العثور على اسم المدير في عنوان التقرير', ui.ButtonSet.OK);
    return;
  }

  // التحقق من وجود استحقاق سابق
  const existing = checkExistingCommissionAccrual(projectCode, managerName);

  if (existing.exists) {
    const diff = commissionAmount - existing.amount;

    if (Math.abs(diff) < 0.01) {
      // موجود بنفس القيمة
      sheet.getRange(row, 7).setValue('موجود ✅');
      ui.alert('ℹ️ تنبيه',
        'استحقاق العمولة موجود بالفعل!\n\n' +
        '• المشروع: ' + projectCode + '\n' +
        '• القيمة: $' + existing.amount.toFixed(2),
        ui.ButtonSet.OK);
      return;
    } else if (diff > 0) {
      // موجود بقيمة أقل - سؤال لإضافة الفرق
      const response = ui.alert('⚠️ يوجد استحقاق سابق',
        'يوجد استحقاق عمولة سابق بقيمة مختلفة:\n\n' +
        '• المشروع: ' + projectCode + '\n' +
        '• العمولة السابقة: $' + existing.amount.toFixed(2) + '\n' +
        '• العمولة الجديدة: $' + commissionAmount.toFixed(2) + '\n' +
        '• الفرق: $' + diff.toFixed(2) + '\n\n' +
        'هل تريد إضافة الفرق؟',
        ui.ButtonSet.YES_NO);

      if (response === ui.Button.YES) {
        const success = insertCommissionAccrual(projectCode, managerName, diff);
        if (success) {
          sheet.getRange(row, 7).setValue('تم إضافة الفرق ✅');
          ui.alert('✅ تم', 'تم إدراج فرق العمولة: $' + diff.toFixed(2), ui.ButtonSet.OK);
        } else {
          ui.alert('❌ خطأ', 'فشل في إدراج استحقاق العمولة', ui.ButtonSet.OK);
        }
      }
      return;
    } else {
      // موجود بقيمة أعلى
      sheet.getRange(row, 7).setValue('موجود (أعلى) ⚠️');
      ui.alert('⚠️ تنبيه',
        'يوجد استحقاق سابق بقيمة أعلى!\n\n' +
        '• العمولة السابقة: $' + existing.amount.toFixed(2) + '\n' +
        '• العمولة الحالية: $' + commissionAmount.toFixed(2),
        ui.ButtonSet.OK);
      return;
    }
  }

  // تأكيد الإدراج
  const confirm = ui.alert('تأكيد إدراج العمولة',
    'سيتم إدراج استحقاق عمولة:\n\n' +
    '• المشروع: ' + projectCode + '\n' +
    '• المدير: ' + managerName + '\n' +
    '• العمولة: $' + commissionAmount.toFixed(2) + '\n\n' +
    'هل تريد المتابعة؟',
    ui.ButtonSet.YES_NO);

  if (confirm !== ui.Button.YES) return;

  // إدراج الاستحقاق
  const success = insertCommissionAccrual(projectCode, managerName, commissionAmount);

  if (success) {
    sheet.getRange(row, 7).setValue('تم ✅');
    ui.alert('✅ تم بنجاح',
      'تم إدراج استحقاق العمولة:\n\n' +
      '• المشروع: ' + projectCode + '\n' +
      '• المدير: ' + managerName + '\n' +
      '• العمولة: $' + commissionAmount.toFixed(2),
      ui.ButtonSet.OK);
  } else {
    ui.alert('❌ خطأ', 'فشل في إدراج استحقاق العمولة', ui.ButtonSet.OK);
  }
}

// ==================== 📋 دفتر الأستاذ المساعد ====================

/**
 * إنشاء دفتر الأستاذ المساعد
 * يعرض كشف حساب كامل لكل طرف مقسم إلى 3 أقسام:
 * 1. الموردين (مستحقات علينا)
 * 2. العملاء (مستحقات لنا - فواتير وتأمينات)
 * 3. الممولين (قروض وتمويل)
 */
function generateDetailedPayablesReport() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const transSheet = ss.getSheetByName(CONFIG.SHEETS.TRANSACTIONS);
  const partiesSheet = ss.getSheetByName(CONFIG.SHEETS.PARTIES);

  if (!transSheet) {
    ui.alert('⚠️ لم يتم العثور على دفتر الحركات المالية');
    return;
  }

  // قراءة قاعدة بيانات الأطراف للحصول على نوع كل طرف
  const partyTypes = {};
  if (partiesSheet) {
    const partiesData = partiesSheet.getDataRange().getValues();
    for (let i = 1; i < partiesData.length; i++) {
      const name = String(partiesData[i][0] || '').trim();
      const type = String(partiesData[i][1] || '').trim();
      if (name) {
        partyTypes[name] = type; // مورد / عميل / ممول
      }
    }
  }

  const lastRow = transSheet.getLastRow();
  if (lastRow < 2) {
    ui.alert('ℹ️ لا توجد بيانات للتقرير');
    return;
  }

  // قراءة البيانات
  const data = transSheet.getRange(2, 1, lastRow - 1, 22).getValues();

  // تجميع الحركات حسب الطرف
  const parties = {};

  for (let i = 0; i < data.length; i++) {
    const date = data[i][1];                           // B - التاريخ
    const natureType = String(data[i][2] || '');       // C - طبيعة الحركة
    const projectCode = String(data[i][4] || '');      // E - كود المشروع
    const item = String(data[i][6] || '');             // G - البند
    const details = String(data[i][7] || '');          // H - التفاصيل
    const partyName = String(data[i][8] || '').trim(); // I - الطرف
    const amountUsd = Number(data[i][12]) || 0;        // M - المبلغ بالدولار
    const movementType = String(data[i][13] || '');    // N - نوع الحركة

    if (!partyName || amountUsd <= 0) continue;

    // تحديد نوع الحركة
    const isDebitAccrual = movementType.indexOf('مدين استحقاق') !== -1 || movementType.indexOf('مدين') !== -1;
    const isCreditPayment = movementType.indexOf('دائن دفعة') !== -1 || movementType.indexOf('دائن') !== -1;

    // تضمين حركات التمويل
    const isFinancing = natureType.indexOf('تمويل') !== -1 || natureType.indexOf('سداد تمويل') !== -1;
    // تضمين حركات الإيرادات
    const isRevenue = natureType.indexOf('إيراد') !== -1 || natureType.indexOf('تحصيل') !== -1;

    if (!isDebitAccrual && !isCreditPayment && !isFinancing && !isRevenue) continue;

    if (!parties[partyName]) {
      parties[partyName] = {
        items: [],
        totalDebit: 0,
        totalCredit: 0,
        type: partyTypes[partyName] || 'غير محدد'
      };
    }

    // تحديد المدين والدائن بناءً على نوع الحركة وطبيعتها
    let debit = 0;
    let credit = 0;

    if (natureType.indexOf('🏦 تمويل') !== -1) {
      // تمويل دخول = نحن مدينون للممول
      credit = amountUsd; // الممول أعطانا فلوس
    } else if (natureType.indexOf('💳 سداد تمويل') !== -1) {
      // سداد تمويل = نسدد للممول
      debit = amountUsd; // نحن ندفع للممول
    } else if (natureType.indexOf('📈 استحقاق إيراد') !== -1) {
      // استحقاق إيراد = العميل يدين لنا
      debit = amountUsd;
    } else if (isDebitAccrual) {
      debit = amountUsd;
    } else if (isCreditPayment) {
      credit = amountUsd;
    }

    parties[partyName].items.push({
      date: date,
      item: item,
      details: details,
      debit: debit,
      credit: credit,
      projectCode: projectCode,
      rowNum: i + 2
    });

    parties[partyName].totalDebit += debit;
    parties[partyName].totalCredit += credit;
  }

  // فلترة الأطراف: فقط من لديهم رصيد متبقي != 0
  const allParties = Object.keys(parties).filter(name => {
    const balance = Math.abs(parties[name].totalDebit - parties[name].totalCredit);
    return balance > 0.01;
  });

  if (allParties.length === 0) {
    ui.alert('✅ ممتاز', 'لا توجد أرصدة مستحقة!', ui.ButtonSet.OK);
    return;
  }

  // تقسيم الأطراف حسب النوع
  const vendors = allParties.filter(name => parties[name].type === 'مورد');
  const clients = allParties.filter(name => parties[name].type === 'عميل');
  const funders = allParties.filter(name => parties[name].type === 'ممول');
  const others = allParties.filter(name => !['مورد', 'عميل', 'ممول'].includes(parties[name].type));

  // ترتيب أبجدي
  vendors.sort((a, b) => a.localeCompare(b, 'ar'));
  clients.sort((a, b) => a.localeCompare(b, 'ar'));
  funders.sort((a, b) => a.localeCompare(b, 'ar'));
  others.sort((a, b) => a.localeCompare(b, 'ar'));

  // إنشاء أو إعادة استخدام الشيت
  const reportSheetName = 'دفتر الأستاذ المساعد';
  let reportSheet = ss.getSheetByName(reportSheetName);

  if (reportSheet) {
    reportSheet.clear();
  } else {
    reportSheet = ss.insertSheet(reportSheetName);
  }

  // === بناء التقرير ===
  let currentRow = 1;
  const numCols = 8;

  // العنوان الرئيسي
  reportSheet.getRange(currentRow, 1, 1, numCols).merge();
  reportSheet.getRange(currentRow, 1)
    .setValue('📋 دفتر الأستاذ المساعد')
    .setFontSize(16)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setBackground('#4a86e8')
    .setFontColor('white');
  currentRow++;

  // تاريخ التقرير
  reportSheet.getRange(currentRow, 1, 1, numCols).merge();
  reportSheet.getRange(currentRow, 1)
    .setValue('تاريخ التقرير: ' + Utilities.formatDate(new Date(), 'Asia/Riyadh', 'yyyy-MM-dd HH:mm'))
    .setFontSize(10)
    .setHorizontalAlignment('center')
    .setBackground('#cfe2f3');
  currentRow += 2;

  // متغيرات الإحصائيات
  let totalVendorsBalance = 0;
  let totalClientsBalance = 0;
  let totalFundersBalance = 0;
  let totalTransactions = 0;

  // دالة لرسم قسم كامل
  function drawSection(sectionTitle, partyList, sectionColor, sectionIcon) {
    if (partyList.length === 0) return 0;

    let sectionTotal = 0;

    // عنوان القسم
    reportSheet.getRange(currentRow, 1, 1, numCols).merge();
    reportSheet.getRange(currentRow, 1)
      .setValue(sectionIcon + ' ' + sectionTitle + ' (' + partyList.length + ')')
      .setFontSize(14)
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setBackground(sectionColor)
      .setFontColor('white');
    currentRow += 2;

    for (const partyName of partyList) {
      const party = parties[partyName];
      const partyBalance = party.totalDebit - party.totalCredit;
      sectionTotal += partyBalance;

      // ترتيب الحركات حسب التاريخ
      party.items.sort((a, b) => new Date(a.date) - new Date(b.date));

      // صف اسم الطرف
      reportSheet.getRange(currentRow, 1, 1, numCols).merge();
      const balanceText = partyBalance >= 0 ? 'مدين بـ $' + partyBalance.toFixed(2) : 'دائن بـ $' + Math.abs(partyBalance).toFixed(2);
      reportSheet.getRange(currentRow, 1)
        .setValue('👤 ' + partyName + ' (' + balanceText + ')')
        .setFontWeight('bold')
        .setFontSize(11)
        .setBackground('#fff2cc')
        .setFontColor('#7f6000');
      currentRow++;

      // هيدر الجدول
      const headers = ['#', 'التاريخ', 'البند', 'التفاصيل', 'مدين ($)', 'دائن ($)', 'الرصيد ($)', 'المشروع'];
      reportSheet.getRange(currentRow, 1, 1, numCols).setValues([headers]);
      reportSheet.getRange(currentRow, 1, 1, numCols)
        .setBackground('#1e88e5')
        .setFontColor('white')
        .setFontWeight('bold')
        .setHorizontalAlignment('center');
      currentRow++;

      // تفاصيل الحركات
      let runningBalance = 0;
      let itemNum = 0;

      for (const trans of party.items) {
        itemNum++;
        totalTransactions++;
        runningBalance += trans.debit - trans.credit;

        const formattedDate = trans.date ? Utilities.formatDate(new Date(trans.date), 'Asia/Riyadh', 'yyyy-MM-dd') : '';

        const rowData = [
          itemNum,
          formattedDate,
          trans.item,
          trans.details,
          trans.debit > 0 ? trans.debit : '',
          trans.credit > 0 ? trans.credit : '',
          runningBalance,
          trans.projectCode
        ];

        reportSheet.getRange(currentRow, 1, 1, numCols).setValues([rowData]);

        if (itemNum % 2 === 0) {
          reportSheet.getRange(currentRow, 1, 1, numCols).setBackground('#f5f5f5');
        }

        if (trans.debit > 0) {
          reportSheet.getRange(currentRow, 5).setFontColor('#cc0000');
        }
        if (trans.credit > 0) {
          reportSheet.getRange(currentRow, 6).setFontColor('#006600');
        }

        currentRow++;
      }

      // صف إجمالي الطرف
      reportSheet.getRange(currentRow, 1, 1, 4).merge();
      reportSheet.getRange(currentRow, 1)
        .setValue('إجمالي ' + partyName)
        .setFontWeight('bold')
        .setBackground('#e8eaf6');
      reportSheet.getRange(currentRow, 5)
        .setValue(party.totalDebit)
        .setFontWeight('bold')
        .setBackground('#ffcdd2')
        .setNumberFormat('$#,##0.00');
      reportSheet.getRange(currentRow, 6)
        .setValue(party.totalCredit)
        .setFontWeight('bold')
        .setBackground('#c8e6c9')
        .setNumberFormat('$#,##0.00');
      reportSheet.getRange(currentRow, 7)
        .setValue(partyBalance)
        .setFontWeight('bold')
        .setBackground('#fff9c4')
        .setNumberFormat('$#,##0.00');
      reportSheet.getRange(currentRow, 8).setBackground('#e8eaf6');
      currentRow += 2;
    }

    // إجمالي القسم
    reportSheet.getRange(currentRow, 1, 1, 6).merge();
    reportSheet.getRange(currentRow, 1)
      .setValue('📊 إجمالي ' + sectionTitle + ':')
      .setFontWeight('bold')
      .setFontSize(11)
      .setBackground(sectionColor)
      .setFontColor('white');
    reportSheet.getRange(currentRow, 7, 1, 2).merge();
    reportSheet.getRange(currentRow, 7)
      .setValue(sectionTotal)
      .setNumberFormat('$#,##0.00')
      .setFontWeight('bold')
      .setFontSize(11)
      .setBackground(sectionColor)
      .setFontColor('white');
    currentRow += 3;

    return sectionTotal;
  }

  // رسم الأقسام الثلاثة
  totalVendorsBalance = drawSection('الموردين (مستحقات علينا)', vendors, '#00695c', '🏭');
  totalClientsBalance = drawSection('العملاء (مستحقات لنا)', clients, '#f9a825', '👥');
  totalFundersBalance = drawSection('الممولين', funders, '#6a1b9a', '🏦');

  // إضافة قسم "أخرى" إذا وجد
  if (others.length > 0) {
    drawSection('أخرى (غير مصنف)', others, '#757575', '📁');
  }

  // === الملخص الإجمالي ===
  reportSheet.getRange(currentRow, 1, 1, numCols).merge();
  reportSheet.getRange(currentRow, 1)
    .setValue('📊 الملخص الإجمالي')
    .setFontWeight('bold')
    .setFontSize(14)
    .setBackground('#4a86e8')
    .setFontColor('white')
    .setHorizontalAlignment('center');
  currentRow++;

  // مستحقات الموردين
  reportSheet.getRange(currentRow, 1, 1, 5).merge();
  reportSheet.getRange(currentRow, 1)
    .setValue('🏭 مستحقات الموردين (علينا):')
    .setBackground('#e0f2f1');
  reportSheet.getRange(currentRow, 6, 1, 3).merge();
  reportSheet.getRange(currentRow, 6)
    .setValue(totalVendorsBalance)
    .setNumberFormat('$#,##0.00')
    .setFontWeight('bold')
    .setBackground('#e0f2f1')
    .setFontColor('#00695c');
  currentRow++;

  // مستحقات العملاء
  reportSheet.getRange(currentRow, 1, 1, 5).merge();
  reportSheet.getRange(currentRow, 1)
    .setValue('👥 مستحقات العملاء (لنا):')
    .setBackground('#fff8e1');
  reportSheet.getRange(currentRow, 6, 1, 3).merge();
  reportSheet.getRange(currentRow, 6)
    .setValue(totalClientsBalance)
    .setNumberFormat('$#,##0.00')
    .setFontWeight('bold')
    .setBackground('#fff8e1')
    .setFontColor('#f9a825');
  currentRow++;

  // مستحقات الممولين
  reportSheet.getRange(currentRow, 1, 1, 5).merge();
  reportSheet.getRange(currentRow, 1)
    .setValue('🏦 مستحقات الممولين:')
    .setBackground('#f3e5f5');
  reportSheet.getRange(currentRow, 6, 1, 3).merge();
  reportSheet.getRange(currentRow, 6)
    .setValue(totalFundersBalance)
    .setNumberFormat('$#,##0.00')
    .setFontWeight('bold')
    .setBackground('#f3e5f5')
    .setFontColor('#6a1b9a');
  currentRow++;

  // صافي الموقف
  const netPosition = totalClientsBalance - totalVendorsBalance - totalFundersBalance;
  reportSheet.getRange(currentRow, 1, 1, 5).merge();
  reportSheet.getRange(currentRow, 1)
    .setValue('💰 صافي الموقف (لنا - علينا):')
    .setFontWeight('bold')
    .setFontSize(12)
    .setBackground('#fff9c4');
  reportSheet.getRange(currentRow, 6, 1, 3).merge();
  reportSheet.getRange(currentRow, 6)
    .setValue(netPosition)
    .setNumberFormat('$#,##0.00')
    .setFontWeight('bold')
    .setFontSize(12)
    .setBackground('#fff9c4')
    .setFontColor(netPosition >= 0 ? '#2e7d32' : '#c62828');

  // تنسيق الأعمدة
  reportSheet.setColumnWidth(1, 40);
  reportSheet.setColumnWidth(2, 90);
  reportSheet.setColumnWidth(3, 150);
  reportSheet.setColumnWidth(4, 180);
  reportSheet.setColumnWidth(5, 100);
  reportSheet.setColumnWidth(6, 100);
  reportSheet.setColumnWidth(7, 100);
  reportSheet.setColumnWidth(8, 100);

  reportSheet.setFrozenRows(3);
  ss.setActiveSheet(reportSheet);

  ui.alert('✅ تم إنشاء دفتر الأستاذ المساعد',
    'الملخص:\n\n' +
    '🏭 الموردين: ' + vendors.length + ' (علينا: $' + totalVendorsBalance.toFixed(2) + ')\n' +
    '👥 العملاء: ' + clients.length + ' (لنا: $' + totalClientsBalance.toFixed(2) + ')\n' +
    '🏦 الممولين: ' + funders.length + ' ($' + totalFundersBalance.toFixed(2) + ')\n\n' +
    '💰 صافي الموقف: $' + netPosition.toFixed(2),
    ui.ButtonSet.OK);
}

// ==================== إخفاء/إظهار الشيتات ====================

/**
 * إخفاء/إظهار شيتات التقارير
 */
function toggleReportsVisibility() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const reportSheets = [
    CONFIG.SHEETS.PROJECT_REPORT,
    CONFIG.SHEETS.VENDORS_REPORT,
    CONFIG.SHEETS.FUNDERS_REPORT,
    CONFIG.SHEETS.EXPENSES_REPORT,
    CONFIG.SHEETS.REVENUE_REPORT,
    CONFIG.SHEETS.CASHFLOW,
    CONFIG.SHEETS.VENDOR_STATEMENT,
    CONFIG.SHEETS.CLIENT_STATEMENT,
    CONFIG.SHEETS.FUNDER_STATEMENT
  ];

  const result = toggleSheetsVisibility_(ss, reportSheets, 'التقارير');
  ui.alert(result.title, result.message, ui.ButtonSet.OK);
}

/**
 * إخفاء/إظهار التقارير التشغيلية
 * يشمل: تقارير ربحية المشاريع، دفتر الأستاذ المساعد، تقرير الاستحقاقات، التنبيهات والاستحقاقات
 */
function toggleOperationalReportsVisibility() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const operationalSheets = [
    'تقارير ربحية المشاريع',      // تقارير ربحية المشاريع
    'دفتر الأستاذ المساعد',        // دفتر الأستاذ المساعد
    'تقرير الاستحقاقات',           // تقرير الاستحقاقات الشامل
    CONFIG.SHEETS.ALERTS           // التنبيهات والاستحقاقات
  ];

  const result = toggleSheetsVisibility_(ss, operationalSheets, 'التقارير التشغيلية');
  ui.alert(result.title, result.message, ui.ButtonSet.OK);
}

/**
 * إخفاء/إظهار شيتات حسابات البنوك والخزنة
 */
function toggleBankAccountsVisibility() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const bankSheets = [
    CONFIG.SHEETS.BANK_USD,
    CONFIG.SHEETS.BANK_TRY,
    CONFIG.SHEETS.CASH_USD,
    CONFIG.SHEETS.CASH_TRY,
    CONFIG.SHEETS.CARD_TRY
  ];

  const result = toggleSheetsVisibility_(ss, bankSheets, 'حسابات البنوك والخزنة');
  ui.alert(result.title, result.message, ui.ButtonSet.OK);
}

/**
 * إخفاء/إظهار شيتات قواعد البيانات
 */
function toggleDatabasesVisibility() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const dbSheets = [
    CONFIG.SHEETS.PROJECTS,
    CONFIG.SHEETS.PARTIES,
    CONFIG.SHEETS.ITEMS
  ];

  const result = toggleSheetsVisibility_(ss, dbSheets, 'قواعد البيانات');
  ui.alert(result.title, result.message, ui.ButtonSet.OK);
}

/**
 * إخفاء/إظهار شيتات الدفاتر والقوائم المحاسبية
 * يشمل: شجرة الحسابات، دفتر الأستاذ، ميزان المراجعة، قيود اليومية، قائمة الدخل، المركز المالي
 */
function toggleAccountingVisibility() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const accountingSheets = [
    CONFIG.SHEETS.CHART_OF_ACCOUNTS,   // شجرة الحسابات
    CONFIG.SHEETS.GENERAL_LEDGER,      // دفتر الأستاذ العام
    CONFIG.SHEETS.TRIAL_BALANCE,       // ميزان المراجعة
    CONFIG.SHEETS.JOURNAL_ENTRIES,     // قيود اليومية
    CONFIG.SHEETS.INCOME_STATEMENT,    // قائمة الدخل
    CONFIG.SHEETS.BALANCE_SHEET        // المركز المالي
  ];

  const result = toggleSheetsVisibility_(ss, accountingSheets, 'الدفاتر والقوائم المحاسبية');
  ui.alert(result.title, result.message, ui.ButtonSet.OK);
}

/**
 * دالة مساعدة لتبديل حالة إظهار/إخفاء مجموعة شيتات
 * @param {Spreadsheet} ss - الجدول
 * @param {string[]} sheetNames - أسماء الشيتات
 * @param {string} groupName - اسم المجموعة للعرض
 * @returns {Object} - عنوان ورسالة النتيجة
 */
function toggleSheetsVisibility_(ss, sheetNames, groupName) {
  let existingSheets = [];
  let hiddenCount = 0;
  let visibleCount = 0;

  // جمع الشيتات الموجودة وحساب حالتها
  for (const name of sheetNames) {
    const sheet = ss.getSheetByName(name);
    if (sheet) {
      existingSheets.push(sheet);
      if (sheet.isSheetHidden()) {
        hiddenCount++;
      } else {
        visibleCount++;
      }
    }
  }

  if (existingSheets.length === 0) {
    return {
      title: '⚠️ لا توجد شيتات',
      message: 'لم يتم العثور على أي شيت من مجموعة ' + groupName
    };
  }

  // تحديد الإجراء: إذا أغلبها مخفي → نظهر، وإلا → نخفي
  const shouldShow = hiddenCount >= visibleCount;

  let processedCount = 0;
  let skippedCount = 0;

  for (const sheet of existingSheets) {
    try {
      if (shouldShow) {
        sheet.showSheet();
      } else {
        sheet.hideSheet();
      }
      processedCount++;
    } catch (e) {
      // قد يفشل الإخفاء إذا كان الشيت الوحيد المرئي
      skippedCount++;
    }
  }

  const action = shouldShow ? 'إظهار' : 'إخفاء';
  const icon = shouldShow ? '👁️' : '🙈';

  let message = 'تم ' + action + ' ' + processedCount + ' شيت من ' + groupName;
  if (skippedCount > 0) {
    message += '\n⚠️ تم تخطي ' + skippedCount + ' شيت (لا يمكن إخفاء كل الشيتات)';
  }

  return {
    title: icon + ' تم ' + action + ' ' + groupName,
    message: message
  };
}

// ==================== فلتر الشيتات ====================

/**
 * تفعيل أو إلغاء الفلتر على الشيت الحالي
 * إذا كان هناك فلتر موجود، يتم إزالته
 * إذا لم يكن هناك فلتر، يتم إنشاء فلتر جديد على صف الهيدر
 */
function toggleFilter() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const ui = SpreadsheetApp.getUi();

  // التحقق من وجود فلتر حالي
  const existingFilter = sheet.getFilter();

  if (existingFilter) {
    // إزالة الفلتر الموجود
    existingFilter.remove();
    ui.alert('🔓 تم إزالة الفلتر', 'تم إلغاء الفلتر من شيت "' + sheet.getName() + '"', ui.ButtonSet.OK);
  } else {
    // إنشاء فلتر جديد
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();

    if (lastRow < 2 || lastCol < 1) {
      ui.alert('⚠️ تنبيه', 'الشيت فارغ أو لا يحتوي على بيانات كافية لإنشاء فلتر.', ui.ButtonSet.OK);
      return;
    }

    // إنشاء الفلتر من الصف الأول حتى آخر صف
    const range = sheet.getRange(1, 1, lastRow, lastCol);
    range.createFilter();

    ui.alert('🔍 تم تفعيل الفلتر',
      'تم إضافة فلتر على شيت "' + sheet.getName() + '"\n\n' +
      '• اضغط على السهم ▼ في أي عمود للفلترة\n' +
      '• يمكنك اختيار قيم محددة أو البحث\n' +
      '• لإزالة الفلتر، اختر هذا الأمر مرة أخرى',
      ui.ButtonSet.OK);
  }
}

/**
 * تفعيل الفلتر على دفتر الحركات المالية مباشرة
 */
function toggleTransactionsFilter() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.TRANSACTIONS);
  const ui = SpreadsheetApp.getUi();

  if (!sheet) {
    ui.alert('⚠️ خطأ', 'لم يتم العثور على شيت "دفتر الحركات المالية"', ui.ButtonSet.OK);
    return;
  }

  // الانتقال للشيت أولاً
  ss.setActiveSheet(sheet);

  const existingFilter = sheet.getFilter();

  if (existingFilter) {
    existingFilter.remove();
    ui.alert('🔓 تم إزالة الفلتر', 'تم إلغاء الفلتر من دفتر الحركات المالية', ui.ButtonSet.OK);
  } else {
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();

    if (lastRow < 2) {
      ui.alert('⚠️ تنبيه', 'لا توجد بيانات كافية لإنشاء فلتر.', ui.ButtonSet.OK);
      return;
    }

    const range = sheet.getRange(1, 1, lastRow, lastCol);
    range.createFilter();

    ui.alert('🔍 تم تفعيل الفلتر',
      'تم إضافة فلتر على دفتر الحركات المالية\n\n' +
      'يمكنك الآن الفلترة حسب:\n' +
      '• المشروع\n' +
      '• المورد/العميل\n' +
      '• نوع الحركة\n' +
      '• التاريخ\n' +
      '• أي عمود آخر',
      ui.ButtonSet.OK);
  }
}

// ==================== نموذج إضافة حركة (Transaction Form) ====================

/**
 * دالة اختبار لتشخيص مشكلة الأذونات
 */
function testFormPermissions() {
  const ui = SpreadsheetApp.getUi();
  let results = [];

  try {
    // اختبار 1: الوصول للسبريدشيت
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    results.push('✅ الوصول للسبريدشيت: نجح');
    results.push('   اسم الملف: ' + ss.getName());

    // اختبار 2: قراءة شيت المشاريع
    const projectsSheet = ss.getSheetByName(CONFIG.SHEETS.PROJECTS);
    if (projectsSheet) {
      const lastRow = projectsSheet.getLastRow();
      results.push('✅ شيت المشاريع: موجود (' + lastRow + ' صف)');
    } else {
      results.push('⚠️ شيت المشاريع: غير موجود');
    }

    // اختبار 3: قراءة شيت الأطراف
    const partiesSheet = ss.getSheetByName(CONFIG.SHEETS.PARTIES);
    if (partiesSheet) {
      const lastRow = partiesSheet.getLastRow();
      results.push('✅ شيت الأطراف: موجود (' + lastRow + ' صف)');
    } else {
      results.push('⚠️ شيت الأطراف: غير موجود');
    }

    // اختبار 4: قراءة شيت البنود
    const itemsSheet = ss.getSheetByName(CONFIG.SHEETS.ITEMS);
    if (itemsSheet) {
      const lastRow = itemsSheet.getLastRow();
      results.push('✅ شيت البنود: موجود (' + lastRow + ' صف)');
    } else {
      results.push('⚠️ شيت البنود: غير موجود');
    }

    // اختبار 5: قراءة CONFIG
    results.push('✅ CONFIG.NATURE_TYPES: ' + CONFIG.NATURE_TYPES.length + ' عنصر');
    results.push('✅ CONFIG.PAYMENT_METHODS: ' + CONFIG.PAYMENT_METHODS.length + ' عنصر');

    // اختبار 6: HtmlService
    try {
      const html = HtmlService.createHtmlOutput('<p>test</p>');
      results.push('✅ HtmlService: يعمل');
    } catch (e) {
      results.push('❌ HtmlService: ' + e.message);
    }

  } catch (e) {
    results.push('❌ خطأ عام: ' + e.message);
  }

  ui.alert('🔍 نتائج الاختبار', results.join('\n'), ui.ButtonSet.OK);
}

// ═══════════════════════════════════════════════════════════════════════════════
// دوال تعريف المستخدم (User Identification)
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * عرض نافذة تعريف المستخدم
 * تظهر عند فتح الشيت أو من القائمة
 */
function showUserIdentificationDialog() {
  try {
    // جلب قائمة المستخدمين النشطين
    const users = getActiveUsersForForm();

    // جلب المستخدم الحالي المحفوظ (إن وجد)
    let currentUser = null;
    try {
      const savedName = PropertiesService.getUserProperties().getProperty('currentUserName');
      const savedEmail = PropertiesService.getUserProperties().getProperty('currentUserEmail');
      if (savedName) {
        currentUser = { name: savedName, email: savedEmail || '' };
      }
    } catch (e) { /* تجاهل */ }

    // إعداد البيانات للقالب
    const usersData = {
      users: users,
      currentUser: currentUser
    };

    // إنشاء قالب HTML
    const template = HtmlService.createTemplateFromFile('UserIdentification');
    template.usersData = usersData;

    const html = template.evaluate()
      .setWidth(380)
      .setHeight(420);

    SpreadsheetApp.getUi().showModalDialog(html, '👤 تعريف المستخدم');
  } catch (e) {
    console.log('خطأ في عرض نافذة تعريف المستخدم:', e.message);
    // لا نعرض رسالة خطأ للمستخدم حتى لا نزعجه
  }
}

/**
 * حفظ هوية المستخدم الحالي
 * @param {string} userName - اسم المستخدم
 * @param {string} userEmail - إيميل المستخدم
 */
function saveCurrentUserIdentity(userName, userEmail) {
  try {
    const userProps = PropertiesService.getUserProperties();
    userProps.setProperty('currentUserName', userName || '');
    userProps.setProperty('currentUserEmail', userEmail || '');

    // تسجيل في الـ console للتأكد
    console.log('تم حفظ هوية المستخدم:', userName, userEmail);

    return { success: true };
  } catch (e) {
    console.log('خطأ في حفظ هوية المستخدم:', e.message);
    throw new Error('فشل في حفظ البيانات: ' + e.message);
  }
}

/**
 * جلب هوية المستخدم الحالي المحفوظة
 * @returns {Object} بيانات المستخدم {name, email} أو null
 */
function getCurrentUserIdentity() {
  try {
    const userProps = PropertiesService.getUserProperties();
    const name = userProps.getProperty('currentUserName');
    const email = userProps.getProperty('currentUserEmail');

    if (name) {
      return { name: name, email: email || '' };
    }
    return null;
  } catch (e) {
    console.log('خطأ في جلب هوية المستخدم:', e.message);
    return null;
  }
}

/**
 * مسح هوية المستخدم الحالي (تسجيل خروج)
 */
function clearCurrentUserIdentity() {
  try {
    const userProps = PropertiesService.getUserProperties();
    userProps.deleteProperty('currentUserName');
    userProps.deleteProperty('currentUserEmail');
    return { success: true };
  } catch (e) {
    console.log('خطأ في مسح هوية المستخدم:', e.message);
    return { success: false };
  }
}

/**
 * التحقق مما إذا كان المستخدم قد عرّف نفسه
 * @returns {boolean}
 */
function isUserIdentified() {
  const identity = getCurrentUserIdentity();
  return identity !== null && identity.name !== '';
}

/**
 * دالة تُستدعى من Installable Trigger عند فتح الشيت
 * تعرض نافذة تعريف المستخدم إذا لم يكن معرّفاً
 */
function onOpenInstallable() {
  try {
    // التحقق مما إذا كان المستخدم قد عرّف نفسه
    if (!isUserIdentified()) {
      showUserIdentificationDialog();
    }
  } catch (e) {
    console.log('خطأ في onOpenInstallable:', e.message);
  }
}

/**
 * تفعيل نافذة تعريف المستخدم التلقائية عند فتح الشيت
 * يجب تشغيل هذه الدالة مرة واحدة فقط
 */
function installUserIdentificationTrigger() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // حذف أي trigger موجود مسبقاً لنفس الدالة
  const triggers = ScriptApp.getUserTriggers(ss);
  let existingTrigger = false;

  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === 'onOpenInstallable') {
      existingTrigger = true;
      break;
    }
  }

  if (existingTrigger) {
    const response = ui.alert(
      '⚠️ تنبيه',
      'نافذة تعريف المستخدم التلقائية مُفعّلة بالفعل.\n\nهل تريد إعادة تفعيلها؟',
      ui.ButtonSet.YES_NO
    );

    if (response !== ui.Button.YES) {
      return;
    }

    // حذف الـ trigger الموجود
    for (const trigger of triggers) {
      if (trigger.getHandlerFunction() === 'onOpenInstallable') {
        ScriptApp.deleteTrigger(trigger);
      }
    }
  }

  // إنشاء trigger جديد
  ScriptApp.newTrigger('onOpenInstallable')
    .forSpreadsheet(ss)
    .onOpen()
    .create();

  ui.alert(
    '✅ تم التفعيل',
    'تم تفعيل نافذة تعريف المستخدم التلقائية.\n\n' +
    'الآن عند فتح أي مستخدم للشيت، ستظهر له نافذة لتعريف نفسه.',
    ui.ButtonSet.OK
  );
}

/**
 * إلغاء تفعيل نافذة تعريف المستخدم التلقائية
 */
function uninstallUserIdentificationTrigger() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const triggers = ScriptApp.getUserTriggers(ss);
  let found = false;

  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === 'onOpenInstallable') {
      ScriptApp.deleteTrigger(trigger);
      found = true;
    }
  }

  if (found) {
    ui.alert(
      '✅ تم الإلغاء',
      'تم إلغاء نافذة تعريف المستخدم التلقائية.',
      ui.ButtonSet.OK
    );
  } else {
    ui.alert(
      'ℹ️ معلومة',
      'نافذة تعريف المستخدم التلقائية غير مُفعّلة أصلاً.',
      ui.ButtonSet.OK
    );
  }
}

// ═══════════════════════════════════════════════════════════════════════════════

/**
 * عرض نموذج إضافة حركة جديدة
 * يمرر البيانات مباشرة للنموذج لتجنب مشاكل الأذونات
 */
function showTransactionForm() {
  try {
    // ═══════════════════════════════════════════════════════════════
    // حفظ إيميل المستخدم الحالي قبل فتح النموذج
    // (لأن google.script.run لا يمكنه الحصول على الإيميل)
    // ═══════════════════════════════════════════════════════════════
    try {
      const userEmail = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail();
      if (userEmail) {
        PropertiesService.getUserProperties().setProperty('currentUserEmail', userEmail);
      }
    } catch (emailErr) {
      console.log('تعذر حفظ إيميل المستخدم:', emailErr.message);
    }

    // جلب البيانات أولاً
    const formData = getSmartFormData();

    // إنشاء قالب HTML مع البيانات
    const template = HtmlService.createTemplateFromFile('TransactionForm');
    template.formData = formData;

    const html = template.evaluate()
      .setWidth(520)
      .setHeight(750);

    SpreadsheetApp.getUi().showModalDialog(html, '➕ إضافة حركة جديدة');
  } catch (e) {
    SpreadsheetApp.getUi().alert('❌ خطأ في فتح النموذج',
      'حدث خطأ: ' + e.message + '\n\n' +
      'الرجاء المحاولة مرة أخرى أو الاتصال بالدعم الفني.',
      SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * جلب بيانات القوائم المنسدلة للنموذج
 * @returns {Object} بيانات القوائم
 */
function getSmartFormData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // تاريخ اليوم
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy');

  // جلب المشاريع (كود + اسم)
  const projectsSheet = ss.getSheetByName(CONFIG.SHEETS.PROJECTS);
  const projects = [];
  if (projectsSheet && projectsSheet.getLastRow() > 1) {
    const projectData = projectsSheet.getRange(2, 1, projectsSheet.getLastRow() - 1, 2).getValues();
    projectData.forEach(row => {
      if (row[0]) {
        projects.push({
          code: String(row[0]).trim(),
          name: String(row[1] || '').trim(),
          display: `${row[1]} (${row[0]})`  // اسم المشروع (الكود)
        });
      }
    });
  }

  // جلب الأطراف (موردين/عملاء/ممولين)
  const partiesSheet = ss.getSheetByName(CONFIG.SHEETS.PARTIES);
  const parties = [];
  if (partiesSheet && partiesSheet.getLastRow() > 1) {
    const partyData = partiesSheet.getRange(2, 1, partiesSheet.getLastRow() - 1, 2).getValues();
    partyData.forEach(row => {
      if (row[0]) {
        parties.push({
          name: String(row[0]).trim(),
          type: String(row[1] || '').trim()
        });
      }
    });
  }

  // جلب البنود والتصنيفات
  const itemsSheet = ss.getSheetByName(CONFIG.SHEETS.ITEMS);
  const items = [];
  const classifications = [];
  if (itemsSheet && itemsSheet.getLastRow() > 1) {
    const itemData = itemsSheet.getRange(2, 1, itemsSheet.getLastRow() - 1, 3).getValues();
    const classSet = new Set();
    itemData.forEach(row => {
      if (row[0]) {
        items.push(String(row[0]).trim());
      }
      if (row[2] && !classSet.has(row[2])) {
        classSet.add(row[2]);
        classifications.push(String(row[2]).trim());
      }
    });
  }

  return {
    today: today,
    projects: projects,
    parties: parties,
    items: items,
    classifications: classifications,
    natureTypes: CONFIG.NATURE_TYPES,
    movementTypes: CONFIG.MOVEMENT.TYPES,
    currencies: CONFIG.CURRENCIES.LIST.slice(0, 3),  // USD, TRY, EGP
    paymentMethods: CONFIG.PAYMENT_METHODS,
    paymentTerms: CONFIG.PAYMENT_TERMS.LIST,
    users: getActiveUsersForForm()  // قائمة المستخدمين النشطين
  };
}

/**
 * جلب قائمة المستخدمين النشطين للنموذج
 * ⭐ محدّث للهيكل الجديد مع Checkboxes
 * @returns {Array} قائمة المستخدمين
 */
function getActiveUsersForForm() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const usersSheet = ss.getSheetByName(CONFIG.SHEETS.BOT_USERS);

    if (!usersSheet || usersSheet.getLastRow() <= 1) {
      return [];
    }

    const columns = BOT_CONFIG.BOT_USERS_COLUMNS;
    const data = usersSheet.getDataRange().getValues();
    const users = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const name = row[columns.NAME.index - 1];
      const email = row[columns.EMAIL.index - 1];

      // ⭐ الهيكل الجديد: Checkboxes بدلاً من نص
      const permSheet = row[columns.PERM_SHEET.index - 1];  // صلاحية الشيت (checkbox)
      const isActive = row[columns.IS_ACTIVE.index - 1];    // نشط (checkbox)

      // فقط المستخدمين النشطين الذين لديهم صلاحية الشيت
      if (name && isActive === true && permSheet === true) {
        users.push({
          name: name,
          email: email || '',
          display: email ? `${name} (${email})` : name
        });
      }
    }

    return users;
  } catch (e) {
    console.log('خطأ في جلب المستخدمين:', e.message);
    return [];
  }
}

/**
 * تخزين بيانات النموذج مؤقتاً باستخدام ScriptProperties
 * @param {string} jsonData بيانات النموذج بصيغة JSON
 */
function cacheTempFormData(jsonData) {
  const props = PropertiesService.getScriptProperties();
  props.setProperty('pendingTransaction', jsonData);
  props.setProperty('pendingTransactionTime', new Date().toISOString());
  return { success: true };
}

/**
 * معالجة البيانات المخزنة وحفظ الحركة
 */
function processPendingTransaction() {
  const props = PropertiesService.getScriptProperties();
  const jsonData = props.getProperty('pendingTransaction');

  if (!jsonData) {
    SpreadsheetApp.getUi().alert('⚠️ لا توجد بيانات معلقة للحفظ');
    return;
  }

  const formData = JSON.parse(jsonData);
  const result = saveTransactionData(formData);

  // حذف البيانات المؤقتة
  props.deleteProperty('pendingTransaction');
  props.deleteProperty('pendingTransactionTime');

  SpreadsheetApp.getUi().alert(
    '✅ تمت إضافة الحركة بنجاح!',
    'رقم الحركة: ' + result.transNum + '\n' + result.summary,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * إدخال حركة يدوياً عبر JSON
 * حل احتياطي في حال فشل النموذج
 */
function manualTransactionEntry() {
  const ui = SpreadsheetApp.getUi();

  const response = ui.prompt(
    '📝 إدخال حركة يدوياً',
    'الصق بيانات الحركة (JSON) هنا:',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const jsonData = response.getResponseText().trim();
  if (!jsonData) {
    ui.alert('⚠️ لم يتم إدخال بيانات');
    return;
  }

  try {
    const formData = JSON.parse(jsonData);
    const result = saveTransactionData(formData);

    ui.alert(
      '✅ تمت إضافة الحركة بنجاح!',
      'رقم الحركة: ' + result.transNum + '\n' + result.summary,
      ui.ButtonSet.OK
    );
  } catch (e) {
    ui.alert('❌ خطأ في تنسيق البيانات: ' + e.message);
  }
}

/**
 * إدخال حركة سريعة عبر نوافذ متتالية (بديل للنموذج HTML)
 * يعمل بدون مشاكل الصلاحيات
 */
function quickTransactionEntry() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // حفظ إيميل المستخدم للتسجيل
  try {
    const userEmail = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail();
    if (userEmail) {
      PropertiesService.getUserProperties().setProperty('currentUserEmail', userEmail);
    }
  } catch (e) { /* تجاهل */ }

  // جلب البيانات للقوائم
  const projectsSheet = ss.getSheetByName(CONFIG.SHEETS.PROJECTS);
  const partiesSheet = ss.getSheetByName(CONFIG.SHEETS.PARTIES);

  // جمع بيانات المشاريع
  let projectsList = '';
  if (projectsSheet && projectsSheet.getLastRow() > 1) {
    const projects = projectsSheet.getRange(2, 1, projectsSheet.getLastRow() - 1, 2).getValues();
    projectsList = projects.filter(r => r[0]).map(r => `${r[0]} - ${r[1]}`).join('\n');
  }

  // 1. اختيار المشروع
  const projectResponse = ui.prompt(
    '📁 الخطوة 1/7: المشروع',
    'أدخل كود المشروع:\n\nالمشاريع المتاحة:\n' + projectsList.substring(0, 500),
    ui.ButtonSet.OK_CANCEL
  );
  if (projectResponse.getSelectedButton() !== ui.Button.OK) return;
  const projectCode = projectResponse.getResponseText().trim();

  // 2. طبيعة الحركة
  const natureResponse = ui.prompt(
    '📋 الخطوة 2/7: طبيعة الحركة',
    'اختر رقم طبيعة الحركة:\n\n' +
    '1. استحقاق مصروف\n' +
    '2. دفعة مصروف\n' +
    '3. استحقاق إيراد\n' +
    '4. تحصيل إيراد\n' +
    '5. تمويل\n' +
    '6. تحويل داخلي',
    ui.ButtonSet.OK_CANCEL
  );
  if (natureResponse.getSelectedButton() !== ui.Button.OK) return;
  const natureTypes = ['استحقاق مصروف', 'دفعة مصروف', 'استحقاق إيراد', 'تحصيل إيراد', 'تمويل', 'تحويل داخلي'];
  const natureType = natureTypes[parseInt(natureResponse.getResponseText()) - 1] || 'استحقاق مصروف';

  // 3. البند والتصنيف
  const itemResponse = ui.prompt(
    '📄 الخطوة 3/7: البند',
    'أدخل اسم البند (مثال: مونتاج، تصوير، إيجار):',
    ui.ButtonSet.OK_CANCEL
  );
  if (itemResponse.getSelectedButton() !== ui.Button.OK) return;
  const item = itemResponse.getResponseText().trim();

  // 4. المورد/الجهة
  const partyResponse = ui.prompt(
    '👤 الخطوة 4/7: المورد/الجهة',
    'أدخل اسم المورد أو الجهة (أو اتركه فارغاً):',
    ui.ButtonSet.OK_CANCEL
  );
  if (partyResponse.getSelectedButton() !== ui.Button.OK) return;
  const partyName = partyResponse.getResponseText().trim();

  // 5. المبلغ والعملة
  const amountResponse = ui.prompt(
    '💰 الخطوة 5/7: المبلغ',
    'أدخل المبلغ والعملة (مثال: 1000 USD أو 5000 TRY):',
    ui.ButtonSet.OK_CANCEL
  );
  if (amountResponse.getSelectedButton() !== ui.Button.OK) return;
  const amountParts = amountResponse.getResponseText().trim().split(/\s+/);
  const amount = parseFloat(amountParts[0]) || 0;
  const currency = (amountParts[1] || 'USD').toUpperCase();

  // 6. سعر الصرف (إذا لزم)
  let exchangeRate = 1;
  if (currency !== 'USD') {
    const rateResponse = ui.prompt(
      '💱 الخطوة 6/7: سعر الصرف',
      `أدخل سعر صرف ${currency} مقابل الدولار:`,
      ui.ButtonSet.OK_CANCEL
    );
    if (rateResponse.getSelectedButton() !== ui.Button.OK) return;
    exchangeRate = parseFloat(rateResponse.getResponseText()) || 1;
  }

  // 7. التفاصيل
  const detailsResponse = ui.prompt(
    '📝 الخطوة 7/7: التفاصيل',
    'أدخل تفاصيل الحركة (اختياري):',
    ui.ButtonSet.OK_CANCEL
  );
  if (detailsResponse.getSelectedButton() !== ui.Button.OK) return;
  const details = detailsResponse.getResponseText().trim();

  // تجميع البيانات
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy');
  const formData = {
    date: today,
    natureType: natureType,
    classification: natureType.includes('مصروف') ? 'مصروفات مباشرة' : 'إيرادات',
    projectCode: projectCode,
    item: item,
    partyName: partyName,
    details: details,
    amount: amount.toString(),
    currency: currency,
    exchangeRate: exchangeRate.toString(),
    paymentMethod: 'تحويل بنكي',
    refNumber: '',
    paymentTerm: '',
    weeksCount: '',
    customDueDate: '',
    notes: ''
  };

  // حفظ الحركة
  try {
    const result = saveTransactionData(formData);
    ui.alert(
      '✅ تمت إضافة الحركة بنجاح!',
      `رقم الحركة: ${result.transNum}\n${result.summary}`,
      ui.ButtonSet.OK
    );
  } catch (e) {
    ui.alert('❌ خطأ', e.message, ui.ButtonSet.OK);
  }
}

/**
 * حفظ الحركة من النموذج
 * @param {Object} formData بيانات النموذج
 * @returns {Object} نتيجة الحفظ
 */
function saveTransactionData(formData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.TRANSACTIONS);

  if (!sheet) {
    throw new Error('شيت دفتر الحركات المالية غير موجود');
  }

  // حساب رقم الحركة الجديد
  const lastRow = sheet.getLastRow();
  const newRow = lastRow + 1;
  const newTransNum = lastRow > 1 ?
    (Number(sheet.getRange(lastRow, 1).getValue()) || 0) + 1 : 1;

  // تحويل التاريخ
  const dateParts = formData.date.split('/');
  const transDate = new Date(dateParts[2], dateParts[1] - 1, dateParts[0]);

  // البيانات الأساسية
  const amount = Number(formData.amount) || 0;
  const exchangeRate = Number(formData.exchangeRate) || 1;

  // تحديد نوع الحركة (N)
  let movementType = '';
  if (formData.natureType.includes('استحقاق')) {
    movementType = 'مدين استحقاق';
  } else if (formData.natureType.includes('دفعة') || formData.natureType.includes('تحصيل') ||
    formData.natureType.includes('سداد') || formData.natureType.includes('استرداد')) {
    movementType = 'دائن دفعة';
  }

  // جلب اسم المشروع من الكود (F)
  let projectName = '';
  let projectDeliveryDate = null;
  const projectsSheet = ss.getSheetByName(CONFIG.SHEETS.PROJECTS);
  if (formData.projectCode && projectsSheet && projectsSheet.getLastRow() > 1) {
    const projectData = projectsSheet.getRange(2, 1, projectsSheet.getLastRow() - 1, 11).getValues();
    const found = projectData.find(row => String(row[0]).trim() === formData.projectCode);
    if (found) {
      projectName = found[1];
      // العمود K (index 10) = تاريخ التسليم المتوقع
      if (found[10] instanceof Date) {
        projectDeliveryDate = found[10];
      }
    }
  }

  // ═══════════════════════════════════════════════════════════════
  // حساب القيم المحسوبة (بدلاً من الاعتماد على معادلات الشيت)
  // ═══════════════════════════════════════════════════════════════

  // M: القيمة بالدولار
  let amountUsd = 0;
  if (amount > 0) {
    if (formData.currency === 'USD' || formData.currency === 'دولار' || !formData.currency) {
      amountUsd = amount;
    } else if (exchangeRate > 0) {
      amountUsd = amount / exchangeRate;
    }
  }

  // O: الرصيد (مجموع استحقاقات - مجموع دفعات لنفس الطرف)
  let balance = 0;
  if (formData.partyName && amountUsd > 0) {
    // جلب كل الحركات السابقة لنفس الطرف
    if (lastRow > 1) {
      const allData = sheet.getRange(2, 9, lastRow - 1, 6).getValues(); // I, J, K, L, M, N (columns 9-14)
      for (let i = 0; i < allData.length; i++) {
        const partyInRow = String(allData[i][0]).trim(); // I
        const amountUsdInRow = Number(allData[i][4]) || 0; // M
        const movementTypeInRow = String(allData[i][5]).trim(); // N

        if (partyInRow === formData.partyName) {
          if (movementTypeInRow === 'مدين استحقاق') {
            balance += amountUsdInRow;
          } else if (movementTypeInRow === 'دائن دفعة') {
            balance -= amountUsdInRow;
          }
        }
      }
    }
    // إضافة الحركة الحالية للرصيد
    if (movementType === 'مدين استحقاق') {
      balance += amountUsd;
    } else if (movementType === 'دائن دفعة') {
      balance -= amountUsd;
    }
  }

  // U: تاريخ الاستحقاق
  let dueDate = '';
  if (movementType === 'مدين استحقاق' && formData.paymentTerm) {
    if (formData.paymentTerm === 'فوري') {
      dueDate = transDate;
    } else if (formData.paymentTerm === 'بعد التسليم') {
      // تاريخ التسليم + (عدد الأسابيع × 7 أيام)
      if (projectDeliveryDate) {
        const weeksToAdd = Number(formData.weeksCount) || 0;
        dueDate = new Date(projectDeliveryDate);
        dueDate.setDate(dueDate.getDate() + (weeksToAdd * 7));
      }
    } else if (formData.paymentTerm === 'تاريخ مخصص' && formData.customDueDate) {
      const dueParts = formData.customDueDate.split('/');
      dueDate = new Date(dueParts[2], dueParts[1] - 1, dueParts[0]);
    }
  }

  // V: حالة السداد
  let paymentStatus = '';
  if (movementType === 'مدين استحقاق') {
    paymentStatus = balance <= 0 ? 'مدفوع بالكامل' : 'معلق';
  } else if (movementType === 'دائن دفعة') {
    paymentStatus = 'عملية دفع/تحصيل';
  }

  // W: الشهر (YYYY-MM)
  const monthStr = Utilities.formatDate(transDate, Session.getScriptTimeZone(), 'yyyy-MM');

  // ═══════════════════════════════════════════════════════════════
  // كتابة جميع القيم للصف الجديد
  // ═══════════════════════════════════════════════════════════════

  // A: رقم الحركة
  sheet.getRange(newRow, 1).setValue(newTransNum);

  // B: التاريخ
  sheet.getRange(newRow, 2).setValue(transDate).setNumberFormat('dd/mm/yyyy');

  // C: طبيعة الحركة
  sheet.getRange(newRow, 3).setValue(formData.natureType);

  // D: تصنيف الحركة
  sheet.getRange(newRow, 4).setValue(formData.classification);

  // E: كود المشروع
  sheet.getRange(newRow, 5).setValue(formData.projectCode);

  // F: اسم المشروع
  sheet.getRange(newRow, 6).setValue(projectName);

  // G: البند
  sheet.getRange(newRow, 7).setValue(formData.item);

  // H: التفاصيل
  sheet.getRange(newRow, 8).setValue(formData.details || '');

  // I: اسم المورد/الجهة
  sheet.getRange(newRow, 9).setValue(formData.partyName || '');

  // J: المبلغ بالعملة الأصلية
  sheet.getRange(newRow, 10).setValue(amount).setNumberFormat('#,##0.00');

  // K: العملة
  sheet.getRange(newRow, 11).setValue(formData.currency);

  // L: سعر الصرف
  sheet.getRange(newRow, 12).setValue(exchangeRate).setNumberFormat('#,##0.0000');

  // M: القيمة بالدولار (محسوبة)
  sheet.getRange(newRow, 13).setValue(amountUsd).setNumberFormat('#,##0.00');

  // N: نوع الحركة
  sheet.getRange(newRow, 14).setValue(movementType);

  // O: الرصيد (محسوب)
  if (formData.partyName) {
    sheet.getRange(newRow, 15).setValue(balance).setNumberFormat('#,##0.00');
  }

  // P: رقم مرجعي
  sheet.getRange(newRow, 16).setValue(formData.refNumber || '');

  // Q: طريقة الدفع
  sheet.getRange(newRow, 17).setValue(formData.paymentMethod || '');

  // R: نوع شرط الدفع
  sheet.getRange(newRow, 18).setValue(formData.paymentTerm || '');

  // S: عدد الأسابيع
  sheet.getRange(newRow, 19).setValue(formData.weeksCount || '');

  // T: تاريخ مخصص
  if (formData.customDueDate) {
    const customParts = formData.customDueDate.split('/');
    const customDate = new Date(customParts[2], customParts[1] - 1, customParts[0]);
    sheet.getRange(newRow, 20).setValue(customDate).setNumberFormat('dd/mm/yyyy');
  }

  // U: تاريخ الاستحقاق (محسوب)
  if (dueDate) {
    sheet.getRange(newRow, 21).setValue(dueDate).setNumberFormat('dd/mm/yyyy');
  }

  // V: حالة السداد (محسوبة)
  sheet.getRange(newRow, 22).setValue(paymentStatus);

  // W: الشهر (محسوب)
  sheet.getRange(newRow, 23).setValue(monthStr);

  // X: ملاحظات
  sheet.getRange(newRow, 24).setValue(formData.notes || '');

  // Y: كشف (رابط) - نتركه فارغاً

  // Z: رقم الأوردر - نتركه فارغاً (يُستخدم فقط في الأوردر المشترك)

  // AA: عدد الوحدات (اختياري - لحساب تكلفة الوحدة)
  if (formData.unitCount && Number(formData.unitCount) > 0) {
    sheet.getRange(newRow, 27).setValue(Number(formData.unitCount));
  }

  // ═══════════════════════════════════════════════════════════════
  // تأكيد الكتابة
  // ═══════════════════════════════════════════════════════════════
  SpreadsheetApp.flush();

  // ═══════════════════════════════════════════════════════════════
  // إعادة حساب الرصيد (O) وحالة السداد (V) لكل حركات نفس الطرف
  // هذا يضمن تحديث الحركات السابقة عند إضافة حركة جديدة
  // ═══════════════════════════════════════════════════════════════
  if (formData.partyName) {
    recalculatePartyBalance_(sheet, newRow);
    SpreadsheetApp.flush();
  }

  // ═══════════════════════════════════════════════════════════════
  // تسجيل النشاط
  // ═══════════════════════════════════════════════════════════════
  const summaryText = `${formData.natureType} - ${formData.partyName || formData.item} - ${amount} ${formData.currency}`;

  // استخدام إيميل المستخدم المحدد من النموذج أو اسمه
  const userIdentifier = formData.submittedByEmail || formData.submittedBy || '';

  logActivity(
    'إضافة حركة',
    CONFIG.SHEETS.TRANSACTIONS,
    newRow,
    newTransNum,
    summaryText,
    {
      projectCode: formData.projectCode,
      projectName: projectName,
      item: formData.item,
      partyName: formData.partyName,
      amount: amount,
      currency: formData.currency,
      amountUsd: amountUsd,
      movementType: movementType,
      submittedBy: formData.submittedBy  // إضافة اسم المستخدم للتفاصيل
    },
    userIdentifier  // تمرير الإيميل أو الاسم كمعرف للمستخدم
  );

  return {
    success: true,
    row: newRow,
    transNum: newTransNum,
    summary: summaryText
  };
}

// ==================== نظام الأوردر المشترك ====================
/**
 * إنشاء رقم أوردر فريد
 * الصيغة: ORD-YYMMDD-NNN
 * مثال: ORD-250121-001
 */
function generateOrderNumber() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.TRANSACTIONS);

  if (!sheet) {
    throw new Error('شيت دفتر الحركات المالية غير موجود');
  }

  const today = new Date();
  const dateStr = Utilities.formatDate(today, Session.getScriptTimeZone(), 'yyMMdd');
  const prefix = `ORD-${dateStr}-`;

  // البحث عن آخر رقم أوردر لنفس اليوم
  const lastRow = sheet.getLastRow();
  let maxSeq = 0;

  if (lastRow > 1) {
    // عمود Z (26) = رقم الأوردر
    const orderNumbers = sheet.getRange(2, 26, lastRow - 1, 1).getValues();

    for (const row of orderNumbers) {
      const orderNum = String(row[0] || '').trim();
      if (orderNum.startsWith(prefix)) {
        const seqPart = orderNum.substring(prefix.length);
        const seq = parseInt(seqPart, 10);
        if (!isNaN(seq) && seq > maxSeq) {
          maxSeq = seq;
        }
      }
    }
  }

  // الرقم التسلسلي الجديد
  const newSeq = String(maxSeq + 1).padStart(3, '0');
  return prefix + newSeq;
}

/**
 * عرض نموذج الأوردر المشترك
 */
function showSharedOrderForm() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // جلب قوائم البيانات
  const vendors = getVendorsList_();
  const items = getItemsList_();
  const projects = getProjectsList_();

  // إنشاء رقم الأوردر المقترح
  const suggestedOrderNumber = generateOrderNumber();

  const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html dir="rtl">
    <head>
      <base target="_top">
      <style>
        * { box-sizing: border-box; font-family: Arial, sans-serif; }
        body { padding: 20px; background: #f5f5f5; margin: 0; }
        h2 { color: #1a73e8; margin-bottom: 20px; text-align: center; }
        .form-group { margin-bottom: 15px; }
        label { display: block; margin-bottom: 5px; font-weight: bold; color: #333; }
        input, select, textarea {
          width: 100%;
          padding: 10px;
          border: 1px solid #ddd;
          border-radius: 5px;
          font-size: 14px;
        }
        input:focus, select:focus { border-color: #1a73e8; outline: none; }
        .order-info {
          background: #e8f5e9;
          padding: 15px;
          border-radius: 8px;
          margin-bottom: 20px;
          border: 1px solid #4caf50;
        }
        .order-number {
          font-size: 18px;
          font-weight: bold;
          color: #2e7d32;
          text-align: center;
        }
        .projects-section {
          background: #fff;
          padding: 15px;
          border-radius: 8px;
          border: 1px solid #ddd;
          margin-bottom: 15px;
        }
        .projects-section h3 { margin-top: 0; color: #555; }
        .project-row {
          display: flex;
          flex-wrap: wrap;
          gap: 8px;
          align-items: center;
          margin-bottom: 10px;
          padding: 10px;
          background: #f9f9f9;
          border-radius: 5px;
        }
        .project-row select { flex: 2; min-width: 150px; }
        .project-row .guest-count { width: 60px; flex: 0 0 60px; }
        .project-row .guest-details { flex: 2; min-width: 120px; }
        .project-row .remove-btn {
          background: #f44336;
          color: white;
          border: none;
          padding: 8px 12px;
          border-radius: 5px;
          cursor: pointer;
        }
        .add-project-btn {
          background: #4caf50;
          color: white;
          border: none;
          padding: 10px 20px;
          border-radius: 5px;
          cursor: pointer;
          width: 100%;
        }
        .summary-section {
          background: #fff3e0;
          padding: 15px;
          border-radius: 8px;
          margin-bottom: 15px;
          border: 1px solid #ff9800;
        }
        .summary-row {
          display: flex;
          justify-content: space-between;
          margin-bottom: 8px;
        }
        .summary-row.total {
          font-weight: bold;
          font-size: 16px;
          border-top: 2px solid #ff9800;
          padding-top: 10px;
        }
        .payment-section {
          background: #e3f2fd;
          padding: 15px;
          border-radius: 8px;
          margin-bottom: 15px;
          border: 1px solid #2196f3;
        }
        .payment-section h3 { margin-top: 0; color: #1565c0; }
        .payment-row {
          display: flex;
          gap: 15px;
          margin-bottom: 10px;
        }
        .payment-row .form-group { margin-bottom: 0; }
        .btn-primary {
          background: #1a73e8;
          color: white;
          border: none;
          padding: 12px 30px;
          border-radius: 5px;
          cursor: pointer;
          font-size: 16px;
          width: 100%;
        }
        .btn-primary:hover { background: #1557b0; }
        .btn-primary:disabled { background: #ccc; cursor: not-allowed; }
        .error { color: #f44336; font-size: 12px; }
        .loading { text-align: center; padding: 20px; display: none; }
        .spinner {
          border: 3px solid #f3f3f3;
          border-top: 3px solid #1a73e8;
          border-radius: 50%;
          width: 30px;
          height: 30px;
          animation: spin 1s linear infinite;
          margin: 0 auto;
        }
        @keyframes spin {
          0% { transform: rotate(0deg); }
          100% { transform: rotate(360deg); }
        }
      </style>
    </head>
    <body>
      <h2>📦 أوردر مشترك (تقسيم بين مشاريع)</h2>

      <div class="order-info">
        <div class="order-number">رقم الأوردر: ${suggestedOrderNumber}</div>
      </div>

      <div id="formContent">
        <div class="form-group">
          <label>📅 التاريخ</label>
          <input type="date" id="orderDate" value="${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd')}">
        </div>

        <div class="form-group">
          <label>👤 المورد</label>
          <select id="vendor">
            <option value="">اختر المورد...</option>
            ${vendors.map(v => '<option value="' + v + '">' + v + '</option>').join('')}
          </select>
        </div>

        <div class="form-group">
          <label>📁 البند</label>
          <select id="item" onchange="onItemChange()">
            <option value="">اختر البند...</option>
            ${items.map(i => '<option value="' + i + '">' + i + '</option>').join('')}
          </select>
        </div>

        <div class="form-group" id="unitCountSection" style="display:none; background: #e8f5e9; padding: 10px; border-radius: 5px; border: 1px solid #a5d6a7;">
          <label>📊 عدد الوحدات الإجمالي <span id="unitTypeLabel" style="color: #2e7d32;"></span></label>
          <input type="number" id="totalUnitCount" min="0" step="1" placeholder="اختياري - سيتم توزيعه على المشاريع">
          <small style="color: #666; font-size: 11px;">💡 سيتم توزيع الوحدات على المشاريع بنفس نسبة الضيوف</small>
        </div>

        <div class="form-group">
          <label>💵 إجمالي المبلغ (USD)</label>
          <input type="number" id="totalAmount" step="0.01" min="0" placeholder="0.00" onchange="calculateDistribution()">
        </div>

        <div class="form-group">
          <label>📝 وصف الأوردر (اختياري)</label>
          <input type="text" id="orderDescription" placeholder="مثال: تصوير حلقات شهر يناير...">
        </div>

        <div class="payment-section">
          <h3>💳 بيانات الدفع</h3>
          <div class="payment-row">
            <div class="form-group" style="flex:1;">
              <label>طريقة الدفع</label>
              <select id="paymentMethod">
                <option value="">اختر...</option>
                <option value="نقدي">نقدي</option>
                <option value="تحويل بنكي">تحويل بنكي</option>
                <option value="شيك">شيك</option>
                <option value="بطاقة">بطاقة</option>
                <option value="أخرى">أخرى</option>
              </select>
            </div>
            <div class="form-group" style="flex:1;">
              <label>نوع شرط الدفع</label>
              <select id="paymentTermType" onchange="togglePaymentFields()">
                <option value="فوري">فوري</option>
                <option value="بعد التسليم">بعد التسليم</option>
                <option value="تاريخ مخصص">تاريخ مخصص</option>
              </select>
            </div>
          </div>
          <div class="payment-row" id="weeksRow" style="display:none;">
            <div class="form-group">
              <label>عدد الأسابيع بعد التسليم</label>
              <input type="number" id="weeksCount" min="0" value="0" placeholder="0">
            </div>
          </div>
          <div class="payment-row" id="customDateRow" style="display:none;">
            <div class="form-group">
              <label>تاريخ الاستحقاق المخصص</label>
              <input type="date" id="customDueDate">
            </div>
          </div>
        </div>

        <div class="projects-section">
          <h3>🎬 المشاريع والضيوف</h3>
          <div id="projectsList">
            <div class="project-row" data-index="0">
              <select class="project-select" onchange="calculateDistribution()">
                <option value="">اختر المشروع...</option>
                ${projects.map(p => '<option value="' + p.code + '">' + p.code + ' - ' + p.name + '</option>').join('')}
              </select>
              <input type="number" class="guest-count" min="1" value="1" placeholder="عدد" onchange="calculateDistribution()">
              <input type="text" class="guest-details" placeholder="أسماء الضيوف...">
              <button type="button" class="remove-btn" onclick="removeProject(0)" style="display:none;">✕</button>
            </div>
          </div>
          <button type="button" class="add-project-btn" onclick="addProject()">+ إضافة مشروع</button>
        </div>

        <div class="summary-section" id="summarySection" style="display:none;">
          <h3>📊 ملخص التوزيع</h3>
          <div id="distributionSummary"></div>
        </div>

        <button type="button" class="btn-primary" id="submitBtn" onclick="submitOrder()">
          💾 حفظ الأوردر المشترك
        </button>
      </div>

      <div class="loading" id="loading">
        <div class="spinner"></div>
        <p>جاري حفظ الأوردر...</p>
      </div>

      <script>
        const projects = ${JSON.stringify(projects)};
        let projectIndex = 1;

        // جدول ربط البنود بأنواع الوحدات
        const unitTypes = {
          'تصوير': 'مقابلة',
          'مونتاج': 'دقيقة',
          'مكساج': 'دقيقة',
          'دوبلاج': 'دقيقة',
          'تلوين': 'دقيقة',
          'جرافيك - رسم': 'رسمة',
          'فيكسر': 'ضيف',
          'تعليق صوتي': 'دقيقة',
          'اقتباسات': 'دقيقة'
        };

        // عند تغيير البند - إظهار/إخفاء حقل الوحدات
        function onItemChange() {
          const item = document.getElementById('item').value;
          const unitSection = document.getElementById('unitCountSection');
          const unitTypeLabel = document.getElementById('unitTypeLabel');
          const unitType = unitTypes[item];

          if (unitType) {
            unitSection.style.display = 'block';
            unitTypeLabel.textContent = '(' + unitType + ')';
          } else {
            unitSection.style.display = 'none';
            unitTypeLabel.textContent = '';
            document.getElementById('totalUnitCount').value = '';
          }
        }

        function togglePaymentFields() {
          const termType = document.getElementById('paymentTermType').value;
          document.getElementById('weeksRow').style.display = termType === 'بعد التسليم' ? 'flex' : 'none';
          document.getElementById('customDateRow').style.display = termType === 'تاريخ مخصص' ? 'flex' : 'none';
        }

        function addProject() {
          const container = document.getElementById('projectsList');
          const div = document.createElement('div');
          div.className = 'project-row';
          div.dataset.index = projectIndex;
          div.innerHTML = \`
            <select class="project-select" onchange="calculateDistribution()">
              <option value="">اختر المشروع...</option>
              \${projects.map(p => '<option value="' + p.code + '">' + p.code + ' - ' + p.name + '</option>').join('')}
            </select>
            <input type="number" class="guest-count" min="1" value="1" placeholder="عدد" onchange="calculateDistribution()">
            <input type="text" class="guest-details" placeholder="أسماء الضيوف...">
            <button type="button" class="remove-btn" onclick="removeProject(\${projectIndex})">✕</button>
          \`;
          container.appendChild(div);
          projectIndex++;
          updateRemoveButtons();
          calculateDistribution();
        }

        function removeProject(index) {
          const row = document.querySelector('.project-row[data-index="' + index + '"]');
          if (row) row.remove();
          updateRemoveButtons();
          calculateDistribution();
        }

        function updateRemoveButtons() {
          const rows = document.querySelectorAll('.project-row');
          rows.forEach((row, i) => {
            const btn = row.querySelector('.remove-btn');
            btn.style.display = rows.length > 1 ? 'block' : 'none';
          });
        }

        function calculateDistribution() {
          const totalAmount = parseFloat(document.getElementById('totalAmount').value) || 0;
          const rows = document.querySelectorAll('.project-row');
          const summarySection = document.getElementById('summarySection');
          const summaryDiv = document.getElementById('distributionSummary');

          if (totalAmount <= 0 || rows.length === 0) {
            summarySection.style.display = 'none';
            return;
          }

          let totalGuests = 0;
          const projectData = [];

          rows.forEach(row => {
            const select = row.querySelector('.project-select');
            const guestInput = row.querySelector('.guest-count');
            const projectCode = select.value;
            const guests = parseInt(guestInput.value) || 0;

            if (projectCode && guests > 0) {
              const projectInfo = projects.find(p => p.code === projectCode);
              projectData.push({
                code: projectCode,
                name: projectInfo ? projectInfo.name : '',
                guests: guests
              });
              totalGuests += guests;
            }
          });

          if (totalGuests === 0 || projectData.length === 0) {
            summarySection.style.display = 'none';
            return;
          }

          let html = '';
          projectData.forEach(p => {
            const share = (p.guests / totalGuests) * totalAmount;
            html += '<div class="summary-row">';
            html += '<span>' + p.code + ' (' + p.guests + ' ضيوف)</span>';
            html += '<span>$' + share.toFixed(2) + '</span>';
            html += '</div>';
          });

          html += '<div class="summary-row total">';
          html += '<span>الإجمالي (' + totalGuests + ' ضيوف)</span>';
          html += '<span>$' + totalAmount.toFixed(2) + '</span>';
          html += '</div>';

          summaryDiv.innerHTML = html;
          summarySection.style.display = 'block';
        }

        function submitOrder() {
          // التحقق من البيانات
          const vendor = document.getElementById('vendor').value;
          const item = document.getElementById('item').value;
          const totalAmount = parseFloat(document.getElementById('totalAmount').value) || 0;
          const orderDate = document.getElementById('orderDate').value;
          const orderDescription = document.getElementById('orderDescription').value;

          if (!vendor) {
            alert('برجاء اختيار المورد');
            return;
          }
          if (!item) {
            alert('برجاء اختيار البند');
            return;
          }
          if (totalAmount <= 0) {
            alert('برجاء إدخال المبلغ');
            return;
          }

          // جمع بيانات المشاريع مع التفاصيل الخاصة بكل مشروع
          const rows = document.querySelectorAll('.project-row');
          const projectsData = [];
          let totalGuests = 0;

          rows.forEach(row => {
            const select = row.querySelector('.project-select');
            const guestInput = row.querySelector('.guest-count');
            const detailsInput = row.querySelector('.guest-details');
            const projectCode = select.value;
            const guests = parseInt(guestInput.value) || 0;
            const guestDetails = detailsInput ? detailsInput.value.trim() : '';

            if (projectCode && guests > 0) {
              projectsData.push({
                code: projectCode,
                guests: guests,
                details: guestDetails  // تفاصيل/أسماء ضيوف هذا المشروع
              });
              totalGuests += guests;
            }
          });

          if (projectsData.length === 0) {
            alert('برجاء إضافة مشروع واحد على الأقل');
            return;
          }

          // جمع بيانات الدفع
          const paymentMethod = document.getElementById('paymentMethod').value;
          const paymentTermType = document.getElementById('paymentTermType').value;
          const weeksCount = parseInt(document.getElementById('weeksCount').value) || 0;
          const customDueDate = document.getElementById('customDueDate').value;

          // إظهار التحميل
          document.getElementById('formContent').style.display = 'none';
          document.getElementById('loading').style.display = 'block';

          // عدد الوحدات الإجمالي (اختياري)
          const totalUnitCount = parseInt(document.getElementById('totalUnitCount').value) || 0;

          // إرسال البيانات
          const orderData = {
            orderNumber: '${suggestedOrderNumber}',
            orderDate: orderDate,
            vendor: vendor,
            item: item,
            totalAmount: totalAmount,
            totalUnitCount: totalUnitCount,  // عدد الوحدات الإجمالي (سيتم توزيعه على المشاريع)
            orderDescription: orderDescription,  // وصف عام للأوردر
            projects: projectsData,  // كل مشروع له تفاصيله الخاصة
            totalGuests: totalGuests,
            // بيانات الدفع
            paymentMethod: paymentMethod,
            paymentTermType: paymentTermType,
            weeksCount: weeksCount,
            customDueDate: customDueDate
          };

          google.script.run
            .withSuccessHandler(function(result) {
              if (result.success) {
                alert('تم حفظ الأوردر بنجاح!\\n\\n' + result.message);
                google.script.host.close();
              } else {
                alert('خطأ: ' + result.error);
                document.getElementById('formContent').style.display = 'block';
                document.getElementById('loading').style.display = 'none';
              }
            })
            .withFailureHandler(function(error) {
              alert('خطأ: ' + error.message);
              document.getElementById('formContent').style.display = 'block';
              document.getElementById('loading').style.display = 'none';
            })
            .saveSharedOrder(orderData);
        }
      </script>
    </body>
    </html>
  `).setWidth(500).setHeight(700);

  ui.showModalDialog(html, '📦 أوردر مشترك');
}

/**
 * جلب قائمة الموردين
 */
function getVendorsList_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.PARTIES);
  if (!sheet || sheet.getLastRow() < 2) return [];

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
  // فقط الموردين (النوع = مورد)
  return data.filter(row => String(row[1]).includes('مورد')).map(row => row[0]).filter(Boolean);
}

/**
 * جلب قائمة البنود
 */
function getItemsList_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.ITEMS);
  if (!sheet || sheet.getLastRow() < 2) return [];

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
  return data.map(row => row[0]).filter(Boolean);
}

/**
 * جلب قائمة المشاريع
 */
function getProjectsList_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.PROJECTS);
  if (!sheet || sheet.getLastRow() < 2) return [];

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
  return data.filter(row => row[0]).map(row => ({ code: row[0], name: row[1] || '' }));
}

/**
 * حفظ الأوردر المشترك كحركات متعددة
 */
function saveSharedOrder(orderData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.TRANSACTIONS);
    const projectsSheet = ss.getSheetByName(CONFIG.SHEETS.PROJECTS);

    if (!sheet) {
      return { success: false, error: 'شيت دفتر الحركات المالية غير موجود' };
    }

    // تحويل التاريخ
    const dateParts = orderData.orderDate.split('-');
    const transDate = new Date(dateParts[0], dateParts[1] - 1, dateParts[2]);

    // جلب بيانات المشاريع
    let projectsData = {};
    if (projectsSheet && projectsSheet.getLastRow() > 1) {
      const pData = projectsSheet.getRange(2, 1, projectsSheet.getLastRow() - 1, 2).getValues();
      for (const row of pData) {
        projectsData[row[0]] = row[1]; // code -> name
      }
    }

    const savedRows = [];
    const totalGuests = orderData.totalGuests;
    const totalUnitCount = orderData.totalUnitCount || 0; // عدد الوحدات الإجمالي (اختياري)

    // حفظ حركة لكل مشروع
    for (const project of orderData.projects) {
      // حساب حصة المشروع بناءً على عدد الضيوف
      const share = (project.guests / totalGuests) * orderData.totalAmount;
      // حساب حصة المشروع من الوحدات (بنفس نسبة الضيوف)
      const unitShare = totalUnitCount > 0 ? Math.round((project.guests / totalGuests) * totalUnitCount) : 0;
      const projectName = projectsData[project.code] || '';

      // حساب رقم الحركة الجديد
      const lastRow = sheet.getLastRow();
      const newRow = lastRow + 1;
      const newTransNum = lastRow > 1 ?
        (Number(sheet.getRange(lastRow, 1).getValue()) || 0) + 1 : 1;

      // التفاصيل: أسماء الضيوف الخاصة بهذا المشروع + معلومات التقسيم
      let detailsText = '';
      if (project.details) {
        // إذا كان هناك تفاصيل خاصة بالمشروع (أسماء الضيوف)
        detailsText = `مقابلة ${project.details} (${project.guests} من ${totalGuests})`;
      } else if (orderData.orderDescription) {
        // إذا كان هناك وصف عام للأوردر
        detailsText = `${orderData.orderDescription} | ${project.guests} ضيوف من ${totalGuests}`;
      } else {
        // الافتراضي
        detailsText = `أوردر مشترك - ${project.guests} ضيوف من ${totalGuests}`;
      }

      // الشهر
      const monthStr = Utilities.formatDate(transDate, Session.getScriptTimeZone(), 'yyyy-MM');

      // كتابة البيانات
      // A: رقم الحركة
      sheet.getRange(newRow, 1).setValue(newTransNum);

      // B: التاريخ
      sheet.getRange(newRow, 2).setValue(transDate).setNumberFormat('dd/mm/yyyy');

      // C: طبيعة الحركة
      sheet.getRange(newRow, 3).setValue('استحقاق مصروف');

      // D: تصنيف الحركة
      sheet.getRange(newRow, 4).setValue('مصروفات مباشرة');

      // E: كود المشروع
      sheet.getRange(newRow, 5).setValue(project.code);

      // F: اسم المشروع
      sheet.getRange(newRow, 6).setValue(projectName);

      // G: البند
      sheet.getRange(newRow, 7).setValue(orderData.item);

      // H: التفاصيل
      sheet.getRange(newRow, 8).setValue(detailsText);

      // I: اسم المورد
      sheet.getRange(newRow, 9).setValue(orderData.vendor);

      // J: المبلغ
      sheet.getRange(newRow, 10).setValue(share).setNumberFormat('#,##0.00');

      // K: العملة
      sheet.getRange(newRow, 11).setValue('USD');

      // L: سعر الصرف
      sheet.getRange(newRow, 12).setValue(1);

      // M: القيمة بالدولار
      sheet.getRange(newRow, 13).setValue(share).setNumberFormat('#,##0.00');

      // N: نوع الحركة
      sheet.getRange(newRow, 14).setValue('مدين استحقاق');

      // O: الرصيد (سيُحسب لاحقاً)
      // skip - will be calculated

      // P: رقم مرجعي
      sheet.getRange(newRow, 16).setValue(orderData.orderNumber);

      // Q: طريقة الدفع
      sheet.getRange(newRow, 17).setValue(orderData.paymentMethod || '');

      // R: نوع شرط الدفع
      sheet.getRange(newRow, 18).setValue(orderData.paymentTermType || 'فوري');

      // S: عدد الأسابيع
      sheet.getRange(newRow, 19).setValue(orderData.weeksCount || '');

      // T: تاريخ مخصص
      if (orderData.customDueDate) {
        const dateParts = orderData.customDueDate.split('-');
        const customDate = new Date(dateParts[0], dateParts[1] - 1, dateParts[2]);
        sheet.getRange(newRow, 20).setValue(customDate).setNumberFormat('dd/mm/yyyy');
      }

      // W: الشهر
      sheet.getRange(newRow, 23).setValue(monthStr);

      // Z: رقم الأوردر (عمود 26)
      sheet.getRange(newRow, 26).setValue(orderData.orderNumber);

      // AA: عدد الوحدات (عمود 27) - حصة المشروع من الوحدات
      if (unitShare > 0) {
        sheet.getRange(newRow, 27).setValue(unitShare);
      }

      savedRows.push({
        row: newRow,
        transNum: newTransNum,
        project: project.code,
        amount: share,
        details: project.details || '',
        guests: project.guests,
        units: unitShare  // إضافة الوحدات للملخص
      });
    }

    SpreadsheetApp.flush();

    // ═══════════════════════════════════════════════════════════════
    // إعادة حساب الأرصدة وتواريخ الاستحقاق لجميع الحركات المضافة
    // ═══════════════════════════════════════════════════════════════
    for (const saved of savedRows) {
      // حساب تاريخ الاستحقاق (U)
      calculateDueDate_(ss, sheet, saved.row);
    }
    // إعادة حساب الرصيد (O) وحالة السداد (V) لكل حركات الطرف (مرة واحدة تكفي)
    recalculatePartyBalance_(sheet, savedRows[0].row);

    SpreadsheetApp.flush();

    // تسجيل النشاط
    logActivity(
      'أوردر مشترك',
      CONFIG.SHEETS.TRANSACTIONS,
      savedRows[0].row,
      orderData.orderNumber,
      `${orderData.vendor} - ${orderData.totalAmount} USD - ${orderData.projects.length} مشاريع`,
      {
        orderNumber: orderData.orderNumber,
        vendor: orderData.vendor,
        totalAmount: orderData.totalAmount,
        projectsCount: orderData.projects.length,
        totalGuests: totalGuests
      }
    );

    // رسالة النجاح مع تفاصيل الضيوف والوحدات
    const projectsList = savedRows.map(r => {
      let line = `• ${r.project}: $${r.amount.toFixed(2)} (${r.guests} ضيوف)`;
      if (r.units > 0) {
        line += ` [${r.units} وحدات]`;
      }
      if (r.details) {
        line += `\\n  ↳ ${r.details}`;
      }
      return line;
    }).join('\\n');

    // إضافة معلومات الوحدات للرسالة إذا كانت موجودة
    let unitsInfo = '';
    if (totalUnitCount > 0) {
      unitsInfo = `\\nإجمالي الوحدات: ${totalUnitCount}`;
    }

    return {
      success: true,
      message: `رقم الأوردر: ${orderData.orderNumber}\\n` +
               `المورد: ${orderData.vendor}\\n` +
               `الإجمالي: $${orderData.totalAmount.toFixed(2)}${unitsInfo}\\n` +
               `عدد الحركات: ${savedRows.length}\\n\\n` +
               `التوزيع:\\n${projectsList}`
    };

  } catch (error) {
    console.error('خطأ في saveSharedOrder:', error);
    return { success: false, error: error.message };
  }
}

// ==================== تقرير مقارنة الميزانية المجمع ====================
/**
 * عرض نموذج اختيار المشاريع للتقرير المجمع
 */
function showCombinedBudgetReportForm() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // جلب قائمة المشاريع
  const projects = getProjectsList_();

  if (projects.length === 0) {
    ui.alert('⚠️ لا توجد مشاريع في قاعدة البيانات');
    return;
  }

  // بناء HTML قائمة المشاريع
  let projectsHtml = '';
  for (let i = 0; i < projects.length; i++) {
    const p = projects[i];
    projectsHtml += '<label class="project-item">' +
      '<input type="checkbox" name="project" value="' + p.code + '" onchange="updateCount()">' +
      '<div class="project-info">' +
      '<div class="project-code">' + p.code + '</div>' +
      '<div class="project-name">' + p.name + '</div>' +
      '</div></label>';
  }

  const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html dir="rtl">
    <head>
      <base target="_top">
      <style>
        * { box-sizing: border-box; font-family: Arial, sans-serif; }
        body { padding: 20px; background: #f5f5f5; margin: 0; }
        h2 { color: #1a73e8; margin-bottom: 20px; text-align: center; }
        .info-box {
          background: #e3f2fd;
          padding: 12px;
          border-radius: 8px;
          margin-bottom: 15px;
          border-right: 4px solid #1a73e8;
          font-size: 13px;
        }
        .projects-list {
          max-height: 350px;
          overflow-y: auto;
          border: 1px solid #ddd;
          border-radius: 8px;
          background: #fff;
          padding: 10px;
          margin-bottom: 15px;
        }
        .project-item {
          display: flex;
          align-items: center;
          padding: 10px;
          border-bottom: 1px solid #eee;
          cursor: pointer;
          transition: background 0.2s;
        }
        .project-item:hover { background: #f5f5f5; }
        .project-item:last-child { border-bottom: none; }
        .project-item input[type="checkbox"] {
          width: 18px;
          height: 18px;
          margin-left: 10px;
        }
        .project-info { flex: 1; }
        .project-code { font-weight: bold; color: #1a73e8; }
        .project-name { color: #666; font-size: 13px; }
        .selection-controls {
          display: flex;
          gap: 10px;
          margin-bottom: 10px;
        }
        .selection-controls button {
          padding: 8px 15px;
          border: 1px solid #ddd;
          border-radius: 5px;
          background: #fff;
          cursor: pointer;
          font-size: 12px;
        }
        .selection-controls button:hover { background: #f0f0f0; }
        .selected-count {
          text-align: center;
          padding: 10px;
          background: #fff3e0;
          border-radius: 5px;
          margin-bottom: 15px;
          font-weight: bold;
        }
        .btn-primary {
          background: #1a73e8;
          color: white;
          border: none;
          padding: 12px 30px;
          border-radius: 5px;
          cursor: pointer;
          font-size: 16px;
          width: 100%;
        }
        .btn-primary:hover { background: #1557b0; }
        .btn-primary:disabled { background: #ccc; cursor: not-allowed; }
        .loading { text-align: center; padding: 20px; display: none; }
        .spinner {
          border: 3px solid #f3f3f3;
          border-top: 3px solid #1a73e8;
          border-radius: 50%;
          width: 30px;
          height: 30px;
          animation: spin 1s linear infinite;
          margin: 0 auto;
        }
        @keyframes spin {
          0% { transform: rotate(0deg); }
          100% { transform: rotate(360deg); }
        }
      </style>
    </head>
    <body>
      <h2>📊 تقرير مقارنة الميزانية المجمع</h2>

      <div class="info-box">
        💡 اختر المشاريع التي تريد مقارنة ميزانياتها بالمصروفات الفعلية.
        سيتم إنشاء تقرير يعرض كل مشروع على حدة مع إجمالي مجمع.
      </div>

      <div id="formContent">
        <div class="selection-controls">
          <button type="button" onclick="selectAll()">تحديد الكل</button>
          <button type="button" onclick="deselectAll()">إلغاء الكل</button>
        </div>

        <div class="projects-list">
          ${projectsHtml}
        </div>

        <div class="selected-count" id="selectedCount">
          تم اختيار: 0 مشاريع
        </div>

        <button type="button" class="btn-primary" id="submitBtn" onclick="generateReport()" disabled>
          📊 إنشاء التقرير المجمع
        </button>
      </div>

      <div class="loading" id="loading">
        <div class="spinner"></div>
        <p>جاري إنشاء التقرير...</p>
      </div>

      <script>
        function selectAll() {
          document.querySelectorAll('input[name="project"]').forEach(cb => cb.checked = true);
          updateCount();
        }

        function deselectAll() {
          document.querySelectorAll('input[name="project"]').forEach(cb => cb.checked = false);
          updateCount();
        }

        function updateCount() {
          const checked = document.querySelectorAll('input[name="project"]:checked').length;
          document.getElementById('selectedCount').textContent = 'تم اختيار: ' + checked + ' مشاريع';
          document.getElementById('submitBtn').disabled = checked === 0;
        }

        function generateReport() {
          const selected = [];
          document.querySelectorAll('input[name="project"]:checked').forEach(cb => {
            selected.push(cb.value);
          });

          if (selected.length === 0) {
            alert('برجاء اختيار مشروع واحد على الأقل');
            return;
          }

          document.getElementById('formContent').style.display = 'none';
          document.getElementById('loading').style.display = 'block';

          google.script.run
            .withSuccessHandler(function(result) {
              if (result.success) {
                alert('✅ تم إنشاء التقرير بنجاح!\\n\\n' + result.message);
                google.script.host.close();
              } else {
                alert('❌ خطأ: ' + result.error);
                document.getElementById('formContent').style.display = 'block';
                document.getElementById('loading').style.display = 'none';
              }
            })
            .withFailureHandler(function(error) {
              alert('❌ خطأ: ' + error.message);
              document.getElementById('formContent').style.display = 'block';
              document.getElementById('loading').style.display = 'none';
            })
            .generateCombinedBudgetReport(selected);
        }
      </script>
    </body>
    </html>
  `).setWidth(450).setHeight(550);

  ui.showModalDialog(html, '📊 تقرير مقارنة الميزانية المجمع');
}

/**
 * إنشاء تقرير مقارنة الميزانية المجمع لعدة مشاريع
 */
function generateCombinedBudgetReport(projectCodes) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // 1️⃣ جلب بيانات المشاريع
    const projectsSheet = ss.getSheetByName(CONFIG.SHEETS.PROJECTS);
    const budgetSheet = ss.getSheetByName(CONFIG.SHEETS.BUDGETS);
    const transSheet = ss.getSheetByName(CONFIG.SHEETS.TRANSACTIONS);

    if (!projectsSheet || !transSheet) {
      return { success: false, error: 'الشيتات المطلوبة غير موجودة' };
    }

    const projectsData = projectsSheet.getDataRange().getValues();
    const transData = transSheet.getDataRange().getValues();
    const budgetData = budgetSheet ? budgetSheet.getDataRange().getValues() : [];

    // تحديد أعمدة الحركات
    const transHeaders = transData[0];
    const colC = transHeaders.indexOf('طبيعة الحركة') !== -1 ? transHeaders.indexOf('طبيعة الحركة') : 2;
    const colE = transHeaders.indexOf('كود المشروع') !== -1 ? transHeaders.indexOf('كود المشروع') : 4;
    const colG = transHeaders.indexOf('البند') !== -1 ? transHeaders.indexOf('البند') : 6;
    const colM = transHeaders.indexOf('القيمة بالدولار') !== -1 ? transHeaders.indexOf('القيمة بالدولار') : 12;

    // 2️⃣ جمع البيانات لكل مشروع
    const projectsReport = [];
    let grandTotalPlanned = 0;
    let grandTotalActual = 0;

    for (const code of projectCodes) {
      const projectCode = code.trim().toUpperCase();

      // معلومات المشروع
      let projectInfo = null;
      for (let i = 1; i < projectsData.length; i++) {
        if (String(projectsData[i][0]).trim().toUpperCase() === projectCode) {
          projectInfo = {
            code: projectsData[i][0],
            name: projectsData[i][1] || ''
          };
          break;
        }
      }

      if (!projectInfo) continue;

      // الميزانية المخططة
      const plannedBudget = {};
      let totalPlanned = 0;

      if (budgetData.length > 1) {
        for (let i = 1; i < budgetData.length; i++) {
          const budgetProjCode = String(budgetData[i][0]).trim().toUpperCase();
          if (budgetProjCode === projectCode) {
            const item = String(budgetData[i][2] || '').trim();
            const amount = Number(budgetData[i][3]) || 0;
            if (item && amount > 0) {
              plannedBudget[item] = (plannedBudget[item] || 0) + amount;
              totalPlanned += amount;
            }
          }
        }
      }

      // المصروفات الفعلية (استحقاق مصروف فقط)
      const actualExpenses = {};
      let totalActual = 0;

      for (let i = 1; i < transData.length; i++) {
        const rowProjCode = String(transData[i][colE] || '').trim().toUpperCase();
        if (rowProjCode !== projectCode) continue;

        const item = String(transData[i][colG] || '').trim();
        const amountUsd = Number(transData[i][colM]) || 0;
        const natureType = String(transData[i][colC] || '').trim();

        if (natureType.indexOf('استحقاق مصروف') === -1) continue;
        if (!item || amountUsd <= 0) continue;

        actualExpenses[item] = (actualExpenses[item] || 0) + amountUsd;
        totalActual += amountUsd;
      }

      // تجميع البنود
      const allItems = new Set([...Object.keys(plannedBudget), ...Object.keys(actualExpenses)]);
      const itemsData = [];

      for (const item of allItems) {
        const planned = plannedBudget[item] || 0;
        const actual = actualExpenses[item] || 0;
        const variance = planned - actual;
        itemsData.push({ item, planned, actual, variance });
      }

      // ترتيب حسب المصروف الفعلي
      itemsData.sort((a, b) => b.actual - a.actual);

      projectsReport.push({
        code: projectInfo.code,
        name: projectInfo.name,
        items: itemsData,
        totalPlanned,
        totalActual,
        totalVariance: totalPlanned - totalActual
      });

      grandTotalPlanned += totalPlanned;
      grandTotalActual += totalActual;
    }

    if (projectsReport.length === 0) {
      return { success: false, error: 'لم يتم العثور على أي من المشاريع المحددة' };
    }

    // 3️⃣ إنشاء شيت التقرير
    const reportSheetName = 'تقرير ميزانية مجمع';
    let reportSheet = ss.getSheetByName(reportSheetName);
    if (reportSheet) {
      ss.deleteSheet(reportSheet);
    }
    reportSheet = ss.insertSheet(reportSheetName);
    reportSheet.setRightToLeft(true);
    reportSheet.setTabColor('#9c27b0');

    // عرض الأعمدة
    reportSheet.setColumnWidth(1, 200);  // البند
    reportSheet.setColumnWidth(2, 120);  // المخطط
    reportSheet.setColumnWidth(3, 120);  // الفعلي
    reportSheet.setColumnWidth(4, 120);  // الفرق
    reportSheet.setColumnWidth(5, 80);   // النسبة

    let currentRow = 1;

    // ═══════════════════════════════════════════════════════════
    // العنوان الرئيسي
    // ═══════════════════════════════════════════════════════════
    reportSheet.getRange(currentRow, 1, 1, 5).merge()
      .setValue('📊 تقرير مقارنة الميزانية المجمع')
      .setBackground('#9c27b0')
      .setFontColor('#ffffff')
      .setFontWeight('bold')
      .setFontSize(16)
      .setHorizontalAlignment('center');
    currentRow++;

    reportSheet.getRange(currentRow, 1, 1, 5).merge()
      .setValue('عدد المشاريع: ' + projectsReport.length + ' | تاريخ التقرير: ' + new Date().toLocaleDateString('ar-EG'))
      .setBackground('#e1bee7')
      .setHorizontalAlignment('center');
    currentRow += 2;

    // ═══════════════════════════════════════════════════════════
    // تقرير كل مشروع
    // ═══════════════════════════════════════════════════════════
    for (const project of projectsReport) {
      // عنوان المشروع
      reportSheet.getRange(currentRow, 1, 1, 5).merge()
        .setValue('🎬 ' + project.code + ' - ' + project.name)
        .setBackground('#7b1fa2')
        .setFontColor('#ffffff')
        .setFontWeight('bold')
        .setFontSize(12);
      currentRow++;

      // رأس الجدول
      const headers = ['البند', 'المخطط ($)', 'الفعلي ($)', 'الفرق ($)', '%'];
      reportSheet.getRange(currentRow, 1, 1, 5).setValues([headers])
        .setBackground('#ce93d8')
        .setFontWeight('bold')
        .setHorizontalAlignment('center');
      currentRow++;

      // بيانات البنود
      if (project.items.length > 0) {
        for (const item of project.items) {
          const percentage = item.planned > 0 ? Math.round((item.actual / item.planned) * 100) : (item.actual > 0 ? 999 : 0);
          const rowData = [
            item.item,
            item.planned || '',
            item.actual || '',
            item.variance,
            percentage + '%'
          ];
          reportSheet.getRange(currentRow, 1, 1, 5).setValues([rowData]);

          // تلوين الفرق
          const varianceCell = reportSheet.getRange(currentRow, 4);
          if (item.variance < 0) {
            varianceCell.setBackground('#ffcdd2').setFontColor('#c62828');
          } else if (item.variance > 0) {
            varianceCell.setBackground('#c8e6c9').setFontColor('#2e7d32');
          }

          currentRow++;
        }
      } else {
        reportSheet.getRange(currentRow, 1, 1, 5).merge()
          .setValue('لا توجد بيانات')
          .setFontStyle('italic')
          .setHorizontalAlignment('center');
        currentRow++;
      }

      // إجمالي المشروع
      reportSheet.getRange(currentRow, 1, 1, 5).setValues([[
        'إجمالي ' + project.code,
        project.totalPlanned,
        project.totalActual,
        project.totalVariance,
        project.totalPlanned > 0 ? Math.round((project.totalActual / project.totalPlanned) * 100) + '%' : '-'
      ]])
        .setBackground('#f3e5f5')
        .setFontWeight('bold');

      // تلوين إجمالي الفرق
      const totalVarianceCell = reportSheet.getRange(currentRow, 4);
      if (project.totalVariance < 0) {
        totalVarianceCell.setBackground('#ffcdd2').setFontColor('#c62828');
      } else if (project.totalVariance > 0) {
        totalVarianceCell.setBackground('#c8e6c9').setFontColor('#2e7d32');
      }

      currentRow += 2;
    }

    // ═══════════════════════════════════════════════════════════
    // الإجمالي الكلي
    // ═══════════════════════════════════════════════════════════
    reportSheet.getRange(currentRow, 1, 1, 5).merge()
      .setValue('📊 الإجمالي الكلي لجميع المشاريع')
      .setBackground('#4a148c')
      .setFontColor('#ffffff')
      .setFontWeight('bold')
      .setFontSize(12);
    currentRow++;

    const grandVariance = grandTotalPlanned - grandTotalActual;
    reportSheet.getRange(currentRow, 1, 1, 5).setValues([[
      'الإجمالي',
      grandTotalPlanned,
      grandTotalActual,
      grandVariance,
      grandTotalPlanned > 0 ? Math.round((grandTotalActual / grandTotalPlanned) * 100) + '%' : '-'
    ]])
      .setBackground('#e1bee7')
      .setFontWeight('bold')
      .setFontSize(12);

    // تلوين الفرق الكلي
    const grandVarianceCell = reportSheet.getRange(currentRow, 4);
    if (grandVariance < 0) {
      grandVarianceCell.setBackground('#ffcdd2').setFontColor('#c62828').setFontWeight('bold');
    } else if (grandVariance > 0) {
      grandVarianceCell.setBackground('#c8e6c9').setFontColor('#2e7d32').setFontWeight('bold');
    }

    currentRow += 2;

    // ═══════════════════════════════════════════════════════════
    // 📊 مقارنة البنود المجمعة من كل المشاريع
    // ═══════════════════════════════════════════════════════════

    // تجميع كل البنود من كل المشاريع
    const aggregatedItems = {};
    for (const project of projectsReport) {
      for (const item of project.items) {
        if (!aggregatedItems[item.item]) {
          aggregatedItems[item.item] = { planned: 0, actual: 0 };
        }
        aggregatedItems[item.item].planned += item.planned;
        aggregatedItems[item.item].actual += item.actual;
      }
    }

    // تحويل إلى مصفوفة وترتيب حسب الفعلي
    const aggregatedArray = [];
    for (const itemName in aggregatedItems) {
      const data = aggregatedItems[itemName];
      aggregatedArray.push({
        item: itemName,
        planned: data.planned,
        actual: data.actual,
        variance: data.planned - data.actual
      });
    }
    aggregatedArray.sort((a, b) => b.actual - a.actual);

    // عنوان القسم
    reportSheet.getRange(currentRow, 1, 1, 5).merge()
      .setValue('📊 مقارنة البنود المجمعة (كل المشاريع)')
      .setBackground('#ff6f00')
      .setFontColor('#ffffff')
      .setFontWeight('bold')
      .setFontSize(14)
      .setHorizontalAlignment('center');
    currentRow++;

    reportSheet.getRange(currentRow, 1, 1, 5).merge()
      .setValue('هذا القسم يجمع كل بند من جميع المشاريع المختارة لمعرفة هل البند ككل ضمن الحد المسموح')
      .setBackground('#ffe0b2')
      .setFontSize(11)
      .setHorizontalAlignment('center');
    currentRow++;

    // رأس الجدول
    reportSheet.getRange(currentRow, 1, 1, 5).setValues([['البند', 'المخطط الكلي ($)', 'الفعلي الكلي ($)', 'الفرق ($)', 'الحالة']])
      .setBackground('#ffb74d')
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
    currentRow++;

    // بيانات البنود المجمعة
    for (const item of aggregatedArray) {
      const percentage = item.planned > 0 ? Math.round((item.actual / item.planned) * 100) : (item.actual > 0 ? 999 : 0);
      let status = '';
      if (item.variance > 0) {
        status = '✅ وفر ' + Math.round(item.variance) + '$';
      } else if (item.variance < 0) {
        status = '❌ تجاوز ' + Math.round(Math.abs(item.variance)) + '$';
      } else {
        status = '⚖️ متطابق';
      }

      reportSheet.getRange(currentRow, 1, 1, 5).setValues([[
        item.item,
        item.planned,
        item.actual,
        item.variance,
        status
      ]]);

      // تلوين الفرق والحالة
      const varCell = reportSheet.getRange(currentRow, 4);
      const statusCell = reportSheet.getRange(currentRow, 5);
      if (item.variance < 0) {
        varCell.setBackground('#ffcdd2').setFontColor('#c62828');
        statusCell.setBackground('#ffcdd2').setFontColor('#c62828');
      } else if (item.variance > 0) {
        varCell.setBackground('#c8e6c9').setFontColor('#2e7d32');
        statusCell.setBackground('#c8e6c9').setFontColor('#2e7d32');
      }

      currentRow++;
    }

    // إجمالي البنود المجمعة
    currentRow++;
    const aggTotalPlanned = aggregatedArray.reduce((sum, i) => sum + i.planned, 0);
    const aggTotalActual = aggregatedArray.reduce((sum, i) => sum + i.actual, 0);
    const aggTotalVariance = aggTotalPlanned - aggTotalActual;
    let aggStatus = aggTotalVariance >= 0 ? '✅ ضمن الميزانية' : '❌ تجاوز الميزانية';

    reportSheet.getRange(currentRow, 1, 1, 5).setValues([[
      '📊 الإجمالي الكلي للبنود',
      aggTotalPlanned,
      aggTotalActual,
      aggTotalVariance,
      aggStatus
    ]])
      .setBackground('#ff6f00')
      .setFontColor('#ffffff')
      .setFontWeight('bold')
      .setFontSize(12);

    // تنسيقات الأرقام
    const lastDataRow = currentRow;
    reportSheet.getRange(4, 2, lastDataRow - 3, 3).setNumberFormat('$#,##0.00');

    // تفعيل الشيت
    ss.setActiveSheet(reportSheet);

    return {
      success: true,
      message: 'عدد المشاريع: ' + projectsReport.length + '\n' +
               'إجمالي المخطط: $' + grandTotalPlanned.toLocaleString() + '\n' +
               'إجمالي الفعلي: $' + grandTotalActual.toLocaleString() + '\n' +
               'الفرق: $' + grandVariance.toLocaleString()
    };

  } catch (error) {
    console.error('خطأ في generateCombinedBudgetReport:', error);
    return { success: false, error: error.message };
  }
}

// ==================== تقرير تكلفة الوحدة ====================

/**
 * عرض نموذج اختيار مشروع واحد لتقرير تكلفة الوحدة
 */
function showUnitCostReportSingle() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const projectsSheet = ss.getSheetByName(CONFIG.SHEETS.PROJECTS);

  if (!projectsSheet || projectsSheet.getLastRow() < 2) {
    ui.alert('❌ خطأ', 'لا توجد مشاريع في قاعدة البيانات!', ui.ButtonSet.OK);
    return;
  }

  // جلب قائمة المشاريع
  const projectsData = projectsSheet.getRange(2, 1, projectsSheet.getLastRow() - 1, 2).getValues();
  const projectsList = projectsData.filter(row => row[0]).map(row => row[0] + ' - ' + row[1]).join('\n');

  const response = ui.prompt(
    '📊 تقرير تكلفة الوحدة - مشروع واحد',
    'أدخل كود المشروع:\n\nالمشاريع المتاحة:\n' + projectsList,
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) return;

  const projectCode = response.getResponseText().trim().toUpperCase();
  if (!projectCode) {
    ui.alert('❌ خطأ', 'يرجى إدخال كود المشروع!', ui.ButtonSet.OK);
    return;
  }

  // التحقق من وجود المشروع
  const projectExists = projectsData.some(row => String(row[0]).toUpperCase() === projectCode);
  if (!projectExists) {
    ui.alert('❌ خطأ', 'المشروع "' + projectCode + '" غير موجود!', ui.ButtonSet.OK);
    return;
  }

  generateUnitCostReport([projectCode]);
}

/**
 * عرض نموذج اختيار عدة مشاريع لتقرير تكلفة الوحدة
 */
function showUnitCostReportMultiple() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const projectsSheet = ss.getSheetByName(CONFIG.SHEETS.PROJECTS);

  if (!projectsSheet || projectsSheet.getLastRow() < 2) {
    ui.alert('❌ خطأ', 'لا توجد مشاريع في قاعدة البيانات!', ui.ButtonSet.OK);
    return;
  }

  // جلب قائمة المشاريع
  const projectsData = projectsSheet.getRange(2, 1, projectsSheet.getLastRow() - 1, 2).getValues();
  const projectsList = projectsData.filter(row => row[0]).map(row => row[0] + ' - ' + row[1]).join('\n');

  const response = ui.prompt(
    '📊 تقرير تكلفة الوحدة - عدة مشاريع',
    'أدخل أكواد المشاريع (مفصولة بفاصلة):\n\nمثال: F001, F002, F003\n\nالمشاريع المتاحة:\n' + projectsList,
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) return;

  const input = response.getResponseText().trim();
  if (!input) {
    ui.alert('❌ خطأ', 'يرجى إدخال أكواد المشاريع!', ui.ButtonSet.OK);
    return;
  }

  const projectCodes = input.split(',').map(code => code.trim().toUpperCase()).filter(code => code);

  if (projectCodes.length === 0) {
    ui.alert('❌ خطأ', 'يرجى إدخال كود مشروع واحد على الأقل!', ui.ButtonSet.OK);
    return;
  }

  generateUnitCostReport(projectCodes);
}

/**
 * توليد تقرير تكلفة الوحدة لكل المشاريع
 */
function generateUnitCostReportAll() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const projectsSheet = ss.getSheetByName(CONFIG.SHEETS.PROJECTS);

  if (!projectsSheet || projectsSheet.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('❌ خطأ', 'لا توجد مشاريع في قاعدة البيانات!', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  // جلب كل المشاريع
  const projectsData = projectsSheet.getRange(2, 1, projectsSheet.getLastRow() - 1, 1).getValues();
  const allProjectCodes = projectsData.filter(row => row[0]).map(row => String(row[0]).trim());

  if (allProjectCodes.length === 0) {
    SpreadsheetApp.getUi().alert('❌ خطأ', 'لا توجد مشاريع!', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  generateUnitCostReport(allProjectCodes);
}

/**
 * توليد تقرير تكلفة الوحدة
 * @param {string[]} projectCodes - قائمة أكواد المشاريع
 */
function generateUnitCostReport(projectCodes) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ui = SpreadsheetApp.getUi();
    const transSheet = ss.getSheetByName(CONFIG.SHEETS.TRANSACTIONS);
    const projectsSheet = ss.getSheetByName(CONFIG.SHEETS.PROJECTS);

    if (!transSheet) {
      ui.alert('❌ خطأ', 'شيت دفتر الحركات المالية غير موجود!', ui.ButtonSet.OK);
      return;
    }

    // جلب بيانات المشاريع
    let projectsData = {};
    if (projectsSheet && projectsSheet.getLastRow() > 1) {
      const pData = projectsSheet.getRange(2, 1, projectsSheet.getLastRow() - 1, 2).getValues();
      for (const row of pData) {
        projectsData[String(row[0]).toUpperCase()] = row[1];
      }
    }

    // جلب بيانات الحركات
    const lastRow = transSheet.getLastRow();
    if (lastRow < 2) {
      ui.alert('❌ خطأ', 'لا توجد حركات في دفتر الحركات!', ui.ButtonSet.OK);
      return;
    }

    // الأعمدة: A=رقم, B=تاريخ, C=طبيعة, D=تصنيف, E=كود المشروع, F=اسم المشروع, G=البند,
    //          H=التفاصيل, I=المورد, M=القيمة بالدولار (13), AA=عدد الوحدات (27)
    const allData = transSheet.getRange(2, 1, lastRow - 1, 27).getValues();

    // فلترة الحركات للمشاريع المحددة (استحقاقات مصروف فقط)
    const filteredData = allData.filter(row => {
      const projectCode = String(row[4] || '').toUpperCase(); // E = كود المشروع
      const natureType = String(row[2] || ''); // C = طبيعة الحركة
      return projectCodes.map(c => c.toUpperCase()).includes(projectCode) &&
             natureType === 'استحقاق مصروف';
    });

    if (filteredData.length === 0) {
      ui.alert('⚠️ تنبيه', 'لا توجد حركات استحقاق مصروف للمشاريع المحددة!', ui.ButtonSet.OK);
      return;
    }

    // ═══════════════════════════════════════════════════════════════
    // تجميع البيانات
    // ═══════════════════════════════════════════════════════════════

    // 1. تجميع حسب البند (الإجمالي)
    const byItem = {};
    // 2. تجميع حسب المشروع والبند
    const byProjectItem = {};
    // 3. تجميع حسب المورد والبند
    const byVendorItem = {};

    for (const row of filteredData) {
      const projectCode = String(row[4] || '').toUpperCase();
      const item = String(row[6] || ''); // G = البند
      const vendor = String(row[8] || ''); // I = المورد
      const amountUsd = Number(row[12]) || 0; // M = القيمة بالدولار
      const unitCount = Number(row[26]) || 0; // AA = عدد الوحدات

      // تجميع حسب البند
      if (!byItem[item]) {
        byItem[item] = { totalAmount: 0, totalUnits: 0, unitType: CONFIG.UNIT_TYPES[item] || '' };
      }
      byItem[item].totalAmount += amountUsd;
      byItem[item].totalUnits += unitCount;

      // تجميع حسب المشروع والبند
      const projectKey = projectCode + '|' + item;
      if (!byProjectItem[projectKey]) {
        byProjectItem[projectKey] = {
          projectCode: projectCode,
          projectName: projectsData[projectCode] || '',
          item: item,
          totalAmount: 0,
          totalUnits: 0,
          unitType: CONFIG.UNIT_TYPES[item] || ''
        };
      }
      byProjectItem[projectKey].totalAmount += amountUsd;
      byProjectItem[projectKey].totalUnits += unitCount;

      // تجميع حسب المورد والبند
      if (vendor) {
        const vendorKey = vendor + '|' + item;
        if (!byVendorItem[vendorKey]) {
          byVendorItem[vendorKey] = {
            vendor: vendor,
            item: item,
            totalAmount: 0,
            totalUnits: 0,
            unitType: CONFIG.UNIT_TYPES[item] || ''
          };
        }
        byVendorItem[vendorKey].totalAmount += amountUsd;
        byVendorItem[vendorKey].totalUnits += unitCount;
      }
    }

    // ═══════════════════════════════════════════════════════════════
    // إنشاء شيت التقرير
    // ═══════════════════════════════════════════════════════════════

    const reportName = 'تقرير تكلفة الوحدة';
    let reportSheet = ss.getSheetByName(reportName);
    if (reportSheet) {
      ss.deleteSheet(reportSheet);
    }
    reportSheet = ss.insertSheet(reportName);
    reportSheet.setTabColor('#4caf50');

    let currentRow = 1;

    // ═══════════════════════════════════════════════════════════════
    // العنوان الرئيسي
    // ═══════════════════════════════════════════════════════════════

    const reportTitle = projectCodes.length === 1 ?
      '📊 تقرير تكلفة الوحدة - مشروع ' + projectCodes[0] :
      '📊 تقرير تكلفة الوحدة - ' + projectCodes.length + ' مشاريع';

    reportSheet.getRange(currentRow, 1, 1, 6).merge()
      .setValue(reportTitle)
      .setBackground('#4caf50')
      .setFontColor('#ffffff')
      .setFontWeight('bold')
      .setFontSize(16)
      .setHorizontalAlignment('center');
    currentRow += 2;

    // معلومات التقرير
    const reportDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
    reportSheet.getRange(currentRow, 1).setValue('تاريخ التقرير: ' + reportDate);
    reportSheet.getRange(currentRow, 4).setValue('المشاريع: ' + projectCodes.join(', '));
    currentRow += 2;

    // ═══════════════════════════════════════════════════════════════
    // القسم الأول: ملخص تكلفة الوحدات (حسب البند)
    // ═══════════════════════════════════════════════════════════════

    reportSheet.getRange(currentRow, 1, 1, 6).merge()
      .setValue('📋 القسم الأول: ملخص تكلفة الوحدات حسب البند')
      .setBackground('#1565c0')
      .setFontColor('#ffffff')
      .setFontWeight('bold')
      .setFontSize(13);
    currentRow++;

    // رأس الجدول
    reportSheet.getRange(currentRow, 1, 1, 6).setValues([[
      'البند', 'نوع الوحدة', 'العدد الكلي', 'التكلفة الكلية ($)', 'تكلفة الوحدة ($)', 'ملاحظات'
    ]])
      .setBackground('#1976d2')
      .setFontColor('#ffffff')
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
    currentRow++;

    // بيانات البنود
    const itemsArray = Object.entries(byItem).sort((a, b) => b[1].totalAmount - a[1].totalAmount);
    let grandTotalAmount = 0;
    let grandTotalUnits = 0;

    for (const [item, data] of itemsArray) {
      const unitCost = data.totalUnits > 0 ? data.totalAmount / data.totalUnits : 0;
      const notes = data.totalUnits === 0 ? '⚠️ لا توجد وحدات مسجلة' : '';

      reportSheet.getRange(currentRow, 1, 1, 6).setValues([[
        item,
        data.unitType || '-',
        data.totalUnits || '-',
        data.totalAmount,
        data.totalUnits > 0 ? unitCost : '-',
        notes
      ]]);

      // تلوين الصفوف
      if (currentRow % 2 === 0) {
        reportSheet.getRange(currentRow, 1, 1, 6).setBackground('#e3f2fd');
      }

      // تحذير إذا لا توجد وحدات
      if (data.totalUnits === 0 && data.unitType) {
        reportSheet.getRange(currentRow, 6).setBackground('#fff3e0').setFontColor('#e65100');
      }

      grandTotalAmount += data.totalAmount;
      grandTotalUnits += data.totalUnits;
      currentRow++;
    }

    // صف الإجمالي
    reportSheet.getRange(currentRow, 1, 1, 6).setValues([[
      '📊 الإجمالي',
      '-',
      grandTotalUnits,
      grandTotalAmount,
      '-',
      ''
    ]])
      .setBackground('#4caf50')
      .setFontColor('#ffffff')
      .setFontWeight('bold');
    currentRow += 2;

    // ═══════════════════════════════════════════════════════════════
    // القسم الثاني: تفصيل حسب المشروع
    // ═══════════════════════════════════════════════════════════════

    reportSheet.getRange(currentRow, 1, 1, 6).merge()
      .setValue('🎬 القسم الثاني: تفصيل حسب المشروع')
      .setBackground('#7b1fa2')
      .setFontColor('#ffffff')
      .setFontWeight('bold')
      .setFontSize(13);
    currentRow++;

    // رأس الجدول
    reportSheet.getRange(currentRow, 1, 1, 6).setValues([[
      'المشروع', 'البند', 'نوع الوحدة', 'العدد', 'التكلفة ($)', 'تكلفة الوحدة ($)'
    ]])
      .setBackground('#9c27b0')
      .setFontColor('#ffffff')
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
    currentRow++;

    // ترتيب حسب المشروع ثم البند
    const projectItemsArray = Object.values(byProjectItem).sort((a, b) => {
      if (a.projectCode !== b.projectCode) return a.projectCode.localeCompare(b.projectCode);
      return b.totalAmount - a.totalAmount;
    });

    let lastProject = '';
    for (const data of projectItemsArray) {
      const unitCost = data.totalUnits > 0 ? data.totalAmount / data.totalUnits : 0;
      const projectDisplay = data.projectCode + (data.projectName ? ' - ' + data.projectName : '');

      // إضافة فاصل بين المشاريع
      if (lastProject && lastProject !== data.projectCode) {
        reportSheet.getRange(currentRow, 1, 1, 6).setBackground('#f5f5f5');
        currentRow++;
      }
      lastProject = data.projectCode;

      reportSheet.getRange(currentRow, 1, 1, 6).setValues([[
        projectDisplay,
        data.item,
        data.unitType || '-',
        data.totalUnits || '-',
        data.totalAmount,
        data.totalUnits > 0 ? unitCost : '-'
      ]]);

      currentRow++;
    }
    currentRow++;

    // ═══════════════════════════════════════════════════════════════
    // القسم الثالث: تفصيل حسب المورد
    // ═══════════════════════════════════════════════════════════════

    reportSheet.getRange(currentRow, 1, 1, 6).merge()
      .setValue('👤 القسم الثالث: تفصيل حسب المورد')
      .setBackground('#00695c')
      .setFontColor('#ffffff')
      .setFontWeight('bold')
      .setFontSize(13);
    currentRow++;

    // رأس الجدول
    reportSheet.getRange(currentRow, 1, 1, 6).setValues([[
      'المورد', 'البند', 'نوع الوحدة', 'العدد', 'التكلفة ($)', 'تكلفة الوحدة ($)'
    ]])
      .setBackground('#009688')
      .setFontColor('#ffffff')
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
    currentRow++;

    // ترتيب حسب المورد ثم البند
    const vendorItemsArray = Object.values(byVendorItem).sort((a, b) => {
      if (a.vendor !== b.vendor) return a.vendor.localeCompare(b.vendor);
      return b.totalAmount - a.totalAmount;
    });

    let lastVendor = '';
    for (const data of vendorItemsArray) {
      const unitCost = data.totalUnits > 0 ? data.totalAmount / data.totalUnits : 0;

      // إضافة فاصل بين الموردين
      if (lastVendor && lastVendor !== data.vendor) {
        reportSheet.getRange(currentRow, 1, 1, 6).setBackground('#f5f5f5');
        currentRow++;
      }
      lastVendor = data.vendor;

      reportSheet.getRange(currentRow, 1, 1, 6).setValues([[
        data.vendor,
        data.item,
        data.unitType || '-',
        data.totalUnits || '-',
        data.totalAmount,
        data.totalUnits > 0 ? unitCost : '-'
      ]]);

      currentRow++;
    }

    // ═══════════════════════════════════════════════════════════════
    // تنسيقات عامة
    // ═══════════════════════════════════════════════════════════════

    // تنسيق الأعمدة
    reportSheet.setColumnWidth(1, 200);
    reportSheet.setColumnWidth(2, 150);
    reportSheet.setColumnWidth(3, 100);
    reportSheet.setColumnWidth(4, 120);
    reportSheet.setColumnWidth(5, 130);
    reportSheet.setColumnWidth(6, 150);

    // تنسيق الأرقام
    reportSheet.getRange(1, 4, currentRow, 1).setNumberFormat('$#,##0.00');
    reportSheet.getRange(1, 5, currentRow, 1).setNumberFormat('$#,##0.00');

    // تفعيل الشيت
    ss.setActiveSheet(reportSheet);

    ui.alert(
      '✅ تم إنشاء التقرير',
      'تقرير تكلفة الوحدة جاهز!\n\n' +
      '• عدد المشاريع: ' + projectCodes.length + '\n' +
      '• عدد البنود: ' + itemsArray.length + '\n' +
      '• إجمالي التكلفة: $' + grandTotalAmount.toLocaleString(),
      ui.ButtonSet.OK
    );

  } catch (error) {
    console.error('خطأ في generateUnitCostReport:', error);
    SpreadsheetApp.getUi().alert('❌ خطأ', 'حدث خطأ: ' + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}