// =============================================
// SEEN Accounting - Constants & Enums
// Migrated from Config.js (Google Apps Script)
// =============================================

// طبيعة الحركة
export const NATURE_TYPES = {
  EXPENSE_ACCRUAL: 'استحقاق مصروف',
  EXPENSE_PAYMENT: 'دفعة مصروف',
  REVENUE_ACCRUAL: 'استحقاق إيراد',
  REVENUE_COLLECTION: 'تحصيل إيراد',
  FINANCING: 'تمويل',
  FINANCING_RECEIVED: 'استلام تمويل',
  FINANCING_REPAYMENT: 'سداد تمويل',
  INSURANCE_PAID: 'تأمين مدفوع للقناة',
  INSURANCE_RETURNED: 'استرداد تأمين من القناة',
  INTERNAL_TRANSFER: 'تحويل داخلي',
};

// تصنيف الحركة
export const CLASSIFICATION_TYPES = {
  DIRECT_EXPENSE: 'مصروفات مباشرة',
  GENERAL_EXPENSE: 'مصروفات عمومية',
  REVENUE: 'ايراد',
  LONG_TERM_FINANCING: 'تمويل طويل الأجل',
  LONG_TERM_REPAYMENT: 'سداد تمويل طويل الأجل',
  SHORT_ADVANCE: 'سلفة قصيرة',
  SHORT_ADVANCE_REPAY: 'سداد سلفة قصيرة',
  INSURANCE: 'تأمين',
  INSURANCE_RETURN: 'استرداد تأمين',
  TRANSFER_TO_BANK: 'تحويل للبنك',
  TRANSFER_TO_CASH: 'تحويل للخزنة',
};

// نوع الحركة (مدين/دائن)
export const MOVEMENT_TYPES = {
  DEBIT: 'مدين استحقاق',
  CREDIT: 'دائن دفعة',
};

// العملات
export const CURRENCIES = {
  USD: { code: 'USD', name: 'دولار', symbol: '$' },
  TRY: { code: 'TRY', name: 'ليرة', symbol: '₺' },
  EGP: { code: 'EGP', name: 'جنيه مصري', symbol: 'ج.م' },
};

// طرق الدفع
export const PAYMENT_METHODS = {
  CASH: 'نقدي',
  BANK_TRANSFER: 'تحويل بنكي',
  CHECK: 'شيك',
  CARD: 'بطاقة',
  OTHER: 'أخرى',
};

// شروط الدفع
export const PAYMENT_TERMS = {
  IMMEDIATE: 'فوري',
  AFTER_DELIVERY: 'بعد التسليم',
  CUSTOM_DATE: 'تاريخ مخصص',
};

// حالة السداد
export const PAYMENT_STATUS = {
  PAID: 'مدفوع بالكامل',
  PENDING: 'معلق',
  IN_PROCESS: 'عملية دفع/تحصيل',
};

// أنواع الأطراف
export const PARTY_TYPES = {
  VENDOR: 'مورد',
  CLIENT: 'عميل',
  FUNDER: 'ممول',
};

// حالة المراجعة
export const REVIEW_STATUS = {
  PENDING: 'قيد الانتظار',
  APPROVED: 'معتمد',
  REJECTED: 'مرفوض',
  NEEDS_EDIT: 'يحتاج تعديل',
};

// أنواع الوحدات للبنود
export const UNIT_TYPES = {
  'تصوير': 'مقابلة',
  'مونتاج': 'دقيقة',
  'مكساج': 'دقيقة',
  'دوبلاج': 'دقيقة',
  'تلوين': 'دقيقة',
  'جرافيك - رسم': 'رسمة',
  'فيكسر': 'ضيف',
  'تعليق صوتي': 'دقيقة',
  'اقتباسات': 'دقيقة',
};

// تحديد نوع الحركة من طبيعة الحركة
export function getMovementType(nature) {
  const debitNatures = [
    NATURE_TYPES.EXPENSE_ACCRUAL,
    NATURE_TYPES.REVENUE_ACCRUAL,
    NATURE_TYPES.FINANCING,
    NATURE_TYPES.INSURANCE_PAID,
  ];
  return debitNatures.includes(nature)
    ? MOVEMENT_TYPES.DEBIT
    : MOVEMENT_TYPES.CREDIT;
}

// قواعد الحقول المطلوبة حسب التصنيف
export const REQUIRED_FIELDS = {
  ALWAYS: ['amount', 'party'],
  BY_CLASSIFICATION: {
    [CLASSIFICATION_TYPES.DIRECT_EXPENSE]: ['project'],
    [CLASSIFICATION_TYPES.REVENUE]: ['project'],
    [CLASSIFICATION_TYPES.GENERAL_EXPENSE]: [],
  },
};
