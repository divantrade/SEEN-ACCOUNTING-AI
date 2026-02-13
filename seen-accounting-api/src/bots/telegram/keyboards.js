import { InlineKeyboard } from 'grammy';
import {
  NATURE_TYPES,
  CLASSIFICATION_TYPES,
  CURRENCIES,
  PAYMENT_METHODS,
  PAYMENT_TERMS,
  PARTY_TYPES,
} from '../../config/constants.js';

// ==================== لوحات المفاتيح ====================

export function mainMenuKeyboard() {
  return new InlineKeyboard()
    .text('📝 مصروف', 'nature:expense')
    .text('💰 إيراد', 'nature:revenue')
    .row()
    .text('🏦 تمويل', 'nature:finance')
    .text('🔒 تأمين', 'nature:insurance')
    .row()
    .text('🔄 تحويل داخلي', 'nature:transfer')
    .row()
    .text('📊 حالة الحركات', 'cmd:status');
}

export function expenseNatureKeyboard() {
  return new InlineKeyboard()
    .text('📋 استحقاق مصروف', `nat:${NATURE_TYPES.EXPENSE_ACCRUAL}`)
    .row()
    .text('💸 دفعة مصروف', `nat:${NATURE_TYPES.EXPENSE_PAYMENT}`)
    .row()
    .text('🔙 رجوع', 'back:main');
}

export function revenueNatureKeyboard() {
  return new InlineKeyboard()
    .text('📋 استحقاق إيراد', `nat:${NATURE_TYPES.REVENUE_ACCRUAL}`)
    .row()
    .text('💰 تحصيل إيراد', `nat:${NATURE_TYPES.REVENUE_COLLECTION}`)
    .row()
    .text('🔙 رجوع', 'back:main');
}

export function classificationKeyboard(nature) {
  const kb = new InlineKeyboard();

  if (nature.includes('مصروف')) {
    kb.text('🎬 مصروفات مباشرة', `cls:${CLASSIFICATION_TYPES.DIRECT_EXPENSE}`);
    kb.row();
    kb.text('🏢 مصروفات عمومية', `cls:${CLASSIFICATION_TYPES.GENERAL_EXPENSE}`);
  } else if (nature.includes('إيراد')) {
    kb.text('💵 ايراد', `cls:${CLASSIFICATION_TYPES.REVENUE}`);
  }

  kb.row().text('🔙 رجوع', 'back:nature');
  return kb;
}

export function currencyKeyboard() {
  return new InlineKeyboard()
    .text(`$ دولار`, 'cur:USD')
    .text(`₺ ليرة`, 'cur:TRY')
    .text(`ج.م جنيه`, 'cur:EGP')
    .row()
    .text('🔙 رجوع', 'back:amount');
}

export function paymentMethodKeyboard() {
  return new InlineKeyboard()
    .text('💵 نقدي', `pm:${PAYMENT_METHODS.CASH}`)
    .text('🏦 تحويل بنكي', `pm:${PAYMENT_METHODS.BANK_TRANSFER}`)
    .row()
    .text('📝 شيك', `pm:${PAYMENT_METHODS.CHECK}`)
    .text('💳 بطاقة', `pm:${PAYMENT_METHODS.CARD}`)
    .row()
    .text('🔙 رجوع', 'back:details');
}

export function paymentTermKeyboard() {
  return new InlineKeyboard()
    .text('⚡ فوري', `pt:${PAYMENT_TERMS.IMMEDIATE}`)
    .row()
    .text('📦 بعد التسليم', `pt:${PAYMENT_TERMS.AFTER_DELIVERY}`)
    .row()
    .text('📅 تاريخ مخصص', `pt:${PAYMENT_TERMS.CUSTOM_DATE}`)
    .row()
    .text('🔙 رجوع', 'back:payment_method');
}

export function confirmationKeyboard() {
  return new InlineKeyboard()
    .text('✅ تأكيد وإرسال', 'confirm:yes')
    .row()
    .text('✏️ تعديل', 'confirm:edit')
    .text('❌ إلغاء', 'confirm:cancel')
    .row()
    .text('📎 إرفاق ملف', 'confirm:attach');
}

export function partyTypeKeyboard() {
  return new InlineKeyboard()
    .text('🏭 مورد', `pty:${PARTY_TYPES.VENDOR}`)
    .text('👤 عميل', `pty:${PARTY_TYPES.CLIENT}`)
    .text('💼 ممول', `pty:${PARTY_TYPES.FUNDER}`)
    .row()
    .text('🔙 رجوع', 'back:party');
}

export function projectSelectionKeyboard(projects) {
  const kb = new InlineKeyboard();
  projects.forEach((p, i) => {
    kb.text(`${p.name}`, `proj:${p.id}`);
    if ((i + 1) % 2 === 0) kb.row();
  });
  kb.row().text('🔙 رجوع', 'back:classification');
  return kb;
}

export function reviewKeyboard(pendingId) {
  return new InlineKeyboard()
    .text('✅ اعتماد', `review:approve:${pendingId}`)
    .text('❌ رفض', `review:reject:${pendingId}`)
    .row()
    .text('✏️ طلب تعديل', `review:edit:${pendingId}`);
}
