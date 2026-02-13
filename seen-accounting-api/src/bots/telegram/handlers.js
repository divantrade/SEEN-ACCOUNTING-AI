import { getDB } from '../../config/database.js';
import { getRedis } from '../../config/redis.js';
import { checkTelegramAuth } from '../../modules/auth/auth.service.js';
import { transactionService } from '../../modules/transactions/transactions.service.js';
import { getMovementType, CURRENCIES } from '../../config/constants.js';
import { fuzzySearch } from '../../shared/fuzzy-search.js';
import { logger } from '../../shared/logger.js';
import {
  mainMenuKeyboard,
  expenseNatureKeyboard,
  revenueNatureKeyboard,
  classificationKeyboard,
  currencyKeyboard,
  paymentMethodKeyboard,
  paymentTermKeyboard,
  confirmationKeyboard,
  projectSelectionKeyboard,
  reviewKeyboard,
} from './keyboards.js';

const db = () => getDB();
const redis = () => getRedis();

// ==================== حالة المحادثة ====================

async function getState(userId) {
  const state = await db().conversationState.findUnique({
    where: { userId },
  });
  return state || { state: 'idle', context: {} };
}

async function setState(userId, state, context = {}) {
  await db().conversationState.upsert({
    where: { userId },
    create: {
      userId,
      botType: 'traditional',
      state,
      context,
    },
    update: { state, context },
  });
}

async function clearState(userId) {
  await db().conversationState.deleteMany({ where: { userId } });
}

// ==================== تسجيل الأوامر ====================

export function registerCommands(bot) {
  bot.command('start', handleStart);
  bot.command('expense', handleExpense);
  bot.command('revenue', handleRevenue);
  bot.command('status', handleStatus);
  bot.command('cancel', handleCancel);
  bot.command('help', handleHelp);

  // Handle callback queries (inline keyboard buttons)
  bot.on('callback_query:data', handleCallback);

  // Handle text messages
  bot.on('message:text', handleTextMessage);

  // Handle file/photo attachments
  bot.on('message:document', handleAttachment);
  bot.on('message:photo', handleAttachment);
}

// ==================== معالجة الأوامر ====================

async function handleStart(ctx) {
  const user = await checkTelegramAuth(ctx.chat.id, 'traditional_bot');
  if (!user) {
    return ctx.reply('⛔ غير مصرح لك باستخدام البوت.\nتواصل مع المسؤول للحصول على صلاحية.');
  }

  await clearState(user.id);
  await ctx.reply(
    `مرحباً ${user.name} 👋\n\nاختر نوع الحركة المالية:`,
    { reply_markup: mainMenuKeyboard() }
  );
}

async function handleExpense(ctx) {
  const user = await checkTelegramAuth(ctx.chat.id, 'traditional_bot');
  if (!user) return ctx.reply('⛔ غير مصرح لك.');

  await ctx.reply('اختر نوع المصروف:', { reply_markup: expenseNatureKeyboard() });
}

async function handleRevenue(ctx) {
  const user = await checkTelegramAuth(ctx.chat.id, 'traditional_bot');
  if (!user) return ctx.reply('⛔ غير مصرح لك.');

  await ctx.reply('اختر نوع الإيراد:', { reply_markup: revenueNatureKeyboard() });
}

async function handleStatus(ctx) {
  const user = await checkTelegramAuth(ctx.chat.id, 'traditional_bot');
  if (!user) return ctx.reply('⛔ غير مصرح لك.');

  const pending = await db().pendingTransaction.count({
    where: { submittedById: user.id, reviewStatus: 'PENDING' },
  });
  const approved = await db().pendingTransaction.count({
    where: { submittedById: user.id, reviewStatus: 'APPROVED' },
  });
  const rejected = await db().pendingTransaction.count({
    where: { submittedById: user.id, reviewStatus: 'REJECTED' },
  });

  await ctx.reply(
    `📊 *حالة حركاتك:*\n\n` +
    `⏳ قيد الانتظار: ${pending}\n` +
    `✅ معتمدة: ${approved}\n` +
    `❌ مرفوضة: ${rejected}`,
    { parse_mode: 'Markdown' }
  );
}

async function handleCancel(ctx) {
  const user = await checkTelegramAuth(ctx.chat.id, 'traditional_bot');
  if (!user) return;

  await clearState(user.id);
  await ctx.reply('🚫 تم إلغاء العملية الحالية.\n\nاختر عملية جديدة:', {
    reply_markup: mainMenuKeyboard(),
  });
}

async function handleHelp(ctx) {
  await ctx.reply(
    `📖 *أوامر البوت:*\n\n` +
    `/start - بدء محادثة جديدة\n` +
    `/expense - تسجيل مصروف\n` +
    `/revenue - تسجيل إيراد\n` +
    `/status - حالة الحركات\n` +
    `/cancel - إلغاء العملية\n` +
    `/help - المساعدة`,
    { parse_mode: 'Markdown' }
  );
}

// ==================== معالجة الأزرار ====================

async function handleCallback(ctx) {
  const user = await checkTelegramAuth(ctx.chat.id, 'traditional_bot');
  if (!user) return ctx.answerCallbackQuery('⛔ غير مصرح');

  const data = ctx.callbackQuery.data;
  const [action, value, extra] = data.split(':');

  try {
    switch (action) {
      case 'nature':
        await handleNatureSelection(ctx, user, value);
        break;
      case 'nat':
        await handleNatureType(ctx, user, value);
        break;
      case 'cls':
        await handleClassification(ctx, user, value);
        break;
      case 'proj':
        await handleProjectSelection(ctx, user, value);
        break;
      case 'cur':
        await handleCurrency(ctx, user, value);
        break;
      case 'pm':
        await handlePaymentMethod(ctx, user, value);
        break;
      case 'pt':
        await handlePaymentTerm(ctx, user, value);
        break;
      case 'confirm':
        await handleConfirmation(ctx, user, value);
        break;
      case 'review':
        await handleReview(ctx, user, value, extra);
        break;
      case 'back':
        await handleBack(ctx, user, value);
        break;
      default:
        break;
    }
    await ctx.answerCallbackQuery();
  } catch (err) {
    logger.error('Callback error:', err);
    await ctx.answerCallbackQuery('❌ حدث خطأ');
  }
}

async function handleNatureSelection(ctx, user, category) {
  switch (category) {
    case 'expense':
      await ctx.editMessageText('اختر نوع المصروف:', {
        reply_markup: expenseNatureKeyboard(),
      });
      break;
    case 'revenue':
      await ctx.editMessageText('اختر نوع الإيراد:', {
        reply_markup: revenueNatureKeyboard(),
      });
      break;
    default:
      break;
  }
}

async function handleNatureType(ctx, user, nature) {
  await setState(user.id, 'waiting_classification', { nature });
  await ctx.editMessageText(
    `✅ طبيعة الحركة: *${nature}*\n\nاختر التصنيف:`,
    { parse_mode: 'Markdown', reply_markup: classificationKeyboard(nature) }
  );
}

async function handleClassification(ctx, user, classification) {
  const convState = await getState(user.id);
  const context = { ...convState.context, classification };

  // If direct expense or revenue, ask for project
  if (classification === 'مصروفات مباشرة' || classification === 'ايراد') {
    const projects = await db().project.findMany({
      where: { status: 'نشط' },
      orderBy: { name: 'asc' },
    });
    await setState(user.id, 'waiting_project', context);
    await ctx.editMessageText(
      `✅ التصنيف: *${classification}*\n\nاختر المشروع:`,
      { parse_mode: 'Markdown', reply_markup: projectSelectionKeyboard(projects) }
    );
  } else {
    // General expense - skip project
    await setState(user.id, 'waiting_item', context);
    await ctx.editMessageText(
      `✅ التصنيف: *${classification}*\n\n📝 اكتب اسم البند:`,
      { parse_mode: 'Markdown' }
    );
  }
}

async function handleProjectSelection(ctx, user, projectId) {
  const project = await db().project.findUnique({ where: { id: Number(projectId) } });
  const convState = await getState(user.id);
  const context = {
    ...convState.context,
    projectId: project.id,
    projectName: project.name,
    projectCode: project.code,
  };
  await setState(user.id, 'waiting_item', context);
  await ctx.editMessageText(
    `✅ المشروع: *${project.name}*\n\n📝 اكتب اسم البند:`,
    { parse_mode: 'Markdown' }
  );
}

async function handleCurrency(ctx, user, currency) {
  const convState = await getState(user.id);
  const context = { ...convState.context, currency };

  if (currency === 'USD') {
    context.exchangeRate = 1;
    await setState(user.id, 'waiting_details', context);
    await ctx.editMessageText(
      `✅ العملة: *${CURRENCIES[currency].name}*\n\n📝 اكتب التفاصيل (أو اكتب "تخطي"):`,
      { parse_mode: 'Markdown' }
    );
  } else {
    await setState(user.id, 'waiting_exchange_rate', context);
    await ctx.editMessageText(
      `✅ العملة: *${CURRENCIES[currency].name}*\n\n💱 اكتب سعر الصرف مقابل الدولار:`,
      { parse_mode: 'Markdown' }
    );
  }
}

async function handlePaymentMethod(ctx, user, method) {
  const convState = await getState(user.id);
  const context = { ...convState.context, paymentMethod: method };
  await setState(user.id, 'waiting_payment_term', context);
  await ctx.editMessageText(
    `✅ طريقة الدفع: *${method}*\n\nاختر شرط الدفع:`,
    { parse_mode: 'Markdown', reply_markup: paymentTermKeyboard() }
  );
}

async function handlePaymentTerm(ctx, user, term) {
  const convState = await getState(user.id);
  const context = { ...convState.context, paymentTerm: term };

  if (term === 'فوري') {
    context.dueDate = new Date().toISOString().split('T')[0];
    await setState(user.id, 'waiting_confirmation', context);
    await showConfirmation(ctx, context);
  } else if (term === 'بعد التسليم') {
    await setState(user.id, 'waiting_weeks', context);
    await ctx.editMessageText('📅 كم أسبوع بعد التسليم؟', { parse_mode: 'Markdown' });
  } else if (term === 'تاريخ مخصص') {
    await setState(user.id, 'waiting_custom_date', context);
    await ctx.editMessageText(
      '📅 اكتب تاريخ الاستحقاق بصيغة: YYYY-MM-DD',
      { parse_mode: 'Markdown' }
    );
  }
}

async function handleConfirmation(ctx, user, action) {
  const convState = await getState(user.id);

  if (action === 'yes') {
    // Save to pending transactions
    const context = convState.context;
    const movementType = getMovementType(context.nature);
    const amountUsd = Number(context.amount) * Number(context.exchangeRate || 1);

    await db().pendingTransaction.create({
      data: {
        transactionDate: new Date(context.date || new Date()),
        nature: context.nature,
        classification: context.classification,
        projectId: context.projectId || null,
        itemId: context.itemId || null,
        details: context.details || null,
        partyId: context.partyId || null,
        partyNameRaw: context.partyName || null,
        amount: context.amount,
        currency: context.currency || 'USD',
        exchangeRate: context.exchangeRate || 1,
        amountUsd,
        movementType,
        paymentMethod: context.paymentMethod || null,
        paymentTerm: context.paymentTerm || null,
        paymentTermWeeks: context.weeks || null,
        dueDate: context.dueDate ? new Date(context.dueDate) : null,
        paymentStatus: context.paymentTerm === 'فوري' ? 'PAID' : 'PENDING',
        notes: context.notes || null,
        attachmentUrl: context.attachmentUrl || null,
        unitCount: context.unitCount || null,
        submittedById: user.id,
        source: 'telegram_bot',
        isNewParty: context.isNewParty || false,
      },
    });

    await clearState(user.id);
    await ctx.editMessageText(
      '✅ تم إرسال الحركة بنجاح!\n\n' +
      'ستتم مراجعتها واعتمادها من المسؤول.\n\n' +
      'اختر عملية جديدة:',
    );
    await ctx.reply('القائمة الرئيسية:', { reply_markup: mainMenuKeyboard() });

    // Notify reviewer
    await notifyReviewer(ctx, context, user);
  } else if (action === 'cancel') {
    await clearState(user.id);
    await ctx.editMessageText('🚫 تم إلغاء الحركة.');
    await ctx.reply('اختر عملية جديدة:', { reply_markup: mainMenuKeyboard() });
  } else if (action === 'edit') {
    await ctx.editMessageText('✏️ ما الحقل الذي تريد تعديله؟\n(اكتب اسم الحقل)');
    await setState(user.id, 'waiting_edit_field', convState.context);
  }
}

async function handleReview(ctx, user, action, pendingId) {
  if (!user.permReview) {
    return ctx.answerCallbackQuery('⛔ لا تملك صلاحية المراجعة');
  }

  const id = Number(pendingId);

  if (action === 'approve') {
    // Use the review service
    const pending = await db().pendingTransaction.findUnique({
      where: { id },
      include: { project: true, party: true },
    });
    if (!pending) return ctx.answerCallbackQuery('الحركة غير موجودة');

    const transaction = await transactionService.create({
      transactionDate: pending.transactionDate,
      nature: pending.nature,
      classification: pending.classification,
      projectId: pending.projectId,
      itemId: pending.itemId,
      details: pending.details,
      partyId: pending.partyId,
      amount: Number(pending.amount),
      currency: pending.currency,
      exchangeRate: Number(pending.exchangeRate || 1),
      paymentMethod: pending.paymentMethod,
      paymentTerm: pending.paymentTerm,
      dueDate: pending.dueDate,
      paymentStatus: pending.paymentStatus,
      notes: pending.notes,
      source: pending.source,
      createdById: pending.submittedById,
    });

    await db().pendingTransaction.update({
      where: { id },
      data: {
        reviewStatus: 'APPROVED',
        reviewedById: user.id,
        reviewedAt: new Date(),
        approvedTransactionId: transaction.id,
      },
    });

    await ctx.editMessageText(`✅ تم اعتماد الحركة #${id}\nرقم الحركة في الدفتر: #${transaction.id}`);
  } else if (action === 'reject') {
    await db().pendingTransaction.update({
      where: { id },
      data: {
        reviewStatus: 'REJECTED',
        reviewedById: user.id,
        reviewedAt: new Date(),
      },
    });
    await ctx.editMessageText(`❌ تم رفض الحركة #${id}`);
  }
}

async function handleBack(ctx, user, target) {
  switch (target) {
    case 'main':
      await clearState(user.id);
      await ctx.editMessageText('اختر نوع الحركة المالية:', {
        reply_markup: mainMenuKeyboard(),
      });
      break;
    default:
      break;
  }
}

// ==================== معالجة الرسائل النصية ====================

async function handleTextMessage(ctx) {
  const user = await checkTelegramAuth(ctx.chat.id, 'traditional_bot');
  if (!user) return;

  const convState = await getState(user.id);
  const text = ctx.message.text.trim();

  switch (convState.state) {
    case 'waiting_item': {
      // Fuzzy search for items
      const items = await db().item.findMany();
      const matches = fuzzySearch(text, items, 'name', 0.3);

      let itemId = null;
      let itemName = text;
      if (matches.length > 0 && matches[0].score > 0.7) {
        itemId = matches[0].item.id;
        itemName = matches[0].item.name;
      }

      const context = { ...convState.context, itemId, itemName };
      await setState(user.id, 'waiting_party', context);
      await ctx.reply(`✅ البند: *${itemName}*\n\n👤 اكتب اسم الطرف (المورد/العميل/الممول):`, {
        parse_mode: 'Markdown',
      });
      break;
    }

    case 'waiting_party': {
      const parties = await db().party.findMany({ where: { isActive: true } });
      const matches = fuzzySearch(text, parties, 'name', 0.3);

      let partyId = null;
      let partyName = text;
      let isNewParty = true;
      if (matches.length > 0 && matches[0].score > 0.7) {
        partyId = matches[0].item.id;
        partyName = matches[0].item.name;
        isNewParty = false;
      }

      const context = { ...convState.context, partyId, partyName, isNewParty };
      await setState(user.id, 'waiting_amount', context);
      await ctx.reply(
        `✅ الطرف: *${partyName}*${isNewParty ? ' (جديد)' : ''}\n\n💵 اكتب المبلغ:`,
        { parse_mode: 'Markdown' }
      );
      break;
    }

    case 'waiting_amount': {
      const amount = parseFloat(text.replace(/,/g, ''));
      if (isNaN(amount) || amount <= 0) {
        return ctx.reply('❌ المبلغ غير صالح. أدخل رقماً صحيحاً:');
      }

      const context = { ...convState.context, amount };
      await setState(user.id, 'waiting_currency', context);
      await ctx.reply(`✅ المبلغ: *${amount.toLocaleString()}*\n\nاختر العملة:`, {
        parse_mode: 'Markdown',
        reply_markup: currencyKeyboard(),
      });
      break;
    }

    case 'waiting_exchange_rate': {
      const rate = parseFloat(text);
      if (isNaN(rate) || rate <= 0) {
        return ctx.reply('❌ سعر الصرف غير صالح. أدخل رقماً صحيحاً:');
      }

      const context = { ...convState.context, exchangeRate: rate };
      await setState(user.id, 'waiting_details', context);
      await ctx.reply(`✅ سعر الصرف: *${rate}*\n\n📝 اكتب التفاصيل (أو اكتب "تخطي"):`, {
        parse_mode: 'Markdown',
      });
      break;
    }

    case 'waiting_details': {
      const details = text === 'تخطي' ? null : text;
      const context = { ...convState.context, details, date: new Date().toISOString().split('T')[0] };
      await setState(user.id, 'waiting_payment_method', context);
      await ctx.reply('اختر طريقة الدفع:', { reply_markup: paymentMethodKeyboard() });
      break;
    }

    case 'waiting_weeks': {
      const weeks = parseInt(text);
      if (isNaN(weeks) || weeks <= 0) {
        return ctx.reply('❌ أدخل عدد أسابيع صحيح:');
      }
      const dueDate = new Date();
      dueDate.setDate(dueDate.getDate() + weeks * 7);
      const context = {
        ...convState.context,
        weeks,
        dueDate: dueDate.toISOString().split('T')[0],
      };
      await setState(user.id, 'waiting_confirmation', context);
      await showConfirmation(ctx, context);
      break;
    }

    case 'waiting_custom_date': {
      if (!/^\d{4}-\d{2}-\d{2}$/.test(text)) {
        return ctx.reply('❌ صيغة التاريخ غير صحيحة.\nاستخدم: YYYY-MM-DD');
      }
      const context = { ...convState.context, dueDate: text };
      await setState(user.id, 'waiting_confirmation', context);
      await showConfirmation(ctx, context);
      break;
    }

    default:
      break;
  }
}

// ==================== معالجة المرفقات ====================

async function handleAttachment(ctx) {
  const user = await checkTelegramAuth(ctx.chat.id, 'traditional_bot');
  if (!user) return;

  const convState = await getState(user.id);
  if (convState.state !== 'waiting_confirmation') {
    return ctx.reply('📎 يمكنك إرفاق ملف فقط أثناء تأكيد الحركة.');
  }

  const file = ctx.message.document || ctx.message.photo?.[ctx.message.photo.length - 1];
  if (!file) return;

  const fileInfo = await ctx.api.getFile(file.file_id);
  const fileUrl = `https://api.telegram.org/file/bot${ctx.api.token}/${fileInfo.file_path}`;

  const context = { ...convState.context, attachmentUrl: fileUrl };
  await setState(user.id, 'waiting_confirmation', context);
  await ctx.reply('✅ تم إرفاق الملف.\n\nهل تريد تأكيد الحركة؟', {
    reply_markup: confirmationKeyboard(),
  });
}

// ==================== عرض التأكيد ====================

async function showConfirmation(ctx, context) {
  const amountUsd = Number(context.amount) * Number(context.exchangeRate || 1);
  const currencyInfo = CURRENCIES[context.currency] || CURRENCIES.USD;

  let message = `📋 *ملخص الحركة:*\n\n`;
  message += `• الطبيعة: ${context.nature}\n`;
  message += `• التصنيف: ${context.classification}\n`;
  if (context.projectName) message += `• المشروع: ${context.projectName}\n`;
  if (context.itemName) message += `• البند: ${context.itemName}\n`;
  message += `• الطرف: ${context.partyName}${context.isNewParty ? ' (جديد)' : ''}\n`;
  message += `• المبلغ: ${Number(context.amount).toLocaleString()} ${currencyInfo.symbol}\n`;
  if (context.currency !== 'USD') {
    message += `• سعر الصرف: ${context.exchangeRate}\n`;
    message += `• القيمة بالدولار: $${amountUsd.toLocaleString()}\n`;
  }
  if (context.details) message += `• التفاصيل: ${context.details}\n`;
  if (context.paymentMethod) message += `• طريقة الدفع: ${context.paymentMethod}\n`;
  if (context.paymentTerm) message += `• شرط الدفع: ${context.paymentTerm}\n`;
  if (context.dueDate) message += `• تاريخ الاستحقاق: ${context.dueDate}\n`;
  if (context.attachmentUrl) message += `\n📎 مرفق: نعم\n`;

  message += `\n*هل تريد تأكيد الحركة؟*`;

  await ctx.reply(message, {
    parse_mode: 'Markdown',
    reply_markup: confirmationKeyboard(),
  });
}

// ==================== إشعار المراجع ====================

async function notifyReviewer(ctx, context, submitter) {
  const { env } = await import('../../config/env.js');
  if (!env.accountantChatId) return;

  const amountUsd = Number(context.amount) * Number(context.exchangeRate || 1);
  const pendingCount = await db().pendingTransaction.count({
    where: { reviewStatus: 'PENDING' },
  });

  // Get the latest pending transaction ID
  const latest = await db().pendingTransaction.findFirst({
    where: { submittedById: submitter.id },
    orderBy: { submittedAt: 'desc' },
  });

  let message = `🔔 *حركة جديدة تنتظر المراجعة*\n\n`;
  message += `• المُدخل: ${submitter.name}\n`;
  message += `• الطبيعة: ${context.nature}\n`;
  if (context.projectName) message += `• المشروع: ${context.projectName}\n`;
  message += `• الطرف: ${context.partyName}\n`;
  message += `• المبلغ: $${amountUsd.toLocaleString()}\n`;
  message += `\n📊 إجمالي المعلقة: ${pendingCount}`;

  try {
    await ctx.api.sendMessage(env.accountantChatId, message, {
      parse_mode: 'Markdown',
      reply_markup: latest ? reviewKeyboard(latest.id) : undefined,
    });
  } catch (err) {
    logger.error('Failed to notify reviewer:', err.message);
  }
}
