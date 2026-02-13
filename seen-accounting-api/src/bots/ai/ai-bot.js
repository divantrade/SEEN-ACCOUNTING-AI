import { Bot } from 'grammy';
import { env } from '../../config/env.js';
import { getDB } from '../../config/database.js';
import { logger } from '../../shared/logger.js';
import { checkTelegramAuth } from '../../modules/auth/auth.service.js';
import { getMovementType } from '../../config/constants.js';
import { fuzzySearch } from '../../shared/fuzzy-search.js';
import { callGemini, parseGeminiJSON } from './gemini-client.js';
import { buildSystemPrompt } from './inference-rules.js';
import { confirmationKeyboard, reviewKeyboard } from '../telegram/keyboards.js';

let bot;
const db = () => getDB();

export async function startAIBot() {
  if (!env.telegramAiBotToken) {
    logger.warn('TELEGRAM_AI_BOT_TOKEN not set, skipping AI bot');
    return;
  }

  bot = new Bot(env.telegramAiBotToken);

  // Commands
  bot.command('start', handleAIStart);
  bot.command('help', handleAIHelp);
  bot.command('cancel', handleAICancel);

  // Callback queries
  bot.on('callback_query:data', handleAICallback);

  // Text messages - main AI handler
  bot.on('message:text', handleAIMessage);

  // Attachments
  bot.on('message:document', handleAIAttachment);
  bot.on('message:photo', handleAIAttachment);

  // Error handler
  bot.catch((err) => {
    logger.error('AI bot error:', err.message);
  });

  // Set commands
  await bot.api.setMyCommands([
    { command: 'start', description: 'بدء محادثة جديدة' },
    { command: 'help', description: 'المساعدة' },
    { command: 'cancel', description: 'إلغاء العملية' },
  ]);

  bot.start({
    onStart: () => logger.info('AI Telegram bot is running'),
    allowed_updates: ['message', 'callback_query'],
  });

  return bot;
}

// ==================== الأوامر ====================

async function handleAIStart(ctx) {
  const user = await checkTelegramAuth(ctx.chat.id, 'ai_bot');
  if (!user) {
    return ctx.reply('⛔ غير مصرح لك باستخدام البوت الذكي.');
  }

  await clearAIState(user.id);
  await ctx.reply(
    `مرحباً ${user.name} 🤖\n\n` +
    `أنا المساعد الذكي لنظام SEEN المحاسبي.\n\n` +
    `اكتب الحركة المالية بلغة طبيعية وأنا هأفهمها:\n\n` +
    `*أمثلة:*\n` +
    `• "اتفقت مع أحمد على تصوير 3 مقابلات لمشروع كلينتون بـ 500 دولار"\n` +
    `• "دفعت لمحمد 200 ليرة مونتاج لمشروع القاهرة"\n` +
    `• "حولت 1000 دولار لشركة النور إيجار المكتب"\n\n` +
    `ابدأ بكتابة الحركة 👇`,
    { parse_mode: 'Markdown' }
  );
}

async function handleAIHelp(ctx) {
  await ctx.reply(
    `🤖 *البوت الذكي - المساعدة*\n\n` +
    `اكتب الحركة المالية بلغة عادية:\n\n` +
    `✅ "اتفقت مع فلان على كذا بمبلغ كذا"\n` +
    `✅ "دفعت لفلان مبلغ كذا"\n` +
    `✅ "استلمت من العميل مبلغ كذا"\n\n` +
    `البوت سيفهم:\n` +
    `• نوع الحركة (مصروف/إيراد/تمويل)\n` +
    `• المشروع والبند\n` +
    `• الطرف والمبلغ والعملة\n` +
    `• طريقة وشرط الدفع\n\n` +
    `/cancel - إلغاء العملية الحالية`,
    { parse_mode: 'Markdown' }
  );
}

async function handleAICancel(ctx) {
  const user = await checkTelegramAuth(ctx.chat.id, 'ai_bot');
  if (!user) return;

  await clearAIState(user.id);
  await ctx.reply('🚫 تم الإلغاء.\n\nاكتب حركة جديدة في أي وقت.');
}

// ==================== المعالجة الرئيسية ====================

async function handleAIMessage(ctx) {
  const user = await checkTelegramAuth(ctx.chat.id, 'ai_bot');
  if (!user) return ctx.reply('⛔ غير مصرح لك.');

  const text = ctx.message.text.trim();
  const convState = await getAIState(user.id);

  // If waiting for missing field
  if (convState.state === 'ai_waiting_missing_field') {
    return handleMissingField(ctx, user, text, convState);
  }

  // If waiting for confirmation edit
  if (convState.state === 'ai_waiting_edit') {
    return handleAIEdit(ctx, user, text, convState);
  }

  // Main AI processing
  await ctx.reply('🔄 جاري تحليل الحركة...');

  try {
    // Load context data
    const [projects, items, parties] = await Promise.all([
      db().project.findMany({ where: { status: 'نشط' } }),
      db().item.findMany(),
      db().party.findMany({ where: { isActive: true } }),
    ]);

    // Build prompt and call Gemini
    const systemPrompt = buildSystemPrompt(projects, items, parties);
    const response = await callGemini(systemPrompt, text);
    const parsed = parseGeminiJSON(response);

    if (!parsed.success) {
      return ctx.reply('❌ لم أستطع فهم الحركة.\nحاول مرة أخرى بصيغة أوضح.');
    }

    // Resolve IDs from names
    const resolvedData = await resolveNames(parsed, projects, items, parties);

    // Check for missing required fields
    if (parsed.missing_fields?.length > 0) {
      await setAIState(user.id, 'ai_waiting_missing_field', {
        ...resolvedData,
        missingFields: parsed.missing_fields,
        currentMissing: 0,
      });

      const field = parsed.missing_fields[0];
      return ctx.reply(`📝 يرجى تحديد: *${getFieldLabel(field)}*`, {
        parse_mode: 'Markdown',
      });
    }

    // Show confirmation
    await setAIState(user.id, 'ai_waiting_confirmation', resolvedData);
    await showAIConfirmation(ctx, resolvedData);
  } catch (err) {
    logger.error('AI processing error:', err);
    await ctx.reply('❌ حدث خطأ في المعالجة.\nحاول مرة أخرى.');
  }
}

// ==================== معالجة الحقول الناقصة ====================

async function handleMissingField(ctx, user, text, convState) {
  const context = convState.context;
  const missingFields = context.missingFields;
  const currentIdx = context.currentMissing;
  const field = missingFields[currentIdx];

  // Set the missing field value
  context[field] = text;

  // Check if more fields needed
  if (currentIdx + 1 < missingFields.length) {
    context.currentMissing = currentIdx + 1;
    const nextField = missingFields[currentIdx + 1];
    await setAIState(user.id, 'ai_waiting_missing_field', context);
    return ctx.reply(`📝 يرجى تحديد: *${getFieldLabel(nextField)}*`, {
      parse_mode: 'Markdown',
    });
  }

  // All fields collected - show confirmation
  await setAIState(user.id, 'ai_waiting_confirmation', context);
  await showAIConfirmation(ctx, context);
}

// ==================== معالجة الأزرار ====================

async function handleAICallback(ctx) {
  const user = await checkTelegramAuth(ctx.chat.id, 'ai_bot');
  if (!user) return ctx.answerCallbackQuery('⛔ غير مصرح');

  const data = ctx.callbackQuery.data;
  const [action, value, extra] = data.split(':');

  try {
    if (action === 'confirm') {
      const convState = await getAIState(user.id);

      if (value === 'yes') {
        await saveAITransaction(ctx, user, convState.context);
      } else if (value === 'cancel') {
        await clearAIState(user.id);
        await ctx.editMessageText('🚫 تم إلغاء الحركة.\n\nاكتب حركة جديدة في أي وقت.');
      } else if (value === 'edit') {
        await setAIState(user.id, 'ai_waiting_edit', convState.context);
        await ctx.editMessageText(
          '✏️ اكتب التعديل المطلوب:\n' +
          '(مثلاً: "المبلغ 500" أو "المشروع كلينتون")'
        );
      }
    }
    await ctx.answerCallbackQuery();
  } catch (err) {
    logger.error('AI callback error:', err);
    await ctx.answerCallbackQuery('❌ حدث خطأ');
  }
}

// ==================== حفظ الحركة ====================

async function saveAITransaction(ctx, user, context) {
  try {
    const movementType = getMovementType(context.nature);
    const amount = Number(context.amount);
    const exchangeRate = Number(context.exchange_rate || context.exchangeRate || 1);
    const amountUsd = amount * exchangeRate;

    await db().pendingTransaction.create({
      data: {
        transactionDate: new Date(),
        nature: context.nature,
        classification: context.classification,
        projectId: context.projectId || null,
        itemId: context.itemId || null,
        details: context.details || null,
        partyId: context.partyId || null,
        partyNameRaw: context.party || null,
        amount,
        currency: context.currency || 'USD',
        exchangeRate,
        amountUsd,
        movementType,
        paymentMethod: context.payment_method || context.paymentMethod || null,
        paymentTerm: context.payment_term || context.paymentTerm || null,
        dueDate: context.due_date ? new Date(context.due_date) : null,
        paymentStatus: 'PENDING',
        notes: context.details || null,
        unitCount: context.unit_count || context.unitCount || null,
        attachmentUrl: context.attachmentUrl || null,
        submittedById: user.id,
        source: 'ai_bot',
        isNewParty: !context.partyId,
      },
    });

    await clearAIState(user.id);
    await ctx.editMessageText(
      '✅ تم إرسال الحركة للمراجعة!\n\nاكتب حركة جديدة في أي وقت.'
    );

    // Notify reviewer
    if (env.accountantChatId) {
      const latest = await db().pendingTransaction.findFirst({
        where: { submittedById: user.id },
        orderBy: { submittedAt: 'desc' },
      });

      await ctx.api.sendMessage(
        env.accountantChatId,
        `🤖 *حركة من البوت الذكي*\n\n` +
        `• المُدخل: ${user.name}\n` +
        `• الطبيعة: ${context.nature}\n` +
        `• المبلغ: $${amountUsd.toLocaleString()}\n` +
        `• الطرف: ${context.party || 'غير محدد'}`,
        {
          parse_mode: 'Markdown',
          reply_markup: latest ? reviewKeyboard(latest.id) : undefined,
        }
      ).catch((err) => logger.error('Failed to notify reviewer:', err.message));
    }
  } catch (err) {
    logger.error('Save AI transaction error:', err);
    await ctx.reply('❌ فشل حفظ الحركة. حاول مرة أخرى.');
  }
}

// ==================== وظائف مساعدة ====================

async function resolveNames(parsed, projects, items, parties) {
  const resolved = { ...parsed };

  // Resolve project
  if (parsed.project) {
    const projMatches = fuzzySearch(parsed.project, projects, 'name', 0.4);
    if (projMatches.length > 0) {
      resolved.projectId = projMatches[0].item.id;
      resolved.projectCode = projMatches[0].item.code;
      resolved.projectName = projMatches[0].item.name;
    }
  }

  // Resolve item
  if (parsed.item) {
    const itemMatches = fuzzySearch(parsed.item, items, 'name', 0.4);
    if (itemMatches.length > 0) {
      resolved.itemId = itemMatches[0].item.id;
      resolved.itemName = itemMatches[0].item.name;
    }
  }

  // Resolve party
  if (parsed.party) {
    const partyMatches = fuzzySearch(parsed.party, parties, 'name', 0.4);
    if (partyMatches.length > 0) {
      resolved.partyId = partyMatches[0].item.id;
      resolved.partyName = partyMatches[0].item.name;
    }
  }

  return resolved;
}

async function showAIConfirmation(ctx, context) {
  const amount = Number(context.amount || 0);
  const exchangeRate = Number(context.exchange_rate || context.exchangeRate || 1);
  const amountUsd = amount * exchangeRate;
  const currency = context.currency || 'USD';

  let msg = `🤖 *تم تحليل الحركة:*\n\n`;
  msg += `• الطبيعة: ${context.nature || '—'}\n`;
  msg += `• التصنيف: ${context.classification || '—'}\n`;
  if (context.projectName || context.project) {
    msg += `• المشروع: ${context.projectName || context.project}\n`;
  }
  if (context.itemName || context.item) {
    msg += `• البند: ${context.itemName || context.item}\n`;
  }
  msg += `• الطرف: ${context.partyName || context.party || '—'}${!context.partyId ? ' (جديد)' : ''}\n`;
  msg += `• المبلغ: ${amount.toLocaleString()} ${currency}\n`;
  if (currency !== 'USD' && exchangeRate !== 1) {
    msg += `• سعر الصرف: ${exchangeRate}\n`;
    msg += `• بالدولار: $${amountUsd.toLocaleString()}\n`;
  }
  if (context.details) msg += `• التفاصيل: ${context.details}\n`;
  if (context.payment_method) msg += `• طريقة الدفع: ${context.payment_method}\n`;
  if (context.unit_count) msg += `• عدد الوحدات: ${context.unit_count}\n`;
  if (context.confidence) msg += `\n🎯 الثقة: ${Math.round(context.confidence * 100)}%\n`;

  msg += `\n*هل البيانات صحيحة؟*`;

  await ctx.reply(msg, {
    parse_mode: 'Markdown',
    reply_markup: confirmationKeyboard(),
  });
}

async function handleAIEdit(ctx, user, text, convState) {
  // Re-process with Gemini to understand the edit
  try {
    const [projects, items, parties] = await Promise.all([
      db().project.findMany({ where: { status: 'نشط' } }),
      db().item.findMany(),
      db().party.findMany({ where: { isActive: true } }),
    ]);

    const editPrompt = `البيانات الحالية: ${JSON.stringify(convState.context)}\n\nالتعديل المطلوب: ${text}\n\nأعد JSON محدث بالتعديل فقط.`;
    const systemPrompt = buildSystemPrompt(projects, items, parties);
    const response = await callGemini(systemPrompt, editPrompt);
    const parsed = parseGeminiJSON(response);

    const updatedContext = { ...convState.context, ...parsed };
    const resolved = await resolveNames(updatedContext, projects, items, parties);

    await setAIState(user.id, 'ai_waiting_confirmation', resolved);
    await showAIConfirmation(ctx, resolved);
  } catch (err) {
    logger.error('AI edit error:', err);
    await ctx.reply('❌ لم أفهم التعديل. حاول مرة أخرى.');
  }
}

async function handleAIAttachment(ctx) {
  const user = await checkTelegramAuth(ctx.chat.id, 'ai_bot');
  if (!user) return;

  const convState = await getAIState(user.id);
  if (convState.state !== 'ai_waiting_confirmation') {
    return ctx.reply('📎 أرسل المرفق بعد كتابة الحركة.');
  }

  const file = ctx.message.document || ctx.message.photo?.[ctx.message.photo.length - 1];
  if (!file) return;

  const fileInfo = await ctx.api.getFile(file.file_id);
  const fileUrl = `https://api.telegram.org/file/bot${ctx.api.token}/${fileInfo.file_path}`;

  const context = { ...convState.context, attachmentUrl: fileUrl };
  await setAIState(user.id, 'ai_waiting_confirmation', context);
  await ctx.reply('✅ تم إرفاق الملف.', { reply_markup: confirmationKeyboard() });
}

// ==================== State Management ====================

async function getAIState(userId) {
  const state = await db().conversationState.findUnique({ where: { userId } });
  return state || { state: 'ai_idle', context: {} };
}

async function setAIState(userId, state, context = {}) {
  await db().conversationState.upsert({
    where: { userId },
    create: { userId, botType: 'ai', state, context },
    update: { state, context },
  });
}

async function clearAIState(userId) {
  await db().conversationState.deleteMany({ where: { userId } });
}

function getFieldLabel(field) {
  const labels = {
    amount: 'المبلغ',
    party: 'اسم الطرف',
    project: 'المشروع',
    item: 'البند',
    currency: 'العملة',
    payment_method: 'طريقة الدفع',
    details: 'التفاصيل',
  };
  return labels[field] || field;
}

export function getAIBotInstance() {
  return bot;
}

export async function stopAIBot() {
  if (bot) {
    await bot.stop();
    bot = null;
  }
}
