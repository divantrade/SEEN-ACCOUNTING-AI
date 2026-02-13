import { Bot } from 'grammy';
import { env } from '../../config/env.js';
import { logger } from '../../shared/logger.js';
import { registerCommands } from './handlers.js';

let bot;

export async function startTraditionalBot() {
  if (!env.telegramBotToken) {
    logger.warn('TELEGRAM_BOT_TOKEN not set, skipping traditional bot');
    return;
  }

  bot = new Bot(env.telegramBotToken);

  // Register all handlers
  registerCommands(bot);

  // Error handler
  bot.catch((err) => {
    logger.error('Traditional bot error:', err.message);
  });

  // Set bot commands menu
  await bot.api.setMyCommands([
    { command: 'start', description: 'بدء محادثة جديدة' },
    { command: 'expense', description: 'تسجيل مصروف' },
    { command: 'revenue', description: 'تسجيل إيراد' },
    { command: 'status', description: 'حالة الحركات' },
    { command: 'cancel', description: 'إلغاء العملية' },
    { command: 'help', description: 'المساعدة' },
  ]);

  // Start bot with long polling
  bot.start({
    onStart: () => logger.info('Traditional Telegram bot is running'),
    allowed_updates: ['message', 'callback_query'],
  });

  return bot;
}

export function getTraditionalBot() {
  return bot;
}

export async function stopTraditionalBot() {
  if (bot) {
    await bot.stop();
    bot = null;
  }
}
