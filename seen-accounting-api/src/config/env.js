import 'dotenv/config';

export const env = {
  // Server
  nodeEnv: process.env.NODE_ENV || 'development',
  port: parseInt(process.env.PORT || '3000', 10),
  host: process.env.HOST || '0.0.0.0',
  isDev: process.env.NODE_ENV !== 'production',

  // Database
  databaseUrl: process.env.DATABASE_URL,

  // Redis
  redisUrl: process.env.REDIS_URL || 'redis://localhost:6379',

  // JWT
  jwtSecret: process.env.JWT_SECRET || 'dev-secret-change-in-production',
  jwtExpiresIn: process.env.JWT_EXPIRES_IN || '7d',

  // Telegram
  telegramBotToken: process.env.TELEGRAM_BOT_TOKEN,
  telegramAiBotToken: process.env.TELEGRAM_AI_BOT_TOKEN,
  accountantChatId: process.env.ACCOUNTANT_CHAT_ID,

  // Gemini
  geminiApiKey: process.env.GEMINI_API_KEY,
  geminiModel: process.env.GEMINI_MODEL || 'gemini-2.0-flash',

  // File Storage
  uploadDir: process.env.UPLOAD_DIR || './uploads',
  maxFileSize: parseInt(process.env.MAX_FILE_SIZE || '10485760', 10),
};
