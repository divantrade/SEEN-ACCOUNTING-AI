import Fastify from 'fastify';
import cors from '@fastify/cors';
import helmet from '@fastify/helmet';
import rateLimit from '@fastify/rate-limit';
import jwt from '@fastify/jwt';
import { env } from './config/env.js';
import { connectDB, disconnectDB } from './config/database.js';
import { getRedis, disconnectRedis } from './config/redis.js';
import { logger } from './shared/logger.js';
import { AppError } from './shared/errors.js';

// Route imports
import { transactionRoutes } from './modules/transactions/transactions.routes.js';
import { projectRoutes } from './modules/projects/projects.routes.js';
import { partyRoutes } from './modules/parties/parties.routes.js';
import { itemRoutes } from './modules/items/items.routes.js';
import { authRoutes } from './modules/auth/auth.routes.js';
import { reviewRoutes } from './modules/review/review.routes.js';
import { reportRoutes } from './modules/reports/reports.routes.js';

// Bot imports
import { startTraditionalBot } from './bots/telegram/traditional-bot.js';
import { startAIBot } from './bots/ai/ai-bot.js';

async function buildApp() {
  const app = Fastify({
    logger: env.isDev,
    trustProxy: true,
  });

  // ==================== Plugins ====================

  await app.register(cors, {
    origin: env.isDev ? true : [/* production domains */],
    credentials: true,
  });

  await app.register(helmet);

  await app.register(rateLimit, {
    max: 100,
    timeWindow: '1 minute',
  });

  await app.register(jwt, {
    secret: env.jwtSecret,
    sign: { expiresIn: env.jwtExpiresIn },
  });

  // ==================== Decorators ====================

  // Auth decorator - authenticate JWT
  app.decorate('authenticate', async function (request, reply) {
    try {
      await request.jwtVerify();
    } catch (err) {
      reply.status(401).send({ error: 'غير مصرح', code: 'UNAUTHORIZED' });
    }
  });

  // ==================== Error Handler ====================

  app.setErrorHandler((error, request, reply) => {
    logger.error({
      message: error.message,
      stack: error.stack,
      url: request.url,
      method: request.method,
    });

    if (error instanceof AppError) {
      return reply.status(error.statusCode).send({
        error: error.message,
        code: error.code,
      });
    }

    // Prisma errors
    if (error.code === 'P2002') {
      return reply.status(409).send({
        error: 'هذا العنصر موجود بالفعل',
        code: 'DUPLICATE',
      });
    }
    if (error.code === 'P2025') {
      return reply.status(404).send({
        error: 'العنصر غير موجود',
        code: 'NOT_FOUND',
      });
    }

    // Validation errors (Fastify/Zod)
    if (error.validation) {
      return reply.status(400).send({
        error: 'بيانات غير صالحة',
        code: 'VALIDATION_ERROR',
        details: error.validation,
      });
    }

    return reply.status(500).send({
      error: 'خطأ داخلي في الخادم',
      code: 'INTERNAL_ERROR',
    });
  });

  // ==================== Routes ====================

  // Health check
  app.get('/health', async () => ({
    status: 'ok',
    timestamp: new Date().toISOString(),
    uptime: process.uptime(),
  }));

  // API routes
  app.register(authRoutes, { prefix: '/api/auth' });
  app.register(transactionRoutes, { prefix: '/api/transactions' });
  app.register(projectRoutes, { prefix: '/api/projects' });
  app.register(partyRoutes, { prefix: '/api/parties' });
  app.register(itemRoutes, { prefix: '/api/items' });
  app.register(reviewRoutes, { prefix: '/api/review' });
  app.register(reportRoutes, { prefix: '/api/reports' });

  return app;
}

// ==================== Start Server ====================

async function start() {
  try {
    // Connect to database
    const db = await connectDB();
    logger.info('PostgreSQL connected');

    // Connect to Redis
    getRedis();

    // Build and start Fastify
    const app = await buildApp();
    await app.listen({ port: env.port, host: env.host });
    logger.info(`Server running on http://${env.host}:${env.port}`);

    // Start Telegram bots
    if (env.telegramBotToken) {
      await startTraditionalBot();
      logger.info('Traditional Telegram bot started');
    }
    if (env.telegramAiBotToken) {
      await startAIBot();
      logger.info('AI Telegram bot started');
    }

    // Graceful shutdown
    const shutdown = async (signal) => {
      logger.info(`${signal} received, shutting down...`);
      await app.close();
      await disconnectDB();
      await disconnectRedis();
      process.exit(0);
    };

    process.on('SIGTERM', () => shutdown('SIGTERM'));
    process.on('SIGINT', () => shutdown('SIGINT'));
  } catch (err) {
    logger.error('Failed to start server:', err);
    process.exit(1);
  }
}

start();
