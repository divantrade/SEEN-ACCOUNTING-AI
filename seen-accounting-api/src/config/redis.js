import Redis from 'ioredis';
import { env } from './env.js';
import { logger } from '../shared/logger.js';

let redis;

export function getRedis() {
  if (!redis) {
    redis = new Redis(env.redisUrl, {
      maxRetriesPerRequest: 3,
      retryStrategy(times) {
        const delay = Math.min(times * 200, 5000);
        return delay;
      },
    });

    redis.on('connect', () => logger.info('Redis connected'));
    redis.on('error', (err) => logger.error('Redis error:', err.message));
  }
  return redis;
}

export async function disconnectRedis() {
  if (redis) {
    await redis.quit();
    redis = null;
  }
}
