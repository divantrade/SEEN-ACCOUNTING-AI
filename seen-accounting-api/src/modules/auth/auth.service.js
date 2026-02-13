import { getDB } from '../../config/database.js';
import { getRedis } from '../../config/redis.js';

const AUTH_CACHE_TTL = 3600;       // 1 hour for authorized users
const UNAUTH_CACHE_TTL = 300;      // 5 minutes for unauthorized

const db = () => getDB();

/**
 * Check Telegram user authorization
 * Used by both traditional and AI bots
 */
export async function checkTelegramAuth(chatId, permissionType) {
  const redis = getRedis();
  const cacheKey = `auth:${chatId}:${permissionType}`;

  // Check cache first
  const cached = await redis.get(cacheKey);
  if (cached) return JSON.parse(cached);

  // Find user by chat ID
  let user = await db().user.findFirst({
    where: {
      telegramChatId: BigInt(chatId),
      isActive: true,
    },
  });

  if (!user) {
    // Cache negative result
    await redis.setex(cacheKey, UNAUTH_CACHE_TTL, JSON.stringify(null));
    return null;
  }

  // Check permission
  const permMap = {
    traditional_bot: 'permTraditionalBot',
    ai_bot: 'permAiBot',
    review: 'permReview',
  };

  const field = permMap[permissionType];
  if (!field || !user[field]) {
    await redis.setex(cacheKey, UNAUTH_CACHE_TTL, JSON.stringify(null));
    return null;
  }

  // Serialize user (BigInt → string for JSON)
  const serialized = {
    id: user.id,
    name: user.name,
    telegramChatId: user.telegramChatId.toString(),
    permReview: user.permReview,
  };

  // Cache positive result
  await redis.setex(cacheKey, AUTH_CACHE_TTL, JSON.stringify(serialized));
  return serialized;
}

/**
 * Register or update Telegram chat ID for a user
 */
export async function linkTelegramChat(chatId, username) {
  // Try to find by username
  const user = await db().user.findFirst({
    where: {
      telegramUsername: { equals: username.replace('@', ''), mode: 'insensitive' },
      isActive: true,
    },
  });

  if (user && !user.telegramChatId) {
    await db().user.update({
      where: { id: user.id },
      data: { telegramChatId: BigInt(chatId) },
    });

    // Clear auth cache
    const redis = getRedis();
    const keys = await redis.keys(`auth:${chatId}:*`);
    if (keys.length) await redis.del(...keys);

    return user;
  }

  return user;
}
