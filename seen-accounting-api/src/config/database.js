import { PrismaClient } from '@prisma/client';
import { env } from './env.js';

// Singleton Prisma client
let prisma;

export function getDB() {
  if (!prisma) {
    prisma = new PrismaClient({
      log: env.isDev ? ['query', 'error', 'warn'] : ['error'],
    });
  }
  return prisma;
}

export async function connectDB() {
  const db = getDB();
  await db.$connect();
  return db;
}

export async function disconnectDB() {
  if (prisma) {
    await prisma.$disconnect();
    prisma = null;
  }
}
