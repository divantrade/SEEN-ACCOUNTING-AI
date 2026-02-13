/**
 * =============================================
 * Migration Script: Google Sheets → PostgreSQL
 * =============================================
 *
 * الخطوات:
 * 1. صدّر كل شيت من Google Sheets كـ CSV
 * 2. ضع ملفات CSV في مجلد data/
 * 3. شغل: node src/database/migrate-from-sheets.js
 *
 * الملفات المطلوبة:
 * - data/projects.csv       (قاعدة بيانات المشاريع)
 * - data/parties.csv        (قاعدة بيانات الأطراف)
 * - data/items.csv          (قاعدة بيانات البنود)
 * - data/transactions.csv   (دفتر الحركات المالية)
 * - data/users.csv          (المستخدمين المصرح لهم)
 */

import { parse } from 'csv-parse/sync';
import { readFileSync, existsSync } from 'fs';
import { PrismaClient } from '@prisma/client';
import 'dotenv/config';

const prisma = new PrismaClient();

// ==================== CSV Reader ====================

function readCSV(filename) {
  const filepath = `data/${filename}`;
  if (!existsSync(filepath)) {
    console.warn(`⚠️  File not found: ${filepath} - skipping`);
    return [];
  }

  const content = readFileSync(filepath, 'utf-8');
  return parse(content, {
    columns: true,
    skip_empty_lines: true,
    trim: true,
    bom: true,
  });
}

// ==================== Type Mappers ====================

const NATURE_MAP = {
  'استحقاق مصروف': 'EXPENSE_ACCRUAL',
  'دفعة مصروف': 'EXPENSE_PAYMENT',
  'استحقاق إيراد': 'REVENUE_ACCRUAL',
  'تحصيل إيراد': 'REVENUE_COLLECTION',
  'تمويل': 'FINANCING',
  'استلام تمويل': 'FINANCING_RECEIVED',
  'سداد تمويل': 'FINANCING_REPAYMENT',
  'تأمين مدفوع للقناة': 'INSURANCE_PAID',
  'استرداد تأمين من القناة': 'INSURANCE_RETURNED',
  'تحويل داخلي': 'INTERNAL_TRANSFER',
};

const CLASSIFICATION_MAP = {
  'مصروفات مباشرة': 'DIRECT_EXPENSE',
  'مصروفات عمومية': 'GENERAL_EXPENSE',
  'ايراد': 'REVENUE',
  'تمويل طويل الأجل': 'LONG_TERM_FINANCING',
  'سداد تمويل طويل الأجل': 'LONG_TERM_REPAYMENT',
  'سلفة قصيرة': 'SHORT_ADVANCE',
  'سداد سلفة قصيرة': 'SHORT_ADVANCE_REPAY',
  'تأمين': 'INSURANCE',
  'استرداد تأمين': 'INSURANCE_RETURN',
  'تحويل للبنك': 'TRANSFER_TO_BANK',
  'تحويل للخزنة': 'TRANSFER_TO_CASH',
};

const PARTY_TYPE_MAP = {
  'مورد': 'VENDOR',
  'عميل': 'CLIENT',
  'ممول': 'FUNDER',
};

const MOVEMENT_MAP = {
  'مدين استحقاق': 'DEBIT',
  'دائن دفعة': 'CREDIT',
};

const PAYMENT_METHOD_MAP = {
  'نقدي': 'CASH',
  'تحويل بنكي': 'BANK_TRANSFER',
  'شيك': 'CHECK',
  'بطاقة': 'CARD',
  'أخرى': 'OTHER',
};

const PAYMENT_STATUS_MAP = {
  'مدفوع بالكامل': 'PAID',
  'معلق': 'PENDING',
  'عملية دفع/تحصيل': 'IN_PROCESS',
};

// ==================== Date Parser ====================

function parseDate(dateStr) {
  if (!dateStr) return null;

  // Try dd/mm/yyyy
  const parts = dateStr.split('/');
  if (parts.length === 3) {
    const [day, month, year] = parts;
    return new Date(`${year}-${month.padStart(2, '0')}-${day.padStart(2, '0')}`);
  }

  // Try yyyy-mm-dd
  const date = new Date(dateStr);
  return isNaN(date.getTime()) ? null : date;
}

function parseNumber(str) {
  if (!str) return null;
  const num = parseFloat(str.replace(/,/g, ''));
  return isNaN(num) ? null : num;
}

// ==================== Migration Functions ====================

async function migrateProjects() {
  const rows = readCSV('projects.csv');
  if (rows.length === 0) return;

  console.log(`📁 Migrating ${rows.length} projects...`);

  for (const row of rows) {
    const code = row['كود المشروع'] || row['code'];
    const name = row['اسم المشروع'] || row['name'];
    if (!code || !name) continue;

    await prisma.project.upsert({
      where: { code },
      create: { code, name },
      update: { name },
    });
  }

  console.log(`✅ Projects migrated`);
}

async function migrateParties() {
  const rows = readCSV('parties.csv');
  if (rows.length === 0) return;

  console.log(`👥 Migrating ${rows.length} parties...`);

  for (const row of rows) {
    const name = row['اسم الطرف'] || row['name'];
    const typeAr = row['نوع الطرف'] || row['type'];
    if (!name || !typeAr) continue;

    const type = PARTY_TYPE_MAP[typeAr];
    if (!type) {
      console.warn(`  ⚠️  Unknown party type: ${typeAr} for ${name}`);
      continue;
    }

    await prisma.party.upsert({
      where: { name_type: { name, type } },
      create: {
        name,
        type,
        specialization: row['التخصص / الفئة'] || null,
        phone: row['رقم الهاتف'] || null,
        email: row['البريد الإلكتروني'] || null,
        city: row['المدينة / الدولة'] || null,
        preferredPayment: row['طريقة الدفع المفضلة'] || null,
        bankDetails: row['بيانات الحساب البنكي'] || null,
        notes: row['ملاحظات'] || null,
      },
      update: {},
    });
  }

  console.log(`✅ Parties migrated`);
}

async function migrateItems() {
  const rows = readCSV('items.csv');
  if (rows.length === 0) return;

  console.log(`📋 Migrating ${rows.length} items...`);

  for (const row of rows) {
    const name = row['البند'] || row['name'];
    if (!name) continue;

    const natureAr = row['طبيعة الحركة'] || null;
    const classAr = row['تصنيف الحركة'] || null;

    await prisma.item.upsert({
      where: { name },
      create: {
        name,
        nature: natureAr ? NATURE_MAP[natureAr] : null,
        classification: classAr ? CLASSIFICATION_MAP[classAr] : null,
      },
      update: {},
    });
  }

  console.log(`✅ Items migrated`);
}

async function migrateUsers() {
  const rows = readCSV('users.csv');
  if (rows.length === 0) return;

  console.log(`👤 Migrating ${rows.length} users...`);

  for (const row of rows) {
    const name = row['الاسم'] || row['name'];
    if (!name) continue;

    const chatId = row['معرّف المحادثة'] || row['chat_id'];

    await prisma.user.create({
      data: {
        name,
        email: row['الإيميل'] || null,
        phone: row['رقم الهاتف'] || null,
        telegramUsername: row['اسم المستخدم تليجرام'] || null,
        telegramChatId: chatId ? BigInt(chatId) : null,
        permTraditionalBot: row['بوت تقليدي'] === 'TRUE',
        permAiBot: row['بوت ذكي'] === 'TRUE',
        permSheet: row['شيت'] === 'TRUE',
        permReview: row['مراجعة'] === 'TRUE',
        isActive: row['نشط'] !== 'FALSE',
      },
    });
  }

  console.log(`✅ Users migrated`);
}

async function migrateTransactions() {
  const rows = readCSV('transactions.csv');
  if (rows.length === 0) return;

  console.log(`💰 Migrating ${rows.length} transactions...`);

  // Pre-load lookups
  const projects = await prisma.project.findMany();
  const projectMap = new Map(projects.map((p) => [p.code, p.id]));
  const projectNameMap = new Map(projects.map((p) => [p.name, p.id]));

  const parties = await prisma.party.findMany();
  const partyMap = new Map(parties.map((p) => [p.name, p.id]));

  const items = await prisma.item.findMany();
  const itemMap = new Map(items.map((i) => [i.name, i.id]));

  let migrated = 0;
  let skipped = 0;

  for (const row of rows) {
    try {
      const dateStr = row['التاريخ'] || row['date'];
      const date = parseDate(dateStr);
      if (!date) {
        skipped++;
        continue;
      }

      const natureAr = row['طبيعة الحركة'] || '';
      const classAr = row['تصنيف الحركة'] || '';
      const nature = NATURE_MAP[natureAr];
      const classification = CLASSIFICATION_MAP[classAr];

      if (!nature || !classification) {
        console.warn(`  ⚠️  Unknown nature/classification: ${natureAr}/${classAr}`);
        skipped++;
        continue;
      }

      const projectCode = row['كود المشروع'] || '';
      const projectName = row['اسم المشروع'] || '';
      const projectId = projectMap.get(projectCode) || projectNameMap.get(projectName) || null;

      const partyName = row['الطرف'] || row['اسم المورد/الجهة'] || '';
      const partyId = partyMap.get(partyName) || null;

      const itemName = row['البند'] || '';
      const itemId = itemMap.get(itemName) || null;

      const amount = parseNumber(row['المبلغ بالعملة الأصلية']) || 0;
      const currency = row['العملة'] || 'USD';
      const exchangeRate = parseNumber(row['سعر الصرف']) || 1;
      const amountUsd = parseNumber(row['القيمة بالدولار']) || amount * exchangeRate;

      const movementAr = row['نوع الحركة'] || '';
      const movementType = MOVEMENT_MAP[movementAr] || 'DEBIT';

      const paymentMethodAr = row['طريقة الدفع'] || '';
      const paymentMethod = PAYMENT_METHOD_MAP[paymentMethodAr] || null;

      const paymentStatusAr = row['حالة السداد'] || '';
      const paymentStatus = PAYMENT_STATUS_MAP[paymentStatusAr] || 'PENDING';

      await prisma.transaction.create({
        data: {
          transactionDate: date,
          nature,
          classification,
          projectId,
          itemId,
          details: row['التفاصيل'] || null,
          partyId,
          amount,
          currency,
          exchangeRate,
          amountUsd,
          movementType,
          referenceNumber: row['رقم مرجعي'] || null,
          paymentMethod,
          paymentStatus,
          dueDate: parseDate(row['تاريخ الاستحقاق']),
          month: row['الشهر'] || date.toISOString().slice(0, 7),
          notes: row['ملاحظات'] || null,
          source: 'migrated',
        },
      });

      migrated++;
    } catch (err) {
      console.error(`  ❌ Error migrating row:`, err.message);
      skipped++;
    }
  }

  console.log(`✅ Transactions migrated: ${migrated} success, ${skipped} skipped`);
}

// ==================== Main ====================

async function main() {
  console.log('🚀 Starting migration from Google Sheets to PostgreSQL...\n');

  try {
    // Order matters: dependencies first
    await migrateProjects();
    await migrateParties();
    await migrateItems();
    await migrateUsers();
    await migrateTransactions();

    console.log('\n🎉 Migration completed successfully!');
  } catch (err) {
    console.error('\n❌ Migration failed:', err);
    process.exit(1);
  } finally {
    await prisma.$disconnect();
  }
}

main();
