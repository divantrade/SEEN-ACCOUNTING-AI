import { PrismaClient } from '@prisma/client';
import bcrypt from 'bcryptjs';

const prisma = new PrismaClient();

async function main() {
  console.log('🌱 Seeding database...\n');

  // Create admin user
  const passwordHash = await bcrypt.hash('admin123', 12);
  const admin = await prisma.user.upsert({
    where: { telegramChatId: 0n },
    create: {
      name: 'مدير النظام',
      email: 'admin@seen.com',
      passwordHash,
      permAdmin: true,
      permReview: true,
      permTraditionalBot: true,
      permAiBot: true,
      permSheet: true,
      telegramChatId: 0n,
    },
    update: {},
  });
  console.log(`✅ Admin user created: ${admin.name}`);

  // Create sample items
  const items = [
    { name: 'تصوير', nature: 'EXPENSE_ACCRUAL', classification: 'DIRECT_EXPENSE', unitType: 'مقابلة' },
    { name: 'مونتاج', nature: 'EXPENSE_ACCRUAL', classification: 'DIRECT_EXPENSE', unitType: 'دقيقة' },
    { name: 'مكساج', nature: 'EXPENSE_ACCRUAL', classification: 'DIRECT_EXPENSE', unitType: 'دقيقة' },
    { name: 'دوبلاج', nature: 'EXPENSE_ACCRUAL', classification: 'DIRECT_EXPENSE', unitType: 'دقيقة' },
    { name: 'تلوين', nature: 'EXPENSE_ACCRUAL', classification: 'DIRECT_EXPENSE', unitType: 'دقيقة' },
    { name: 'جرافيك - رسم', nature: 'EXPENSE_ACCRUAL', classification: 'DIRECT_EXPENSE', unitType: 'رسمة' },
    { name: 'فيكسر', nature: 'EXPENSE_ACCRUAL', classification: 'DIRECT_EXPENSE', unitType: 'ضيف' },
    { name: 'تعليق صوتي', nature: 'EXPENSE_ACCRUAL', classification: 'DIRECT_EXPENSE', unitType: 'دقيقة' },
    { name: 'ترجمة', nature: 'EXPENSE_ACCRUAL', classification: 'DIRECT_EXPENSE', unitType: null },
    { name: 'إيجار مكتب', nature: 'EXPENSE_ACCRUAL', classification: 'GENERAL_EXPENSE', unitType: null },
    { name: 'رواتب', nature: 'EXPENSE_ACCRUAL', classification: 'GENERAL_EXPENSE', unitType: null },
    { name: 'سفر وإقامة', nature: 'EXPENSE_ACCRUAL', classification: 'GENERAL_EXPENSE', unitType: null },
  ];

  for (const item of items) {
    await prisma.item.upsert({
      where: { name: item.name },
      create: item,
      update: {},
    });
  }
  console.log(`✅ ${items.length} items seeded`);

  console.log('\n🎉 Seeding completed!');
  console.log('\n📋 Default admin login:');
  console.log('   Email: admin@seen.com');
  console.log('   Password: admin123');
  console.log('   ⚠️  Change this password immediately!\n');
}

main()
  .catch((err) => {
    console.error('Seeding failed:', err);
    process.exit(1);
  })
  .finally(() => prisma.$disconnect());
