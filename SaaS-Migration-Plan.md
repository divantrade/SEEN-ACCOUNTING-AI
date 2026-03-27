# خطة تحويل SEEN Accounting إلى KodLab SaaS

## الرؤية

تحويل نظام SEEN المحاسبي الشخصي (Google Sheets + Telegram Bot) إلى منتج SaaS تجاري
باسم **KodLab** يخدم عملاء متعددين في أي صناعة - وليس الإنتاج الفني فقط.

---

## الفرق بين النظام الحالي والمنتج الجديد

| الحالي (SEEN - شخصي) | الجديد (KodLab - SaaS) |
|---|---|
| تليجرام بوت فقط | واجهة ويب كاملة + تليجرام بوت |
| Google Sheets | PostgreSQL (Supabase) |
| مستخدم واحد / فريق صغير | عملاء متعددين (Multi-tenant) |
| مجاني | باشتراك شهري |
| بدون login | Login + صلاحيات |
| للإنتاج الفني فقط | لأي صناعة |
| غير قابل للبيع | منتج تجاري |

---

## الوضع الحالي

```
Google Sheets = قاعدة البيانات + الواجهة + السيرفر (كل شيء في مكان واحد)
Apps Script   = المنطق والتقارير والبوت
```

### نقاط القوة الحالية (تُنقل للنظام الجديد)
- محرك ذكاء اصطناعي يعمل بـ Gemini Function Calling (SmartAgent)
- 12 تقرير PDF جاهز
- بحث عربي ذكي بـ 5 مستويات مطابقة (FuzzySearch)
- نظام مراجعة واعتماد حركات (ReviewSystem)
- تعلم تفضيلات المستخدم
- كشف الحركات المكررة
- إدخال ذكي (Guided Data Entry) - البوت يسأل عن الناقص فقط
- استنتاج تلقائي من تخصص الطرف (مونتير → بند مونتاج)
- تاريخ مشاريع الطرف (شريف عمل في فيلم X وY)

### القيود التي تدفع للتحول
- بطء مع البيانات الكبيرة (Google Sheets)
- مستخدم واحد / شركة واحدة
- لا يمكن تخصيص الواجهة
- حدود Apps Script (6 دقائق تشغيل، 100MB بيانات)
- لا يمكن بيعه كمنتج

---

## البنية المستهدفة

```
┌──────────────────────────────────────────────────────┐
│                    الواجهة (Frontend)                  │
│                      Next.js + React                  │
│                                                       │
│   🔐 تسجيل دخول / تسجيل    📊 لوحة تحكم             │
│   💰 إدخال حركات            📄 تقارير PDF             │
│   👥 إدارة أطراف            🎬 إدارة مشاريع          │
│   ⚙️ إعدادات الشركة         💳 إدارة الاشتراك         │
│   🤖 ربط بوت تليجرام       📱 متجاوب مع الموبايل     │
└───────────────────────────────┬──────────────────────┘
                                │ API
┌───────────────────────────────┴──────────────────────┐
│                    السيرفر (Backend)                   │
│                   Next.js API Routes                  │
│                                                       │
│   🔑 Auth (Supabase)         📊 Reports Engine        │
│   🏢 Multi-tenancy           🤖 Telegram Bot Manager  │
│   💰 Transactions API        🧠 Gemini AI / Claude    │
│   📄 PDF Generator           💳 Billing (Stripe)      │
│   🔍 Arabic Fuzzy Search     📁 File Storage          │
└───────────────────────────────┬──────────────────────┘
                                │
┌───────────────────────────────┴──────────────────────┐
│                  قاعدة البيانات (PostgreSQL)           │
│                        Supabase                       │
│                                                       │
│   شركة A ──→ حركاتها، أطرافها، مشاريعها              │
│   شركة B ──→ حركاتها، أطرافها، مشاريعها              │
│   شركة C ──→ حركاتها، أطرافها، مشاريعها              │
│   (كل شركة معزولة تماماً عن الأخرى via RLS)          │
└──────────────────────────────────────────────────────┘
```

---

## الأدوات والتقنيات

| المكون | الأداة | السبب | التكلفة |
|---|---|---|---|
| الواجهة | Next.js + React + Tailwind CSS (RTL) | سريعة، تدعم العربية RTL، SEO | مجاني |
| السيرفر | Next.js API Routes | نفس المشروع = أبسط في الإدارة | مجاني |
| قاعدة البيانات | Supabase (PostgreSQL) + Prisma ORM | Auth جاهز، RLS، مجاني للبداية، Prisma لـ type safety | $0 → $25/شهر |
| الاستضافة | Vercel | deploy تلقائي من Git، مجاني للبداية | $0 → $20/شهر |
| الدفع والاشتراكات | Stripe | اشتراكات شهرية/سنوية، فواتير تلقائية | 2.9% لكل معاملة |
| توليد PDF | `@react-pdf/renderer` | تقارير احترافية بدون سيرفر منفصل | مجاني |
| بوت تليجرام | Telegram Webhook via API Route | لا يحتاج سيرفر منفصل | مجاني |
| الذكاء الاصطناعي | Gemini / Claude API | نفس المنطق الحالي مع تحسينات | حسب الاستخدام |
| تخزين الملفات | Supabase Storage | بدل Google Drive | ضمن خطة Supabase |
| UI Components | shadcn/ui | مكونات جاهزة وقابلة للتخصيص | مجاني |
| الـ ORM | Prisma | Type-safe queries، migrations تلقائية، schema واضح | مجاني |

### التكلفة المتوقعة
| المرحلة | التكلفة الشهرية |
|---|---|
| البداية (أول 10 شركات) | $0 (كل شيء مجاني) |
| النمو (10-100 شركة) | ~$45/شهر |
| التوسع (100+ شركة) | ~$200/شهر |

---

## ربط بوت تليجرام (استراتيجية Multi-tenant)

> مدمج من فرع `excel-to-web-app` (خطة KodLab)

### النموذج: بوت واحد لكل العملاء (@KodLabBot)

```
العميل يسجل في الويب → يدخل إعدادات → يضغط "ربط تليجرام"
    ↓
الموقع يولّد كود تفعيل فريد (مثل: KL-A7X9)
    ↓
العميل يرسل الكود للبوت @KodLabBot في تليجرام
    ↓
البوت يربط chat_id بـ company_id → جاهز للاستخدام
```

- **بوت واحد** لكل العملاء (مش بوت لكل شركة)
- العميل يربط حسابه من **صفحة الإعدادات** بالويب عبر كود تفعيل
- البوت يعرف الشركة من `chat_id` → `user.company_id`
- صلاحيات البوت قابلة للتحكم من الويب (permAiBot, permTraditionalBot)
- أكثر من مستخدم في نفس الشركة يستخدمون البوت

### الميزة الرئيسية: الإدخال الذكي (Guided Data Entry)

البوت **مش بيعلّم محاسبة** - البوت بيتأكد إن البيانات **كاملة** قبل التسجيل:

1. المستخدم يكتب رسالة (مثال: "دفعت لشريف 5000 دولار")
2. الـ AI يفهم ويبحث عن "شريف" تلقائياً → يجده مونتير
3. يستنتج: مونتير → بند "مونتاج" → تصنيف "مصروفات مباشرة"
4. يبحث عن مشاريع شريف السابقة → يقترح المشروع
5. لو فيه ناقص → يسأل عن الناقص بس
6. يعرض ملخص ويطلب تأكيد → يسجل

الحقول المطلوبة **قابلة للتخصيص** من إعدادات الويب لكل نوع حركة ولكل صناعة.

---

## تصميم قاعدة البيانات (Prisma Schema)

> مدمج من أفضل ما في فرع `review-project-plan` (12 موديل Prisma مفصل) + فرعنا الحالي (Multi-tenant)

### Enums (الأنواع المحددة)

```prisma
enum PartyType {
  VENDOR  @map("مورد")
  CLIENT  @map("عميل")
  FUNDER  @map("ممول")
}

enum NatureType {
  EXPENSE_ACCRUAL       @map("استحقاق مصروف")
  EXPENSE_PAYMENT       @map("دفعة مصروف")
  REVENUE_ACCRUAL       @map("استحقاق إيراد")
  REVENUE_COLLECTION    @map("تحصيل إيراد")
  FINANCING_RECEIVED    @map("استلام تمويل")
  FINANCING_REPAYMENT   @map("سداد تمويل")
  INSURANCE_PAID        @map("تأمين مدفوع للقناة")
  INSURANCE_RETURNED    @map("استرداد تأمين من القناة")
  INTERNAL_TRANSFER     @map("تحويل داخلي")
}

enum ClassificationType {
  DIRECT_EXPENSE        @map("مصروفات مباشرة")
  GENERAL_EXPENSE       @map("مصروفات عمومية")
  REVENUE               @map("ايراد")
  LONG_TERM_FINANCING   @map("تمويل طويل الأجل")
  LONG_TERM_REPAYMENT   @map("سداد تمويل طويل الأجل")
  SHORT_ADVANCE         @map("سلفة قصيرة")
  SHORT_ADVANCE_REPAY   @map("سداد سلفة قصيرة")
  INSURANCE             @map("تأمين")
  INSURANCE_RETURN      @map("استرداد تأمين")
  TRANSFER_TO_BANK      @map("تحويل للبنك")
  TRANSFER_TO_CASH      @map("تحويل للخزنة")
}

enum MovementType     { DEBIT @map("مدين استحقاق")  CREDIT @map("دائن دفعة") }
enum PaymentMethod    { CASH  BANK_TRANSFER  CHECK  CARD  OTHER }
enum PaymentTerm      { IMMEDIATE  AFTER_DELIVERY  CUSTOM_DATE }
enum PaymentStatus    { PAID  PENDING  IN_PROCESS }
enum ReviewStatus     { PENDING  APPROVED  REJECTED  NEEDS_EDIT }
enum CurrencyCode     { USD  TRY  EGP }
```

### الجداول (13 موديل)

```prisma
// ===================== الشركات (Multi-tenant) =====================
model Company {
  id                String    @id @default(uuid())
  name              String
  email             String?
  timezone          String    @default("Asia/Istanbul")
  defaultCurrency   CurrencyCode @default(USD)
  currencies        CurrencyCode[]
  subscriptionPlan  String    @default("free")
  stripeCustomerId  String?
  settings          Json      @default("{}")
  createdAt         DateTime  @default(now())

  users             User[]
  projects          Project[]
  parties           Party[]
  items             Item[]
  transactions      Transaction[]
  pendingTransactions PendingTransaction[]
  pendingParties    PendingParty[]
  exchangeRates     ExchangeRate[]
  budgets           Budget[]
  activityLogs      ActivityLog[]

  @@map("companies")
}

// ===================== المستخدمون =====================
model User {
  id                Int       @id @default(autoincrement())
  companyId         String    @map("company_id")
  name              String
  email             String?
  phone             String?
  telegramUsername  String?   @map("telegram_username")
  telegramChatId   BigInt?   @unique @map("telegram_chat_id")
  role              String    @default("accountant") // admin, accountant, viewer
  permTraditionalBot Boolean @default(false)
  permAiBot         Boolean  @default(false)
  permSheet         Boolean  @default(false)
  permReview        Boolean  @default(false)
  permAdmin         Boolean  @default(false)
  isActive          Boolean  @default(true)
  passwordHash      String?
  lastLogin         DateTime?
  createdAt         DateTime @default(now())
  updatedAt         DateTime @updatedAt

  company             Company @relation(fields: [companyId], references: [id])
  createdTransactions Transaction[] @relation("TransactionCreator")
  submittedPending    PendingTransaction[] @relation("PendingSubmitter")
  reviewedPending     PendingTransaction[] @relation("PendingReviewer")
  submittedParties    PendingParty[] @relation("PartySubmitter")
  reviewedParties     PendingParty[] @relation("PartyReviewer")
  activityLogs        ActivityLog[]
  conversationState   ConversationState?
  rejectedTransactions RejectedTransaction[]

  @@index([companyId])
  @@index([telegramChatId])
  @@map("users")
}

// ===================== المشاريع =====================
model Project {
  id          Int       @id @default(autoincrement())
  companyId   String    @map("company_id")
  code        String    @db.VarChar(50)
  name        String
  status      String    @default("نشط")
  description String?
  createdAt   DateTime  @default(now())
  updatedAt   DateTime  @updatedAt

  company      Company @relation(fields: [companyId], references: [id])
  transactions Transaction[]
  pendingTransactions PendingTransaction[]
  budgets      Budget[]

  @@unique([companyId, code])
  @@index([companyId])
  @@map("projects")
}

// ===================== الأطراف =====================
model Party {
  id               Int       @id @default(autoincrement())
  companyId        String    @map("company_id")
  name             String
  type             PartyType
  specialization   String?   // تخصص الطرف (مونتير، مخرج، مصور...)
  phone            String?
  email            String?
  city             String?
  preferredPayment String?
  bankDetails      String?
  notes            String?
  isActive         Boolean   @default(true)
  createdAt        DateTime  @default(now())
  updatedAt        DateTime  @updatedAt

  company      Company @relation(fields: [companyId], references: [id])
  transactions Transaction[]
  pendingTransactions PendingTransaction[]
  pendingParties PendingParty[]

  @@unique([companyId, name, type])
  @@index([companyId])
  @@index([name])
  @@map("parties")
}

// ===================== البنود المحاسبية =====================
model Item {
  id             Int                 @id @default(autoincrement())
  companyId      String              @map("company_id")
  name           String
  nature         NatureType?
  classification ClassificationType?
  unitType       String?
  createdAt      DateTime            @default(now())

  company      Company @relation(fields: [companyId], references: [id])
  transactions Transaction[]
  pendingTransactions PendingTransaction[]

  @@unique([companyId, name])
  @@index([companyId])
  @@map("items")
}

// ===================== الحركات المالية (القلب) =====================
model Transaction {
  id               Int                @id @default(autoincrement())
  companyId        String             @map("company_id")
  transactionDate  DateTime           @db.Date
  nature           NatureType
  classification   ClassificationType
  projectId        Int?               @map("project_id")
  itemId           Int?               @map("item_id")
  partyId          Int?               @map("party_id")
  details          String?
  amount           Decimal            @db.Decimal(14, 2)
  currency         CurrencyCode       @default(USD)
  exchangeRate     Decimal            @default(1.0) @db.Decimal(10, 6)
  amountUsd        Decimal?           @db.Decimal(14, 2)
  movementType     MovementType
  referenceNumber  String?
  paymentMethod    PaymentMethod?
  paymentTerm      PaymentTerm?       @default(IMMEDIATE)
  paymentTermWeeks Int?
  paymentTermDate  DateTime?          @db.Date
  dueDate          DateTime?          @db.Date
  paymentStatus    PaymentStatus      @default(PENDING)
  month            String?            // "2026-03" للفلترة السريعة
  notes            String?
  attachmentUrl    String?
  unitCount        Int?
  source           String             @default("manual") // manual, telegram_bot, web
  createdById      Int?               @map("created_by")
  createdAt        DateTime           @default(now())
  updatedAt        DateTime           @updatedAt

  company   Company  @relation(fields: [companyId], references: [id])
  project   Project? @relation(fields: [projectId], references: [id])
  item      Item?    @relation(fields: [itemId], references: [id])
  party     Party?   @relation(fields: [partyId], references: [id])
  createdBy User?    @relation("TransactionCreator", fields: [createdById], references: [id])
  pendingTransaction PendingTransaction?

  @@index([companyId])
  @@index([transactionDate])
  @@index([projectId])
  @@index([partyId])
  @@index([month])
  @@index([paymentStatus])
  @@index([dueDate])
  @@map("transactions")
}

// ===================== حركات البوت المعلقة =====================
model PendingTransaction {
  id               Int                @id @default(autoincrement())
  companyId        String             @map("company_id")
  // ... نفس حقول Transaction (كلها optional) ...
  partyNameRaw     String?            // الاسم كما كتبه المستخدم قبل المطابقة
  isNewParty       Boolean            @default(false)
  reviewStatus     ReviewStatus       @default(PENDING)
  submittedById    Int?
  submittedAt      DateTime           @default(now())
  reviewedById     Int?
  reviewedAt       DateTime?
  reviewNotes      String?
  attachmentUrl    String?
  source           String             @default("telegram_bot")
  approvedTransactionId Int?          @unique

  company             Company     @relation(fields: [companyId], references: [id])
  submittedBy          User?      @relation("PendingSubmitter", fields: [submittedById], references: [id])
  reviewedBy           User?      @relation("PendingReviewer", fields: [reviewedById], references: [id])
  approvedTransaction  Transaction? @relation(fields: [approvedTransactionId], references: [id])

  @@index([companyId])
  @@index([reviewStatus])
  @@map("pending_transactions")
}

// ===================== أطراف معلقة للمراجعة =====================
model PendingParty {
  id               Int          @id @default(autoincrement())
  companyId        String       @map("company_id")
  name             String
  type             PartyType
  specialization   String?
  phone            String?
  reviewStatus     ReviewStatus @default(PENDING)
  submittedById    Int?
  reviewedById     Int?
  reviewNotes      String?
  approvedPartyId  Int?

  submittedBy   User?  @relation("PartySubmitter", fields: [submittedById], references: [id])
  reviewedBy    User?  @relation("PartyReviewer", fields: [reviewedById], references: [id])
  approvedParty Party? @relation(fields: [approvedPartyId], references: [id])

  @@map("pending_parties")
}

// ===================== حالات المحادثة (البوت) =====================
model ConversationState {
  userId    Int      @id @map("user_id")
  botType   String   @default("ai") // ai, traditional
  state     String   @default("idle")
  context   Json     @default("{}")
  updatedAt DateTime @updatedAt

  user User @relation(fields: [userId], references: [id])
  @@map("conversation_states")
}

// ===================== الموازنات =====================
model Budget {
  id        Int          @id @default(autoincrement())
  companyId String       @map("company_id")
  projectId Int
  itemName  String?
  amount    Decimal      @db.Decimal(14, 2)
  currency  CurrencyCode @default(USD)
  notes     String?
  createdAt DateTime     @default(now())

  company Company @relation(fields: [companyId], references: [id])
  project Project @relation(fields: [projectId], references: [id])
  @@index([companyId, projectId])
  @@map("budgets")
}

// ===================== أسعار الصرف =====================
model ExchangeRate {
  id        Int          @id @default(autoincrement())
  companyId String       @map("company_id")
  currency  CurrencyCode
  rate      Decimal      @db.Decimal(10, 6)
  validDate DateTime     @db.Date

  company Company @relation(fields: [companyId], references: [id])
  @@unique([companyId, currency, validDate])
  @@map("exchange_rates")
}

// ===================== سجل النشاط =====================
model ActivityLog {
  id         Int      @id @default(autoincrement())
  companyId  String   @map("company_id")
  userId     Int?
  action     String   // "create_transaction", "approve_pending", "reject_party"
  entityType String?  // "transaction", "party", "project"
  entityId   Int?
  details    Json?
  ipAddress  String?
  createdAt  DateTime @default(now())

  company Company @relation(fields: [companyId], references: [id])
  user    User?   @relation(fields: [userId], references: [id])
  @@index([companyId, createdAt])
  @@map("activity_log")
}

// ===================== أرشيف المرفوضات =====================
model RejectedTransaction {
  id                   Int      @id @default(autoincrement())
  pendingTransactionId Int?
  data                 Json     // نسخة كاملة من البيانات المرفوضة
  rejectedById         Int?
  rejectedAt           DateTime @default(now())
  rejectionReason      String?

  rejectedBy User? @relation(fields: [rejectedById], references: [id])
  @@map("rejected_transactions")
}
```

### Row Level Security (عزل البيانات)

```sql
-- كل شركة ترى بياناتها فقط
ALTER TABLE transactions ENABLE ROW LEVEL SECURITY;

CREATE POLICY "company_isolation" ON transactions
    FOR ALL
    USING (company_id = (SELECT company_id FROM users WHERE id = auth.uid()));

-- نفس السياسة لكل الجداول التي تحتوي company_id
-- هذا يمنع أي شركة من الوصول لبيانات شركة أخرى
```

---

## خطة التنفيذ

### المرحلة 1: الأساس (3-4 أسابيع)

```
الأسبوع 1: البنية التحتية
├── إنشاء مشروع Next.js جديد
├── إعداد Supabase (قاعدة بيانات + Auth)
├── إنشاء الجداول الأساسية
├── إعداد Row Level Security
└── نظام تسجيل الدخول والصلاحيات

الأسبوع 2-3: الحركات المالية
├── API: إنشاء/تعديل/حذف/عرض الحركات
├── واجهة إدخال الحركات (form)
├── جدول عرض الحركات (مع فلاتر وبحث)
├── نظام المراجعة والاعتماد
└── إدارة المشاريع والأطراف والبنود

الأسبوع 4: البحث والبنية
├── تحويل FuzzySearch إلى TypeScript
├── Full-text search بالعربية في PostgreSQL
├── API لإدارة المشاريع والأطراف
└── اختبارات أساسية
```

### المرحلة 2: التقارير (2-3 أسابيع)

```
الأسبوع 5-6: محرك التقارير
├── تحويل منطق التقارير من Main.js
├── توليد PDF (كشوفات حساب، ربحية، مصروفات...)
├── تقرير تحليل المصروفات الديناميكي
├── القوائم المالية (دخل، مركز مالي، ميزان مراجعة)
└── التدفقات النقدية

الأسبوع 7: لوحة التحكم
├── Dashboard بالرسوم البيانية
├── ملخص الأرصدة والاستحقاقات
├── أداء المشاريع
└── تنبيهات المدفوعات القادمة
```

### المرحلة 3: البوت الذكي (2 أسابيع)

```
الأسبوع 8: ربط تليجرام
├── Telegram Webhook endpoint
├── تحويل SmartAgent + AgentTools
├── ربط Gemini/Claude API
└── تعلم تفضيلات المستخدم

الأسبوع 9: إدارة البوتات
├── كل شركة تربط البوت الخاص بها
├── لوحة إعداد البوت من الواجهة
├── إشعارات المراجعة عبر البوت
└── إرسال التقارير PDF عبر البوت
```

### المرحلة 4: ميزات SaaS (2-3 أسابيع)

```
الأسبوع 10: الاشتراكات
├── خطط الأسعار (مجاني / أساسي / احترافي)
├── ربط Stripe للدفع
├── إدارة الاشتراك من الواجهة
└── حدود كل خطة (عدد حركات، مستخدمين، مشاريع)

الأسبوع 11: تجربة المستخدم
├── صفحة هبوط تسويقية
├── Onboarding (إعداد ذاتي للشركة الجديدة)
├── استيراد بيانات من Excel/CSV
└── دليل استخدام تفاعلي

الأسبوع 12: الإطلاق
├── اختبارات شاملة
├── نقل بيانات SEEN كأول عميل
├── مراقبة الأداء
└── الإطلاق التجريبي (Beta)
```

---

## خطط الاشتراك المقترحة

| الميزة | مجاني | أساسي ($29/شهر) | احترافي ($79/شهر) |
|---|---|---|---|
| عدد الحركات/شهر | 100 | 1,000 | غير محدود |
| المستخدمين | 1 | 3 | 10 |
| المشاريع | 3 | 15 | غير محدود |
| التقارير | 5 أساسية | 12 تقرير | 12 + تقارير مخصصة |
| بوت تليجرام | لا | نعم | نعم |
| الذكاء الاصطناعي | لا | أساسي | كامل + تعلم |
| تصدير PDF | لا | نعم | نعم |
| دعم فني | مجتمع | بريد إلكتروني | أولوية |
| نسخ احتياطي | يومي | يومي | كل ساعة |

---

## تحويل الكود الحالي

### ما يُنقل مباشرة (المنطق)
| الملف الحالي | المقابل في SaaS | ملاحظات |
|---|---|---|
| `SmartAgent.js` | `lib/ai/smart-agent.ts` | تحويل لـ TypeScript، نفس Function Calling |
| `AgentTools.js` | `lib/ai/agent-tools.ts` | Prisma queries بدل Sheet getValues() |
| `FuzzySearch.js` | `lib/search/fuzzy-search.ts` | + PostgreSQL `pg_trgm` full-text search |
| `ReportsConfig.js` | `lib/reports/config.ts` | نفس البنية |
| `AIConfig.js` (الثوابت) | `lib/ai/config.ts` | الأنواع والتصنيفات وقواعد الاستنتاج |

### ما يُعاد بناؤه (يعتمد على Sheets)
| الملف الحالي | المقابل في SaaS | ملاحظات |
|---|---|---|
| `Main.js` (التقارير) | `lib/reports/*.ts` | SQL بدل getValues() |
| `BotSheets.js` | `lib/db/transactions.ts` | Prisma بدل Sheet API |
| `PDFGenerator.js` | `lib/pdf/generator.ts` | @react-pdf بدل Sheets PDF export |
| `ReviewSystem.js` | `lib/review/system.ts` | نفس المنطق مع PendingTransaction model |
| `AIBot.js` (state machine) | `lib/bot/state-machine.ts` | ConversationState model بدل PropertiesService |

### ما يُحذف (خاص بـ Google)
| الملف | السبب |
|---|---|
| `DriveAttachments.js` | يُستبدل بـ Supabase Storage |
| `appsscript.json` | خاص بـ Apps Script |
| `.clasp.json` | خاص بـ clasp |

### بنية المجلدات المقترحة للـ SaaS

```
kodlab/
├── prisma/
│   ├── schema.prisma          ← السكيما المفصلة أعلاه
│   ├── seed.ts                ← بيانات تجريبية
│   └── migrations/
├── src/
│   ├── app/                   ← Next.js App Router
│   │   ├── (auth)/            ← تسجيل دخول / تسجيل
│   │   ├── (dashboard)/       ← لوحة التحكم
│   │   ├── api/
│   │   │   ├── webhook/telegram/  ← Telegram Bot webhook
│   │   │   ├── transactions/
│   │   │   ├── reports/
│   │   │   ├── parties/
│   │   │   └── projects/
│   ├── lib/
│   │   ├── ai/
│   │   │   ├── smart-agent.ts     ← من SmartAgent.js
│   │   │   ├── agent-tools.ts     ← من AgentTools.js
│   │   │   ├── inference-rules.ts ← قواعد الاستنتاج التلقائي
│   │   │   └── config.ts
│   │   ├── bot/
│   │   │   ├── state-machine.ts   ← من AIBot.js
│   │   │   ├── handlers.ts        ← من TelegramBot.js
│   │   │   └── keyboards.ts
│   │   ├── db/
│   │   │   ├── prisma.ts          ← Prisma client singleton
│   │   │   └── transactions.ts    ← من BotSheets.js
│   │   ├── reports/
│   │   │   ├── generator.ts       ← من Main.js
│   │   │   └── pdf.ts             ← من PDFGenerator.js
│   │   ├── search/
│   │   │   └── fuzzy-search.ts    ← من FuzzySearch.js
│   │   └── review/
│   │       └── system.ts          ← من ReviewSystem.js
│   └── scripts/
│       └── migrate-from-sheets.ts ← نقل بيانات SEEN
└── package.json
```

---

## المخاطر والتحديات

| المخاطر | الحل |
|---|---|
| فقدان بيانات أثناء النقل | تشغيل النظامين بالتوازي حتى التأكد |
| تكلفة Gemini API مع عدة شركات | حدود استخدام لكل خطة + caching |
| أمان البيانات المالية | Row Level Security + تشفير + HTTPS |
| أداء التقارير الكبيرة | تقارير مُولدة مسبقاً (cached) + pagination |
| دعم RTL العربية في الواجهة | Tailwind RTL + مكتبات عربية |

---

## قرارات معمارية مهمة

> مدمجة من قرارات الفروع الثلاثة

1. **Supabase + Vercel وليس VPS** - أرخص، أبسط، auto-scaling، لا حاجة لإدارة سيرفر
2. **Prisma ORM وليس SQL مباشر** - type safety، migrations، schema واضح ومقروء
3. **بوت واحد لكل العملاء** (@KodLabBot) - مش بوت لكل شركة
4. **نجرب على SEEN Bot أولاً** - ثم التبديل لـ @KodLabBot آخر خطوة (تغيير BOT_TOKEN)
5. **قاعدة بيانات واحدة + RLS** - مش نسخة لكل عميل (كابوس في الصيانة)
6. **الإدخال الذكي مش تعليم محاسبة** - البوت يسأل عن الناقص ويستنتج الباقي

---

## مصادر هذه الخطة

هذا الملف يدمج أفضل ما في 3 مصادر:

| المصدر | ما أخذنا منه |
|---|---|
| `claude/excel-to-web-app-lhzfF` (KodLab Plan) | رؤية المنتج، ربط البوت، Guided Entry، قرارات معمارية |
| `claude/review-project-plan-GEt1L` (VPS Plan) | Prisma Schema المفصل (12 model + enums)، بنية الكود |
| `claude/analyze-test-coverage-Pjqg5` (هذا الفرع) | Next.js + Supabase stack، خطة التنفيذ، التسعير، تحويل الكود |

---

## الخطوة التالية

1. إنشاء repo جديد `kodlab` وتجهيز مشروع Next.js
2. إعداد Supabase وتشغيل `prisma migrate` بالسكيما أعلاه
3. بناء Auth (تسجيل دخول + إنشاء شركة)
4. بناء API الحركات المالية (CRUD + review)
5. ربط Telegram webhook + نقل SmartAgent
6. نقل بيانات SEEN كأول عميل تجريبي

> آخر تحديث: 2026-03-27
> الحالة: مرحلة التخطيط - مكتمل ومدمج من 3 فروع
