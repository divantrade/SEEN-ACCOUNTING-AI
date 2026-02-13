# خطة الانتقال إلى VPS + PostgreSQL
# SEEN Accounting AI - Migration Plan

> **الهدف:** بناء بنية تحتية متينة واحترافية تدعم النمو المستقبلي
> **التاريخ:** 2026-02-13
> **الحالة:** مرحلة التخطيط

---

## 1. لماذا VPS + PostgreSQL؟

### المشاكل الحالية مع Google Sheets
| المشكلة | التأثير |
|---|---|
| حد 10 مليون خلية | سقف نمو ثابت |
| بطء مع البيانات الكبيرة | تجربة مستخدم سيئة |
| لا يوجد Foreign Keys حقيقية | بيانات قد تتناقض |
| لا يوجد Transactions (ACID) | احتمال فقدان بيانات |
| الصلاحيات محدودة | أمان ضعيف على مستوى الصف |
| Google Apps Script - 6 دقائق حد أقصى | عمليات محدودة |
| لا يوجد API حقيقي | صعوبة التوسع |

### ماذا يقدم VPS + PostgreSQL
| الميزة | الفائدة |
|---|---|
| قاعدة بيانات لا محدودة عملياً | نمو بدون قيود |
| ACID Transactions | سلامة البيانات 100% |
| Foreign Keys + Constraints | تكامل مرجعي مضمون |
| Row Level Security | أمان على مستوى الصف |
| Full-Text Search | بحث سريع بالعربي |
| API سريع (< 100ms) | تجربة مستخدم ممتازة |
| WebSocket Support | تحديثات فورية |
| ملكية كاملة للبيانات | استقلال تام |

---

## 2. البنية المعمارية الجديدة

```
┌─────────────────────────────────────────────────────────────┐
│                        VPS Server                           │
│                   (Ubuntu 22.04 LTS)                        │
│                                                             │
│  ┌──────────────────────────────────────────────────────┐   │
│  │              Nginx (Reverse Proxy + SSL)              │   │
│  │              Port 80/443                              │   │
│  └──────────┬──────────────────────┬────────────────────┘   │
│             │                      │                        │
│  ┌──────────▼──────────┐  ┌───────▼────────────────────┐   │
│  │   Node.js API       │  │   Web Dashboard (Future)   │   │
│  │   (Express/Fastify) │  │   (React/Next.js)          │   │
│  │   Port 3000         │  │   Port 3001                │   │
│  └──────────┬──────────┘  └────────────────────────────┘   │
│             │                                               │
│  ┌──────────▼──────────────────────────────────────────┐   │
│  │              PostgreSQL 16                            │   │
│  │              Port 5432                                │   │
│  │              + pgvector (للبحث الذكي مستقبلاً)       │   │
│  └──────────────────────────────────────────────────────┘   │
│                                                             │
│  ┌──────────────────────────────────────────────────────┐   │
│  │              Redis (Cache + Sessions)                 │   │
│  │              Port 6379                                │   │
│  └──────────────────────────────────────────────────────┘   │
│                                                             │
└─────────────────────────────────────────────────────────────┘
          │                    │                    │
          ▼                    ▼                    ▼
   ┌─────────────┐   ┌──────────────┐   ┌──────────────────┐
   │ Telegram Bot │   │ Gemini API   │   │ File Storage     │
   │ API          │   │ (AI/NLP)     │   │ (Local/S3)       │
   └─────────────┘   └──────────────┘   └──────────────────┘
```

---

## 3. اختيار التقنيات

### Backend Stack
| التقنية | السبب |
|---|---|
| **Node.js 20 LTS** | نفس لغة المشروع الحالي (JavaScript) - سهولة نقل المنطق |
| **Fastify** | أسرع من Express بـ 2-3x، مناسب لـ API |
| **PostgreSQL 16** | أقوى قاعدة بيانات مفتوحة المصدر |
| **Prisma ORM** | Type-safe، سهل التعامل مع PostgreSQL |
| **Redis** | Cache للصلاحيات وحالات المحادثة |
| **Nginx** | Reverse proxy + SSL termination |
| **PM2** | Process manager لـ Node.js |
| **Docker** (اختياري) | تسهيل النشر والتحديث |

### لماذا Node.js وليس Python أو غيره؟
1. **المشروع الحالي بـ JavaScript** - نقل المنطق مباشرة
2. **أداء ممتاز** لـ I/O operations (API, Database, Telegram)
3. **npm ecosystem** غني جداً
4. **Telegram Bot libraries** ناضجة (grammy/telegraf)
5. **منحنى تعلم أقل** - نفس اللغة

---

## 4. تصميم قاعدة البيانات PostgreSQL

### 4.1 الجداول الأساسية (Core Tables)

```sql
-- =============================================
-- 1. المشاريع (Projects)
-- =============================================
CREATE TABLE projects (
    id              SERIAL PRIMARY KEY,
    code            VARCHAR(50) UNIQUE NOT NULL,
    name            VARCHAR(255) NOT NULL,
    status          VARCHAR(50) DEFAULT 'نشط',
    description     TEXT,
    created_at      TIMESTAMPTZ DEFAULT NOW(),
    updated_at      TIMESTAMPTZ DEFAULT NOW()
);

CREATE INDEX idx_projects_code ON projects(code);
CREATE INDEX idx_projects_name ON projects(name);

-- =============================================
-- 2. الأطراف (Parties - Vendors/Clients/Funders)
-- =============================================
CREATE TYPE party_type AS ENUM ('مورد', 'عميل', 'ممول');

CREATE TABLE parties (
    id                  SERIAL PRIMARY KEY,
    name                VARCHAR(255) NOT NULL,
    type                party_type NOT NULL,
    specialization      VARCHAR(255),
    phone               VARCHAR(30),
    email               VARCHAR(255),
    city                VARCHAR(100),
    preferred_payment   VARCHAR(100),
    bank_details        TEXT,
    notes               TEXT,
    is_active           BOOLEAN DEFAULT TRUE,
    created_at          TIMESTAMPTZ DEFAULT NOW(),
    updated_at          TIMESTAMPTZ DEFAULT NOW()
);

CREATE INDEX idx_parties_name ON parties(name);
CREATE INDEX idx_parties_type ON parties(type);
CREATE UNIQUE INDEX idx_parties_name_type ON parties(name, type);

-- =============================================
-- 3. البنود (Items / Expense Categories)
-- =============================================
CREATE TYPE nature_type AS ENUM (
    'استحقاق مصروف', 'دفعة مصروف',
    'استحقاق إيراد', 'تحصيل إيراد',
    'تمويل', 'استلام تمويل', 'سداد تمويل',
    'تأمين مدفوع للقناة', 'استرداد تأمين من القناة',
    'تحويل داخلي'
);

CREATE TYPE classification_type AS ENUM (
    'مصروفات مباشرة', 'مصروفات عمومية',
    'ايراد',
    'تمويل طويل الأجل', 'سداد تمويل طويل الأجل',
    'سلفة قصيرة', 'سداد سلفة قصيرة',
    'تأمين', 'استرداد تأمين',
    'تحويل للبنك', 'تحويل للخزنة'
);

CREATE TABLE items (
    id              SERIAL PRIMARY KEY,
    name            VARCHAR(255) NOT NULL UNIQUE,
    nature          nature_type,
    classification  classification_type,
    unit_type       VARCHAR(50),  -- مقابلة، دقيقة، رسمة، ضيف
    created_at      TIMESTAMPTZ DEFAULT NOW()
);

-- =============================================
-- 4. العملات وأسعار الصرف (Currencies)
-- =============================================
CREATE TYPE currency_code AS ENUM ('USD', 'TRY', 'EGP');

CREATE TABLE exchange_rates (
    id          SERIAL PRIMARY KEY,
    currency    currency_code NOT NULL,
    rate        DECIMAL(10,6) NOT NULL,  -- سعر الصرف مقابل الدولار
    valid_date  DATE NOT NULL,
    created_at  TIMESTAMPTZ DEFAULT NOW(),
    UNIQUE(currency, valid_date)
);

-- =============================================
-- 5. الحركات المالية (Transactions - القلب)
-- =============================================
CREATE TYPE movement_type AS ENUM ('مدين استحقاق', 'دائن دفعة');
CREATE TYPE payment_method_type AS ENUM ('نقدي', 'تحويل بنكي', 'شيك', 'بطاقة', 'أخرى');
CREATE TYPE payment_term_type AS ENUM ('فوري', 'بعد التسليم', 'تاريخ مخصص');
CREATE TYPE payment_status_type AS ENUM ('مدفوع بالكامل', 'معلق', 'عملية دفع/تحصيل');

CREATE TABLE transactions (
    id                  SERIAL PRIMARY KEY,
    transaction_date    DATE NOT NULL,
    nature              nature_type NOT NULL,
    classification      classification_type NOT NULL,
    project_id          INTEGER REFERENCES projects(id),
    item_id             INTEGER REFERENCES items(id),
    details             TEXT,
    party_id            INTEGER REFERENCES parties(id),
    amount              DECIMAL(14,2) NOT NULL,
    currency            currency_code NOT NULL DEFAULT 'USD',
    exchange_rate       DECIMAL(10,6) NOT NULL DEFAULT 1.0,
    amount_usd          DECIMAL(14,2) GENERATED ALWAYS AS (amount * exchange_rate) STORED,
    movement_type       movement_type NOT NULL,
    reference_number    VARCHAR(100),
    payment_method      payment_method_type,
    payment_term        payment_term_type DEFAULT 'فوري',
    payment_term_weeks  INTEGER,
    payment_term_date   DATE,
    due_date            DATE,
    payment_status      payment_status_type DEFAULT 'معلق',
    month               VARCHAR(7) GENERATED ALWAYS AS (TO_CHAR(transaction_date, 'YYYY-MM')) STORED,
    notes               TEXT,
    attachment_url      TEXT,
    -- حقول التتبع
    created_by          INTEGER REFERENCES users(id),
    created_at          TIMESTAMPTZ DEFAULT NOW(),
    updated_at          TIMESTAMPTZ DEFAULT NOW(),
    source              VARCHAR(50) DEFAULT 'manual'  -- manual, telegram_bot, ai_bot
);

-- فهارس للأداء
CREATE INDEX idx_transactions_date ON transactions(transaction_date);
CREATE INDEX idx_transactions_project ON transactions(project_id);
CREATE INDEX idx_transactions_party ON transactions(party_id);
CREATE INDEX idx_transactions_nature ON transactions(nature);
CREATE INDEX idx_transactions_month ON transactions(month);
CREATE INDEX idx_transactions_status ON transactions(payment_status);
CREATE INDEX idx_transactions_due_date ON transactions(due_date);

-- =============================================
-- 6. الموازنات (Budgets)
-- =============================================
CREATE TABLE budgets (
    id          SERIAL PRIMARY KEY,
    project_id  INTEGER NOT NULL REFERENCES projects(id),
    item_id     INTEGER REFERENCES items(id),
    amount      DECIMAL(14,2) NOT NULL,
    currency    currency_code DEFAULT 'USD',
    notes       TEXT,
    created_at  TIMESTAMPTZ DEFAULT NOW(),
    updated_at  TIMESTAMPTZ DEFAULT NOW()
);

CREATE INDEX idx_budgets_project ON budgets(project_id);
```

### 4.2 جداول إدارة البوت (Bot Management)

```sql
-- =============================================
-- 7. المستخدمون (Users)
-- =============================================
CREATE TABLE users (
    id                      SERIAL PRIMARY KEY,
    name                    VARCHAR(255) NOT NULL,
    email                   VARCHAR(255),
    phone                   VARCHAR(30),
    telegram_username       VARCHAR(255),
    telegram_chat_id        BIGINT UNIQUE,
    -- الصلاحيات
    perm_traditional_bot    BOOLEAN DEFAULT FALSE,
    perm_ai_bot             BOOLEAN DEFAULT FALSE,
    perm_sheet              BOOLEAN DEFAULT FALSE,
    perm_review             BOOLEAN DEFAULT FALSE,
    perm_admin              BOOLEAN DEFAULT FALSE,  -- صلاحية جديدة
    is_active               BOOLEAN DEFAULT TRUE,
    password_hash           VARCHAR(255),  -- للوحة التحكم المستقبلية
    last_login              TIMESTAMPTZ,
    created_at              TIMESTAMPTZ DEFAULT NOW(),
    updated_at              TIMESTAMPTZ DEFAULT NOW()
);

CREATE INDEX idx_users_telegram ON users(telegram_chat_id);
CREATE INDEX idx_users_username ON users(telegram_username);

-- =============================================
-- 8. حركات البوت المعلقة (Pending Bot Transactions)
-- =============================================
CREATE TYPE review_status AS ENUM ('قيد الانتظار', 'معتمد', 'مرفوض', 'يحتاج تعديل');

CREATE TABLE pending_transactions (
    id                  SERIAL PRIMARY KEY,
    -- نفس حقول الحركة المالية
    transaction_date    DATE,
    nature              nature_type,
    classification      classification_type,
    project_id          INTEGER REFERENCES projects(id),
    item_id             INTEGER REFERENCES items(id),
    details             TEXT,
    party_id            INTEGER REFERENCES parties(id),
    party_name_raw      VARCHAR(255),  -- الاسم الخام قبل المطابقة
    amount              DECIMAL(14,2),
    currency            currency_code DEFAULT 'USD',
    exchange_rate       DECIMAL(10,6) DEFAULT 1.0,
    amount_usd          DECIMAL(14,2),
    movement_type       movement_type,
    payment_method      payment_method_type,
    payment_term        payment_term_type,
    payment_term_weeks  INTEGER,
    payment_term_date   DATE,
    due_date            DATE,
    payment_status      payment_status_type,
    notes               TEXT,
    unit_count          INTEGER,
    -- حقول المراجعة
    review_status       review_status DEFAULT 'قيد الانتظار',
    submitted_by        INTEGER REFERENCES users(id),
    submitted_at        TIMESTAMPTZ DEFAULT NOW(),
    reviewed_by         INTEGER REFERENCES users(id),
    reviewed_at         TIMESTAMPTZ,
    review_notes        TEXT,
    attachment_url      TEXT,
    is_new_party        BOOLEAN DEFAULT FALSE,
    source              VARCHAR(50) DEFAULT 'telegram_bot',
    -- الحركة المعتمدة
    approved_transaction_id INTEGER REFERENCES transactions(id)
);

CREATE INDEX idx_pending_review ON pending_transactions(review_status);
CREATE INDEX idx_pending_submitted ON pending_transactions(submitted_by);

-- =============================================
-- 9. أطراف البوت المعلقة (Pending Bot Parties)
-- =============================================
CREATE TABLE pending_parties (
    id                  SERIAL PRIMARY KEY,
    name                VARCHAR(255) NOT NULL,
    type                party_type NOT NULL,
    specialization      VARCHAR(255),
    phone               VARCHAR(30),
    email               VARCHAR(255),
    city                VARCHAR(100),
    preferred_payment   VARCHAR(100),
    bank_details        TEXT,
    notes               TEXT,
    review_status       review_status DEFAULT 'قيد الانتظار',
    submitted_by        INTEGER REFERENCES users(id),
    submitted_at        TIMESTAMPTZ DEFAULT NOW(),
    reviewed_by         INTEGER REFERENCES users(id),
    reviewed_at         TIMESTAMPTZ,
    review_notes        TEXT,
    approved_party_id   INTEGER REFERENCES parties(id)
);

-- =============================================
-- 10. حالات المحادثة (Conversation States)
-- =============================================
CREATE TABLE conversation_states (
    user_id         INTEGER PRIMARY KEY REFERENCES users(id),
    bot_type        VARCHAR(20) NOT NULL,  -- 'traditional' or 'ai'
    state           VARCHAR(100) NOT NULL DEFAULT 'idle',
    context         JSONB DEFAULT '{}',    -- بيانات المحادثة الحالية
    last_message_id BIGINT,
    updated_at      TIMESTAMPTZ DEFAULT NOW()
);

-- =============================================
-- 11. سجل النشاط (Activity Log)
-- =============================================
CREATE TABLE activity_log (
    id          SERIAL PRIMARY KEY,
    user_id     INTEGER REFERENCES users(id),
    action      VARCHAR(100) NOT NULL,
    entity_type VARCHAR(50),  -- transaction, party, project, etc.
    entity_id   INTEGER,
    details     JSONB,
    ip_address  INET,
    created_at  TIMESTAMPTZ DEFAULT NOW()
);

CREATE INDEX idx_activity_date ON activity_log(created_at);
CREATE INDEX idx_activity_user ON activity_log(user_id);

-- =============================================
-- 12. أرشيف المرفوضات (Rejected Archive)
-- =============================================
CREATE TABLE rejected_transactions (
    id                      SERIAL PRIMARY KEY,
    pending_transaction_id  INTEGER,
    data                    JSONB NOT NULL,  -- كل بيانات الحركة المرفوضة
    rejected_by             INTEGER REFERENCES users(id),
    rejected_at             TIMESTAMPTZ DEFAULT NOW(),
    rejection_reason        TEXT
);
```

### 4.3 Views للتقارير (بدل الشيتات المنفصلة)

```sql
-- =============================================
-- كشف حساب طرف (Party Statement)
-- =============================================
CREATE OR REPLACE VIEW party_statement AS
SELECT
    t.id,
    t.transaction_date,
    t.nature,
    t.classification,
    p.code AS project_code,
    p.name AS project_name,
    i.name AS item_name,
    t.details,
    pa.name AS party_name,
    pa.type AS party_type,
    t.amount,
    t.currency,
    t.exchange_rate,
    t.amount_usd,
    t.movement_type,
    t.payment_method,
    t.payment_status,
    t.due_date,
    SUM(
        CASE WHEN t.movement_type = 'مدين استحقاق'
             THEN t.amount_usd
             ELSE -t.amount_usd
        END
    ) OVER (PARTITION BY t.party_id ORDER BY t.transaction_date, t.id) AS running_balance
FROM transactions t
LEFT JOIN projects p ON t.project_id = p.id
LEFT JOIN items i ON t.item_id = i.id
LEFT JOIN parties pa ON t.party_id = pa.id;

-- =============================================
-- تقرير المشروع التفصيلي (Project Report)
-- =============================================
CREATE OR REPLACE VIEW project_report AS
SELECT
    p.code AS project_code,
    p.name AS project_name,
    t.nature,
    t.classification,
    i.name AS item_name,
    pa.name AS party_name,
    COUNT(*) AS transaction_count,
    SUM(t.amount_usd) AS total_usd,
    SUM(CASE WHEN t.movement_type = 'مدين استحقاق' THEN t.amount_usd ELSE 0 END) AS total_debit,
    SUM(CASE WHEN t.movement_type = 'دائن دفعة' THEN t.amount_usd ELSE 0 END) AS total_credit
FROM transactions t
JOIN projects p ON t.project_id = p.id
LEFT JOIN items i ON t.item_id = i.id
LEFT JOIN parties pa ON t.party_id = pa.id
GROUP BY p.code, p.name, t.nature, t.classification, i.name, pa.name;

-- =============================================
-- التنبيهات والاستحقاقات (Due Alerts)
-- =============================================
CREATE OR REPLACE VIEW due_alerts AS
SELECT
    t.id,
    t.transaction_date,
    t.due_date,
    t.due_date - CURRENT_DATE AS days_remaining,
    pa.name AS party_name,
    pa.type AS party_type,
    p.name AS project_name,
    t.amount,
    t.currency,
    t.amount_usd,
    t.payment_status,
    t.notes
FROM transactions t
LEFT JOIN parties pa ON t.party_id = pa.id
LEFT JOIN projects p ON t.project_id = p.id
WHERE t.payment_status = 'معلق'
  AND t.due_date IS NOT NULL
ORDER BY t.due_date ASC;

-- =============================================
-- التدفقات النقدية (Cash Flow)
-- =============================================
CREATE OR REPLACE VIEW cash_flow AS
SELECT
    t.month,
    SUM(CASE WHEN t.nature IN ('تحصيل إيراد', 'استلام تمويل')
             THEN t.amount_usd ELSE 0 END) AS cash_in,
    SUM(CASE WHEN t.nature IN ('دفعة مصروف', 'سداد تمويل')
             THEN t.amount_usd ELSE 0 END) AS cash_out,
    SUM(CASE WHEN t.nature IN ('تحصيل إيراد', 'استلام تمويل')
             THEN t.amount_usd ELSE 0 END) -
    SUM(CASE WHEN t.nature IN ('دفعة مصروف', 'سداد تمويل')
             THEN t.amount_usd ELSE 0 END) AS net_cash_flow
FROM transactions t
GROUP BY t.month
ORDER BY t.month;

-- =============================================
-- لوحة التحكم (Dashboard)
-- =============================================
CREATE OR REPLACE VIEW dashboard AS
SELECT
    (SELECT COUNT(*) FROM projects WHERE status = 'نشط') AS active_projects,
    (SELECT COUNT(*) FROM parties WHERE is_active = TRUE) AS active_parties,
    (SELECT COUNT(*) FROM transactions) AS total_transactions,
    (SELECT COALESCE(SUM(amount_usd), 0) FROM transactions
     WHERE nature LIKE '%مصروف%') AS total_expenses,
    (SELECT COALESCE(SUM(amount_usd), 0) FROM transactions
     WHERE nature LIKE '%إيراد%') AS total_revenue,
    (SELECT COUNT(*) FROM pending_transactions
     WHERE review_status = 'قيد الانتظار') AS pending_reviews,
    (SELECT COUNT(*) FROM transactions
     WHERE payment_status = 'معلق' AND due_date < CURRENT_DATE) AS overdue_count;
```

---

## 5. هيكل مجلدات المشروع الجديد

```
seen-accounting-api/
├── src/
│   ├── config/
│   │   ├── database.js          # PostgreSQL connection
│   │   ├── redis.js             # Redis connection
│   │   ├── constants.js         # ENUMs and constants (from Config.js)
│   │   └── env.js               # Environment variables
│   │
│   ├── database/
│   │   ├── migrations/          # Database migrations (Prisma)
│   │   ├── seeds/               # Initial data
│   │   └── schema.prisma        # Prisma schema
│   │
│   ├── modules/
│   │   ├── transactions/
│   │   │   ├── transactions.routes.js
│   │   │   ├── transactions.service.js
│   │   │   └── transactions.validation.js
│   │   │
│   │   ├── projects/
│   │   │   ├── projects.routes.js
│   │   │   └── projects.service.js
│   │   │
│   │   ├── parties/
│   │   │   ├── parties.routes.js
│   │   │   └── parties.service.js
│   │   │
│   │   ├── items/
│   │   │   ├── items.routes.js
│   │   │   └── items.service.js
│   │   │
│   │   ├── reports/
│   │   │   ├── reports.routes.js
│   │   │   ├── reports.service.js
│   │   │   └── pdf-generator.js
│   │   │
│   │   ├── auth/
│   │   │   ├── auth.routes.js
│   │   │   ├── auth.service.js
│   │   │   └── auth.middleware.js
│   │   │
│   │   └── review/
│   │       ├── review.routes.js
│   │       └── review.service.js
│   │
│   ├── bots/
│   │   ├── telegram/
│   │   │   ├── traditional-bot.js   # نقل TelegramBot.js
│   │   │   ├── commands.js          # أوامر البوت
│   │   │   ├── keyboards.js         # لوحات المفاتيح
│   │   │   └── handlers.js          # معالجة الرسائل
│   │   │
│   │   └── ai/
│   │       ├── ai-bot.js            # نقل AIBot.js
│   │       ├── ai-agent.js          # نقل AIAgent.js
│   │       ├── gemini-client.js     # Gemini API client
│   │       └── inference-rules.js   # قواعد الاستنتاج
│   │
│   ├── shared/
│   │   ├── fuzzy-search.js          # نقل FuzzySearch.js
│   │   ├── logger.js                # Winston logger
│   │   └── errors.js                # Custom error classes
│   │
│   └── app.js                       # Entry point
│
├── prisma/
│   └── schema.prisma
│
├── .env                             # Environment variables
├── .env.example
├── docker-compose.yml               # PostgreSQL + Redis
├── ecosystem.config.js              # PM2 config
├── package.json
└── README.md
```

---

## 6. خطة الترحيل (Migration Strategy)

### المرحلة 1: البنية التحتية (أسبوع 1)
```
[ ] إعداد VPS (Ubuntu 22.04 LTS)
[ ] تثبيت PostgreSQL 16
[ ] تثبيت Node.js 20 LTS
[ ] تثبيت Redis
[ ] تثبيت Nginx + SSL (Let's Encrypt)
[ ] تثبيت PM2
[ ] إعداد Firewall (UFW)
[ ] إعداد SSH key authentication
[ ] إعداد النسخ الاحتياطي التلقائي
```

### المرحلة 2: قاعدة البيانات (أسبوع 2)
```
[ ] إنشاء قاعدة البيانات والجداول
[ ] إعداد Prisma ORM
[ ] كتابة Migration scripts
[ ] ترحيل البيانات من Google Sheets
    - تصدير كل شيت كـ CSV
    - تنظيف البيانات
    - استيراد إلى PostgreSQL
[ ] التحقق من سلامة البيانات
[ ] إعداد النسخ الاحتياطي اليومي
```

### المرحلة 3: API الأساسي (أسبوع 3-4)
```
[ ] إعداد مشروع Fastify
[ ] إنشاء CRUD endpoints للجداول الأساسية
    - /api/transactions
    - /api/projects
    - /api/parties
    - /api/items
[ ] نظام المصادقة (JWT)
[ ] Validation middleware
[ ] Error handling
[ ] Rate limiting
[ ] Logging (Winston)
```

### المرحلة 4: بوت تيليجرام (أسبوع 5-6)
```
[ ] نقل Traditional Bot إلى Grammy/Telegraf
[ ] ربط البوت بـ PostgreSQL بدل Google Sheets
[ ] نقل AI Bot مع Gemini integration
[ ] نظام المراجعة (Review System)
[ ] اختبار كامل للبوت
[ ] التحويل من Polling إلى Webhook (أسرع)
```

### المرحلة 5: التقارير (أسبوع 7)
```
[ ] نقل منطق التقارير
[ ] إنشاء SQL Views للتقارير
[ ] PDF Generator (PDFKit بدل Google Sheets PDF)
[ ] إرسال التقارير عبر تيليجرام
```

### المرحلة 6: الاختبار والتشغيل (أسبوع 8)
```
[ ] اختبار شامل لكل الوظائف
[ ] Performance testing
[ ] تشغيل النظامين بالتوازي (Google + VPS) لمدة أسبوع
[ ] التأكد من تطابق النتائج
[ ] إيقاف النظام القديم
```

---

## 7. تقدير التكلفة الشهرية

### VPS Recommendations

| المزود | المواصفات | السعر/شهر | ملاحظات |
|---|---|---|---|
| **Hetzner** (الأفضل) | 4 vCPU, 8GB RAM, 80GB NVMe | ~$7 | أوروبا - أداء ممتاز |
| **DigitalOcean** | 2 vCPU, 4GB RAM, 80GB SSD | $24 | سهل الاستخدام |
| **Contabo** | 6 vCPU, 16GB RAM, 200GB | ~$7 | ألمانيا - سعر ممتاز |
| **Vultr** | 2 vCPU, 4GB RAM, 80GB NVMe | $24 | أداء جيد |

### التكلفة الإجمالية المتوقعة
| البند | التكلفة/شهر |
|---|---|
| VPS (Hetzner/Contabo) | $7-12 |
| Domain + SSL | مجاني (Let's Encrypt) |
| Backup Storage | $1-3 |
| **الإجمالي** | **$8-15/شهر** |

> **مقارنة:** Google Workspace مجاني حالياً لكن محدود. VPS يعطيك قوة وتحكم أكبر بفرق بسيط.

---

## 8. الأمان (Security Checklist)

```
[ ] SSH Key-only authentication (disable password login)
[ ] UFW Firewall (allow only 22, 80, 443)
[ ] PostgreSQL: listen only on localhost
[ ] Redis: bind to localhost only
[ ] SSL/TLS for all connections
[ ] Environment variables for all secrets
[ ] Rate limiting on API
[ ] SQL injection protection (Prisma ORM)
[ ] Input validation (Zod/Joi)
[ ] Regular security updates (unattended-upgrades)
[ ] Daily encrypted backups
[ ] Fail2ban for SSH brute force protection
[ ] API keys rotation policy
```

---

## 9. النسخ الاحتياطي (Backup Strategy)

```bash
# Automated daily backup script
# PostgreSQL dump
pg_dump -U seen_user seen_accounting | gzip > /backups/db_$(date +%Y%m%d).sql.gz

# Keep last 30 days locally
find /backups -mtime +30 -delete

# Upload to external storage (optional)
# rclone copy /backups remote:seen-backups
```

### Backup Schedule
| النوع | التكرار | الاحتفاظ |
|---|---|---|
| Full Database Dump | يومياً 3:00 صباحاً | 30 يوم |
| WAL Archiving | مستمر | 7 أيام |
| External Copy | أسبوعياً | 90 يوم |

---

## 10. خطة الطوارئ (Disaster Recovery)

| السيناريو | الحل | وقت الاستعادة |
|---|---|---|
| تعطل VPS | استعادة من Snapshot + آخر backup | < 1 ساعة |
| تلف قاعدة البيانات | استعادة من pg_dump | < 30 دقيقة |
| اختراق أمني | إعادة بناء من backup نظيف | < 2 ساعة |
| فشل المزود | نقل لمزود آخر | < 4 ساعات |

---

## 11. مقارنة قبل وبعد

| المعيار | الحالي (Google Sheets) | الجديد (VPS + PostgreSQL) |
|---|---|---|
| سرعة الاستجابة | 2-5 ثواني | < 100ms |
| حد البيانات | 10M خلية | لا محدود |
| المستخدمين المتزامنين | ~50 | آلاف |
| النسخ الاحتياطي | يدوي | تلقائي يومي |
| الأمان | متوسط | عالي |
| التوسع المستقبلي | محدود جداً | مفتوح |
| Web Dashboard | مستحيل | ممكن |
| Mobile App | مستحيل | ممكن |
| Multi-tenant | مستحيل | ممكن |
| التكلفة | مجاني | $8-15/شهر |

---

## 12. الخطوة التالية

بعد اعتماد هذه الخطة، نبدأ بالترتيب:

1. **اختيار مزود VPS** (أنصح بـ Hetzner أو Contabo)
2. **إعداد السيرفر** (يمكنني كتابة script تلقائي)
3. **إنشاء قاعدة البيانات** (تطبيق الـ SQL أعلاه)
4. **ترحيل البيانات** (تصدير Google Sheets → PostgreSQL)
5. **بناء API** (نقل المنطق الحالي)
6. **ربط البوت** (تحويل من Google Apps Script إلى Node.js)
