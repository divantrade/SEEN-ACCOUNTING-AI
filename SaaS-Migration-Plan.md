# خطة تحويل SEEN Accounting إلى SaaS

## الرؤية

تحويل نظام SEEN المحاسبي من نظام داخلي يعمل على Google Sheets إلى منتج SaaS يمكن لأي شركة إنتاج إعلامي الاشتراك فيه واستخدامه.

---

## الوضع الحالي

```
Google Sheets = قاعدة البيانات + الواجهة + السيرفر (كل شيء في مكان واحد)
Apps Script   = المنطق والتقارير والبوت
```

### نقاط القوة الحالية (تُنقل للنظام الجديد)
- محرك ذكاء اصطناعي يعمل بـ Gemini Function Calling
- 12 تقرير PDF جاهز
- بحث عربي ذكي بـ 5 مستويات مطابقة
- نظام مراجعة واعتماد حركات
- تعلم تفضيلات المستخدم
- كشف الحركات المكررة

### القيود التي تدفع للتحول
- بطء مع البيانات الكبيرة (Google Sheets)
- مستخدم واحد / شركة واحدة
- لا يمكن تخصيص الواجهة
- حدود Apps Script (6 دقائق تشغيل، 100MB بيانات)

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
| الواجهة | Next.js + React + Tailwind | سريعة، تدعم العربية RTL، SEO | مجاني |
| السيرفر | Next.js API Routes | نفس المشروع = أبسط في الإدارة | مجاني |
| قاعدة البيانات | Supabase (PostgreSQL) | Auth جاهز، Row Level Security، مجاني للبداية | $0 → $25/شهر |
| الاستضافة | Vercel | deploy تلقائي من Git، مجاني للبداية | $0 → $20/شهر |
| الدفع والاشتراكات | Stripe | اشتراكات شهرية/سنوية، فواتير تلقائية | 2.9% لكل معاملة |
| توليد PDF | `@react-pdf/renderer` | تقارير احترافية بدون سيرفر منفصل | مجاني |
| بوت تليجرام | Telegram Webhook via API Route | لا يحتاج سيرفر منفصل | مجاني |
| الذكاء الاصطناعي | Gemini / Claude API | نفس المنطق الحالي مع تحسينات | حسب الاستخدام |
| تخزين الملفات | Supabase Storage | بدل Google Drive | ضمن خطة Supabase |
| UI Components | shadcn/ui | مكونات جاهزة وقابلة للتخصيص | مجاني |

### التكلفة المتوقعة
| المرحلة | التكلفة الشهرية |
|---|---|
| البداية (أول 10 شركات) | $0 (كل شيء مجاني) |
| النمو (10-100 شركة) | ~$45/شهر |
| التوسع (100+ شركة) | ~$200/شهر |

---

## تصميم قاعدة البيانات

### الجداول الأساسية

```sql
-- الشركات (tenants)
companies
├── id UUID PRIMARY KEY
├── name TEXT
├── email TEXT
├── timezone TEXT DEFAULT 'Asia/Istanbul'
├── default_currency TEXT DEFAULT 'USD'
├── currencies TEXT[] DEFAULT '{USD,TRY,EGP}'
├── subscription_plan TEXT DEFAULT 'free'
├── stripe_customer_id TEXT
├── created_at TIMESTAMP
└── settings JSONB

-- المستخدمون
users
├── id UUID PRIMARY KEY (Supabase Auth)
├── company_id UUID → companies.id
├── full_name TEXT
├── role TEXT ('admin', 'accountant', 'viewer')
├── telegram_chat_id TEXT
├── phone TEXT
└── created_at TIMESTAMP

-- الحركات المالية (أهم جدول)
transactions
├── id UUID PRIMARY KEY
├── company_id UUID → companies.id
├── transaction_id TEXT (BOT-20250327-...)
├── date DATE
├── nature TEXT (استحقاق مصروف، دفعة، ...)
├── classification TEXT (مصروفات مباشرة، عمومية، ...)
├── project_id UUID → projects.id
├── item_id UUID → items.id
├── party_id UUID → parties.id
├── details TEXT
├── amount DECIMAL(15,2)
├── currency TEXT
├── exchange_rate DECIMAL(10,4)
├── amount_usd DECIMAL(15,2)
├── movement_type TEXT (مدين، دائن، تسوية)
├── payment_method TEXT
├── payment_term TEXT
├── review_status TEXT DEFAULT 'pending'
├── reviewed_by UUID → users.id
├── created_by UUID → users.id
├── created_at TIMESTAMP
└── attachments TEXT[]

-- المشاريع
projects
├── id UUID PRIMARY KEY
├── company_id UUID → companies.id
├── code TEXT
├── name TEXT
├── client TEXT
├── status TEXT ('active', 'completed', 'cancelled')
├── budget DECIMAL(15,2)
└── created_at TIMESTAMP

-- الأطراف (موردين، عملاء، ممولين)
parties
├── id UUID PRIMARY KEY
├── company_id UUID → companies.id
├── name TEXT
├── type TEXT ('مورد', 'عميل', 'ممول')
├── phone TEXT
├── email TEXT
├── specialization TEXT
└── created_at TIMESTAMP

-- البنود المحاسبية
items
├── id UUID PRIMARY KEY
├── company_id UUID → companies.id
├── name TEXT
├── nature TEXT
├── classification TEXT
├── unit TEXT
└── created_at TIMESTAMP

-- أسعار الصرف
exchange_rates
├── id UUID PRIMARY KEY
├── company_id UUID → companies.id
├── from_currency TEXT
├── to_currency TEXT
├── rate DECIMAL(10,4)
└── date DATE
```

### Row Level Security (عزل البيانات)

```sql
-- كل شركة ترى بياناتها فقط
ALTER TABLE transactions ENABLE ROW LEVEL SECURITY;

CREATE POLICY "company_isolation" ON transactions
    FOR ALL
    USING (company_id = (auth.jwt() ->> 'company_id')::UUID);

-- نفس السياسة لكل الجداول
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
| `SmartAgent.js` | `lib/ai/smart-agent.ts` | تحويل لـ TypeScript |
| `AgentTools.js` | `lib/ai/agent-tools.ts` | ربط بـ PostgreSQL بدل Sheets |
| `FuzzySearch.js` | `lib/search/fuzzy-search.ts` | + PostgreSQL full-text search |
| `ReportsConfig.js` | `lib/reports/config.ts` | نفس البنية |

### ما يُعاد بناؤه (يعتمد على Sheets)
| الملف الحالي | المقابل في SaaS | ملاحظات |
|---|---|---|
| `Main.js` (التقارير) | `lib/reports/*.ts` | SQL بدل getValues() |
| `BotSheets.js` | `lib/db/transactions.ts` | Supabase بدل Sheet API |
| `PDFGenerator.js` | `lib/pdf/generator.ts` | @react-pdf بدل Sheets PDF export |
| `ReviewSystem.js` | `lib/review/system.ts` | نفس المنطق، DB مختلف |

### ما يُحذف (خاص بـ Google)
| الملف | السبب |
|---|---|
| `DriveAttachments.js` | يُستبدل بـ Supabase Storage |
| `AIConfig.js` (جزئياً) | الإعدادات تنتقل لـ env vars |
| `appsscript.json` | خاص بـ Apps Script |
| `.clasp.json` | خاص بـ clasp |

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

## الخطوة التالية

1. إنشاء مشروع Next.js جديد (repo منفصل)
2. إعداد Supabase وإنشاء الجداول
3. بناء API الحركات المالية
4. بناء واجهة إدخال أولية
5. نقل بيانات SEEN كاختبار
