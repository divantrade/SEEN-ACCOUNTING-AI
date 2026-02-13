import { z } from 'zod';

export const createTransactionSchema = z.object({
  transactionDate: z.string().regex(/^\d{4}-\d{2}-\d{2}$/, 'التاريخ يجب أن يكون بصيغة YYYY-MM-DD'),
  nature: z.string().min(1, 'طبيعة الحركة مطلوبة'),
  classification: z.string().min(1, 'تصنيف الحركة مطلوب'),
  projectId: z.number().int().positive().optional().nullable(),
  itemId: z.number().int().positive().optional().nullable(),
  details: z.string().optional().nullable(),
  partyId: z.number().int().positive().optional().nullable(),
  amount: z.number().positive('المبلغ يجب أن يكون أكبر من صفر'),
  currency: z.enum(['USD', 'TRY', 'EGP']).default('USD'),
  exchangeRate: z.number().positive().default(1),
  referenceNumber: z.string().optional().nullable(),
  paymentMethod: z.string().optional().nullable(),
  paymentTerm: z.string().optional().default('IMMEDIATE'),
  paymentTermWeeks: z.number().int().positive().optional().nullable(),
  paymentTermDate: z.string().optional().nullable(),
  dueDate: z.string().optional().nullable(),
  paymentStatus: z.string().optional().default('PENDING'),
  notes: z.string().optional().nullable(),
  unitCount: z.number().int().positive().optional().nullable(),
});

export const updateTransactionSchema = createTransactionSchema.partial();

export const listTransactionsSchema = z.object({
  page: z.coerce.number().int().positive().default(1),
  limit: z.coerce.number().int().positive().max(200).default(50),
  projectId: z.coerce.number().int().positive().optional(),
  partyId: z.coerce.number().int().positive().optional(),
  nature: z.string().optional(),
  classification: z.string().optional(),
  paymentStatus: z.string().optional(),
  dateFrom: z.string().optional(),
  dateTo: z.string().optional(),
  month: z.string().optional(),
  search: z.string().optional(),
});
