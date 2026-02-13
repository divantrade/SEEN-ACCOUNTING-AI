import { getDB } from '../../config/database.js';
import { NotFoundError } from '../../shared/errors.js';
import { fuzzySearch } from '../../shared/fuzzy-search.js';
import { z } from 'zod';

const createSchema = z.object({
  name: z.string().min(1, 'اسم الطرف مطلوب'),
  type: z.enum(['VENDOR', 'CLIENT', 'FUNDER']),
  specialization: z.string().optional().nullable(),
  phone: z.string().optional().nullable(),
  email: z.string().email().optional().nullable(),
  city: z.string().optional().nullable(),
  preferredPayment: z.string().optional().nullable(),
  bankDetails: z.string().optional().nullable(),
  notes: z.string().optional().nullable(),
});

export async function partyRoutes(app) {
  const db = () => getDB();

  // قائمة الأطراف
  app.get('/', { preHandler: [app.authenticate] }, async (request) => {
    const { search, type } = request.query;
    const where = { isActive: true };
    if (type) where.type = type;
    if (search) {
      where.OR = [
        { name: { contains: search, mode: 'insensitive' } },
        { specialization: { contains: search, mode: 'insensitive' } },
        { phone: { contains: search } },
      ];
    }
    return db().party.findMany({
      where,
      orderBy: { name: 'asc' },
      include: {
        _count: { select: { transactions: true } },
      },
    });
  });

  // بحث ذكي (Fuzzy Search)
  app.get('/search', { preHandler: [app.authenticate] }, async (request) => {
    const { q, type } = request.query;
    if (!q) return [];

    const where = { isActive: true };
    if (type) where.type = type;
    const allParties = await db().party.findMany({ where, select: { id: true, name: true, type: true } });
    const results = fuzzySearch(q, allParties, 'name', 0.3);
    return results.map((r) => ({ ...r.item, score: r.score }));
  });

  // طرف واحد مع كشف حسابه
  app.get('/:id', { preHandler: [app.authenticate] }, async (request) => {
    const party = await db().party.findUnique({
      where: { id: Number(request.params.id) },
      include: {
        _count: { select: { transactions: true } },
      },
    });
    if (!party) throw new NotFoundError('الطرف');

    // Get balance
    const balance = await db().transaction.aggregate({
      where: { partyId: party.id },
      _sum: { amountUsd: true },
    });

    return { ...party, totalAmountUsd: balance._sum.amountUsd || 0 };
  });

  // إنشاء طرف
  app.post('/', { preHandler: [app.authenticate] }, async (request, reply) => {
    const data = createSchema.parse(request.body);
    const party = await db().party.create({ data });
    return reply.status(201).send(party);
  });

  // تحديث طرف
  app.put('/:id', { preHandler: [app.authenticate] }, async (request) => {
    const data = createSchema.partial().parse(request.body);
    return db().party.update({
      where: { id: Number(request.params.id) },
      data,
    });
  });

  // تعطيل طرف (soft delete)
  app.delete('/:id', { preHandler: [app.authenticate] }, async (request) => {
    return db().party.update({
      where: { id: Number(request.params.id) },
      data: { isActive: false },
    });
  });
}
