import { getDB } from '../../config/database.js';
import { z } from 'zod';

const createSchema = z.object({
  name: z.string().min(1, 'اسم البند مطلوب'),
  nature: z.string().optional().nullable(),
  classification: z.string().optional().nullable(),
  unitType: z.string().optional().nullable(),
});

export async function itemRoutes(app) {
  const db = () => getDB();

  // قائمة البنود
  app.get('/', { preHandler: [app.authenticate] }, async () => {
    return db().item.findMany({
      orderBy: { name: 'asc' },
      include: {
        _count: { select: { transactions: true } },
      },
    });
  });

  // إنشاء بند
  app.post('/', { preHandler: [app.authenticate] }, async (request, reply) => {
    const data = createSchema.parse(request.body);
    const item = await db().item.create({ data });
    return reply.status(201).send(item);
  });

  // تحديث بند
  app.put('/:id', { preHandler: [app.authenticate] }, async (request) => {
    const data = createSchema.partial().parse(request.body);
    return db().item.update({
      where: { id: Number(request.params.id) },
      data,
    });
  });

  // حذف بند
  app.delete('/:id', { preHandler: [app.authenticate] }, async (request, reply) => {
    const id = Number(request.params.id);
    const hasTransactions = await db().transaction.count({ where: { itemId: id } });
    if (hasTransactions > 0) {
      return reply.status(400).send({
        error: 'لا يمكن حذف بند له حركات مالية',
        code: 'HAS_TRANSACTIONS',
      });
    }
    await db().item.delete({ where: { id } });
    return reply.status(204).send();
  });
}
