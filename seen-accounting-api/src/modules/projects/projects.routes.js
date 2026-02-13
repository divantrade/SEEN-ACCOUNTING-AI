import { getDB } from '../../config/database.js';
import { NotFoundError } from '../../shared/errors.js';
import { z } from 'zod';

const createSchema = z.object({
  code: z.string().min(1, 'كود المشروع مطلوب'),
  name: z.string().min(1, 'اسم المشروع مطلوب'),
  status: z.string().optional().default('نشط'),
  description: z.string().optional().nullable(),
});

export async function projectRoutes(app) {
  const db = () => getDB();

  // قائمة المشاريع
  app.get('/', { preHandler: [app.authenticate] }, async (request) => {
    const { search, status } = request.query;
    const where = {};
    if (status) where.status = status;
    if (search) {
      where.OR = [
        { name: { contains: search, mode: 'insensitive' } },
        { code: { contains: search, mode: 'insensitive' } },
      ];
    }
    return db().project.findMany({
      where,
      orderBy: { name: 'asc' },
      include: {
        _count: { select: { transactions: true } },
      },
    });
  });

  // مشروع واحد مع إحصائياته
  app.get('/:id', { preHandler: [app.authenticate] }, async (request) => {
    const project = await db().project.findUnique({
      where: { id: Number(request.params.id) },
      include: {
        _count: { select: { transactions: true } },
        budgets: true,
      },
    });
    if (!project) throw new NotFoundError('المشروع');

    // Get financial summary
    const summary = await db().transaction.aggregate({
      where: { projectId: project.id },
      _sum: { amountUsd: true },
      _count: true,
    });

    return { ...project, financialSummary: summary };
  });

  // إنشاء مشروع
  app.post('/', { preHandler: [app.authenticate] }, async (request, reply) => {
    const data = createSchema.parse(request.body);
    const project = await db().project.create({ data });
    return reply.status(201).send(project);
  });

  // تحديث مشروع
  app.put('/:id', { preHandler: [app.authenticate] }, async (request) => {
    const data = createSchema.partial().parse(request.body);
    return db().project.update({
      where: { id: Number(request.params.id) },
      data,
    });
  });

  // حذف مشروع
  app.delete('/:id', { preHandler: [app.authenticate] }, async (request, reply) => {
    const id = Number(request.params.id);
    const hasTransactions = await db().transaction.count({ where: { projectId: id } });
    if (hasTransactions > 0) {
      return reply.status(400).send({
        error: 'لا يمكن حذف مشروع له حركات مالية',
        code: 'HAS_TRANSACTIONS',
      });
    }
    await db().project.delete({ where: { id } });
    return reply.status(204).send();
  });
}
