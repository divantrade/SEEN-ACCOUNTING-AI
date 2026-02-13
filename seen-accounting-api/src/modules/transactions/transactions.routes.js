import { transactionService } from './transactions.service.js';
import {
  createTransactionSchema,
  updateTransactionSchema,
  listTransactionsSchema,
} from './transactions.validation.js';

export async function transactionRoutes(app) {
  // ==================== قائمة الحركات ====================
  app.get('/', { preHandler: [app.authenticate] }, async (request) => {
    const filters = listTransactionsSchema.parse(request.query);
    return transactionService.list(filters);
  });

  // ==================== إحصائيات ====================
  app.get('/stats', { preHandler: [app.authenticate] }, async (request) => {
    return transactionService.getStats(request.query);
  });

  // ==================== التنبيهات ====================
  app.get('/alerts', { preHandler: [app.authenticate] }, async () => {
    return transactionService.getDueAlerts();
  });

  // ==================== حركة واحدة ====================
  app.get('/:id', { preHandler: [app.authenticate] }, async (request) => {
    return transactionService.getById(Number(request.params.id));
  });

  // ==================== إنشاء حركة ====================
  app.post('/', { preHandler: [app.authenticate] }, async (request, reply) => {
    const data = createTransactionSchema.parse(request.body);
    const transaction = await transactionService.create({
      ...data,
      createdById: request.user.id,
    });
    return reply.status(201).send(transaction);
  });

  // ==================== تحديث حركة ====================
  app.put('/:id', { preHandler: [app.authenticate] }, async (request) => {
    const data = updateTransactionSchema.parse(request.body);
    return transactionService.update(Number(request.params.id), data);
  });

  // ==================== حذف حركة ====================
  app.delete('/:id', { preHandler: [app.authenticate] }, async (request, reply) => {
    await transactionService.delete(Number(request.params.id));
    return reply.status(204).send();
  });

  // ==================== كشف حساب طرف ====================
  app.get(
    '/party-statement/:partyId',
    { preHandler: [app.authenticate] },
    async (request) => {
      return transactionService.getPartyStatement(
        Number(request.params.partyId),
        request.query
      );
    }
  );
}
