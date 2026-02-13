import { getDB } from '../../config/database.js';
import { requirePermission } from '../auth/auth.middleware.js';
import { transactionService } from '../transactions/transactions.service.js';
import { NotFoundError } from '../../shared/errors.js';
import { logger } from '../../shared/logger.js';

export async function reviewRoutes(app) {
  const db = () => getDB();

  // ==================== قائمة الحركات المعلقة ====================
  app.get(
    '/pending',
    { preHandler: [requirePermission('review')] },
    async (request) => {
      const { status = 'PENDING' } = request.query;
      return db().pendingTransaction.findMany({
        where: { reviewStatus: status },
        include: {
          project: { select: { code: true, name: true } },
          party: { select: { name: true, type: true } },
          item: { select: { name: true } },
          submittedBy: { select: { name: true } },
        },
        orderBy: { submittedAt: 'desc' },
      });
    }
  );

  // ==================== اعتماد حركة ====================
  app.post(
    '/:id/approve',
    { preHandler: [requirePermission('review')] },
    async (request) => {
      const id = Number(request.params.id);
      const { notes } = request.body || {};

      const pending = await db().pendingTransaction.findUnique({
        where: { id },
        include: { project: true, party: true, item: true },
      });
      if (!pending) throw new NotFoundError('الحركة المعلقة');

      // Create approved transaction in main ledger
      const transaction = await transactionService.create({
        transactionDate: pending.transactionDate,
        nature: pending.nature,
        classification: pending.classification,
        projectId: pending.projectId,
        itemId: pending.itemId,
        details: pending.details,
        partyId: pending.partyId,
        amount: Number(pending.amount),
        currency: pending.currency,
        exchangeRate: Number(pending.exchangeRate || 1),
        paymentMethod: pending.paymentMethod,
        paymentTerm: pending.paymentTerm,
        paymentTermWeeks: pending.paymentTermWeeks,
        paymentTermDate: pending.paymentTermDate,
        dueDate: pending.dueDate,
        paymentStatus: pending.paymentStatus,
        notes: pending.notes,
        attachmentUrl: pending.attachmentUrl,
        unitCount: pending.unitCount,
        source: pending.source,
        createdById: pending.submittedById,
      });

      // Update pending status
      await db().pendingTransaction.update({
        where: { id },
        data: {
          reviewStatus: 'APPROVED',
          reviewedById: request.user.id,
          reviewedAt: new Date(),
          reviewNotes: notes || null,
          approvedTransactionId: transaction.id,
        },
      });

      // Log activity
      await db().activityLog.create({
        data: {
          userId: request.user.id,
          action: 'approve_transaction',
          entityType: 'transaction',
          entityId: transaction.id,
          details: { pendingId: id },
        },
      });

      logger.info(`Transaction #${id} approved by user #${request.user.id}`);
      return { message: 'تم اعتماد الحركة بنجاح', transaction };
    }
  );

  // ==================== رفض حركة ====================
  app.post(
    '/:id/reject',
    { preHandler: [requirePermission('review')] },
    async (request) => {
      const id = Number(request.params.id);
      const { reason } = request.body || {};

      const pending = await db().pendingTransaction.findUnique({ where: { id } });
      if (!pending) throw new NotFoundError('الحركة المعلقة');

      // Move to rejected archive
      await db().rejectedTransaction.create({
        data: {
          pendingTransactionId: id,
          data: pending,
          rejectedById: request.user.id,
          rejectionReason: reason || null,
        },
      });

      // Update pending status
      await db().pendingTransaction.update({
        where: { id },
        data: {
          reviewStatus: 'REJECTED',
          reviewedById: request.user.id,
          reviewedAt: new Date(),
          reviewNotes: reason || null,
        },
      });

      logger.info(`Transaction #${id} rejected by user #${request.user.id}`);
      return { message: 'تم رفض الحركة' };
    }
  );

  // ==================== طلب تعديل حركة ====================
  app.post(
    '/:id/request-edit',
    { preHandler: [requirePermission('review')] },
    async (request) => {
      const id = Number(request.params.id);
      const { notes } = request.body || {};

      await db().pendingTransaction.update({
        where: { id },
        data: {
          reviewStatus: 'NEEDS_EDIT',
          reviewedById: request.user.id,
          reviewedAt: new Date(),
          reviewNotes: notes || null,
        },
      });

      return { message: 'تم طلب تعديل الحركة' };
    }
  );

  // ==================== أطراف معلقة ====================
  app.get(
    '/pending-parties',
    { preHandler: [requirePermission('review')] },
    async () => {
      return db().pendingParty.findMany({
        where: { reviewStatus: 'PENDING' },
        include: { submittedBy: { select: { name: true } } },
        orderBy: { submittedAt: 'desc' },
      });
    }
  );

  // ==================== اعتماد طرف ====================
  app.post(
    '/party/:id/approve',
    { preHandler: [requirePermission('review')] },
    async (request) => {
      const id = Number(request.params.id);
      const pending = await db().pendingParty.findUnique({ where: { id } });
      if (!pending) throw new NotFoundError('الطرف المعلق');

      const party = await db().party.create({
        data: {
          name: pending.name,
          type: pending.type,
          specialization: pending.specialization,
          phone: pending.phone,
          email: pending.email,
          city: pending.city,
          preferredPayment: pending.preferredPayment,
          bankDetails: pending.bankDetails,
          notes: pending.notes,
        },
      });

      await db().pendingParty.update({
        where: { id },
        data: {
          reviewStatus: 'APPROVED',
          reviewedById: request.user.id,
          reviewedAt: new Date(),
          approvedPartyId: party.id,
        },
      });

      return { message: 'تم اعتماد الطرف بنجاح', party };
    }
  );
}
