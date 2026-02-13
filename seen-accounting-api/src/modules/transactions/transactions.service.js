import { getDB } from '../../config/database.js';
import { getMovementType } from '../../config/constants.js';
import { NotFoundError } from '../../shared/errors.js';

const db = () => getDB();

export const transactionService = {
  // ==================== إنشاء حركة مالية ====================
  async create(data) {
    const movementType = getMovementType(data.nature);
    const amountUsd = Number(data.amount) * Number(data.exchangeRate || 1);
    const month = data.transactionDate
      ? new Date(data.transactionDate).toISOString().slice(0, 7)
      : null;

    return db().transaction.create({
      data: {
        transactionDate: new Date(data.transactionDate),
        nature: data.nature,
        classification: data.classification,
        projectId: data.projectId || null,
        itemId: data.itemId || null,
        details: data.details || null,
        partyId: data.partyId || null,
        amount: data.amount,
        currency: data.currency || 'USD',
        exchangeRate: data.exchangeRate || 1,
        amountUsd,
        movementType,
        referenceNumber: data.referenceNumber || null,
        paymentMethod: data.paymentMethod || null,
        paymentTerm: data.paymentTerm || 'IMMEDIATE',
        paymentTermWeeks: data.paymentTermWeeks || null,
        paymentTermDate: data.paymentTermDate ? new Date(data.paymentTermDate) : null,
        dueDate: data.dueDate ? new Date(data.dueDate) : null,
        paymentStatus: data.paymentStatus || 'PENDING',
        month,
        notes: data.notes || null,
        attachmentUrl: data.attachmentUrl || null,
        unitCount: data.unitCount || null,
        source: data.source || 'manual',
        createdById: data.createdById || null,
      },
      include: {
        project: true,
        party: true,
        item: true,
      },
    });
  },

  // ==================== قراءة حركة واحدة ====================
  async getById(id) {
    const transaction = await db().transaction.findUnique({
      where: { id },
      include: {
        project: true,
        party: true,
        item: true,
        createdBy: { select: { id: true, name: true } },
      },
    });
    if (!transaction) throw new NotFoundError('الحركة المالية');
    return transaction;
  },

  // ==================== قائمة الحركات مع فلاتر ====================
  async list({
    page = 1,
    limit = 50,
    projectId,
    partyId,
    nature,
    classification,
    paymentStatus,
    dateFrom,
    dateTo,
    month,
    search,
  } = {}) {
    const where = {};

    if (projectId) where.projectId = projectId;
    if (partyId) where.partyId = partyId;
    if (nature) where.nature = nature;
    if (classification) where.classification = classification;
    if (paymentStatus) where.paymentStatus = paymentStatus;
    if (month) where.month = month;

    if (dateFrom || dateTo) {
      where.transactionDate = {};
      if (dateFrom) where.transactionDate.gte = new Date(dateFrom);
      if (dateTo) where.transactionDate.lte = new Date(dateTo);
    }

    if (search) {
      where.OR = [
        { details: { contains: search, mode: 'insensitive' } },
        { notes: { contains: search, mode: 'insensitive' } },
        { party: { name: { contains: search, mode: 'insensitive' } } },
        { project: { name: { contains: search, mode: 'insensitive' } } },
      ];
    }

    const skip = (page - 1) * limit;

    const [transactions, total] = await Promise.all([
      db().transaction.findMany({
        where,
        include: {
          project: { select: { id: true, code: true, name: true } },
          party: { select: { id: true, name: true, type: true } },
          item: { select: { id: true, name: true } },
        },
        orderBy: [{ transactionDate: 'desc' }, { id: 'desc' }],
        skip,
        take: limit,
      }),
      db().transaction.count({ where }),
    ]);

    return {
      data: transactions,
      pagination: {
        page,
        limit,
        total,
        pages: Math.ceil(total / limit),
      },
    };
  },

  // ==================== تحديث حركة ====================
  async update(id, data) {
    // Recalculate if amount or exchange rate changed
    const updateData = { ...data };
    if (data.amount !== undefined || data.exchangeRate !== undefined) {
      const existing = await this.getById(id);
      const amount = data.amount ?? Number(existing.amount);
      const rate = data.exchangeRate ?? Number(existing.exchangeRate);
      updateData.amountUsd = amount * rate;
    }
    if (data.nature) {
      updateData.movementType = getMovementType(data.nature);
    }
    if (data.transactionDate) {
      updateData.month = new Date(data.transactionDate).toISOString().slice(0, 7);
      updateData.transactionDate = new Date(data.transactionDate);
    }
    if (data.dueDate) updateData.dueDate = new Date(data.dueDate);
    if (data.paymentTermDate) updateData.paymentTermDate = new Date(data.paymentTermDate);

    return db().transaction.update({
      where: { id },
      data: updateData,
      include: { project: true, party: true, item: true },
    });
  },

  // ==================== حذف حركة ====================
  async delete(id) {
    await this.getById(id); // Verify exists
    return db().transaction.delete({ where: { id } });
  },

  // ==================== إحصائيات سريعة ====================
  async getStats({ projectId, dateFrom, dateTo } = {}) {
    const where = {};
    if (projectId) where.projectId = projectId;
    if (dateFrom || dateTo) {
      where.transactionDate = {};
      if (dateFrom) where.transactionDate.gte = new Date(dateFrom);
      if (dateTo) where.transactionDate.lte = new Date(dateTo);
    }

    const [expenses, revenue, pending] = await Promise.all([
      db().transaction.aggregate({
        where: { ...where, nature: { in: ['EXPENSE_ACCRUAL', 'EXPENSE_PAYMENT'] } },
        _sum: { amountUsd: true },
        _count: true,
      }),
      db().transaction.aggregate({
        where: { ...where, nature: { in: ['REVENUE_ACCRUAL', 'REVENUE_COLLECTION'] } },
        _sum: { amountUsd: true },
        _count: true,
      }),
      db().transaction.count({
        where: { ...where, paymentStatus: 'PENDING' },
      }),
    ]);

    return {
      totalExpenses: expenses._sum.amountUsd || 0,
      expenseCount: expenses._count,
      totalRevenue: revenue._sum.amountUsd || 0,
      revenueCount: revenue._count,
      pendingCount: pending,
    };
  },

  // ==================== كشف حساب طرف ====================
  async getPartyStatement(partyId, { dateFrom, dateTo } = {}) {
    const where = { partyId };
    if (dateFrom || dateTo) {
      where.transactionDate = {};
      if (dateFrom) where.transactionDate.gte = new Date(dateFrom);
      if (dateTo) where.transactionDate.lte = new Date(dateTo);
    }

    const transactions = await db().transaction.findMany({
      where,
      include: {
        project: { select: { code: true, name: true } },
        item: { select: { name: true } },
      },
      orderBy: [{ transactionDate: 'asc' }, { id: 'asc' }],
    });

    // Calculate running balance
    let balance = 0;
    return transactions.map((t) => {
      balance += t.movementType === 'DEBIT'
        ? Number(t.amountUsd)
        : -Number(t.amountUsd);
      return { ...t, runningBalance: balance };
    });
  },

  // ==================== التنبيهات والاستحقاقات ====================
  async getDueAlerts() {
    return db().transaction.findMany({
      where: {
        paymentStatus: 'PENDING',
        dueDate: { not: null },
      },
      include: {
        party: { select: { name: true, type: true } },
        project: { select: { name: true } },
      },
      orderBy: { dueDate: 'asc' },
    });
  },
};
