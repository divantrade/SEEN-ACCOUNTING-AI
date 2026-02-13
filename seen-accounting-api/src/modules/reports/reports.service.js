import { getDB } from '../../config/database.js';

const db = () => getDB();

export const reportService = {
  // ==================== لوحة التحكم (Dashboard) ====================
  async getDashboard() {
    const [
      activeProjects,
      activeParties,
      totalTransactions,
      expenseSum,
      revenueSum,
      pendingReviews,
      overdueCount,
    ] = await Promise.all([
      db().project.count({ where: { status: 'نشط' } }),
      db().party.count({ where: { isActive: true } }),
      db().transaction.count(),
      db().transaction.aggregate({
        where: { nature: { in: ['EXPENSE_ACCRUAL', 'EXPENSE_PAYMENT'] } },
        _sum: { amountUsd: true },
      }),
      db().transaction.aggregate({
        where: { nature: { in: ['REVENUE_ACCRUAL', 'REVENUE_COLLECTION'] } },
        _sum: { amountUsd: true },
      }),
      db().pendingTransaction.count({ where: { reviewStatus: 'PENDING' } }),
      db().transaction.count({
        where: {
          paymentStatus: 'PENDING',
          dueDate: { lt: new Date() },
        },
      }),
    ]);

    return {
      activeProjects,
      activeParties,
      totalTransactions,
      totalExpenses: expenseSum._sum.amountUsd || 0,
      totalRevenue: revenueSum._sum.amountUsd || 0,
      netProfit: (revenueSum._sum.amountUsd || 0) - (expenseSum._sum.amountUsd || 0),
      pendingReviews,
      overdueCount,
    };
  },

  // ==================== تقرير التدفقات النقدية ====================
  async getCashFlow({ dateFrom, dateTo } = {}) {
    const where = {};
    if (dateFrom || dateTo) {
      where.transactionDate = {};
      if (dateFrom) where.transactionDate.gte = new Date(dateFrom);
      if (dateTo) where.transactionDate.lte = new Date(dateTo);
    }

    const transactions = await db().transaction.findMany({
      where,
      select: {
        month: true,
        nature: true,
        amountUsd: true,
      },
      orderBy: { transactionDate: 'asc' },
    });

    // Group by month
    const monthly = {};
    for (const t of transactions) {
      if (!t.month) continue;
      if (!monthly[t.month]) {
        monthly[t.month] = { month: t.month, cashIn: 0, cashOut: 0 };
      }

      const amount = Number(t.amountUsd || 0);
      const isInflow = ['REVENUE_COLLECTION', 'FINANCING_RECEIVED'].includes(t.nature);
      const isOutflow = ['EXPENSE_PAYMENT', 'FINANCING_REPAYMENT'].includes(t.nature);

      if (isInflow) monthly[t.month].cashIn += amount;
      if (isOutflow) monthly[t.month].cashOut += amount;
    }

    return Object.values(monthly).map((m) => ({
      ...m,
      netCashFlow: m.cashIn - m.cashOut,
    }));
  },

  // ==================== تقرير المشاريع ====================
  async getProjectsReport() {
    const projects = await db().project.findMany({
      where: { status: 'نشط' },
      include: {
        _count: { select: { transactions: true } },
      },
    });

    const result = [];
    for (const project of projects) {
      const [expenses, revenue] = await Promise.all([
        db().transaction.aggregate({
          where: { projectId: project.id, nature: { in: ['EXPENSE_ACCRUAL', 'EXPENSE_PAYMENT'] } },
          _sum: { amountUsd: true },
        }),
        db().transaction.aggregate({
          where: { projectId: project.id, nature: { in: ['REVENUE_ACCRUAL', 'REVENUE_COLLECTION'] } },
          _sum: { amountUsd: true },
        }),
      ]);

      result.push({
        id: project.id,
        code: project.code,
        name: project.name,
        transactionCount: project._count.transactions,
        totalExpenses: expenses._sum.amountUsd || 0,
        totalRevenue: revenue._sum.amountUsd || 0,
        netProfit: (revenue._sum.amountUsd || 0) - (expenses._sum.amountUsd || 0),
      });
    }

    return result;
  },

  // ==================== تقرير الموردين ====================
  async getVendorsReport() {
    const vendors = await db().party.findMany({
      where: { type: 'VENDOR', isActive: true },
      include: {
        _count: { select: { transactions: true } },
      },
    });

    const result = [];
    for (const vendor of vendors) {
      const totals = await db().transaction.aggregate({
        where: { partyId: vendor.id },
        _sum: { amountUsd: true },
      });

      const pending = await db().transaction.count({
        where: { partyId: vendor.id, paymentStatus: 'PENDING' },
      });

      result.push({
        id: vendor.id,
        name: vendor.name,
        specialization: vendor.specialization,
        transactionCount: vendor._count.transactions,
        totalAmountUsd: totals._sum.amountUsd || 0,
        pendingCount: pending,
      });
    }

    return result.sort((a, b) => Number(b.totalAmountUsd) - Number(a.totalAmountUsd));
  },

  // ==================== تقرير الممولين ====================
  async getFundersReport() {
    const funders = await db().party.findMany({
      where: { type: 'FUNDER', isActive: true },
      include: {
        _count: { select: { transactions: true } },
      },
    });

    const result = [];
    for (const funder of funders) {
      const received = await db().transaction.aggregate({
        where: { partyId: funder.id, nature: 'FINANCING_RECEIVED' },
        _sum: { amountUsd: true },
      });

      const repaid = await db().transaction.aggregate({
        where: { partyId: funder.id, nature: 'FINANCING_REPAYMENT' },
        _sum: { amountUsd: true },
      });

      result.push({
        id: funder.id,
        name: funder.name,
        transactionCount: funder._count.transactions,
        totalReceived: received._sum.amountUsd || 0,
        totalRepaid: repaid._sum.amountUsd || 0,
        balance: (received._sum.amountUsd || 0) - (repaid._sum.amountUsd || 0),
      });
    }

    return result;
  },

  // ==================== تقرير المصروفات حسب البند ====================
  async getExpensesByItem({ projectId, dateFrom, dateTo } = {}) {
    const where = { nature: { in: ['EXPENSE_ACCRUAL', 'EXPENSE_PAYMENT'] } };
    if (projectId) where.projectId = projectId;
    if (dateFrom || dateTo) {
      where.transactionDate = {};
      if (dateFrom) where.transactionDate.gte = new Date(dateFrom);
      if (dateTo) where.transactionDate.lte = new Date(dateTo);
    }

    const transactions = await db().transaction.findMany({
      where,
      include: { item: { select: { name: true } } },
    });

    const byItem = {};
    for (const t of transactions) {
      const itemName = t.item?.name || 'غير محدد';
      if (!byItem[itemName]) {
        byItem[itemName] = { item: itemName, count: 0, totalUsd: 0 };
      }
      byItem[itemName].count++;
      byItem[itemName].totalUsd += Number(t.amountUsd || 0);
    }

    return Object.values(byItem).sort((a, b) => b.totalUsd - a.totalUsd);
  },
};
