import { reportService } from './reports.service.js';
import { transactionService } from '../transactions/transactions.service.js';
import { generateStatementPDF, generateProjectReportPDF } from './pdf-generator.js';
import { getDB } from '../../config/database.js';
import { NotFoundError } from '../../shared/errors.js';

export async function reportRoutes(app) {
  const db = () => getDB();

  // ==================== لوحة التحكم ====================
  app.get('/dashboard', { preHandler: [app.authenticate] }, async () => {
    return reportService.getDashboard();
  });

  // ==================== التدفقات النقدية ====================
  app.get('/cash-flow', { preHandler: [app.authenticate] }, async (request) => {
    return reportService.getCashFlow(request.query);
  });

  // ==================== تقرير المشاريع ====================
  app.get('/projects', { preHandler: [app.authenticate] }, async () => {
    return reportService.getProjectsReport();
  });

  // ==================== تقرير الموردين ====================
  app.get('/vendors', { preHandler: [app.authenticate] }, async () => {
    return reportService.getVendorsReport();
  });

  // ==================== تقرير الممولين ====================
  app.get('/funders', { preHandler: [app.authenticate] }, async () => {
    return reportService.getFundersReport();
  });

  // ==================== تقرير المصروفات حسب البند ====================
  app.get('/expenses-by-item', { preHandler: [app.authenticate] }, async (request) => {
    return reportService.getExpensesByItem(request.query);
  });

  // ==================== التنبيهات ====================
  app.get('/alerts', { preHandler: [app.authenticate] }, async () => {
    return transactionService.getDueAlerts();
  });

  // ==================== PDF - كشف حساب طرف ====================
  app.get(
    '/pdf/party-statement/:partyId',
    { preHandler: [app.authenticate] },
    async (request, reply) => {
      const partyId = Number(request.params.partyId);
      const party = await db().party.findUnique({ where: { id: partyId } });
      if (!party) throw new NotFoundError('الطرف');

      const transactions = await transactionService.getPartyStatement(partyId, request.query);
      const pdf = await generateStatementPDF(party, transactions);

      reply.header('Content-Type', 'application/pdf');
      reply.header('Content-Disposition', `attachment; filename="statement-${party.name}.pdf"`);
      return reply.send(pdf);
    }
  );

  // ==================== PDF - تقرير مشروع ====================
  app.get(
    '/pdf/project/:projectId',
    { preHandler: [app.authenticate] },
    async (request, reply) => {
      const projectId = Number(request.params.projectId);
      const project = await db().project.findUnique({
        where: { id: projectId },
        include: { budgets: true },
      });
      if (!project) throw new NotFoundError('المشروع');

      const transactions = await db().transaction.findMany({
        where: { projectId },
        include: {
          item: { select: { name: true } },
          party: { select: { name: true } },
        },
        orderBy: { transactionDate: 'asc' },
      });

      const pdf = await generateProjectReportPDF(project, transactions, project.budgets);

      reply.header('Content-Type', 'application/pdf');
      reply.header('Content-Disposition', `attachment; filename="project-${project.code}.pdf"`);
      return reply.send(pdf);
    }
  );
}
