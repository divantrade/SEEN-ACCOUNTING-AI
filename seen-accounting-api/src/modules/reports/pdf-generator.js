import PDFDocument from 'pdfkit';
import { CURRENCIES } from '../../config/constants.js';

/**
 * Generate a party statement PDF
 */
export function generateStatementPDF(party, transactions, options = {}) {
  return new Promise((resolve, reject) => {
    const doc = new PDFDocument({
      size: 'A4',
      layout: 'landscape',
      margins: { top: 40, bottom: 40, left: 40, right: 40 },
    });

    const chunks = [];
    doc.on('data', (chunk) => chunks.push(chunk));
    doc.on('end', () => resolve(Buffer.concat(chunks)));
    doc.on('error', reject);

    // Header
    doc.fontSize(18).text('SEEN Accounting System', { align: 'center' });
    doc.fontSize(14).text(`كشف حساب: ${party.name}`, { align: 'center' });
    doc.fontSize(10).text(`النوع: ${party.type} | التاريخ: ${new Date().toLocaleDateString('ar-EG')}`, {
      align: 'center',
    });
    doc.moveDown();

    // Summary
    if (transactions.length > 0) {
      const lastBalance = transactions[transactions.length - 1].runningBalance;
      doc.fontSize(12).text(`الرصيد: $${Number(lastBalance).toLocaleString()}`, { align: 'right' });
      doc.moveDown();
    }

    // Table header
    const tableTop = doc.y;
    const colWidths = [60, 70, 100, 120, 80, 80, 80, 80];
    const headers = ['#', 'التاريخ', 'الطبيعة', 'التفاصيل', 'المبلغ', 'العملة', 'بالدولار', 'الرصيد'];

    let x = 40;
    doc.fontSize(9).fillColor('#333');
    headers.forEach((header, i) => {
      doc.text(header, x, tableTop, { width: colWidths[i], align: 'center' });
      x += colWidths[i];
    });

    doc.moveTo(40, tableTop + 15).lineTo(760, tableTop + 15).stroke();

    // Table rows
    let y = tableTop + 20;
    transactions.forEach((t, idx) => {
      if (y > 540) {
        doc.addPage();
        y = 40;
      }

      x = 40;
      const row = [
        idx + 1,
        new Date(t.transactionDate).toLocaleDateString('en-GB'),
        t.nature || '',
        (t.details || '').substring(0, 25),
        Number(t.amount).toLocaleString(),
        t.currency,
        `$${Number(t.amountUsd).toLocaleString()}`,
        `$${Number(t.runningBalance).toLocaleString()}`,
      ];

      doc.fontSize(8).fillColor('#000');
      row.forEach((cell, i) => {
        doc.text(String(cell), x, y, { width: colWidths[i], align: 'center' });
        x += colWidths[i];
      });

      y += 15;
    });

    // Footer
    doc.fontSize(8).fillColor('#888');
    doc.text(
      `تم إنشاء هذا التقرير بواسطة SEEN Accounting System - ${new Date().toISOString()}`,
      40,
      560,
      { align: 'center' }
    );

    doc.end();
  });
}

/**
 * Generate a project report PDF
 */
export function generateProjectReportPDF(project, transactions, budget) {
  return new Promise((resolve, reject) => {
    const doc = new PDFDocument({
      size: 'A4',
      layout: 'landscape',
      margins: { top: 40, bottom: 40, left: 40, right: 40 },
    });

    const chunks = [];
    doc.on('data', (chunk) => chunks.push(chunk));
    doc.on('end', () => resolve(Buffer.concat(chunks)));
    doc.on('error', reject);

    // Header
    doc.fontSize(18).text('SEEN Accounting System', { align: 'center' });
    doc.fontSize(14).text(`تقرير مشروع: ${project.name} (${project.code})`, { align: 'center' });
    doc.moveDown();

    // Summary stats
    const totalExpenses = transactions
      .filter((t) => t.nature?.includes('مصروف'))
      .reduce((sum, t) => sum + Number(t.amountUsd || 0), 0);

    const totalRevenue = transactions
      .filter((t) => t.nature?.includes('إيراد'))
      .reduce((sum, t) => sum + Number(t.amountUsd || 0), 0);

    doc.fontSize(11);
    doc.text(`إجمالي المصروفات: $${totalExpenses.toLocaleString()}`);
    doc.text(`إجمالي الإيرادات: $${totalRevenue.toLocaleString()}`);
    doc.text(`صافي الربح: $${(totalRevenue - totalExpenses).toLocaleString()}`);
    doc.text(`عدد الحركات: ${transactions.length}`);
    doc.moveDown();

    // Transaction list
    const tableTop = doc.y;
    const colWidths = [40, 70, 100, 100, 100, 80, 80];
    const headers = ['#', 'التاريخ', 'البند', 'الطرف', 'التفاصيل', 'المبلغ', 'بالدولار'];

    let x = 40;
    doc.fontSize(9).fillColor('#333');
    headers.forEach((header, i) => {
      doc.text(header, x, tableTop, { width: colWidths[i], align: 'center' });
      x += colWidths[i];
    });

    doc.moveTo(40, tableTop + 15).lineTo(760, tableTop + 15).stroke();

    let y = tableTop + 20;
    transactions.forEach((t, idx) => {
      if (y > 540) {
        doc.addPage();
        y = 40;
      }

      x = 40;
      const row = [
        idx + 1,
        new Date(t.transactionDate).toLocaleDateString('en-GB'),
        t.item?.name || '',
        t.party?.name || '',
        (t.details || '').substring(0, 20),
        `${Number(t.amount).toLocaleString()} ${t.currency}`,
        `$${Number(t.amountUsd).toLocaleString()}`,
      ];

      doc.fontSize(8).fillColor('#000');
      row.forEach((cell, i) => {
        doc.text(String(cell), x, y, { width: colWidths[i], align: 'center' });
        x += colWidths[i];
      });

      y += 15;
    });

    doc.end();
  });
}
