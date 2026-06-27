import type { CartItem, DateRange, Expense, ReportSummary, Sale } from "@/types/pos";

export function money(value: number) {
  return `PHP ${Number(value || 0).toFixed(2)}`;
}

export function dateKey(date: Date | string) {
  const d = new Date(date);
  const yyyy = d.getFullYear();
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  const dd = String(d.getDate()).padStart(2, "0");
  return `${yyyy}-${mm}-${dd}`;
}

export function todayKey() {
  return dateKey(new Date());
}

export function cartSubtotal(items: CartItem[]) {
  return items.reduce((sum, item) => sum + item.price * item.quantity, 0);
}

export function saleTotal(subtotal: number, discount: number) {
  return Math.max(subtotal - Math.max(discount || 0, 0), 0);
}

export function changeDue(cashReceived: number, total: number) {
  return Math.max((cashReceived || 0) - total, 0);
}

export function inRange(date: string, range: DateRange) {
  return date >= range.from && date <= range.to;
}

export function buildReport(sales: Sale[], expenses: Expense[], range: DateRange): ReportSummary {
  const rangeSales = sales.filter((sale) => inRange(sale.date, range));
  const rangeExpenses = expenses.filter((expense) => inRange(expense.date, range));
  const grossSales = rangeSales.reduce((sum, sale) => sum + sale.total, 0);
  const expenseTotal = rangeExpenses.reduce((sum, expense) => sum + expense.amount, 0);
  const serviceMap = new Map<string, { quantity: number; total: number }>();
  const expenseMap = new Map<string, number>();

  for (const sale of rangeSales) {
    for (const item of sale.items) {
      const key = `${item.serviceName} - ${item.optionLabel}`;
      const current = serviceMap.get(key) ?? { quantity: 0, total: 0 };
      current.quantity += item.quantity;
      current.total += item.lineTotal;
      serviceMap.set(key, current);
    }
  }

  for (const expense of rangeExpenses) {
    expenseMap.set(expense.category, (expenseMap.get(expense.category) ?? 0) + expense.amount);
  }

  return {
    grossSales,
    expenses: expenseTotal,
    netIncome: grossSales - expenseTotal,
    transactions: rangeSales.length,
    topServices: [...serviceMap.entries()]
      .map(([name, value]) => ({ name, ...value }))
      .sort((a, b) => b.total - a.total)
      .slice(0, 8),
    expenseSummary: [...expenseMap.entries()]
      .map(([category, total]) => ({ category, total }))
      .sort((a, b) => b.total - a.total),
  };
}
