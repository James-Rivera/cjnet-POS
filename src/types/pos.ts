export type Service = {
  id: string;
  name: string;
  category: string;
  optionLabel: string;
  price: number;
  isCustomPrice?: boolean;
  sortOrder?: number;
  isActive?: boolean;
};

export type CartItem = {
  id: string;
  serviceId: string;
  name: string;
  category: string;
  optionLabel: string;
  price: number;
  quantity: number;
};

export type SaleItem = {
  id: string;
  saleId?: string;
  serviceId?: string;
  serviceName: string;
  category: string;
  optionLabel: string;
  quantity: number;
  unitPrice: number;
  lineTotal: number;
};

export type Sale = {
  id: string;
  createdAt: string;
  soldAt: string;
  date: string;
  customerNote: string;
  subtotal: number;
  discount: number;
  total: number;
  cashReceived: number;
  changeDue: number;
  items: SaleItem[];
  cashierId?: string | null;
};

export type Expense = {
  id: string;
  date: string;
  category: string;
  description: string;
  amount: number;
  createdAt: string;
  createdBy?: string | null;
};

export type Customer = {
  id: string;
  name: string;
  reference?: string | null;
  createdAt: string;
};

export type CashierUser = {
  id: string;
  displayName: string;
  role: "owner" | "admin" | "cashier";
};

export type PosSnapshot = {
  services: Service[];
  sales: Sale[];
  expenses: Expense[];
};

export type DateRange = {
  from: string;
  to: string;
};

export type ReportSummary = {
  grossSales: number;
  expenses: number;
  netIncome: number;
  transactions: number;
  topServices: Array<{ name: string; quantity: number; total: number }>;
  expenseSummary: Array<{ category: string; total: number }>;
};
