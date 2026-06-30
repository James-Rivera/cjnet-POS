export type Service = {
  id: string;
  name: string;
  category: string;
  optionLabel: string;
  price: number;
  isCustomPrice?: boolean;
  groupName?: string | null;
  requiresTracking?: boolean;
  baseFee?: number;
  serviceFee?: number;
  usesMaya?: boolean;
  mayaDeductionAmount?: number;
  mayaDeductionMode?: "fixed" | "pass_through";
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
  requiresTracking?: boolean;
  baseFee?: number;
  serviceFee?: number;
  passThroughFee?: number;
  revenueAmount?: number;
  pricingBreakdown?: PricingBreakdown | null;
  bundleId?: string | null;
  bundleLabel?: string | null;
  isUncategorizedCustom?: boolean;
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
  baseFee?: number;
  serviceFee?: number;
  passThroughFee?: number;
  revenueAmount?: number;
  pricingBreakdown?: PricingBreakdown | null;
  bundleId?: string | null;
  bundleLabel?: string | null;
  isUncategorizedCustom?: boolean;
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
  status?: "completed" | "voided";
  needsFollowUp?: boolean;
  voidedAt?: string | null;
  voidedBy?: string | null;
  voidReason?: string | null;
};

export type Expense = {
  id: string;
  date: string;
  category: string;
  description: string;
  amount: number;
  createdAt: string;
  createdBy?: string | null;
  status?: "active" | "voided";
  voidedAt?: string | null;
  voidedBy?: string | null;
  voidReason?: string | null;
};

export type DailyClosing = {
  id: string;
  closingDate: string;
  cashCounted?: boolean;
  openingCash: number;
  expectedCash: number;
  actualCash: number;
  cashDifference: number;
  walletBalance?: number | null;
  notes: string;
  summary: DailyClosingSummary;
  closedBy?: string | null;
  closedAt: string;
  createdAt: string;
  updatedAt?: string | null;
};

export type DailyClosingSummary = {
  collected: number;
  passThrough: number;
  earned: number;
  expenses: number;
  net: number;
  transactions: number;
  groupedSales?: number;
  bundledSales: number;
  uncategorizedCustom: number;
  pendingOnline: number;
  topServices: Array<{ name: string; quantity: number; total: number }>;
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
  role: "owner" | "manager" | "staff" | "admin" | "cashier";
};

export type PosSnapshot = {
  services: Service[];
  sales: Sale[];
  expenses: Expense[];
  closings: DailyClosing[];
  mayaSettings: MayaSettings;
  mayaLedger: MayaLedgerEntry[];
};

export type MayaSettings = {
  id: number;
  trackingEnabled: boolean;
  currentBalance: number;
  lowBalanceThreshold: number;
  updatedBy?: string | null;
  createdAt?: string | null;
  updatedAt?: string | null;
};

export type MayaLedgerEntry = {
  id: string;
  entryType: "sale_deduction" | "sale_reversal" | "top_up" | "adjustment";
  amount: number;
  direction: "in" | "out" | "adjustment";
  balanceAfter?: number | null;
  saleId?: string | null;
  saleItemId?: string | null;
  serviceId?: string | null;
  serviceName?: string | null;
  notes: string;
  reason: string;
  createdBy?: string | null;
  createdAt: string;
};

export type DateRange = {
  from: string;
  to: string;
};

export type ReportSummary = {
  grossSales: number;
  passThroughFees: number;
  serviceRevenue: number;
  expenses: number;
  netIncome: number;
  transactions: number;
  topServices: Array<{ name: string; quantity: number; total: number }>;
  expenseSummary: Array<{ category: string; total: number }>;
  uncategorizedCustom: number;
};

export type PricingBreakdown = {
  baseFee: number;
  serviceFee: number;
  total: number;
};
