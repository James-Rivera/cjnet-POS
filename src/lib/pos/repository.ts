"use client";

import type { CartItem, DailyClosing, Expense, MayaLedgerEntry, MayaSettings, PosSnapshot, Sale, SaleItem, Service } from "@/types/pos";
import { DEFAULT_SERVICES } from "@/lib/pos/defaults";
import { changeDue, dateKey, todayKey } from "@/lib/pos/calculations";
import { createSupabaseBrowserClient } from "@/lib/pos/supabase-client";
import { formatPosError } from "@/lib/pos/errors";

const STORAGE_KEY = "cjnet_pos_next_snapshot";

type SaveSaleInput = {
  cart: CartItem[];
  discount: number;
  cashReceived: number;
  customerNote: string;
  needsFollowUp?: boolean;
};

type SaveExpenseInput = Omit<Expense, "id" | "createdAt">;
type SaveDailyClosingInput = Omit<DailyClosing, "id" | "closedBy" | "closedAt" | "createdAt" | "updatedAt">;
type SaveMayaSettingsInput = Pick<MayaSettings, "trackingEnabled" | "currentBalance" | "lowBalanceThreshold">;

const DEFAULT_MAYA_SETTINGS: MayaSettings = {
  id: 1,
  trackingEnabled: false,
  currentBalance: 0,
  lowBalanceThreshold: 500,
  updatedBy: null,
  createdAt: null,
  updatedAt: null,
};

function uid() {
  const cryptoApi = typeof globalThis !== "undefined" ? globalThis.crypto : undefined;
  if (cryptoApi?.randomUUID) return cryptoApi.randomUUID();

  if (cryptoApi?.getRandomValues) {
    const bytes = cryptoApi.getRandomValues(new Uint8Array(16));
    bytes[6] = (bytes[6] & 0x0f) | 0x40;
    bytes[8] = (bytes[8] & 0x3f) | 0x80;
    const hex = [...bytes].map((byte) => byte.toString(16).padStart(2, "0"));
    return `${hex.slice(0, 4).join("")}-${hex.slice(4, 6).join("")}-${hex.slice(6, 8).join("")}-${hex.slice(8, 10).join("")}-${hex.slice(10, 16).join("")}`;
  }

  return "xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx".replace(/[xy]/g, (char) => {
    const random = Math.floor(Math.random() * 16);
    const value = char === "x" ? random : (random & 0x3) | 0x8;
    return value.toString(16);
  });
}

function emptySnapshot(): PosSnapshot {
  return { services: DEFAULT_SERVICES, sales: [], expenses: [], closings: [], mayaSettings: DEFAULT_MAYA_SETTINGS, mayaLedger: [] };
}

function withRequiredDefaultServices(services: Service[]) {
  const existingIds = new Set(services.map((service) => service.id));
  const requiredServices = DEFAULT_SERVICES.filter((service) => ["police-clearance-custom", "psa-custom", "gcash-cash-in"].includes(service.id) && !existingIds.has(service.id));
  return [...services, ...requiredServices].sort((a, b) => (a.sortOrder ?? 100) - (b.sortOrder ?? 100));
}

function readLocalSnapshot(): PosSnapshot {
  if (typeof window === "undefined") return emptySnapshot();
  const raw = localStorage.getItem(STORAGE_KEY);
  if (!raw) {
    const initial = emptySnapshot();
    localStorage.setItem(STORAGE_KEY, JSON.stringify(initial));
    return initial;
  }

  try {
    const parsed = JSON.parse(raw) as PosSnapshot;
    const services = withRequiredDefaultServices(parsed.services?.length ? parsed.services : DEFAULT_SERVICES);
    const snapshot = {
      services,
      sales: parsed.sales ?? [],
      expenses: parsed.expenses ?? [],
      closings: parsed.closings ?? [],
      mayaSettings: parsed.mayaSettings ?? DEFAULT_MAYA_SETTINGS,
      mayaLedger: parsed.mayaLedger ?? [],
    };
    if (services.length !== (parsed.services?.length ?? 0)) {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(snapshot));
    }
    return {
      services: snapshot.services,
      sales: snapshot.sales,
      expenses: snapshot.expenses,
      closings: snapshot.closings,
      mayaSettings: snapshot.mayaSettings,
      mayaLedger: snapshot.mayaLedger,
    };
  } catch {
    const initial = emptySnapshot();
    localStorage.setItem(STORAGE_KEY, JSON.stringify(initial));
    return initial;
  }
}

function writeLocalSnapshot(snapshot: PosSnapshot) {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(snapshot));
}

function categoryId(category: string) {
  const known: Record<string, string> = {
    Xerox: "xerox",
    Printing: "printing",
    "Online Services": "online-services",
    GCash: "gcash",
    Finishing: "finishing",
    Custom: "custom",
  };
  return known[category] ?? (category.trim().toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/(^-|-$)/g, "") || "custom");
}

function numericValue(value: unknown) {
  return Number(value ?? 0);
}

function isMissingOptionalSaleItemColumn(error: unknown) {
  if (!error || typeof error !== "object") return false;
  const candidate = error as { code?: unknown; message?: unknown; details?: unknown };
  const text = `${candidate.message ?? ""} ${candidate.details ?? ""}`.toLowerCase();
  return candidate.code === "PGRST204" && (
    text.includes("bundle_id") ||
    text.includes("bundle_label") ||
    text.includes("is_uncategorized_custom")
  );
}

function isMissingMayaSchema(error: unknown) {
  if (!error || typeof error !== "object") return false;
  const candidate = error as { code?: unknown; message?: unknown; details?: unknown };
  const text = `${candidate.message ?? ""} ${candidate.details ?? ""}`.toLowerCase();
  return (candidate.code === "PGRST205" || candidate.code === "42P01") && (
    text.includes("maya_settings") ||
    text.includes("maya_ledger_entries")
  );
}

function isMissingMayaServiceColumn(error: unknown) {
  if (!error || typeof error !== "object") return false;
  const candidate = error as { code?: unknown; message?: unknown; details?: unknown };
  const text = `${candidate.message ?? ""} ${candidate.details ?? ""}`.toLowerCase();
  return (candidate.code === "PGRST204" || candidate.code === "42703") && (
    text.includes("uses_maya") ||
    text.includes("maya_deduction_amount") ||
    text.includes("maya_deduction_mode")
  );
}

function assertMayaSchema(error: unknown) {
  if (isMissingMayaSchema(error)) {
    throw new Error("Maya tracking tables are not installed yet. Run migration 202606300003_maya_ledger.sql, then reload.");
  }
  assertNoError(error, "Could not save Maya data.");
}

function toService(row: Record<string, unknown>): Service {
  return {
    id: String(row.id),
    name: String(row.name),
    category: String(row.category),
    optionLabel: String(row.option_label),
    price: Number(row.price ?? 0),
    isCustomPrice: Boolean(row.is_custom_price),
    groupName: row.group_name ? String(row.group_name) : null,
    requiresTracking: Boolean(row.requires_tracking),
    baseFee: numericValue(row.base_fee),
    serviceFee: numericValue(row.service_fee),
    usesMaya: Boolean(row.uses_maya),
    mayaDeductionAmount: numericValue(row.maya_deduction_amount),
    mayaDeductionMode: String(row.maya_deduction_mode ?? "pass_through") === "fixed" ? "fixed" : "pass_through",
    isActive: Boolean(row.is_active ?? true),
    sortOrder: Number(row.sort_order ?? 100),
  };
}

function toExpense(row: Record<string, unknown>): Expense {
  return {
    id: String(row.id),
    date: String(row.expense_date),
    category: String(row.category),
    description: String(row.description ?? ""),
    amount: Number(row.amount ?? 0),
    createdAt: String(row.created_at),
    createdBy: row.created_by ? String(row.created_by) : null,
    status: String(row.status ?? "active") === "voided" ? "voided" : "active",
    voidedAt: row.voided_at ? String(row.voided_at) : null,
    voidedBy: row.voided_by ? String(row.voided_by) : null,
    voidReason: row.void_reason ? String(row.void_reason) : null,
  };
}

function toDailyClosing(row: Record<string, unknown>): DailyClosing {
  return {
    id: String(row.id),
    closingDate: String(row.closing_date),
    cashCounted: Boolean(row.cash_counted ?? true),
    openingCash: numericValue(row.opening_cash),
    expectedCash: numericValue(row.expected_cash),
    actualCash: numericValue(row.actual_cash),
    cashDifference: numericValue(row.cash_difference),
    walletBalance: row.wallet_balance === null || row.wallet_balance === undefined ? null : numericValue(row.wallet_balance),
    notes: String(row.notes ?? ""),
    summary: row.summary && typeof row.summary === "object"
      ? row.summary as DailyClosing["summary"]
      : {
        collected: 0,
        passThrough: 0,
        earned: 0,
        expenses: 0,
        net: 0,
        transactions: 0,
        groupedSales: 0,
        bundledSales: 0,
        uncategorizedCustom: 0,
        pendingOnline: 0,
        topServices: [],
      },
    closedBy: row.closed_by ? String(row.closed_by) : null,
    closedAt: String(row.closed_at),
    createdAt: String(row.created_at),
    updatedAt: row.updated_at ? String(row.updated_at) : null,
  };
}

function toMayaSettings(row: Record<string, unknown> | null | undefined): MayaSettings {
  if (!row) return DEFAULT_MAYA_SETTINGS;
  return {
    id: Number(row.id ?? 1),
    trackingEnabled: Boolean(row.tracking_enabled),
    currentBalance: numericValue(row.current_balance),
    lowBalanceThreshold: numericValue(row.low_balance_threshold || 500),
    updatedBy: row.updated_by ? String(row.updated_by) : null,
    createdAt: row.created_at ? String(row.created_at) : null,
    updatedAt: row.updated_at ? String(row.updated_at) : null,
  };
}

function toMayaLedgerEntry(row: Record<string, unknown>): MayaLedgerEntry {
  const service = row.services && typeof row.services === "object" ? row.services as Record<string, unknown> : null;
  return {
    id: String(row.id),
    entryType: String(row.entry_type) as MayaLedgerEntry["entryType"],
    amount: numericValue(row.amount),
    direction: String(row.direction) as MayaLedgerEntry["direction"],
    balanceAfter: row.balance_after === null || row.balance_after === undefined ? null : numericValue(row.balance_after),
    saleId: row.sale_id ? String(row.sale_id) : null,
    saleItemId: row.sale_item_id ? String(row.sale_item_id) : null,
    serviceId: row.service_id ? String(row.service_id) : null,
    serviceName: service?.name ? String(service.name) : null,
    notes: String(row.notes ?? ""),
    reason: String(row.reason ?? ""),
    createdBy: row.created_by ? String(row.created_by) : null,
    createdAt: String(row.created_at),
  };
}

function toSale(row: Record<string, unknown>): Sale {
  const items = ((row.sale_items as Record<string, unknown>[] | null) ?? []).map((item): SaleItem => ({
    id: String(item.id),
    saleId: String(item.sale_id),
    serviceId: item.service_id ? String(item.service_id) : undefined,
    serviceName: String(item.service_name),
    category: String(item.category),
    optionLabel: String(item.option_label),
    quantity: Number(item.quantity ?? 0),
    unitPrice: Number(item.unit_price ?? 0),
    lineTotal: Number(item.line_total ?? 0),
    baseFee: numericValue(item.base_fee),
    serviceFee: numericValue(item.service_fee),
    passThroughFee: numericValue(item.pass_through_fee),
    revenueAmount: numericValue(item.revenue_amount),
    pricingBreakdown: item.pricing_breakdown && typeof item.pricing_breakdown === "object" ? item.pricing_breakdown as SaleItem["pricingBreakdown"] : null,
    bundleId: item.bundle_id ? String(item.bundle_id) : null,
    bundleLabel: item.bundle_label ? String(item.bundle_label) : null,
    isUncategorizedCustom: Boolean(item.is_uncategorized_custom),
  }));

  return {
    id: String(row.id),
    createdAt: String(row.created_at),
    soldAt: String(row.sold_at),
    date: dateKey(String(row.sold_at)),
    customerNote: String(row.customer_note ?? ""),
    subtotal: Number(row.subtotal ?? 0),
    discount: Number(row.discount ?? 0),
    total: Number(row.total ?? 0),
    cashReceived: Number(row.cash_received ?? 0),
    changeDue: Number(row.change_due ?? 0),
    cashierId: row.cashier_id ? String(row.cashier_id) : null,
    status: String(row.status ?? "completed") === "voided" ? "voided" : "completed",
    needsFollowUp: Boolean(row.needs_follow_up),
    voidedAt: row.voided_at ? String(row.voided_at) : null,
    voidedBy: row.voided_by ? String(row.voided_by) : null,
    voidReason: row.void_reason ? String(row.void_reason) : null,
    items,
  };
}

function saleItemFromCart(item: CartItem): SaleItem {
  const baseFee = item.passThroughFee ?? item.baseFee ?? 0;
  const serviceFee = item.serviceFee ?? item.revenueAmount ?? Math.max(item.price - baseFee, 0);
  const lineBaseFee = baseFee * item.quantity;
  const lineServiceFee = serviceFee * item.quantity;
  const lineTotal = item.price * item.quantity;

  return {
    id: uid(),
    serviceId: item.serviceId,
    serviceName: item.name,
    category: item.category,
    optionLabel: item.optionLabel,
    quantity: item.quantity,
    unitPrice: item.price,
    lineTotal,
    baseFee: lineBaseFee,
    serviceFee: lineServiceFee,
    passThroughFee: lineBaseFee,
    revenueAmount: lineServiceFee,
    pricingBreakdown: baseFee || serviceFee ? { baseFee: lineBaseFee, serviceFee: lineServiceFee, total: lineTotal } : null,
    bundleId: item.bundleId ?? null,
    bundleLabel: item.bundleLabel ?? null,
    isUncategorizedCustom: Boolean(item.isUncategorizedCustom),
  };
}

function assertNoError(error: unknown, fallback: string) {
  if (error) throw new Error(formatPosError(error, fallback));
}

function errorCode(error: unknown) {
  return error && typeof error === "object" && "code" in error ? String((error as { code?: unknown }).code) : "";
}

function uniqueServiceIds(items: SaleItem[]) {
  return [...new Set(items.map((item) => item.serviceId).filter((id): id is string => Boolean(id)))];
}

function mayaEntryEffect(entry: Pick<MayaLedgerEntry, "amount" | "direction">) {
  if (entry.direction === "in") return Math.abs(entry.amount);
  if (entry.direction === "out") return -Math.abs(entry.amount);
  return entry.amount;
}

function expectedMayaBalance(settings: MayaSettings, ledger: MayaLedgerEntry[]) {
  return ledger.reduce((balance, entry) => balance + mayaEntryEffect(entry), settings.currentBalance);
}

function mayaDeductionForItem(item: SaleItem, service: Service) {
  if (!service.usesMaya) return 0;
  return Math.max(item.passThroughFee ?? item.baseFee ?? 0, 0);
}

export class PosRepository {
  private supabase = createSupabaseBrowserClient();

  get mode() {
    return this.supabase ? "supabase" : "local";
  }

  async loadSnapshot(): Promise<PosSnapshot> {
    if (!this.supabase) return readLocalSnapshot();

    const [servicesResult, salesResult, expensesResult, closingsResult, mayaSettingsResult, mayaLedgerResult] = await Promise.all([
      this.supabase.from("services").select("*").eq("is_active", true).order("sort_order", { ascending: true }),
      this.supabase.from("sales").select("*, sale_items(*)").order("sold_at", { ascending: false }),
      this.supabase.from("expenses").select("*").order("expense_date", { ascending: false }),
      this.supabase.from("daily_closings").select("*").order("closing_date", { ascending: false }),
      this.supabase.from("maya_settings").select("*").eq("id", 1).maybeSingle(),
      this.supabase.from("maya_ledger_entries").select("*, services(name)").order("created_at", { ascending: false }),
    ]);

    assertNoError(servicesResult.error, "Could not load services.");
    assertNoError(salesResult.error, "Could not load sales.");
    assertNoError(expensesResult.error, "Could not load expenses.");
    assertNoError(closingsResult.error, "Could not load daily closings.");
    if (mayaSettingsResult.error && !isMissingMayaSchema(mayaSettingsResult.error)) {
      assertNoError(mayaSettingsResult.error, "Could not load Maya settings.");
    }
    if (mayaLedgerResult.error && !isMissingMayaSchema(mayaLedgerResult.error)) {
      assertNoError(mayaLedgerResult.error, "Could not load Maya ledger.");
    }

    return {
      services: servicesResult.data?.map(toService) ?? DEFAULT_SERVICES,
      sales: salesResult.data?.map(toSale) ?? [],
      expenses: expensesResult.data?.map(toExpense) ?? [],
      closings: closingsResult.data?.map(toDailyClosing) ?? [],
      mayaSettings: mayaSettingsResult.error ? DEFAULT_MAYA_SETTINGS : toMayaSettings(mayaSettingsResult.data),
      mayaLedger: mayaLedgerResult.error ? [] : mayaLedgerResult.data?.map(toMayaLedgerEntry) ?? [],
    };
  }

  async saveSale(input: SaveSaleInput): Promise<Sale> {
    const subtotal = input.cart.reduce((sum, item) => sum + item.price * item.quantity, 0);
    const discount = Math.max(input.discount || 0, 0);
    const total = Math.max(subtotal - discount, 0);
    const cash = Math.max(input.cashReceived || 0, 0);
    const now = new Date().toISOString();
    const userId = this.supabase ? await this.getCurrentUserId() : null;
    const saleItems: SaleItem[] = input.cart.map(saleItemFromCart);

    if (this.supabase) {
      if (!userId) throw new Error("Please sign in again.");
      const requestedServiceIds = uniqueServiceIds(saleItems);
      const validServiceIds = new Set<string>();
      const serviceConfig = new Map<string, Service>();
      if (requestedServiceIds.length) {
        const servicesResult = await this.supabase.from("services").select("*").in("id", requestedServiceIds);
        assertNoError(servicesResult.error, "Could not verify service records before saving the sale.");
        for (const service of servicesResult.data ?? []) {
          const mapped = toService(service);
          validServiceIds.add(mapped.id);
          serviceConfig.set(mapped.id, mapped);
        }
      }

      const saleId = uid();
      const saleInsert = {
        id: saleId,
        customer_note: input.customerNote,
        needs_follow_up: Boolean(input.needsFollowUp),
        subtotal,
        discount,
        total,
        cash_received: cash,
        change_due: changeDue(cash, total),
        sold_at: now,
        cashier_id: userId,
        status: "completed",
      };
      const saleResult = await this.supabase.from("sales").insert(saleInsert);
      assertNoError(saleResult.error, "Could not save sale.");

      const itemRows = saleItems.map((item) => ({
        id: item.id,
        sale_id: saleId,
        service_id: item.serviceId && validServiceIds.has(item.serviceId) ? item.serviceId : null,
        service_name: item.serviceName,
        category: item.category,
        option_label: item.optionLabel,
        quantity: item.quantity,
        unit_price: item.unitPrice,
        line_total: item.lineTotal,
        base_fee: item.baseFee ?? 0,
        service_fee: item.serviceFee ?? 0,
        pass_through_fee: item.passThroughFee ?? 0,
        revenue_amount: item.revenueAmount ?? item.lineTotal,
        pricing_breakdown: item.pricingBreakdown,
        bundle_id: item.bundleId ?? null,
        bundle_label: item.bundleLabel ?? null,
        is_uncategorized_custom: Boolean(item.isUncategorizedCustom),
      }));
      const itemsResult = await this.supabase.from("sale_items").insert(itemRows);
      if (isMissingOptionalSaleItemColumn(itemsResult.error)) {
        const compatibleRows = itemRows.map((item) => ({
          id: item.id,
          sale_id: item.sale_id,
          service_id: item.service_id,
          service_name: item.service_name,
          category: item.category,
          option_label: item.option_label,
          quantity: item.quantity,
          unit_price: item.unit_price,
          line_total: item.line_total,
          base_fee: item.base_fee,
          service_fee: item.service_fee,
          pass_through_fee: item.pass_through_fee,
          revenue_amount: item.revenue_amount,
          pricing_breakdown: item.pricing_breakdown,
        }));
        const compatibleResult = await this.supabase.from("sale_items").insert(compatibleRows);
        assertNoError(compatibleResult.error, "Sale was created, but line items could not be saved.");
      } else {
        assertNoError(itemsResult.error, "Sale was created, but line items could not be saved.");
      }
      await this.recordMayaSaleDeductions({
        id: saleId,
        createdAt: now,
        soldAt: now,
        date: todayKey(),
        customerNote: input.customerNote,
        subtotal,
        discount,
        total,
        cashReceived: cash,
        changeDue: changeDue(cash, total),
        cashierId: userId,
        status: "completed",
        needsFollowUp: Boolean(input.needsFollowUp),
        items: saleItems.map((item) => ({ ...item, saleId })),
      }, serviceConfig);
      return {
        id: saleId,
        createdAt: now,
        soldAt: now,
        date: todayKey(),
        customerNote: input.customerNote,
        subtotal,
        discount,
        total,
        cashReceived: cash,
        changeDue: changeDue(cash, total),
        cashierId: userId,
        status: "completed",
        needsFollowUp: Boolean(input.needsFollowUp),
        items: saleItems.map((item) => ({ ...item, saleId })),
      };
    }

    const sale: Sale = {
      id: uid(),
      createdAt: now,
      soldAt: now,
      date: todayKey(),
      customerNote: input.customerNote,
      subtotal,
      discount,
      total,
      cashReceived: cash,
      changeDue: changeDue(cash, total),
      status: "completed",
      needsFollowUp: Boolean(input.needsFollowUp),
      items: saleItems.map((item) => ({ ...item, saleId: undefined })),
    };
    const snapshot = readLocalSnapshot();
    const serviceConfig = new Map(snapshot.services.map((service) => [service.id, service]));
    const mayaEntries = this.buildMayaSaleDeductionEntries(sale, serviceConfig, snapshot.mayaSettings, snapshot.mayaLedger, null);
    writeLocalSnapshot({ ...snapshot, sales: [sale, ...snapshot.sales], mayaLedger: [...mayaEntries, ...snapshot.mayaLedger] });
    return sale;
  }

  async deleteSale(id: string) {
    await this.recordMayaSaleReversal(id, "Sale deleted");
    if (this.supabase) {
      const result = await this.supabase.from("sales").delete().eq("id", id);
      assertNoError(result.error, "Could not delete sale.");
      return;
    }
    const snapshot = readLocalSnapshot();
    writeLocalSnapshot({ ...snapshot, sales: snapshot.sales.filter((sale) => sale.id !== id) });
  }

  async voidSale(id: string, reason: string): Promise<void> {
    if (!reason.trim()) throw new Error("Enter a void reason.");
    if (this.supabase) {
      const result = await this.supabase.rpc("void_sale", { sale_id: id, reason: reason.trim() });
      assertNoError(result.error, "Could not void sale.");
      await this.recordMayaSaleReversal(id, `Sale voided: ${reason.trim()}`);
      return;
    }
    const now = new Date().toISOString();
    const snapshot = readLocalSnapshot();
    const reversalEntries = this.buildMayaSaleReversalEntries(id, `Sale voided: ${reason.trim()}`, snapshot.mayaSettings, snapshot.mayaLedger, null);
    writeLocalSnapshot({
      ...snapshot,
      mayaLedger: [...reversalEntries, ...snapshot.mayaLedger],
      sales: snapshot.sales.map((sale) => sale.id === id ? { ...sale, status: "voided", voidedAt: now, voidReason: reason.trim() } : sale),
    });
  }

  async saveExpense(input: SaveExpenseInput): Promise<Expense> {
    const now = new Date().toISOString();
    if (this.supabase) {
      const userId = await this.getCurrentUserId();
      if (!userId) throw new Error("Please sign in again.");
      const id = uid();
      const result = await this.supabase
        .from("expenses")
        .insert({
          id,
          expense_date: input.date,
          category: input.category,
          description: input.description,
          amount: input.amount,
          created_by: userId,
          updated_by: userId,
          status: "active",
        });
      assertNoError(result.error, "Could not save expense.");
      return {
        id,
        date: input.date,
        category: input.category,
        description: input.description,
        amount: input.amount,
        createdAt: now,
        createdBy: userId,
        status: "active",
      };
    }

    const expense: Expense = {
      ...input,
      id: uid(),
      createdAt: now,
      status: "active",
    };
    const snapshot = readLocalSnapshot();
    writeLocalSnapshot({ ...snapshot, expenses: [expense, ...snapshot.expenses] });
    return expense;
  }

  async deleteExpense(id: string) {
    if (this.supabase) {
      const result = await this.supabase.from("expenses").delete().eq("id", id);
      assertNoError(result.error, "Could not delete expense.");
      return;
    }
    const snapshot = readLocalSnapshot();
    writeLocalSnapshot({ ...snapshot, expenses: snapshot.expenses.filter((expense) => expense.id !== id) });
  }

  async voidExpense(id: string, reason: string): Promise<void> {
    if (!reason.trim()) throw new Error("Enter a void reason.");
    if (this.supabase) {
      const result = await this.supabase.rpc("void_expense", { expense_id: id, reason: reason.trim() });
      assertNoError(result.error, "Could not void expense.");
      return;
    }
    const now = new Date().toISOString();
    const snapshot = readLocalSnapshot();
    writeLocalSnapshot({
      ...snapshot,
      expenses: snapshot.expenses.map((expense) => expense.id === id ? { ...expense, status: "voided", voidedAt: now, voidReason: reason.trim() } : expense),
    });
  }

  getExpectedMayaBalance(settings: MayaSettings, ledger: MayaLedgerEntry[]) {
    return expectedMayaBalance(settings, ledger);
  }

  async loadMayaSettings(): Promise<MayaSettings> {
    if (this.supabase) {
      const result = await this.supabase.from("maya_settings").select("*").eq("id", 1).maybeSingle();
      if (isMissingMayaSchema(result.error)) return DEFAULT_MAYA_SETTINGS;
      assertNoError(result.error, "Could not load Maya settings.");
      return toMayaSettings(result.data);
    }
    return readLocalSnapshot().mayaSettings;
  }

  async saveMayaSettings(input: SaveMayaSettingsInput): Promise<MayaSettings> {
    const now = new Date().toISOString();
    if (this.supabase) {
      const userId = await this.getCurrentUserId();
      if (!userId) throw new Error("Please sign in again.");
      const payload = {
        id: 1,
        tracking_enabled: input.trackingEnabled,
        current_balance: Math.max(input.currentBalance, 0),
        low_balance_threshold: Math.max(input.lowBalanceThreshold, 0),
        updated_by: userId,
      };
      const result = await this.supabase.from("maya_settings").upsert(payload, { onConflict: "id" });
      assertMayaSchema(result.error);
      return { ...toMayaSettings(payload), createdAt: now, updatedAt: now };
    }
    const snapshot = readLocalSnapshot();
    const settings: MayaSettings = {
      ...snapshot.mayaSettings,
      trackingEnabled: input.trackingEnabled,
      currentBalance: Math.max(input.currentBalance, 0),
      lowBalanceThreshold: Math.max(input.lowBalanceThreshold, 0),
      updatedAt: now,
    };
    writeLocalSnapshot({ ...snapshot, mayaSettings: settings });
    return settings;
  }

  async listMayaLedger(): Promise<MayaLedgerEntry[]> {
    if (this.supabase) {
      const result = await this.supabase.from("maya_ledger_entries").select("*, services(name)").order("created_at", { ascending: false });
      if (isMissingMayaSchema(result.error)) return [];
      assertNoError(result.error, "Could not load Maya ledger.");
      return result.data?.map(toMayaLedgerEntry) ?? [];
    }
    return readLocalSnapshot().mayaLedger;
  }

  async addMayaTopUp(amount: number, notes: string): Promise<MayaLedgerEntry> {
    const value = Math.max(Number(amount || 0), 0);
    if (value <= 0) throw new Error("Enter a top-up amount.");
    return this.addMayaLedgerEntry({
      entryType: "top_up",
      direction: "in",
      amount: value,
      notes: notes.trim() || "Maya top-up",
      reason: "",
    });
  }

  async addMayaAdjustment(amount: number, reason: string, notes: string): Promise<MayaLedgerEntry> {
    const value = Number(amount || 0);
    if (value === 0) throw new Error("Enter an adjustment amount.");
    if (!reason.trim()) throw new Error("Enter a reason for the adjustment.");
    return this.addMayaLedgerEntry({
      entryType: "adjustment",
      direction: "adjustment",
      amount: value,
      notes: notes.trim(),
      reason: reason.trim(),
    });
  }

  async deleteMayaLedgerEntry(id: string): Promise<void> {
    if (this.supabase) {
      const result = await this.supabase.from("maya_ledger_entries").delete().eq("id", id);
      assertMayaSchema(result.error);
      return;
    }
    const snapshot = readLocalSnapshot();
    writeLocalSnapshot({ ...snapshot, mayaLedger: snapshot.mayaLedger.filter((entry) => entry.id !== id) });
  }

  private buildMayaSaleDeductionEntries(sale: Sale, serviceConfig: Map<string, Service>, settings: MayaSettings, ledger: MayaLedgerEntry[], userId: string | null) {
    if (!settings.trackingEnabled) return [];
    let runningBalance = expectedMayaBalance(settings, ledger);
    const entries: MayaLedgerEntry[] = [];
    const createdAt = new Date().toISOString();
    for (const item of sale.items) {
      const service = item.serviceId ? serviceConfig.get(item.serviceId) : null;
      if (!service?.usesMaya) continue;
      const amount = mayaDeductionForItem(item, service);
      if (amount <= 0) continue;
      runningBalance -= amount;
      entries.push({
        id: uid(),
        entryType: "sale_deduction",
        direction: "out",
        amount,
        balanceAfter: runningBalance,
        saleId: sale.id,
        saleItemId: item.id,
        serviceId: item.serviceId ?? null,
        serviceName: item.serviceName,
        notes: `${item.serviceName} ${item.optionLabel}`,
        reason: "",
        createdBy: userId,
        createdAt,
      });
    }
    return entries;
  }

  async recordMayaSaleDeductions(sale: Sale, serviceConfig?: Map<string, Service>): Promise<void> {
    try {
      if (this.supabase) {
        const [settings, ledger, userId] = await Promise.all([
          this.loadMayaSettings(),
          this.listMayaLedger(),
          this.getCurrentUserId(),
        ]);
        const config = serviceConfig ?? new Map<string, Service>();
        if (!serviceConfig) {
          const serviceIds = uniqueServiceIds(sale.items);
          if (serviceIds.length) {
            const servicesResult = await this.supabase.from("services").select("*").in("id", serviceIds);
            assertNoError(servicesResult.error, "Could not load Maya service config.");
            for (const service of servicesResult.data ?? []) {
              const mapped = toService(service);
              config.set(mapped.id, mapped);
            }
          }
        }
        const entries = this.buildMayaSaleDeductionEntries(sale, config, settings, ledger, userId);
        if (!entries.length) return;
        const rows = entries.map((entry) => ({
          id: entry.id,
          entry_type: entry.entryType,
          amount: entry.amount,
          direction: entry.direction,
          balance_after: entry.balanceAfter ?? null,
          sale_id: entry.saleId ?? null,
          sale_item_id: entry.saleItemId ?? null,
          service_id: entry.serviceId ?? null,
          notes: entry.notes,
          reason: entry.reason,
          created_by: userId,
          created_at: entry.createdAt,
        }));
        const result = await this.supabase.from("maya_ledger_entries").insert(rows);
        assertMayaSchema(result.error);
        return;
      }

      const snapshot = readLocalSnapshot();
      const config = serviceConfig ?? new Map(snapshot.services.map((service) => [service.id, service]));
      const entries = this.buildMayaSaleDeductionEntries(sale, config, snapshot.mayaSettings, snapshot.mayaLedger, null);
      if (entries.length) writeLocalSnapshot({ ...snapshot, mayaLedger: [...entries, ...snapshot.mayaLedger] });
    } catch (error) {
      console.error("[CJNET POS] Maya sale deduction recording failed", error);
    }
  }

  private buildMayaSaleReversalEntries(saleId: string, reason: string, settings: MayaSettings, ledger: MayaLedgerEntry[], userId: string | null) {
    const saleEntries = ledger.filter((entry) => entry.saleId === saleId);
    const deducted = saleEntries
      .filter((entry) => entry.entryType === "sale_deduction")
      .reduce((sum, entry) => sum + Math.abs(entry.amount), 0);
    const reversed = saleEntries
      .filter((entry) => entry.entryType === "sale_reversal")
      .reduce((sum, entry) => sum + Math.abs(entry.amount), 0);
    const remaining = Math.max(deducted - reversed, 0);
    if (remaining <= 0) return [];

    let remainingToReverse = remaining;
    let runningBalance = expectedMayaBalance(settings, ledger);
    const createdAt = new Date().toISOString();
    const entries: MayaLedgerEntry[] = [];

    for (const deduction of saleEntries.filter((entry) => entry.entryType === "sale_deduction")) {
      if (remainingToReverse <= 0) break;
      const amount = Math.min(Math.abs(deduction.amount), remainingToReverse);
      remainingToReverse -= amount;
      runningBalance += amount;
      entries.push({
        id: uid(),
        entryType: "sale_reversal",
        direction: "in",
        amount,
        balanceAfter: runningBalance,
        saleId,
        saleItemId: deduction.saleItemId ?? null,
        serviceId: deduction.serviceId ?? null,
        serviceName: deduction.serviceName ?? null,
        notes: `Maya reversal for sale ${saleId.slice(0, 8)}. ${deduction.notes}`.trim(),
        reason,
        createdBy: userId,
        createdAt,
      });
    }

    return entries;
  }

  async recordMayaSaleReversal(saleId: string, reason: string): Promise<void> {
    try {
      if (this.supabase) {
        const [settings, ledger, userId] = await Promise.all([
          this.loadMayaSettings(),
          this.listMayaLedger(),
          this.getCurrentUserId(),
        ]);
        if (!userId) throw new Error("Please sign in again.");
        const entries = this.buildMayaSaleReversalEntries(saleId, reason, settings, ledger, userId);
        if (!entries.length) return;
        const rows = entries.map((entry) => ({
          id: entry.id,
          entry_type: entry.entryType,
          amount: entry.amount,
          direction: entry.direction,
          balance_after: entry.balanceAfter ?? null,
          sale_id: entry.saleId ?? null,
          sale_item_id: entry.saleItemId ?? null,
          service_id: entry.serviceId ?? null,
          notes: entry.notes,
          reason: entry.reason,
          created_by: userId,
          created_at: entry.createdAt,
        }));
        const result = await this.supabase.from("maya_ledger_entries").insert(rows);
        assertMayaSchema(result.error);
        return;
      }

      const snapshot = readLocalSnapshot();
      const entries = this.buildMayaSaleReversalEntries(saleId, reason, snapshot.mayaSettings, snapshot.mayaLedger, null);
      if (entries.length) writeLocalSnapshot({ ...snapshot, mayaLedger: [...entries, ...snapshot.mayaLedger] });
    } catch (error) {
      console.error("[CJNET POS] Maya sale reversal recording failed", error);
      throw new Error(formatPosError(error, "Sale was changed, but Maya reversal failed. Please add a manual Maya adjustment."));
    }
  }

  private async addMayaLedgerEntry(input: Pick<MayaLedgerEntry, "entryType" | "direction" | "amount" | "notes" | "reason">): Promise<MayaLedgerEntry> {
    const now = new Date().toISOString();
    if (this.supabase) {
      const [settings, ledger, userId] = await Promise.all([
        this.loadMayaSettings(),
        this.listMayaLedger(),
        this.getCurrentUserId(),
      ]);
      if (!userId) throw new Error("Please sign in again.");
      const balanceAfter = expectedMayaBalance(settings, ledger) + mayaEntryEffect(input);
      const id = uid();
      const row = {
        id,
        entry_type: input.entryType,
        amount: input.amount,
        direction: input.direction,
        balance_after: balanceAfter,
        notes: input.notes,
        reason: input.reason,
        created_by: userId,
        created_at: now,
      };
      const result = await this.supabase.from("maya_ledger_entries").insert(row);
      assertMayaSchema(result.error);
      return toMayaLedgerEntry(row);
    }

    const snapshot = readLocalSnapshot();
    const entry: MayaLedgerEntry = {
      id: uid(),
      entryType: input.entryType,
      amount: input.amount,
      direction: input.direction,
      balanceAfter: expectedMayaBalance(snapshot.mayaSettings, snapshot.mayaLedger) + mayaEntryEffect(input),
      saleId: null,
      saleItemId: null,
      serviceId: null,
      serviceName: null,
      notes: input.notes,
      reason: input.reason,
      createdBy: null,
      createdAt: now,
    };
    writeLocalSnapshot({ ...snapshot, mayaLedger: [entry, ...snapshot.mayaLedger] });
    return entry;
  }

  async saveDailyClosing(input: SaveDailyClosingInput): Promise<DailyClosing> {
    const now = new Date().toISOString();
    if (this.supabase) {
      const userId = await this.getCurrentUserId();
      if (!userId) throw new Error("Please sign in again.");
      const id = uid();
      const payload = {
        id,
        closing_date: input.closingDate,
        cash_counted: Boolean(input.cashCounted),
        opening_cash: input.openingCash,
        expected_cash: input.expectedCash,
        actual_cash: input.actualCash,
        cash_difference: input.cashDifference,
        wallet_balance: input.walletBalance ?? null,
        notes: input.notes,
        summary: input.summary,
        closed_by: userId,
        closed_at: now,
      };
      const result = await this.supabase.from("daily_closings").upsert(payload, { onConflict: "closing_date,closed_by" });
      assertNoError(result.error, "Could not save daily closing.");
      return {
        ...input,
        id,
        cashCounted: Boolean(input.cashCounted),
        closedBy: userId,
        closedAt: now,
        createdAt: now,
        updatedAt: now,
      };
    }

    const closing: DailyClosing = {
      ...input,
      cashCounted: Boolean(input.cashCounted),
      id: uid(),
      closedAt: now,
      createdAt: now,
      updatedAt: now,
    };
    const snapshot = readLocalSnapshot();
    writeLocalSnapshot({
      ...snapshot,
      closings: [closing, ...snapshot.closings.filter((item) => item.closingDate !== input.closingDate)],
    });
    return closing;
  }

  async saveService(service: Service): Promise<Service> {
    if (this.supabase) {
      const userId = await this.getCurrentUserId();
      if (!userId) throw new Error("Please sign in again.");
      const serviceCategoryId = categoryId(service.category);
      const categoryResult = await this.supabase
        .from("service_categories")
        .upsert({ id: serviceCategoryId, name: service.category, sort_order: service.sortOrder ?? 100, is_active: true });
      assertNoError(categoryResult.error, "Could not save category. Owner access may be required.");

      const servicePayload = {
        id: service.id,
        name: service.name,
        category_id: serviceCategoryId,
        category: service.category,
        option_label: service.optionLabel,
        price: service.price,
        is_custom_price: Boolean(service.isCustomPrice || service.price <= 0),
        group_name: service.groupName || null,
        requires_tracking: Boolean(service.requiresTracking),
        base_fee: service.baseFee ?? 0,
        service_fee: service.serviceFee ?? 0,
        uses_maya: Boolean(service.usesMaya),
        maya_deduction_amount: service.mayaDeductionAmount ?? 0,
        maya_deduction_mode: service.mayaDeductionMode ?? "pass_through",
        is_active: service.isActive ?? true,
        sort_order: service.sortOrder ?? 100,
        created_by: userId,
        updated_by: userId,
      };
      const result = await this.supabase.from("services").upsert(servicePayload);
      if (isMissingMayaServiceColumn(result.error)) {
        const compatiblePayload = {
          id: servicePayload.id,
          name: servicePayload.name,
          category_id: servicePayload.category_id,
          category: servicePayload.category,
          option_label: servicePayload.option_label,
          price: servicePayload.price,
          is_custom_price: servicePayload.is_custom_price,
          group_name: servicePayload.group_name,
          requires_tracking: servicePayload.requires_tracking,
          base_fee: servicePayload.base_fee,
          service_fee: servicePayload.service_fee,
          is_active: servicePayload.is_active,
          sort_order: servicePayload.sort_order,
          created_by: servicePayload.created_by,
          updated_by: servicePayload.updated_by,
        };
        const compatibleResult = await this.supabase.from("services").upsert(compatiblePayload);
        assertNoError(compatibleResult.error, "Could not save service. Owner access may be required.");
      } else {
        assertNoError(result.error, "Could not save service. Owner access may be required.");
      }
      const pricePayload = {
        service_id: service.id,
        price: service.price,
        is_custom_price: Boolean(service.isCustomPrice || service.price <= 0),
        base_fee: service.baseFee ?? 0,
        service_fee: service.serviceFee ?? 0,
        created_by: userId,
        effective_from: new Date().toISOString(),
      };
      const legacyPricePayload = {
        service_id: service.id,
        price: service.price,
        is_custom_price: Boolean(service.isCustomPrice || service.price <= 0),
        created_by: userId,
        effective_from: pricePayload.effective_from,
      };
      const priceResult = await this.supabase
        .from("price_settings")
        .upsert(
          pricePayload,
          { onConflict: "service_id" },
        );
      const normalizedPricePayload = errorCode(priceResult.error) === "42703" ? legacyPricePayload : pricePayload;
      const normalizedPriceError = errorCode(priceResult.error) === "42703"
        ? (await this.supabase.from("price_settings").upsert(legacyPricePayload, { onConflict: "service_id" })).error
        : priceResult.error;
      if (errorCode(normalizedPriceError) === "42P10") {
        const historyResult = await this.supabase.from("price_settings").insert(normalizedPricePayload);
        assertNoError(historyResult.error, "Service saved, but price history could not be updated.");
      } else {
        assertNoError(normalizedPriceError, "Service saved, but price history could not be updated.");
      }
      return service;
    }

    const snapshot = readLocalSnapshot();
    const services = snapshot.services.some((item) => item.id === service.id)
      ? snapshot.services.map((item) => (item.id === service.id ? service : item))
      : [...snapshot.services, service];
    writeLocalSnapshot({ ...snapshot, services });
    return service;
  }

  async deleteService(id: string) {
    if (this.supabase) {
      const result = await this.supabase.from("services").update({ is_active: false }).eq("id", id);
      assertNoError(result.error, "Could not delete service. Owner access may be required.");
      return;
    }
    const snapshot = readLocalSnapshot();
    writeLocalSnapshot({ ...snapshot, services: snapshot.services.filter((service) => service.id !== id) });
  }

  async resetServices() {
    if (this.supabase) {
      await Promise.all(DEFAULT_SERVICES.map((service) => this.saveService(service)));
      return;
    }
    const snapshot = readLocalSnapshot();
    writeLocalSnapshot({ ...snapshot, services: DEFAULT_SERVICES });
  }

  exportSnapshot(): PosSnapshot {
    return readLocalSnapshot();
  }

  importSnapshot(snapshot: PosSnapshot) {
    writeLocalSnapshot(snapshot);
  }

  private async getCurrentUserId() {
    if (!this.supabase) return null;
    const { data, error } = await this.supabase.auth.getUser();
    assertNoError(error, "Could not verify the signed-in user.");
    return data.user?.id ?? null;
  }
}
