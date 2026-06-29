"use client";

import type { CartItem, Expense, PosSnapshot, Sale, SaleItem, Service } from "@/types/pos";
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
  return { services: DEFAULT_SERVICES, sales: [], expenses: [] };
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
    };
    if (services.length !== (parsed.services?.length ?? 0)) {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(snapshot));
    }
    return {
      services: snapshot.services,
      sales: snapshot.sales,
      expenses: snapshot.expenses,
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

export class PosRepository {
  private supabase = createSupabaseBrowserClient();

  get mode() {
    return this.supabase ? "supabase" : "local";
  }

  async loadSnapshot(): Promise<PosSnapshot> {
    if (!this.supabase) return readLocalSnapshot();

    const [servicesResult, salesResult, expensesResult] = await Promise.all([
      this.supabase.from("services").select("*").eq("is_active", true).order("sort_order", { ascending: true }),
      this.supabase.from("sales").select("*, sale_items(*)").order("sold_at", { ascending: false }),
      this.supabase.from("expenses").select("*").order("expense_date", { ascending: false }),
    ]);

    assertNoError(servicesResult.error, "Could not load services.");
    assertNoError(salesResult.error, "Could not load sales.");
    assertNoError(expensesResult.error, "Could not load expenses.");

    return {
      services: servicesResult.data?.map(toService) ?? DEFAULT_SERVICES,
      sales: salesResult.data?.map(toSale) ?? [],
      expenses: expensesResult.data?.map(toExpense) ?? [],
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
      if (requestedServiceIds.length) {
        const servicesResult = await this.supabase.from("services").select("id").in("id", requestedServiceIds);
        assertNoError(servicesResult.error, "Could not verify service records before saving the sale.");
        for (const service of servicesResult.data ?? []) {
          validServiceIds.add(String(service.id));
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
      }));
      const itemsResult = await this.supabase.from("sale_items").insert(itemRows);
      assertNoError(itemsResult.error, "Sale was created, but line items could not be saved.");
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
    writeLocalSnapshot({ ...snapshot, sales: [sale, ...snapshot.sales] });
    return sale;
  }

  async deleteSale(id: string) {
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
      return;
    }
    const now = new Date().toISOString();
    const snapshot = readLocalSnapshot();
    writeLocalSnapshot({
      ...snapshot,
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

  async saveService(service: Service): Promise<Service> {
    if (this.supabase) {
      const userId = await this.getCurrentUserId();
      if (!userId) throw new Error("Please sign in again.");
      const serviceCategoryId = categoryId(service.category);
      const categoryResult = await this.supabase
        .from("service_categories")
        .upsert({ id: serviceCategoryId, name: service.category, sort_order: service.sortOrder ?? 100, is_active: true });
      assertNoError(categoryResult.error, "Could not save category. Owner access may be required.");

      const result = await this.supabase
        .from("services")
        .upsert({
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
          is_active: service.isActive ?? true,
          sort_order: service.sortOrder ?? 100,
          created_by: userId,
          updated_by: userId,
        });
      assertNoError(result.error, "Could not save service. Owner access may be required.");
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
