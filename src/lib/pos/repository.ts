"use client";

import type { CartItem, Expense, PosSnapshot, Sale, SaleItem, Service } from "@/types/pos";
import { DEFAULT_SERVICES } from "@/lib/pos/defaults";
import { changeDue, dateKey, todayKey } from "@/lib/pos/calculations";
import { createSupabaseBrowserClient } from "@/lib/pos/supabase-client";

const STORAGE_KEY = "cjnet_pos_next_snapshot";

type SaveSaleInput = {
  cart: CartItem[];
  discount: number;
  cashReceived: number;
  customerNote: string;
};

type SaveExpenseInput = Omit<Expense, "id" | "createdAt">;

function uid(prefix: string) {
  if (typeof crypto !== "undefined" && "randomUUID" in crypto) {
    return crypto.randomUUID();
  }
  return `${prefix}_${Date.now()}_${Math.random().toString(16).slice(2)}`;
}

function emptySnapshot(): PosSnapshot {
  return { services: DEFAULT_SERVICES, sales: [], expenses: [] };
}

function withRequiredDefaultServices(services: Service[]) {
  const existingIds = new Set(services.map((service) => service.id));
  const requiredServices = DEFAULT_SERVICES.filter((service) => ["police-clearance-custom", "psa-custom"].includes(service.id) && !existingIds.has(service.id));
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
    Finishing: "finishing",
    Custom: "custom",
  };
  return known[category] ?? (category.trim().toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/(^-|-$)/g, "") || "custom");
}

function toService(row: Record<string, unknown>): Service {
  return {
    id: String(row.id),
    name: String(row.name),
    category: String(row.category),
    optionLabel: String(row.option_label),
    price: Number(row.price ?? 0),
    isCustomPrice: Boolean(row.is_custom_price),
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
    items,
  };
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

    if (servicesResult.error) throw servicesResult.error;
    if (salesResult.error) throw salesResult.error;
    if (expensesResult.error) throw expensesResult.error;

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
    const saleItems: SaleItem[] = input.cart.map((item) => ({
      id: uid("sale_item"),
      serviceId: item.serviceId,
      serviceName: item.name,
      category: item.category,
      optionLabel: item.optionLabel,
      quantity: item.quantity,
      unitPrice: item.price,
      lineTotal: item.price * item.quantity,
    }));

    if (this.supabase) {
      if (!userId) throw new Error("Please sign in again.");
      const saleInsert = {
        customer_note: input.customerNote,
        subtotal,
        discount,
        total,
        cash_received: cash,
        change_due: changeDue(cash, total),
        sold_at: now,
        cashier_id: userId,
        status: "completed",
      };
      const saleResult = await this.supabase.from("sales").insert(saleInsert).select("*").single();
      if (saleResult.error) throw saleResult.error;

      const saleId = saleResult.data.id as string;
      const itemRows = saleItems.map((item) => ({
        sale_id: saleId,
        service_id: item.serviceId,
        service_name: item.serviceName,
        category: item.category,
        option_label: item.optionLabel,
        quantity: item.quantity,
        unit_price: item.unitPrice,
        line_total: item.lineTotal,
      }));
      const itemsResult = await this.supabase.from("sale_items").insert(itemRows).select("*");
      if (itemsResult.error) throw itemsResult.error;
      return toSale({ ...saleResult.data, sale_items: itemsResult.data });
    }

    const sale: Sale = {
      id: uid("sale"),
      createdAt: now,
      soldAt: now,
      date: todayKey(),
      customerNote: input.customerNote,
      subtotal,
      discount,
      total,
      cashReceived: cash,
      changeDue: changeDue(cash, total),
      items: saleItems.map((item) => ({ ...item, saleId: undefined })),
    };
    const snapshot = readLocalSnapshot();
    writeLocalSnapshot({ ...snapshot, sales: [sale, ...snapshot.sales] });
    return sale;
  }

  async deleteSale(id: string) {
    if (this.supabase) {
      const result = await this.supabase.from("sales").delete().eq("id", id);
      if (result.error) throw result.error;
      return;
    }
    const snapshot = readLocalSnapshot();
    writeLocalSnapshot({ ...snapshot, sales: snapshot.sales.filter((sale) => sale.id !== id) });
  }

  async saveExpense(input: SaveExpenseInput): Promise<Expense> {
    if (this.supabase) {
      const userId = await this.getCurrentUserId();
      if (!userId) throw new Error("Please sign in again.");
      const result = await this.supabase
        .from("expenses")
        .insert({
          expense_date: input.date,
          category: input.category,
          description: input.description,
          amount: input.amount,
          created_by: userId,
        })
        .select("*")
        .single();
      if (result.error) throw result.error;
      return toExpense(result.data);
    }

    const expense: Expense = {
      ...input,
      id: uid("expense"),
      createdAt: new Date().toISOString(),
    };
    const snapshot = readLocalSnapshot();
    writeLocalSnapshot({ ...snapshot, expenses: [expense, ...snapshot.expenses] });
    return expense;
  }

  async deleteExpense(id: string) {
    if (this.supabase) {
      const result = await this.supabase.from("expenses").delete().eq("id", id);
      if (result.error) throw result.error;
      return;
    }
    const snapshot = readLocalSnapshot();
    writeLocalSnapshot({ ...snapshot, expenses: snapshot.expenses.filter((expense) => expense.id !== id) });
  }

  async saveService(service: Service): Promise<Service> {
    if (this.supabase) {
      const userId = await this.getCurrentUserId();
      if (!userId) throw new Error("Please sign in again.");
      const serviceCategoryId = categoryId(service.category);
      const categoryResult = await this.supabase
        .from("service_categories")
        .upsert({ id: serviceCategoryId, name: service.category, sort_order: service.sortOrder ?? 100, is_active: true });
      if (categoryResult.error) throw categoryResult.error;

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
          is_active: service.isActive ?? true,
          sort_order: service.sortOrder ?? 100,
          created_by: userId,
          updated_by: userId,
        })
        .select("*")
        .single();
      if (result.error) throw result.error;
      const priceResult = await this.supabase
        .from("price_settings")
        .upsert(
          {
            service_id: service.id,
            price: service.price,
            is_custom_price: Boolean(service.isCustomPrice || service.price <= 0),
            created_by: userId,
            effective_from: new Date().toISOString(),
          },
          { onConflict: "service_id" },
        )
        .select("*")
        .single();
      if (priceResult.error) throw priceResult.error;
      return toService(result.data);
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
      if (result.error) throw result.error;
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
    if (error) throw error;
    return data.user?.id ?? null;
  }
}
