"use client";

import Image from "next/image";
import Link from "next/link";
import { ChangeEvent, FormEvent, RefObject, ReactNode, useCallback, useEffect, useMemo, useRef, useState } from "react";
import { useRouter } from "next/navigation";
import type { CartItem, Expense, Sale, Service } from "@/types/pos";
import { buildReport, cartSubtotal, changeDue, dateKey, money, saleTotal, todayKey } from "@/lib/pos/calculations";
import { formatPosError, logPosError } from "@/lib/pos/errors";
import { PosRepository } from "@/lib/pos/repository";
import { createSupabaseBrowserClient, hasSupabaseConfig } from "@/lib/pos/supabase-client";

type Tab = "register" | "dashboard" | "audit" | "sales" | "expenses" | "reports" | "prices" | "staff" | "settings";
type DashboardPeriod = "today" | "week" | "month";
type AppRole = "owner" | "manager" | "staff";
type Status = { tone: "success" | "error" | "info"; message: string } | null;
type ServiceGroup = { id: string; name: string; category: string; services: Service[]; sortOrder: number };

const tabs: Array<{ id: Tab; label: string }> = [
  { id: "dashboard", label: "Dashboard" },
  { id: "register", label: "Register" },
  { id: "sales", label: "Sales" },
  { id: "expenses", label: "Expenses" },
  { id: "reports", label: "Reports" },
  { id: "prices", label: "Prices" },
  { id: "audit", label: "Audit" },
  { id: "staff", label: "Staff" },
  { id: "settings", label: "Settings" },
];

const expenseCategories = ["Food", "Ink", "Bond Paper", "Lamination Film", "Internet", "Electricity", "Rent", "Salary"];
const quickCash = [20, 50, 100, 500, 1000];
const roleAccess: Record<AppRole, Tab[]> = {
  owner: ["dashboard", "register", "sales", "expenses", "reports", "prices", "audit", "staff", "settings"],
  manager: ["dashboard", "register", "sales", "expenses", "reports", "audit"],
  staff: ["register", "sales", "expenses"],
};

function uid(prefix: string) {
  if (typeof crypto !== "undefined" && "randomUUID" in crypto) return crypto.randomUUID();
  return `${prefix}_${Date.now()}_${Math.random().toString(16).slice(2)}`;
}

function niceDate(value: string) {
  return new Date(value).toLocaleString([], {
    year: "numeric",
    month: "short",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
  });
}

function downloadFile(filename: string, content: string, type: string) {
  const blob = new Blob([content], { type });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

function csvEscape(value: unknown) {
  return `"${String(value ?? "").replaceAll('"', '""')}"`;
}

function htmlEscape(value: unknown) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
}

function excelTable(title: string, headers: string[], rows: Array<Array<string | number>>) {
  const headerHtml = headers.map((header) => `<th>${htmlEscape(header)}</th>`).join("");
  const rowsHtml = rows.length
    ? rows.map((row) => `<tr>${row.map((cell) => `<td>${htmlEscape(cell)}</td>`).join("")}</tr>`).join("")
    : `<tr><td colspan="${headers.length}">No records</td></tr>`;

  return `
    <h2>${htmlEscape(title)}</h2>
    <table>
      <thead><tr>${headerHtml}</tr></thead>
      <tbody>${rowsHtml}</tbody>
    </table>
  `;
}

function downloadExcel(filename: string, title: string, sections: string[]) {
  const html = `
    <html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">
      <head>
        <meta charset="utf-8" />
        <!--[if gte mso 9]><xml><x:ExcelWorkbook><x:ExcelWorksheets><x:ExcelWorksheet><x:Name>${htmlEscape(title).slice(0, 31)}</x:Name><x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions></x:ExcelWorksheet></x:ExcelWorksheets></x:ExcelWorkbook></xml><![endif]-->
        <style>
          body { font-family: Arial, sans-serif; color: #151515; }
          h1 { font-size: 20px; margin: 0 0 12px; }
          h2 { font-size: 15px; margin: 22px 0 8px; }
          table { border-collapse: collapse; margin-bottom: 14px; width: 100%; }
          th { background: #ffd400; border: 1px solid #d8cba8; font-weight: 700; text-align: left; }
          td { border: 1px solid #eadfca; }
          th, td { padding: 7px 9px; }
        </style>
      </head>
      <body>
        <h1>${htmlEscape(title)}</h1>
        ${sections.join("")}
      </body>
    </html>
  `;
  downloadFile(filename, html, "application/vnd.ms-excel;charset=utf-8");
}

function getDashboardRange(period: DashboardPeriod) {
  const now = new Date();
  if (period === "month") {
    return { from: dateKey(new Date(now.getFullYear(), now.getMonth(), 1)), to: todayKey() };
  }
  if (period === "week") {
    const from = new Date(now);
    from.setDate(now.getDate() - 6);
    return { from: dateKey(from), to: todayKey() };
  }
  return { from: todayKey(), to: todayKey() };
}

function buildSalesTrend(sales: Sale[], range: { from: string; to: string }) {
  const start = new Date(`${range.from}T00:00:00`);
  const end = new Date(`${range.to}T00:00:00`);
  const rows: Array<{ date: string; label: string; total: number }> = [];

  for (const cursor = new Date(start); cursor <= end; cursor.setDate(cursor.getDate() + 1)) {
    const key = dateKey(cursor);
    const total = sales.filter((sale) => sale.status !== "voided" && sale.date === key).reduce((sum, sale) => sum + sale.total, 0);
    rows.push({
      date: key,
      label: cursor.toLocaleDateString([], { month: "short", day: "numeric" }),
      total,
    });
  }

  return rows;
}

function currentMonthKey() {
  return todayKey().slice(0, 7);
}

function monthRange(month: string) {
  const [year, monthNumber] = month.split("-").map(Number);
  const start = new Date(year, monthNumber - 1, 1);
  const end = new Date(year, monthNumber, 0);
  return { from: dateKey(start), to: dateKey(end) };
}

function monthTitle(month: string) {
  const [year, monthNumber] = month.split("-").map(Number);
  return new Date(year, monthNumber - 1, 1).toLocaleDateString([], { month: "long", year: "numeric" });
}

function hasOnlineService(item: Pick<CartItem, "category">) {
  return item.category === "Online Services";
}

function needsReference(item: Pick<CartItem, "requiresTracking">) {
  return Boolean(item.requiresTracking);
}

function isGcashService(name: string, category: string) {
  return `${name} ${category}`.toLowerCase().includes("gcash");
}

function gcashServiceFee(amount: number) {
  if (amount <= 0) return 0;
  return Math.max(1, Math.floor(amount / 1000)) * 15;
}

function buildCashierActivity(sales: Sale[], range: { from: string; to: string }) {
  const activity = new Map<string, { count: number; total: number }>();
  for (const sale of sales) {
    if (sale.date < range.from || sale.date > range.to) continue;
    const name = sale.cashierId ? `Cashier ${sale.cashierId.slice(0, 8)}` : "Local cashier";
    const current = activity.get(name) ?? { count: 0, total: 0 };
    current.count += 1;
    current.total += sale.total;
    activity.set(name, current);
  }
  return [...activity.entries()].map(([name, value]) => ({ name, ...value })).sort((a, b) => b.total - a.total);
}

function canAccessTab(role: AppRole, tab: Tab) {
  return roleAccess[role].includes(tab);
}

function roleLabel(role: AppRole) {
  if (role === "owner") return "Owner";
  if (role === "manager") return "Manager";
  return "Staff";
}

function normalizeGroupKey(value: string) {
  return value.trim().toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/(^-|-$)/g, "") || "service";
}

function inferredGroupName(service: Service) {
  if (service.groupName?.trim()) return service.groupName.trim();
  const text = `${service.name} ${service.optionLabel}`.toLowerCase();
  const knownGroups = [
    { match: ["sss"], name: "SSS" },
    { match: ["pag-ibig", "pag ibig", "hdmf"], name: "PAG-IBIG" },
    { match: ["philhealth"], name: "PhilHealth" },
    { match: ["nbi"], name: "NBI" },
    { match: ["psa"], name: "PSA" },
    { match: ["police clearance"], name: "Police clearance" },
    { match: ["laminat"], name: "Lamination" },
    { match: ["folder"], name: "Folders" },
    { match: ["black and white"], name: "Black and white print" },
    { match: ["colored", "color print"], name: "Colored print" },
    { match: ["xerox"], name: "Xerox" },
  ];
  return knownGroups.find((group) => group.match.some((keyword) => text.includes(keyword)))?.name ?? service.name;
}

function buildServiceGroups(services: Service[]) {
  const groups = new Map<string, ServiceGroup>();
  for (const service of services) {
    const name = inferredGroupName(service);
    const key = `${service.category}:${normalizeGroupKey(name)}`;
    const current = groups.get(key) ?? { id: key, name, category: service.category, services: [], sortOrder: service.sortOrder ?? 100 };
    current.services.push(service);
    current.sortOrder = Math.min(current.sortOrder, service.sortOrder ?? 100);
    groups.set(key, current);
  }
  return [...groups.values()]
    .map((group) => ({ ...group, services: group.services.sort((a, b) => (a.sortOrder ?? 100) - (b.sortOrder ?? 100)) }))
    .sort((a, b) => a.sortOrder - b.sortOrder || a.name.localeCompare(b.name));
}

function categorySummary(category: string, services: Service[]) {
  const count = services.filter((service) => service.category === category).length;
  return `${count} option${count === 1 ? "" : "s"}`;
}

function canVoidSale(sale: Sale, role: AppRole, userId: string | null, today: string) {
  if (sale.status === "voided") return false;
  if (role === "owner" || role === "manager") return true;
  return sale.date === today && (!sale.cashierId || sale.cashierId === userId);
}

function canVoidExpense(expense: Expense, role: AppRole, userId: string | null, today: string) {
  if (expense.status === "voided") return false;
  if (role === "owner" || role === "manager") return true;
  return expense.date === today && (!expense.createdBy || expense.createdBy === userId);
}

export function PosApp() {
  const repo = useMemo(() => new PosRepository(), []);
  const router = useRouter();
  const supabase = useMemo(() => createSupabaseBrowserClient(), []);
  const [role, setRole] = useState<AppRole>(() => (hasSupabaseConfig() ? "staff" : "owner"));
  const [currentUserId, setCurrentUserId] = useState<string | null>(null);
  const [accountName, setAccountName] = useState(() => (hasSupabaseConfig() ? "Cashier" : "Local mode"));
  const [authLoading, setAuthLoading] = useState(() => hasSupabaseConfig());
  const [activeTab, setActiveTab] = useState<Tab>("register");
  const [services, setServices] = useState<Service[]>([]);
  const [sales, setSales] = useState<Sale[]>([]);
  const [expenses, setExpenses] = useState<Expense[]>([]);
  const [cart, setCart] = useState<CartItem[]>([]);
  const [loading, setLoading] = useState(true);
  const [saving, setSaving] = useState(false);
  const [status, setStatus] = useState<Status>(null);
  const [search, setSearch] = useState("");
  const [category, setCategory] = useState("");
  const [variantGroup, setVariantGroup] = useState<ServiceGroup | null>(null);
  const [discount, setDiscount] = useState(0);
  const [cashReceived, setCashReceived] = useState(0);
  const [customerNote, setCustomerNote] = useState("");
  const [needsFollowUp, setNeedsFollowUp] = useState(false);
  const [customOpen, setCustomOpen] = useState(false);
  const [customService, setCustomService] = useState<Service | null>(null);
  const [customName, setCustomName] = useState("");
  const [customCategory, setCustomCategory] = useState("Custom");
  const [customPrice, setCustomPrice] = useState("");
  const [customBaseFee, setCustomBaseFee] = useState("");
  const [customServiceFee, setCustomServiceFee] = useState("");
  const [customGcashMode, setCustomGcashMode] = useState<"cash-in" | "cash-out">("cash-in");
  const [customQty, setCustomQty] = useState(1);
  const [expenseForm, setExpenseForm] = useState({ date: todayKey(), category: "", description: "", amount: "" });
  const [serviceForm, setServiceForm] = useState({ name: "", category: "", optionLabel: "", price: "", baseFee: "", serviceFee: "", requiresTracking: false });
  const [salesFilters, setSalesFilters] = useState({ from: todayKey(), to: todayKey(), search: "" });
  const [reportRange, setReportRange] = useState({ from: todayKey(), to: todayKey() });
  const [dashboardPeriod, setDashboardPeriod] = useState<DashboardPeriod>("today");
  const [auditMonth, setAuditMonth] = useState(currentMonthKey());
  const importRef = useRef<HTMLInputElement>(null);

  const subtotal = cartSubtotal(cart);
  const total = saleTotal(subtotal, discount);
  const change = changeDue(cashReceived, total);
  const today = todayKey();
  const todayReport = buildReport(sales, expenses, { from: today, to: today });
  const todayCollected = todayReport.grossSales;
  const todayNonIncome = todayReport.passThroughFees;
  const todayEarned = todayReport.serviceRevenue;
  const todayExpenses = todayReport.expenses;
  const todayNet = todayReport.netIncome;
  const todayTransactions = todayReport.transactions;
  const allowedTabs = tabs.filter((tab) => canAccessTab(role, tab.id));
  const categories = useMemo(() => Array.from(new Set(services.map((service) => service.category))).sort((a, b) => {
    const preferred = ["Printing", "Xerox", "Online Services", "GCash", "Lamination", "Finishing", "Typing", "Materials", "Rush ID", "Custom"];
    const left = preferred.indexOf(a);
    const right = preferred.indexOf(b);
    if (left !== -1 || right !== -1) return (left === -1 ? 999 : left) - (right === -1 ? 999 : right);
    return a.localeCompare(b);
  }), [services]);
  const filteredServices = services.filter((service) => {
    const term = search.trim().toLowerCase();
    const haystack = `${service.name} ${service.category} ${service.optionLabel}`.toLowerCase();
    return (!category || service.category === category || Boolean(term)) && (!term || haystack.includes(term));
  });
  const serviceGroups = useMemo(() => buildServiceGroups(filteredServices), [filteredServices]);
  const filteredSales = sales.filter((sale) => {
    const saleText = `${sale.customerNote} ${sale.items.map((item) => `${item.serviceName} ${item.optionLabel}`).join(" ")}`.toLowerCase();
    const rangeAllowed = role === "owner" || role === "manager" || sale.date === today;
    return rangeAllowed && sale.date >= salesFilters.from && sale.date <= salesFilters.to && (!salesFilters.search || saleText.includes(salesFilters.search.toLowerCase()));
  });
  const report = buildReport(sales, expenses, reportRange);
  const dashboardRange = getDashboardRange(dashboardPeriod);
  const dashboardReport = buildReport(sales, expenses, dashboardRange);
  const dashboardTrend = buildSalesTrend(sales, dashboardRange);
  const cashierActivity = buildCashierActivity(sales, dashboardRange);
  const auditRange = monthRange(auditMonth);
  const auditReport = buildReport(sales, expenses, auditRange);
  const auditTrend = buildSalesTrend(sales, auditRange);
  const recentSales = sales.slice(0, 6);
  const cartHasOnlineService = cart.some(hasOnlineService);
  const saleNeedsReference = needsFollowUp || cart.some(needsReference);
  const visibleTab = canAccessTab(role, activeTab) ? activeTab : role === "owner" ? "dashboard" : "register";
  const accessAllowed = canAccessTab(role, visibleTab);

  const refreshSnapshot = useCallback(async () => {
    const snapshot = await repo.loadSnapshot();
    setServices(snapshot.services);
    setSales(snapshot.sales);
    setExpenses(snapshot.expenses);
    return snapshot;
  }, [repo]);

  const refreshAfterWrite = useCallback(async (successMessage: string) => {
    try {
      await refreshSnapshot();
      setStatus({ tone: "success", message: successMessage });
    } catch (error) {
      logPosError("refreshSnapshot after mutation", error);
      setStatus({ tone: "error", message: "Saved, but refresh failed. Please reload." });
    }
  }, [refreshSnapshot]);

  useEffect(() => {
    if (!supabase) return;

    const client = supabase;

    let cancelled = false;

    async function loadAuth() {
      const { data: authData, error: authError } = await client.auth.getUser();
      if (authError || !authData.user) {
        router.replace("/login");
        return;
      }

      const { data: profile, error: profileError } = await client
        .from("profiles")
        .select("full_name, role, status")
        .eq("id", authData.user.id)
        .single();

      if (profileError || !profile || profile.status === "disabled") {
        await client.auth.signOut();
        router.replace("/login?error=access_denied");
        return;
      }

      if (cancelled) return;
      setCurrentUserId(authData.user.id);
      setRole(profile.role === "owner" ? "owner" : profile.role === "manager" ? "manager" : "staff");
      setAccountName(profile.full_name?.trim() || authData.user.email || "Cashier");
      setAuthLoading(false);
    }

    void loadAuth();
    return () => {
      cancelled = true;
    };
  }, [router, supabase]);

  useEffect(() => {
    async function load() {
      try {
        setLoading(true);
        await refreshSnapshot();
      } catch (error) {
        logPosError("loadSnapshot", error);
        setStatus({ tone: "error", message: formatPosError(error, "Could not load POS data.") });
      } finally {
        setLoading(false);
      }
    }
    load();
  }, [refreshSnapshot]);

  useEffect(() => {
    if (!status) return;
    const timer = window.setTimeout(() => setStatus(null), 3200);
    return () => window.clearTimeout(timer);
  }, [status]);

  function addToCart(service: Service, quantity = 1, price = service.price, name = service.name, optionLabel = service.optionLabel) {
    setCart((current) => {
      const existing = current.find((item) => item.serviceId === service.id && item.price === price && item.name === name);
      if (existing) {
        return current.map((item) => (item.id === existing.id ? { ...item, quantity: item.quantity + quantity } : item));
      }
      return [
        ...current,
        {
          id: uid("cart"),
          serviceId: service.id,
          name,
          category: service.category,
          optionLabel,
          price,
          quantity,
          requiresTracking: Boolean(service.requiresTracking),
          baseFee: service.baseFee ?? 0,
          serviceFee: service.serviceFee ?? Math.max(price - (service.baseFee ?? 0), 0),
          passThroughFee: service.baseFee ?? 0,
          revenueAmount: service.serviceFee ?? Math.max(price - (service.baseFee ?? 0), 0),
          pricingBreakdown: service.baseFee || service.serviceFee ? { baseFee: service.baseFee ?? 0, serviceFee: service.serviceFee ?? Math.max(price - (service.baseFee ?? 0), 0), total: price } : null,
        },
      ];
    });
  }

  function onServiceClick(service: Service) {
    const hasFeeBreakdown = (service.baseFee ?? 0) > 0 || (service.serviceFee ?? 0) > 0;
    if (hasFeeBreakdown) {
      addToCart(service, 1, (service.baseFee ?? 0) + (service.serviceFee ?? 0));
      return;
    }
    if (service.isCustomPrice || service.price <= 0 || service.category === "Online Services") {
      setCustomService(service);
      setCustomName(service.name);
      setCustomCategory(service.category);
      setCustomPrice("");
      setCustomBaseFee(service.baseFee ? String(service.baseFee) : "");
      setCustomServiceFee(service.serviceFee ? String(service.serviceFee) : "");
      setCustomGcashMode("cash-in");
      setCustomQty(1);
      setCustomOpen(true);
      return;
    }
    addToCart(service);
  }

  async function saveSale() {
    if (saving) return;
    if (!cart.length) {
      setStatus({ tone: "error", message: "Add items first." });
      return;
    }
    if (saleNeedsReference && !customerNote.trim()) {
      setStatus({ tone: "error", message: "Enter customer/reference details for follow-up or tracked services." });
      return;
    }
    try {
      setSaving(true);
      const note = customerNote.trim();
      const sale = await repo.saveSale({ cart, discount, cashReceived, customerNote: note, needsFollowUp });
      setSales((current) => [sale, ...current]);
      setCart([]);
      setDiscount(0);
      setCashReceived(0);
      setCustomerNote("");
      setNeedsFollowUp(false);
      await refreshAfterWrite(`Sale saved: ${money(sale.total)}`);
    } catch (error) {
      logPosError("saveSale", error);
      setStatus({ tone: "error", message: formatPosError(error, "Could not save sale.") });
    } finally {
      setSaving(false);
    }
  }

  async function saveExpense(event: FormEvent) {
    event.preventDefault();
    const amount = Number(expenseForm.amount);
    if (!expenseForm.category.trim() || amount <= 0) {
      setStatus({ tone: "error", message: "Enter expense category and amount." });
      return;
    }
    try {
      setSaving(true);
      const expense = await repo.saveExpense({
        date: expenseForm.date || todayKey(),
        category: expenseForm.category.trim(),
        description: expenseForm.description.trim(),
        amount,
      });
      setExpenses((current) => [expense, ...current]);
      setExpenseForm({ date: todayKey(), category: "", description: "", amount: "" });
      await refreshAfterWrite("Expense saved.");
    } catch (error) {
      logPosError("saveExpense", error);
      setStatus({ tone: "error", message: formatPosError(error, "Could not save expense.") });
    } finally {
      setSaving(false);
    }
  }

  async function saveServiceSettings(service: Service, updates: Partial<Service>) {
    const baseFee = Math.max(Number(updates.baseFee ?? service.baseFee ?? 0), 0);
    const serviceFee = Math.max(Number(updates.serviceFee ?? service.serviceFee ?? 0), 0);
    const hasFeeBreakdown = baseFee > 0 || serviceFee > 0;
    const price = hasFeeBreakdown ? baseFee + serviceFee : Math.max(Number(updates.price ?? service.price ?? 0), 0);
    const isCustomPrice = hasFeeBreakdown
      ? false
      : Boolean(updates.isCustomPrice ?? service.isCustomPrice ?? price <= 0) || service.category === "Online Services";
    try {
      setSaving(true);
      const updated = await repo.saveService({
        ...service,
        ...updates,
        price,
        baseFee,
        serviceFee,
        isCustomPrice,
      });
      setServices((current) => current.map((item) => (item.id === updated.id ? updated : item)));
      await refreshAfterWrite("Service settings saved.");
    } catch (error) {
      logPosError("saveServiceSettings", error);
      setStatus({ tone: "error", message: formatPosError(error, "Could not save service settings.") });
    } finally {
      setSaving(false);
    }
  }

  async function addService(event: FormEvent) {
    event.preventDefault();
    if (!serviceForm.name.trim() || !serviceForm.category.trim() || !serviceForm.optionLabel.trim()) {
      setStatus({ tone: "error", message: "Complete the service details." });
      return;
    }
    const baseFee = Number(serviceForm.baseFee || 0);
    const serviceFee = Number(serviceForm.serviceFee || 0);
    const hasFeeBreakdown = baseFee > 0 || serviceFee > 0;
    const price = hasFeeBreakdown ? baseFee + serviceFee : Number(serviceForm.price || 0);
    const category = serviceForm.category.trim();
    const service: Service = {
      id: `service-${Date.now()}`,
      name: serviceForm.name.trim(),
      category,
      optionLabel: serviceForm.optionLabel.trim(),
      price,
      baseFee,
      serviceFee,
      requiresTracking: serviceForm.requiresTracking,
      isCustomPrice: hasFeeBreakdown ? false : price <= 0 || category === "Online Services",
      sortOrder: services.length + 100,
    };
    try {
      setSaving(true);
      const saved = await repo.saveService(service);
      setServices((current) => [...current, saved]);
      setServiceForm({ name: "", category: "", optionLabel: "", price: "", baseFee: "", serviceFee: "", requiresTracking: false });
      await refreshAfterWrite("Service added.");
    } catch (error) {
      logPosError("addService", error);
      setStatus({ tone: "error", message: formatPosError(error, "Could not add service.") });
    } finally {
      setSaving(false);
    }
  }

  async function deleteSale(id: string) {
    if (!window.confirm("Delete this sale?")) return;
    try {
      await repo.deleteSale(id);
      setSales((current) => current.filter((sale) => sale.id !== id));
      await refreshAfterWrite("Sale deleted.");
    } catch (error) {
      logPosError("deleteSale", error);
      setStatus({ tone: "error", message: formatPosError(error, "Could not delete sale.") });
    }
  }

  async function voidSale(id: string) {
    const reason = window.prompt("Reason for void/correction");
    if (reason === null) return;
    try {
      await repo.voidSale(id, reason);
      const now = new Date().toISOString();
      setSales((current) => current.map((sale) => sale.id === id ? { ...sale, status: "voided", voidedAt: now, voidReason: reason.trim() } : sale));
      await refreshAfterWrite("Transaction voided.");
    } catch (error) {
      logPosError("voidSale", error);
      setStatus({ tone: "error", message: formatPosError(error, "Could not void transaction.") });
    }
  }

  async function deleteExpense(id: string) {
    if (!window.confirm("Delete this expense?")) return;
    try {
      await repo.deleteExpense(id);
      setExpenses((current) => current.filter((expense) => expense.id !== id));
      await refreshAfterWrite("Expense deleted.");
    } catch (error) {
      logPosError("deleteExpense", error);
      setStatus({ tone: "error", message: formatPosError(error, "Could not delete expense.") });
    }
  }

  async function voidExpense(id: string) {
    const reason = window.prompt("Reason for void/correction");
    if (reason === null) return;
    try {
      await repo.voidExpense(id, reason);
      const now = new Date().toISOString();
      setExpenses((current) => current.map((expense) => expense.id === id ? { ...expense, status: "voided", voidedAt: now, voidReason: reason.trim() } : expense));
      await refreshAfterWrite("Expense voided.");
    } catch (error) {
      logPosError("voidExpense", error);
      setStatus({ tone: "error", message: formatPosError(error, "Could not void expense.") });
    }
  }

  async function deleteService(id: string) {
    if (!window.confirm("Delete this service?")) return;
    try {
      await repo.deleteService(id);
      setServices((current) => current.filter((service) => service.id !== id));
      await refreshAfterWrite("Service deleted.");
    } catch (error) {
      logPosError("deleteService", error);
      setStatus({ tone: "error", message: formatPosError(error, "Could not delete service.") });
    }
  }

  function exportSalesCsv() {
    const rows = [["id", "date", "items", "note", "subtotal", "discount", "total", "cash", "change"]];
    for (const sale of filteredSales) {
      rows.push([
        sale.id,
        sale.soldAt,
        sale.items.map((item) => `${item.quantity}x ${item.serviceName} ${item.optionLabel}`).join("; "),
        sale.customerNote,
        String(sale.subtotal),
        String(sale.discount),
        String(sale.total),
        String(sale.cashReceived),
        String(sale.changeDue),
      ]);
    }
    downloadFile(`sales-${todayKey()}.csv`, rows.map((row) => row.map(csvEscape).join(",")).join("\n"), "text/csv");
  }

  function exportExpensesCsv() {
    const rows = [["id", "date", "category", "description", "amount"]];
    for (const expense of expenses) rows.push([expense.id, expense.date, expense.category, expense.description, String(expense.amount)]);
    downloadFile(`expenses-${todayKey()}.csv`, rows.map((row) => row.map(csvEscape).join(",")).join("\n"), "text/csv");
  }

  function exportExpensesExcel() {
    downloadExcel(`expenses-${todayKey()}.xls`, "CJNET Expenses", [
      excelTable("Expenses", ["Date", "Category", "Description", "Amount"], expenses.map((expense) => [expense.date, expense.category, expense.description, expense.amount])),
      excelTable("Summary by category", ["Category", "Total"], buildReport([], expenses, { from: "0000-01-01", to: "9999-12-31" }).expenseSummary.map((item) => [item.category, item.total])),
    ]);
    setStatus({ tone: "success", message: "Expenses Excel exported." });
  }

  function exportReportExcel() {
    downloadExcel(`report-${reportRange.from}-to-${reportRange.to}.xls`, "CJNET Report", [
      excelTable("Date range", ["From", "To"], [[reportRange.from, reportRange.to]]),
      excelTable("Summary", ["Metric", "Value"], [
        ["Collected", report.grossSales],
        ["Non-income", report.passThroughFees],
        ["Earned", report.serviceRevenue],
        ["Expenses", report.expenses],
        ["Net", report.netIncome],
        ["Transactions", report.transactions],
      ]),
      excelTable("Top services", ["Service", "Quantity", "Total"], report.topServices.map((item) => [item.name, item.quantity, item.total])),
      excelTable("Expense summary", ["Category", "Total"], report.expenseSummary.map((item) => [item.category, item.total])),
    ]);
    setStatus({ tone: "success", message: "Report Excel exported." });
  }

  function exportAuditExcel() {
    const averageSale = auditReport.transactions ? auditReport.grossSales / auditReport.transactions : 0;
    downloadExcel(`monthly-audit-${auditMonth}.xls`, `CJNET Monthly Audit - ${monthTitle(auditMonth)}`, [
      excelTable("Month", ["Month", "From", "To"], [[monthTitle(auditMonth), auditRange.from, auditRange.to]]),
      excelTable("Summary", ["Metric", "Value"], [
        ["Collected", auditReport.grossSales],
        ["Non-income", auditReport.passThroughFees],
        ["Earned", auditReport.serviceRevenue],
        ["Expenses", auditReport.expenses],
        ["Net", auditReport.netIncome],
        ["Transactions", auditReport.transactions],
        ["Average sale", averageSale],
      ]),
      excelTable("Daily sales chart data", ["Date", "Label", "Sales"], auditTrend.map((row) => [row.date, row.label, row.total])),
      excelTable("Top services", ["Service", "Quantity", "Total"], auditReport.topServices.map((item) => [item.name, item.quantity, item.total])),
      excelTable("Expense audit", ["Category", "Total"], auditReport.expenseSummary.map((item) => [item.category, item.total])),
    ]);
    setStatus({ tone: "success", message: "Monthly audit Excel exported." });
  }

  function exportBackup() {
    downloadFile(
      `cjnet-pos-backup-${todayKey()}.json`,
      JSON.stringify({ exportedAt: new Date().toISOString(), services, sales, expenses }, null, 2),
      "application/json",
    );
  }

  async function importBackup(event: ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];
    if (!file) return;
    try {
      const parsed = JSON.parse(await file.text()) as { services: Service[]; sales: Sale[]; expenses: Expense[] };
      if (!Array.isArray(parsed.services) || !Array.isArray(parsed.sales) || !Array.isArray(parsed.expenses)) {
        throw new Error("Invalid backup file.");
      }
      repo.importSnapshot({ services: parsed.services, sales: parsed.sales, expenses: parsed.expenses });
      setServices(parsed.services);
      setSales(parsed.sales);
      setExpenses(parsed.expenses);
      setStatus({ tone: "success", message: "Backup restored." });
    } catch (error) {
      setStatus({ tone: "error", message: error instanceof Error ? error.message : "Could not restore backup." });
    } finally {
      event.target.value = "";
    }
  }

  if (loading || authLoading) {
    return (
      <main className="app-shell flex items-center justify-center">
        <div className="rounded-[1.35rem] border border-surface-border bg-surface-card px-6 py-5 shadow-[var(--shadow-soft)]">
          <p className="text-sm font-semibold text-text-secondary">Loading CJNET POS...</p>
        </div>
      </main>
    );
  }

  return (
    <main className="app-shell">
      <section className="mx-auto flex w-full max-w-[1440px] min-w-0 flex-col gap-3">
        <header className="pos-header rounded-[1.1rem] border border-surface-border bg-white px-3 py-2 shadow-[var(--shadow-soft)]">
          <div className="flex flex-col gap-2 lg:flex-row lg:items-center lg:justify-between">
            <div className="flex flex-wrap items-center gap-3">
              <Link href="/" className="utility-logo-link pos-logo-link w-[132px]">
                <Image src="/logo.png" alt="CJ NET" width={920} height={311} className="h-auto w-full" priority />
              </Link>
              <div className="rounded-full border border-surface-border bg-[rgba(255,212,0,0.14)] px-3 py-1 text-xs font-semibold text-foreground">
                {accountName} / {roleLabel(role)}
              </div>
              {supabase ? (
                <button
                  type="button"
                  className="secondary-btn !min-h-9 !px-3"
                  onClick={async () => {
                    await supabase.auth.signOut();
                    router.replace("/login");
                  }}
                >
                  Sign out
                </button>
              ) : null}
            </div>
            <nav className="flex gap-2 overflow-x-auto">
              {allowedTabs.map((tab) => (
                <button key={tab.id} type="button" onClick={() => setActiveTab(tab.id)} className={activeTab === tab.id ? "pos-nav pos-nav-active" : "pos-nav"}>
                  {tab.label}
                </button>
              ))}
            </nav>
          </div>
        </header>

        <section className={`summary-bar ${visibleTab === "register" ? "summary-bar-compact" : ""} grid gap-2 rounded-[1.1rem] border border-surface-border bg-[rgba(255,255,255,0.72)] p-2 shadow-[var(--shadow-card)] md:grid-cols-5`}>
          <SummaryTile label="Collected" value={money(todayCollected)} subtext={`${todayTransactions} transactions`} />
          <SummaryTile label="Non-income" value={money(todayNonIncome)} subtext="Pass-through cash" />
          <SummaryTile label="Earned" value={money(todayEarned)} subtext="Actual shop revenue" />
          <SummaryTile label="Expenses" value={money(todayExpenses)} subtext="Recorded today" />
          <SummaryTile label="Net" value={money(todayNet)} subtext="Earned minus expenses" />
        </section>

        {!accessAllowed ? <AccessDenied role={role} /> : null}

        {accessAllowed && visibleTab === "register" ? (
          <RegisterView
            services={services}
            serviceGroups={serviceGroups}
            categories={categories}
            category={category}
            search={search}
            cart={cart}
            subtotal={subtotal}
            discount={discount}
            total={total}
            cashReceived={cashReceived}
            change={change}
            customerNote={customerNote}
            needsFollowUp={needsFollowUp}
            showReference={cartHasOnlineService || saleNeedsReference}
            requiresReference={saleNeedsReference}
            saving={saving}
            onCategory={setCategory}
            onSearch={setSearch}
            onServiceClick={onServiceClick}
            onGroupClick={setVariantGroup}
            onCart={setCart}
            onDiscount={setDiscount}
            onCash={setCashReceived}
            onCustomerNote={setCustomerNote}
            onNeedsFollowUp={setNeedsFollowUp}
            onClear={() => {
              setCart([]);
              setDiscount(0);
              setCashReceived(0);
              setCustomerNote("");
              setNeedsFollowUp(false);
            }}
            onSave={saveSale}
            onCustom={() => {
              setCustomOpen(true);
              setCustomService(null);
              setCustomName("");
              setCustomCategory(category || "Custom");
              setCustomPrice("");
              setCustomBaseFee("");
              setCustomServiceFee("");
              setCustomGcashMode("cash-in");
              setCustomQty(1);
            }}
          />
        ) : null}

        {accessAllowed && visibleTab === "dashboard" ? (
          <DashboardView
            period={dashboardPeriod}
            range={dashboardRange}
            todayCollected={todayCollected}
            todayNonIncome={todayNonIncome}
            todayEarned={todayEarned}
            todayExpenses={todayExpenses}
            todayNet={todayNet}
            todayTransactions={todayTransactions}
            report={dashboardReport}
            trend={dashboardTrend}
            recentSales={recentSales}
            cashierActivity={cashierActivity}
            onPeriod={setDashboardPeriod}
          />
        ) : null}

        {accessAllowed && visibleTab === "audit" ? (
          <AuditView
            month={auditMonth}
            title={monthTitle(auditMonth)}
            range={auditRange}
            report={auditReport}
            trend={auditTrend}
            onMonth={setAuditMonth}
            onExportExcel={exportAuditExcel}
          />
        ) : null}

        {accessAllowed && visibleTab === "sales" ? <SalesView filters={salesFilters} sales={filteredSales} role={role} userId={currentUserId} today={today} onFilters={setSalesFilters} onExport={exportSalesCsv} onVoid={voidSale} onDelete={deleteSale} /> : null}
        {accessAllowed && visibleTab === "expenses" ? (
          <ExpensesView form={expenseForm} expenses={expenses} role={role} userId={currentUserId} today={today} saving={saving} onForm={setExpenseForm} onSubmit={saveExpense} onExportCsv={exportExpensesCsv} onExportExcel={exportExpensesExcel} onVoid={voidExpense} onDelete={deleteExpense} />
        ) : null}
        {accessAllowed && visibleTab === "reports" ? (
          <ReportsView
            range={reportRange}
            report={report}
            onRange={setReportRange}
            onExportExcel={exportReportExcel}
            onExportBackup={exportBackup}
            onImportClick={() => importRef.current?.click()}
            importRef={importRef}
            onImport={importBackup}
          />
        ) : null}
        {accessAllowed && visibleTab === "prices" ? (
          <PricesView
            services={services}
            form={serviceForm}
            onForm={setServiceForm}
            onAdd={addService}
            onSaveSettings={saveServiceSettings}
            onDelete={deleteService}
            onReset={async () => {
              if (!window.confirm("Reset services to default prices?")) return;
              try {
                await repo.resetServices();
                await refreshAfterWrite("Default prices restored.");
              } catch (error) {
                logPosError("resetServices", error);
                setStatus({ tone: "error", message: formatPosError(error, "Could not reset services.") });
              }
            }}
          />
        ) : null}
        {accessAllowed && visibleTab === "staff" ? <StaffView /> : null}
        {accessAllowed && visibleTab === "settings" ? <SettingsView /> : null}
      </section>

      {customOpen ? (
        <CustomDialog
          open={customOpen}
          name={customName}
          category={customCategory}
          price={customPrice}
          baseFee={customBaseFee}
          serviceFee={customServiceFee}
          gcashMode={customGcashMode}
          quantity={customQty}
          onName={setCustomName}
          onCategory={setCustomCategory}
          onPrice={setCustomPrice}
          onBaseFee={setCustomBaseFee}
          onServiceFee={setCustomServiceFee}
          onGcashMode={setCustomGcashMode}
          onQuantity={setCustomQty}
          onClose={() => {
            setCustomOpen(false);
            setCustomService(null);
            setCustomName("");
            setCustomBaseFee("");
            setCustomServiceFee("");
            setCustomGcashMode("cash-in");
          }}
          onAdd={() => {
            const gcash = isGcashService(customName, customCategory);
            const isOnline = customCategory === "Online Services";
            const baseFee = Number(customBaseFee || 0);
            const serviceFee = gcash ? gcashServiceFee(baseFee) : Number(customServiceFee || 0);
            const price = gcash || isOnline || baseFee || serviceFee ? baseFee + serviceFee : Number(customPrice);
            if (!customName.trim() || price <= 0) {
              setStatus({ tone: "error", message: isOnline || gcash ? "Enter the bill amount and earned fee." : "Enter a name and price." });
              return;
            }
            const service = customService ?? {
              id: `custom-${Date.now()}`,
              name: customName.trim(),
              category: customCategory.trim() || "Custom",
              optionLabel: "Custom",
              price,
              isCustomPrice: true,
            };
            addToCart(
              {
                ...service,
                name: customName.trim(),
                category: customCategory.trim() || "Custom",
                baseFee,
                serviceFee: baseFee || serviceFee ? serviceFee : Math.max(price - baseFee, 0),
              },
              gcash ? 1 : Math.max(customQty, 1),
              price,
              customName.trim(),
              gcash ? (customGcashMode === "cash-in" ? "Cash in" : "Cash out") : customService?.optionLabel ?? "Custom",
            );
            setCustomOpen(false);
            setCustomService(null);
            setCustomName("");
            setCustomBaseFee("");
            setCustomServiceFee("");
            setCustomGcashMode("cash-in");
          }}
        />
      ) : null}

      {variantGroup ? (
        <VariantDialog
          group={variantGroup}
          onClose={() => setVariantGroup(null)}
          onSelect={(service) => {
            setVariantGroup(null);
            onServiceClick(service);
          }}
        />
      ) : null}

      {status ? <StatusToast status={status} /> : null}
    </main>
  );
}

function SummaryTile({ label, value, subtext }: { label: string; value: string; subtext: string }) {
  return (
    <div className="summary-tile rounded-xl bg-white px-3 py-2">
      <p className="text-[11px] font-semibold uppercase tracking-[0.08em] text-text-secondary">{label}</p>
      <p className="mt-1 truncate text-base font-bold text-foreground">{value}</p>
      <p className="summary-subtext mt-1 truncate text-[11px] font-medium text-text-secondary">{subtext}</p>
    </div>
  );
}

function StatusToast({ status }: { status: Exclude<Status, null> }) {
  const className = status.tone === "error" ? "status-toast status-toast-error" : "status-toast status-toast-success";
  return (
    <div className={className} role="status" aria-live="polite">
      <span className="status-toast-dot" />
      <span>{status.message}</span>
    </div>
  );
}

function DashboardView({ period, range, todayCollected, todayNonIncome, todayEarned, todayExpenses, todayNet, todayTransactions, report, trend, recentSales, cashierActivity, onPeriod }: {
  period: DashboardPeriod;
  range: { from: string; to: string };
  todayCollected: number;
  todayNonIncome: number;
  todayEarned: number;
  todayExpenses: number;
  todayNet: number;
  todayTransactions: number;
  report: ReturnType<typeof buildReport>;
  trend: Array<{ date: string; label: string; total: number }>;
  recentSales: Sale[];
  cashierActivity: Array<{ name: string; count: number; total: number }>;
  onPeriod: (period: DashboardPeriod) => void;
}) {
  const periodOptions: Array<{ id: DashboardPeriod; label: string }> = [
    { id: "today", label: "Daily" },
    { id: "week", label: "Weekly" },
    { id: "month", label: "Monthly" },
  ];

  return (
    <section>
      <SectionHeader
        title="Dashboard"
        action={
          <div className="flex gap-2 overflow-x-auto">
            {periodOptions.map((option) => (
              <button key={option.id} type="button" onClick={() => onPeriod(option.id)} className={period === option.id ? "category-pill category-pill-active" : "category-pill"}>
                {option.label}
              </button>
            ))}
          </div>
        }
      />

      <div className="grid gap-3 md:grid-cols-6">
        <ReportCard label="Collected" value={money(todayCollected)} subtext="Cash received today" />
        <ReportCard label="Non-income" value={money(todayNonIncome)} subtext="Pass-through cash" />
        <ReportCard label="Earned" value={money(todayEarned)} subtext="Actual shop revenue" />
        <ReportCard label="Expenses" value={money(todayExpenses)} subtext="Shop costs recorded" />
        <ReportCard label="Net" value={money(todayNet)} subtext="Earned minus expenses" />
        <ReportCard label="Transactions" value={String(todayTransactions)} subtext="Completed receipts" />
      </div>

      <div className="mt-4 grid gap-3 xl:grid-cols-[minmax(0,1.2fr)_minmax(320px,0.8fr)]">
        <TrendCard rows={trend} range={range} />
        <ListCard title="Best-selling services" rows={report.topServices.slice(0, 5).map((item) => ({ label: item.name, meta: `${item.quantity} sold`, value: money(item.total) }))} empty="No service sales yet." />
      </div>

      <div className="mt-4 grid gap-3 xl:grid-cols-[minmax(0,1fr)_minmax(320px,0.8fr)]">
        <ListCard title="Recent sales" rows={recentSales.map((sale) => ({ label: sale.items.map((item) => `${item.quantity}x ${item.serviceName}`).join(", "), meta: niceDate(sale.soldAt), value: money(sale.total) }))} empty="No sales recorded yet." />
        <ListCard title="Cashier activity" rows={cashierActivity.map((item) => ({ label: item.name, meta: `${item.count} transactions`, value: money(item.total) }))} empty="No cashier activity yet." />
      </div>
    </section>
  );
}

function AuditView({ month, title, range, report, trend, onMonth, onExportExcel }: {
  month: string;
  title: string;
  range: { from: string; to: string };
  report: ReturnType<typeof buildReport>;
  trend: Array<{ date: string; label: string; total: number }>;
  onMonth: (month: string) => void;
  onExportExcel: () => void;
}) {
  const averageSale = report.transactions ? report.grossSales / report.transactions : 0;

  return (
    <section>
      <SectionHeader
        title="Monthly Audit"
        action={
          <div className="grid gap-2 sm:grid-cols-[11rem_auto_auto]">
            <input className="input-field" type="month" value={month} onChange={(event) => onMonth(event.target.value || currentMonthKey())} />
            <button type="button" className="secondary-btn" onClick={() => onMonth(currentMonthKey())}>
              This month
            </button>
            <button type="button" className="secondary-btn" onClick={onExportExcel}>
              Export Excel
            </button>
          </div>
        }
      />

      <div className="mb-4 rounded-[18px] border border-surface-border bg-white/75 px-4 py-3 shadow-[var(--shadow-card)]">
        <p className="text-xs font-medium text-text-secondary">{range.from} to {range.to}</p>
        <h2 className="mt-1 text-base font-bold">{title}</h2>
      </div>

      <div className="grid gap-3 md:grid-cols-7">
        <ReportCard label="Collected" value={money(report.grossSales)} />
        <ReportCard label="Non-income" value={money(report.passThroughFees)} />
        <ReportCard label="Earned" value={money(report.serviceRevenue)} />
        <ReportCard label="Expenses" value={money(report.expenses)} />
        <ReportCard label="Net" value={money(report.netIncome)} />
        <ReportCard label="Transactions" value={String(report.transactions)} />
        <ReportCard label="Average sale" value={money(averageSale)} />
      </div>

      <div className="mt-4 grid gap-3 xl:grid-cols-[minmax(0,1.25fr)_minmax(320px,0.75fr)]">
        <TrendCard rows={trend} range={range} title="Monthly sales chart" />
        <ListCard title="Top services" rows={report.topServices.map((item) => ({ label: item.name, meta: `${item.quantity} sold`, value: money(item.total) }))} empty="No service sales this month." />
      </div>

      <div className="mt-4 grid gap-3 xl:grid-cols-2">
        <ListCard title="Expense audit" rows={report.expenseSummary.map((item) => ({ label: item.category, value: money(item.total) }))} empty="No expenses this month." />
        <ListCard title="Audit totals" rows={[
          { label: "Collected after discounts", value: money(report.grossSales) },
          { label: "Non-income", value: money(report.passThroughFees) },
          { label: "Earned", value: money(report.serviceRevenue) },
          { label: "Recorded expenses", value: money(report.expenses) },
          { label: "Net after expenses", value: money(report.netIncome) },
          { label: "Sales count", value: String(report.transactions) },
        ]} empty="No audit data." />
      </div>
    </section>
  );
}

function RegisterView(props: {
  services: Service[];
  serviceGroups: ServiceGroup[];
  categories: string[];
  category: string;
  search: string;
  cart: CartItem[];
  subtotal: number;
  discount: number;
  total: number;
  cashReceived: number;
  change: number;
  customerNote: string;
  needsFollowUp: boolean;
  showReference: boolean;
  requiresReference: boolean;
  saving: boolean;
  onCategory: (value: string) => void;
  onSearch: (value: string) => void;
  onServiceClick: (service: Service) => void;
  onGroupClick: (group: ServiceGroup) => void;
  onCart: (updater: (current: CartItem[]) => CartItem[]) => void;
  onDiscount: (value: number) => void;
  onCash: (value: number) => void;
  onCustomerNote: (value: string) => void;
  onNeedsFollowUp: (value: boolean) => void;
  onClear: () => void;
  onSave: () => void;
  onCustom: () => void;
}) {
  const showingSearch = props.search.trim().length > 0;
  const showingCategories = !props.category && !showingSearch;

  return (
    <section>
      <div className="register-grid grid items-start gap-3 xl:grid-cols-[minmax(0,1fr)_clamp(320px,27vw,370px)]">
        <section className="section-card register-services flex min-w-0 flex-col rounded-[16px] p-4">
          <div className="mb-3 flex items-center justify-between gap-3">
            <div>
              <h2 className="text-base font-bold">{props.category ? props.category : "Choose category"}</h2>
              {props.category ? <p className="mt-1 text-xs text-text-secondary">Select a service group or variant.</p> : null}
            </div>
            <button type="button" onClick={props.onCustom} className="secondary-btn !min-h-10 !px-3 !py-2">
              Custom
            </button>
          </div>

          <div className="mb-3 grid gap-2 md:grid-cols-[minmax(0,1fr)_190px]">
            <input className="input-field" value={props.search} onChange={(event) => props.onSearch(event.target.value)} placeholder="Search service" />
            <select className="input-field" value={props.category} onChange={(event) => props.onCategory(event.target.value)}>
              <option value="">Categories</option>
              {props.categories.map((category) => (
                <option key={category} value={category}>
                  {category}
                </option>
              ))}
            </select>
          </div>

          {showingCategories ? (
            <div className="register-card-grid register-category-grid grid grid-cols-[repeat(auto-fill,minmax(165px,1fr))] gap-3">
              {props.categories.map((category) => (
                <button key={category} type="button" onClick={() => props.onCategory(category)} className="category-card">
                  <span className="text-base font-bold">{category}</span>
                  <span className="mt-1 text-[11px] font-semibold text-text-secondary">{categorySummary(category, props.services)}</span>
                </button>
              ))}
            </div>
          ) : (
            <>
              <div className="mb-3 flex flex-wrap items-center gap-2">
                {props.category ? (
                  <button type="button" onClick={() => props.onCategory("")} className="secondary-btn !min-h-9 !px-3 !py-2">
                    Back
                  </button>
                ) : null}
                {props.category ? <span className="category-pill category-pill-active">{props.category}</span> : null}
                {showingSearch && !props.category ? <span className="category-pill category-pill-active">Search results</span> : null}
              </div>
              <div className="register-card-grid register-service-grid grid grid-cols-[repeat(auto-fill,minmax(160px,1fr))] gap-3">
                {props.serviceGroups.map((group) => (
                  group.services.length > 1 ? (
                    <ServiceGroupCard key={group.id} group={group} onClick={() => props.onGroupClick(group)} />
                  ) : (
                    <ServiceCard key={group.services[0].id} service={group.services[0]} onClick={() => props.onServiceClick(group.services[0])} />
                  )
                ))}
              </div>
              {!props.serviceGroups.length ? <p className="mt-4 text-sm text-text-secondary">No services found.</p> : null}
            </>
          )}
        </section>

        <aside className="checkout-panel section-card sticky top-3 flex h-[calc(100vh-168px)] max-h-[calc(100vh-168px)] min-h-[500px] min-w-0 flex-col rounded-[16px] p-4 max-xl:static max-xl:h-auto max-xl:max-h-none max-xl:min-h-0">
          <div className="receipt-header flex items-start justify-between gap-3">
            <div>
              <h2 className="text-base font-bold">Current Sale</h2>
              <p className="receipt-count mt-1 text-xs text-text-secondary">{props.cart.reduce((sum, item) => sum + item.quantity, 0)} items in receipt</p>
            </div>
            <button type="button" onClick={props.onClear} className="danger-btn !min-h-9 !px-3 !py-2">
              Clear
            </button>
          </div>

          <div className={props.cart.length ? "receipt-list mt-3 grid gap-2 overflow-auto pr-1" : "receipt-empty mt-3 grid place-items-center rounded-2xl border-2 border-dashed border-[rgba(23,23,23,0.12)] bg-white/70 p-4 text-center text-sm text-text-secondary"}>
            {props.cart.length ? (
              props.cart.map((item) => <CartLine key={item.id} item={item} onCart={props.onCart} />)
            ) : (
              <div className="grid justify-items-center gap-2">
                <p className="font-semibold text-foreground">No items yet</p>
                <p className="max-w-48 text-xs leading-5 text-text-secondary">Tap a service to start the receipt.</p>
              </div>
            )}
          </div>

          <div className="mt-auto grid gap-2 pt-3">
            {props.showReference ? (
              <label className="block">
                <span className="mb-2 block text-xs font-medium text-text-secondary">
                  Customer / reference / notes {props.requiresReference ? "" : "(optional)"}
                </span>
                <input className="input-field" value={props.customerNote} onChange={(event) => props.onCustomerNote(event.target.value)} placeholder="Name, reference, remarks" />
              </label>
            ) : null}

            <label className="flex items-center gap-2 rounded-xl border border-[rgba(23,23,23,0.08)] bg-white/70 p-2 text-xs font-semibold">
              <input
                type="checkbox"
                checked={props.needsFollowUp}
                onChange={(event) => props.onNeedsFollowUp(event.target.checked)}
                className="h-5 w-5"
              />
              Needs follow-up / claiming later
            </label>

            <div className="payment-panel subtle-panel grid gap-2 rounded-2xl p-3">
              <TotalRow label="Subtotal" value={money(props.subtotal)} />
              <div className="flex items-center justify-between gap-3">
                <label className="text-xs font-medium text-text-secondary">Discount</label>
                <input className="input-field max-w-[150px]" type="number" min="0" step="0.01" value={props.discount} onChange={(event) => props.onDiscount(Number(event.target.value || 0))} />
              </div>
              <TotalRow label="Total" value={money(props.total)} variant="total" />
              <div className="flex items-center justify-between gap-3">
                <label className="text-xs font-medium text-text-secondary">Cash</label>
                <input className="input-field max-w-[150px]" type="number" min="0" step="0.01" value={props.cashReceived || ""} onChange={(event) => props.onCash(Number(event.target.value || 0))} placeholder="0.00" />
              </div>
              <div className="flex gap-2 overflow-x-auto">
                {quickCash.map((amount) => (
                  <button key={amount} type="button" onClick={() => props.onCash(amount)} className={props.cashReceived === amount ? "cash-pill cash-pill-active" : "cash-pill"}>
                    {amount}
                  </button>
                ))}
              </div>
              <TotalRow label="Change" value={money(props.change)} variant="change" />
            </div>

            <button type="button" onClick={props.onSave} className="primary-btn w-full text-base font-bold" disabled={props.saving}>
              {props.saving ? "Saving..." : "Save Sale"}
            </button>
          </div>
        </aside>
      </div>
    </section>
  );
}

function ServiceCard({ service, onClick, compact = false }: { service: Service; onClick: () => void; compact?: boolean }) {
  return (
    <button type="button" data-service-id={service.id} onClick={onClick} className={`service-card ${compact ? "h-[104px]" : "h-[112px]"}`}>
      <div>
        <p className="line-clamp-2 text-[14px] font-semibold leading-5 text-foreground">{service.name}</p>
        <p className="mt-1 line-clamp-2 text-xs font-normal leading-4 text-text-secondary">
          {service.category} / {service.optionLabel}
        </p>
      </div>
      <span className="price-badge">{service.isCustomPrice || service.price <= 0 ? "Custom price" : money(service.price)}</span>
    </button>
  );
}

function ServiceGroupCard({ group, onClick }: { group: ServiceGroup; onClick: () => void }) {
  const minPrice = Math.min(...group.services.map((service) => service.price).filter((price) => price > 0));
  const hasCustom = group.services.some((service) => service.isCustomPrice || service.price <= 0);
  return (
    <button type="button" onClick={onClick} className="service-card h-[112px]">
      <div>
        <p className="line-clamp-2 text-[14px] font-semibold leading-5 text-foreground">{group.name}</p>
        <p className="mt-1 line-clamp-2 text-xs font-normal leading-4 text-text-secondary">
          {group.category} / {group.services.length} options
        </p>
      </div>
      <span className="price-badge">{hasCustom ? "Choose option" : `From ${money(Number.isFinite(minPrice) ? minPrice : 0)}`}</span>
    </button>
  );
}

function VariantDialog({ group, onClose, onSelect }: { group: ServiceGroup; onClose: () => void; onSelect: (service: Service) => void }) {
  return (
    <div className="fixed inset-0 z-50 grid place-items-center bg-black/40 p-4">
      <div className="w-full max-w-xl rounded-[18px] bg-white p-5 shadow-[0_24px_70px_rgba(20,23,31,0.18)]">
        <div className="mb-4 flex items-start justify-between gap-3">
          <div>
            <h2 className="text-lg font-bold">{group.name}</h2>
            <p className="mt-1 text-sm text-text-secondary">{group.category}</p>
          </div>
          <button type="button" className="secondary-btn !min-h-9 !px-3 !py-2" onClick={onClose}>Close</button>
        </div>
        <div className="grid gap-2">
          {group.services.map((service) => (
            <button key={service.id} type="button" className="service-card !h-auto min-h-[74px]" onClick={() => onSelect(service)}>
              <div>
                <p className="text-sm font-semibold">{service.name}</p>
                <p className="mt-1 text-xs text-text-secondary">{service.optionLabel}</p>
              </div>
              <span className="price-badge">{service.isCustomPrice || service.price <= 0 ? "Custom price" : money(service.price)}</span>
            </button>
          ))}
        </div>
      </div>
    </div>
  );
}

function CartLine({ item, onCart }: { item: CartItem; onCart: (updater: (current: CartItem[]) => CartItem[]) => void }) {
  return (
    <div className="rounded-xl border border-[rgba(23,23,23,0.08)] bg-white p-2.5">
      <div className="flex justify-between gap-3">
        <div className="min-w-0">
          <p className="truncate text-sm font-semibold">{item.name}</p>
          <p className="mt-1 text-xs text-text-secondary">
            {item.category} / {item.optionLabel} / {money(item.price)} each
          </p>
          {item.passThroughFee ? (
            <p className="mt-1 text-[11px] text-text-secondary">
              {isGcashService(item.name, item.category) ? "GCash amount" : "Pass-through"} {money(item.passThroughFee)} + earned {money(item.revenueAmount ?? Math.max(item.price - item.passThroughFee, 0))}
            </p>
          ) : null}
        </div>
        <p className="text-sm font-bold">{money(item.price * item.quantity)}</p>
      </div>
      <div className="mt-2 flex items-center justify-between gap-2">
        <button type="button" className="danger-btn !min-h-9 !px-3" onClick={() => onCart((current) => current.filter((cartItem) => cartItem.id !== item.id))}>
          Remove
        </button>
        <div className="flex items-center gap-2">
          <button type="button" className="qty-btn" onClick={() => onCart((current) => current.map((cartItem) => (cartItem.id === item.id ? { ...cartItem, quantity: Math.max(1, cartItem.quantity - 1) } : cartItem)))}>
            -
          </button>
          <input
            className="qty-input"
            type="number"
            min="1"
            step="1"
            inputMode="numeric"
            value={item.quantity}
            onChange={(event) => {
              const value = Math.max(1, Number(event.target.value || 1));
              onCart((current) => current.map((cartItem) => (cartItem.id === item.id ? { ...cartItem, quantity: value } : cartItem)));
            }}
          />
          <button type="button" className="qty-btn" onClick={() => onCart((current) => current.map((cartItem) => (cartItem.id === item.id ? { ...cartItem, quantity: cartItem.quantity + 1 } : cartItem)))}>
            +
          </button>
        </div>
      </div>
    </div>
  );
}

function TotalRow({ label, value, variant = "default" }: { label: string; value: string; variant?: "default" | "total" | "change" }) {
  if (variant !== "default") {
    return (
      <div className={`total-row total-row-${variant}`}>
        <span>{label}</span>
        <strong>{value}</strong>
      </div>
    );
  }

  return (
    <div className="flex items-center justify-between gap-3">
      <span className="text-xs font-medium text-text-secondary">{label}</span>
      <strong className="text-sm font-bold">{value}</strong>
    </div>
  );
}

function AccessDenied({ role }: { role: AppRole }) {
  return (
    <section className="section-card rounded-[18px] p-6">
      <h1 className="text-2xl font-bold">Access denied</h1>
      <p className="mt-2 max-w-xl text-sm leading-6 text-text-secondary">
        This area is protected for owner access. Current role: {role}.
      </p>
    </section>
  );
}

function StaffView() {
  return (
    <section>
      <SectionHeader title="Staff" />
      <div className="grid gap-3 lg:grid-cols-2">
        <ListCard title="Accounts" rows={[
          { label: "Owner", meta: "Full access", value: "Active" },
          { label: "Staff", meta: "Cashier access", value: "Active" },
        ]} empty="No staff roles configured." />
        <ListCard title="Recent staff activity" rows={[]} empty="No staff activity yet." />
      </div>
    </section>
  );
}

function SettingsView() {
  return (
    <section>
      <SectionHeader title="Settings" />
      <div className="grid gap-3 lg:grid-cols-2">
        <ListCard title="Cashier permissions" rows={[
          { label: "Staff expenses", meta: "Add only", value: "On" },
          { label: "Staff discounts", meta: "Register", value: "On" },
        ]} empty="No settings configured." />
        <ListCard title="Owner controls" rows={[
          { label: "Price changes", meta: "Owner only", value: "Locked" },
          { label: "Audit logs", meta: "Owner only", value: "Locked" },
        ]} empty="No security settings." />
      </div>
    </section>
  );
}

function SalesView({ filters, sales, role, userId, today, onFilters, onExport, onVoid, onDelete }: {
  filters: { from: string; to: string; search: string };
  sales: Sale[];
  role: AppRole;
  userId: string | null;
  today: string;
  onFilters: (filters: { from: string; to: string; search: string }) => void;
  onExport: () => void;
  onVoid: (id: string) => void;
  onDelete: (id: string) => void;
}) {
  return (
    <section>
      <SectionHeader title="Sales" action={<button className="secondary-btn" onClick={onExport}>Export CSV</button>} />
      <div className="filter-row">
        <input className="input-field" type="date" value={filters.from} onChange={(event) => onFilters({ ...filters, from: event.target.value })} />
        <input className="input-field" type="date" value={filters.to} onChange={(event) => onFilters({ ...filters, to: event.target.value })} />
        <input className="input-field" placeholder="Search note or item" value={filters.search} onChange={(event) => onFilters({ ...filters, search: event.target.value })} />
      </div>
      <DataTable empty="No sales for this filter.">
        {sales.map((sale) => (
          <tr key={sale.id}>
            <td>{niceDate(sale.soldAt)}</td>
            <td>{sale.items.map((item) => `${item.quantity}x ${item.serviceName} ${item.optionLabel}`).join(", ")}</td>
            <td>
              {sale.customerNote}
              {sale.status === "voided" ? <p className="mt-1 text-xs font-semibold text-brand-red">Voided: {sale.voidReason || "No reason recorded"}</p> : null}
            </td>
            <td className="text-right font-bold">{sale.status === "voided" ? <span className="text-text-secondary line-through">{money(sale.total)}</span> : money(sale.total)}</td>
            <td className="text-right">
              <div className="flex justify-end gap-2">
                {canVoidSale(sale, role, userId, today) ? <button className="danger-btn !min-h-9" onClick={() => onVoid(sale.id)}>Void</button> : <span className="text-xs text-text-secondary">Locked</span>}
                {role === "owner" ? <button className="danger-btn !min-h-9" onClick={() => onDelete(sale.id)}>Delete</button> : null}
              </div>
            </td>
          </tr>
        ))}
      </DataTable>
    </section>
  );
}

function ExpensesView({ form, expenses, role, userId, today, saving, onForm, onSubmit, onExportCsv, onExportExcel, onVoid, onDelete }: {
  form: { date: string; category: string; description: string; amount: string };
  expenses: Expense[];
  role: AppRole;
  userId: string | null;
  today: string;
  saving: boolean;
  onForm: (form: { date: string; category: string; description: string; amount: string }) => void;
  onSubmit: (event: FormEvent) => void;
  onExportCsv: () => void;
  onExportExcel: () => void;
  onVoid: (id: string) => void;
  onDelete: (id: string) => void;
}) {
  return (
    <section>
      <SectionHeader
        title="Expenses"
        action={
          <div className="flex gap-2">
            <button className="secondary-btn" onClick={onExportCsv}>Export CSV</button>
            <button className="secondary-btn" onClick={onExportExcel}>Export Excel</button>
          </div>
        }
      />
      <form onSubmit={onSubmit} className="form-grid">
        <input className="input-field" type="date" value={form.date} onChange={(event) => onForm({ ...form, date: event.target.value })} />
        <input className="input-field" list="expenseCategories" placeholder="Category" value={form.category} onChange={(event) => onForm({ ...form, category: event.target.value })} />
        <datalist id="expenseCategories">{expenseCategories.map((category) => <option key={category} value={category} />)}</datalist>
        <input className="input-field" placeholder="Description" value={form.description} onChange={(event) => onForm({ ...form, description: event.target.value })} />
        <input className="input-field" type="number" min="0.01" step="0.01" placeholder="Amount" value={form.amount} onChange={(event) => onForm({ ...form, amount: event.target.value })} />
        <button className="primary-btn" disabled={saving}>{saving ? "Saving..." : "Add"}</button>
      </form>
      <DataTable empty="No expenses recorded yet.">
        {expenses.map((expense) => (
          <tr key={expense.id}>
            <td>{expense.date}</td>
            <td>{expense.category}</td>
            <td>
              {expense.description}
              {expense.status === "voided" ? <p className="mt-1 text-xs font-semibold text-brand-red">Voided: {expense.voidReason || "No reason recorded"}</p> : null}
            </td>
            <td className="text-right font-bold">{expense.status === "voided" ? <span className="text-text-secondary line-through">{money(expense.amount)}</span> : money(expense.amount)}</td>
            <td className="text-right">
              <div className="flex justify-end gap-2">
                {canVoidExpense(expense, role, userId, today) ? <button className="danger-btn !min-h-9" onClick={() => onVoid(expense.id)}>Void</button> : <span className="text-xs text-text-secondary">Locked</span>}
                {role === "owner" ? <button className="danger-btn !min-h-9" onClick={() => onDelete(expense.id)}>Delete</button> : null}
              </div>
            </td>
          </tr>
        ))}
      </DataTable>
    </section>
  );
}

function ReportsView({ range, report, onRange, onExportExcel, onExportBackup, onImportClick, importRef, onImport }: {
  range: { from: string; to: string };
  report: ReturnType<typeof buildReport>;
  onRange: (range: { from: string; to: string }) => void;
  onExportExcel: () => void;
  onExportBackup: () => void;
  onImportClick: () => void;
  importRef: RefObject<HTMLInputElement | null>;
  onImport: (event: ChangeEvent<HTMLInputElement>) => void;
}) {
  return (
    <section>
      <SectionHeader title="Reports" action={<div className="flex flex-wrap gap-2"><button className="secondary-btn" onClick={onExportExcel}>Export Excel</button><button className="secondary-btn" onClick={onExportBackup}>Backup JSON</button><button className="secondary-btn" onClick={onImportClick}>Restore JSON</button><input ref={importRef} type="file" accept="application/json,.json" onChange={onImport} className="hidden" /></div>} />
      <div className="filter-row">
        <input className="input-field" type="date" value={range.from} onChange={(event) => onRange({ ...range, from: event.target.value })} />
        <input className="input-field" type="date" value={range.to} onChange={(event) => onRange({ ...range, to: event.target.value })} />
        <button className="secondary-btn" onClick={() => onRange({ from: todayKey(), to: todayKey() })}>Today</button>
        <button className="secondary-btn" onClick={() => {
          const now = new Date();
          onRange({ from: dateKey(new Date(now.getFullYear(), now.getMonth(), 1)), to: todayKey() });
        }}>This month</button>
      </div>
      <div className="my-4 grid gap-3 md:grid-cols-5">
        <ReportCard label="Collected" value={money(report.grossSales)} />
        <ReportCard label="Non-income" value={money(report.passThroughFees)} />
        <ReportCard label="Earned" value={money(report.serviceRevenue)} />
        <ReportCard label="Expenses" value={money(report.expenses)} />
        <ReportCard label="Net" value={money(report.netIncome)} />
        <ReportCard label="Transactions" value={String(report.transactions)} />
      </div>
      <div className="grid gap-3 md:grid-cols-2">
        <ListCard title="Top services" rows={report.topServices.map((item) => ({ label: item.name, meta: `${item.quantity} sold`, value: money(item.total) }))} empty="No service sales in this range." />
        <ListCard title="Expense summary" rows={report.expenseSummary.map((item) => ({ label: item.category, value: money(item.total) }))} empty="No expenses in this range." />
      </div>
    </section>
  );
}

function PricesView({ services, form, onForm, onAdd, onSaveSettings, onDelete, onReset }: {
  services: Service[];
  form: { name: string; category: string; optionLabel: string; price: string; baseFee: string; serviceFee: string; requiresTracking: boolean };
  onForm: (form: { name: string; category: string; optionLabel: string; price: string; baseFee: string; serviceFee: string; requiresTracking: boolean }) => void;
  onAdd: (event: FormEvent) => void;
  onSaveSettings: (service: Service, updates: Partial<Service>) => void;
  onDelete: (id: string) => void;
  onReset: () => void;
}) {
  const [drafts, setDrafts] = useState<Record<string, { price: string; baseFee: string; serviceFee: string; requiresTracking: boolean }>>({});
  const categories = useMemo(() => ["all", ...Array.from(new Set(services.map((service) => service.category))).sort()], [services]);
  const [activeCategory, setActiveCategory] = useState("Online Services");
  const makeDraft = (service: Service) => ({
    price: String(service.price ?? 0),
    baseFee: String(service.baseFee ?? 0),
    serviceFee: String(service.serviceFee ?? 0),
    requiresTracking: Boolean(service.requiresTracking),
  });
  const serviceDraft = (service: Service) => drafts[service.id] ?? makeDraft(service);
  const setDraft = (service: Service, patch: Partial<{ price: string; baseFee: string; serviceFee: string; requiresTracking: boolean }>) => {
    setDrafts((current) => ({ ...current, [service.id]: { ...(current[service.id] ?? makeDraft(service)), ...patch } }));
  };
  const selectedCategory = activeCategory === "all" || categories.includes(activeCategory) ? activeCategory : "all";
  const visibleServices = services.filter((service) => selectedCategory === "all" || service.category === selectedCategory);

  return (
    <section>
      <SectionHeader title="Prices" action={<button className="secondary-btn" onClick={onReset}>Reset defaults</button>} />
      <div className="grid gap-4 xl:grid-cols-[0.85fr_1.6fr] xl:items-start">
        <form onSubmit={onAdd} className="section-card grid gap-3 rounded-[18px] p-4">
          <div>
            <h2 className="text-base font-bold">Add service</h2>
            <p className="text-xs font-semibold uppercase tracking-[0.12em] text-text-secondary">Owner setup</p>
          </div>
          <input className="input-field" placeholder="Service name" value={form.name} onChange={(event) => onForm({ ...form, name: event.target.value })} />
          <div className="grid gap-3 sm:grid-cols-2">
            <input className="input-field" placeholder="Category" value={form.category} onChange={(event) => onForm({ ...form, category: event.target.value })} />
            <input className="input-field" placeholder="Option label" value={form.optionLabel} onChange={(event) => onForm({ ...form, optionLabel: event.target.value })} />
          </div>
          <div className="grid gap-3 sm:grid-cols-3">
            <label className="grid gap-1 text-xs font-bold uppercase tracking-[0.08em] text-text-secondary">
              Customer charge
              <input className="input-field !text-sm" type="number" min="0" step="0.01" value={form.price} onChange={(event) => onForm({ ...form, price: event.target.value })} />
            </label>
            <label className="grid gap-1 text-xs font-bold uppercase tracking-[0.08em] text-text-secondary">
              Gov/pass-through
              <input className="input-field !text-sm" type="number" min="0" step="0.01" value={form.baseFee} onChange={(event) => onForm({ ...form, baseFee: event.target.value })} />
            </label>
            <label className="grid gap-1 text-xs font-bold uppercase tracking-[0.08em] text-text-secondary">
              CJNET earns
              <input className="input-field !text-sm" type="number" min="0" step="0.01" value={form.serviceFee} onChange={(event) => onForm({ ...form, serviceFee: event.target.value })} />
            </label>
          </div>
          <label className="flex items-center gap-2 rounded-xl border border-[rgba(23,23,23,0.08)] bg-white px-3 py-2 text-xs font-semibold">
            <input type="checkbox" checked={form.requiresTracking} onChange={(event) => onForm({ ...form, requiresTracking: event.target.checked })} />
            Require reference
          </label>
          <button className="primary-btn">Add service</button>
        </form>

        <div className="section-card rounded-[18px] p-4">
          <div className="mb-3 flex flex-col gap-3 lg:flex-row lg:items-center lg:justify-between">
            <div>
              <h2 className="text-base font-bold">Fee table</h2>
              <p className="text-xs font-semibold uppercase tracking-[0.12em] text-text-secondary">{visibleServices.length} services</p>
            </div>
            <div className="flex flex-wrap gap-2">
              {categories.map((item) => (
                <button
                  key={item}
                  type="button"
                  className={selectedCategory === item ? "primary-btn !min-h-9 !px-3 !py-2 !text-xs" : "secondary-btn !min-h-9 !px-3 !py-2 !text-xs"}
                  onClick={() => setActiveCategory(item)}
                >
                  {item === "all" ? "All" : item}
                </button>
              ))}
            </div>
          </div>
          <div className="grid gap-3">
            {visibleServices.map((service) => {
              const draft = serviceDraft(service);
              const baseFee = Number(draft.baseFee || 0);
              const serviceFee = Number(draft.serviceFee || 0);
              const usesBreakdown = service.category === "Online Services" || baseFee > 0 || serviceFee > 0;
              const computedTotal = usesBreakdown ? baseFee + serviceFee : Number(draft.price || 0);
              return (
                <div key={service.id} data-service-id={service.id} className="price-row grid gap-3 border-b border-[rgba(23,23,23,0.08)] pb-3 last:border-b-0 xl:grid-cols-[minmax(150px,1fr)_minmax(0,2fr)_auto] xl:items-center">
                  <div>
                    <p className="font-semibold">{service.name}</p>
                    <p className="text-sm text-text-secondary">{service.optionLabel} - {service.category}</p>
                  </div>
                  {usesBreakdown ? (
                    <div className="grid gap-2 sm:grid-cols-3">
                      <div className="rounded-xl border border-[rgba(23,23,23,0.08)] bg-[rgba(255,248,230,0.72)] px-3 py-2">
                        <p className="text-[0.68rem] font-bold uppercase tracking-[0.08em] text-text-secondary">Customer charge</p>
                        <p className="font-bold">{money(computedTotal)}</p>
                      </div>
                      <label className="grid gap-1 text-[0.68rem] font-bold uppercase tracking-[0.08em] text-text-secondary">
                        Gov/pass-through
                        <input className="input-field !min-h-10 !text-sm" type="number" min="0" step="0.01" aria-label={`${service.name} bill or pass-through fee`} value={draft.baseFee} onChange={(event) => setDraft(service, { baseFee: event.target.value })} />
                      </label>
                      <label className="grid gap-1 text-[0.68rem] font-bold uppercase tracking-[0.08em] text-text-secondary">
                        CJNET earns
                        <input className="input-field !min-h-10 !text-sm" type="number" min="0" step="0.01" aria-label={`${service.name} earned fee`} value={draft.serviceFee} onChange={(event) => setDraft(service, { serviceFee: event.target.value })} />
                      </label>
                    </div>
                  ) : (
                    <label className="grid gap-1 text-[0.68rem] font-bold uppercase tracking-[0.08em] text-text-secondary">
                      Price
                      <input className="input-field !min-h-10 !text-sm" type="number" min="0" step="0.01" aria-label={`${service.name} price`} value={draft.price} onChange={(event) => setDraft(service, { price: event.target.value })} />
                    </label>
                  )}
                  <div className="grid gap-2 sm:grid-cols-[auto_auto_auto] xl:grid-cols-1">
                    <label className="flex items-center gap-2 text-xs font-semibold text-text-secondary">
                      <input type="checkbox" checked={draft.requiresTracking} onChange={(event) => setDraft(service, { requiresTracking: event.target.checked })} />
                      Ref
                    </label>
                    <button type="button" className="secondary-btn !min-h-10 !py-2" onClick={() => onSaveSettings(service, { price: computedTotal, baseFee: usesBreakdown ? baseFee : 0, serviceFee: usesBreakdown ? serviceFee : 0, requiresTracking: draft.requiresTracking })}>Save</button>
                    <button type="button" className="danger-btn !min-h-10 !py-2" onClick={() => onDelete(service.id)}>Delete</button>
                  </div>
                </div>
              );
            })}
          </div>
          {visibleServices.length === 0 ? (
            <p className="py-8 text-center text-sm text-text-secondary">No services in this category.</p>
          ) : null}
        </div>
      </div>
    </section>
  );
}

function SectionHeader({ title, action }: { title: string; action?: ReactNode }) {
  return (
    <div className="mb-4 flex flex-col gap-3 sm:flex-row sm:items-center sm:justify-between">
      <h1 className="text-2xl font-bold">{title}</h1>
      {action}
    </div>
  );
}

function DataTable({ children, empty }: { children: ReactNode; empty: string }) {
  const rows = Array.isArray(children) ? children.filter(Boolean) : children;
  const isEmpty = Array.isArray(rows) && rows.length === 0;
  return (
    <div className="section-card overflow-auto rounded-[18px]">
      <table className="w-full border-collapse">
        <tbody>
          {isEmpty ? (
            <tr><td className="p-4 text-sm text-text-secondary">{empty}</td></tr>
          ) : (
            children
          )}
        </tbody>
      </table>
    </div>
  );
}

function ReportCard({ label, value, subtext }: { label: string; value: string; subtext?: string }) {
  return (
    <div className="dashboard-card section-card rounded-[18px] p-4">
      <p className="text-xs font-semibold text-text-secondary">{label}</p>
      <p className="mt-2 text-2xl font-bold">{value}</p>
      {subtext ? <p className="mt-1 text-xs font-normal text-text-secondary">{subtext}</p> : null}
    </div>
  );
}

function TrendCard({ rows, range, title = "Sales trend" }: { rows: Array<{ date: string; label: string; total: number }>; range: { from: string; to: string }; title?: string }) {
  const max = Math.max(...rows.map((row) => row.total), 1);
  return (
    <div className="section-card rounded-[18px] p-4">
      <div className="flex items-start justify-between gap-3">
        <div>
          <h2 className="text-base font-bold">{title}</h2>
          <p className="mt-1 text-xs font-normal text-text-secondary">{range.from} to {range.to}</p>
        </div>
        <strong className="text-sm font-bold">{money(rows.reduce((sum, row) => sum + row.total, 0))}</strong>
      </div>
      <div className="mt-4 grid gap-3">
        {rows.map((row) => (
          <div key={row.date} className="grid grid-cols-[4.5rem_minmax(0,1fr)_5.5rem] items-center gap-3">
            <span className="text-xs font-medium text-text-secondary">{row.label}</span>
            <div className="h-3 overflow-hidden rounded-full bg-[rgba(23,23,23,0.06)]">
              <div className="h-full rounded-full bg-brand-yellow" style={{ width: `${Math.max((row.total / max) * 100, row.total > 0 ? 8 : 0)}%` }} />
            </div>
            <span className="text-right text-xs font-bold">{money(row.total)}</span>
          </div>
        ))}
      </div>
    </div>
  );
}

function ListCard({ title, rows, empty }: { title: string; rows: Array<{ label: string; meta?: string; value: string }>; empty: string }) {
  return (
    <div className="section-card rounded-[18px] p-4">
      <h2 className="text-base font-bold">{title}</h2>
      <div className="mt-3 grid gap-2">
        {rows.length ? rows.map((row, index) => (
          <div key={`${row.label}-${row.value}-${index}`} className="flex justify-between gap-3 border-b border-[rgba(23,23,23,0.08)] py-2 last:border-b-0">
            <div><p className="font-medium">{row.label}</p>{row.meta ? <p className="text-xs text-text-secondary">{row.meta}</p> : null}</div>
            <strong>{row.value}</strong>
          </div>
        )) : <p className="text-sm text-text-secondary">{empty}</p>}
      </div>
    </div>
  );
}

function CustomDialog(props: {
  open: boolean;
  name: string;
  category: string;
  price: string;
  baseFee: string;
  serviceFee: string;
  gcashMode: "cash-in" | "cash-out";
  quantity: number;
  onName: (value: string) => void;
  onCategory: (value: string) => void;
  onPrice: (value: string) => void;
  onBaseFee: (value: string) => void;
  onServiceFee: (value: string) => void;
  onGcashMode: (value: "cash-in" | "cash-out") => void;
  onQuantity: (value: number) => void;
  onClose: () => void;
  onAdd: () => void;
}) {
  if (!props.open) return null;
  const gcash = isGcashService(props.name, props.category);
  const isOnline = props.category === "Online Services";
  const amount = Number(props.baseFee || 0);
  const earnedFee = gcash ? gcashServiceFee(amount) : Number(props.serviceFee || 0);
  const computedTotal = Number(props.baseFee || 0) + Number(props.serviceFee || 0);
  return (
    <div className="fixed inset-0 z-50 grid place-items-center bg-black/40 p-4">
      <div className="w-full max-w-lg rounded-[18px] bg-white p-5 shadow-[0_24px_70px_rgba(20,23,31,0.18)]">
        <h2 className="text-lg font-bold">Custom item</h2>
        <div className="mt-4 grid gap-3">
          <input className="input-field" placeholder="Name or service" value={props.name} onChange={(event) => props.onName(event.target.value)} />
          <input className="input-field" placeholder="Category" value={props.category} onChange={(event) => props.onCategory(event.target.value)} />
          {gcash ? (
            <div className="grid gap-3">
              <div className="grid grid-cols-2 gap-2 rounded-2xl border border-[rgba(23,23,23,0.08)] bg-white p-1">
                <button
                  type="button"
                  className={props.gcashMode === "cash-in" ? "primary-btn !min-h-10 !py-2" : "secondary-btn !min-h-10 !py-2"}
                  onClick={() => props.onGcashMode("cash-in")}
                >
                  Cash In
                </button>
                <button
                  type="button"
                  className={props.gcashMode === "cash-out" ? "primary-btn !min-h-10 !py-2" : "secondary-btn !min-h-10 !py-2"}
                  onClick={() => props.onGcashMode("cash-out")}
                >
                  Cash Out
                </button>
              </div>
              <input className="input-field" type="number" min="0.01" step="0.01" placeholder="GCash amount" value={props.baseFee} onChange={(event) => props.onBaseFee(event.target.value)} />
              <div className="grid gap-2 rounded-2xl border border-[rgba(23,23,23,0.08)] bg-[rgba(255,248,230,0.72)] px-4 py-3 text-sm">
                <div className="flex items-center justify-between gap-3">
                  <span className="font-semibold text-text-secondary">Mode</span>
                  <strong>{props.gcashMode === "cash-in" ? "Cash In" : "Cash Out"}</strong>
                </div>
                <div className="flex items-center justify-between gap-3">
                  <span className="font-semibold text-text-secondary">Earned fee</span>
                  <strong>{money(earnedFee)}</strong>
                </div>
                <div className="flex items-center justify-between gap-3">
                  <span className="font-semibold text-text-secondary">Customer charge</span>
                  <strong>{money(amount + earnedFee)}</strong>
                </div>
              </div>
            </div>
          ) : isOnline ? (
            <div className="grid gap-3 sm:grid-cols-2">
              <input className="input-field" type="number" min="0" step="0.01" placeholder="Gov/pass-through fee" value={props.baseFee} onChange={(event) => props.onBaseFee(event.target.value)} />
              <input className="input-field" type="number" min="0" step="0.01" placeholder="CJNET service fee" value={props.serviceFee} onChange={(event) => props.onServiceFee(event.target.value)} />
              <div className="sm:col-span-2 rounded-2xl border border-[rgba(23,23,23,0.08)] bg-[rgba(255,248,230,0.72)] px-4 py-3 text-sm font-bold">
                Customer charge: {money(computedTotal)}
              </div>
            </div>
          ) : (
            <input className="input-field" type="number" min="0.01" step="0.01" placeholder="Price" value={props.price} onChange={(event) => props.onPrice(event.target.value)} />
          )}
          {!gcash ? (
            <div className="grid gap-2">
              <label className="text-xs font-medium text-text-secondary">Pages / quantity</label>
              <input
                className="input-field font-mono text-lg tabular-nums"
                type="number"
                min="1"
                step="1"
                inputMode="numeric"
                value={props.quantity}
                onChange={(event) => props.onQuantity(Math.max(1, Number(event.target.value || 1)))}
              />
              <p className="text-xs leading-5 text-text-secondary">Type the full page count directly, such as 500 or 1200.</p>
            </div>
          ) : null}
          <div className="grid grid-cols-2 gap-3 pt-1">
            <button className="secondary-btn" onClick={props.onClose}>Cancel</button>
            <button className="primary-btn" onClick={props.onAdd}>Add</button>
          </div>
        </div>
      </div>
    </div>
  );
}
