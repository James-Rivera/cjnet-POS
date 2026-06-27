"use client";

import Image from "next/image";
import Link from "next/link";
import { ChangeEvent, FormEvent, RefObject, ReactNode, useEffect, useMemo, useRef, useState } from "react";
import { useRouter } from "next/navigation";
import type { CartItem, Expense, Sale, Service } from "@/types/pos";
import { buildReport, cartSubtotal, changeDue, dateKey, money, saleTotal, todayKey } from "@/lib/pos/calculations";
import { PosRepository } from "@/lib/pos/repository";
import { createSupabaseBrowserClient, hasSupabaseConfig } from "@/lib/pos/supabase-client";

type Tab = "register" | "dashboard" | "audit" | "sales" | "expenses" | "reports" | "prices" | "staff" | "settings";
type DashboardPeriod = "today" | "week" | "month";
type AppRole = "owner" | "staff";
type Status = { tone: "success" | "error" | "info"; message: string } | null;

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

const expenseCategories = ["Ink", "Bond Paper", "Lamination Film", "Internet", "Electricity", "Rent", "Salary"];
const quickCash = [20, 50, 100, 500, 1000];
const roleAccess: Record<AppRole, Tab[]> = {
  owner: ["dashboard", "register", "sales", "expenses", "reports", "prices", "audit", "staff", "settings"],
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
    const total = sales.filter((sale) => sale.date === key).reduce((sum, sale) => sum + sale.total, 0);
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

function needsReference(item: Pick<CartItem, "name" | "category" | "optionLabel">) {
  const text = `${item.name} ${item.category} ${item.optionLabel}`.toLowerCase();
  return item.category === "Online Services" || ["nbi", "police", "psa", "government", "clearance"].some((keyword) => text.includes(keyword));
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

export function PosApp() {
  const repo = useMemo(() => new PosRepository(), []);
  const router = useRouter();
  const supabase = useMemo(() => createSupabaseBrowserClient(), []);
  const [role, setRole] = useState<AppRole>(() => (hasSupabaseConfig() ? "staff" : "owner"));
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
  const [category, setCategory] = useState("all");
  const [discount, setDiscount] = useState(0);
  const [cashReceived, setCashReceived] = useState(0);
  const [customerNote, setCustomerNote] = useState("");
  const [customOpen, setCustomOpen] = useState(false);
  const [customService, setCustomService] = useState<Service | null>(null);
  const [customName, setCustomName] = useState("");
  const [customCategory, setCustomCategory] = useState("Custom");
  const [customPrice, setCustomPrice] = useState("");
  const [customQty, setCustomQty] = useState(1);
  const [expenseForm, setExpenseForm] = useState({ date: todayKey(), category: "", description: "", amount: "" });
  const [serviceForm, setServiceForm] = useState({ name: "", category: "", optionLabel: "", price: "" });
  const [salesFilters, setSalesFilters] = useState({ from: todayKey(), to: todayKey(), search: "" });
  const [reportRange, setReportRange] = useState({ from: todayKey(), to: todayKey() });
  const [dashboardPeriod, setDashboardPeriod] = useState<DashboardPeriod>("today");
  const [auditMonth, setAuditMonth] = useState(currentMonthKey());
  const importRef = useRef<HTMLInputElement>(null);

  const subtotal = cartSubtotal(cart);
  const total = saleTotal(subtotal, discount);
  const change = changeDue(cashReceived, total);
  const today = todayKey();
  const todaySales = sales.filter((sale) => sale.date === today).reduce((sum, sale) => sum + sale.total, 0);
  const todayExpenses = expenses.filter((expense) => expense.date === today).reduce((sum, expense) => sum + expense.amount, 0);
  const todayTransactions = sales.filter((sale) => sale.date === today).length;
  const allowedTabs = tabs.filter((tab) => canAccessTab(role, tab.id));
  const categories = useMemo(() => ["all", ...Array.from(new Set(services.map((service) => service.category)))], [services]);
  const filteredServices = services.filter((service) => {
    const term = search.trim().toLowerCase();
    const haystack = `${service.name} ${service.category} ${service.optionLabel}`.toLowerCase();
    return (category === "all" || service.category === category) && (!term || haystack.includes(term));
  });
  const filteredSales = sales.filter((sale) => {
    const saleText = `${sale.customerNote} ${sale.items.map((item) => `${item.serviceName} ${item.optionLabel}`).join(" ")}`.toLowerCase();
    const rangeAllowed = role === "owner" || sale.date === today;
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
  const saleNeedsReference = cart.some(needsReference);
  const visibleTab = canAccessTab(role, activeTab) ? activeTab : role === "owner" ? "dashboard" : "register";
  const accessAllowed = canAccessTab(role, visibleTab);

  useEffect(() => {
    if (!supabase) return;

    let cancelled = false;

    async function loadAuth() {
      const { data: authData, error: authError } = await supabase.auth.getUser();
      if (authError || !authData.user) {
        router.replace("/login");
        return;
      }

      const { data: profile, error: profileError } = await supabase
        .from("profiles")
        .select("full_name, role, status")
        .eq("id", authData.user.id)
        .single();

      if (profileError || !profile || profile.status === "disabled") {
        await supabase.auth.signOut();
        router.replace("/login?error=access_denied");
        return;
      }

      if (cancelled) return;
      setRole(profile.role === "owner" ? "owner" : "staff");
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
        const snapshot = await repo.loadSnapshot();
        setServices(snapshot.services);
        setSales(snapshot.sales);
        setExpenses(snapshot.expenses);
      } catch (error) {
        setStatus({ tone: "error", message: error instanceof Error ? error.message : "Could not load POS data." });
      } finally {
        setLoading(false);
      }
    }
    load();
  }, [repo]);

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
        },
      ];
    });
  }

  function onServiceClick(service: Service) {
    if (service.isCustomPrice || service.price <= 0) {
      setCustomService(service);
      setCustomName(service.name);
      setCustomCategory(service.category);
      setCustomPrice("");
      setCustomQty(1);
      setCustomOpen(true);
      return;
    }
    addToCart(service);
  }

  async function saveSale() {
    if (!cart.length) {
      setStatus({ tone: "error", message: "Add items first." });
      return;
    }
    if (saleNeedsReference && !customerNote.trim()) {
      setStatus({ tone: "error", message: "Reference required for government services." });
      return;
    }
    try {
      setSaving(true);
      const sale = await repo.saveSale({ cart, discount, cashReceived, customerNote: saleNeedsReference ? customerNote.trim() : "" });
      setSales((current) => [sale, ...current]);
      setCart([]);
      setDiscount(0);
      setCashReceived(0);
      setCustomerNote("");
      setStatus({ tone: "success", message: `Sale saved: ${money(sale.total)}` });
    } catch (error) {
      setStatus({ tone: "error", message: error instanceof Error ? error.message : "Could not save sale." });
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
      const expense = await repo.saveExpense({
        date: expenseForm.date || todayKey(),
        category: expenseForm.category.trim(),
        description: expenseForm.description.trim(),
        amount,
      });
      setExpenses((current) => [expense, ...current]);
      setExpenseForm({ date: todayKey(), category: "", description: "", amount: "" });
      setStatus({ tone: "success", message: "Expense saved." });
    } catch (error) {
      setStatus({ tone: "error", message: error instanceof Error ? error.message : "Could not save expense." });
    }
  }

  async function saveServicePrice(service: Service, price: number) {
    try {
      const updated = await repo.saveService({ ...service, price, isCustomPrice: price <= 0 });
      setServices((current) => current.map((item) => (item.id === updated.id ? updated : item)));
      setStatus({ tone: "success", message: "Price saved." });
    } catch (error) {
      setStatus({ tone: "error", message: error instanceof Error ? error.message : "Could not save price." });
    }
  }

  async function addService(event: FormEvent) {
    event.preventDefault();
    if (!serviceForm.name.trim() || !serviceForm.category.trim() || !serviceForm.optionLabel.trim()) {
      setStatus({ tone: "error", message: "Complete the service details." });
      return;
    }
    const service: Service = {
      id: `service-${Date.now()}`,
      name: serviceForm.name.trim(),
      category: serviceForm.category.trim(),
      optionLabel: serviceForm.optionLabel.trim(),
      price: Number(serviceForm.price || 0),
      isCustomPrice: Number(serviceForm.price || 0) <= 0,
      sortOrder: services.length + 100,
    };
    const saved = await repo.saveService(service);
    setServices((current) => [...current, saved]);
    setServiceForm({ name: "", category: "", optionLabel: "", price: "" });
    setStatus({ tone: "success", message: "Service added." });
  }

  async function deleteSale(id: string) {
    if (!window.confirm("Delete this sale?")) return;
    await repo.deleteSale(id);
    setSales((current) => current.filter((sale) => sale.id !== id));
    setStatus({ tone: "success", message: "Sale deleted." });
  }

  async function deleteExpense(id: string) {
    if (!window.confirm("Delete this expense?")) return;
    await repo.deleteExpense(id);
    setExpenses((current) => current.filter((expense) => expense.id !== id));
    setStatus({ tone: "success", message: "Expense deleted." });
  }

  async function deleteService(id: string) {
    if (!window.confirm("Delete this service?")) return;
    await repo.deleteService(id);
    setServices((current) => current.filter((service) => service.id !== id));
    setStatus({ tone: "success", message: "Service deleted." });
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
        ["Gross sales", report.grossSales],
        ["Expenses", report.expenses],
        ["Net income", report.netIncome],
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
        ["Gross sales", auditReport.grossSales],
        ["Expenses", auditReport.expenses],
        ["Net income", auditReport.netIncome],
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
      <section className="mx-auto flex w-full max-w-[1440px] flex-col gap-4">
        <header className="rounded-[1.35rem] border border-surface-border bg-white px-4 py-3 shadow-[var(--shadow-soft)]">
          <div className="flex flex-col gap-3 lg:flex-row lg:items-center lg:justify-between">
            <div className="flex flex-wrap items-center gap-3">
              <Link href="/" className="utility-logo-link w-[158px]">
                <Image src="/logo.png" alt="CJ NET" width={920} height={311} className="h-auto w-full" priority />
              </Link>
              <div className="rounded-full border border-surface-border bg-[rgba(255,212,0,0.14)] px-3 py-1 text-xs font-semibold text-foreground">
                {accountName} · {role === "owner" ? "Owner" : "Staff"}
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

        <section className="summary-bar grid gap-3 rounded-[1.35rem] border border-surface-border bg-[rgba(255,255,255,0.72)] p-3 shadow-[var(--shadow-card)] md:grid-cols-4">
          <SummaryTile label="Today" value={new Date().toLocaleDateString([], { weekday: "long", month: "long", day: "numeric", year: "numeric" })} subtext={`${todayTransactions} transactions`} />
          <SummaryTile label="Sales" value={money(todaySales)} subtext="Today's gross sales" />
          <SummaryTile label="Expenses" value={money(todayExpenses)} subtext="Recorded today" />
          <SummaryTile label="Net" value={money(todaySales - todayExpenses)} subtext="Sales minus expenses" />
        </section>

        {!accessAllowed ? <AccessDenied role={role} /> : null}

        {accessAllowed && visibleTab === "register" ? (
          <RegisterView
            filteredServices={filteredServices}
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
            requiresReference={saleNeedsReference}
            saving={saving}
            onCategory={setCategory}
            onSearch={setSearch}
            onServiceClick={onServiceClick}
            onCart={setCart}
            onDiscount={setDiscount}
            onCash={setCashReceived}
            onCustomerNote={setCustomerNote}
            onClear={() => {
              setCart([]);
              setDiscount(0);
              setCashReceived(0);
              setCustomerNote("");
            }}
            onSave={saveSale}
            onCustom={() => {
              setCustomOpen(true);
              setCustomService(null);
              setCustomName("");
              setCustomCategory("Custom");
              setCustomPrice("");
              setCustomQty(1);
            }}
          />
        ) : null}

        {accessAllowed && visibleTab === "dashboard" ? (
          <DashboardView
            period={dashboardPeriod}
            range={dashboardRange}
            todaySales={todaySales}
            todayExpenses={todayExpenses}
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

        {accessAllowed && visibleTab === "sales" ? <SalesView filters={salesFilters} sales={filteredSales} role={role} onFilters={setSalesFilters} onExport={exportSalesCsv} onDelete={deleteSale} /> : null}
        {accessAllowed && visibleTab === "expenses" ? (
          <ExpensesView form={expenseForm} expenses={expenses} role={role} onForm={setExpenseForm} onSubmit={saveExpense} onExportCsv={exportExpensesCsv} onExportExcel={exportExpensesExcel} onDelete={deleteExpense} />
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
            onSave={saveServicePrice}
            onDelete={deleteService}
            onReset={async () => {
              if (!window.confirm("Reset services to default prices?")) return;
              await repo.resetServices();
              const snapshot = await repo.loadSnapshot();
              setServices(snapshot.services);
              setStatus({ tone: "success", message: "Default prices restored." });
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
          quantity={customQty}
          onName={setCustomName}
          onCategory={setCustomCategory}
          onPrice={setCustomPrice}
          onQuantity={setCustomQty}
          onClose={() => {
            setCustomOpen(false);
            setCustomService(null);
            setCustomName("");
          }}
          onAdd={() => {
            const price = Number(customPrice);
            if (!customName.trim() || price <= 0) {
              setStatus({ tone: "error", message: "Enter a name and price." });
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
            addToCart({ ...service, name: customName.trim(), category: customCategory.trim() || "Custom" }, Math.max(customQty, 1), price, customName.trim(), "Custom");
            setCustomOpen(false);
            setCustomService(null);
            setCustomName("");
          }}
        />
      ) : null}

      {status ? <StatusToast status={status} /> : null}
    </main>
  );
}

function SummaryTile({ label, value, subtext }: { label: string; value: string; subtext: string }) {
  return (
    <div className="summary-tile rounded-2xl bg-white px-4 py-3">
      <p className="text-[11px] font-semibold uppercase tracking-[0.08em] text-text-secondary">{label}</p>
      <p className="mt-1 truncate text-base font-bold text-foreground">{value}</p>
      <p className="mt-1 truncate text-[11px] font-medium text-text-secondary">{subtext}</p>
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

function DashboardView({ period, range, todaySales, todayExpenses, todayTransactions, report, trend, recentSales, cashierActivity, onPeriod }: {
  period: DashboardPeriod;
  range: { from: string; to: string };
  todaySales: number;
  todayExpenses: number;
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

      <div className="grid gap-3 md:grid-cols-4">
        <ReportCard label="Today's sales" value={money(todaySales)} subtext="Cash collected today" />
        <ReportCard label="Today's expenses" value={money(todayExpenses)} subtext="Shop costs recorded" />
        <ReportCard label="Net income" value={money(todaySales - todayExpenses)} subtext="Sales minus expenses" />
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

      <div className="grid gap-3 md:grid-cols-5">
        <ReportCard label="Gross sales" value={money(report.grossSales)} />
        <ReportCard label="Expenses" value={money(report.expenses)} />
        <ReportCard label="Net income" value={money(report.netIncome)} />
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
          { label: "Sales after discounts", value: money(report.grossSales) },
          { label: "Recorded expenses", value: money(report.expenses) },
          { label: "Net business income", value: money(report.netIncome) },
          { label: "Sales count", value: String(report.transactions) },
        ]} empty="No audit data." />
      </div>
    </section>
  );
}

function RegisterView(props: {
  filteredServices: Service[];
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
  requiresReference: boolean;
  saving: boolean;
  onCategory: (value: string) => void;
  onSearch: (value: string) => void;
  onServiceClick: (service: Service) => void;
  onCart: (updater: (current: CartItem[]) => CartItem[]) => void;
  onDiscount: (value: number) => void;
  onCash: (value: number) => void;
  onCustomerNote: (value: string) => void;
  onClear: () => void;
  onSave: () => void;
  onCustom: () => void;
}) {
  return (
    <section>
      <h1 className="mb-4 text-2xl font-bold text-foreground">Register</h1>
      <div className="grid gap-4 lg:grid-cols-[minmax(0,1fr)_390px]">
        <section className="section-card rounded-[18px] p-6">
          <div className="mb-4 flex items-center justify-between gap-4">
            <h2 className="text-base font-bold">Services</h2>
            <button type="button" onClick={props.onCustom} className="secondary-btn !py-3">
              Custom
            </button>
          </div>

          <div className="mb-3 grid gap-3 md:grid-cols-[minmax(0,1fr)_220px]">
            <input className="input-field" value={props.search} onChange={(event) => props.onSearch(event.target.value)} placeholder="Search service" />
            <select className="input-field" value={props.category} onChange={(event) => props.onCategory(event.target.value)}>
              {props.categories.map((category) => (
                <option key={category} value={category}>
                  {category === "all" ? "All services" : category}
                </option>
              ))}
            </select>
          </div>

          <div className="mb-4 flex gap-2 overflow-x-auto pb-1">
            {props.categories.map((category) => (
              <button key={category} type="button" onClick={() => props.onCategory(category)} className={category === props.category ? "category-pill category-pill-active" : "category-pill"}>
                {category === "all" ? "All" : category}
              </button>
            ))}
          </div>

          <div className="grid grid-cols-[repeat(auto-fill,minmax(150px,1fr))] gap-3">
            {props.filteredServices.map((service) => (
              <ServiceCard key={service.id} service={service} onClick={() => props.onServiceClick(service)} />
            ))}
          </div>
        </section>

        <aside className="checkout-panel section-card sticky top-4 flex max-h-[calc(100vh-32px)] min-h-[640px] flex-col rounded-[18px] p-6 max-lg:static max-lg:max-h-none">
          <div className="receipt-header flex items-start justify-between gap-4">
            <div>
              <h2 className="text-base font-bold">Current Sale</h2>
              <p className="receipt-count mt-2 text-sm text-text-secondary">{props.cart.reduce((sum, item) => sum + item.quantity, 0)} items in receipt</p>
            </div>
            <button type="button" onClick={props.onClear} className="danger-btn !py-3">
              Clear
            </button>
          </div>

          <div className={props.cart.length ? "receipt-list mt-4 grid gap-3 overflow-auto pr-1" : "receipt-empty mt-4 grid min-h-44 place-items-center rounded-2xl border-2 border-dashed border-[rgba(23,23,23,0.12)] bg-white/70 p-6 text-center text-sm text-text-secondary"}>
            {props.cart.length ? (
              props.cart.map((item) => <CartLine key={item.id} item={item} onCart={props.onCart} />)
            ) : (
              <div className="grid justify-items-center gap-2">
                <p className="font-semibold text-foreground">No items yet</p>
                <p className="max-w-48 text-xs leading-5 text-text-secondary">Tap a service to start the receipt.</p>
              </div>
            )}
          </div>

          <div className="mt-auto grid gap-3 pt-4">
            {props.requiresReference ? (
              <label className="block">
                <span className="mb-2 block text-xs font-medium text-text-secondary">Government reference</span>
                <input className="input-field" value={props.customerNote} onChange={(event) => props.onCustomerNote(event.target.value)} placeholder="Name, reference, remarks" />
              </label>
            ) : null}

            <div className="payment-panel subtle-panel grid gap-3 rounded-2xl p-4">
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
    <button type="button" data-service-id={service.id} onClick={onClick} className={`service-card ${compact ? "h-[104px]" : "h-[116px]"}`}>
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

function CartLine({ item, onCart }: { item: CartItem; onCart: (updater: (current: CartItem[]) => CartItem[]) => void }) {
  return (
    <div className="rounded-2xl border border-[rgba(23,23,23,0.08)] bg-white p-3">
      <div className="flex justify-between gap-3">
        <div className="min-w-0">
          <p className="truncate text-sm font-semibold">{item.name}</p>
          <p className="mt-1 text-xs text-text-secondary">
            {item.category} / {item.optionLabel} / {money(item.price)} each
          </p>
        </div>
        <p className="text-sm font-bold">{money(item.price * item.quantity)}</p>
      </div>
      <div className="mt-3 flex items-center justify-between gap-3">
        <button type="button" className="danger-btn !min-h-9 !px-3" onClick={() => onCart((current) => current.filter((cartItem) => cartItem.id !== item.id))}>
          Remove
        </button>
        <div className="flex items-center gap-2">
          <button type="button" className="qty-btn" onClick={() => onCart((current) => current.map((cartItem) => (cartItem.id === item.id ? { ...cartItem, quantity: Math.max(1, cartItem.quantity - 1) } : cartItem)))}>
            -
          </button>
          <span className="w-8 text-center text-sm font-bold">{item.quantity}</span>
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

function SalesView({ filters, sales, role, onFilters, onExport, onDelete }: {
  filters: { from: string; to: string; search: string };
  sales: Sale[];
  role: AppRole;
  onFilters: (filters: { from: string; to: string; search: string }) => void;
  onExport: () => void;
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
            <td>{sale.customerNote}</td>
            <td className="text-right font-bold">{money(sale.total)}</td>
            <td className="text-right">{role === "owner" ? <button className="danger-btn !min-h-9" onClick={() => onDelete(sale.id)}>Delete</button> : <span className="text-xs text-text-secondary">Locked</span>}</td>
          </tr>
        ))}
      </DataTable>
    </section>
  );
}

function ExpensesView({ form, expenses, role, onForm, onSubmit, onExportCsv, onExportExcel, onDelete }: {
  form: { date: string; category: string; description: string; amount: string };
  expenses: Expense[];
  role: AppRole;
  onForm: (form: { date: string; category: string; description: string; amount: string }) => void;
  onSubmit: (event: FormEvent) => void;
  onExportCsv: () => void;
  onExportExcel: () => void;
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
        <button className="primary-btn">Add</button>
      </form>
      <DataTable empty="No expenses recorded yet.">
        {expenses.map((expense) => (
          <tr key={expense.id}>
            <td>{expense.date}</td>
            <td>{expense.category}</td>
            <td>{expense.description}</td>
            <td className="text-right font-bold">{money(expense.amount)}</td>
            <td className="text-right">{role === "owner" ? <button className="danger-btn !min-h-9" onClick={() => onDelete(expense.id)}>Delete</button> : <span className="text-xs text-text-secondary">Locked</span>}</td>
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
      <div className="my-4 grid gap-3 md:grid-cols-4">
        <ReportCard label="Gross sales" value={money(report.grossSales)} />
        <ReportCard label="Expenses" value={money(report.expenses)} />
        <ReportCard label="Net income" value={money(report.netIncome)} />
        <ReportCard label="Transactions" value={String(report.transactions)} />
      </div>
      <div className="grid gap-3 md:grid-cols-2">
        <ListCard title="Top services" rows={report.topServices.map((item) => ({ label: item.name, meta: `${item.quantity} sold`, value: money(item.total) }))} empty="No service sales in this range." />
        <ListCard title="Expense summary" rows={report.expenseSummary.map((item) => ({ label: item.category, value: money(item.total) }))} empty="No expenses in this range." />
      </div>
    </section>
  );
}

function PricesView({ services, form, onForm, onAdd, onSave, onDelete, onReset }: {
  services: Service[];
  form: { name: string; category: string; optionLabel: string; price: string };
  onForm: (form: { name: string; category: string; optionLabel: string; price: string }) => void;
  onAdd: (event: FormEvent) => void;
  onSave: (service: Service, price: number) => void;
  onDelete: (id: string) => void;
  onReset: () => void;
}) {
  const [drafts, setDrafts] = useState<Record<string, string>>({});
  return (
    <section>
      <SectionHeader title="Prices" action={<button className="secondary-btn" onClick={onReset}>Reset defaults</button>} />
      <form onSubmit={onAdd} className="form-grid">
        <input className="input-field" placeholder="Service name" value={form.name} onChange={(event) => onForm({ ...form, name: event.target.value })} />
        <input className="input-field" placeholder="Category" value={form.category} onChange={(event) => onForm({ ...form, category: event.target.value })} />
        <input className="input-field" placeholder="Option label" value={form.optionLabel} onChange={(event) => onForm({ ...form, optionLabel: event.target.value })} />
        <input className="input-field" type="number" min="0" step="0.01" placeholder="Price" value={form.price} onChange={(event) => onForm({ ...form, price: event.target.value })} />
        <button className="primary-btn">Add</button>
      </form>
      <div className="section-card rounded-[18px] p-4">
        {services.map((service) => (
          <div key={service.id} data-service-id={service.id} className="price-row grid gap-3 border-b border-[rgba(23,23,23,0.08)] py-3 last:border-b-0 md:grid-cols-[minmax(0,1fr)_120px_auto_auto] md:items-center">
            <div>
              <p className="font-semibold">{service.name} - {service.optionLabel}</p>
              <p className="text-sm text-text-secondary">{service.category}</p>
            </div>
            <input className="input-field" type="number" min="0" step="0.01" value={drafts[service.id] ?? service.price} onChange={(event) => setDrafts((current) => ({ ...current, [service.id]: event.target.value }))} />
            <button className="secondary-btn" onClick={() => onSave(service, Number(drafts[service.id] ?? service.price))}>Save</button>
            <button className="danger-btn" onClick={() => onDelete(service.id)}>Delete</button>
          </div>
        ))}
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
        {rows.length ? rows.map((row) => (
          <div key={`${row.label}-${row.value}`} className="flex justify-between gap-3 border-b border-[rgba(23,23,23,0.08)] py-2 last:border-b-0">
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
  quantity: number;
  onName: (value: string) => void;
  onCategory: (value: string) => void;
  onPrice: (value: string) => void;
  onQuantity: (value: number) => void;
  onClose: () => void;
  onAdd: () => void;
}) {
  if (!props.open) return null;
  return (
    <div className="fixed inset-0 z-50 grid place-items-center bg-black/40 p-4">
      <div className="w-full max-w-md rounded-[18px] bg-white p-5 shadow-[0_24px_70px_rgba(20,23,31,0.18)]">
        <h2 className="text-lg font-bold">Custom item</h2>
        <div className="mt-4 grid gap-3">
          <input className="input-field" placeholder="Name or service" value={props.name} onChange={(event) => props.onName(event.target.value)} />
          <input className="input-field" placeholder="Category" value={props.category} onChange={(event) => props.onCategory(event.target.value)} />
          <input className="input-field" type="number" min="0.01" step="0.01" placeholder="Price" value={props.price} onChange={(event) => props.onPrice(event.target.value)} />
          <input className="input-field" type="number" min="1" step="1" value={props.quantity} onChange={(event) => props.onQuantity(Math.max(1, Number(event.target.value || 1)))} />
          <div className="grid grid-cols-2 gap-3">
            <button className="secondary-btn" onClick={props.onClose}>Cancel</button>
            <button className="primary-btn" onClick={props.onAdd}>Add</button>
          </div>
        </div>
      </div>
    </div>
  );
}
