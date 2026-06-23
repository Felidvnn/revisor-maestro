"use client";

import dynamic from "next/dynamic";
import Link from "next/link";
import type { ReactNode } from "react";
import { useMemo, useState } from "react";
import {
  AlertTriangle,
  ArrowDownWideNarrow,
  BarChart3,
  CheckCircle2,
  Download,
  Eye,
  EyeOff,
  FileSpreadsheet,
  Filter,
  Layers3,
  LocateFixed,
  Lock,
  Map as MapIcon,
  MapPinned,
  Search,
  Target,
  TrendingUp,
  Unlock,
  UploadCloud,
  Route
} from "lucide-react";
import clsx from "clsx";
import {
  buildReviewWorkbook,
  buildTemplateWorkbook,
  compareClients,
  downloadWorkbook,
  formatDays,
  normalizeText,
  parseClients,
  parseExcelRows,
  parseKml,
  parseSchedules,
  type AnalysisMode,
  type ClientRecord,
  type ReviewRecord,
  type ReviewStatus,
  type ZonePolygon,
  type ZoneSchedule
} from "@/lib/reviewer";

const ZoneMap = dynamic(() => import("@/components/ZoneMap"), {
  ssr: false,
  loading: () => <div className="grid h-full min-h-[460px] place-items-center text-sm text-slate-500">Preparando mapa...</div>
});

type FileState = {
  master?: string;
  kml?: string;
  schedule?: string;
};

type SortKey = "priority" | "kilosDesc" | "kilosAsc" | "boxesDesc" | "nameAsc" | "status" | "zoneAsc";
type AppTab = "map" | "analysis";

const statusCopy: Record<ReviewStatus, string> = {
  match: "OK",
  mismatch: "Alerta",
  noSchedule: "Sin calendario",
  noLocation: "Sin ubicacion",
  outsideZone: "Fuera de zona",
  noKml: "Sin KML"
};

const modeCopy: Record<AnalysisMode, string> = {
  geo: "Zona KML",
  days: "Dias",
  combined: "Ambos"
};

export default function DashboardApp() {
  const [clients, setClients] = useState<ClientRecord[]>([]);
  const [schedules, setSchedules] = useState<ZoneSchedule[]>([]);
  const [zones, setZones] = useState<ZonePolygon[]>([]);
  const [files, setFiles] = useState<FileState>({});
  const [analysisMode, setAnalysisMode] = useState<AnalysisMode>("geo");
  const [activeTab, setActiveTab] = useState<AppTab>("map");
  const [showZones, setShowZones] = useState(true);
  const [lockMapView, setLockMapView] = useState(false);
  const [sortKey, setSortKey] = useState<SortKey>("priority");
  const [office, setOffice] = useState("all");
  const [zone, setZone] = useState("all");
  const [customerType, setCustomerType] = useState("all");
  const [status, setStatus] = useState<ReviewStatus | "all">("all");
  const [query, setQuery] = useState("");
  const [error, setError] = useState("");

  const reviewed = useMemo(() => compareClients(clients, schedules, zones, analysisMode), [analysisMode, clients, schedules, zones]);
  const filtered = useMemo(() => {
    const needle = normalizeText(query);
    return reviewed.filter((record) => {
      const matchesOffice = office === "all" || record.office === office;
      const matchesZone = zone === "all" || record.zoneCode === zone;
      const matchesType = customerType === "all" || record.customerType === customerType;
      const matchesStatus = status === "all" || record.status === status;
      const matchesQuery =
        !needle ||
        normalizeText(`${record.id} ${record.name} ${record.address} ${record.population} ${record.zoneName} ${record.customerType}`).includes(needle);
      return matchesOffice && matchesZone && matchesType && matchesStatus && matchesQuery;
    });
  }, [customerType, office, query, reviewed, status, zone]);

  const summary = useMemo(() => getSummary(reviewed), [reviewed]);
  const calendarSummary = useMemo(() => getCalendarSummary(schedules), [schedules]);
  const offices = useMemo(() => uniqueSorted(clients.map((client) => client.office)), [clients]);
  const customerTypes = useMemo(() => uniqueSorted(clients.map((client) => client.customerType)), [clients]);
  const zonesForFilter = useMemo(
    () =>
      uniqueSorted(clients.map((client) => client.zoneCode).filter(Boolean)).map((zoneCode) => ({
        code: zoneCode,
        name: clients.find((client) => client.zoneCode === zoneCode)?.zoneName || zoneCode
      })),
    [clients]
  );

  async function handleMaster(file: File) {
    try {
      setError("");
      const rows = await parseExcelRows(file);
      const parsed = parseClients(rows);
      setClients(parsed);
      setFiles((current) => ({ ...current, master: file.name }));
    } catch {
      setError("No pude leer el Excel de maestro. Revisa que sea .xlsx y que tenga encabezados.");
    }
  }

  async function handleSchedule(file: File) {
    try {
      setError("");
      const rows = await parseExcelRows(file);
      const parsed = parseSchedules(rows);
      setSchedules(parsed);
      setFiles((current) => ({ ...current, schedule: file.name }));
    } catch {
      setError("No pude leer el calendario de zonas. Puedes usar la plantilla generada por la app.");
    }
  }

  async function handleKml(file: File) {
    try {
      setError("");
      const text = await file.text();
      setZones(parseKml(text));
      setFiles((current) => ({ ...current, kml: file.name }));
    } catch {
      setError("No pude leer el KML. Revisa que el archivo tenga poligonos de zonas.");
    }
  }

  function downloadTemplate() {
    const baseZones = getTemplateZones(clients, zones);
    if (!baseZones.length) {
      window.location.href = "/plantilla_zonas_reparto.xlsx";
      return;
    }
    downloadWorkbook(buildTemplateWorkbook(baseZones), "plantilla_zonas_reparto.xlsx");
  }

  function downloadReview() {
    const workbook = buildReviewWorkbook(sortRecords(filtered, sortKey), analysisMode);
    const stamp = new Date().toISOString().slice(0, 10);
    downloadWorkbook(workbook, `consolidado_revisor_maestro_${analysisMode}_${stamp}.xlsx`);
  }

  return (
    <main className="min-h-screen bg-slate-50 text-slate-950">
      <div className="mx-auto flex w-full max-w-[1560px] flex-col gap-5 px-4 py-4 sm:px-6 lg:px-8">
        <header className="flex flex-col gap-4 rounded-lg border border-slate-200 bg-white px-5 py-4 shadow-panel lg:flex-row lg:items-center lg:justify-between">
          <div>
            <p className="text-xs font-semibold uppercase tracking-[0.18em] text-teal-700">Revisor maestro</p>
            <h1 className="mt-1 text-2xl font-semibold tracking-normal text-slate-950 sm:text-3xl">Control de dias de despacho por zona</h1>
            <p className="mt-2 max-w-3xl text-sm leading-6 text-slate-600">
              Valida clientes contra su zona geografica KML, calendario de visita o ambos criterios a la vez.
            </p>
          </div>
          <div className="flex flex-col gap-2 sm:flex-row">
            <Link
              href="/rutas"
              className="inline-flex h-11 items-center justify-center gap-2 rounded-md bg-white px-4 text-sm font-semibold text-slate-800 ring-1 ring-slate-200 transition hover:bg-slate-50"
            >
              <Route className="h-4 w-4" />
              Analisis OSRM
            </Link>
            <button
              onClick={downloadReview}
              disabled={!filtered.length}
              className="inline-flex h-11 items-center justify-center gap-2 rounded-md bg-teal-700 px-4 text-sm font-semibold text-white transition hover:bg-teal-800 disabled:cursor-not-allowed disabled:bg-slate-300"
            >
              <Download className="h-4 w-4" />
              Exportar Excel
            </button>
            <button
              onClick={downloadTemplate}
              className="inline-flex h-11 items-center justify-center gap-2 rounded-md bg-slate-950 px-4 text-sm font-semibold text-white transition hover:bg-slate-800"
            >
              <Download className="h-4 w-4" />
              Plantilla zonas
            </button>
          </div>
        </header>

        <section className="grid gap-3 lg:grid-cols-3">
          <UploadCard
            title="Maestro clientes"
            icon={<FileSpreadsheet className="h-5 w-5" />}
            filename={files.master}
            accept=".xlsx,.xls"
            onFile={handleMaster}
          />
          <UploadCard
            title="Zonas KML"
            icon={<MapPinned className="h-5 w-5" />}
            filename={files.kml}
            accept=".kml"
            onFile={handleKml}
          />
          <UploadCard
            title="Calendario zonas"
            icon={<UploadCloud className="h-5 w-5" />}
            filename={files.schedule}
            accept=".xlsx,.xls"
            onFile={handleSchedule}
          />
        </section>

        {error && (
          <div className="flex items-center gap-2 rounded-md border border-red-200 bg-red-50 px-4 py-3 text-sm text-red-700">
            <AlertTriangle className="h-4 w-4" />
            {error}
          </div>
        )}

        {files.schedule && calendarSummary.total > 0 && calendarSummary.active === 0 && (
          <div className="flex items-center gap-2 rounded-md border border-amber-200 bg-amber-50 px-4 py-3 text-sm text-amber-800">
            <AlertTriangle className="h-4 w-4" />
            El calendario cargado tiene zonas, pero ninguna tiene dias activos.
          </div>
        )}

        <section className="rounded-lg border border-slate-200 bg-white p-3 shadow-panel">
          <div className="flex flex-col gap-3 lg:flex-row lg:items-center lg:justify-between">
            <div>
              <p className="text-sm font-semibold">Tipo de analisis</p>
              <p className="text-sm text-slate-500">{modeDescription(analysisMode)}</p>
            </div>
            <ModeSwitch value={analysisMode} onChange={setAnalysisMode} />
          </div>
        </section>

        <TabSwitch value={activeTab} onChange={setActiveTab} />

        {activeTab === "map" && (
          <section className="grid gap-3 sm:grid-cols-2 xl:grid-cols-5">
            <Metric label="Clientes" value={summary.total.toLocaleString("es-CL")} icon={<FileSpreadsheet className="h-4 w-4" />} />
            <Metric label="Alertas" value={summary.alerts.toLocaleString("es-CL")} tone="danger" icon={<AlertTriangle className="h-4 w-4" />} />
            <Metric label="Conformidad" value={`${summary.compliance}%`} tone="success" icon={<CheckCircle2 className="h-4 w-4" />} />
            <Metric label={pendingLabel(analysisMode)} value={summary.pending.toLocaleString("es-CL")} tone="warn" icon={<Filter className="h-4 w-4" />} />
            <Metric label={referenceLabel(analysisMode)} value={referenceValue(analysisMode, zones.length, calendarSummary)} icon={<MapIcon className="h-4 w-4" />} />
          </section>
        )}

        {activeTab === "map" ? (
          <section className="grid gap-4 xl:h-[780px] xl:grid-cols-[minmax(0,1fr)_430px]">
            <div className="flex min-h-[620px] flex-col overflow-hidden rounded-lg border border-slate-200 bg-white shadow-panel xl:h-full xl:min-h-0">
              <div className="flex min-h-[73px] shrink-0 flex-col gap-3 border-b border-slate-200 px-4 py-3 lg:flex-row lg:items-center lg:justify-between">
                <div>
                  <h2 className="text-base font-semibold">Mapa operacional</h2>
                  <p className="text-sm text-slate-500">{filtered.length.toLocaleString("es-CL")} clientes visibles - {modeCopy[analysisMode]}</p>
                </div>
                <Legend />
              </div>
              <div className="min-h-[560px] flex-1">
                <ZoneMap clients={filtered} zones={zones} showZones={showZones} fitEnabled={!lockMapView} />
              </div>
            </div>

            <aside className="grid min-h-[680px] grid-rows-[auto_minmax(0,1fr)] gap-4 xl:h-full xl:min-h-0">
              <FilterPanel
                office={office}
                setOffice={setOffice}
                offices={offices}
                customerType={customerType}
                setCustomerType={setCustomerType}
                customerTypes={customerTypes}
                zone={zone}
                setZone={setZone}
                zonesForFilter={zonesForFilter}
                status={status}
                setStatus={setStatus}
                sortKey={sortKey}
                setSortKey={setSortKey}
                showZones={showZones}
                setShowZones={setShowZones}
                lockMapView={lockMapView}
                setLockMapView={setLockMapView}
                query={query}
                setQuery={setQuery}
              />

              <ReviewList records={filtered} mode={analysisMode} sortKey={sortKey} />
            </aside>
          </section>
        ) : (
          <>
            <AnalysisFilterBar
              office={office}
              setOffice={setOffice}
              offices={offices}
              customerType={customerType}
              setCustomerType={setCustomerType}
              customerTypes={customerTypes}
              status={status}
              setStatus={setStatus}
              query={query}
              setQuery={setQuery}
            />
            <AnalysisPanel records={filtered} mode={analysisMode} />
          </>
        )}
      </div>
    </main>
  );
}

function ModeSwitch({ value, onChange }: { value: AnalysisMode; onChange: (value: AnalysisMode) => void }) {
  const modes: Array<{ value: AnalysisMode; label: string; icon: ReactNode }> = [
    { value: "geo", label: "Zona KML", icon: <LocateFixed className="h-4 w-4" /> },
    { value: "days", label: "Dias", icon: <FileSpreadsheet className="h-4 w-4" /> },
    { value: "combined", label: "Ambos", icon: <Layers3 className="h-4 w-4" /> }
  ];

  return (
    <div className="grid grid-cols-3 rounded-md border border-slate-200 bg-slate-50 p-1">
      {modes.map((mode) => (
        <button
          key={mode.value}
          type="button"
          onClick={() => onChange(mode.value)}
          className={clsx(
            "inline-flex h-10 items-center justify-center gap-2 rounded px-3 text-sm font-semibold transition",
            value === mode.value ? "bg-slate-950 text-white shadow-sm" : "text-slate-600 hover:bg-white"
          )}
        >
          {mode.icon}
          {mode.label}
        </button>
      ))}
    </div>
  );
}

function TabSwitch({ value, onChange }: { value: AppTab; onChange: (value: AppTab) => void }) {
  const tabs: Array<{ value: AppTab; label: string; icon: ReactNode }> = [
    { value: "map", label: "Mapa", icon: <MapIcon className="h-4 w-4" /> },
    { value: "analysis", label: "Analisis", icon: <BarChart3 className="h-4 w-4" /> }
  ];

  return (
    <div className="flex rounded-lg border border-slate-200 bg-white p-1 shadow-panel">
      {tabs.map((tab) => (
        <button
          key={tab.value}
          type="button"
          onClick={() => onChange(tab.value)}
          className={clsx(
            "inline-flex h-10 flex-1 items-center justify-center gap-2 rounded-md text-sm font-semibold transition",
            value === tab.value ? "bg-slate-950 text-white" : "text-slate-600 hover:bg-slate-50"
          )}
        >
          {tab.icon}
          {tab.label}
        </button>
      ))}
    </div>
  );
}

function FilterPanel({
  office,
  setOffice,
  offices,
  customerType,
  setCustomerType,
  customerTypes,
  zone,
  setZone,
  zonesForFilter,
  status,
  setStatus,
  sortKey,
  setSortKey,
  showZones,
  setShowZones,
  lockMapView,
  setLockMapView,
  query,
  setQuery
}: {
  office: string;
  setOffice: (value: string) => void;
  offices: string[];
  customerType: string;
  setCustomerType: (value: string) => void;
  customerTypes: string[];
  zone: string;
  setZone: (value: string) => void;
  zonesForFilter: Array<{ code: string; name: string }>;
  status: ReviewStatus | "all";
  setStatus: (value: ReviewStatus | "all") => void;
  sortKey: SortKey;
  setSortKey: (value: SortKey) => void;
  showZones: boolean;
  setShowZones: (value: (current: boolean) => boolean) => void;
  lockMapView: boolean;
  setLockMapView: (value: (current: boolean) => boolean) => void;
  query: string;
  setQuery: (value: string) => void;
}) {
  return (
    <div className="rounded-lg border border-slate-200 bg-white p-4 shadow-panel">
      <div className="mb-3 flex items-center gap-2 text-sm font-semibold">
        <Filter className="h-4 w-4" />
        Filtros
      </div>
      <div className="grid gap-3 sm:grid-cols-2 xl:grid-cols-1">
        <Select label="Oficina" value={office} onChange={setOffice} options={[{ label: "Todas", value: "all" }, ...offices.map((item) => ({ label: item, value: item }))]} />
        <Select
          label="Tipo cliente"
          value={customerType}
          onChange={setCustomerType}
          options={[{ label: "Todos", value: "all" }, ...customerTypes.map((item) => ({ label: item, value: item }))]}
        />
        <Select
          label="Zona"
          value={zone}
          onChange={setZone}
          options={[{ label: "Todas", value: "all" }, ...zonesForFilter.map((item) => ({ label: `${item.code} - ${item.name}`, value: item.code }))]}
        />
        <Select
          label="Estado"
          value={status}
          onChange={(value) => setStatus(value as ReviewStatus | "all")}
          options={[
            { label: "Todos", value: "all" },
            { label: "Alertas", value: "mismatch" },
            { label: "OK", value: "match" },
            { label: "Fuera de zona", value: "outsideZone" },
            { label: "Sin KML", value: "noKml" },
            { label: "Sin calendario", value: "noSchedule" },
            { label: "Sin ubicacion", value: "noLocation" }
          ]}
        />
        <Select
          label="Orden"
          value={sortKey}
          onChange={(value) => setSortKey(value as SortKey)}
          options={[
            { label: "Prioridad de alerta", value: "priority" },
            { label: "Kilos prom. mayor", value: "kilosDesc" },
            { label: "Kilos prom. menor", value: "kilosAsc" },
            { label: "Cajas prom. mayor", value: "boxesDesc" },
            { label: "Nombre cliente", value: "nameAsc" },
            { label: "Estado", value: "status" },
            { label: "Zona maestro", value: "zoneAsc" }
          ]}
        />
        <button
          type="button"
          onClick={() => setShowZones((current) => !current)}
          className={clsx(
            "inline-flex h-10 w-full items-center justify-center gap-2 rounded-md border px-3 text-sm font-semibold transition",
            showZones
              ? "border-teal-200 bg-teal-50 text-teal-800 hover:bg-teal-100"
              : "border-slate-200 bg-slate-50 text-slate-600 hover:bg-white"
          )}
        >
          {showZones ? <Eye className="h-4 w-4" /> : <EyeOff className="h-4 w-4" />}
          {showZones ? "Ocultar zonas KML" : "Mostrar zonas KML"}
        </button>
        <button
          type="button"
          onClick={() => setLockMapView((current) => !current)}
          className={clsx(
            "inline-flex h-10 w-full items-center justify-center gap-2 rounded-md border px-3 text-sm font-semibold transition",
            lockMapView
              ? "border-slate-300 bg-slate-950 text-white hover:bg-slate-800"
              : "border-slate-200 bg-slate-50 text-slate-600 hover:bg-white"
          )}
        >
          {lockMapView ? <Lock className="h-4 w-4" /> : <Unlock className="h-4 w-4" />}
          {lockMapView ? "Zoom bloqueado" : "Bloquear zoom"}
        </button>
        <label className="block sm:col-span-2 xl:col-span-1">
          <span className="mb-1 block text-xs font-semibold uppercase tracking-[0.12em] text-slate-500">Buscar</span>
          <div className="flex h-10 items-center gap-2 rounded-md border border-slate-200 bg-slate-50 px-3">
            <Search className="h-4 w-4 text-slate-400" />
            <input
              value={query}
              onChange={(event) => setQuery(event.target.value)}
              className="min-w-0 flex-1 bg-transparent text-sm outline-none"
              placeholder="Cliente, codigo, direccion..."
            />
          </div>
        </label>
      </div>
    </div>
  );
}

function AnalysisFilterBar({
  office,
  setOffice,
  offices,
  customerType,
  setCustomerType,
  customerTypes,
  status,
  setStatus,
  query,
  setQuery
}: {
  office: string;
  setOffice: (value: string) => void;
  offices: string[];
  customerType: string;
  setCustomerType: (value: string) => void;
  customerTypes: string[];
  status: ReviewStatus | "all";
  setStatus: (value: ReviewStatus | "all") => void;
  query: string;
  setQuery: (value: string) => void;
}) {
  return (
    <section className="rounded-lg border border-slate-200 bg-white p-4 shadow-panel">
      <div className="mb-3 flex items-center gap-2 text-sm font-semibold">
        <Filter className="h-4 w-4" />
        Filtros de analisis
      </div>
      <div className="grid gap-3 md:grid-cols-2 xl:grid-cols-[1fr_1fr_1fr_1.4fr]">
        <Select label="Oficina" value={office} onChange={setOffice} options={[{ label: "Todas", value: "all" }, ...offices.map((item) => ({ label: item, value: item }))]} />
        <Select
          label="Tipo cliente"
          value={customerType}
          onChange={setCustomerType}
          options={[{ label: "Todos", value: "all" }, ...customerTypes.map((item) => ({ label: item, value: item }))]}
        />
        <Select
          label="Estado"
          value={status}
          onChange={(value) => setStatus(value as ReviewStatus | "all")}
          options={[
            { label: "Todos", value: "all" },
            { label: "Alertas", value: "mismatch" },
            { label: "OK", value: "match" },
            { label: "Fuera de zona", value: "outsideZone" },
            { label: "Sin KML", value: "noKml" },
            { label: "Sin calendario", value: "noSchedule" },
            { label: "Sin ubicacion", value: "noLocation" }
          ]}
        />
        <label className="block">
          <span className="mb-1 block text-xs font-semibold uppercase tracking-[0.12em] text-slate-500">Buscar</span>
          <div className="flex h-10 items-center gap-2 rounded-md border border-slate-200 bg-slate-50 px-3">
            <Search className="h-4 w-4 text-slate-400" />
            <input
              value={query}
              onChange={(event) => setQuery(event.target.value)}
              className="min-w-0 flex-1 bg-transparent text-sm outline-none"
              placeholder="Cliente, codigo, direccion..."
            />
          </div>
        </label>
      </div>
    </section>
  );
}

function AnalysisPanel({ records, mode }: { records: ReviewRecord[]; mode: AnalysisMode }) {
  const analytics = useMemo(() => buildAnalytics(records), [records]);

  return (
    <section className="space-y-4">
      <div className="grid gap-3 md:grid-cols-2 xl:grid-cols-4">
        <InsightMetric
          label="Clientes en vista"
          value={records.length.toLocaleString("es-CL")}
          detail={`${modeCopy[mode]} activo`}
          icon={<Target className="h-4 w-4" />}
        />
        <InsightMetric
          label="Kilos en alerta"
          value={Math.round(analytics.alertKilos).toLocaleString("es-CL")}
          detail={`${analytics.alertKiloShare}% del volumen filtrado`}
          tone="danger"
          icon={<TrendingUp className="h-4 w-4" />}
        />
        <InsightMetric
          label="Coordenadas utiles"
          value={`${analytics.locationRate}%`}
          detail={`${analytics.located.toLocaleString("es-CL")} clientes con lat/lon`}
          tone="success"
          icon={<LocateFixed className="h-4 w-4" />}
        />
        <InsightMetric
          label="Desviacion geografica"
          value={analytics.geoAlerts.toLocaleString("es-CL")}
          detail="ZR distinta o fuera de zona KML"
          tone="warn"
          icon={<MapPinned className="h-4 w-4" />}
        />
      </div>

      <div className="grid gap-4 xl:grid-cols-[minmax(0,1.15fr)_minmax(360px,0.85fr)]">
        <Panel title="Composicion de estados" subtitle="Distribucion de la vista filtrada">
          <BarList items={analytics.statusBars} />
        </Panel>

        <Panel title="Calidad de coordenadas" subtitle="Cobertura geografica para auditoria KML">
          <div className="grid gap-3 sm:grid-cols-3">
            <QualityTile label="Con ubicacion" value={analytics.located} total={records.length} tone="success" />
            <QualityTile label="Sin ubicacion" value={analytics.noLocation} total={records.length} tone="neutral" />
            <QualityTile label="Fuera KML" value={analytics.outsideZone} total={records.length} tone="warn" />
          </div>
          <div className="mt-4">
            <BarList items={analytics.geoBars} compact />
          </div>
        </Panel>
      </div>

      <div className="grid gap-4 xl:grid-cols-3">
        <Panel title="Riesgo por tipo de cliente" subtitle="Alertas y volumen por segmento">
          <BarList items={analytics.topCustomerTypes} emptyText="No hay alertas por tipo de cliente en la vista actual." />
        </Panel>

        <Panel title="Zonas con mayor riesgo" subtitle="Alertas ponderadas por volumen de clientes">
          <BarList items={analytics.topZones} emptyText="No hay alertas por zona en la vista actual." />
        </Panel>

        <Panel title="Desviaciones de dias" subtitle="Dias faltantes y sobrantes contra calendario">
          <DayDeviationChart items={analytics.dayDeviation} />
        </Panel>
      </div>

      <div className="grid gap-4 xl:grid-cols-[minmax(0,1fr)_420px]">
        <Panel title="Clientes alertados de mayor volumen" subtitle="Prioridad comercial para correccion de maestro">
          <TopClients records={analytics.topAlertClients} />
        </Panel>

        <Panel title="Lectura operativa" subtitle="Senales accionables para limpieza">
          <RecommendationList analytics={analytics} />
        </Panel>
      </div>
    </section>
  );
}

function InsightMetric({
  label,
  value,
  detail,
  icon,
  tone = "default"
}: {
  label: string;
  value: string;
  detail: string;
  icon: ReactNode;
  tone?: "default" | "success" | "danger" | "warn";
}) {
  const tones = {
    default: "bg-slate-950 text-white",
    success: "bg-emerald-600 text-white",
    danger: "bg-red-600 text-white",
    warn: "bg-amber-500 text-white"
  };

  return (
    <div className="rounded-lg border border-slate-200 bg-white p-4 shadow-panel">
      <div className="flex items-start justify-between gap-3">
        <div>
          <p className="text-sm font-medium text-slate-500">{label}</p>
          <p className="mt-2 text-3xl font-semibold tracking-normal">{value}</p>
        </div>
        <span className={clsx("grid h-8 w-8 place-items-center rounded-md", tones[tone])}>{icon}</span>
      </div>
      <p className="mt-2 text-sm text-slate-500">{detail}</p>
    </div>
  );
}

function Panel({ title, subtitle, children }: { title: string; subtitle: string; children: ReactNode }) {
  return (
    <div className="rounded-lg border border-slate-200 bg-white p-4 shadow-panel">
      <div className="mb-4">
        <h2 className="text-base font-semibold">{title}</h2>
        <p className="text-sm text-slate-500">{subtitle}</p>
      </div>
      {children}
    </div>
  );
}

function BarList({ items, compact = false, emptyText = "Sin datos para mostrar." }: { items: BarItem[]; compact?: boolean; emptyText?: string }) {
  const max = Math.max(...items.map((item) => item.value), 0);

  if (!items.length || max === 0) {
    return <div className="grid min-h-[160px] place-items-center rounded-md bg-slate-50 px-4 text-center text-sm text-slate-500">{emptyText}</div>;
  }

  return (
    <div className={clsx("space-y-3", compact && "space-y-2")}>
      {items.map((item) => (
        <div key={item.label}>
          <div className="mb-1 flex items-center justify-between gap-3 text-sm">
            <span className="min-w-0 truncate font-medium text-slate-700">{item.label}</span>
            <span className="shrink-0 font-semibold text-slate-950">{item.value.toLocaleString("es-CL")}</span>
          </div>
          <div className="h-2.5 overflow-hidden rounded-full bg-slate-100">
            <div className={clsx("h-full rounded-full", item.color)} style={{ width: `${Math.max((item.value / max) * 100, 2)}%` }} />
          </div>
          {item.detail && <p className="mt-1 text-xs text-slate-500">{item.detail}</p>}
        </div>
      ))}
    </div>
  );
}

function QualityTile({ label, value, total, tone }: { label: string; value: number; total: number; tone: "success" | "warn" | "neutral" }) {
  const pctValue = total ? Math.round((value / total) * 100) : 0;
  const toneClass = {
    success: "text-emerald-700 bg-emerald-50",
    warn: "text-orange-700 bg-orange-50",
    neutral: "text-slate-700 bg-slate-100"
  }[tone];

  return (
    <div className={clsx("rounded-md px-3 py-3", toneClass)}>
      <p className="text-xs font-semibold uppercase tracking-[0.12em]">{label}</p>
      <p className="mt-2 text-2xl font-semibold">{pctValue}%</p>
      <p className="text-xs">{value.toLocaleString("es-CL")} clientes</p>
    </div>
  );
}

function DayDeviationChart({ items }: { items: DayDeviation[] }) {
  const max = Math.max(...items.flatMap((item) => [item.missing, item.extra]), 0);

  if (max === 0) {
    return <div className="grid min-h-[210px] place-items-center rounded-md bg-slate-50 px-4 text-center text-sm text-slate-500">No hay desviaciones de dias en la vista actual.</div>;
  }

  return (
    <div className="grid grid-cols-7 items-end gap-3 pt-2">
      {items.map((item) => (
        <div key={item.day} className="text-center">
          <div className="mx-auto flex h-36 w-full max-w-12 items-end justify-center gap-1">
            <div
              className="w-4 rounded-t bg-red-500"
              style={{ height: `${Math.max((item.missing / max) * 100, item.missing ? 5 : 0)}%` }}
              title={`Faltan ${item.missing}`}
            />
            <div
              className="w-4 rounded-t bg-amber-500"
              style={{ height: `${Math.max((item.extra / max) * 100, item.extra ? 5 : 0)}%` }}
              title={`Sobran ${item.extra}`}
            />
          </div>
          <p className="mt-2 text-xs font-semibold text-slate-700">{item.day}</p>
          <p className="text-[11px] text-slate-500">{item.missing}/{item.extra}</p>
        </div>
      ))}
      <div className="col-span-7 mt-3 flex justify-center gap-4 text-xs text-slate-500">
        <LegendItem color="#ef4444" label="Faltan" />
        <LegendItem color="#f59e0b" label="Sobran" />
      </div>
    </div>
  );
}

function TopClients({ records }: { records: ReviewRecord[] }) {
  if (!records.length) {
    return <div className="grid min-h-[180px] place-items-center rounded-md bg-slate-50 px-4 text-center text-sm text-slate-500">No hay clientes alertados en la vista actual.</div>;
  }

  return (
    <div className="divide-y divide-slate-100">
      {records.map((record) => (
        <div key={`${record.id}-${record.zoneCode}`} className="grid grid-cols-[minmax(0,1fr)_auto] gap-3 py-3">
          <div className="min-w-0">
            <p className="truncate text-sm font-semibold">{record.name}</p>
            <p className="truncate text-xs text-slate-500">{record.id} - {record.customerType || "Sin tipo"} - {record.zoneCode}</p>
            <p className="truncate text-xs text-slate-400">{record.population || record.address}</p>
          </div>
          <div className="text-right">
            <p className="text-sm font-semibold">{Math.round(record.yearlyKilos).toLocaleString("es-CL")}</p>
            <p className="text-xs text-slate-500">kg prom.</p>
          </div>
        </div>
      ))}
    </div>
  );
}

function RecommendationList({ analytics }: { analytics: Analytics }) {
  const recommendations = buildRecommendations(analytics);

  return (
    <div className="space-y-3">
      {recommendations.map((item) => (
        <div key={item.title} className="rounded-md border border-slate-200 bg-slate-50 px-3 py-3">
          <p className="text-sm font-semibold text-slate-800">{item.title}</p>
          <p className="mt-1 text-sm leading-5 text-slate-600">{item.detail}</p>
        </div>
      ))}
    </div>
  );
}

function UploadCard({
  title,
  icon,
  filename,
  accept,
  onFile
}: {
  title: string;
  icon: ReactNode;
  filename?: string;
  accept: string;
  onFile: (file: File) => void;
}) {
  return (
    <label className="group flex cursor-pointer items-center gap-4 rounded-lg border border-slate-200 bg-white p-4 shadow-panel transition hover:border-teal-400">
      <span className="grid h-11 w-11 shrink-0 place-items-center rounded-md bg-teal-50 text-teal-700">{icon}</span>
      <span className="min-w-0 flex-1">
        <span className="block text-sm font-semibold">{title}</span>
        <span className="block truncate text-sm text-slate-500">{filename || "Seleccionar archivo"}</span>
      </span>
      <UploadCloud className="h-5 w-5 shrink-0 text-slate-400 transition group-hover:text-teal-700" />
      <input
        className="sr-only"
        type="file"
        accept={accept}
        onChange={(event) => {
          const file = event.target.files?.[0];
          if (file) onFile(file);
        }}
      />
    </label>
  );
}

function Metric({
  label,
  value,
  icon,
  tone = "default"
}: {
  label: string;
  value: string;
  icon: ReactNode;
  tone?: "default" | "success" | "danger" | "warn";
}) {
  const tones = {
    default: "bg-slate-950 text-white",
    success: "bg-emerald-600 text-white",
    danger: "bg-red-600 text-white",
    warn: "bg-amber-500 text-white"
  };

  return (
    <div className="rounded-lg border border-slate-200 bg-white p-4 shadow-panel">
      <div className="flex items-center justify-between gap-3">
        <span className="text-sm font-medium text-slate-500">{label}</span>
        <span className={clsx("grid h-8 w-8 place-items-center rounded-md", tones[tone])}>{icon}</span>
      </div>
      <div className="mt-3 text-3xl font-semibold tracking-normal">{value}</div>
    </div>
  );
}

function Select({
  label,
  value,
  options,
  onChange
}: {
  label: string;
  value: string;
  options: Array<{ label: string; value: string }>;
  onChange: (value: string) => void;
}) {
  return (
    <label className="block">
      <span className="mb-1 block text-xs font-semibold uppercase tracking-[0.12em] text-slate-500">{label}</span>
      <select
        value={value}
        onChange={(event) => onChange(event.target.value)}
        className="h-10 w-full rounded-md border border-slate-200 bg-slate-50 px-3 text-sm outline-none transition focus:border-teal-500"
      >
        {options.map((option) => (
          <option key={`${label}-${option.value}`} value={option.value}>
            {option.label}
          </option>
        ))}
      </select>
    </label>
  );
}

function ReviewList({ records, mode, sortKey }: { records: ReviewRecord[]; mode: AnalysisMode; sortKey: SortKey }) {
  const priority = sortRecords(records, sortKey).slice(0, 120);

  return (
    <div className="flex min-h-[360px] flex-1 flex-col overflow-hidden rounded-lg border border-slate-200 bg-white shadow-panel">
      <div className="border-b border-slate-200 px-4 py-3">
        <div className="flex items-start justify-between gap-3">
          <div>
            <h2 className="text-base font-semibold">Clientes revisados</h2>
            <p className="text-sm text-slate-500">Primeros {priority.length.toLocaleString("es-CL")} registros - {modeCopy[mode]}</p>
          </div>
          <ArrowDownWideNarrow className="mt-1 h-4 w-4 shrink-0 text-slate-400" />
        </div>
      </div>
      <div className="min-h-0 flex-1 overflow-y-auto">
        {priority.length ? (
          <div className="divide-y divide-slate-100">
            {priority.map((record) => (
              <article key={`${record.id}-${record.zoneCode}`} className="p-4">
                <div className="mb-2 flex items-start justify-between gap-3">
                  <div className="min-w-0">
                    <h3 className="truncate text-sm font-semibold">{record.name}</h3>
                    <p className="truncate text-xs text-slate-500">{record.id} - {record.population || record.address}</p>
                  </div>
                  <StatusBadge status={record.status} />
                </div>
                <div className="grid grid-cols-2 gap-2 text-xs">
                  <Info label="ZR maestro" value={record.zoneCode || "Sin zona"} />
                  <Info label="ZR KML" value={record.kmlZoneCode || statusCopy[record.geoStatus]} />
                  <Info label="Tipo cliente" value={record.customerType || "Sin tipo"} />
                  <Info label="Dias maestro" value={formatDays(record.actualDays)} />
                  <Info label="Calendario zona" value={formatDays(record.expectedDays)} />
                  <Info label="Kilos prom." value={Math.round(record.yearlyKilos).toLocaleString("es-CL")} />
                </div>
                {record.issues.length > 0 && (
                  <p className="mt-2 rounded-md bg-red-50 px-3 py-2 text-xs text-red-700">
                    {record.issues.join(" | ")}
                  </p>
                )}
              </article>
            ))}
          </div>
        ) : (
          <div className="grid min-h-[280px] place-items-center px-6 text-center text-sm text-slate-500">
            Carga los archivos para comenzar la revision.
          </div>
        )}
      </div>
    </div>
  );
}

function StatusBadge({ status }: { status: ReviewStatus }) {
  return (
    <span
      className={clsx(
        "shrink-0 rounded-md px-2 py-1 text-xs font-semibold",
        status === "match" && "bg-emerald-50 text-emerald-700",
        status === "mismatch" && "bg-red-50 text-red-700",
        status === "noSchedule" && "bg-amber-50 text-amber-700",
        status === "outsideZone" && "bg-orange-50 text-orange-700",
        status === "noKml" && "bg-blue-50 text-blue-700",
        status === "noLocation" && "bg-slate-100 text-slate-600"
      )}
    >
      {statusCopy[status]}
    </span>
  );
}

function Info({ label, value }: { label: string; value: string }) {
  return (
    <div className="rounded-md bg-slate-50 px-3 py-2">
      <div className="font-medium text-slate-400">{label}</div>
      <div className="mt-1 font-semibold text-slate-800">{value}</div>
    </div>
  );
}

function Legend() {
  return (
    <div className="flex flex-wrap items-center gap-3 text-xs text-slate-500">
      <LegendItem color="#1f9d6a" label="OK" />
      <LegendItem color="#de4f3f" label="Alerta" />
      <LegendItem color="#f97316" label="Fuera de zona" />
      <LegendItem color="#d39122" label="Pendiente" />
    </div>
  );
}

function LegendItem({ color, label }: { color: string; label: string }) {
  return (
    <span className="inline-flex items-center gap-1.5">
      <span className="h-2.5 w-2.5 rounded-full" style={{ backgroundColor: color }} />
      {label}
    </span>
  );
}

function getSummary(records: ReviewRecord[]) {
  const total = records.length;
  const alerts = records.filter((record) => record.status === "mismatch" || record.status === "outsideZone").length;
  const match = records.filter((record) => record.status === "match").length;
  const pending = records.filter((record) => ["noSchedule", "noLocation", "noKml"].includes(record.status)).length;
  const comparable = match + alerts;
  const compliance = comparable ? Math.round((match / comparable) * 100) : 0;

  return { total, alerts, match, pending, compliance };
}

type BarItem = {
  label: string;
  value: number;
  detail?: string;
  color: string;
};

type DayDeviation = {
  day: string;
  missing: number;
  extra: number;
};

type Analytics = {
  located: number;
  noLocation: number;
  outsideZone: number;
  geoAlerts: number;
  alertKilos: number;
  alertKiloShare: number;
  locationRate: number;
  statusBars: BarItem[];
  geoBars: BarItem[];
  topCustomerTypes: BarItem[];
  topZones: BarItem[];
  dayDeviation: DayDeviation[];
  topAlertClients: ReviewRecord[];
};

function buildAnalytics(records: ReviewRecord[]): Analytics {
  const total = records.length;
  const located = records.filter((record) => typeof record.lat === "number" && typeof record.lng === "number").length;
  const noLocation = records.filter((record) => record.status === "noLocation" || record.geoStatus === "noLocation").length;
  const outsideZone = records.filter((record) => record.geoStatus === "outsideZone").length;
  const geoAlerts = records.filter((record) => record.geoStatus === "mismatch" || record.geoStatus === "outsideZone").length;
  const totalKilos = records.reduce((sum, record) => sum + record.yearlyKilos, 0);
  const alertRecords = records.filter((record) => record.status !== "match");
  const alertKilos = alertRecords.reduce((sum, record) => sum + record.yearlyKilos, 0);

  return {
    located,
    noLocation,
    outsideZone,
    geoAlerts,
    alertKilos,
    alertKiloShare: totalKilos ? Math.round((alertKilos / totalKilos) * 100) : 0,
    locationRate: total ? Math.round((located / total) * 100) : 0,
    statusBars: statusDistribution(records),
    geoBars: geoDistribution(records),
    topCustomerTypes: topAlertCustomerTypes(records),
    topZones: topAlertZones(records),
    dayDeviation: dayDeviation(records),
    topAlertClients: sortRecords(alertRecords, "kilosDesc").slice(0, 8)
  };
}

function statusDistribution(records: ReviewRecord[]): BarItem[] {
  const colors: Record<ReviewStatus, string> = {
    match: "bg-emerald-600",
    mismatch: "bg-red-600",
    outsideZone: "bg-orange-500",
    noSchedule: "bg-amber-500",
    noLocation: "bg-slate-500",
    noKml: "bg-blue-600"
  };
  const order: ReviewStatus[] = ["mismatch", "outsideZone", "noSchedule", "noLocation", "noKml", "match"];

  return order
    .map((status) => ({
      label: statusCopy[status],
      value: records.filter((record) => record.status === status).length,
      color: colors[status]
    }))
    .filter((item) => item.value > 0);
}

function geoDistribution(records: ReviewRecord[]): BarItem[] {
  return [
    {
      label: "ZR distinta",
      value: records.filter((record) => record.geoStatus === "mismatch").length,
      color: "bg-red-600"
    },
    {
      label: "Fuera de poligono",
      value: records.filter((record) => record.geoStatus === "outsideZone").length,
      color: "bg-orange-500"
    },
    {
      label: "Sin coordenadas",
      value: records.filter((record) => record.geoStatus === "noLocation").length,
      color: "bg-slate-500"
    },
    {
      label: "Zona correcta",
      value: records.filter((record) => record.geoStatus === "match").length,
      color: "bg-emerald-600"
    }
  ].filter((item) => item.value > 0);
}

function topAlertZones(records: ReviewRecord[]): BarItem[] {
  const zones = new Map<string, { count: number; kilos: number; name: string }>();
  for (const record of records) {
    if (record.status === "match") continue;
    const key = record.zoneCode || "Sin zona";
    const current = zones.get(key) ?? { count: 0, kilos: 0, name: record.zoneName || key };
    current.count += 1;
    current.kilos += record.yearlyKilos;
    zones.set(key, current);
  }

  return Array.from(zones.entries())
    .map(([code, data]) => ({
      label: `${code} - ${data.name}`,
      value: data.count,
      detail: `${Math.round(data.kilos).toLocaleString("es-CL")} kg prom. asociados`,
      color: "bg-red-600"
    }))
    .sort((a, b) => b.value - a.value)
    .slice(0, 8);
}

function topAlertCustomerTypes(records: ReviewRecord[]): BarItem[] {
  const types = new Map<string, { count: number; kilos: number; total: number }>();
  for (const record of records) {
    const key = record.customerType || "Sin tipo";
    const current = types.get(key) ?? { count: 0, kilos: 0, total: 0 };
    current.total += 1;
    if (record.status !== "match") {
      current.count += 1;
      current.kilos += record.yearlyKilos;
    }
    types.set(key, current);
  }

  return Array.from(types.entries())
    .map(([type, data]) => ({
      label: type,
      value: data.count,
      detail: `${data.total ? Math.round((data.count / data.total) * 100) : 0}% del tipo - ${Math.round(data.kilos).toLocaleString("es-CL")} kg prom.`,
      color: "bg-teal-700"
    }))
    .filter((item) => item.value > 0)
    .sort((a, b) => b.value - a.value)
    .slice(0, 8);
}

function dayDeviation(records: ReviewRecord[]): DayDeviation[] {
  const days = [
    { key: "L", label: "Lun" },
    { key: "M", label: "Mar" },
    { key: "W", label: "Mie" },
    { key: "J", label: "Jue" },
    { key: "V", label: "Vie" },
    { key: "S", label: "Sab" },
    { key: "D", label: "Dom" }
  ] as const;

  return days.map((day) => ({
    day: day.label,
    missing: records.filter((record) => record.missingDays.includes(day.key)).length,
    extra: records.filter((record) => record.extraDays.includes(day.key)).length
  }));
}

function buildRecommendations(analytics: Analytics) {
  const recommendations = [];

  if (analytics.geoAlerts > 0) {
    recommendations.push({
      title: "Priorizar correccion geografica",
      detail: `${analytics.geoAlerts.toLocaleString("es-CL")} clientes presentan desviacion de zona KML o estan fuera de poligono. Conviene revisar primero los de mayor volumen.`
    });
  }

  if (analytics.noLocation > 0) {
    recommendations.push({
      title: "Completar coordenadas",
      detail: `${analytics.noLocation.toLocaleString("es-CL")} clientes no pueden validarse geograficamente por latitud/longitud faltante o invalida.`
    });
  }

  if (analytics.alertKiloShare >= 20) {
    recommendations.push({
      title: "Impacto operacional alto",
      detail: `Las alertas concentran ${analytics.alertKiloShare}% del volumen filtrado. Esto puede afectar planificacion de rutas, dias de visita y carga operativa.`
    });
  }

  if (!recommendations.length) {
    recommendations.push({
      title: "Vista sin desviaciones criticas",
      detail: "No se observan alertas relevantes en el corte actual. Mantiene esta vista como referencia para contrastar otras oficinas o zonas."
    });
  }

  return recommendations;
}

function getCalendarSummary(schedules: ZoneSchedule[]) {
  return {
    total: schedules.length,
    active: schedules.filter((schedule) => schedule.expectedDays.length > 0).length
  };
}

function uniqueSorted(values: string[]) {
  return Array.from(new Set(values.filter(Boolean))).sort((a, b) => a.localeCompare(b, "es"));
}

function statusRank(status: ReviewStatus) {
  return { mismatch: 0, outsideZone: 1, noKml: 2, noSchedule: 3, noLocation: 4, match: 5 }[status];
}

function sortRecords(records: ReviewRecord[], sortKey: SortKey) {
  const sorted = [...records];
  const byPriority = (a: ReviewRecord, b: ReviewRecord) =>
    statusRank(a.status) - statusRank(b.status) ||
    b.yearlyKilos - a.yearlyKilos ||
    a.name.localeCompare(b.name, "es");

  return sorted.sort((a, b) => {
    if (sortKey === "priority") return byPriority(a, b);
    if (sortKey === "kilosDesc") return b.yearlyKilos - a.yearlyKilos || byPriority(a, b);
    if (sortKey === "kilosAsc") return a.yearlyKilos - b.yearlyKilos || byPriority(a, b);
    if (sortKey === "boxesDesc") return b.yearlyBoxes - a.yearlyBoxes || byPriority(a, b);
    if (sortKey === "nameAsc") return a.name.localeCompare(b.name, "es") || byPriority(a, b);
    if (sortKey === "status") return statusRank(a.status) - statusRank(b.status) || a.name.localeCompare(b.name, "es");
    return a.zoneCode.localeCompare(b.zoneCode, "es") || byPriority(a, b);
  });
}

function modeDescription(mode: AnalysisMode) {
  const copy: Record<AnalysisMode, string> = {
    geo: "Compara el codigo ZR del maestro contra el poligono KML donde cae el cliente.",
    days: "Compara dias del maestro contra el calendario esperado de la zona.",
    combined: "Levanta alerta si falla la zona KML o los dias de despacho."
  };
  return copy[mode];
}

function pendingLabel(mode: AnalysisMode) {
  if (mode === "geo") return "Sin validar";
  if (mode === "days") return "Sin calendario";
  return "Pendientes";
}

function referenceLabel(mode: AnalysisMode) {
  if (mode === "days") return "Zonas calendario";
  if (mode === "geo") return "Zonas KML";
  return "Referencias";
}

function referenceValue(mode: AnalysisMode, kmlZones: number, calendarSummary: { total: number; active: number }) {
  if (mode === "days") return `${calendarSummary.active}/${calendarSummary.total}`;
  if (mode === "geo") return kmlZones.toLocaleString("es-CL");
  return `${kmlZones.toLocaleString("es-CL")} KML`;
}

function getTemplateZones(clients: ClientRecord[], polygons: ZonePolygon[]) {
  const byCode = new Map<string, { zoneCode: string; zoneName: string }>();
  for (const client of clients) {
    if (client.zoneCode) byCode.set(client.zoneCode, { zoneCode: client.zoneCode, zoneName: client.zoneName || client.zoneCode });
  }
  for (const polygon of polygons) {
    if (polygon.code && !byCode.has(polygon.code)) byCode.set(polygon.code, { zoneCode: polygon.code, zoneName: polygon.name });
  }
  return Array.from(byCode.values()).sort((a, b) => a.zoneCode.localeCompare(b.zoneCode));
}
