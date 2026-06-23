"use client";

import dynamic from "next/dynamic";
import Link from "next/link";
import type { ReactNode } from "react";
import { useMemo, useState } from "react";
import { ArrowLeft, Calculator, Crosshair, Eye, EyeOff, FileSpreadsheet, Lock, MapPinned, Navigation, Play, Route, Server, Unlock, UploadCloud } from "lucide-react";
import {
  centroidFromClients,
  clientsInsideZone,
  parseClients,
  parseExcelRows,
  parseKml,
  type ClientRecord,
  type ZonePolygon
} from "@/lib/reviewer";
import type { RouteMetric } from "@/components/OsrmRouteMap";

const OsrmRouteMap = dynamic(() => import("@/components/OsrmRouteMap"), {
  ssr: false,
  loading: () => <div className="grid h-full min-h-[560px] place-items-center text-sm text-slate-500">Preparando mapa...</div>
});

type ZoneRouteInput = {
  zone: ZonePolygon;
  clients: ClientRecord[];
  centroid: [number, number] | null;
};

type ZoneRouteResult = RouteMetric & {
  zoneName: string;
  zoneCode: string;
  clientCount: number;
  failedRoutes: number;
};

type ProgressState = {
  done: number;
  total: number;
  label: string;
};

export default function RouteAnalysisApp() {
  const [clients, setClients] = useState<ClientRecord[]>([]);
  const [zones, setZones] = useState<ZonePolygon[]>([]);
  const [selectedZoneIds, setSelectedZoneIds] = useState<string[]>([]);
  const [originLat, setOriginLat] = useState("");
  const [originLng, setOriginLng] = useState("");
  const [pickingOrigin, setPickingOrigin] = useState(false);
  const [showZones, setShowZones] = useState(true);
  const [showClients, setShowClients] = useState(true);
  const [lockMapView, setLockMapView] = useState(false);
  const [osrmBaseUrl, setOsrmBaseUrl] = useState("https://router.project-osrm.org");
  const [results, setResults] = useState<ZoneRouteResult[]>([]);
  const [progress, setProgress] = useState<ProgressState>({ done: 0, total: 0, label: "" });
  const [running, setRunning] = useState(false);
  const [error, setError] = useState("");

  const selectedInputs = useMemo(() => {
    return zones
      .filter((zone) => selectedZoneIds.includes(zone.id))
      .map((zone) => {
        const zoneClients = clientsInsideZone(clients, zone);
        return {
          zone,
          clients: zoneClients,
          centroid: centroidFromClients(zoneClients)
        };
      });
  }, [clients, selectedZoneIds, zones]);

  const origin = useMemo(() => parseOrigin(originLat, originLng), [originLat, originLng]);
  const routeMetrics = useMemo<RouteMetric[]>(() => results.map(({ zoneName, zoneCode, clientCount, failedRoutes, ...metric }) => metric), [results]);
  const selectedClients = useMemo(() => {
    const byClient = new Map<string, ClientRecord>();
    for (const input of selectedInputs) {
      for (const client of input.clients) byClient.set(client.id, client);
    }
    return Array.from(byClient.values());
  }, [selectedInputs]);
  const zoneClientCounts = useMemo(() => {
    return Object.fromEntries(zones.map((zone) => [zone.id, clientsInsideZone(clients, zone).length]));
  }, [clients, zones]);
  const progressPct = progress.total ? Math.round((progress.done / progress.total) * 100) : 0;

  async function handleMaster(file: File) {
    setError("");
    const rows = await parseExcelRows(file);
    setClients(parseClients(rows));
    setResults([]);
  }

  async function handleKml(file: File) {
    setError("");
    const text = await file.text();
    const parsed = parseKml(text);
    setZones(parsed);
    setSelectedZoneIds([]);
    setResults([]);
  }

  async function runAnalysis() {
    if (!origin) {
      setError("Ingresa latitud y longitud validas para la sucursal.");
      return;
    }
    if (!selectedInputs.length) {
      setError("Selecciona al menos un poligono KML para analizar.");
      return;
    }

    setError("");
    setRunning(true);
    setResults([]);
    const totalRequests = selectedInputs.reduce((sum, item) => sum + item.clients.length + (item.centroid ? 1 : 0), 0);
    let done = 0;
    const nextResults: ZoneRouteResult[] = [];

    try {
      for (const input of selectedInputs) {
        const clientRoutes: Array<{ km: number; minutes: number }> = [];
        let failedRoutes = 0;

        for (const client of input.clients) {
          if (typeof client.lat !== "number" || typeof client.lng !== "number") continue;
          setProgress({ done, total: totalRequests, label: `${input.zone.name} - ${client.name}` });
          const route = await fetchRoute(osrmBaseUrl, origin, [client.lat, client.lng]);
          done += 1;
          if (route) clientRoutes.push(route);
          else failedRoutes += 1;
          setProgress({ done, total: totalRequests, label: `${input.zone.name} - clientes ${done}/${totalRequests}` });
        }

        let centroidRoute: { km: number; minutes: number; path: Array<[number, number]> } | null = null;
        if (input.centroid) {
          setProgress({ done, total: totalRequests, label: `${input.zone.name} - centroide` });
          centroidRoute = await fetchRoute(osrmBaseUrl, origin, input.centroid, true);
          done += 1;
          setProgress({ done, total: totalRequests, label: `${input.zone.name} - centroide` });
        }

        nextResults.push({
          zoneId: input.zone.id,
          zoneName: input.zone.name,
          zoneCode: input.zone.code,
          clientCount: input.clients.length,
          centroid: input.centroid,
          avgKm: average(clientRoutes.map((route) => route.km)),
          avgMinutes: average(clientRoutes.map((route) => route.minutes)),
          centroidKm: centroidRoute?.km ?? null,
          centroidMinutes: centroidRoute?.minutes ?? null,
          centroidPath: centroidRoute?.path ?? [],
          failedRoutes
        });
        setResults([...nextResults]);
      }
    } catch {
      setError("No pude completar el analisis OSRM. Revisa el endpoint, CORS o conectividad del servidor.");
    } finally {
      setRunning(false);
    }
  }

  return (
    <main className="min-h-screen bg-slate-50 text-slate-950">
      <div className="mx-auto flex w-full max-w-[1560px] flex-col gap-5 px-4 py-4 sm:px-6 lg:px-8">
        <header className="flex flex-col gap-4 rounded-lg border border-slate-200 bg-white px-5 py-4 shadow-panel lg:flex-row lg:items-center lg:justify-between">
          <div>
            <Link href="/" className="inline-flex items-center gap-2 text-sm font-semibold text-teal-700">
              <ArrowLeft className="h-4 w-4" />
              Revisor maestro
            </Link>
            <h1 className="mt-2 text-2xl font-semibold tracking-normal text-slate-950 sm:text-3xl">Analisis OSRM por poligono</h1>
            <p className="mt-2 max-w-3xl text-sm leading-6 text-slate-600">
              Calcula minutos y kilometros desde una sucursal hacia clientes y centroides por zona KML seleccionada.
            </p>
          </div>
        </header>

        <section className="grid gap-3 xl:grid-cols-[1fr_1fr_1.2fr]">
          <UploadCard title="Maestro clientes" icon={<FileSpreadsheet className="h-5 w-5" />} accept=".xlsx,.xls" onFile={handleMaster} value={`${clients.length.toLocaleString("es-CL")} clientes`} />
          <UploadCard title="Zonas KML" icon={<MapPinned className="h-5 w-5" />} accept=".kml" onFile={handleKml} value={`${zones.length.toLocaleString("es-CL")} poligonos`} />
          <div className="rounded-lg border border-slate-200 bg-white p-4 shadow-panel">
            <div className="mb-3 flex items-center gap-2 text-sm font-semibold">
              <Navigation className="h-4 w-4" />
              Sucursal / origen
            </div>
            <div className="grid gap-3 sm:grid-cols-2">
              <TextInput label="Latitud" value={originLat} onChange={setOriginLat} placeholder="-29.9533" />
              <TextInput label="Longitud" value={originLng} onChange={setOriginLng} placeholder="-71.3436" />
            </div>
            <button
              type="button"
              onClick={() => setPickingOrigin((current) => !current)}
              className={
                pickingOrigin
                  ? "mt-3 inline-flex h-10 w-full items-center justify-center gap-2 rounded-md bg-slate-950 px-3 text-sm font-semibold text-white transition hover:bg-slate-800"
                  : "mt-3 inline-flex h-10 w-full items-center justify-center gap-2 rounded-md border border-slate-200 bg-slate-50 px-3 text-sm font-semibold text-slate-700 transition hover:bg-white"
              }
            >
              <Crosshair className="h-4 w-4" />
              {pickingOrigin ? "Click en el mapa..." : "Elegir origen en mapa"}
            </button>
          </div>
        </section>

        <section className="grid gap-4 xl:grid-cols-[380px_minmax(0,1fr)]">
          <aside className="space-y-4">
            <div className="rounded-lg border border-slate-200 bg-white p-4 shadow-panel">
              <div className="mb-3 flex items-center gap-2 text-sm font-semibold">
                <Server className="h-4 w-4" />
                OSRM
              </div>
              <TextInput label="Endpoint" value={osrmBaseUrl} onChange={setOsrmBaseUrl} placeholder="http://localhost:5000" />
              <p className="mt-2 text-xs leading-5 text-slate-500">
                Usa tu servidor local o el publico de OSRM. La llamada se ejecuta desde el navegador.
              </p>
              <button
                onClick={runAnalysis}
                disabled={running || !clients.length || !zones.length}
                className="mt-4 inline-flex h-11 w-full items-center justify-center gap-2 rounded-md bg-slate-950 px-4 text-sm font-semibold text-white transition hover:bg-slate-800 disabled:cursor-not-allowed disabled:bg-slate-300"
              >
                <Play className="h-4 w-4" />
                {running ? "Calculando..." : "Calcular OSRM"}
              </button>
            </div>

            <div className="rounded-lg border border-slate-200 bg-white p-4 shadow-panel">
              <div className="mb-3 flex items-center gap-2 text-sm font-semibold">
                <MapPinned className="h-4 w-4" />
                Vista mapa
              </div>
              <div className="grid gap-2">
                <MapToggle
                  active={showZones}
                  onClick={() => setShowZones((current) => !current)}
                  activeLabel="Ocultar KML"
                  inactiveLabel="Mostrar KML"
                  icon={showZones ? <Eye className="h-4 w-4" /> : <EyeOff className="h-4 w-4" />}
                />
                <MapToggle
                  active={showClients}
                  onClick={() => setShowClients((current) => !current)}
                  activeLabel="Ocultar clientes"
                  inactiveLabel="Mostrar clientes"
                  icon={showClients ? <Eye className="h-4 w-4" /> : <EyeOff className="h-4 w-4" />}
                />
                <MapToggle
                  active={lockMapView}
                  onClick={() => setLockMapView((current) => !current)}
                  activeLabel="Zoom bloqueado"
                  inactiveLabel="Bloquear zoom"
                  icon={lockMapView ? <Lock className="h-4 w-4" /> : <Unlock className="h-4 w-4" />}
                  darkActive
                />
              </div>
            </div>

            <div className="rounded-lg border border-slate-200 bg-white p-4 shadow-panel">
              <div className="mb-3 flex items-center justify-between gap-3">
                <div className="flex items-center gap-2 text-sm font-semibold">
                  <Route className="h-4 w-4" />
                  Poligonos
                </div>
                <div className="flex gap-2">
                  <button className="text-xs font-semibold text-teal-700" onClick={() => setSelectedZoneIds(zones.map((zone) => zone.id))}>
                    Todos
                  </button>
                  <button className="text-xs font-semibold text-slate-500" onClick={() => setSelectedZoneIds([])}>
                    Limpiar
                  </button>
                </div>
              </div>
              <div className="max-h-[360px] space-y-2 overflow-y-auto pr-1">
                {zones.map((zone) => {
                  const checked = selectedZoneIds.includes(zone.id);
                  const count = clientsInsideZone(clients, zone).length;
                  return (
                    <label key={zone.id} className="flex cursor-pointer items-start gap-3 rounded-md border border-slate-200 bg-slate-50 px-3 py-2 text-sm">
                      <input
                        type="checkbox"
                        checked={checked}
                        onChange={() =>
                          setSelectedZoneIds((current) =>
                            current.includes(zone.id) ? current.filter((item) => item !== zone.id) : [...current, zone.id]
                          )
                        }
                        className="mt-1"
                      />
                      <span className="min-w-0">
                        <span className="block truncate font-semibold">{zone.name}</span>
                        <span className="text-xs text-slate-500">{count.toLocaleString("es-CL")} clientes dentro</span>
                      </span>
                    </label>
                  );
                })}
                {!zones.length && <div className="rounded-md bg-slate-50 px-3 py-8 text-center text-sm text-slate-500">Carga un KML para seleccionar poligonos.</div>}
              </div>
            </div>

            <ProgressPanel progress={progress} running={running} progressPct={progressPct} error={error} />
          </aside>

          <div className="overflow-hidden rounded-lg border border-slate-200 bg-white shadow-panel">
            <div className="flex min-h-[73px] items-center justify-between gap-3 border-b border-slate-200 px-4 py-3">
              <div>
                <h2 className="text-base font-semibold">Mapa de accesibilidad</h2>
                <p className="text-sm text-slate-500">{selectedInputs.length.toLocaleString("es-CL")} poligonos seleccionados - click en KML para seleccionar</p>
              </div>
              <div className="text-right text-sm">
                <p className="font-semibold">{selectedClients.length.toLocaleString("es-CL")} clientes visibles</p>
                <p className="text-slate-500">{results.length.toLocaleString("es-CL")} zonas calculadas</p>
              </div>
            </div>
            <div className="h-[500px] overflow-hidden xl:h-[560px]">
              <OsrmRouteMap
                clients={selectedClients}
                zones={zones}
                selectedZoneIds={selectedZoneIds}
                origin={origin}
                metrics={routeMetrics}
                zoneClientCounts={zoneClientCounts}
                onToggleZone={toggleZone}
                pickingOrigin={pickingOrigin}
                onPickOrigin={(point) => {
                  setOriginLat(point[0].toFixed(6));
                  setOriginLng(point[1].toFixed(6));
                  setPickingOrigin(false);
                }}
                showZones={showZones}
                showClients={showClients}
                fitEnabled={!lockMapView}
              />
            </div>
          </div>
        </section>

        <ResultsTable results={results} />
      </div>
    </main>
  );

  function toggleZone(zoneId: string) {
    setSelectedZoneIds((current) => (current.includes(zoneId) ? current.filter((item) => item !== zoneId) : [...current, zoneId]));
  }
}

function MapToggle({
  active,
  onClick,
  activeLabel,
  inactiveLabel,
  icon,
  darkActive = false
}: {
  active: boolean;
  onClick: () => void;
  activeLabel: string;
  inactiveLabel: string;
  icon: ReactNode;
  darkActive?: boolean;
}) {
  const activeClass = darkActive
    ? "border-slate-300 bg-slate-950 text-white hover:bg-slate-800"
    : "border-teal-200 bg-teal-50 text-teal-800 hover:bg-teal-100";

  return (
    <button
      type="button"
      onClick={onClick}
      className={`inline-flex h-10 w-full items-center justify-center gap-2 rounded-md border px-3 text-sm font-semibold transition ${
        active ? activeClass : "border-slate-200 bg-slate-50 text-slate-600 hover:bg-white"
      }`}
    >
      {icon}
      {active ? activeLabel : inactiveLabel}
    </button>
  );
}

function UploadCard({
  title,
  icon,
  accept,
  value,
  onFile
}: {
  title: string;
  icon: ReactNode;
  accept: string;
  value: string;
  onFile: (file: File) => void;
}) {
  return (
    <label className="group flex cursor-pointer items-center gap-4 rounded-lg border border-slate-200 bg-white p-4 shadow-panel transition hover:border-teal-400">
      <span className="grid h-11 w-11 shrink-0 place-items-center rounded-md bg-teal-50 text-teal-700">{icon}</span>
      <span className="min-w-0 flex-1">
        <span className="block text-sm font-semibold">{title}</span>
        <span className="block truncate text-sm text-slate-500">{value}</span>
      </span>
      <UploadCloud className="h-5 w-5 shrink-0 text-slate-400 transition group-hover:text-teal-700" />
      <input className="sr-only" type="file" accept={accept} onChange={(event) => event.target.files?.[0] && onFile(event.target.files[0])} />
    </label>
  );
}

function TextInput({ label, value, onChange, placeholder }: { label: string; value: string; onChange: (value: string) => void; placeholder: string }) {
  return (
    <label className="block">
      <span className="mb-1 block text-xs font-semibold uppercase tracking-[0.12em] text-slate-500">{label}</span>
      <input
        value={value}
        onChange={(event) => onChange(event.target.value)}
        placeholder={placeholder}
        className="h-10 w-full rounded-md border border-slate-200 bg-slate-50 px-3 text-sm outline-none transition focus:border-teal-500"
      />
    </label>
  );
}

function ProgressPanel({ progress, running, progressPct, error }: { progress: ProgressState; running: boolean; progressPct: number; error: string }) {
  return (
    <div className="rounded-lg border border-slate-200 bg-white p-4 shadow-panel">
      <div className="mb-3 flex items-center gap-2 text-sm font-semibold">
        <Calculator className="h-4 w-4" />
        Progreso
      </div>
      <div className="h-3 overflow-hidden rounded-full bg-slate-100">
        <div className="h-full rounded-full bg-teal-700 transition-all" style={{ width: `${progressPct}%` }} />
      </div>
      <div className="mt-2 flex justify-between text-xs text-slate-500">
        <span>{running ? progress.label || "Calculando rutas..." : "En espera"}</span>
        <span>{progressPct}%</span>
      </div>
      {error && <p className="mt-3 rounded-md bg-red-50 px-3 py-2 text-sm text-red-700">{error}</p>}
    </div>
  );
}

function ResultsTable({ results }: { results: ZoneRouteResult[] }) {
  if (!results.length) return null;
  return (
    <section className="rounded-lg border border-slate-200 bg-white shadow-panel">
      <div className="border-b border-slate-200 px-4 py-3">
        <h2 className="text-base font-semibold">Resultado por poligono</h2>
        <p className="text-sm text-slate-500">Promedio a clientes y ruta al centroide de clientes dentro del KML.</p>
      </div>
      <div className="overflow-x-auto">
        <table className="min-w-full divide-y divide-slate-200 text-sm">
          <thead className="bg-slate-50 text-left text-xs uppercase tracking-[0.08em] text-slate-500">
            <tr>
              <th className="px-4 py-3">Poligono</th>
              <th className="px-4 py-3">Clientes</th>
              <th className="px-4 py-3">Km prom.</th>
              <th className="px-4 py-3">Min prom.</th>
              <th className="px-4 py-3">Km centroide</th>
              <th className="px-4 py-3">Min centroide</th>
              <th className="px-4 py-3">Fallidas</th>
            </tr>
          </thead>
          <tbody className="divide-y divide-slate-100">
            {results.map((result) => (
              <tr key={result.zoneId}>
                <td className="px-4 py-3 font-semibold">{result.zoneName}</td>
                <td className="px-4 py-3">{result.clientCount.toLocaleString("es-CL")}</td>
                <td className="px-4 py-3">{formatMetric(result.avgKm, "km")}</td>
                <td className="px-4 py-3">{formatMetric(result.avgMinutes, "min")}</td>
                <td className="px-4 py-3">{formatMetric(result.centroidKm, "km")}</td>
                <td className="px-4 py-3">{formatMetric(result.centroidMinutes, "min")}</td>
                <td className="px-4 py-3">{result.failedRoutes.toLocaleString("es-CL")}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </section>
  );
}

async function fetchRoute(baseUrl: string, origin: [number, number], destination: [number, number], includeGeometry = false) {
  const base = baseUrl.replace(/\/$/, "");
  const coordinates = `${origin[1]},${origin[0]};${destination[1]},${destination[0]}`;
  const overview = includeGeometry ? "full&geometries=geojson" : "false";
  const response = await fetch(`${base}/route/v1/driving/${coordinates}?overview=${overview}&alternatives=false&steps=false`);
  if (!response.ok) return null;
  const data = await response.json();
  const route = data?.routes?.[0];
  if (!route) return null;
  return {
    km: route.distance / 1000,
    minutes: route.duration / 60,
    path: Array.isArray(route.geometry?.coordinates)
      ? route.geometry.coordinates.map(([lng, lat]: [number, number]) => [lat, lng] as [number, number])
      : []
  };
}

function parseOrigin(lat: string, lng: string): [number, number] | null {
  const parsedLat = Number(lat.replace(",", "."));
  const parsedLng = Number(lng.replace(",", "."));
  if (!Number.isFinite(parsedLat) || !Number.isFinite(parsedLng)) return null;
  return [parsedLat, parsedLng];
}

function average(values: number[]) {
  if (!values.length) return null;
  return values.reduce((sum, value) => sum + value, 0) / values.length;
}

function formatMetric(value: number | null, suffix: string) {
  return value === null ? "Sin dato" : `${value.toLocaleString("es-CL", { maximumFractionDigits: 1 })} ${suffix}`;
}
