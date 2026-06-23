"use client";

import { useEffect, useMemo } from "react";
import { CircleMarker, MapContainer, Marker, Pane, Polygon, Polyline, Popup, TileLayer, Tooltip, useMap, useMapEvents } from "react-leaflet";
import type { DivIcon, LatLngBoundsExpression, LatLngExpression } from "leaflet";
import L from "leaflet";
import type { ClientRecord, ZonePolygon } from "@/lib/reviewer";

export type RouteMetric = {
  zoneId: string;
  centroid: [number, number] | null;
  avgMinutes: number | null;
  avgKm: number | null;
  centroidMinutes: number | null;
  centroidKm: number | null;
  centroidPath: Array<[number, number]>;
};

type OsrmRouteMapProps = {
  clients: ClientRecord[];
  zones: ZonePolygon[];
  selectedZoneIds: string[];
  origin: [number, number] | null;
  metrics: RouteMetric[];
  zoneClientCounts: Record<string, number>;
  onToggleZone: (zoneId: string) => void;
  pickingOrigin: boolean;
  onPickOrigin: (point: [number, number]) => void;
  showZones: boolean;
  showClients: boolean;
  fitEnabled: boolean;
};

export default function OsrmRouteMap({
  clients,
  zones,
  selectedZoneIds,
  origin,
  metrics,
  zoneClientCounts,
  onToggleZone,
  pickingOrigin,
  onPickOrigin,
  showZones,
  showClients,
  fitEnabled
}: OsrmRouteMapProps) {
  const selectedZoneSet = useMemo(() => new Set(selectedZoneIds), [selectedZoneIds]);
  const locatedClients = useMemo(
    () => clients.filter((client) => typeof client.lat === "number" && typeof client.lng === "number"),
    [clients]
  );
  const metricsByZone = useMemo(() => new Map(metrics.map((metric) => [metric.zoneId, metric])), [metrics]);
  const bounds = useMemo(() => getBounds(locatedClients, zones, origin, metrics), [locatedClients, metrics, origin, zones]);
  const mapProps = bounds ? { bounds } : { center: [-29.95, -71.33] as LatLngExpression, zoom: 10 };

  return (
    <MapContainer {...mapProps} scrollWheelZoom className={pickingOrigin ? "h-full w-full cursor-crosshair" : "h-full w-full"}>
      <TileLayer
        attribution='&copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a>'
        url="https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png"
      />
      <FitBounds bounds={bounds} enabled={fitEnabled} />
      <PickOrigin enabled={pickingOrigin} onPick={onPickOrigin} />
      <Pane name="osrm-zones" style={{ zIndex: 390 }}>
        {showZones && zones.map((zone, index) => {
          const selected = selectedZoneSet.has(zone.id);
          const metric = metricsByZone.get(zone.id);
          return (
            <Polygon
              key={zone.id}
              pane="osrm-zones"
              positions={zone.rings as LatLngExpression[][]}
              eventHandlers={{ click: () => onToggleZone(zone.id) }}
              bubblingMouseEvents={false}
              pathOptions={{
                color: selected ? "#dc2626" : zoneColor(index),
                fillColor: selected ? "#dc2626" : zoneColor(index),
                fillOpacity: selected ? 0.18 : 0.07,
                opacity: selected ? 0.9 : 0.42,
                weight: selected ? 3 : 1.5
              }}
            >
              <Tooltip sticky>{zone.name}</Tooltip>
              <Popup pane="popupPane">
                <div className="min-w-56 space-y-1 text-sm">
                  <strong>{zone.name}</strong>
                  {zone.code && <div>Codigo: {zone.code}</div>}
                  <div>Clientes dentro: {(zoneClientCounts[zone.id] ?? 0).toLocaleString("es-CL")}</div>
                  <div>Km prom.: {formatMetric(metric?.avgKm ?? null, "km")}</div>
                  <div>Min prom.: {formatMetric(metric?.avgMinutes ?? null, "min")}</div>
                  <div>Km centroide: {formatMetric(metric?.centroidKm ?? null, "km")}</div>
                  <div>Min centroide: {formatMetric(metric?.centroidMinutes ?? null, "min")}</div>
                  <button className="mt-2 text-xs font-semibold text-teal-700" type="button">
                    {selected ? "Seleccionado" : "Click para seleccionar"}
                  </button>
                </div>
              </Popup>
            </Polygon>
          );
        })}
      </Pane>
      <Pane name="osrm-routes" style={{ zIndex: 560 }}>
        {metrics.map((metric, index) =>
          metric.centroidPath.length ? (
            <Polyline
              key={`${metric.zoneId}-route`}
              pane="osrm-routes"
              positions={metric.centroidPath}
              bubblingMouseEvents={false}
              pathOptions={{ color: zoneColor(index), opacity: 0.86, weight: 4 }}
            />
          ) : null
        )}
      </Pane>
      <Pane name="osrm-clients" style={{ zIndex: 620 }}>
        {showClients && locatedClients.map((client) => (
          <CircleMarker
            key={`${client.id}-${client.lat}-${client.lng}`}
            pane="osrm-clients"
            center={[client.lat as number, client.lng as number]}
            radius={3.8}
            bubblingMouseEvents={false}
            pathOptions={{
              color: "#ffffff",
              fillColor: "#0f766e",
              fillOpacity: 0.82,
              opacity: 1,
              weight: 1.2
            }}
          >
            <Popup pane="popupPane">
              <div className="min-w-52 space-y-1 text-sm">
                <strong>{client.name}</strong>
                <div>{client.address}</div>
                <div>Zona maestro: {client.zoneCode}</div>
                <div>Tipo: {client.customerType || "Sin tipo"}</div>
              </div>
            </Popup>
          </CircleMarker>
        ))}
        {origin && (
          <Marker position={origin} icon={originIcon()}>
            <Popup pane="popupPane">Punto de inicio</Popup>
          </Marker>
        )}
        {metrics.map((metric) =>
          metric.centroid ? (
            <Marker key={metric.zoneId} position={metric.centroid} icon={centroidIcon()}>
              <Popup pane="popupPane">
                <div className="space-y-1 text-sm">
                  <strong>Centroide clientes</strong>
                  <div>Km centroide: {formatMetric(metric.centroidKm, "km")}</div>
                  <div>Min centroide: {formatMetric(metric.centroidMinutes, "min")}</div>
                </div>
              </Popup>
            </Marker>
          ) : null
        )}
      </Pane>
    </MapContainer>
  );
}

function PickOrigin({ enabled, onPick }: { enabled: boolean; onPick: (point: [number, number]) => void }) {
  useMapEvents({
    click(event) {
      if (enabled) onPick([event.latlng.lat, event.latlng.lng]);
    }
  });

  return null;
}

function FitBounds({ bounds, enabled }: { bounds: LatLngBoundsExpression | null; enabled: boolean }) {
  const map = useMap();

  useEffect(() => {
    if (enabled && bounds) map.fitBounds(bounds, { padding: [28, 28], maxZoom: 12 });
  }, [bounds, enabled, map]);

  return null;
}

function getBounds(
  clients: ClientRecord[],
  zones: ZonePolygon[],
  origin: [number, number] | null,
  metrics: RouteMetric[]
): LatLngBoundsExpression | null {
  const points: Array<[number, number]> = [];
  if (origin) points.push(origin);
  for (const client of clients) points.push([client.lat as number, client.lng as number]);
  for (const zone of zones) {
    for (const ring of zone.rings) points.push(...ring);
  }
  for (const metric of metrics) {
    if (metric.centroid) points.push(metric.centroid);
  }
  if (!points.length) return null;
  const lats = points.map(([lat]) => lat);
  const lngs = points.map(([, lng]) => lng);
  return [
    [Math.min(...lats), Math.min(...lngs)],
    [Math.max(...lats), Math.max(...lngs)]
  ];
}

function originIcon(): DivIcon {
  return L.divIcon({
    className: "osrm-marker osrm-marker-origin",
    html: "O",
    iconSize: [28, 28],
    iconAnchor: [14, 14]
  });
}

function centroidIcon(): DivIcon {
  return L.divIcon({
    className: "osrm-marker osrm-marker-centroid",
    html: "C",
    iconSize: [24, 24],
    iconAnchor: [12, 12]
  });
}

function formatMetric(value: number | null, suffix: string) {
  return value === null ? "Sin dato" : `${value.toLocaleString("es-CL", { maximumFractionDigits: 1 })} ${suffix}`;
}

function zoneColor(index: number) {
  const palette = ["#2563eb", "#12a594", "#8b5cf6", "#ef7b45", "#0f766e", "#be4bdb", "#3b82f6", "#84a52a"];
  return palette[index % palette.length];
}
