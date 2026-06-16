"use client";

import { useEffect, useMemo } from "react";
import { CircleMarker, MapContainer, Pane, Polygon, Popup, TileLayer, Tooltip, useMap } from "react-leaflet";
import type { LatLngBoundsExpression, LatLngExpression } from "leaflet";
import { formatDays, type ReviewRecord, type ZonePolygon } from "@/lib/reviewer";

type ZoneMapProps = {
  clients: ReviewRecord[];
  zones: ZonePolygon[];
  showZones: boolean;
  fitEnabled: boolean;
};

const statusColors = {
  match: "#1f9d6a",
  mismatch: "#de4f3f",
  noSchedule: "#d39122",
  noLocation: "#64748b",
  outsideZone: "#f97316",
  noKml: "#2563eb"
};

export default function ZoneMap({ clients, zones, showZones, fitEnabled }: ZoneMapProps) {
  const locatedClients = useMemo(
    () => clients.filter((client) => typeof client.lat === "number" && typeof client.lng === "number"),
    [clients]
  );
  const bounds = useMemo(() => getBounds(locatedClients, zones), [locatedClients, zones]);
  const mapProps = bounds
    ? { bounds }
    : { center: [-29.95, -71.33] as LatLngExpression, zoom: 11 };

  return (
    <MapContainer
      {...mapProps}
      scrollWheelZoom
      className="h-full min-h-[460px] w-full"
    >
      <TileLayer
        attribution='&copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a>'
        url="https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png"
      />
      <FitBounds bounds={bounds} enabled={fitEnabled} />
      <Pane name="zones" style={{ zIndex: 390 }}>
        {showZones &&
          zones.map((zone, index) => (
            <Polygon
              key={zone.id}
              pane="zones"
              positions={zone.rings as LatLngExpression[][]}
              bubblingMouseEvents={false}
              pathOptions={{
                color: zoneColor(index),
                fillColor: zoneColor(index),
                fillOpacity: 0.09,
                opacity: 0.5,
                weight: 2
              }}
            >
              <Tooltip sticky>{zone.name}</Tooltip>
              <Popup pane="popupPane">
                <div className="space-y-1">
                  <strong>{zone.name}</strong>
                  {zone.code && <div>Codigo: {zone.code}</div>}
                </div>
              </Popup>
            </Polygon>
          ))}
      </Pane>
      <Pane name="clients" style={{ zIndex: 620 }}>
        {locatedClients.map((client) => (
          <CircleMarker
            key={`${client.id}-${client.zoneCode}`}
            pane="clients"
            center={[client.lat as number, client.lng as number]}
            radius={client.status === "mismatch" || client.status === "outsideZone" ? 5.5 : 4.2}
            bubblingMouseEvents={false}
            pathOptions={{
              color: "#ffffff",
              fillColor: statusColors[client.status],
              fillOpacity: 0.92,
              opacity: 1,
              weight: 1.5
            }}
          >
            <Popup pane="popupPane">
              <div className="min-w-52 space-y-1 text-sm">
                <strong>{client.name}</strong>
                <div>{client.address}</div>
                <div>Tipo cliente: {client.customerType || "Sin tipo"}</div>
                <div>ZR maestro: {client.zoneCode} - {client.zoneName}</div>
                <div>ZR KML: {client.kmlZoneCode || "Sin zona"} {client.kmlZoneName ? `- ${client.kmlZoneName}` : ""}</div>
                <div>Dias maestro: {formatDays(client.actualDays)}</div>
                <div>Calendario zona: {formatDays(client.expectedDays)}</div>
                {client.issues.length > 0 && <div>Alertas: {client.issues.join(" | ")}</div>}
              </div>
            </Popup>
          </CircleMarker>
        ))}
      </Pane>
    </MapContainer>
  );
}

function FitBounds({ bounds, enabled }: { bounds: LatLngBoundsExpression | null; enabled: boolean }) {
  const map = useMap();

  useEffect(() => {
    if (enabled && bounds) map.fitBounds(bounds, { padding: [28, 28], maxZoom: 13 });
  }, [bounds, enabled, map]);

  return null;
}

function getBounds(clients: ReviewRecord[], zones: ZonePolygon[]): LatLngBoundsExpression | null {
  const points: Array<[number, number]> = [];
  for (const client of clients) {
    if (typeof client.lat === "number" && typeof client.lng === "number") points.push([client.lat, client.lng]);
  }
  for (const zone of zones) {
    for (const ring of zone.rings) points.push(...ring);
  }

  if (!points.length) return null;
  const lats = points.map(([lat]) => lat);
  const lngs = points.map(([, lng]) => lng);
  return [
    [Math.min(...lats), Math.min(...lngs)],
    [Math.max(...lats), Math.max(...lngs)]
  ];
}

function zoneColor(index: number) {
  const palette = ["#2563eb", "#12a594", "#8b5cf6", "#ef7b45", "#0f766e", "#be4bdb", "#3b82f6", "#84a52a"];
  return palette[index % palette.length];
}
