import ExcelJS from "exceljs";

export type DayKey = "L" | "M" | "W" | "J" | "V" | "S" | "D";

export const dayLabels: Record<DayKey, string> = {
  L: "Lun",
  M: "Mar",
  W: "Mie",
  J: "Jue",
  V: "Vie",
  S: "Sab",
  D: "Dom"
};

export const dayNames: Record<DayKey, string> = {
  L: "Lunes",
  M: "Martes",
  W: "Miercoles",
  J: "Jueves",
  V: "Viernes",
  S: "Sabado",
  D: "Domingo"
};

const dayColumns: Record<DayKey, string[]> = {
  L: ["PROM_LUNES", "LUNES"],
  M: ["PROM_MARTES", "MARTES"],
  W: ["PROM_MIERCOLES", "MIERCOLES"],
  J: ["PROM_JUEVES", "JUEVES"],
  V: ["PROM_VIERNES", "VIERNES"],
  S: ["PROM_SABADO", "SABADO"],
  D: ["PROM_DOMINGO", "DOMINGO"]
};

const dayOrder: DayKey[] = ["L", "M", "W", "J", "V", "S", "D"];

export type ClientRecord = {
  id: string;
  office: string;
  name: string;
  address: string;
  customerType: string;
  chain: string;
  population: string;
  zoneCode: string;
  zoneName: string;
  size: string;
  lat: number | null;
  lng: number | null;
  actualDays: DayKey[];
  actualCode: string;
  yearlyKilos: number;
  yearlyBoxes: number;
};

export type ZoneSchedule = {
  zoneCode: string;
  zoneName: string;
  expectedDays: DayKey[];
  expectedCode: string;
};

export type ZonePolygon = {
  id: string;
  name: string;
  code: string;
  rings: Array<Array<[number, number]>>;
};

export type KmlFeatureInfo = {
  id: string;
  name: string;
  code: string;
  description: string;
  polygonCount: number;
  ringCount: number;
  pointCount: number;
  minLat: number | null;
  maxLat: number | null;
  minLng: number | null;
  maxLng: number | null;
  centroidLat: number | null;
  centroidLng: number | null;
  areaKm2: number | null;
  attributes: Record<string, string>;
};

export type AnalysisMode = "geo" | "days" | "combined";
export type ReviewStatus = "match" | "mismatch" | "noSchedule" | "noLocation" | "outsideZone" | "noKml";
export type DayStatus = "match" | "mismatch" | "noSchedule" | "noLocation";
export type GeoStatus = "match" | "mismatch" | "outsideZone" | "noLocation" | "noKml";

export type ReviewRecord = ClientRecord & {
  expectedDays: DayKey[];
  expectedCode: string;
  status: ReviewStatus;
  dayStatus: DayStatus;
  geoStatus: GeoStatus;
  kmlZoneCode: string;
  kmlZoneName: string;
  missingDays: DayKey[];
  extraDays: DayKey[];
  issues: string[];
};

type SheetRow = Record<string, unknown>;

type ExcelCellPrimitive = string | number | boolean | Date | null | undefined;

export async function parseExcelRows(file: File): Promise<SheetRow[]> {
  const buffer = await file.arrayBuffer();
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);
  const worksheet = workbook.worksheets[0];
  if (!worksheet) return [];

  const headers = getWorksheetHeaders(worksheet);
  const rows: SheetRow[] = [];

  worksheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return;
    const record: SheetRow = {};
    let hasValue = false;

    headers.forEach((header, index) => {
      const value = cellToValue(row.getCell(index + 1).value);
      record[header] = value;
      if (value !== "") hasValue = true;
    });

    if (hasValue) rows.push(record);
  });

  return rows;
}

function getWorksheetHeaders(worksheet: ExcelJS.Worksheet) {
  const headerRow = worksheet.getRow(1);
  const headers: string[] = [];
  headerRow.eachCell({ includeEmpty: false }, (cell, columnNumber) => {
    headers[columnNumber - 1] = String(cellToValue(cell.value) ?? "").trim();
  });
  return headers;
}

function cellToValue(value: ExcelJS.CellValue): ExcelCellPrimitive {
  if (value === null || value === undefined) return "";
  if (value instanceof Date) return value;
  if (typeof value !== "object") return value;
  if ("result" in value) return cellToValue(value.result as ExcelJS.CellValue);
  if ("text" in value) return value.text ?? "";
  if ("richText" in value) return value.richText.map((part) => part.text).join("");
  if ("hyperlink" in value && "text" in value) return String(value.text ?? "");
  return String(value);
}

function styleHeader(worksheet: ExcelJS.Worksheet) {
  const header = worksheet.getRow(1);
  header.font = { bold: true, color: { argb: "FFFFFFFF" } };
  header.fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FF111827" }
  };
  header.alignment = { vertical: "middle", horizontal: "center" };
}

function addJsonSheet(workbook: ExcelJS.Workbook, name: string, rows: Array<Record<string, unknown>>, widths?: number[]) {
  const worksheet = workbook.addWorksheet(name);
  const headers = rows[0] ? Object.keys(rows[0]) : ["SIN_DATOS"];
  worksheet.columns = headers.map((header, index) => ({
    header,
    key: header,
    width: widths?.[index] ?? Math.min(Math.max(header.length + 4, 12), 34)
  }));

  if (rows.length) {
    rows.forEach((row) => worksheet.addRow(row));
  }

  styleHeader(worksheet);
  worksheet.views = [{ state: "frozen", ySplit: 1 }];
  worksheet.autoFilter = {
    from: { row: 1, column: 1 },
    to: { row: Math.max(1, worksheet.rowCount), column: headers.length }
  };
  return worksheet;
}

export function buildTemplateWorkbook(zones: Array<Pick<ZoneSchedule, "zoneCode" | "zoneName">>) {
  const rows = zones.length
    ? zones
    : [{ zoneCode: "8000000000", zoneName: "Zona ejemplo" }];
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "Revisor Maestro";
  const worksheet = workbook.addWorksheet("Zonas");
  worksheet.columns = [
    { header: "COD_ZONA_REPARTO", key: "zoneCode", width: 18 },
    { header: "ZONA_DE_REPARTO", key: "zoneName", width: 32 },
    { header: "DIAS_VISITA", key: "visitDays", width: 16 },
    { header: "LUNES", key: "monday", width: 10 },
    { header: "MARTES", key: "tuesday", width: 10 },
    { header: "MIERCOLES", key: "wednesday", width: 12 },
    { header: "JUEVES", key: "thursday", width: 10 },
    { header: "VIERNES", key: "friday", width: 10 },
    { header: "SABADO", key: "saturday", width: 10 },
    { header: "DOMINGO", key: "sunday", width: 10 }
  ];
  rows.forEach((zone) => {
    worksheet.addRow({
      zoneCode: zone.zoneCode,
      zoneName: zone.zoneName,
      visitDays: "",
      monday: "",
      tuesday: "",
      wednesday: "",
      thursday: "",
      friday: "",
      saturday: "",
      sunday: ""
    });
  });
  styleHeader(worksheet);
  worksheet.views = [{ state: "frozen", ySplit: 1 }];

  const guide = workbook.addWorksheet("Guia");
  guide.columns = [
    { header: "Campo", key: "field", width: 18 },
    { header: "Detalle", key: "detail", width: 110 }
  ];
  guide.addRows([
    {
      field: "DIAS_VISITA",
      detail: "Usa codigos compactos: L=lunes, M=martes, W o X=miercoles, J=jueves, V=viernes, S=sabado, D=domingo."
    },
    { field: "Ejemplos", detail: "LJ, MV, W, LMWJV" },
    {
      field: "Columnas por dia",
      detail: "Tambien puedes dejar DIAS_VISITA vacio y marcar SI, X o 1 en las columnas LUNES...DOMINGO."
    }
  ]);
  styleHeader(guide);
  return workbook;
}

export async function downloadWorkbook(workbook: ExcelJS.Workbook, filename: string) {
  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

export function buildReviewWorkbook(records: ReviewRecord[], mode: AnalysisMode) {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "Revisor Maestro";
  const clientRows = records.map((record) => reviewRow(record, mode));
  const alertRows = records.filter((record) => record.status !== "match").map((record) => reviewRow(record, mode));
  addJsonSheet(workbook, "Resumen", [
    {
      ANALISIS: analysisModeLabel(mode),
      CLIENTES: records.length,
      ALERTAS: alertRows.length,
      OK: records.filter((record) => record.status === "match").length,
      GENERADO: new Date().toLocaleString("es-CL")
    }
  ]);
  addJsonSheet(workbook, "Alertas", alertRows, reviewColumnWidths());
  addJsonSheet(workbook, "Clientes", clientRows, reviewColumnWidths());
  return workbook;
}

export function buildKmlWorkbook(features: KmlFeatureInfo[]) {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "Revisor Maestro";
  const attributeKeys = Array.from(new Set(features.flatMap((feature) => Object.keys(feature.attributes)))).sort((a, b) =>
    a.localeCompare(b, "es")
  );
  const rows = features.map((feature, index) => {
    const row: Record<string, unknown> = {
      ORDEN: index + 1,
      NOMBRE: feature.name,
      CODIGO_ZONA: feature.code,
      DESCRIPCION: feature.description,
      POLIGONOS: feature.polygonCount,
      ANILLOS: feature.ringCount,
      PUNTOS: feature.pointCount,
      LAT_MIN: feature.minLat ?? "",
      LAT_MAX: feature.maxLat ?? "",
      LNG_MIN: feature.minLng ?? "",
      LNG_MAX: feature.maxLng ?? "",
      CENTROIDE_LAT: feature.centroidLat ?? "",
      CENTROIDE_LNG: feature.centroidLng ?? "",
      AREA_KM2_APROX: feature.areaKm2 ?? ""
    };

    for (const key of attributeKeys) {
      row[`KML_${safeColumnName(key)}`] = feature.attributes[key] ?? "";
    }
    return row;
  });

  addJsonSheet(workbook, "Resumen", [
    {
      POLIGONOS_KML: features.length,
      COLUMNAS_EXTENDIDAS: attributeKeys.length,
      PUNTOS_COORDENADAS: features.reduce((sum, feature) => sum + feature.pointCount, 0),
      AREA_KM2_APROX: roundNumber(features.reduce((sum, feature) => sum + (feature.areaKm2 ?? 0), 0)),
      GENERADO: new Date().toLocaleString("es-CL")
    }
  ]);
  addJsonSheet(workbook, "KML", rows);
  return workbook;
}

export function parseClients(rows: SheetRow[]): ClientRecord[] {
  return rows
    .map((row, index) => {
      const get = rowGetter(row);
      const frequency = asText(get(["FRECUENCIA_DESPACHO", "DIAS_DESPACHO"]));
      const actualDays = parseDayCode(frequency);
      const fallbackDays = daysFromPromColumns(get);

      return {
        id: asText(get(["COD_CLIENTE_LOCAL", "CODIGO_CLIENTE"])) || `fila-${index + 2}`,
        office: asText(get(["OFICINA_VENTAS", "OFICINA DE VENTAS"])) || "Sin oficina",
        name: asText(get(["CLIENTE_LOCAL", "CLIENTE"])) || "Sin nombre",
        address: asText(get(["CALLE", "DIRECCION"])),
        customerType: asText(get(["TIPO_DE_CLIENTE", "TIPO CLIENTE"])),
        chain: asText(get(["CADENA"])),
        population: asText(get(["POBLACION", "COMUNA"])),
        zoneCode: asText(get(["COD_ZONA_REPARTO", "CODIGO_ZONA_REPARTO"])),
        zoneName: asText(get(["ZONA_DE_REPARTO", "ZONA REPARTO"])),
        size: asText(get(["TALLA"])),
        lat: asNumber(get(["GRADO_LATITUD", "LATITUD", "LAT"])),
        lng: asNumber(get(["GRADO_LONGITUD", "LONGITUD", "LON", "LNG"])),
        actualDays: actualDays.length ? actualDays : fallbackDays,
        actualCode: actualDays.length ? daysToCode(actualDays) : daysToCode(fallbackDays),
        yearlyKilos: asNumber(get(["PROM_KILOS_ANIO", "PROM_KILOS_ANO", "KILOS_ANIO", "KILOS_ANO"])) ?? 0,
        yearlyBoxes: asNumber(get(["PROM_CAJAS_ANIO", "PROM_CAJAS_ANO", "CAJAS_ANIO", "CAJAS_ANO"])) ?? 0
      };
    })
    .filter((client) => client.id || client.name || client.zoneCode);
}

export function parseSchedules(rows: SheetRow[]): ZoneSchedule[] {
  const seen = new Map<string, ZoneSchedule>();

  for (const row of rows) {
    const get = rowGetter(row);
    const zoneCode = asText(get(["COD_ZONA_REPARTO", "CODIGO_ZONA_REPARTO", "ZONA"]));
    if (!zoneCode) continue;
    const zoneKey = zoneCodeKey(zoneCode);

    const explicitDays = parseDayCode(asText(get(["DIAS_VISITA", "DIAS_DESPACHO"])));
    const booleanDays = daysFromBooleanColumns(get);
    const expectedDays = explicitDays.length ? explicitDays : booleanDays;

    seen.set(zoneKey, {
      zoneCode,
      zoneName: asText(get(["ZONA_DE_REPARTO", "ZONA REPARTO", "NOMBRE_ZONA"])) || zoneCode,
      expectedDays,
      expectedCode: daysToCode(expectedDays)
    });
  }

  return Array.from(seen.values());
}

export function compareClients(
  clients: ClientRecord[],
  schedules: ZoneSchedule[],
  zones: ZonePolygon[] = [],
  mode: AnalysisMode = "days"
): ReviewRecord[] {
  const schedulesByZone = new Map(schedules.map((schedule) => [zoneCodeKey(schedule.zoneCode), schedule]));
  const indexedZones = zones.map((zone) => ({ zone, bounds: getZoneBounds(zone) }));

  return clients.map((client) => {
    const masterZoneKey = zoneCodeKey(client.zoneCode);
    const schedule = schedulesByZone.get(masterZoneKey);
    const expectedDays = schedule?.expectedDays ?? [];
    const missingDays = expectedDays.filter((day) => !client.actualDays.includes(day));
    const extraDays = client.actualDays.filter((day) => !expectedDays.includes(day));
    const hasLocation = typeof client.lat === "number" && typeof client.lng === "number";
    let dayStatus: DayStatus = "match";

    if (!hasLocation) dayStatus = "noLocation";
    else if (!schedule || expectedDays.length === 0) dayStatus = "noSchedule";
    else if (missingDays.length || extraDays.length) dayStatus = "mismatch";

    const kmlZone = hasLocation ? findContainingZone(client.lat as number, client.lng as number, indexedZones) : null;
    let geoStatus: GeoStatus = "match";

    if (!hasLocation) geoStatus = "noLocation";
    else if (!zones.length) geoStatus = "noKml";
    else if (!kmlZone) geoStatus = "outsideZone";
    else if (kmlZone.code && client.zoneCode && zoneCodeKey(kmlZone.code) !== masterZoneKey) geoStatus = "mismatch";

    const status = resolveStatus(dayStatus, geoStatus, mode);
    const issues = buildIssues(
      dayStatus,
      geoStatus,
      client.zoneCode,
      kmlZone?.code ?? "",
      kmlZone?.name ?? "",
      missingDays,
      extraDays
    );

    return {
      ...client,
      expectedDays,
      expectedCode: schedule?.expectedCode ?? "",
      status,
      dayStatus,
      geoStatus,
      kmlZoneCode: kmlZone?.code ?? "",
      kmlZoneName: kmlZone?.name ?? "",
      missingDays,
      extraDays,
      issues
    };
  });
}

export function parseKml(kmlText: string): ZonePolygon[] {
  const parser = new DOMParser();
  const xml = parser.parseFromString(kmlText, "text/xml");
  const placemarks = Array.from(xml.getElementsByTagName("Placemark"));

  return placemarks
    .map((placemark, index) => {
      const name = textFromFirst(placemark, "name");
      const polygonNodes = Array.from(placemark.getElementsByTagName("Polygon"));
      const rings = polygonNodes
        .flatMap((polygon) => Array.from(polygon.getElementsByTagName("outerBoundaryIs")))
        .map((outer) => textFromFirst(outer, "coordinates"))
        .map(parseCoordinates)
        .filter((ring) => ring.length >= 3);

      if (!rings.length) return null;

      return {
        id: `${index}-${name || "zona"}`,
        name: cleanName(name || `Zona ${index + 1}`),
        code: extractZoneCode(name),
        rings
      };
    })
    .filter(Boolean) as ZonePolygon[];
}

export function parseKmlFeatureInfo(kmlText: string): KmlFeatureInfo[] {
  const parser = new DOMParser();
  const xml = parser.parseFromString(kmlText, "text/xml");
  const placemarks = Array.from(xml.getElementsByTagName("Placemark"));

  return placemarks.map((placemark, index) => {
    const name = cleanName(textFromFirst(placemark, "name") || `Zona ${index + 1}`);
    const description = textFromFirst(placemark, "description");
    const polygonNodes = Array.from(placemark.getElementsByTagName("Polygon"));
    const rings = polygonNodes
      .flatMap((polygon) => Array.from(polygon.getElementsByTagName("outerBoundaryIs")))
      .map((outer) => textFromFirst(outer, "coordinates"))
      .map(parseCoordinates)
      .filter((ring) => ring.length >= 3);
    const points = rings.flat();
    const bounds = boundsFromPoints(points);
    const centroid = centroidFromPoints(points);

    return {
      id: `${index}-${name}`,
      name,
      code: extractZoneCode(name),
      description,
      polygonCount: polygonNodes.length,
      ringCount: rings.length,
      pointCount: points.length,
      minLat: bounds?.minLat ?? null,
      maxLat: bounds?.maxLat ?? null,
      minLng: bounds?.minLng ?? null,
      maxLng: bounds?.maxLng ?? null,
      centroidLat: centroid?.[0] ?? null,
      centroidLng: centroid?.[1] ?? null,
      areaKm2: rings.length ? roundNumber(rings.reduce((sum, ring) => sum + approximateRingAreaKm2(ring), 0)) : null,
      attributes: {
        ...extractDescriptionAttributes(description),
        ...extractKmlAttributes(placemark)
      }
    };
  });
}

export function findZoneForPoint(lat: number, lng: number, zones: ZonePolygon[]) {
  const indexedZones = zones.map((zone) => ({ zone, bounds: getZoneBounds(zone) }));
  return findContainingZone(lat, lng, indexedZones);
}

export function clientsInsideZone(clients: ClientRecord[], zone: ZonePolygon) {
  const bounds = getZoneBounds(zone);
  return clients.filter((client) => {
    if (typeof client.lat !== "number" || typeof client.lng !== "number") return false;
    if (!bounds || !isPointInBounds(client.lat, client.lng, bounds)) return false;
    return zone.rings.some((ring) => isPointInRing(client.lat as number, client.lng as number, ring));
  });
}

export function centroidFromClients(clients: ClientRecord[], fallbackZone?: ZonePolygon): [number, number] | null {
  const located = clients.filter((client) => typeof client.lat === "number" && typeof client.lng === "number");
  if (located.length) {
    const lat = located.reduce((sum, client) => sum + (client.lat as number), 0) / located.length;
    const lng = located.reduce((sum, client) => sum + (client.lng as number), 0) / located.length;
    return [lat, lng];
  }
  if (!fallbackZone) return null;
  const points = fallbackZone.rings.flat();
  if (!points.length) return null;
  const lat = points.reduce((sum, point) => sum + point[0], 0) / points.length;
  const lng = points.reduce((sum, point) => sum + point[1], 0) / points.length;
  return [lat, lng];
}

export function daysToCode(days: DayKey[]) {
  return dayOrder.filter((day) => days.includes(day)).join("");
}

export function formatDays(days: DayKey[]) {
  return days.length ? dayOrder.filter((day) => days.includes(day)).map((day) => dayLabels[day]).join(", ") : "Sin dias";
}

export function normalizeText(value: unknown) {
  return String(value ?? "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .trim()
    .toUpperCase();
}

export function zoneCodeKey(value: unknown) {
  const normalized = normalizeText(value).replace(/[^A-Z0-9]/g, "");
  if (!normalized) return "";
  return normalized.replace(/^0+(?=\d)/, "");
}

function rowGetter(row: SheetRow) {
  const entries = new Map(Object.entries(row).map(([key, value]) => [normalizeHeader(key), value]));
  return (candidates: string[]) => {
    for (const candidate of candidates) {
      const value = entries.get(normalizeHeader(candidate));
      if (value !== undefined && value !== null && String(value).trim() !== "") return value;
    }
    return "";
  };
}

function normalizeHeader(header: string) {
  return normalizeText(header).replace(/[^A-Z0-9]/g, "");
}

function asText(value: unknown) {
  return String(value ?? "").trim();
}

function asNumber(value: unknown) {
  if (value === null || value === undefined || value === "") return null;
  const parsed = Number(String(value).replace(",", "."));
  return Number.isFinite(parsed) ? parsed : null;
}

function parseDayCode(value: string): DayKey[] {
  const normalizedValue = normalizeText(value);
  const byWords: DayKey[] = [];
  const wordMap: Array<[string, DayKey]> = [
    ["LUNES", "L"],
    ["MARTES", "M"],
    ["MIERCOLES", "W"],
    ["JUEVES", "J"],
    ["VIERNES", "V"],
    ["SABADO", "S"],
    ["DOMINGO", "D"]
  ];

  for (const [word, day] of wordMap) {
    if (normalizedValue.includes(word)) byWords.push(day);
  }
  if (byWords.length) return uniqueDays(byWords);

  const tokenDays = normalizedValue
    .split(/[^A-Z0-9]+/)
    .map(dayFromToken)
    .filter((day): day is DayKey => Boolean(day));
  if (tokenDays.length) return uniqueDays(tokenDays);

  return uniqueDays(
    Array.from(normalizedValue.replace(/[^LMWXJVSD]/g, ""))
      .map((letter) => (letter === "X" ? "W" : letter))
      .filter((letter): letter is DayKey => dayOrder.includes(letter as DayKey))
  );
}

function dayFromToken(token: string): DayKey | null {
  const tokenMap: Record<string, DayKey> = {
    L: "L",
    LU: "L",
    LUN: "L",
    M: "M",
    MA: "M",
    MAR: "M",
    W: "W",
    X: "W",
    MI: "W",
    MIE: "W",
    J: "J",
    JU: "J",
    JUE: "J",
    V: "V",
    VI: "V",
    VIE: "V",
    S: "S",
    SA: "S",
    SAB: "S",
    D: "D",
    DO: "D",
    DOM: "D"
  };

  return tokenMap[token] ?? null;
}

function daysFromPromColumns(get: (candidates: string[]) => unknown) {
  return dayOrder.filter((day) => {
    const value = asNumber(get(dayColumns[day]));
    return value !== null && value > 0;
  });
}

function daysFromBooleanColumns(get: (candidates: string[]) => unknown) {
  return dayOrder.filter((day) => {
    const value = normalizeText(get(dayColumns[day]));
    return ["SI", "S", "YES", "Y", "TRUE", "1", "X"].includes(value);
  });
}

function uniqueDays(days: DayKey[]) {
  return dayOrder.filter((day) => days.includes(day));
}

function textFromFirst(parent: Element, tagName: string) {
  return parent.getElementsByTagName(tagName)[0]?.textContent?.trim() ?? "";
}

function parseCoordinates(value: string) {
  return value
    .trim()
    .split(/\s+/)
    .map((chunk) => {
      const [lng, lat] = chunk.split(",").map(Number);
      return [lat, lng] as [number, number];
    })
    .filter(([lat, lng]) => Number.isFinite(lat) && Number.isFinite(lng));
}

function extractZoneCode(name: string) {
  return name.match(/\d{8,}/)?.[0] ?? "";
}

function cleanName(name: string) {
  return name.replace(/\s+/g, " ").trim();
}

function extractKmlAttributes(placemark: Element) {
  const attributes: Record<string, string> = {};

  Array.from(placemark.getElementsByTagName("Data")).forEach((node) => {
    const key = node.getAttribute("name")?.trim();
    if (!key) return;
    attributes[key] = textFromFirst(node, "value");
  });

  Array.from(placemark.getElementsByTagName("SimpleData")).forEach((node) => {
    const key = node.getAttribute("name")?.trim();
    if (!key) return;
    attributes[key] = node.textContent?.trim() ?? "";
  });

  return attributes;
}

function extractDescriptionAttributes(description: string) {
  const attributes: Record<string, string> = {};
  const plain = description
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/<\/p>|<\/div>|<\/li>/gi, "\n")
    .replace(/<[^>]+>/g, "")
    .replace(/&nbsp;/gi, " ")
    .replace(/&amp;/gi, "&")
    .replace(/&lt;/gi, "<")
    .replace(/&gt;/gi, ">");

  plain.split(/\n+/).forEach((line) => {
    const [rawKey, ...valueParts] = line.split(":");
    const key = rawKey?.trim();
    const value = valueParts.join(":").trim();
    if (key && value) attributes[key] = value;
  });

  return attributes;
}

function safeColumnName(value: string) {
  const normalized = normalizeText(value).replace(/[^A-Z0-9]+/g, "_").replace(/^_+|_+$/g, "");
  return normalized || "DATO";
}

function roundNumber(value: number) {
  return Math.round(value * 1000) / 1000;
}

function boundsFromPoints(points: Array<[number, number]>): Bounds | null {
  if (!points.length) return null;
  const lats = points.map(([lat]) => lat);
  const lngs = points.map(([, lng]) => lng);
  return {
    minLat: Math.min(...lats),
    maxLat: Math.max(...lats),
    minLng: Math.min(...lngs),
    maxLng: Math.max(...lngs)
  };
}

function centroidFromPoints(points: Array<[number, number]>): [number, number] | null {
  if (!points.length) return null;
  const lat = points.reduce((sum, point) => sum + point[0], 0) / points.length;
  const lng = points.reduce((sum, point) => sum + point[1], 0) / points.length;
  return [roundNumber(lat), roundNumber(lng)];
}

function approximateRingAreaKm2(ring: Array<[number, number]>) {
  if (ring.length < 3) return 0;
  const meanLat = ring.reduce((sum, point) => sum + point[0], 0) / ring.length;
  const kmPerLat = 110.574;
  const kmPerLng = 111.32 * Math.cos((meanLat * Math.PI) / 180);
  const projected = ring.map(([lat, lng]) => [lng * kmPerLng, lat * kmPerLat]);
  let area = 0;

  for (let index = 0; index < projected.length; index += 1) {
    const [x1, y1] = projected[index];
    const [x2, y2] = projected[(index + 1) % projected.length];
    area += x1 * y2 - x2 * y1;
  }

  return Math.abs(area) / 2;
}

function reviewRow(record: ReviewRecord, mode: AnalysisMode) {
  return {
    MODO_ANALISIS: analysisModeLabel(mode),
    ESTADO: reviewStatusLabel(record.status),
    ALERTAS: record.issues.join(" | "),
    COD_CLIENTE_LOCAL: record.id,
    CLIENTE_LOCAL: record.name,
    OFICINA_VENTAS: record.office,
    DIRECCION: record.address,
    POBLACION: record.population,
    TIPO_CLIENTE: record.customerType,
    CADENA: record.chain,
    COD_ZONA_REPARTO_MAESTRO: record.zoneCode,
    ZONA_REPARTO_MAESTRO: record.zoneName,
    COD_ZONA_REPARTO_KML: record.kmlZoneCode,
    ZONA_REPARTO_KML: record.kmlZoneName,
    ESTADO_ZONA_KML: reviewStatusLabel(record.geoStatus),
    DIAS_MAESTRO: formatDays(record.actualDays),
    DIAS_CALENDARIO_ZONA: formatDays(record.expectedDays),
    ESTADO_DIAS: reviewStatusLabel(record.dayStatus),
    DIAS_FALTANTES: formatDays(record.missingDays),
    DIAS_SOBRANTES: formatDays(record.extraDays),
    LATITUD: record.lat ?? "",
    LONGITUD: record.lng ?? "",
    PROM_KILOS_ANIO: record.yearlyKilos,
    PROM_CAJAS_ANIO: record.yearlyBoxes
  };
}

function reviewColumnWidths() {
  return [
    16,
    16,
    58,
    18,
    34,
    18,
    34,
    18,
    18,
    18,
    24,
    30,
    24,
    34,
    18,
    18,
    22,
    18,
    18,
    18,
    14,
    14,
    16,
    16
  ];
}

function analysisModeLabel(mode: AnalysisMode) {
  const labels: Record<AnalysisMode, string> = {
    geo: "Zona KML",
    days: "Dias",
    combined: "Ambos"
  };
  return labels[mode];
}

function reviewStatusLabel(status: ReviewStatus | DayStatus | GeoStatus) {
  const labels: Record<ReviewStatus, string> = {
    match: "OK",
    mismatch: "Alerta",
    noSchedule: "Sin calendario",
    noLocation: "Sin ubicacion",
    outsideZone: "Fuera de zona",
    noKml: "Sin KML"
  };
  return labels[status as ReviewStatus];
}

function resolveStatus(dayStatus: DayStatus, geoStatus: GeoStatus, mode: AnalysisMode): ReviewStatus {
  if (mode === "days") return dayStatus;
  if (mode === "geo") return geoStatus;

  if (geoStatus === "noLocation" || dayStatus === "noLocation") return "noLocation";
  if (geoStatus === "mismatch" || dayStatus === "mismatch") return "mismatch";
  if (geoStatus === "outsideZone") return "outsideZone";
  if (geoStatus === "noKml") return "noKml";
  if (dayStatus === "noSchedule") return "noSchedule";
  return "match";
}

function buildIssues(
  dayStatus: DayStatus,
  geoStatus: GeoStatus,
  masterZoneCode: string,
  kmlZoneCode: string,
  kmlZoneName: string,
  missingDays: DayKey[],
  extraDays: DayKey[]
) {
  const issues: string[] = [];

  if (geoStatus === "mismatch") {
    issues.push(
      `Zona distinta: maestro ${masterZoneCode || "sin codigo"}; ubicacion KML ${kmlZoneCode || "sin codigo"}${kmlZoneName ? ` (${kmlZoneName})` : ""}`
    );
  } else if (geoStatus === "outsideZone") {
    issues.push("Ubicacion fuera de las zonas KML cargadas");
  } else if (geoStatus === "noKml") {
    issues.push("Sin KML cargado para validar zona geografica");
  } else if (geoStatus === "noLocation") {
    issues.push("Cliente sin latitud/longitud valida");
  }

  if (dayStatus === "mismatch") {
    issues.push(`Dias distintos: faltan ${formatDays(missingDays)}; sobran ${formatDays(extraDays)}`);
  } else if (dayStatus === "noSchedule") {
    issues.push("Zona sin calendario de visita cargado");
  }

  return issues;
}

function findContainingZone(
  lat: number,
  lng: number,
  zones: Array<{ zone: ZonePolygon; bounds: Bounds | null }>
) {
  return zones.find(({ zone, bounds }) => {
    if (!bounds || !isPointInBounds(lat, lng, bounds)) return false;
    return zone.rings.some((ring) => isPointInRing(lat, lng, ring));
  })?.zone ?? null;
}

type Bounds = {
  minLat: number;
  maxLat: number;
  minLng: number;
  maxLng: number;
};

function getZoneBounds(zone: ZonePolygon): Bounds | null {
  const points = zone.rings.flat();
  return boundsFromPoints(points);
}

function isPointInBounds(lat: number, lng: number, bounds: Bounds) {
  return lat >= bounds.minLat && lat <= bounds.maxLat && lng >= bounds.minLng && lng <= bounds.maxLng;
}

function isPointInRing(lat: number, lng: number, ring: Array<[number, number]>) {
  let inside = false;
  for (let current = 0, previous = ring.length - 1; current < ring.length; previous = current++) {
    const [currentLat, currentLng] = ring[current];
    const [previousLat, previousLng] = ring[previous];
    const intersects =
      currentLng > lng !== previousLng > lng &&
      lat < ((previousLat - currentLat) * (lng - currentLng)) / (previousLng - currentLng) + currentLat;

    if (intersects) inside = !inside;
  }
  return inside;
}
