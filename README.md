# Revisor Maestro

Aplicacion Next.js para revisar diferencias entre los dias de despacho del maestro de clientes y los dias definidos para cada zona de reparto.

## Requisitos

- Node.js 20 o superior
- npm

## Ejecutar localmente

```bash
npm install
npm run dev
```

Luego abre `http://localhost:3000`.

## Archivos de entrada

- Maestro de clientes: Excel con columnas como `COD_CLIENTE_LOCAL`, `OFICINA_VENTAS`, `COD_ZONA_REPARTO`, `GRADO_LATITUD`, `GRADO_LONGITUD` y `FRECUENCIA_DESPACHO`.
- KML de zonas: archivo con poligonos de reparto. El nombre de cada zona puede incluir el codigo, por ejemplo `8006680067 - Coquimbo Centro`.
- Calendario de zonas: Excel con `COD_ZONA_REPARTO`, `ZONA_DE_REPARTO` y `DIAS_VISITA` o columnas por dia.

La plantilla inicial esta en `public/plantilla_zonas_reparto.xlsx` y tambien se puede descargar desde la app.

## Como llenar el calendario

En la columna `DIAS_VISITA` usa los codigos de dias esperados para la zona:

- `L`: lunes
- `M`: martes
- `W` o `X`: miercoles
- `J`: jueves
- `V`: viernes
- `S`: sabado
- `D`: domingo

Ejemplos: `LJ`, `MV`, `W`, `LMWJV`.

Tambien puedes dejar `DIAS_VISITA` vacio y marcar `SI`, `X` o `1` en las columnas `LUNES`, `MARTES`, `MIERCOLES`, `JUEVES`, `VIERNES`, `SABADO`, `DOMINGO`.

## Modos de analisis

- `Zona KML`: valida que el punto del cliente caiga dentro de un poligono KML cuyo codigo coincida con `COD_ZONA_REPARTO`.
- `Dias`: valida los dias del maestro contra el calendario de zonas.
- `Ambos`: levanta alerta si falla cualquiera de las dos validaciones.
