# Proceso de Auditoría FIFO - Documentación

## Guía Rápida de Ejecución

> Los scripts detectan automáticamente el **mes anterior** a la fecha actual.
> Si ejecutas en **marzo 2026**, procesan **febrero 2026**. No hay que configurar el mes.

### Flujo en 2 etapas

El proceso NO es un único bloque de 4 comandos. Va así:

**Etapa A — Sacar series del inventario base** (solo necesita `archivos_base.zip`)

| # | Archivo | De dónde sale |
|---|---------|---------------|
| 1 | `archivos_base.zip` | Inventario base (TXT pipe-delimited con carpetas `Fija/` y `Movil/`) |

```
REM Paso 1: Procesar base FIJA
python procesar_fifo.py archivos_base.zip FIJA

REM Paso 2: Procesar base MOVIL
python procesar_fifo.py archivos_base.zip MOVIL
```

Esto produce `VERIFICAR SALIDA CDVES FIJA {MES}.xlsx` y `VERIFICAR SALIDA CDVES MOVIL {MES}.xlsx`, que contienen las **series** del inventario base. Estas series son el insumo que se usa en SAP para extraer los movimientos del mes.

**Pausa** — con esos dos `.xlsx` en mano, se arma `Fifo.zip` (movimientos del mes filtrados por esas series).

**Etapa B — Verificar despachos FIFO** (necesita `Fifo.zip`)

| # | Archivo | De dónde sale |
|---|---------|---------------|
| 2 | `Fifo.zip` | Movimientos/transacciones del mes (carpetas `Fija/` y `Movil/` con archivos `.XLS`) |

```
REM Paso 3: Verificar FIJA
python verificar_fifo.py Fifo.zip {MES} FIJA

REM Paso 4: Verificar MOVIL
python verificar_fifo.py Fifo.zip {MES} MOVIL
```

> **{MES}**: Reemplazar con el mes que se procesa en MAYÚSCULAS (ej: `MAYO`).
> Los pasos 1-2 no necesitan el nombre del mes (lo detectan solos).
> Los pasos 3-4 sí necesitan el nombre del mes como argumento.

### Verificar resultado

Al terminar deben existir estos 2 archivos nuevos:

| Archivo | Qué contiene |
|---------|--------------|
| `Fifo_Fija_{MES}.xlsb` | Hoja "Data" + Hoja "TD" (CD VES vs ALMACENES) + Hoja "TD_U" (CD VES vs Almacen U) |
| `Fifo_Movil_{MES}.xlsb` | Hoja "Data" + Hoja "TD" con CD VES vs PDV / Televentas / Tienda Virtual / Corporativo |

Abrir cada archivo y revisar las hojas TD: al final tienen **Total Despacho**, **Total No Despachó** y **% Cumplimiento FIFO**.

> **Nota:** ya **no se generan** archivos `*_ALT.xlsb`. El reporte único utiliza Fecha Modificación como Mes/Año de CD VES (con fallback a Fecha Ingreso si Fecha Modificación es inválida).

---

## Ejecución con Claude Code

Si prefieres que Claude Code ejecute el proceso, di:

> "Correr Etapa A del FIFO" — con `archivos_base.zip` ya en `C:\ProKlauz\Fifo\`. Genera las series.
>
> "Correr Etapa B del FIFO" — cuando ya tengas `Fifo.zip` armado con los movimientos. Genera el reporte final.

Claude Code conoce el flujo en 2 etapas y sabe que entre A y B hay que ir a SAP a sacar los movimientos del mes usando las series generadas.

---

## Historial de Ejecuciones

| Mes | Fecha ejecución | FIJA % FIFO | MOVIL % FIFO | Notas |
|-----|-----------------|-------------|--------------|-------|
| Enero 2026 | 2026-02-26 | 98.78% | 99.97% | Bug corregido: columna "Número de serie" vs "Serie" |
| Abril 2026 | 2026-05-28 | TD 99.16% / TD_U 93.75% | PDV 99.72% / Corp 99.98% / Telev 99.90% / TV 99.73% | Primera corrida con: (a) lógica única usando Fecha Modificación para CD VES, (b) nuevo grupo `Almacen U` (Centro P008 + Almacén que empieza con U) y hoja `TD_U`, (c) sin archivos `_ALT` |
| Mayo 2026 | 2026-07-02 | TD 99.88% / TD_U 100.00% | PDV 99.65% / Corp 100.00% / Telev 100.00% / TV 99.51% | Se adaptó `procesar_fifo.py` al nuevo schema SAP: 139 cols (antes 62), alias multi-nombre en `TXT_COL_MAP`, fechas `DD/MM/YYYY` además de `YYYYMMDD`, `Estado` como columna directa (antes segunda ocurrencia de `Status`) |

---

## Detalle Técnico

### Archivos de Entrada

**`archivos_base.zip`** contiene:
```
Fija/Descarga_inventario_Infra_Valorado.txt   (inventario fija - ~770K filas)
Movil/Descarga_inventario_MOVIL.txt           (inventario móvil - ~442K filas)
Fija/Descarga_inventario_Infra_NO_Valorado.txt (NO se usa)
```

> **Formato legado**: También se soportan archivos `.xlsb` individuales en `C:\FIFO\archivos_base\`.

### Scripts

| Orden | Script | Propósito |
|-------|--------|-----------|
| 1 | `procesar_fifo.py` | Clasifica inventario por almacén agrupado y extrae series para verificación |
| 2 | `verificar_fifo.py` | Verifica series despachadas, filtra no despachadas, crea tabla dinámica con análisis FIFO |

### Flujo de archivos intermedios

```
archivos_base.zip
    ├─ procesar_fifo.py FIJA ──→ FIFO_PROCESADO_FIJA.xlsb
    │                          ──→ VERIFICAR SALIDA CDVES FIJA {MES}.xlsx
    └─ procesar_fifo.py MOVIL ─→ FIFO_PROCESADO_MOVIL.xlsb
                               ─→ VERIFICAR SALIDA CDVES MOVIL {MES}.xlsx

Fifo.zip + archivos intermedios
    ├─ verificar_fifo.py FIJA ──→ Fifo_Fija_{MES}.xlsb   (RESULTADO FINAL)
    └─ verificar_fifo.py MOVIL ─→ Fifo_Movil_{MES}.xlsb  (RESULTADO FINAL)
```

### Reglas de Clasificación

**FIJA**

| Grupo | Criterios | Filtro fecha |
|-------|-----------|--------------|
| **CD VES** | Amacén Tipo1='CD VES', Centro='P008', Almacén='P000', Estado='ALMA', Tipo Stock='01' | NO (Mes/Año = Fecha Modificación; si inválida, Fecha Ingreso) |
| **Almacen U** | Centro='P008', Almacén empieza con 'U', Estado='ALMA', Tipo Stock='01' | SÍ (Fecha Modificación en mes anterior; Mes/Año = Fecha Ingreso) |
| **ALMACENES** | Amacén Tipo1 NO en [CD VES, BAJAS DE INVENTARIO, BODEGA FALTANTE, LOGISTICA INVERSA INFRAESTRUCTURA, REFURBISHED] | SÍ (Fecha Modificación en mes anterior; Mes/Año = Fecha Ingreso) |

**MOVIL**

| Grupo | Criterios | Filtro fecha |
|-------|-----------|--------------|
| **CD VES** | Amacén Tipo1='CD VES', Centro='P008', Almacén='H000', Estado='ALMA', Tipo Stock='01' | NO (Mes/Año = Fecha Modificación; si inválida, Fecha Ingreso) |
| **PDV (CADENA)** | Amacén Tipo1='CADENA', Estado en {'ALMA ENCL','ENCL ALMA'}, Tipo Stock='02' | SÍ |
| **PDV (CAV/CAC)** | Amacén Tipo1 en {'CAV','CAC','OTROS CADS'}, Almacén en {'H000','H001'}, Estado='ALMA', Tipo Stock='01' | SÍ |
| **Televentas VES** | Centro='P150', Almacén='H000', Estado='ALMA', Tipo Stock en {'01','07'} | SÍ |
| **Tienda Virtual VES** | Centro='P102', Almacén='H000', Estado='ALMA', Tipo Stock en {'01','07'} | SÍ |
| **Corporativo Ves** | Centro='P101', Almacén en {'H001','H002'}, Estado='ALMA', Tipo Stock en {'01','07'} | SÍ |

### Contenido de las hojas TD

- **`TD` (FIJA)**: pivot con **CD VES + ALMACENES** (Almacen U oculto). FIFO entre estos dos grupos.
- **`TD_U` (FIJA)**: pivot con **CD VES + Almacen U** (ALMACENES oculto). FIFO entre estos dos grupos.
- **`TD` (MOVIL)**: pivot con CD VES como fuente y PDV / Televentas / Tienda Virtual / Corporativo como destinos.
- **Tabla dinámica**: Material x Mes/Año x Almacén Agrupado → Cuenta de Series.
- **Columna "NO DESPACHÓ"**: series de meses anteriores que no se despacharon antes que meses posteriores.
- **Columna "NO CUMPLIÓ FIFO"**: violaciones FIFO (se despachó inventario más nuevo antes que más viejo).
- **Resumen**: Total Despacho, Total No Despachó, % Cumplimiento FIFO (uno por hoja TD).

### Estructura del Directorio

```
C:\FIFO\
├── procesar_fifo.py          # Script principal paso 1
├── verificar_fifo.py         # Script principal paso 2
├── PROCESO_FIFO.md           # Esta documentación
├── archivos_base.zip         # Inventario base
├── Fifo.zip                  # Movimientos del mes
├── historico\                # Archivos procesados anteriormente
└── obsoleto\                 # Scripts que ya no se usan
```

### Requisitos del Sistema

- Windows con Excel instalado (se usa COM automation)
- Python 3.x
- Librerías Python: `pywin32`, `pandas`, `openpyxl`, `pyxlsb`
- Si se ejecuta desde WSL, usar `python.exe` (no `python3`)
