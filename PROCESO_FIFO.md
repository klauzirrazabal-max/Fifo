# Proceso de Auditoría FIFO - Documentación

## Guía Rápida de Ejecución

> Los scripts detectan automáticamente el **mes anterior** a la fecha actual.
> Si ejecutas en **marzo 2026**, procesan **febrero 2026**. No hay que configurar el mes.

### Prerequisitos

Colocar estos 2 archivos en `C:\FIFO\` antes de empezar:

| # | Archivo | De dónde sale |
|---|---------|---------------|
| 1 | `archivos_base.zip` | Inventario base (TXT pipe-delimited con carpetas `Fija/` y `Movil/`) |
| 2 | `Fifo.zip` | Movimientos/transacciones del mes (carpetas `Fija/` y `Movil/` con archivos `.XLS`) |

### Ejecución (4 comandos en orden)

Ejecutar desde `C:\FIFO\` con Python de Windows:

```
REM Paso 1: Procesar base FIJA
python procesar_fifo.py archivos_base.zip FIJA

REM Paso 2: Procesar base MOVIL
python procesar_fifo.py archivos_base.zip MOVIL

REM Paso 3: Verificar FIJA (necesita Fifo.zip)
python verificar_fifo.py Fifo.zip {MES} FIJA

REM Paso 4: Verificar MOVIL (necesita Fifo.zip)
python verificar_fifo.py Fifo.zip {MES} MOVIL
```

> **{MES}**: Reemplazar con el mes que se procesa en MAYÚSCULAS (ej: `FEBRERO`).
> Los pasos 1-2 no necesitan el nombre del mes (lo detectan solos).
> Los pasos 3-4 sí necesitan el nombre del mes como argumento.

### Verificar resultado

Al terminar deben existir estos 2 archivos nuevos:

| Archivo | Qué contiene |
|---------|--------------|
| `Fifo_Fija_{MES}.xlsb` | Hoja "Data" + Hoja "TD" con tabla dinámica y % FIFO |
| `Fifo_Movil_{MES}.xlsb` | Igual pero para inventario móvil |

Abrir cada archivo y revisar la hoja "TD": al final tiene **Total Despacho**, **Total No Despachó** y **% Cumplimiento FIFO**.

---

## Ejecución con Claude Code

Si prefieres que Claude Code ejecute el proceso, solo di:

> "Ejecutar proceso FIFO"

Claude Code sabe ejecutar los 4 pasos en orden. Solo asegúrate de tener `archivos_base.zip` y `Fifo.zip` en `C:\FIFO\` antes de pedirlo.

---

## Historial de Ejecuciones

| Mes | Fecha ejecución | FIJA % FIFO | MOVIL % FIFO | Notas |
|-----|-----------------|-------------|--------------|-------|
| Enero 2026 | 2026-02-26 | 98.78% | 99.97% | Bug corregido: columna "Número de serie" vs "Serie" |

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
| **CD VES** | Amacén Tipo1='CD VES', Centro='P008', Almacén='P000', Estado='ALMA', Tipo Stock='01' | NO (usa Fecha Ingreso como Mes/Año) |
| **ALMACENES** | Amacén Tipo1 NO en [CD VES, BAJAS DE INVENTARIO, BODEGA FALTANTE, LOGISTICA INVERSA INFRAESTRUCTURA, REFURBISHED] | SÍ (Fecha Modificación en mes anterior) |

**MOVIL**

| Grupo | Criterios | Filtro fecha |
|-------|-----------|--------------|
| **CD VES** | Amacén Tipo1='CD VES', Centro='P008', Almacén='H000', Estado='ALMA', Tipo Stock='01' | NO |
| **PDV (CADENA)** | Amacén Tipo1='CADENA', Estado en {'ALMA ENCL','ENCL ALMA'}, Tipo Stock='02' | SÍ |
| **PDV (CAV/CAC)** | Amacén Tipo1 en {'CAV','CAC','OTROS CADS'}, Almacén en {'H000','H001'}, Estado='ALMA', Tipo Stock='01' | SÍ |
| **Televentas VES** | Centro='P150', Almacén='H000', Estado='ALMA', Tipo Stock en {'01','07'} | SÍ |
| **Tienda Virtual VES** | Centro='P102', Almacén='H000', Estado='ALMA', Tipo Stock en {'01','07'} | SÍ |
| **Corporativo Ves** | Centro='P101', Almacén en {'H001','H002'}, Estado='ALMA', Tipo Stock en {'01','07'} | SÍ |

### Contenido de la hoja "TD"

- **Tabla dinámica**: Material x Mes/Año x Almacén Agrupado → Cuenta de Series
- **Columna "NO DESPACHÓ"**: series de meses anteriores que no se despacharon antes que meses posteriores
- **Columna "NO CUMPLIÓ FIFO"**: violaciones FIFO (se despachó inventario más nuevo antes que más viejo)
- **Resumen**: Total Despacho, Total No Despachó, % Cumplimiento FIFO

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
