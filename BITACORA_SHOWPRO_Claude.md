# BITÁCORA DE GENERADORES SHOWPRO — Documento Maestro para Claude

> Pega este archivo como conocimiento / contexto en Claude para que gestione, valide y genere registros compatibles con el Excel `Bitacora_GENERADORES_SHOWPRO.xlsx` (compatible Excel + Google Sheets).

## 1. Contexto
Empresa: SHOWPRO (eventos / faenas).
Equipos iniciales:
- GEN-01: Generador 130 kVA (horómetro base ~1242 hrs)
- GEN-02: Generador 60 kVA (horómetro base ~832 hrs)
- GEN-03 en adelante: agregable por el usuario en Hoja 1.

Intervalo de servicio por defecto: **250 hrs** (configurable por equipo, admite 500).
Archivo Excel con 4 hojas en este orden:
1. `1_MAESTRA_Equipos` — MAESTRA, única entrada manual de equipos.
2. `Bitácora General` — registro diario / por evento y jornada.
3. `Resumen de Equipos` — vista automática, no se edita a mano.
4. `Listas y ayuda` — referencia.

Regla de oro: **1 fila de bitácora = 1 Fecha + 1 Equipo + 1 Jornada.** Si un equipo trabaja día y noche el mismo día, son 2 filas.

## 2. Hoja 1 — `1_MAESTRA_Equipos` (la maestra)
Rango de equipos: filas 5 a 14 (hasta 10 generadores). Fila 7 amarilla = slot del 3er generador.

| Col | Campo | Formato / Regla |
|-----|-------|----------------|
| A | ID | `GEN-01`, `GEN-02`... correlativo, no repetir |
| B | Nombre Equipo * | **Clave única.** Ej: `Generador 130 kVA`. Sin tildes preferente, no duplicar. Este valor alimenta desplegables de Hoja 2 y columna A de Hoja 3 |
| C | Potencia / Capacidad | Texto, ej: `130 kVA` |
| D | N° Serie / Patente | Texto libre |
| E | Horómetro inicial (hrs) | Número entero `#,##0`. Base cuando no hay movimientos en bitácora |
| F | Intervalo servicio (hrs) | Número, por defecto `250`. Admite `500` por equipo |
| G | Fecha ingreso | `dd-mmm-yyyy` |
| H | Estado | Lista: `Operativo, En taller, Fuera de servicio, Standby` |
| I | Ubicación / Base | Texto, ej: `Base ShowPro` |
| J | Observaciones | Texto libre |

### Cómo agregar el 3er generador
1. Completar fila 7 (GEN-03) en Maestra: B, C, D, E, F, G, H.
2. No borrar filas, solo completar.
3. Automáticamente aparece en Hoja 2 columna C y Hoja 3 columna A. No tocar fórmulas.

## 3. Hoja 2 — `Bitácora General`
Rango de datos: filas 5 a 504 (500 registros). Fila 3 = encabezados, fila 4 = grupos, fila 5+ = datos. Filtros en `A3:N504`. Panel congelado en `A5`.

| Col | Campo | Tipo / Validación / Fórmula |
|-----|-------|-----------------------------|
| A | Fecha | Fecha real de la jornada. Formato `dd-mmm-yyyy` |
| B | Evento / Faena / Cliente * | Texto obligatorio. Ej: `Festival ShowPro Día 1`, `Matrimonio Viña`, `Taller – mantención` |
| C | Equipo | Desplegable desde `='1_MAESTRA_Equipos'!$B$5:$B$14`. Rechaza valores fuera de Maestra |
| D | Jornada / Turno | Lista: `Día (08:00-20:00), Noche (20:00-02:00), Día completo 24H, Mañana/Tarde parcial, Standby / Respaldo, Traslado / Taller` |
| E | Horómetro Inicial (hrs) | Número `#,##0`. Foto al inicio |
| F | Horómetro Final (hrs) | Número `#,##0`. Foto al final. Alerta roja si `F<E` |
| G | Horas del día | **Automática, no escribir.** `=IF(AND(E5<>"",F5<>""),F5-E5,"")`. Formato `#,##0.0`. Alerta si `>24` |
| H | Tipo Mantención | Lista: `Sin intervención, Inspección Rutinaria, Preventiva, Correctiva, Emergencia` |
| I | Nivel Aceite | Lista: `OK, Revisado, Rellenado, Cambiados, Limpiado, Requiere Cambio, Crítico` |
| J | Filtros Aire/Combust. | Misma lista que I |
| K | Voltaje Batería (V) | Número `0.0`, ej: `13.8` |
| L | Estado Batería | Lista: `OK, Revisado, Requiere Cambio, Crítico` |
| M | Técnico / Responsable | Texto libre |
| N | Observaciones / Trabajos | Texto libre |

Filas ejemplo:
- `2026-10-26 | Evento Matrimonio Viña – Escenario Principal | Generador 130 kVA | Día (08:00-20:00) | 1242 | 1250 | 8 | Preventiva | OK | Cambiados | 13.8 | OK | Juan Pérez | Cambio filtro aceite y combustible. Prueba con carga OK.`
- `2026-10-26 | Evento Matrimonio Viña – Iluminación | Generador 60 kVA | Día (08:00-20:00) | 832 | 840 | 8 | Inspección Rutinaria | Rellenado | Revisado | 12.6 | OK | Carlos Soto | Control previo a evento. Sin novedades.`

Totales (filas 507-510): `=SUMIF(C5:C504,'1_MAESTRA_Equipos'!B5,G5:G504)` por cada equipo (B5, B6, B7). No incluir en promedios.

## 4. Hoja 3 — `Resumen de Equipos` (solo lectura)
10 filas (4 a 13) ligadas 1:1 a Maestra filas 5 a 14. Si `A=""` la fila queda en blanco.

| Col | Campo | Fórmula (Excel EN, Sheets la convierte) |
|-----|-------|------------------------------------------|
| A | Equipo | `=IF('1_MAESTRA_Equipos'!B5="","",'1_MAESTRA_Equipos'!B5)` |
| B | Intervalo (hrs) | `=IF(A4="","",IF('1_MAESTRA_Equipos'!F5<>"",'1_MAESTRA_Equipos'!F5,250))` |
| C | Última fecha | `=IF(A4="","",IF(MAX(INDEX(('Bitácora General'!C5:C504=A4)*'Bitácora General'!A5:A504,0))=0,"Sin movimientos",MAX(...)))` |
| D | Último horómetro | `=IF(A4="","",IF(MAX(INDEX((C=A)*F,0))=0,'1_MAESTRA_Equipos'!E5,MAX(...)))` |
| E | Horas últimos 30 días | `=IF(A4="","",SUMIFS(G5:G504,C5:C504,A4,A5:A504,">="&TODAY()-30))` |
| F | Última Preventiva | Igual que D pero con condición `*(H="Preventiva")`. Si 0 → `"Sin registro"` |
| G | Horas desde servicio | `=IF(A4="","",IF(C4="Sin movimientos",0,IF(ISNUMBER(F4),D4-F4,D4-Maestra!E5)))` |
| H | Horas restantes | `=IF(A4="","",B4-G4)` con semáforo: `>50` verde, `1-50` amarillo, `<=0` rojo |

## 5. Compatibilidad Excel + Google Sheets
Permitido: `IF, AND, OR, ISNUMBER, IFERROR, SUM, SUMIF, SUMIFS, MAX, INDEX, TODAY, MATCH`.
Prohibido: `MAXIFS, MINIFS, XLOOKUP, XMATCH, FILTER, SORT, UNIQUE, LET, LAMBDA, SEQUENCE, TEXTJOIN, AGGREGATE`, tablas estructuradas, macros VBA, scripts.
Archivo: `.xlsx` sin macros. Validaciones por rango fijo (`$B$5:$B$14`), no rangos dinámicos con `OFFSET`. Formatos número/fecha estándar. Filtros nativos. Import a Sheets: Subir > Abrir > Guardar como hoja de Google.

## 6. Instrucciones para Claude
1. **Agregar equipo:** pide B, C, D, E, F, G, H. Valida B único y E/F numéricos >0. Propón el siguiente `GEN-NN` libre (5-14).
2. **Agregar registro bitácora:** pide siempre A, B, C, D, E, F. Valida C existe en Maestra, E≤F, G=F-E ≤24, B no vacío. Si falta dato, pregunta antes de inventar. Tipo H por defecto `Sin intervención` si solo operó.
3. **Responder con tabla:** usa exactamente las 14 columnas A-N en orden, más `Horas del día = F-E`.
4. **Alertas:** si `F<E` → error; si `Horas restantes <=50` → avisa "programar servicio"; si `<=0` → "servicio vencido"; si `I/J/L = Requiere Cambio / Crítico` → destaca.
5. **Nunca:** inventar horómetros, cambiar Maestra sin confirmación, usar funciones incompatibles, o editar Hoja 3 a mano.
6. **Idioma:** español (Chile). Fechas `dd-mmm-yyyy`. Horas con 1 decimal, horómetros enteros.

## 7. Ratificación v2.1 (errores corregidos el 2026-10-26)
- [x] Jornada `Mañana/Tarde parcial` sin backslash (antes `Ma\ñana`).
- [x] Resumen A7 y A12 restauradas (estaban vacías por borrado accidental, ahora `=IF(Maestra!B8…)` y `=IF(Maestra!B13…)`).
- [x] Code.gs `getEquiposMaestra` reescrito fila por fila (antes mezclaba intervalos si había huecos).
- [x] Code.gs pinta semáforo por `x.fila` real (antes por índice filtrado, fallaba si se saltaba una fila de Maestra).
- [x] Code.gs `onEdit` con guarda `if(!e||!e.range)` y `enviarWhatsApp` detecta placeholder `PEGUE_AQUI`.
- [x] Fórmulas solo con `IF, MAX, INDEX, SUMIFS, TODAY, ISNUMBER` (sin MAXIFS/XLOOKUP/FILTER). Verificado en Excel y Sheets.

## 8. Prompt rápido para copiar en Claude
```
Usa el documento BITÁCORA SHOWPRO como esquema.
Maestra actual: GEN-01 Generador 130 kVA (1242 hrs, 250 hrs), GEN-02 Generador 60 kVA (832 hrs, 250 hrs).
Si te pido agregar un registro, valídalo contra la Maestra y devuélvelo en formato tabla A-N con Horas del día = Final-Inicial.
Si te pido agregar el 3er generador, pídeme nombre, potencia, serie, horómetro inicial, intervalo y estado, y genérame la fila Maestra lista para pegar.
No uses MAXIFS/XLOOKUP/FILTER, solo fórmulas compatibles Excel+Sheets.
```
