# Descargar plantilla qPCR (actualizado)

Repositorio: https://github.com/carlosvibecoding/BASURA

## Archivos (rama `main`)

| Archivo | Enlace directo |
|---------|----------------|
| **Plantilla Excel** | https://github.com/carlosvibecoding/BASURA/raw/main/qPCR_plantilla.xlsx |
| **Macro VBA v4.3** | https://github.com/carlosvibecoding/BASURA/raw/main/qpcr/Modulo_qPCR.bas |
| **Todo el proyecto (ZIP)** | https://github.com/carlosvibecoding/BASURA/archive/refs/heads/main.zip |
| **Procesar sin macro** | https://github.com/carlosvibecoding/BASURA/raw/main/qpcr/procesar_placa.py |

## Comprobar que tienes la macro nueva

Al ejecutar **ProcesarPlaca**, si funciona debe decir:

`Analisis completado: RGS12 (macro 4.3)`

Si el mensaje **no** incluye **(macro 4.3)**, importa de nuevo `Modulo_qPCR.bas` y vuelve a abrir el libro.

## Importar la macro (resumen)

1. Abre `qPCR_plantilla.xlsx` → guardar como **`qPCR_plantilla.xlsm`**
2. **Programador** → **Visual Basic**
3. Elimina módulos viejos `Modulo_qPCR` si hay más de uno
4. **Archivo** → **Importar archivo** → `Modulo_qPCR.bas`
5. Guardar y cerrar el editor

## Grupos de muestras

En **Instrucciones**:

| Celda | Uso |
|-------|-----|
| **B21** | Prefijos que son **control** para el promedio ΔCt (por defecto `C`). Varios: `C,CTRL` |
| **B22** | Nombres para GLOBAL (opcional): `C=Controles;S=Suicidas;A=Alcohólicos` |

Muestras válidas: **letras + número** (`C10`, `S5`, `A12`, `ALC3`…). GLOBAL crea una tabla por cada prefijo distinto (no control).

## Interfaz laboratorio (macro 4.3)

- **Franja superior A1:N** con banner de **doble hélice ADN** (imagen nítida, fina, no tapa el pegado).
- **Panel derecho (columna O)**: título `qPCR · lab`, **ratón de laboratorio sobre placa-botón** (clic = procesar), **Limpiar** (rojo) y **+ Placa** (teal).
- Decoración molecular sutil (esquina ADN) también en **Instrucciones, Resultados, GLOBAL, Datos y Calculos**.
- La decoración ahora va en **PNG incrustado** (hoja oculta `Recursos`), no en cientos de formas VBA → más limpia y rápida.
- **No se escribe en A3** si ya hay datos pegados; el export va desde **A1**.
- Detección del gen de interés en **tabla única** (gen + PPIA + SYP juntos) — sin cambios respecto a 4.2.1.

**Obligatorio:** plantilla nueva + macro 4.3 + guardar como `.xlsm` + cerrar y abrir Excel.

| Control | Acción |
|---------|--------|
| **Ratón** | Procesar placa |
| **Limpiar** | Borra datos |
| **+ Placa** | Marca donde pegar otra placa |

## Pegar datos

En **RAW**, celda **A1**, pegar el export **completo** del StepOne:

- Gen de interés (ej. RGS10 / RGS12)
- **PPIA**
- **SYP**

(en una sola tabla o en bloques apilados)

Varias placas: pegar una debajo de otra (o usar **+ Placa**).

## Hojas intermedias

- **Datos** — muestra × gen con Ct numéricos.
- **Calculos** — valores VBA (PPIA y SYP en columnas separadas).
- **Resultados** — bloque izquierdo PPIA, bloque derecho SYP (ΔCt / ΔΔCt / FC).

## Resultados

- Fila **2**: promedio ΔCt de controles (C) en **Prom. dCt (C)** (PPIA col G, SYP col W).
- Fila **3+**: muestras.
- **Rojo**: indeterminado, Ct SD > 0,3 o FC extremo.
- **Naranja**: un solo duplicado Ct válido.
- **GLOBAL**: tablas por grupo.

## Sin macro (alternativa)

```bash
pip install openpyxl xlrd pillow
python qpcr/procesar_placa.py "PLACA 2 RGS12 060526_data.xls" -o resultados.xlsx
```
