# Descargar plantilla qPCR (actualizado)

Repositorio: https://github.com/carlosvibecoding/BASURA

## Archivos (rama `main`)

| Archivo | Enlace directo |
|---------|----------------|
| **Plantilla Excel** | https://github.com/carlosvibecoding/BASURA/raw/main/qPCR_plantilla.xlsx |
| **Macro VBA v4.1** | https://github.com/carlosvibecoding/BASURA/raw/main/qpcr/Modulo_qPCR.bas |
| **Plantilla (raton + fondo ADN)** | https://github.com/carlosvibecoding/BASURA/raw/main/qPCR_plantilla.xlsx |
| **Todo el proyecto (ZIP)** | https://github.com/carlosvibecoding/BASURA/archive/refs/heads/main.zip |
| **Procesar sin macro** | https://github.com/carlosvibecoding/BASURA/raw/main/qpcr/procesar_placa.py |

## Comprobar que tienes la macro nueva

Al ejecutar **ProcesarPlaca**, si funciona debe decir:

`Analisis completado: RGS12 (macro 4.0)`

Tras importar la macro, abre el libro de nuevo (o ejecuta **InstalarBotones**) para ver los botones en RAW.

Si el mensaje **no** incluye **(macro 3.2)**, importa de nuevo `Modulo_qPCR.bas`.

## Importar la macro (resumen)

1. Abre `qPCR_plantilla.xlsx` → guardar como **`qPCR_plantilla.xlsm`**
2. **Programador** → **Visual Basic**
3. Elimina módulos viejos `Modulo_qPCR` si hay más de uno
4. **Archivo** → **Importar archivo** → `Modulo_qPCR.bas`
5. Guardar y cerrar el editor

## Grupos de muestras (macro 3.8)

En **Instrucciones**:

| Celda | Uso |
|-------|-----|
| **B21** | Prefijos que son **control** para el promedio ΔCt (por defecto `C`). Varios: `C,CTRL` |
| **B22** | Nombres para GLOBAL (opcional): `C=Controles;S=Suicidas;A=Alcohólicos` |

Muestras válidas: **letras + número** (`C10`, `S5`, `A12`, `ALC3`…). GLOBAL crea una tabla por cada prefijo distinto (no control).

## Interfaz laboratorio (macro 4.1)

- **Raton de laboratorio** (imagen fija, estilo mono): **clic = Procesar placa**.
- **Fondo** con doble helice / detalle molecular (sutil, no invasivo).
- Colores tipo laboratorio (verde azulado agua, fondo claro).
- Todo va en `qPCR_plantilla.xlsx` (hoja oculta **Recursos**); no hay que elegir imagenes.

**Obligatorio:** plantilla nueva + macro 4.1 + guardar como `.xlsm` + cerrar y abrir Excel.

| Control | Accion |
|---------|--------|
| **Raton** | Procesar placa |
| **Limpiar** | Borra datos |
| **+ Placa** | Marca donde pegar otra placa |

## Pegar datos

En **RAW**, celda **A1**, pegar el export **completo** del StepOne:

- Bloque gen de interés (ej. RGS12)
- Bloque **PPIA**
- Bloque **SYP**

Varias placas: pegar una debajo de otra (o usar **Añadir placa abajo**).

## Hojas intermedias (macro 3.6)

- **Datos** — muestra × gen con Ct numéricos.
- **Calculos** — solo valores (sin fórmulas Excel); PPIA y SYP en columnas separadas.
- **Resultados** — bloque izquierdo PPIA, bloque derecho SYP (ambos con ΔCt / ΔΔCt / FC).

## Resultados

- Fila **2**: promedio ΔCt de controles (C) en **Prom. dCt (C)** (PPIA col G, SYP col W).
- Fila **3+**: muestras (sin repetir el promedio en cada fila).
- **Rojo**: indeterminado, Ct SD > 0,3 o valor 2^(-ΔΔCt) extremo.
- **Naranja**: solo un duplicado Ct válido (se usa ese valor / media del instrumento).
- **Indeterminado**: ambos Ct "Undetermined" → sin cálculo ΔΔCt.
- **GLOBAL**: muestras ordenadas C2, C3… S1, S2…

## Sin macro (alternativa)

```bash
pip install openpyxl xlrd
python qpcr/procesar_placa.py "PLACA 2 RGS12 060526_data.xls" -o resultados.xlsx
```
