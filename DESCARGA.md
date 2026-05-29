# Descargar plantilla qPCR (actualizado)

Repositorio: https://github.com/carlosvibecoding/BASURA

## Archivos (rama `main`)

| Archivo | Enlace directo |
|---------|----------------|
| **Plantilla Excel** | https://github.com/carlosvibecoding/BASURA/raw/main/qPCR_plantilla.xlsx |
| **Macro VBA v3.7** | https://github.com/carlosvibecoding/BASURA/raw/main/qpcr/Modulo_qPCR.bas |
| **Todo el proyecto (ZIP)** | https://github.com/carlosvibecoding/BASURA/archive/refs/heads/main.zip |
| **Procesar sin macro** | https://github.com/carlosvibecoding/BASURA/raw/main/qpcr/procesar_placa.py |

## Comprobar que tienes la macro nueva

Al ejecutar **ProcesarPlaca**, si funciona debe decir:

`Analisis completado: RGS12 (macro 3.7)`

Tras importar la macro, abre el libro de nuevo (o ejecuta **InstalarBotones**) para ver los botones en RAW.

Si el mensaje **no** incluye **(macro 3.2)**, importa de nuevo `Modulo_qPCR.bas`.

## Importar la macro (resumen)

1. Abre `qPCR_plantilla.xlsx` → guardar como **`qPCR_plantilla.xlsm`**
2. **Programador** → **Visual Basic**
3. Elimina módulos viejos `Modulo_qPCR` si hay más de uno
4. **Archivo** → **Importar archivo** → `Modulo_qPCR.bas`
5. Guardar y cerrar el editor

## Botones en RAW (macro 3.2)

| Botón | Acción |
|-------|--------|
| **Procesar placa** | Calcula Resultados y GLOBAL |
| **Limpiar datos** | Borra RAW, Datos, Calculos, Resultados y GLOBAL |
| **Añadir placa abajo** | Marca dónde pegar otra placa; luego **Procesar** (lee todas) |

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
