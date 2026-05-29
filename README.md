# Plantilla Excel qPCR (ΔΔCt)

Herramienta para analizar exportaciones del termociclador **StepOne** (PPIA y SYP como genes control), con el mismo flujo que las placas de referencia `PLACA 1 RGS10` / `PLACA 2 RGS12`.

## Archivos

| Archivo | Descripción |
|---------|-------------|
| `qPCR_plantilla.xlsx` | Plantilla vacía (RAW + hojas Resultados / GLOBAL) |
| `qpcr/Modulo_qPCR.bas` | Macro VBA `ProcesarPlaca` |
| `qpcr/crear_plantilla_qpcr.py` | Regenerar la plantilla o generar un ejemplo con `--demo` |
| `PLACA 1 RGS10 300426_data.xls` | Ejemplo procesado (referencia) |
| `PLACA 2 RGS12 060526_data.xls` | Ejemplo RAW del termociclador |

## Uso en Excel (primera vez)

1. Abra `qPCR_plantilla.xlsx`.
2. **Alt+F11** → Archivo → **Importar archivo** → seleccione `qpcr/Modulo_qPCR.bas`.
3. En la hoja **RAW**, inserte un botón (Insertar → Formas → asignar macro **ProcesarPlaca**), o ejecute la macro desde el editor VBA.
4. Pegue el export del termociclador en **RAW** desde **A1** (puede apilar varias exportaciones).
5. Pulse **Procesar placa**.

## Cálculos

Para cada control (**PPIA** y **SYP**):

1. **ΔCt** = Ct mean (gen de interés) − Ct mean (control)
2. **Promedio ΔCt** solo en muestras **C** (C1, C12, …)
3. **ΔΔCt** = ΔCt − ese promedio (fijo para toda la placa)
4. **2^(−ΔΔCt)** en la hoja **GLOBAL** (columnas PPIA, SYP y **MEDIA**)

Las muestras con **Ct SD > 0,3** en el gen de interés se marcan en **rojo** (siguen calculándose).

## Regenerar plantilla

```bash
pip install -r qpcr/requirements.txt
python qpcr/crear_plantilla_qpcr.py
python qpcr/crear_plantilla_qpcr.py --demo "PLACA 2 RGS12 060526_data.xls" -o qPCR_ejemplo_placa2.xlsx
```

## Notas

- El gen de interés se detecta solo (único target distinto de PPIA/SYP).
- Las cabeceras del StepOne usan a veces **Cт** con «т» cirílica; la macro y el script lo normalizan.
- Una placa por libro; en RAW puede pegar varios exports y reprocesar todo.
