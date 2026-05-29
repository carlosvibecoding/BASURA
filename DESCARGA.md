# Descargar plantilla qPCR (actualizado)

Repositorio: https://github.com/carlosvibecoding/BASURA

## Archivos (rama `main`)

| Archivo | Enlace directo |
|---------|----------------|
| **Plantilla Excel** | https://github.com/carlosvibecoding/BASURA/raw/main/qPCR_plantilla.xlsx |
| **Macro VBA v3.1** | https://github.com/carlosvibecoding/BASURA/raw/main/qpcr/Modulo_qPCR.bas |
| **Todo el proyecto (ZIP)** | https://github.com/carlosvibecoding/BASURA/archive/refs/heads/main.zip |
| **Procesar sin macro** | https://github.com/carlosvibecoding/BASURA/raw/main/qpcr/procesar_placa.py |

## Comprobar que tienes la macro nueva

Al ejecutar **ProcesarPlaca**, si funciona debe decir:

`Analisis completado: RGS12 (macro 3.1)`

Si el mensaje **no** incluye **(macro 3.1)**, sigues con la macro antigua.

## Importar la macro (resumen)

1. Abre `qPCR_plantilla.xlsx` → guardar como **`qPCR_plantilla.xlsm`**
2. **Programador** → **Visual Basic**
3. Elimina módulos viejos `Modulo_qPCR` si hay más de uno
4. **Archivo** → **Importar archivo** → `Modulo_qPCR.bas`
5. Guardar y cerrar el editor

## Pegar datos

En **RAW**, celda **A1**, pegar el export **completo** del StepOne:

- Bloque gen de interés (ej. RGS12)
- Bloque **PPIA**
- Bloque **SYP**

No basta con pegar solo RGS12.

## Sin macro (alternativa)

```bash
pip install openpyxl xlrd
python qpcr/procesar_placa.py "PLACA 2 RGS12 060526_data.xls" -o resultados.xlsx
```
