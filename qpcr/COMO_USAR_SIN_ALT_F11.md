# Usar la plantilla qPCR (sin Alt+F11)

## Importante

- Al abrir **`qPCR_plantilla.xlsx`** **no verás** `Modulo_qPCR`: la macro es un archivo **aparte** (`Modulo_qPCR.bas`) que hay que añadir al libro, o puedes **no usar macro** (opción B abajo).
- **Alt+F11** no funciona en Excel online ni en muchos Mac sin tecla Fn. Usa la pestaña **Programador**.

---

## Opción A — Macro con pestaña Programador (sin Alt+F11)

### 1. Mostrar la pestaña Programador

**Windows**

1. **Archivo** → **Opciones** → **Personalizar cinta de opciones**.
2. A la derecha, marcar **Programador** (a veces “Desarrollador”).
3. **Aceptar**.

**Mac**

1. **Excel** → **Preferencias** → **Cinta de opciones y barra de herramientas**.
2. Marcar **Programador** / **Desarrollador**.

### 2. Guardar como .xlsm

1. Abre la plantilla.
2. **Archivo** → **Guardar como**.
3. Tipo: **Libro habilitado para macros (.xlsm)**.
4. Guardar.

### 3. Abrir el editor VBA (sin Alt+F11)

1. Pestaña **Programador**.
2. Botón **Visual Basic** (primer bloque de la cinta).

Se abre una ventana aparte (Editor de VBA).

### 4. Importar o pegar la macro

**Importar**

1. En el editor: menú **Archivo** → **Importar archivo…**
2. Elegir **Todos los archivos**.
3. Seleccionar `Modulo_qPCR.bas`.

**Si no hay “Importar”**

1. Menú **Insertar** → **Módulo**.
2. Abrir `Modulo_qPCR.bas` con Bloc de notas.
3. Copiar desde `Option Explicit` hasta el final.
4. Pegar en la ventana del módulo.
5. En **Propiedades** del módulo (panel izquierdo abajo), nombre: `Modulo_qPCR`.

### 5. Guardar y ejecutar

1. En el editor VBA: **Archivo** → **Guardar** (o Ctrl+S).
2. Cerrar el editor.
3. En Excel, hoja **RAW**, pegar el export del termociclador en **A1**.
4. Pestaña **Programador** → **Macros** → elegir **`ProcesarPlaca`** → **Ejecutar**.

(También: **Vista** → **Macros** → **Ver macros** en algunas versiones.)

---

## Opción B — Sin macro (recomendado si VBA no funciona)

Solo necesitas **Python** instalado (una vez).

### 1. Instalar (una vez)

Abrir **Símbolo del sistema** o **Terminal**:

```bash
pip install openpyxl xlrd
```

### 2. Descargar el repo o la carpeta `qpcr`

Debe contener `procesar_placa.py` y `processor.py`.

### 3. Procesar

**Con el export del termociclador (.xls):**

```bash
cd ruta\a\la\carpeta\qpcr
python procesar_placa.py "PLACA 2 RGS12 060526_data.xls" -o resultados.xlsx
```

**Con datos pegados en la plantilla (hoja RAW):**

1. Abre `qPCR_plantilla.xlsx`, pega el export en **RAW**, guarda.
2. Ejecuta:

```bash
python procesar_placa.py "qPCR_plantilla.xlsx" -o resultados.xlsx
```

### 4. Abrir `resultados.xlsx`

Ahí están las hojas **Resultados** y **GLOBAL**, sin configurar macros.

---

## Comprobar que la macro está dentro del libro

Solo después de importar/pegar:

1. **Programador** → **Visual Basic**.
2. Panel izquierdo: **Módulos** → **Modulo_qPCR**.

Si no aparece **Módulos**, la macro no está guardada en ese `.xlsm`.

---

## Enlaces

- Plantilla: `qPCR_plantilla.xlsx`
- Macro: `qpcr/Modulo_qPCR.bas`
- Rama: https://github.com/carlosvibecoding/BASURA/tree/cursor/qpcr-excel-template-c0e4
