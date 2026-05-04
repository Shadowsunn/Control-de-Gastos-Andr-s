# 💰 Control de Gastos

Aplicación de escritorio para registrar, visualizar y exportar gastos personales y de negocio. Construida en Python con interfaz gráfica Tkinter y base de datos SQLite local.

---

## 📋 Tabla de contenidos

- [Descripción general](#descripción-general)
- [Características](#características)
- [Requisitos](#requisitos)
- [Instalación](#instalación)
- [Estructura del proyecto](#estructura-del-proyecto)
- [Uso](#uso)
- [Arquitectura técnica](#arquitectura-técnica)
- [Base de datos](#base-de-datos)
- [API del backend](#api-del-backend)
- [Validaciones](#validaciones)
- [Exportación](#exportación)
- [Limitaciones conocidas](#limitaciones-conocidas)
- [Roadmap](#roadmap)

---

## Descripción general

Control de Gastos es una herramienta de escritorio pensada para llevar el registro diario de gastos, tanto personales como del negocio. Permite registrar cada gasto con descripción, categoría, monto y tipo, consultarlos por día o por mes, ver un resumen por categoría con gráfico visual, y exportar el historial a Excel o CSV.

Toda la información se guarda localmente en un archivo SQLite — sin necesidad de internet ni servidores externos.

---

## Características

- **Registro rápido de gastos** con fecha automática (hoy)
- **Vista diaria** con total acumulado del día y conteo de registros
- **Vista mensual** con desglose Personal vs Negocio
- **Vista por categoría** con gráfico de barras proporcional
- **Edición y eliminación** de gastos registrados
- **Exportación a Excel (.xlsx)** con encabezados estilizados
- **Exportación a CSV** con codificación UTF-8 para compatibilidad con Excel en español
- **Categorías predefinidas** seleccionables desde un menú desplegable
- **Validación de montos** en formato COP (punto como miles, coma como decimal) y formato inglés
- **Base de datos local** SQLite — sin dependencias de red

---

## Requisitos

| Dependencia | Versión mínima | Obligatoria |
|-------------|---------------|-------------|
| Python | 3.8+ | ✅ Sí |
| tkinter | Incluido con Python | ✅ Sí |
| sqlite3 | Incluido con Python | ✅ Sí |
| openpyxl | 3.0+ | ⚠️ Solo para exportar a Excel |

> **Nota:** Si `openpyxl` no está instalado, la exportación a CSV sigue funcionando normalmente. Solo el formato Excel queda deshabilitado con un mensaje de error claro al intentar usarlo.

---

## Instalación

### 1. Clonar el repositorio

```bash
git clone https://github.com/tu-usuario/control-gastos-andres.git
cd control-gastos-andres
```

### 2. (Opcional) Crear entorno virtual

```bash
python -m venv venv

# Windows
venv\Scripts\activate

# macOS / Linux
source venv/bin/activate
```

### 3. Instalar dependencias opcionales

```bash
pip install openpyxl
```

### 4. Ejecutar la aplicación

```bash
python main.py
```

La base de datos `gastos_andres.db` se crea automáticamente en la misma carpeta al primer arranque.

---

## Estructura del proyecto

```
control-gastos-andres/
│
├── main.py              # Interfaz gráfica (Tkinter) — Frontend
├── backend.py           # Lógica de negocio y acceso a datos
├── gastos_andres.db     # Base de datos SQLite (se crea automáticamente)
└── README.md
```

### Responsabilidades por archivo

**`backend.py`** — Todo lo que no se ve en pantalla:
- Conexión y consultas a SQLite
- Validación de datos de entrada
- Conversión y parsing de montos
- Lógica de negocio (tipos válidos, rangos, límites)

**`main.py`** — Todo lo que el usuario ve y toca:
- Ventana principal y tabs (Hoy / Mensual / Categorías)
- Formularios de registro y edición
- Tablas de datos con scroll
- Gráfico de barras
- Llamadas al backend con manejo de errores

---

## Uso

### Tab 📅 Hoy

Registra un gasto nuevo completando los 4 campos:

| Campo | Descripción |
|-------|-------------|
| **Monto ($)** | Valor en pesos. Acepta formato COP (`1.200.000`) o inglés (`1200.50`) |
| **Categoría** | Selecciona del menú o escribe una categoría nueva |
| **Descripción** | Texto libre de hasta 200 caracteres |
| **Tipo** | Personal o Negocio |

La tabla muestra todos los gastos del día actual con el total acumulado y conteo de registros.

**Editar un gasto:** Selecciona una fila y haz clic en ✏️ Editar. Se abre una ventana modal con todos los campos editables, incluyendo la fecha.

**Eliminar un gasto:** Selecciona una fila y haz clic en 🗑️ Eliminar. Se pide confirmación antes de borrar.

> ⚠️ Si editas un gasto y cambias su fecha a un día distinto a hoy, el gasto desaparece de la vista de Hoy (comportamiento esperado — sigue existiendo en la vista Mensual).

### Tab 📊 Mensual

Selecciona mes y año y presiona 🔍 Buscar para ver todos los gastos del período. Muestra:

- Total general del mes
- Subtotal Personal
- Subtotal Negocio

Desde aquí puedes exportar el mes completo a Excel o CSV con el botón 📥 Exportar.

### Tab 🏷️ Categorías

Selecciona mes y año y presiona 🔍 Buscar para ver el gasto total por categoría, ordenado de mayor a menor. Incluye un gráfico de barras que muestra la proporción Personal vs Negocio del mes.

---

## Arquitectura técnica

```
┌─────────────────────────────────────┐
│              main.py                │
│          (Interfaz Tkinter)         │
│                                     │
│  Tab Hoy │ Tab Mensual │ Tab Cat.  │
└──────────────────┬──────────────────┘
                   │ llama funciones
                   ▼
┌─────────────────────────────────────┐
│             backend.py              │
│         (Lógica de negocio)         │
│                                     │
│  _parsear_monto()                   │
│  inicializar_db()                   │
│  agregar_gasto()                    │
│  obtener_gastos_dia()               │
│  obtener_gastos_mes()               │
│  obtener_por_categoria()            │
│  editar_gasto()                     │
│  eliminar_gasto()                   │
└──────────────────┬──────────────────┘
                   │ SQL parametrizado
                   ▼
┌─────────────────────────────────────┐
│          gastos_andres.db           │
│             (SQLite)                │
└─────────────────────────────────────┘
```

### Contrato de retorno del backend

Todas las funciones del backend retornan una tupla `(bool, payload)`:

```python
# Éxito con datos
(True, {"data": [...], "count": N})

# Éxito con mensaje
(True, "Gasto registrado correctamente.")

# Error
(False, "Mensaje descriptivo del error.")
```

El frontend siempre verifica el primer elemento antes de usar el segundo:

```python
exito, resultado = backend.agregar_gasto(...)
if exito:
    # usar resultado
else:
    messagebox.showerror("Error", resultado)
```

---

## Base de datos

### Esquema

```sql
CREATE TABLE IF NOT EXISTS gastos (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    fecha       TEXT    NOT NULL,               -- Formato YYYY-MM-DD
    descripcion TEXT    NOT NULL,
    categoria   TEXT    NOT NULL,
    monto       REAL    NOT NULL,
    tipo        TEXT    NOT NULL CHECK(tipo IN ('Personal', 'Negocio'))
);
```

### Ejemplo de registros

| id | fecha | descripcion | categoria | monto | tipo |
|----|-------|-------------|-----------|-------|------|
| 1 | 2026-04-01 | Mi almuerzo de hoy | Almuerzo | 17500.0 | Personal |
| 2 | 2026-04-01 | 3 llantas marca VVS2 | Compra de repuestos | 320000.0 | Negocio |

### Notas de diseño

- Las fechas se almacenan como `TEXT` en formato ISO 8601 (`YYYY-MM-DD`) para facilitar comparaciones y filtros `LIKE`.
- El campo `tipo` tiene un `CHECK` constraint a nivel de base de datos como segunda línea de defensa, además de la validación en Python.
- Los montos se guardan como `REAL` (float de doble precisión). Para importes en COP sin decimales esto no genera pérdida de precisión.

---

## API del backend

### `inicializar_db()`

Crea la tabla `gastos` si no existe. Seguro para llamar múltiples veces.

```python
exito, mensaje = backend.inicializar_db()
```

---

### `agregar_gasto(descripcion, categoria, monto, tipo)`

Inserta un nuevo gasto con la fecha de hoy.

```python
exito, mensaje = backend.agregar_gasto(
    descripcion="Almuerzo ejecutivo",
    categoria="Almuerzo",
    monto="17.500",       # Acepta string en formato COP o inglés
    tipo="Personal"
)
```

---

### `obtener_gastos_dia(fecha=None)`

Retorna todos los gastos de un día. Si `fecha` es `None`, usa hoy.

```python
exito, resultado = backend.obtener_gastos_dia("2026-04-01")
# resultado = {"data": [(id, fecha, desc, cat, monto, tipo), ...], "count": N}
```

---

### `obtener_gastos_mes(mes, año)`

Retorna todos los gastos de un mes y año.

```python
exito, resultado = backend.obtener_gastos_mes(4, 2026)
# resultado = {"data": [...], "count": N}
```

---

### `obtener_por_categoria(mes, año)`

Retorna el total gastado por categoría en un mes, ordenado de mayor a menor.

```python
exito, resultado = backend.obtener_por_categoria(4, 2026)
# resultado = {"data": [("Almuerzo", 52500.0), ("Transporte", 35000.0)], "count": N}
```

---

### `editar_gasto(id_gasto, fecha, categoria, descripcion, monto, tipo)`

Actualiza todos los campos de un gasto existente.

```python
exito, mensaje = backend.editar_gasto(
    id_gasto=1,
    fecha="2026-04-01",
    categoria="Almuerzo",
    descripcion="Almuerzo corregido",
    monto="18000",
    tipo="Personal"
)
```

---

### `eliminar_gasto(id_gasto)`

Elimina un gasto por su ID.

```python
exito, mensaje = backend.eliminar_gasto(3)
```

---

## Validaciones

### Montos

El backend acepta montos en múltiples formatos gracias a `_parsear_monto()`:

| Entrada | Resultado | Formato |
|---------|-----------|---------|
| `"1.200.000"` | `1200000.0` | COP — miles con punto |
| `"1.200.000,50"` | `1200000.5` | COP — miles + decimal |
| `"1200,50"` | `1200.5` | COP — decimal con coma |
| `"1200.50"` | `1200.5` | Inglés — decimal con punto |
| `"1200"` | `1200.0` | Entero simple |

La regla general: **el separador que aparece de último es el decimal**.

Montos inválidos que se rechazan: vacío, texto no numérico, negativos, cero, y valores superiores a `$999.999.999`.

### Descripción y Categoría

- Máximo 200 caracteres cada una
- No pueden estar vacías ni ser solo espacios

### Tipo

- Solo acepta `"Personal"` o `"Negocio"` (insensible a mayúsculas — `"personal"` se acepta y normaliza)

### Fechas

- Formato obligatorio: `YYYY-MM-DD`
- Validadas con `date.fromisoformat()` — rechaza fechas imposibles como `"2026-13-45"`

### IDs

- Deben ser enteros positivos — rechaza negativos, cero, booleanos y strings

---

## Exportación

### Excel (.xlsx)

Requiere `openpyxl`. El archivo generado incluye:

- Encabezados con fondo azul (`#4A90D9`) y texto blanco en negrita
- Columnas con ancho predefinido optimizado para lectura
- Montos como valores numéricos reales (no texto formateado)

Si `openpyxl` no está instalado, la app muestra un mensaje con las instrucciones de instalación y no genera el archivo.

### CSV

No requiere dependencias adicionales. Usa codificación `UTF-8 con BOM` (`utf-8-sig`) para garantizar compatibilidad con Microsoft Excel en español (evita que las tildes y la `ñ` aparezcan mal).

Columnas exportadas: `Fecha`, `Descripción`, `Categoría`, `Monto`, `Tipo`.

---

## Limitaciones conocidas

**Edición de montos con decimales:** Al abrir la ventana de edición, el monto se pre-carga eliminando todos los separadores de formato (`$`, `.`, `,`). Si el monto original tenía decimales (ej: `$129,50`), el campo mostrará `12950` en lugar de `129.50`. Para corrección, reescribir el monto manualmente en el campo de edición.

**Vista de Hoy solo muestra la fecha actual:** Si se edita un gasto y se cambia su fecha a otro día, desaparece de la vista de Hoy sin aviso. El gasto sigue existiendo y es visible en la vista Mensual.

**Sin soporte multiusuario:** La base de datos es un archivo local. No está diseñada para acceso concurrente desde múltiples usuarios o dispositivos.

---

## Roadmap

Funcionalidades pendientes identificadas durante el desarrollo:

- [ ] Exportar resumen mensual con totales por tipo (Personal / Negocio)
- [ ] Gráfico de torta con porcentaje de gastos por categoría
- [ ] Categorías más usadas con sugerencia automática al tipear
- [ ] Filtro por tipo (Personal / Negocio) en las vistas Mensual y Categorías
- [ ] Búsqueda por descripción o categoría
- [ ] Soporte para múltiples monedas

---

## Licencia

Uso personal. Proyecto privado de Erick Cortés para el cliente Andrés (Ente ficticio creado para simular una petición formal de App).
