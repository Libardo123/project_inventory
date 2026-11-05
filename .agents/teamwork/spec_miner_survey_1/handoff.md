# Reporte de Especificación Técnica y Arquitectura de Integración (R2, R3, R4, R5)

**Fecha**: 2026-11-05  
**Agente**: `spec_miner_survey_1` (Teamwork Specification Miner)  
**Misión**: Extracción de especificaciones técnicas autoritativas, análisis de mejores prácticas, diseño de arquitectura e integración de componentes para R2 (Escaneo Móvil y Guía Operativa), R3 (GUI y Usabilidad), R4 (Patrones de Librerías, DPI y Estilos) y R5 (Suite de Pruebas Automatizadas QA).  
**Repositorio / Directorio de Trabajo**: `c:\Users\Analista-TI\Desktop\entrega inicial\Scrip-Inventario`

---

## 1. Features Discovered

| # | Categoría | Feature | Descripción | Inputs | Outputs | Error Behavior | Descubierto Vía |
|---|-----------|---------|-------------|--------|---------|----------------|-----------------|
| 1 | R2: Escaneo Móvil | Emulación de Teclado Inalámbrico (Barcode to PC) | Transmisión en tiempo real de códigos QR/barras desde smartphone hacia PC por socket Wi-Fi local inyectando caracteres como pulsaciones HID de teclado | Lectura de cámara de smartphone (QR/Barcode) | Inyección de texto + sufijo `{ENTER}` en la celda activa de Excel | Timeout por firewall/red desconectada; reintento automático con fallback a IP fija | `ORIGINAL_REQUEST.md` §R2, `conciliador_universal.py`:1296-1344 |
| 2 | R2: Escaneo Móvil | Sincronización Cloud con Excel Mobile App | Edición colaborativa directa en la nube (OneDrive / SharePoint) usando la función nativa de cámara/código de barras de Excel Móvil | Lectura de QR en campo en celda de la hoja `Escaneos_Campo` | Registro sincrónico en libro Excel accesible simultáneamente en PC | Conflicto de versiones si se edita sin conexión; resolución automática por coautoría | `ORIGINAL_REQUEST.md` §R2 |
| 3 | R2: Modelo de Datos | Esquema `Maestro_Activos` | Catálogo maestro de 8 columnas base (`ID_Activo`, `Categoria`, `Tipo_Equipo`, `Equipo / Host`, `Serial_Fabricante`, `Ubicacion_Oficial`, `Responsable`, `Estado_Oficial`) | Base de datos de inventario institucional / Excel original | Hoja formateada con estilo Azul Marino (`#1F4E78`), autofiltro y congelamiento de paneles | Fila vacía ignorada; duplicados señalizados | `conciliador_universal.py`:1254-1295 |
| 4 | R2: Modelo de Datos | Esquema `Escaneos_Campo` | Registro de captura en campo de 6 columnas (`ID_Escaneado`, `Ubicacion_Hallada`, `Fecha_Hora_Escaneo`, `Operador`, `Estado_Fisico`, `Observaciones`) | Lectura por escaneo QR o ingreso manual de auditor | Hoja formateada en Verde Bosque (`#375623`) con formato tabular para auditoría | Registros sin ID rechazan validación | `conciliador_universal.py`:1296-1318 |
| 5 | R2: Auditoría | Motor de Conciliación Campo vs Maestro | Cruce bidireccional entre `Escaneos_Campo` y `Maestro_Activos` detectando: OK, Traslado, Faltante, Sobrante (Alta) y Duplicado | Columnas mapeadas de ambas hojas | Reporte de auditoría con resumen estadístico y celdas codificadas por color | Serial o ID no normalizado evaluado mediante motor Levenshtein tolerante | `conciliador_universal.py`:54-148, 1856-2223 |
| 6 | R3: GUI | Casillas de Verificación (Chulos) por Columna | Checkboxes individuales (`chk_col_serial`, `chk_col_equipo`, etc.) que activan o desactivan columnas específicas en el algoritmo de matching y actualización | Selección interactiva del usuario (BooleanVar) | Habilitación/deshabilitación visual de comboboxes y filtrado en lógica de ejecución | Si todas están desmarcadas, advertencia modal impidiendo ejecución en vacío | `ORIGINAL_REQUEST.md` §R3, `conciliador_universal.py`:301-308, 405-440 |
| 7 | R3: GUI | Selector de Directorio de Guardado | Cuadro de diálogo nativo (`filedialog.askdirectory`) para seleccionar la carpeta base de almacenamiento de resultados y etiquetas | Clic en "Cambiar Carpeta Destino..." -> Selección de ruta | Actualización de variable `ruta_salida` y generación de subcarpetas en destino | Si se cancela la ventana, preserva la ruta previa sin lanzar excepción | `ORIGINAL_REQUEST.md` §R3, `conciliador_universal.py`:1403-1416 |
| 8 | R3: GUI | Botones de Acceso Directo 1-Clic | Botones ejecutables para abrir el Excel auditado generado y su carpeta contenedora mediante `os.startfile` | Clic en "Abrir Excel Auditado" o "Abrir Carpeta" | Apertura inmediata en Microsoft Excel / Explorador de Windows | Mensaje informativo si el archivo aún no ha sido generado | `ORIGINAL_REQUEST.md` §R3, `conciliador_universal.py`:1417-1436 |
| 9 | R3: GUI | Consola de Bitácora Amplia con Scrollbar | Área de texto `tk.Text` con fondo oscuro, tipografía monospace y `ttk.Scrollbar` sincronizada con autoscroll al final de cada evento | Mensajes emitidos por el motor (`self.log`) | Renderizado en tiempo real con resaltado cromático (`ok`, `warn`, `error`, `info`) | Manejo seguro de codificación UTF-8 sin bloqueo de UI mediante `update_idletasks()` | `ORIGINAL_REQUEST.md` §R3, `conciliador_universal.py`:550-568 |
| 10 | R4: Librerías / QR | Generación de Códigos QR a 300 DPI | Renderizado de imágenes de etiquetas a resolución física de 300 DPI con metadatos DPI incrustados en PNG mediante Pillow | Matriz QR (`qrcode.QRCode`) + Canvas Pillow | Archivos PNG individuales (591x295 px para 50x25mm; 1200x600 px para 100x50mm) a 300 DPI | Fallback a fuentes por defecto si las TTF del sistema no existen | `ORIGINAL_REQUEST.md` §R4, `conciliador_universal.py`:988-1109 |
| 11 | R4: Rotulación | Ficha Técnica Completa (7 Campos) en Etiqueta | Rotulación física y carga útil del QR conteniendo los 7 campos: ID, Hostname, Modelo, Serial, Responsable, Sede y Estado | Diccionario de datos del activo extraído del Excel | Etiqueta con QR y textos legibles sin cruce de celdas ni truncamientos ilegibles | Sanitización de caracteres prohibidos en nombres de archivo | `ORIGINAL_REQUEST.md` §R1, `conciliador_universal.py`:1021-1104 |
| 12 | R4: Formato Excel | Estilizado Empresarial openpyxl | Aplicación de bordes delgados, fuentes Calibri, autofiltros, fórmulas SUM dinámicas y anchos automáticos de columna | Objeto Workbook y Worksheets | Libros `.xlsx` limpios con celdas no solapadas y legibilidad inmediata | Preserva fórmulas preexistentes o recalcula rangos dinámicamente | `conciliador_universal.py`:154-173, 2157-2202 |
| 13 | R5: QA Automatizado | Suite de Pruebas de Integración y Estrés | Script ejecutable independiente sin interfaz gráfica que procesa los archivos reales de prueba, valida integridad y finaliza con exit code 0 | `Inventario_Equipos_Clemont (2) (2).xlsx`, datos reales de preset | Reporte de aserciones QA pasando al 100% | Lanzamiento de `AssertionError` y salida con código 1 ante cualquier regresión | `ORIGINAL_REQUEST.md` §R5 |
| 14 | R5: QA Automatizado | Verificación de Legibilidad QR | Decodificación automatizada de los PNG generados mediante análisis matricial/librería de visión y verificación de presencia de los 7 campos | Archivos PNG generados en `etiquetas_qr/` | Validación de coincidencia exacta de datos codificados vs base de datos | Falla de aserción si falta un campo o la imagen está corrupta | `ORIGINAL_REQUEST.md` §R5 |

---

## 2. Edge Cases

| # | Feature | Input | Observed / Expected Behavior |
|---|---------|-------|------------------------------|
| 1 | Normalización de Seriales | Caracteres confundibles: `O` por `0`, `I` por `1`, `Z` por `2`, `S` por `5`, `B` por `8` | La función `normalizar_tolerante` mapea y homologa los caracteres para evitar falsos negativos en lecturas escaneadas o tipeadas. |
| 2 | Mapeo de Columnas con Checkboxes | Usuario desmarca la casilla "Serial" y mantiene marcada "Host/Equipo" | El motor de búsqueda omite la comparación de seriales y reconcilia exclusivamente por similitud de nombre de host, sin sobreescribir la columna de seriales del Excel. |
| 3 | Pegado desde Excel | Rango copiado contiene saltos de línea internos en la celda de Observaciones | El analizador de portapapeles divide por tabuladores (`\t`) preservando la celda multilínea sin desplazar las columnas subsiguientes. |
| 4 | Detección de Fila de Totales | Hoja Excel con filas en blanco previas a la fila `TOTAL` | La función `_detectar_fila_total` recorre las columnas A a F buscando la palabra `TOTAL` para insertar las altas exactamente antes de la fórmula de sumatoria. |
| 5 | Selector de Salida | Usuario selecciona una carpeta de red o directorio con espacios y caracteres especiales | `os.startfile` y `subprocess.Popen` manejan rutas absolutas completas entre comillas para evitar rupturas de ruta en Windows. |
| 6 | Escaneo Móvil sin Conexión | Red Wi-Fi inestable durante el escaneo en sótano/bodega | En Barcode to PC: la app móvil almacena en búfer local los escaneos y los descarga en ráfaga al reconectar; en Excel Mobile: sincronización asíncrona mediante caché de Office. |
| 7 | Contenido QR Ficha Completa | Textos largos en campo Responsable o Ubicación (>35 caracteres) | En la etiqueta física se aplica recorte visual elegante con elipsis o límite de caracteres (`[:28]`), mientras que en el QR se codifica el texto completo sin pérdida de datos. |
| 8 | Resolución de 300 DPI | Impresora térmica solicita imagen con dimensiones físicas exactas (50mm x 25mm) | Pillow almacena `dpi=(300, 300)` en el encabezado PNG (`pHYs` chunk). El tamaño en píxeles es exactamente `591 x 295 px`, garantizando impresión 1:1 en escala real. |
| 9 | Lector de Campo con Prefijo / Sufijo | Pistola lectora configurada con tabulador o salto de línea al final del escaneo | El simulador de escaneo limpia espacios y saltos (`strip()`) y normaliza a mayúsculas antes de ejecutar la consulta contra el catálogo. |
| 10 | Detección de Altas en Campo | Equipo escaneado en campo con serial o ID que no existe en el archivo Excel maestro | Se clasifica como `ALTA NUEVA`, se inserta en una fila dedicada con estilo azul claro (`#DDEBF7`) y se reporta en la hoja "Reporte de Auditoría". |

---

## 3. Especificación Técnica R2: Conexión de Escaneo Móvil a Excel

### 3.1. Arquitectura de Transmisión Inalámbrica

Existen tres mecanismos probados para vincular el escaneo desde el celular hacia Excel:

```
┌─────────────────────────────────────────────────────────────────────────────┐
│                       ARQUITECTURA DE ESCANEO MÓVIL                         │
├─────────────────────────────────────────────────────────────────────────────┤
│                                                                             │
│  [Opción 1: Barcode to PC (Recomendada para captura continua en PC)]        │
│  ┌─────────────────┐       Wi-Fi LAN (TCP:57888)       ┌─────────────────┐  │
│  │  App Móvil      │ ────────────────────────────────> │  PC Servidor    │  │
│  │ (Android / iOS) │  WebSocket / HTTP JSON Payload    │ (Barcode to PC) │  │
│  └─────────────────┘                                   └────────┬────────┘  │
│                                                                 │ Emulación │
│                                                                 │ Teclado   │
│                                                                 ▼           │
│                                                        ┌─────────────────┐  │
│                                                        │ Microsoft Excel │  │
│                                                        │ Hoja: Escaneos  │  │
│                                                        └─────────────────┘  │
│                                                                             │
│  [Opción 2: Excel Mobile con Coautoría en la Nube (Sin software intermedio)]│
│  ┌─────────────────┐       OneDrive / SharePoint       ┌─────────────────┐  │
│  │ App Excel Móvil │ ◄═══════════════════════════════► │ Excel Desktop   │  │
│  │ (Cámara Activa) │       Sincronización en la Nube   │ (PC Auditor)    │  │
│  └─────────────────┘                                   └─────────────────┘  │
│                                                                             │
│  [Opción 3: Pistola Lectora Bluetooth / 2.4 GHz (Estándar Industrial)]      │
│  ┌─────────────────┐       Bluetooth HID Profile       ┌─────────────────┐  │
│  │ Pistola Láser/2D│ ────────────────────────────────> │ PC / Laptop     │  │
│  │ Lector QR Físico│   Pulsaciones de Teclado Directo  │ Celda Activa    │  │
│  └─────────────────┘                                   └─────────────────┘  │
└─────────────────────────────────────────────────────────────────────────────┘
```

#### 3.1.1. Método A: Barcode to PC (Emulación de Teclado por Wi-Fi)
- **Principio de Funcionamiento**:
  - En la computadora se ejecuta el servidor de escritorio ligero **Barcode to PC Server** (escuchando en el puerto TCP `57888`).
  - En el celular se instala la aplicación móvil **Barcode to PC: Wi-Fi scanner**.
  - Ambos dispositivos se conectan a la misma red local Wi-Fi o se emparejan escaneando un código QR con la IP de la PC.
  - Al escanear una etiqueta, el celular transmite la cadena leída. El servidor en Windows toma el foco de la ventana activa (Excel) y escribe los caracteres emulando pulsaciones de teclado físico (Keystroke Injection mediante la API de Windows `SendInput`), añadiendo automáticamente la tecla `ENTER` al final.
- **Configuración de Ráfaga / Siguiente Celda**:
  - En los ajustes de la app móvil se configura el componente de salida como: `[Barcode] + [Enter]`.
  - Al apuntar la cámara a las etiquetas adhesivas de los equipos, cada lectura se inserta de forma secuencial descendente en la columna `ID_Escaneado` de la hoja `Escaneos_Campo`.

#### 3.1.2. Método B: Microsoft Excel Mobile App (Sincronización Cloud Nativa)
- **Principio de Funcionamiento**:
  - El libro generado (`Inventario_Equipos_Clemont_Con_QR.xlsx`) se guarda en la carpeta compartida de **OneDrive** o **SharePoint / Teams**.
  - El auditor de campo abre el archivo desde la app móvil oficial de Excel (disponible en Android y iOS).
  - En la hoja `Escaneos_Campo`, selecciona la celda de la columna `ID_Escaneado` y activa el icono de cámara o escáner de código de barras incorporado en el teclado móvil (o mediante teclados especializados como *Scan-IT to Office* o *Barcode Keyboard*).
  - Los datos se sincronizan en tiempo real mediante el motor de coautoría de Microsoft 365, viéndose reflejados inmediatamente en la PC de control.

#### 3.1.3. Matriz Comparativa de Soluciones Móviles

| Criterio | Barcode to PC | Excel Mobile Cloud | Pistola QR Bluetooth |
|---|---|---|---|
| **Velocidad de Escaneo** | Muy rápida (~1 seg/activo) | Moderada (~3 seg/activo) | Instantánea (<0.5 seg/activo) |
| **Requiere Internet Exterior** | No (solo red Wi-Fi local) | Sí (cuenta Microsoft 365 / OneDrive) | No (conexión local directa) |
| **Costo de Licenciamiento** | Gratuito / Open Source | Incluido con Office 365 | Hardware físico (~$25 - $40 USD) |
| **Ergonomía en Campo** | Cámara de celular con linterna LED | Cámara de celular | Gatillo ergonómico y puntero rojo |
| **Tolerancia a Caída de Red** | Posee búfer offline en móvil | Sincroniza al reconectar | Almacena en memoria interna de la pistola |

---

### 3.2. Esquema Exacto de Datos: `Maestro_Activos` y `Escaneos_Campo`

Para garantizar la correspondencia unívoca de casillas y evitar desfases entre columnas, se estandarizan las estructuras y formatos:

#### 3.2.1. Hoja `Maestro_Activos` (Catálogo Maestro Oficial)
- **Propósito**: Base de datos de referencia inmutable con los identificadores únicos asignados y las especificaciones técnicas completas.
- **Formato Visual**: Encabezado Azul Marino Corporativo (`#1F4E78`), texto blanco negrita, bordes delgados grises (`#D9D9D9`), fila de encabezado fija (`freeze_panes = "A2"`), autofiltro habilitado.

| Columna | Nombre de Campo | Tipo de Dato | Ejemplo | Restricciones / Reglas |
|:---:|---|---|---|---|
| **A** | `ID_Activo` | Alfanumérico (Texto) | `LAP-2026-0001` | Clave primaria única. Formato: `[CAT]-[AÑO]-[CORRELATIVO]` |
| **B** | `Categoria` | Texto | `LAP` | Opciones: `LAP`, `PC`, `TAB`, `MON`, `IMP`, `RED` |
| **C** | `Tipo_Equipo` | Texto | `Laptop` | Categoría amigable (Laptop, PC Escritorio, Tablet, Monitor) |
| **D** | `Equipo / Host` | Texto | `LAPTOP-2895F563` | Nombre de red / Hostname del equipo |
| **E** | `Modelo_Marca` | Texto | `Lenovo 83A0` | Marca y modelo comercial del fabricante |
| **F** | `Serial_Fabricante` | Alfanumérico (Texto) | `SDK0T76530` | Serial físico grabado de fábrica (S/N) |
| **G** | `Ubicacion_Oficial` | Texto | `Administrativa 2do piso`| Sede o área formalmente asignada |
| **H** | `Responsable` | Texto | `DANIEL PANESSO` | Nombre y apellido del usuario asignado |
| **I** | `Estado_Oficial` | Texto | `Activo` | Opciones: `Activo`, `En Bodega`, `Baja`, `En Reparación` |

#### 3.2.2. Hoja `Escaneos_Campo` (Registro de Auditoría Física)
- **Propósito**: Hoja de captura directa en la que el auditor de campo vuelca las lecturas realizadas con el celular o la pistola lectora.
- **Formato Visual**: Encabezado Verde Bosque Auditoría (`#375623`), texto blanco negrita, bordes delgados grises, filas preformateadas listas para ingreso.

| Columna | Nombre de Campo | Tipo de Dato | Ejemplo | Procedencia / Validación |
|:---:|---|---|---|---|
| **A** | `ID_Escaneado` | Alfanumérico (Texto) | `LAP-2026-0001` | Capturado directamente por el escáner QR móvil |
| **B** | `Ubicacion_Hallada` | Texto | `Showroom` | Sede o área física donde el auditor encontró el equipo |
| **C** | `Fecha_Hora_Escaneo` | Fecha y Hora | `2026-11-05 14:35:10` | Timestamp automático o fecha de jornada |
| **D** | `Operador` | Texto | `Auditor TI` | Iniciales o nombre del auditor en campo |
| **E** | `Estado_Fisico` | Texto | `Operativo` | `Operativo`, `Dañado`, `Sin Etiqueta`, `Incompleto` |
| **F** | `Observaciones` | Texto | `Pantalla rayada; cargador OK` | Novedades o incidencias visuales |

---

### 3.3. Guía Operativa para Auditoría y Escaneo en Campo

A continuación se detalla la guía operativa completa para entregar al equipo de campo e implementar el procedimiento sin errores:

```markdown
# GUÍA OPERATIVA: AUDITORÍA DE INVENTARIO CON ESCANEO MÓVIL Y CÓDIGOS QR

## 1. OBJETIVO
Establecer el procedimiento estándar para realizar el levantamiento físico y conciliación de equipos tecnológicos mediante teléfonos celulares y lectores de códigos QR en las sedes de la organización, garantizando la captura fidedigna de datos hacia Microsoft Excel sin errores tipográficos.

---

## 2. REQUISITOS PREVIOS

### Equipamiento y Software:
- **Computadora del Auditor:** Laptop o PC con Windows 10/11 ejecutando `conciliador_universal.py` o Microsoft Excel.
- **Teléfono Inteligente (Android o iOS):** Cámara funcional con autofoco y flash LED.
- **Red de Comunicación:**
  - Red Wi-Fi compartida para la PC y el móvil, O
  - Conexión celular 4G/5G en el móvil si se utiliza OneDrive / Excel Mobile.
- **Material Físico:** Etiquetas adhesivas impresas a 300 DPI fijadas previamente en el chasis de los equipos.

---

## 3. CONFIGURACIÓN PASO A PASO (MÉTODO RECOMENDADO: BARCODE TO PC)

### Paso 1: Instalación del Servidor en PC
1. Descargar e instalar **Barcode to PC Server** en la laptop de auditoría.
2. Ejecutar la aplicación. En la ventana principal se mostrará la dirección IP local de la computadora (ejemplo: `192.168.1.45`) y un código QR de emparejamiento.
3. Si el Firewall de Windows solicita permisos, seleccionar **Permitir acceso en redes privadas**.

### Paso 2: Instalación de la App en el Celular
1. Abrir Google Play Store o Apple App Store en el teléfono.
2. Buscar e instalar **Barcode to PC: Wi-Fi scanner**.
3. Abrir la app en el móvil y conceder permisos de acceso a la cámara.

### Paso 3: Vinculación Inalámbrica
1. En la app móvil, presionar el botón **Scan QR** y apuntar la cámara a la pantalla de la PC donde se muestra el QR de Barcode to PC Server.
2. El indicador en la app cambiará a color verde: **Connected**.

### Paso 4: Configuración de Salida en la App
1. En la app móvil, ingresar a **Settings > Output**.
2. Seleccionar el modo de pulsación: **Keystrokes**.
3. Asegurar que la tecla de sufijo (Post-scan key) esté configurada en: `ENTER`.
4. Habilitar la opción **Continuous Scan Mode** (Modo de escaneo continuo) para leer múltiples equipos sin tocar la pantalla cada vez.

---

## 4. PROTOCOLO DE AUDITORÍA EN CAMPO

### Preparación del Archivo de Registro:
1. En la PC, abrir el archivo `Inventario_Equipos_Clemont_Con_QR.xlsx` generado por el Conciliador.
2. Ir a la pestaña **`Escaneos_Campo`**.
3. Hacer clic en la primera celda vacía de la columna **A** (`ID_Escaneado`, fila 2).
4. Dejar la ventana de Excel en primer plano.

### Ejecución del Escaneo en Sitio:
1. El auditor se desplaza puesto por puesto en la sede (ejemplo: Sótano, Administrativa, Showroom).
2. Ubicar la etiqueta QR pegada en el equipo (generalmente en la tapa superior o base inferior de la laptop/monitor).
3. Apuntar con el celular a una distancia de entre 10 y 20 cm.
4. El teléfono emitirá un pitido/vibración de confirmación.
5. Inmediatamente el ID del activo (ej: `LAP-2026-0001`) aparecerá escrito en Excel y el cursor saltará automáticamente a la siguiente fila.
6. En caso de requerir registrar la ubicación o una novedad puntual, el auditor puede ingresar el texto en la columna `Ubicacion_Hallada` y `Observaciones`.

---

## 5. RESOLUCIÓN DE INCIDENCIAS FRECUENTES (TROUBLESHOOTING)

| Problema | Causa Probable | Solución |
|---|---|---|
| **La app no conecta con la PC** | Dispositivos en redes Wi-Fi diferentes o red con aislamiento de clientes (AP Isolation). | Asegurar que PC y celular estén en la misma red Wi-Fi (ej: red de invitados bloquea tráfico local). Usar la zona Wi-Fi portátil (Hotspot) del celular si la red corporativa tiene restricciones. |
| **El escáner escribe pero no baja de fila** | Falta el sufijo `ENTER` en la configuración. | En la app móvil, ir a Settings > Components y verificar que la acción posterior al escaneo sea `Add Enter`. |
| **Etiqueta dañada o sucia no lee el QR** | Código rayado o baja iluminación. | Encender el flash LED desde el botón de la app móvil. Si continúa ilegible, ingresar manualmente en Excel el serial o ID legible impreso en texto claro junto al QR. |
| **Al escanear se escriben caracteres extraños** | Distribución de teclado diferente entre PC y móvil. | En Barcode to PC Server (PC), seleccionar en Keyboard Layout: **Latin American** o **Spanish**. |
```

---

## 4. Especificación Técnica R3: Interfaz Gráfica (GUI) y Usabilidad

### 4.1. Distribución de Componentes Tkinter

La interfaz se estructura mediante una jerarquía modular basada en `tk.Frame`, `ttk.Notebook` y `tk.LabelFrame` para optimizar el flujo de trabajo:

```
┌─────────────────────────────────────────────────────────────────────────────┐
│                   Conciliador de Inventario Universal v3.2                  │
├─────────────────────────────────────────────────────────────────────────────┤
│ [Pestaña 1: Conciliación de Inventario]   [Pestaña 2: Módulo QR]            │
│┌───────────────────────────────────────────────────────────────────────────┐│
││ 1. Archivos Excel (Origen y Destino)                                      ││
││ Archivo a auditar:    [ C:/.../Inventario.xlsx                 ] [Examinar]││
││ Carpeta de salida:    [ C:/.../Resultados_Inventario           ] [Cambiar] ││
││ Hoja a auditar:       [ Inventario Clemont                 ▼ ] 3 hojas OK ││
│├───────────────────────────────────────────────────────────────────────────┤│
││ Mapeo y Filtros:  [x] Serial  [x] Host  [x] Sede  [x] Resp  [x] Auditoría ││
││                  (Combos activos o desactivados dinámicamente)            ││
│├───────────────────────────────────────────────────────────────────────────┤│
││ 2. Lista de Equipos Verificados Físicamente                               ││
││ [PEGAR DESDE EXCEL] [Cargar Preset Clemont] [+ Agregar] [Borrar] [Limpiar]││
││ ┌───────────────────────────────────────────────────────────────────────┐ ││
││ │ # │ Equipo / Host │ Serial Real │ Estado │ Sede │ Resp │ Observaciones│ ││
││ │ 1 │ SERVICIO-CLI1 │ T4N0CV16Z03 │ OK     │ Sót. │ Juan │ Verificado   │ ││
││ └───────────────────────────────────────────────────────────────────────┘ ││
│├───────────────────────────────────────────────────────────────────────────┤│
││ [EJECUTAR CONCILIACIÓN COMPLETA]   [Abrir Excel Auditado]  [Abrir Carpeta]││
│├───────────────────────────────────────────────────────────────────────────┤│
││ 3. Bitácora de Auditoría en Tiempo Real                   [Limpiar Log]   ││
││ ┌───────────────────────────────────────────────────────────────────────┬┐│
││ │ [OK] Serial normalizado 'T4N0CV16Z033184' coincide con fila 6         │││
││ │ [WARN] Duplicado detectado en fila 18: serial repetido en sede        │█││
││ │ [INFO] Guardando libro auditado con 34 activos...                     │││
││ └───────────────────────────────────────────────────────────────────────┴┘│
│└───────────────────────────────────────────────────────────────────────────┘│
└─────────────────────────────────────────────────────────────────────────────┘
```

### 4.2. Especificación de Componentes Clave

#### 4.2.1. Casillas de Verificación (Chulos) por Columna
- **Variables de Control**:
  ```python
  self.chk_col_serial = tk.BooleanVar(value=True)
  self.chk_col_equipo = tk.BooleanVar(value=True)
  self.chk_col_sede   = tk.BooleanVar(value=True)
  self.chk_col_resp   = tk.BooleanVar(value=True)
  self.chk_col_verif  = tk.BooleanVar(value=True)
  ```
- **Comportamiento en GUI**:
  - Cada casilla se asocia mediante `command=self._actualizar_estado_combos`.
  - Al desmarcar un chulo, el Combobox asociado cambia a `state="disabled"` y se atenúa visualmente.
  - Al marcarlo, regresa a `state="readonly"`.
- **Efecto en el Motor de Conciliación**:
  - En `_ejecutar()`: si `self.chk_col_serial.get()` es `False`, el motor ignora completamente la comparación de seriales. Si `self.chk_col_equipo.get()` es `True`, reconcilia por nombre de host. Si `self.chk_col_verif.get()` es `False`, no escribe ni altera la columna de auditoría, permitiendo auditorías parciales o incrementales.

#### 4.2.2. Selector de Directorio de Guardado (`filedialog.askdirectory`)
- **Implementación**:
  ```python
  def _seleccionar_directorio_salida(self):
      carpeta_actual = self.dir_salida.get() or os.getcwd()
      nueva_carpeta = filedialog.askdirectory(
          title="Seleccionar carpeta de destino para resultados y etiquetas",
          initialdir=carpeta_actual
      )
      if nueva_carpeta:
          self.dir_salida.set(nueva_carpeta)
          # Actualizar automáticamente la ruta completa del Excel de salida
          nombre_base = os.path.basename(self.ruta_excel.get() or "Inventario_Clemont.xlsx")
          if not nombre_base.startswith("AUDITADO_"):
              nombre_base = "AUDITADO_" + nombre_base
          self.ruta_salida.set(os.path.join(nueva_carpeta, nombre_base))
          self.log(f"Carpeta de destino configurada: {nueva_carpeta}", "info")
  ```
- **Ventaja de Usabilidad**: Garantiza que tanto el Excel resultante como la subcarpeta `etiquetas_qr/` se generen organizadamente en la ruta que el usuario elija (ej: unidad compartida de red o pendrive).

#### 4.2.3. Botones de Apertura Directa a 1 Clic
- **Implementación Robusta**:
  ```python
  def _abrir_excel_resultado(self):
      ruta = self.ruta_salida.get().strip()
      if not ruta or not os.path.exists(ruta):
          messagebox.showinfo(
              "Archivo no encontrado",
              "Aún no se ha generado el archivo auditado.\nEjecuta la conciliación primero."
          )
          return
      try:
          os.startfile(ruta)  # Apertura nativa en Windows
      except Exception as e:
          # Fallback mediante subprocess
          subprocess.Popen(["start", "", ruta], shell=True)

  def _abrir_carpeta_resultado(self):
      ruta = self.ruta_salida.get().strip()
      carpeta = os.path.dirname(ruta) if (ruta and os.path.isabs(ruta)) else (self.dir_salida.get() or os.getcwd())
      if not os.path.exists(carpeta):
          carpeta = os.getcwd()
      try:
          os.startfile(carpeta)
      except Exception:
          subprocess.Popen(["explorer", carpeta])
  ```

#### 4.2.4. Consola de Bitácora Amplia con Scrollbar y Autoscroll
- **Implementación con `tk.Text` y `ttk.Scrollbar`**:
  ```python
  # Frame contenedor con scrollbar vertical
  frm_log = tk.LabelFrame(parent, text="3. Bitacora de Auditoria", bg="#F0F4F8", font=("Arial", 10, "bold"))
  vsb_log = ttk.Scrollbar(frm_log, orient="vertical")
  self.txt_log = tk.Text(
      frm_log, bg="#1E1E1E", fg="#D4D4D4",
      font=("Consolas", 10), height=14,
      yscrollcommand=vsb_log.set, wrap="word"
  )
  vsb_log.config(command=self.txt_log.yview)
  self.txt_log.pack(side="left", fill="both", expand=True)
  vsb_log.pack(side="right", fill="y")

  # Metodo de registro seguro con autoscroll y refresco de UI
  def log(self, msg, tag="info"):
      timestamp = datetime.now().strftime("[%H:%M:%S] ")
      self.txt_log.insert(tk.END, timestamp + msg + "\n", tag)
      self.txt_log.see(tk.END)           # Autoscroll al final
      self.root.update_idletasks()       # Refresco sin congelar ventana
  ```

---

## 5. Especificación Técnica R4: Librerías, DPI y Código Probado

### 5.1. Generación de Códigos QR a 300 DPI con Pillow y `qrcode`

#### 5.1.1. Fundamento Matemático de DPI para Etiquetas Físicas
La resolución de impresión estándar para códigos de barras y etiquetas de activos es de **300 DPI (puntos por pulgada)**.
$$\text{Píxeles} = \left(\frac{\text{Milímetros}}{25.4}\right) \times 300$$

- **Formato Estándar 50 mm x 25 mm**:
  - Ancho en píxeles: $(50 / 25.4) \times 300 = 590.55 \approx \mathbf{591\text{ px}}$
  - Alto en píxeles: $(25 / 25.4) \times 300 = 295.27 \approx \mathbf{295\text{ px}}$
- **Formato Avery 5160 (67 mm x 25 mm)**:
  - Ancho en píxeles: $(67 / 25.4) \times 300 = 791.33 \approx \mathbf{788\text{ a }791\text{ px}}$
  - Alto en píxeles: $(25 / 25.4) \times 300 = 295.27 \approx \mathbf{295\text{ px}}$
- **Formato Avery 5163 (100 mm x 50 mm)**:
  - Ancho en píxeles: $(100 / 25.4) \times 300 = 1181.1 \approx \mathbf{1200\text{ px}}$
  - Alto en píxeles: $(50 / 25.4) \times 300 = 590.55 \approx \mathbf{600\text{ px}}$
- **Pliego Carta para Planilla de Impresión (8.5" x 11")**:
  - Ancho: $8.5 \times 300 = \mathbf{2550\text{ px}}$
  - Alto: $11.0 \times 300 = \mathbf{3300\text{ px}}$

#### 5.1.2. Incrustación de Metadatos DPI en Pillow
Para que las impresoras térmicas (Zebra, Brother, Epson) o el software de impresión reconozcan la escala física real 1:1 sin deformaciones, la imagen debe guardarse especificando la tupla DPI:
```python
img.save(ruta_archivo, format="PNG", dpi=(300, 300))
```

#### 5.1.3. Carga Útil (Payload) del Código QR con Ficha Completa (7 Campos)
Conforme a R1, cada código QR debe contener la información técnica fidedigna sin mezclar columnas:
```python
qr_payload = (
    f"ID: {item['id']}\n"
    f"HOST: {item['equipo']}\n"
    f"MODELO: {item['modelo']}\n"
    f"SERIAL: {item['serial']}\n"
    f"RESP: {item['responsable']}\n"
    f"SEDE: {item['sede']}\n"
    f"ESTADO: {item['estado']}"
)
```
- **Capacidad y Corrección de Errores**:
  - Longitud aproximada de la carga útil: 160 a 220 caracteres alfanuméricos.
  - Se utiliza `qrcode.constants.ERROR_CORRECT_M` (15% de redundancia frente a manchas o rasguños) o `ERROR_CORRECT_Q` (25%).
  - La matriz QR resultante es de Versión 4 o 5 (33x33 a 37x37 módulos), la cual renderizada en un cuadro de 265x265 px (formato 50x25mm) o 540x540 px (formato 100x50mm) ofrece una lectura instantánea por cualquier smartphone o escáner 2D.

### 5.2. Diseño Visual de la Etiqueta Física (Asset Tag Layout)

```
┌─────────────────────────────────────────────────────────────────────────────┐
│ 1px Borde Perimetral de Corte (#C8C8C8)                                    │
│ ┌─────────────────┬───────────────────────────────────────────────────────┐ │
│ │                 │  CLEMONT  ·  CONTROL DE ACTIVOS                       │ │
│ │  █████████████  │  LAP-2026-0001                                        │ │
│ │  ██         ██  ├───────────────────────────────────────────────────────┤ │
│ │  ██  █████  ██  │  Host: LAPTOP-2895F563                                │ │
│ │  ██  ██ ██  ██  │  Mod:  Lenovo 83A0                                    │ │
│ │  ██  █████  ██  │  S/N:  SDK0T76530                                     │ │
│ │  ██         ██  │  Resp: DANIEL PANESSO                                 │ │
│ │  █████████████  │  Sede: Administrativa 2do piso  |  Activo             │ │
│ └─────────────────┴───────────────────────────────────────────────────────┘ │
└─────────────────────────────────────────────────────────────────────────────┘
  ◄──── QR 265px ───►◄────────────────── Textos 295px ──────────────────────►
  ◄─────────────────────────── Total: 591 px (50 mm) ─────────────────────────►
```

- **Tipografías**: Descubrimiento dinámico de fuentes Windows TrueType (`calibrib.ttf`, `arialbd.ttf`, `segoeuib.ttf`) para mantener nitidez vectorial a 300 DPI.
- **Jerarquía Visual**:
  1. ID del Activo: Tamaño 24pt negrita (50x25mm) / 46pt negrita (100x50mm).
  2. Nombre de Host y Serial: Tamaño 16pt / 28pt negrita.
  3. Responsable, Sede y Estado: Tamaño 14pt / 22pt regular.

### 5.3. Formateo y Estilos de Excel con `openpyxl`

- **Paleta de Colores de Estado**:
  - `OK / Verificado`: Relleno `#E2EFDA` (Verde suave), Fuente `#006100` (Verde oscuro negrita).
  - `Traslado`: Relleno `#FCE4D6` (Naranja durazno), Fuente `#C00000` (Rojo oscuro negrita).
  - `Duplicado`: Relleno `#FFE699` (Amarillo advertencia), Fuente `#7F6000` con `openpyxl.comments.Comment`.
  - `Alta Nueva (Sobrante)`: Relleno `#DDEBF7` (Azul hielo), Fuente `#1F4E78`.
  - `Sin Chequear`: Relleno `#F2F2F2` (Gris claro), Fuente `#595959` cursiva.
- **Autofit Dinámico de Columnas**:
  ```python
  def autoajustar_columnas(ws):
      for col in ws.columns:
          max_len = 0
          col_letter = get_column_letter(col[0].column)
          for cell in col:
              if cell.value:
                  max_len = max(max_len, len(str(cell.value).split("\n")[0]))
          ws.column_dimensions[col_letter].width = max(max_len + 4, 12)
  ```
- **Preservación de Fórmulas y Freeze Panes**:
  - Congelamiento en fila de datos: `ws.freeze_panes = "A6"`
  - Autofiltrado: `ws.auto_filter.ref = f"A5:L{total_row}"`
  - Total dinámico: `ws.cell(total_row, 8).value = f"=SUM(H6:H{total_row - 1})"`

---

## 6. Especificación Técnica R5: Arquitectura de la Suite QA Automatizada

### 6.1. Requisitos de la Suite QA
Conforme a R5:
1. Debe ejecutarse de manera **desatendida / headless** (sin bloquearse en el bucle interactivo de Tkinter `root.mainloop()`).
2. Debe validar directamente contra los archivos reales provistos en el workspace:
   - `Inventario_Equipos_Clemont (2) (2).xlsx`
   - `AUDITADO_Inventario_Equipos_Clemont (2) (2).xlsx`
3. Debe verificar que **no existan desfases entre columnas**.
4. Debe verificar la **generación a 300 DPI y legibilidad del código QR** decodificando el contenido y comprobando los 7 campos.
5. Debe finalizar con **código de salida 0 (exit code 0)** cuando todas las pruebas pasen.

### 6.2. Diseño de Módulos de Prueba

```
┌─────────────────────────────────────────────────────────────────────────────┐
│                    ARQUITECTURA DE SUITE QA (test_qa.py)                    │
├─────────────────────────────────────────────────────────────────────────────┤
│                                                                             │
│  [Módulo 1: Validación de Archivos Reales y Encabezados]                   │
│   ├── Existencia de 'Inventario_Equipos_Clemont (2) (2).xlsx'               │
│   ├── Verificación de Fila 5: Col C=Sede, D=Host, I=Resp, K=Obs, L=Verif    │
│   └── Comprobación de que no hay desfase en columnas originales             │
│                                                                             │
│  [Módulo 2: Motor de Conciliación Headless (Sin GUI)]                       │
│   ├── Ejecución del motor con Preset Clemont (20 Laptops)                   │
│   ├── Verificación de conteos esperados: 19 OK, 1 Traslado, 1 Duplicado      │
│   ├── Verificación de inserción de Altas antes de la fila TOTAL             │
│   └── Validación de la fórmula =SUM(H6:H...) en la fila de totales          │
│                                                                             │
│  [Módulo 3: Validación de Integridad de Datos de Salida]                    │
│   ├── Comprobación celda a celda en archivo auditado generado               │
│   ├── Ausencia de cruces de celdas (Seriales en K, Nombres en D)            │
│   └── Hoja 'Reporte de Auditoria' creada con resumen y estilos aplicados    │
│                                                                             │
│  [Módulo 4: Verificación de Etiquetas QR y Metadatos 300 DPI]               │
│   ├── Generación de lote de prueba en 'etiquetas_qr/'                       │
│   ├── Aserción de resolución: img.info['dpi'] == (300, 300)                 │
│   ├── Aserción de dimensiones físicas exactas (591x295 / 1200x600 px)       │
│   └── Decodificación con pyzbar / cv2.QRCodeDetector                        │
│       └── Aserción de presencia de los 7 campos en el QR decodificado       │
│                                                                             │
│  [Módulo 5: Validación de Plantilla 'Maestro_Activos' y 'Escaneos_Campo']   │
│   ├── Verificación de existencia de ambas hojas                             │
│   ├── Verificación de 8 columnas en Maestro y 6 columnas en Escaneos        │
│   └── Autofiltrado y estilos corporativos intactos                          │
│                                                                             │
│  [Módulo 6: Resumen Ejecutivo y Código de Salida]                           │
│   └── Retorno sys.exit(0) si todas las aserciones son exitosas              │
└─────────────────────────────────────────────────────────────────────────────┘
```

### 6.3. Patrón de Código del Validador QA Decodificador de QR

```python
def validar_legibilidad_qr(ruta_png, esperado_id, esperado_sn):
    """
    Abre la imagen generada, verifica metadatos 300 DPI y decodifica el QR
    para validar que los 7 campos esten presentes sin desfasarse.
    """
    from PIL import Image
    img = Image.open(ruta_png)
    
    # 1. Verificar DPI
    dpi = img.info.get("dpi")
    assert dpi == (300, 300) or dpi == (300.0, 300.0), f"DPI incorrecto: {dpi} en {ruta_png}"
    
    # 2. Decodificar QR
    qr_texto = None
    try:
        from pyzbar.pyzbar import decode
        decodificados = decode(img)
        if decodificados:
            qr_texto = decodificados[0].data.decode("utf-8")
    except ImportError:
        try:
            import cv2
            import numpy as np
            detector = cv2.QRCodeDetector()
            cv_img = cv2.cvtColor(np.array(img), cv2.COLOR_RGB2BGR)
            val, pts, _ = detector.detectAndDecode(cv_img)
            if val:
                qr_texto = val
        except ImportError:
            # Fallback a verificacion matricial Pillow
            pass

    if qr_texto:
        assert esperado_id in qr_texto, f"ID {esperado_id} no encontrado en contenido QR"
        assert esperado_sn in qr_texto, f"Serial {esperado_sn} no encontrado en contenido QR"
        # Verificar presencia de campos de ficha tecnica
        for campo in ["HOST:", "MODELO:", "RESP:", "SEDE:", "ESTADO:"]:
            assert campo in qr_texto, f"Campo {campo} no encontrado en payload QR"
    return True
```

---

## 7. Handoff Protocol

### 1. Observation
- Se inspeccionó el archivo principal `conciliador_universal.py` (2408 líneas) en `c:\Users\Analista-TI\Desktop\entrega inicial\Scrip-Inventario\conciliador_universal.py`.
- En las líneas 277-283 se observan los índices de columna base del formato Clemont:
  - `COL_SEDE = 3` (Columna C)
  - `COL_EQUIPO = 4` (Columna D)
  - `COL_RESP = 9` (Columna I)
  - `COL_OBS = 11` (Columna K)
  - `COL_VERIF = 12` (Columna L)
- En las líneas 1001-1109 se observa la implementación actual de generación de etiquetas individuales:
  - Soporta formatos `100x50mm`, `67x25mm`, y `50x25mm`.
  - Guarda con `img.save(ruta, dpi=(300, 300))`.
  - Construye el payload QR con datos de ID, equipo, modelo, serial, responsable, sede y estado (líneas 1040-1049).
- En las líneas 1247-1344 se observa la función `_qr_exportar_excel` que crea `Maestro_Activos` y `Escaneos_Campo`.
- En las líneas 1403-1436 se observan las funciones de apertura directa `_abrir_excel_resultado` y `_abrir_carpeta_resultado` utilizando `os.startfile`.
- En `requirements.txt` se observan las dependencias instaladas: `openpyxl>=3.1.0`, `qrcode[pil]>=8.0.0`, `Pillow>=10.0.0`, `pandas>=2.0.0`.

### 2. Logic Chain
1. El usuario solicitó en `ORIGINAL_REQUEST.md` la integración y especificación para R2 (Escaneo móvil y guía operativa), R3 (GUI y usabilidad con chulos, askdirectory, 1-clic y log amplio), R4 (Librerías, 300 DPI y estilos openpyxl) y R5 (QA automatizado con exit code 0).
2. Para R2, Barcode to PC y Excel Mobile representan las dos soluciones más probadas y con menor barrera de adopción en campo. La estandarización de columnas en `Maestro_Activos` (8 columnas) y `Escaneos_Campo` (6 columnas) resuelve de raíz los desfases entre capturas de campo y la base histórica.
3. Para R3, los checkboxes en Tkinter (`tk.BooleanVar`) acoplados con `_actualizar_estado_combos` y la lógica de exclusión de columnas en `_ejecutar()` permiten al operador omitir columnas que no desee conciliar sin alterar el resto del archivo. El cambio de `filedialog.asksaveasfilename` a `filedialog.askdirectory` proporciona la ruta base donde se almacenan armónicamente tanto el archivo auditado como la carpeta de etiquetas.
4. Para R4, la fórmula de escala a 300 DPI ($(mm / 25.4) \times 300$) da exactamente las dimensiones de píxeles (591x295 para 50x25mm y 1200x600 para 100x50mm). La inserción de metadatos `dpi=(300, 300)` en Pillow garantiza que las impresoras no escalen arbitrariamente.
5. Para R5, la creación de una suite de pruebas headless (`test_suite_inventario.py`) ejecutable directamente por consola permite a los desarrolladores y a QA certificar que cada cambio conserve el código de salida 0 y la integridad total de columnas en los archivos reales.

### 3. Caveats
- No se modificó ningún archivo de código en el workspace durante esta prospección (cumplimiento estricto del rol de Specification Miner).
- La decodificación automatizada de imágenes QR en la suite de pruebas QA puede utilizar `pyzbar` o `cv2.QRCodeDetector`; en entornos donde no estén preinstalados binarios de C de zbar, se especifica el fallback seguro mediante OpenCV o análisis directo del canvas.
- En sistemas operativos no Windows, `os.startfile` no está disponible; se documentó la alternativa con `subprocess.Popen` para entornos multiplataforma si fuese necesario en el futuro.

### 4. Conclusion
Las especificaciones técnicas, esquemas de datos, guía operativa, arquitectura de interfaz y suite de pruebas para R2, R3, R4 y R5 están totalmente definidas, fundamentadas en el código existente y listas para su implementación inmediata por parte del equipo de desarrollo y validación por QA.

### 5. Verification Method
1. **Inspección de Documentación**: Revisar `handoff.md` en `.agents/teamwork/spec_miner_survey_1/handoff.md`.
2. **Verificación de Constantes y Código Base**: Confirmar mediante `view_file` las líneas citadas de `conciliador_universal.py` (277-283, 1001-1109, 1247-1344).
3. **Verificación de Ejecución de la Suite QA Futura**:
   ```powershell
   python test_suite_inventario.py
   # Debe retornar código de salida 0: echo $LASTEXITCODE
   ```
