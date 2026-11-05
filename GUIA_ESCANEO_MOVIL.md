# GUÍA OPERATIVA: AUDITORÍA DE INVENTARIO CON ESCANEO MÓVIL Y CÓDIGOS QR

**Organización**: Clemont · Control y Gestión de Activos de TI  
**Documento**: Manual de Procedimiento Operativo Estándar (SOP)  
**Versión**: 3.2 — Noviembre 2026  
**Alcance**: Levantamiento físico, escaneo inalámbrico y conciliación automatizada en sedes operativas y administrativas.

---

## 1. OBJETIVO DEL PROCEDIMIENTO

Establecer el protocolo estándar para realizar el levantamiento físico, auditoría y conciliación de activos tecnológicos (Laptops, PCs, Tablets, Monitores, Impresoras y Equipos de Red) mediante dispositivos móviles inteligentes y lectores de códigos 2D/QR.

Este procedimiento garantiza:
1. La captura fidedigna de identificadores de activos (`ID_Activo`) y seriales de fábrica (`S/N`) sin errores de digitación ni transposición de caracteres.
2. La transmisión directa e inalámbrica de las lecturas hacia la plantilla corporativa de Microsoft Excel (`Inventario_Equipos_Clemont_Con_QR.xlsx`), específicamente en la hoja de auditoría `Escaneos_Campo`.
3. La conciliación inmediata en tiempo real frente al catálogo oficial `Maestro_Activos` para detectar equipos conciliados, traslados no registrados, faltantes y altas nuevas (sobrantes).

---

## 2. ARQUITECTURA DE CONECTIVIDAD Y MÉTODOS DISPONIBLES

Se disponen de tres soluciones homologadas para el escaneo en campo. El auditor seleccionará el método según la disponibilidad de red y equipamiento:

```
┌─────────────────────────────────────────────────────────────────────────────┐
│                       ARQUITECTURA DE ESCANEO MÓVIL                         │
├─────────────────────────────────────────────────────────────────────────────┤
│                                                                             │
│  [MÉTODO 1: Barcode to PC (Recomendado - Inalámbrico por Wi-Fi Local)]      │
│  ┌─────────────────┐       Wi-Fi LAN (TCP:57888)       ┌─────────────────┐  │
│  │ Teléfono Móvil  │ ────────────────────────────────> │ Laptop Auditor  │  │
│  │ (Android / iOS) │  WebSocket / Keystroke Emulation  │ (Servidor B2PC) │  │
│  └─────────────────┘                                   └────────┬────────┘  │
│                                                                 │ Inyección │
│                                                                 │ [ENTER]   │
│                                                                 ▼           │
│                                                        ┌─────────────────┐  │
│                                                        │ Microsoft Excel │  │
│                                                        │ Hoja: Escaneos  │  │
│                                                        └─────────────────┘  │
│                                                                             │
│  [MÉTODO 2: Excel Mobile con Coautoría en la Nube (Sin Servidor Local)]     │
│  ┌─────────────────┐       OneDrive / SharePoint       ┌─────────────────┐  │
│  │ App Excel Móvil │ ◄═══════════════════════════════► │ Excel Desktop   │  │
│  │ (Cámara Activa) │       Sincronización en la Nube   │ (PC Central TI) │  │
│  └─────────────────┘                                   └─────────────────┘  │
│                                                                             │
│  [MÉTODO 3: Pistola Lectora 2D Inalámbrica (Hardware Industrial)]           │
│  ┌─────────────────┐       Dongle USB 2.4GHz / BT      ┌─────────────────┐  │
│  │ Pistola Óptica  │ ────────────────────────────────> │ Laptop Auditor  │  │
│  │ Lector QR / S/N │      HID Keyboard Scan Directo    │ Celda Activa    │  │
│  └─────────────────┘                                   └─────────────────┘  │
└─────────────────────────────────────────────────────────────────────────────┘
```

### Tabla Comparativa de Soluciones

| Característica | Método 1: Barcode to PC | Método 2: Excel Mobile Cloud | Método 3: Pistola 2D USB/BT |
|---|---|---|---|
| **Velocidad de captura** | ~1 seg por equipo | ~3-4 seg por equipo | < 0.5 seg por equipo |
| **Dependencia de Internet** | Nula (solo red local o hotspot) | Requiere conexión a Internet activa | Nula (conexión local directa) |
| **Costo de licenciamiento** | Gratuito / Código abierto | Incluido en suscripción M365 | Hardware físico dedicado |
| **Operación a dos manos** | Teléfono móvil con linterna LED | Teléfono móvil | Pistola ergonómica con gatillo |
| **Comportamiento sin red** | Memoria búfer en app móvil | Caché offline de Office | Memoria interna de escáner |

---

## 3. ESQUEMA DE DATOS Y CORRESPONDENCIA DE COLUMNAS (R2)

El sistema genera la plantilla normalizada `Inventario_Equipos_Clemont_Con_QR.xlsx` con correspondencia exacta entre las hojas de control:

### 3.1. Hoja: `Maestro_Activos` (Catálogo Oficial)
Contiene la ficha técnica completa pre-registrada de los activos:

| Columna | Nombre de Campo | Tipo de Dato | Ejemplo | Descripción |
|:---:|---|---|---|---|
| **A** | `ID_Activo` | Texto | `LAP-2026-0001` | Identificador único estandarizado del activo |
| **B** | `Categoria` | Texto | `LAP` | Nemónico de categoría (`LAP`, `PC`, `TAB`, `MON`, `IMP`, `RED`) |
| **C** | `Tipo_Equipo` | Texto | `Laptop` | Descripción amigable de la clase de activo |
| **D** | `Equipo / Host` | Texto | `LAPTOP-2895F563` | Nombre de red o hostname asignado |
| **E** | `Modelo_Marca` | Texto | `Vivobook ASUS laptop X1502VA` | Fabricante comercial y modelo de equipo |
| **F** | `Serial_Fabricante` | Alfanumérico | `T4N0CV16Z033184` | Serial físico grabado de fábrica (S/N) |
| **G** | `Ubicacion_Oficial` | Texto | `Sótano` | Sede o centro operativo formal de asignación |
| **H** | `Responsable` | Texto | `ESTEBAN MUÑOZ` | Nombre del colaborador asignado |
| **I** | `Estado_Oficial` | Texto | `Activo` | Estado administrativo actual en sistema |

### 3.2. Hoja: `Escaneos_Campo` (Registro de Captura Física)
Hoja preparada para recibir las lecturas del escáner en sitio:

| Columna | Nombre de Campo | Tipo de Dato | Ejemplo | Modo de Captura |
|:---:|---|---|---|---|
| **A** | `ID_Escaneado` | Texto | `LAP-2026-0001` | **Transmisión automática desde escáner móvil** |
| **B** | `Ubicacion_Hallada` | Texto | `Sótano` | Sede física donde el auditor efectúa la lectura |
| **C** | `Fecha_Hora_Escaneo` | Timestamp | `2026-11-05 14:10:00` | Fecha y hora de captura |
| **D** | `Operador` | Texto | `Auditor TI` | Iniciales o nombre del auditor en campo |
| **E** | `Estado_Fisico` | Texto | `Operativo` | `Operativo`, `Dañado`, `Sin Etiqueta`, `Incompleto` |
| **F** | `Observaciones` | Texto | `Verificado en puesto de trabajo` | Novedades adicionales o detalles del cargador |

---

## 4. GUÍA PASO A PASO: MÉTODO 1 — BARCODE TO PC (RECOMENDADO)

### Fase A: Preparación en la Computadora del Auditor
1. Iniciar la laptop de auditoría y abrir `conciliador_universal.py`.
2. En la pestaña **Módulo QR**, hacer clic en **"Exportar Base con IDs a Excel"**. Se generará el libro `etiquetas_qr/Inventario_Equipos_Clemont_Con_QR.xlsx`.
3. Abrir dicho archivo en Microsoft Excel y posicionar la vista en la hoja **`Escaneos_Campo`**.
4. Hacer clic en la celda **A2** (primera fila vacía de `ID_Escaneado`).
5. Descargar e iniciar la aplicación de escritorio **Barcode to PC Server** (gratuita para Windows).
6. Verificar que la ventana del servidor muestre la dirección IP local de la computadora (ej: `192.168.1.85`) y el código QR de emparejamiento.

### Fase B: Configuración del Teléfono Móvil
1. Conectar el teléfono móvil a la misma red Wi-Fi de la laptop (o a la Zona Wi-Fi / Hotspot emitida por el propio teléfono o laptop si no hay router local).
2. Abrir la app **Barcode to PC: Wi-Fi scanner** (disponible en Google Play Store y Apple App Store).
3. Presionar el botón de escaneo inicial y apuntar la cámara a la pantalla de la PC para leer el código QR del servidor.
4. El indicador superior en el móvil cambiará a **"Connected"** en color verde.
5. Acceder al menú **Settings (Ajustes) > Output**:
   - Modo de transmisión: **Keystrokes**.
   - Tecla de sufijo (Post-scan key): Seleccionar **`ENTER`**.
   - Habilitar **Continuous Scan Mode** (Modo ráfaga) para escanear múltiples activos de forma continua sin oprimir la pantalla en cada toma.

### Fase C: Ejecución del Escaneo en Sitio
1. Asegurarse de que Microsoft Excel permanezca en primer plano en la PC con el cursor en la columna **A** (`ID_Escaneado`).
2. Acercar la cámara del teléfono a la etiqueta adhesiva fijada en el chasis del equipo (distancia recomendada: 12 a 20 cm).
3. El móvil emitirá una vibración y sonido de confirmación al decodificar la etiqueta.
4. Inmediatamente el valor del activo (ej: `LAP-2026-0001` o el serial de fábrica) se escribirá en la celda activa de Excel y el cursor descenderá automáticamente a la celda inferior (`A3`, `A4`, etc.).
5. Si el auditor se encuentra en un área con baja iluminación (como el Sótano), presionar el icono de la **linterna/flash LED** en la interfaz de la app móvil.

---

## 5. GUÍA PASO A PASO: MÉTODO 2 — MICROSOFT EXCEL MOBILE APP

### Fase A: Preparación en la Nube
1. Cargar el archivo generado `Inventario_Equipos_Clemont_Con_QR.xlsx` en la carpeta institucional de **OneDrive Corporativo** o biblioteca de **SharePoint**.
2. Compartir el enlace con permisos de edición para el auditor de campo.

### Fase B: Operación desde el Celular
1. Abrir la aplicación móvil oficial de **Microsoft Excel** en el smartphone.
2. Iniciar sesión con la cuenta corporativa y abrir el archivo compartido.
3. Desplazarse a la pestaña inferior **`Escaneos_Campo`**.
4. Tocar dos veces la celda de la columna **A** (`ID_Escaneado`) para activar el teclado en pantalla.
5. Oprimir el icono de **Cámara / Lector de Código de Barras** (incorporado en el teclado de Microsoft o teclado compatible con escáner).
6. Apuntar la cámara al código QR de la etiqueta. El valor se insertará directamente en la celda.
7. Escribir la sede encontrada en la columna B (`Ubicacion_Hallada`) y pasar a la siguiente fila.
8. La sincronización en tiempo real actualizará simultáneamente la pantalla del coordinador en la PC central.

---

## 6. REGLAS DE CONCILIACIÓN FÍSICA Y CRITERIOS DE CLASIFICACIÓN

Una vez finalizada la jornada de escaneo, el `Conciliador Universal` cruza las lecturas contra el catálogo maestro y clasifica cada equipo bajo los siguientes criterios técnicos:

```
┌───────────────────────────┬─────────────────────────────────────────────────┐
│ ESTADO DE CONCILIACIÓN    │ CRITERIO TÉCNICO Y ACCIÓN TI                    │
├───────────────────────────┼─────────────────────────────────────────────────┤
│ ✔️ CONCILIADO OK          │ El ID/Serial escaneado existe en el inventario  │
│                           │ oficial y la sede hallada coincide exactamente. │
├───────────────────────────┼─────────────────────────────────────────────────┤
│ ⚠️ TRASLADO NO REGISTRADO │ El equipo fue hallado en una sede distinta a su │
│                           │ asignación oficial (ej: registrado en Sótano,   │
│                           │ encontrado en Showroom). Requiere actualización.│
├───────────────────────────┼─────────────────────────────────────────────────┤
│ ➕ ALTA / SOBRANTE        │ Equipo físico encontrado y escaneado que NO     │
│                           │ figuraba en el catálogo inicial. Se inserta en  │
│                           │ la base con código nuevo para su regularización.│
├───────────────────────────┼─────────────────────────────────────────────────┤
│ ❌ FALTANTE               │ Activo registrado en el catálogo maestro que NO │
│                           │ fue localizado durante el recorrido en campo.   │
└───────────────────────────┴─────────────────────────────────────────────────┘
```

---

## 7. RESOLUCIÓN DE INCIDENCIAS EN CAMPO (TROUBLESHOOTING)

| Síntoma Observado | Causa Probable | Solución Inmediata |
|---|---|---|
| **La app móvil indica "Connection Error" con la PC** | Dispositivos en redes distintas o la red Wi-Fi corporativa tiene aislamiento de clientes (AP Isolation). | 1. Activar la **Zona Wi-Fi (Hotspot)** en el teléfono móvil y conectar la laptop a dicha red.<br>2. Comprobar que el Firewall de Windows permita el puerto TCP 57888. |
| **El escaneo escribe el ID pero no cambia de fila** | Falta la configuración del sufijo `ENTER`. | En la app móvil Barcode to PC, ir a `Settings > Output > Post-scan key` y marcar **`ENTER`**. |
| **Se escriben caracteres con tildes o símbolos raros** | Desalineación de distribución de idioma de teclado. | En Barcode to PC Server (en la PC), seleccionar en `Keyboard Layout`: **Latin America** o **Spanish**. |
| **Etiqueta física rayada o con poca luz** | Desgaste físico del adhesivo o penumbra. | 1. Activar la linterna de la app móvil.<br>2. Si el QR está roto, digitar manualmente el código de activo visible impreso en texto legible (ej: `LAP-2026-0004`). |
| **El escáner lee pero Excel no recibe los caracteres** | Excel perdió el foco de la ventana activa. | Hacer clic con el ratón sobre la celda objetivo en Excel para devolver el foco antes de iniciar la ráfaga de escaneo. |

---

## 8. RECOMENDACIONES DE IMPRESIÓN Y FIJACIÓN DE ETIQUETAS

1. **Resolución de Impresión**: Utilizar exclusivamente los archivos generados a **300 DPI** de la carpeta `etiquetas_qr/` para garantizar la nitidez de los módulos QR y la legibilidad tipográfica.
2. **Material del Adhesivo**: Se recomienda papel sintético o poliéster mate resistente a la abrasión y fricción térmica del chasis de los equipos.
3. **Ubicación en el Activo**:
   - **Laptops**: Cara inferior del chasis, en superficie plana sin obstruir ranuras de ventilación ni tornillos de mantenimiento.
   - **Monitores**: Borde trasero inferior, junto a la etiqueta del fabricante.
   - **PCs de Escritorio**: Lateral superior del chasis o bisel frontal liso.
   - **Tablets**: Carcasa protectora posterior o funda oficial.
