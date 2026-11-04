import os
import re
import tkinter as tk
from tkinter import filedialog, messagebox
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side


def normalizar(texto):
  if not texto:
    return ""
  return re.sub(r"[^A-Za-z0-9]", "", str(texto)).upper()


def similitud_serial(s1, s2):
  """Compara dos seriales tolerando confusiones habituales (O/0, Z/2, S/5)."""
  n1 = normalizar(s1).replace("O", "0").replace("Z", "2").replace("S", "5")
  n2 = normalizar(s2).replace("O", "0").replace("Z", "2").replace("S", "5")
  return n1 == n2 or (len(n1) >= 7 and (n1 in n2 or n2 in n1))


class ConciliadorMaestro:

  def __init__(self, root):
    self.root = root
    self.root.title("Auditoría y Conciliación de Inventario - Clemont")
    self.root.geometry("820x650")

    self.ruta_excel = tk.StringVar()

    # Selección de archivo
    frame_top = tk.LabelFrame(
        root, text="1. Seleccionar Libro de Excel Original", padx=10, pady=10
    )
    frame_top.pack(fill="x", padx=15, pady=10)

    tk.Entry(
        frame_top, textvariable=self.ruta_excel, width=65, state="readonly"
    ).pack(side="left", fill="x", expand=True, padx=(0, 10))
    tk.Button(
        frame_top,
        text="Examinar...",
        command=self.seleccionar_archivo,
        bg="#1F4E78",
        fg="white",
        font=("Arial", 9, "bold"),
    ).pack(side="right")

    # Botón de ejecución
    frame_btn = tk.Frame(root, pady=5)
    frame_btn.pack(fill="x", padx=15)

    tk.Button(
        frame_btn,
        text="▶ EJECUTAR AUDITORÍA Y CONCILIACIÓN COMPLETA",
        command=self.ejecutar_proceso,
        bg="#28A745",
        fg="white",
        font=("Arial", 11, "bold"),
        pady=8,
    ).pack(fill="x")

    # Consola de reporte
    frame_log = tk.LabelFrame(
        root, text="2. Bitácora de Procesamiento en Vivo", padx=10, pady=10
    )
    frame_log.pack(fill="both", expand=True, padx=15, pady=(5, 15))

    self.txt_log = tk.Text(frame_log, bg="#F8F9FA", font=("Consolas", 9))
    self.txt_log.pack(fill="both", expand=True)

  def seleccionar_archivo(self):
    ruta = filedialog.askopenfilename(
        title="Selecciona el archivo Excel original",
        filetypes=[("Archivos de Excel", "*.xlsx")],
    )
    if ruta:
      self.ruta_excel.set(ruta)

  def log(self, msj):
    self.txt_log.insert(tk.END, msj + "\n")
    self.txt_log.see(tk.END)

  def ejecutar_proceso(self):
    ruta = self.ruta_excel.get()
    if not ruta or not os.path.exists(ruta):
      messagebox.showerror(
          "Error", "Selecciona un archivo de Excel válido antes de continuar."
      )
      return

    self.txt_log.delete("1.0", tk.END)
    self.log("Iniciando lectura del libro de Excel...")

    try:
      wb = openpyxl.load_workbook(ruta)
    except Exception as e:
      messagebox.showerror(
          "Error al abrir",
          f"Asegúrate de que el archivo esté CERRADO en Excel.\nDetalle: {e}",
      )
      return

    # Localizar la hoja original
    nombre_hoja = "Inventario de Equipos"
    if nombre_hoja not in wb.sheetnames:
      # Si el usuario ya la renombró o duplicó
      candidatas = [h for h in wb.sheetnames if "INVENTARIO" in h.upper()]
      nombre_hoja = candidatas[0] if candidatas else wb.sheetnames[0]

    ws = wb[nombre_hoja]
    self.log(f"-> Hoja objetivo identificada: [{nombre_hoja}]")

    # Estilos estándar
    thin_border = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9"),
    )
    font_base = Font(name="Calibri", size=10)
    font_bold = Font(name="Calibri", size=10, bold=True)
    font_check = Font(name="Calibri", size=11, bold=True, color="008000")
    font_traslado = Font(name="Calibri", size=10, bold=True, color="C00000")
    font_uncheck = Font(name="Calibri", size=10, italic=True, color="595959")

    # Base de datos de equipos físicos verificados
    auditoria_equipos = [
        # Administrativa 2do piso
        {
            "patron": "T4N0CV16Z13918D",
            "sede": "Administrativa 2do piso",
            "equipo": "ANALISTA-GHA",
            "serial": (
                "Serial equipo: T4N0CV16Z13918D; Modelo: Vivobook ASUS X1502VA"
            ),
            "user": "YASMIN RIVERA",
            "verif": "✔️",
        },
        {
            "patron": "T6NXCV",
            "sede": "Administrativa 2do piso",
            "equipo": "LIDER_GHI",
            "serial": (
                "Serial equipo: T6NXCV053386231; Modelo: ASUS EXPERTBOOK"
                " B1503CVA-S75274X"
            ),
            "user": "NICOLAS",
            "verif": "✔️",
        },
        {
            "patron": "5CD50315P1",
            "sede": "Administrativa 2do piso",
            "equipo": "DESKTOP-KHDG30P",
            "serial": (
                "Serial equipo: 5CD50315P1; Modelo: HP LAPTOP 14-dq0533la"
                " [NOTA: Físicamente en Administrativa 2do piso]"
            ),
            "user": "N/A-ANA MARIA",
            "verif": "✔️",
        },
        {
            "patron": "R9N0CX01W86636H",
            "sede": "Administrativa 2do piso",
            "equipo": "Inteligencia",
            "serial": (
                "Serial equipo: R9N0CX01W86636H; Modelo: Vivobook ASUS X1504ZA"
            ),
            "user": "Inteligencia",
            "verif": "✔️",
        },
        {
            "patron": "R8N0CV02K866319",
            "sede": "Administrativa 2do piso",
            "equipo": "DESKTOP-341T3CV",
            "serial": (
                "Serial equipo: R8N0CV02K866319; Modelo: Vivobook ASUS X515EA"
            ),
            "user": "VALENTINA SALAZAR",
            "verif": "✔️",
        },
        {
            "patron": "R9N0CX01W87936A",
            "sede": "Administrativa 2do piso",
            "equipo": "DESKTOP-FR2U7CB",
            "serial": (
                "Serial equipo: R9N0CX01W87936A; Modelo: Vivobook ASUS X1504ZA"
            ),
            "user": "YESENIA LOPEZ",
            "verif": "✔️",
        },
        {
            "patron": "C30D4F5F",
            "sede": "Administrativa 2do piso",
            "equipo": "DESKTOP-QODII8I",
            "serial": (
                "Serial equipo: TANRCX04W552444; Modelo: ASUS TUF Gaming A15"
                " FA506NCG-HN204 [NOTA: Novedad de Traslado]"
            ),
            "user": "EVENTO",
            "verif": "Traslado",
        },
        # Showroom
        {
            "patron": "R8N0CV02K892318",
            "sede": "Showroom",
            "equipo": "DESKTOP-34IT3CV",
            "serial": (
                "Serial equipo: R8N0CV02K892318; Modelo: VIVOBOOK ASUS X515EA"
            ),
            "user": "LIBARDO BEDOYA",
            "verif": "✔️",
        },
        {
            "patron": "R8NOCV02K825318",
            "sede": "Showroom",
            "equipo": "DESKTOP-2CHSE06",
            "serial": (
                "Serial equipo: R8N0CV02K825318; Modelo: VIVOBOOK ASUS X515EA"
            ),
            "user": "MARIA FERNANDA DUQUE",
            "verif": "✔️",
        },
        {
            "patron": "R4N0CV1115888165",
            "sede": "Showroom",
            "equipo": "ALEXANDERPC",
            "serial": (
                "Serial equipo: R4N0CV1115888165; Modelo: VIVOBOOK GO E1504FA"
            ),
            "user": "ALEXANDER LOPERA",
            "verif": "✔️",
        },
        {
            "patron": "T4N0CV16B145189",
            "sede": "Showroom",
            "equipo": "AUXILIAR-CONTABLE-ERIKA",
            "serial": (
                "Serial equipo: T4N0CV16B145189; Modelo: Vivobook ASUS X1502VA"
            ),
            "user": "ERIKA MARCELA LOPEZ",
            "verif": "✔️",
        },
        {
            "patron": "R8N0CV183189349",
            "sede": "Showroom",
            "equipo": "DESKTOP-TQ9MEAZ",
            "serial": (
                "Serial equipo: R8N0CV183189349; Modelo: VIVOBOOK ASUS X515EA"
            ),
            "user": "DANIEL RESTREPO",
            "verif": "✔️",
        },
        {
            "patron": "T4NOCV16Z02518G",
            "sede": "Showroom",
            "equipo": "DESKTOP-5NSDMM5",
            "serial": (
                "Serial equipo: T4NOCV16Z02518G; Modelo: VIVOBOOK ASUS X150VA"
            ),
            "user": "N/A ANGELA",
            "verif": "✔️",
        },
        # Sótano
        {
            "patron": "TN4NOCV16Z033184",
            "sede": "Sótano",
            "equipo": "SERVICIO-CLIENTE1",
            "serial": (
                "Serial equipo: T4N0CV16Z033184; Modelo: Vivobook ASUS laptop"
                " X1502VA"
            ),
            "user": "ESTEBAN MUÑOZ",
            "verif": "✔️",
        },
        {
            "patron": "T4N0CV162047186",
            "sede": "Sótano",
            "equipo": "LIDER-ECOMMERCE",
            "serial": (
                "Serial equipo: T4N0CV16Z047186; Modelo: Vivobook ASUS laptop"
                " X1502VA"
            ),
            "user": "DANIELA RODRIGUEZ",
            "verif": "✔️",
        },
        {
            "patron": "TN40CV16Y99",
            "sede": "Sótano",
            "equipo": "Personalshopper",
            "serial": (
                "Serial equipo: T4N0CV16Y999184; Modelo: Vivobook ASUS laptop"
                " X1502VA"
            ),
            "user": "PERSONAL SHOPPER",
            "verif": "✔️",
        },
        {
            "patron": "RCNOCV02552949G",
            "sede": "Sótano",
            "equipo": "Personal-shopper1",
            "serial": (
                "Serial equipo: RCN0CV02552949G; Modelo: Vivobook ASUS laptop"
                " X1605VA (Core i9)"
            ),
            "user": "PERSONAL SHOPPER",
            "verif": "✔️",
        },
    ]

    total_row = None
    for r in range(6, ws.max_row + 1):
      val_b = str(ws.cell(r, 2).value or "")
      if "TOTAL" in val_b.upper():
        total_row = r
        break

    if not total_row:
      total_row = ws.max_row + 1

    # Actualizar filas existentes
    procesados = set()
    for r in range(6, total_row):
      obs_celda = str(ws.cell(r, 11).value or "")
      equipo_celda = str(ws.cell(r, 4).value or "")
      coincidencia = None

      for eq in auditoria_equipos:
        if (
            similitud_serial(eq["patron"], obs_celda)
            or eq["patron"] in obs_celda
            or eq["patron"] in equipo_celda
        ):
          coincidencia = eq
          procesados.add(eq["patron"])
          break

      if coincidencia:
        ws.cell(r, 4).value = coincidencia["equipo"]
        ws.cell(r, 9).value = coincidencia["user"]
        ws.cell(r, 11).value = coincidencia["serial"]
        ws.cell(r, 12).value = coincidencia["verif"]
        ws.cell(r, 12).font = (
            font_traslado if coincidencia["verif"] == "Traslado" else font_check
        )
        self.log(
            f"✔ Actualizado: {coincidencia['equipo']} ({coincidencia['user']})"
        )
      else:
        # Si la fila tiene datos pero no se verificó
        if ws.cell(r, 2).value:
          ws.cell(r, 12).value = "Sin chequear"
          ws.cell(r, 12).font = font_uncheck

    # Adicionar computadores físicos nuevos que no tenían fila
    altas_computadores = [
        [
            "CLEMONT",
            "Sótano",
            "HAIBER",
            "WINDOWS 11 PRO",
            "Microsoft",
            "Sistema operativo",
            1,
            "HAIBER MONTAÑA",
            "Activo",
            (
                "Serial equipo: TAN0KD004371447; Modelo: ASUS Vivobook S 16"
                " M3607HA (Ryzen 9, 40GB RAM)"
            ),
            "✔️",
        ],
        [
            "CLEMONT",
            "Showroom",
            "Natalia",
            "WINDOWS 11 PRO",
            "Microsoft",
            "Sistema operativo",
            1,
            "NATALIA",
            "Activo",
            (
                "Serial equipo: R8N0CV02K86131C; Modelo: VIVOBOOK ASUS X515EA"
                " [NOTA: Reubicada en Showroom]"
            ),
            "✔️",
        ],
        [
            "CLEMONT",
            "Showroom",
            'MacBook Pro 14"',
            "MACOS",
            "Apple",
            "Sistema operativo",
            1,
            "SHOWROOM",
            "Activo",
            (
                "Serial equipo: L04XFFHFQ4; Modelo: Apple MacBook Pro 14 A3426"
                " (24GB RAM)"
            ),
            "✔️",
        ],
    ]

    for idx, nuevo in enumerate(altas_computadores):
      insert_r = total_row + idx
      ws.insert_rows(insert_r)
      ws.cell(insert_r, 1).value = insert_r - 5
      for col_idx, val in enumerate(nuevo, start=2):
        cell = ws.cell(insert_r, col_idx)
        cell.value = val
        cell.font = font_base
        cell.border = thin_border
        if col_idx == 12:
          cell.font = font_check
          cell.alignment = Alignment(horizontal="center")
      self.log(f"➕ Adicionado: {nuevo[2]} en {nuevo[1]}")

    # Reajustar total
    nuevo_total_row = total_row + len(altas_computadores)
    ws.cell(
        nuevo_total_row, 8
    ).value = f"=SUM(H6:H{nuevo_total_row - 1})"

    # Generar o actualizar sección de Tablets
    tab_start = nuevo_total_row + 2
    ws.cell(tab_start, 2).value = "TABLETS Y DISPOSITIVOS MÓVILES VERIFICADOS"
    ws.cell(tab_start, 2).font = Font(
        name="Calibri", size=11, bold=True, color="7030A0"
    )

    tab_headers = [
        "Ítem",
        "Cliente",
        "Sede",
        "Tipo",
        "Marca / Modelo",
        "Modelo Detallado",
        "Tipo",
        "Cant",
        "Responsable",
        "Estado",
        "Serial",
        "Verificado",
    ]
    for c, h in enumerate(tab_headers, 1):
      cell = ws.cell(tab_start + 1, c)
      cell.value = h
      cell.font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
      cell.fill = PatternFill(
          start_color="7030A0", end_color="7030A0", fill_type="solid"
      )

    tablets = [
        [
            1,
            "CLEMONT",
            "Sótano",
            "Tablet",
            "Galaxy Tab S7 FE",
            "SM-T733",
            "Móvil",
            1,
            "SÓTANO",
            "Activo",
            "Serial: R52W10BDXPM",
            "✔️",
        ],
        [
            2,
            "CLEMONT",
            "Showroom",
            "Tablet",
            "Galaxy Tab A7 Lite",
            "SM-T220",
            "Móvil",
            1,
            "SHOWROOM",
            "Traslado",
            "Serial: R83W60T62XB",
            "Traslado",
        ],
    ]
    for i, t in enumerate(tablets, start=1):
      r = tab_start + 1 + i
      for c, v in enumerate(t, 1):
        cell = ws.cell(r, c)
        cell.value = v
        cell.font = font_base
        cell.border = thin_border
        if c == 12:
          cell.font = font_check if v == "✔️" else font_traslado

    # Generar o actualizar sección de Monitores
    mon_start = tab_start + len(tablets) + 3
    ws.cell(mon_start, 2).value = (
        "MONITORES Y PANTALLAS VERIFICADAS (PERIFÉRICOS)"
    )
    ws.cell(mon_start, 2).font = Font(
        name="Calibri", size=11, bold=True, color="1F4E78"
    )

    monitores = [
        [
            1,
            "CLEMONT",
            "Administrativa 2do piso",
            'Monitor 27"',
            "LG 27G411A",
            "27G411A-BB.AWPEMVN",
            "Periférico",
            1,
            "Puesto 2do Piso",
            "Activo",
            "Serial: 512NTDVHJ778",
            "✔️",
        ],
        [
            2,
            "CLEMONT",
            "Administrativa 2do piso",
            'Monitor 27"',
            "LG 27U411A",
            "27U411A-BD.CWPPMV",
            "Periférico",
            1,
            "Puesto 2do Piso",
            "Activo",
            "Serial: 601TFCJ1R907",
            "✔️",
        ],
        [
            3,
            "CLEMONT",
            "Administrativa 2do piso",
            'Monitor 27"',
            "LG 27U411A",
            "27U411A-BD.CWPPMV",
            "Periférico",
            1,
            "Puesto 2do Piso",
            "Activo",
            "Serial: 601TFQK1R908",
            "✔️",
        ],
        [
            4,
            "CLEMONT",
            "Administrativa 2do piso",
            'Monitor 27"',
            "LG 27G411A",
            "27G411A-BB.AWPEMVN",
            "Periférico",
            1,
            "Puesto 2do Piso",
            "Activo",
            "Serial: 512NTBKHJ789",
            "✔️",
        ],
        [
            5,
            "CLEMONT",
            "Sótano",
            'Monitor 27"',
            "LG 27U411A",
            "27U411A-BD.CWPPMV",
            "Periférico",
            1,
            "Puesto Sótano",
            "Activo",
            "Serial: 601TFPG1S168",
            "✔️",
        ],
        [
            6,
            "CLEMONT",
            "Showroom",
            'Monitor 24"',
            'Janus 24" IPS FHD',
            "J24100SIPSFHD",
            "Periférico",
            1,
            "Puesto Showroom",
            "Activo",
            "Serial: J24100SIPSFHD24120400962",
            "✔️",
        ],
        [
            7,
            "CLEMONT",
            "Showroom",
            'Monitor 19"',
            "ViewSonic VA1903H",
            "VS16216",
            "Periférico",
            1,
            "Puesto Showroom",
            "Activo",
            "Serial: VR4234721931",
            "✔️",
        ],
        [
            8,
            "CLEMONT",
            "Showroom",
            'Monitor 19"',
            "ViewSonic VA1903H",
            "VS16216",
            "Periférico",
            1,
            "Puesto Showroom",
            "Activo",
            "Serial: VR4234721846",
            "✔️",
        ],
        [
            9,
            "CLEMONT",
            "Showroom",
            'Monitor 24"',
            "Samsung S24C310EAL",
            "LS24C310EALXZL",
            "Periférico",
            1,
            "Puesto Showroom",
            "Activo",
            "Serial: 5MXSH9JX902670K",
            "✔️",
        ],
        [
            10,
            "CLEMONT",
            "Showroom",
            'Monitor 21.5"',
            "HP S3 Pro 322pe",
            "HSD-0185-V",
            "Periférico",
            1,
            "Puesto Showroom",
            "Activo",
            "Serial: 1CR4450DX8",
            "✔️",
        ],
        [
            11,
            "CLEMONT",
            "Showroom",
            'Monitor 27"',
            "LG 27G411A",
            "27G411A-BB.AWPEMVN",
            "Periférico",
            1,
            "Puesto Showroom",
            "Activo",
            "Serial: 512NTHMHJ793",
            "✔️",
        ],
        [
            12,
            "CLEMONT",
            "Showroom",
            'Monitor 27"',
            "LG 27G411A",
            "27G411A-B",
            "Periférico",
            1,
            "Puesto Showroom",
            "Activo",
            "Serial: Foto 16",
            "✔️",
        ],
    ]

    for c, h in enumerate(tab_headers, 1):
      cell = ws.cell(mon_start + 1, c)
      cell.value = h
      cell.font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
      cell.fill = PatternFill(
          start_color="203764", end_color="203764", fill_type="solid"
      )

    for i, m in enumerate(monitores, start=1):
      r = mon_start + 1 + i
      for c, v in enumerate(m, 1):
        cell = ws.cell(r, c)
        cell.value = v
        cell.font = font_base
        cell.border = thin_border
        if c == 12:
          cell.font = font_check

    # Crear o sobreescribir la Bitácora de Novedades
    if "Bitácora de Novedades" in wb.sheetnames:
      wb.remove(wb["Bitácora de Novedades"])
    ws_bit = wb.create_sheet(title="Bitácora de Novedades")

    ws_bit["B2"] = "ACTA FORMAL DE AUDITORÍA Y CONCILIACIÓN FÍSICA"
    ws_bit["B2"].font = Font(name="Calibri", size=13, bold=True, color="1F4E78")

    bit_headers = [
        "Ítem",
        "Tipo de Novedad",
        "Sede Reportada",
        "Ubicación Física Real",
        "Equipo",
        "Serial",
        "Responsable",
        "Detalle / Justificación",
    ]
    for c, h in enumerate(bit_headers, start=2):
      cell = ws_bit.cell(4, c)
      cell.value = h
      cell.font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
      cell.fill = PatternFill(
          start_color="1F4E78", end_color="1F4E78", fill_type="solid"
      )

    novedades_data = [
        [
            1,
            "Equipo Nuevo (Alta)",
            "Sótano",
            "Sótano",
            "Vivobook S 16 M3607HA",
            "TAN0KD004371447",
            "Haiber Montaña",
            "Máquina física de gama alta (Ryzen 9, 40GB RAM) encontrada en"
            " sitio.",
        ],
        [
            2,
            "Depuración Serial Doble",
            "Sótano",
            "Sótano",
            "Vivobook 16 Core i9",
            "RCN0CV02552949G",
            "Personal Shopper",
            "Se separó de la celda compartida y quedó como Personal-shopper1.",
        ],
        [
            3,
            "Corrección Serial Incompleto",
            "Sótano",
            "Sótano",
            "Vivobook 15 X1502VA",
            "T4N0CV16Y999184",
            "Personal Shopper",
            "Figuraba como TN40CV16Y99 a nombre de Haiber; es de Personal"
            " Shopper.",
        ],
        [
            4,
            "Reubicación de Sede",
            "Sótano",
            "Administrativa 2do piso",
            "HP Laptop 14",
            "5CD50315P1",
            "Ana María",
            "Registrado en Sótano como SCD50315P1; verificado operando en el"
            " 2do piso.",
        ],
        [
            5,
            "Reubicación de Sede",
            "Sótano",
            "Showroom",
            "ASUS Vivobook X515EA",
            "R8N0CV02K86131C",
            "Natalia",
            "Registrado en Sótano; verificado operando en Showroom.",
        ],
        [
            6,
            "Equipo Nuevo (Alta)",
            "No registrado",
            "Showroom",
            'MacBook Pro 14"',
            "L04XFFHFQ4",
            "Showroom",
            "Portátil Apple nuevo con 24GB RAM verificado en Showroom.",
        ],
        [
            7,
            "Traslado Físico",
            "Administrativa 2do piso",
            "Showroom / Traslado",
            "ASUS TUF Gaming A15",
            "TANRCX04W552444",
            "Evento",
            "Serial figuraba como C30D4F5F. Identificado serial de fábrica y"
            " marcado Traslado.",
        ],
        [
            8,
            "Alta Periféricos",
            "No registrados",
            "Varias Sedes",
            "Tablets y Monitores",
            "Varios",
            "Puestos de trabajo",
            "Se crearon secciones independientes para 2 tablets y 12 monitores.",
        ],
    ]

    for idx, nov in enumerate(novedades_data, start=5):
      for c, v in enumerate(nov, start=2):
        cell = ws_bit.cell(idx, c)
        cell.value = v
        cell.font = font_base
        cell.border = thin_border

    # Guardar en archivo nuevo para proteger el original
    dir_salida, nombre_orig = os.path.split(ruta)
    nombre_salida = (
        "AUDITADO_" + nombre_orig
        if not nombre_orig.startswith("AUDITADO_")
        else nombre_orig
    )
    ruta_guardado = os.path.join(dir_salida, nombre_salida)

    wb.save(ruta_guardado)
    self.log("\n-----------------------------------------------------")
    self.log("✅ PROCESO COMPLETADO EXITOSAMENTE")
    self.log(f"Archivo guardado como: {nombre_salida}")
    messagebox.showinfo(
        "Éxito", f"Inventario conciliado y guardado como:\n{nombre_salida}"
    )


if __name__ == "__main__":
  app_root = tk.Tk()
  ConciliadorMaestro(app_root)
  app_root.mainloop()