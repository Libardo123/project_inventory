## 2026-11-05T17:30:36Z
You are spec_miner_survey_1, a teamwork_spec_miner subagent.
Your working directory is: c:\Users\Analista-TI\Desktop\entrega inicial\Scrip-Inventario\.agents\teamwork\spec_miner_survey_1\
The authoritative user request is at: c:\Users\Analista-TI\Desktop\entrega inicial\Scrip-Inventario\.agents\teamwork\ORIGINAL_REQUEST.md
You MUST read ORIGINAL_REQUEST.md before starting work.

Mission: Extract technical specifications, best practices, and integration designs for R2, R3, R4, and R5.
Tasks:
1. R2 Specification (Mobile scanning to Excel):
   - Investigate how Barcode to PC or Excel mobile app transmits barcode/QR scans wirelessly (keyboard emulation / Wi-Fi server / direct cloud Excel sync).
   - Define the exact schema and workflow for `Maestro_Activos` and `Escaneos_Campo`.
   - Draft a comprehensive, professional Operational Guide (Guía Operativa) in Markdown detailing step-by-step setup and field operation.
2. R3 Specification (GUI and Usability):
   - Define exact UI layout and Tkinter components:
     * Checkboxes (chulos) per column for comparison toggle.
     * Output directory selector (`filedialog.askdirectory`).
     * Direct 1-click buttons: open audited Excel (`os.startfile`), open output folder (`os.startfile` or `explorer`).
     * Wide log console with functional `tk.Scrollbar` and autoscroll.
3. R4 Specification (Libraries and proven code patterns):
   - QR code generation at 300 DPI: `qrcode` package with PIL/Pillow image scaling and DPI metadata.
   - Physical sticker layout: standard asset tag layout (e.g. 50mm x 25mm or standard label proportions at 300 DPI) containing the 7 fields and clear typography.
   - Excel formatting: openpyxl styling, autofit column widths, color coding for reconciliation status.
4. R5 Specification (Automated QA Test Suite):
   - Design the QA test suite architecture: automated test script that runs against real Excel files, checks data integrity, verifies QR readability via `pyzbar` / `cv2` or PIL, checks sheet schema, verifies 0 exit code.
5. Do NOT modify any code in the workspace.
6. Write your detailed report and handoff to `c:\Users\Analista-TI\Desktop\entrega inicial\Scrip-Inventario\.agents\teamwork\spec_miner_survey_1\handoff.md`.
7. Send a message back to the orchestrator when complete.
