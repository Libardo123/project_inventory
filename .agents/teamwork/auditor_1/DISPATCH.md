## 2026-11-05T18:12:37Z

You are auditor_1, a teamwork_auditor subagent.
Your working directory is: c:\Users\Analista-TI\Desktop\entrega inicial\Scrip-Inventario\.agents\teamwork\auditor_1\
The authoritative user request is at: c:\Users\Analista-TI\Desktop\entrega inicial\Scrip-Inventario\.agents\teamwork\ORIGINAL_REQUEST.md
You MUST read ORIGINAL_REQUEST.md before starting work.
Project plan: c:\Users\Analista-TI\Desktop\entrega inicial\Scrip-Inventario\PROJECT.md
Test readiness report: c:\Users\Analista-TI\Desktop\entrega inicial\Scrip-Inventario\TEST_READY.md

Tasks:
Perform a complete forensic integrity audit of the codebase (`conciliador_universal.py`), tests (`test_qa_inventario.py`), and documentation (`GUIA_ESCANEO_MOVIL.md`):
1. Forensic Integrity Checks:
   - Check for hardcoded test results: Ensure test functions don't just assert `True` or compare against mocked constants without invoking the real logic.
   - Check for dummy or facade implementations: Ensure `extraer_seriales`, `extraer_campos_completos`, `_generar_etiqueta_individual`, and `_qr_exportar_excel` contain genuine algorithm logic that actually reads/writes data, generates QR matrices, builds PIL images, and writes OpenPyXL workbooks.
   - Run `python test_qa_inventario.py` via run_command and verify that all 93 tests genuinely execute against real files without fake mocks.
   - Check for fabrication of verification outputs or cheating.
2. Determine your audit verdict:
   - `CLEAN` if no integrity violations are found.
   - `INTEGRITY VIOLATION` if any cheating, facade implementations, or hardcoded dummy results are detected.
3. Write your complete handoff report with forensic evidence to `c:\Users\Analista-TI\Desktop\entrega inicial\Scrip-Inventario\.agents\teamwork\auditor_1\handoff.md` and send message to orchestrator.
