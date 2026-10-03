import sys, openpyxl
sys.stdout.reconfigure(encoding="utf-8")
path = sys.argv[1]
maxr = int(sys.argv[2]) if len(sys.argv) > 2 else 80
wb = openpyxl.load_workbook(path, data_only=True)
wbf = openpyxl.load_workbook(path, data_only=False)
for ws in wb.worksheets:
    wf = wbf[ws.title]
    print(f"===== SHEET '{ws.title}' dims={ws.dimensions} state={ws.sheet_state} max_row={ws.max_row} max_col={ws.max_column}")
    print("MERGED:", [str(r) for r in ws.merged_cells.ranges][:80])
    for r in range(1, min(ws.max_row, maxr) + 1):
        cells = []
        for c in range(1, ws.max_column + 1):
            cell = ws.cell(r, c)
            v = cell.value
            if v is None:
                continue
            f = wf.cell(r, c).value
            extra = ""
            if isinstance(f, str) and f.startswith("="):
                extra = f" [{f[:40]}]"
            ind = cell.alignment.indent if cell.alignment else 0
            if ind:
                extra += f" <ind{ind}>"
            cells.append(f"{cell.coordinate}={v!r}{extra}")
        if cells:
            print(f"R{r}: " + " | ".join(cells))
