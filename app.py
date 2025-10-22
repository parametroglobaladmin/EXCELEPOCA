import io, os, csv, base64
import pandas as pd
from flask import Flask, request, Response
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

app = Flask(__name__)

HTML_PAGE = r"""<!doctype html><html lang='pt'><head><meta charset='utf-8'/><meta name='viewport' content='width=device-width, initial-scale=1'/><title>Excel -&gt; Odoo CSV</title></head><body><h1>Excel -&gt; Odoo CSV</h1><form action='/convert' method='post' enctype='multipart/form-data'><input name='file' type='file' accept='.xlsx' required /><button type='submit'>Converter</button></form></body></html>"""

def excel_to_csv_bytes(xlsx_bytes: bytes, sheet_name: str | None, delimiter: str, quote_all: bool, keep_headers: bool) -> bytes:
    wb = load_workbook(io.BytesIO(xlsx_bytes), data_only=True)
    ws = wb[sheet_name] if sheet_name and sheet_name in wb.sheetnames else wb.active

    image_map = {}
    for image in getattr(ws, "_images", []):
        try:
            if hasattr(image.anchor, "_from"):
                col = image.anchor._from.col + 1
                row = image.anchor._from.row + 1
                coord = ws.cell(row=row, column=col).coordinate
                img_bytes = image._data()
                image_map[coord] = base64.b64encode(img_bytes).decode("utf-8")
        except Exception:
            continue

    data = list(ws.values)
    if not data or not data[0]:
        raise ValueError("Folha vazia ou sem cabeçalhos.")

    columns = data[0]
    rows = data[1:] if len(data) > 1 else []
    df = pd.DataFrame(rows, columns=columns)

    if "IMAGE" in df.columns:
        col_idx = list(df.columns).index("IMAGE") + 1
        col_letter = get_column_letter(col_idx)
        for i in range(len(df)):
            cell_ref = f"{col_letter}{i + 2}"
            if cell_ref in image_map:
                df.at[i, "IMAGE"] = image_map[cell_ref]

    import io as _io
    buf = _io.StringIO()
    quoting = csv.QUOTE_ALL if quote_all else csv.QUOTE_MINIMAL
    df.to_csv(buf, index=False, header=keep_headers, sep=delimiter, quoting=quoting, encoding="utf-8")
    return buf.getvalue().encode("utf-8")

@app.get("/")
def index():
    return Response(HTML_PAGE, mimetype="text/html; charset=utf-8")

@app.post("/convert")
def convert():
    if "file" not in request.files:
        return ("Ficheiro não enviado.", 400)

    f = request.files["file"]
    xlsx = f.read()

    delimiter = request.form.get("delimiter", ";")
    quote_all = request.form.get("quoteall", "1") == "1"
    keep_headers = request.form.get("keep_headers", "1") == "1"
    sheet_name = request.form.get("sheet") or None

    try:
        csv_bytes = excel_to_csv_bytes(xlsx, sheet_name, delimiter, quote_all, keep_headers)
    except Exception as e:
        return ("Erro a converter: " + str(e), 400)

    out_name = (os.path.splitext(f.filename or "export.xlsx")[0] + ".csv").replace('"','')
    return Response(
        csv_bytes,
        headers={"Content-Type": "text/csv; charset=utf-8", "Content-Disposition": "attachment; filename=" + out_name},
        status=200,
    )

if __name__ == "__main__":
    port = int(os.environ.get("PORT", "8000"))
    app.run(host="0.0.0.0", port=port, debug=False)
