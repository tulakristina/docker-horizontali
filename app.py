from flask import Flask, render_template_string, request, send_file
from openpyxl import load_workbook
import io

app = Flask(__name__)

# A very simple HTML form for uploading a file
HTML_TEMPLATE = """
<!doctype html>
<html lang="en">
  <head>
    <meta charset="utf-8">
    <title>Ataskaitų kepėja</title>
  </head>
  <body>
    <h1>Įkelkite Excel lentelę, kurią atsisiuntėte iš Palantire</h1>
    <form action="/" method="post" enctype="multipart/form-data">
      <input type="file" name="source_file" accept=".xlsx" required>
      <br><br>
      <input type="submit" value="Pateikti ataskaitą">
    </form>
  </body>
</html>
"""

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        if "source_file" not in request.files:
            return "Failas neįkeltas", 400

        source_file = request.files["source_file"]

        # Load uploaded workbook using openpyxl (assumes source file has data in D2:D8)
        try:
            source_wb = load_workbook(source_file, data_only=True)
            source_ws = source_wb.active
            values = [source_ws[f"D{row}"].value for row in range(2, 9)]
        except Exception as e:
            return f":( Nepavyko nuskaityti failo: {e}", 500

        # Load a template workbook (assumes 'tuscias.xlsx' is present in the container)
        try:
            template_wb = load_workbook("tuscias.xlsx")
            template_ws = template_wb.active
            target_row = 39
            start_col = 1  # Column A is 1
            for index, value in enumerate(values):
                template_ws.cell(row=target_row, column=start_col + index).value = value
        except Exception as e:
            return f"Error processing template: {e}", 500

        # Save modified workbook into a BytesIO stream and send it for download
        output = io.BytesIO()
        template_wb.save(output)
        output.seek(0)
        return send_file(output,
                         as_attachment=True,
                         download_name="ataskaita.xlsx",
                         mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    return render_template_string(HTML_TEMPLATE)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8080)
