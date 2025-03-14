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
    <title>Ataskaitų kepėja Gertrūda</title>
  </head>
  <body>
    <h1>Įkelkite duomenis iš VDA sistemos excel formatu</h1>
    <form action="/" method="post" enctype="multipart/form-data">
      <input type="file" name="source_file" accept=".xlsx" required>
      <br><br>
      <input type="submit" value="Atsisiųsti ataskaitą">
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
            values1 = [source_ws[f"D{row}"].value for row in range(2, 9)]
            values8 = [source_ws[f"D{row}"].value for row in [11, 12, 10]] # 8
            values10 = [source_ws[f"D{row}"].value for row in [16, 17, 15]] # 10
            values12 = [source_ws[f"D{row}"].value for row in [22, 23, 21]] # 12
            values14 = [source_ws[f"D{row}"].value for row in [28, 29, 27]] # 14
            values16 = [source_ws[f"D{row}"].value for row in [34, 35, 33]] # 16
            values18 = [source_ws[f"D{row}"].value for row in [40, 41, 39]] # 18

            values11_1 = [source_ws[f"G{row}"].value for row in [19, 20, 18]] 
            values11_2 = [source_ws[f"H{row}"].value for row in [19, 20, 18]] 
            values11 = [f"{g}-{h}" for g, h in zip(values11_1, values11_2)]

            values13_1 = [source_ws[f"G{row}"].value for row in [25, 26, 24]] 
            values13_2 = [source_ws[f"H{row}"].value for row in [25, 26, 24]] 
            values13 = [f"{g}-{h}" for g, h in zip(values13_1, values13_2)]

            values15_1 = [source_ws[f"G{row}"].value for row in [31, 32, 30]] 
            values15_2 = [source_ws[f"H{row}"].value for row in [31, 32, 30]] 
            values15 = [f"{g}-{h}" for g, h in zip(values15_1, values15_2)]

            values17_1 = [source_ws[f"G{row}"].value for row in [37, 38, 36]] 
            values17_2 = [source_ws[f"H{row}"].value for row in [37, 38, 36]] 
            values17 = [f"{g}-{h}" for g, h in zip(values17_1, values17_2)]

            values19_1 = [source_ws[f"G{row}"].value for row in [43, 44, 42]] 
            values19_2 = [source_ws[f"H{row}"].value for row in [43, 44, 42]] 
            values19 = [f"{g}-{h}" for g, h in zip(values19_1, values19_2)]

            values21_1 = [source_ws[f"G{row}"].value for row in [49, 50, 48]] 
            values21_2 = [source_ws[f"H{row}"].value for row in [49, 50, 48]] 
            values21 = [f"{g}-{h}" for g, h in zip(values21_1, values21_2)]

            values22_1 = [source_ws[f"G{row}"].value for row in [49, 50, 48]] 
            values22_2 = [source_ws[f"H{row}"].value for row in [49, 50, 48]] 
            values22 = [f"{g}-{h}" for g, h in zip(values22_1, values22_2)]

            values23_1 = [source_ws[f"G{row}"].value for row in [49, 50, 48]] 
            values23_2 = [source_ws[f"H{row}"].value for row in [49, 50, 48]] 
            values23 = [f"{g}-{h}" for g, h in zip(values23_1, values23_2)]

            values24_1 = [source_ws[f"G{row}"].value for row in [49, 50, 48]] 
            values24_2 = [source_ws[f"H{row}"].value for row in [49, 50, 48]] 
            values24 = [f"{g}-{h}" for g, h in zip(values24_1, values24_2)]

            values25_1 = [source_ws[f"G{row}"].value for row in [49, 50, 48]] 
            values25_2 = [source_ws[f"H{row}"].value for row in [49, 50, 48]] 
            values25 = [f"{g}-{h}" for g, h in zip(values25_1, values25_2)]

            values27_1 = [source_ws[f"G{row}"].value for row in [49, 50, 48]] 
            values27_2 = [source_ws[f"H{row}"].value for row in [49, 50, 48]] 
            values27 = [f"{g}-{h}" for g, h in zip(values27_1, values27_2)]

            values28_1 = [source_ws[f"G{row}"].value for row in [49, 50, 48]] 
            values28_2 = [source_ws[f"H{row}"].value for row in [49, 50, 48]] 
            values28 = [f"{g}-{h}" for g, h in zip(values28_1, values28_2)]

            values29_1 = [source_ws[f"G{row}"].value for row in [49, 50, 48]] 
            values29_2 = [source_ws[f"H{row}"].value for row in [49, 50, 48]] 
            values29 = [f"{g}-{h}" for g, h in zip(values29_1, values29_2)]

            values30_1 = [source_ws[f"G{row}"].value for row in [49, 50, 48]] 
            values30_2 = [source_ws[f"H{row}"].value for row in [49, 50, 48]] 
            values30 = [f"{g}-{h}" for g, h in zip(values30_1, values30_2)]


        except Exception as e:
            return f":( Nepavyko nuskaityti failo: {e}", 500

        # Load a template workbook (assumes 'tuscias.xlsx' is present in the container)
        try:
            template_wb = load_workbook("tuscias.xlsx")
            template_ws = template_wb.active

            for index, value in enumerate(values1):
                template_ws.cell(row=39, column= 1 + index).value = value # target_row =39 start_col = 1  # Column A is 1

            for index, value in enumerate(values8):
              col = 2 
              row = 48 + index
              for merged_range in template_ws.merged_cells.ranges:
                if template_ws.cell(row=row, column=col).coordinate in merged_range:
                  template_ws.unmerge_cells(str(merged_range))  # Unmerge before writing
                  break

              template_ws.cell(row=row, column= col).value = value # target_row =48 start_col = 2  # Column A is 1
              #if index == 0:  # Example: Re-merge for first value
              template_ws.merge_cells(start_row=row, start_column=col, end_row=row, end_column=col+1)

              for index, value in enumerate(values10):
                col = 6 
                row = 48 + index
                template_ws.cell(row=row, column= col).value = value 

              for index, value in enumerate(values12):
                col = 8 
                row = 48 + index
                template_ws.cell(row=row, column= col).value = value 
              
              for index, value in enumerate(values14):
                col = 10 
                row = 48 + index
                template_ws.cell(row=row, column= col).value = value

              for index, value in enumerate(values16):
                col = 12 
                row = 48 + index
                template_ws.cell(row=row, column= col).value = value

              for index, value in enumerate(values18):
                col = 14 
                row = 48 + index
                template_ws.cell(row=row, column= col).value = value

            # nelyginiai

              for index, value in enumerate(values11):
                col = 7 
                row = 48 + index
                template_ws.cell(row=row, column= col).value = value

              for index, value in enumerate(values13):
                col = 9 
                row = 48 + index
                template_ws.cell(row=row, column= col).value = value

              for index, value in enumerate(values15):
                col = 11 
                row = 48 + index
                template_ws.cell(row=row, column= col).value = value

              for index, value in enumerate(values17):
                col = 13 
                row = 48 + index
                template_ws.cell(row=row, column= col).value = value

              for index, value in enumerate(values19):
                col = 15 
                row = 48 + index
                template_ws.cell(row=row, column= col).value = value
                

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
