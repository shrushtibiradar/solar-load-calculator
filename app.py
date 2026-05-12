from flask import Flask, render_template, request, send_file
import os
from werkzeug.utils import secure_filename
from extractor import extract_bill_data, fill_excel_template

app = Flask(__name__)
UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = "output"
EXCEL_FOLDER = "excel"

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)
os.makedirs(EXCEL_FOLDER, exist_ok=True)

app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER
TEMPLATE_PATH = os.path.join(
    EXCEL_FOLDER,
    "solar_template.xlsx"
)
@app.route("/")
def home():
    return render_template("index.html")
@app.route("/upload", methods=["POST"])
def upload():
    if "bill" not in request.files:
        return "No file uploaded"

    file = request.files["bill"]

    if file.filename == "":
        return "No selected file"
    filename = secure_filename(file.filename)

    filepath = os.path.join(
        app.config["UPLOAD_FOLDER"],
        filename
    )

    file.save(filepath)

    print("\n========== OCR START ==========\n")
    extracted_data = extract_bill_data(filepath)

    print(extracted_data)

    print("\n========== OCR END ==========\n")
    output_excel = os.path.join(
        OUTPUT_FOLDER,
        "filled_bill.xlsx"
    )
    fill_excel_template(
        TEMPLATE_PATH,
        output_excel,
        extracted_data
    )

    print("\nExcel Generated Successfully\n")
    return send_file(
        output_excel,
        as_attachment=True
    )
if __name__ == "__main__":
    app.run(debug=True)