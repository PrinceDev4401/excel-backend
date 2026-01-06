from flask import Flask, request, send_file, jsonify
import os
from processor import process_excel

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

@app.route("/process", methods=["POST"])
def process():
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    file = request.files["file"]

    input_path = os.path.join(UPLOAD_FOLDER, file.filename)
    output_path = os.path.join(UPLOAD_FOLDER, "GLOWLINK_Repeated_Names.xlsx")

    file.save(input_path)
    process_excel(input_path, output_path)

    return send_file(
        output_path,
        as_attachment=True,
        download_name="GLOWLINK_Repeated_Names.xlsx"
    )

if __name__ == "__main__":
    app.run()
