from flask import Flask, send_from_directory, Response, send_file
from Database import generate_sheet
from datetime import datetime
from Path import PathResource
from openpyxl import Workbook
from openpyxl.writer.excel import save_virtual_workbook

app = Flask(__name__)


@app.route('/store')
def get_stores():
    file = generate_sheet(datetime(2022, 11, 24), datetime(2022, 12, 7))
    sheet = PathResource.resource_path(file)
    return send_file(sheet, as_attachment=True)


if __name__ == "__main__":
    app.run(debug=True)
