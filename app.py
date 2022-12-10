from flask import Flask, send_from_directory, Response, send_file, request
from Database import generate_sheet
from datetime import datetime
from Path import PathResource
from openpyxl import Workbook
from openpyxl.writer.excel import save_virtual_workbook

app = Flask(__name__)


@app.route('/store',methods=['Get'])
def get_stores():
    if request.get_json():
        content = request.get_json()
        s = content['start'].split('-')
        e = content['end'].split('-')
        start = datetime(int(s[0]), int(s[1]),int(s[2]))
        end = datetime(int(e[0]), int(e[1]), int(e[2]))
        file = generate_sheet(start=start, end=end)
    else:
        file = generate_sheet()
    sheet = PathResource.resource_path(file)
    return send_file(sheet, as_attachment=True)


if __name__ == "__main__":
    app.run(debug=True)
