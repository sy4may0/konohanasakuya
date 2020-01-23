#coding: UTF-8
import json
import dotenv
import os
import base64
from flask import Flask, request, jsonify
from flask_cors import CORS 

from WeeklyReport import WeeklyReport

dotenv.load_dotenv()

DISTRIBUTE_REPO = os.environ.get("DISTRIBUTE_REPO")
DISTRIBUTE_URL = os.environ.get("DISTRIBUTE_URL")
PORT = os.environ.get("PORT")

app = Flask(__name__)
CORS(app)

#
# {
#   'user': '唐澤貴洋'
#   'achievements': [{}...]
# }
#
@app.route('/', methods=['POST'])
def build_excel():
    data = request.json

    filename = base64.urlsafe_b64encode(
                    data['user'].encode('utf-8')
               ).decode()
    distribute_path = os.path.join(DISTRIBUTE_REPO, filename)

    weeklyReport = WeeklyReport(data['user'], data['achievements'])
    weeklyReport.writeToExcel(distribute_path)

    download_link = os.path.join(DISTRIBUTE_URL, filename)

    return jsonify({ 'download_link': download_link })

if __name__ == '__main__':
    app.run(debug=True, host='127.0.0.1', port=PORT)