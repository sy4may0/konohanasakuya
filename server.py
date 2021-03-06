#coding: UTF-8
import json
import dotenv
import os
import base64
import urllib.parse
from pymongo import MongoClient
from flask import Flask, make_response, request, jsonify
from flask_cors import CORS 
from waitress import serve

from WeeklyReport import WeeklyReport

dotenv.load_dotenv()

DISTRIBUTE_REPO = os.path.dirname(os.path.abspath(__file__)) + '/public'
DISTRIBUTE_URL = os.environ.get("URL")
HOST = os.environ.get("HOST")
PORT = os.environ.get("PORT")

app = Flask(__name__)
CORS(app)

if not os.path.isdir(DISTRIBUTE_REPO):
    os.mkdir(DISTRIBUTE_REPO)

mongoClient = MongoClient(os.environ.get('MONGO_URI'))

#
# {
#   'user': '唐澤貴洋',
#   'filename': '唐澤貴洋.xlsx',
#   'achievements': [{}...]
# }
#
@app.route('/', methods=['POST'])
def build_excel():
    data = request.json

    # Get Username text.
    db = mongoClient['oshihomimi']
    col = db.get_collection('users')
    userData = col.find_one(filter={'name': data['user']})
    userNameText = userData['text']

    filename = base64.urlsafe_b64encode(
                    data['filename'].encode('utf-8')
               ).decode()
    distribute_path = os.path.join(DISTRIBUTE_REPO, filename)

    weeklyReport = WeeklyReport(userNameText, data['achievements'])
    weeklyReport.writeToExcel(distribute_path)

    download_link = os.path.join(DISTRIBUTE_URL, filename)

    return jsonify({ 'download_link': download_link })

@app.route('/<string:filename>', methods=['GET'])
def download_excel(filename):
    response = make_response()
    distribute_path = os.path.join(DISTRIBUTE_REPO, filename)
    response.data = open(distribute_path, 'rb').read()

    downloadFileName = base64.urlsafe_b64decode(filename).decode('utf-8')
    response.headers['Content-Disposition'] = \
            "attachment;filename*=utf-8''" + \
            urllib.parse.quote(downloadFileName)

    response.mimetype = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    return response

if __name__ == '__main__':
    serve(app, host=HOST, port=PORT)
