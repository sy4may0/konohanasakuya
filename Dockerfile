FROM alpine:latest

RUN apk update
RUN apk add git python3 py3-pip
RUN pip install python-dotenv flask flask-cors waitress pymongo openpyxl
RUN git clone https://github.com/sy4may0/konohanasakuya.git
ADD .env /konohanasakuya/.env

ENTRYPOINT ["python3", "/konohanasakuya/server.py"]
