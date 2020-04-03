FROM alpine:latest

RUN apk update
RUN apk add git python3
RUN pip3 install python-dotenv flask flask-cors waitress pymongo openpyxl
RUN git clone https://github.com/sy4may0/konohanasakuya.git
ADD .env /konohanasakuya/.env

ENTRYPOINT ["python3", "/konohanasakuya/server.py"]
