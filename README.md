# EOB Monday Webhook (Local)

## Install (global is fine)
python -m pip install -r requirements.txt

## Configure
Copy .env.example -> .env and fill in tokens + board id.

## Run
python app.py

Server runs on:
http://127.0.0.1:8000

## Expose to Monday
ngrok http 8000

Set Monday automation webhook URL to:
https://<ngrok-domain>/monday/webhook/export-eob

## Quick handshake test
curl -X POST http://127.0.0.1:8000/monday/webhook/export-eob \
  -H "Content-Type: application/json" \
  -d '{"challenge":"abc"}'
