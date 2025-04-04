# gunicorn.conf.py
bind = "0.0.0.0:$PORT"  # Utilize a variável de ambiente PORT fornecida pelo Render
workers = 3
timeout = 120