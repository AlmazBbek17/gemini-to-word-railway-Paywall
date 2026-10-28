FROM python:3.12-slim

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

ENV MALLOC_ARENA_MAX=2

CMD ["gunicorn", "app:app", "-c", "gunicorn.conf.py"]
