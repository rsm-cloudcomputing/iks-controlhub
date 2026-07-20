FROM python:3.12-slim

# LibreOffice is needed only for the "preview template as PDF" feature.
# It's a sizeable install (~600MB+) -- see README for the lighter alternative
# if you'd rather skip it and only lose that one preview button.
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

RUN python manage.py collectstatic --noinput

EXPOSE 8080
CMD ["sh", "-c", "python manage.py migrate --noinput && python manage.py seed_placeholders && gunicorn iks_platform.wsgi --bind 0.0.0.0:8080"]
