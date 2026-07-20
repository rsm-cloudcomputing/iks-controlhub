release: python manage.py migrate --noinput && python manage.py seed_placeholders
web: gunicorn iks_platform.wsgi --log-file -
