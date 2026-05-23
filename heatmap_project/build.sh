#!/usr/bin/env bash
set -o errexit
pip install --upgrade pip setuptools wheel
pip install pandas==3.0.3 --only-binary=:all:
pip install -r requirements.txt --no-deps
pip install Django==4.2.13 gunicorn==21.2.0 openpyxl==3.1.2 xlsxwriter==3.1.9 whitenoise==6.7.0
python manage.py collectstatic --no-input
