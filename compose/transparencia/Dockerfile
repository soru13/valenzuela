#!/bin/sh
FROM python:3.8-alpine
ENV PYTHONUNBUFFERED 1

# Requirements are installed here to ensure they will be cached.
RUN pip install xlrd & pip install openpyxl

WORKDIR /app
COPY ./ .

CMD [ "python", "./transparencia.py"]
