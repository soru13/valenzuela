#!/bin/sh
FROM nickgryg/alpine-pandas

ENV PYTHONUNBUFFERED 1
LABEL maintainer="jesus.ed13@gmail.com"
RUN apk update \
    # psycopg2 dependencies
    && apk add --virtual build-deps gcc python3-dev musl-dev \
    && apk add postgresql-dev \
    # Pillow dependencies
    && apk add jpeg-dev zlib-dev freetype-dev lcms2-dev openjpeg-dev tiff-dev tk-dev tcl-dev \
    # CFFI dependencies
    && apk add libffi-dev py-cffi \
    # Translations dependencies
    && apk add gettext \
    # https://docs.djangoproject.com/en/dev/ref/django-admin/#dbshell
    && apk add postgresql-client

# Requirements are installed here to ensure they will be cached.
RUN pip install psycopg2 && pip install ipdb

WORKDIR /app
COPY ./ .

CMD [ "python", "./insert_programas.py"]