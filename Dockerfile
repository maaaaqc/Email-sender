FROM python:3.7-alpine3.7
MAINTAINER <qingchuan.ma@nuance.com>

WORKDIR /cd-summary

RUN pip install poetry
RUN apk upgrade --update && \
    apk add git && \
    rm -rf /var/cache/apk/* /tmp/* /root/.cache

ADD pyproject.toml poetry.lock ./
RUN poetry config settings.virtualenvs.create false && \
    poetry install && \
    rm -r /root/.cache

ADD *.py ./

EXPOSE 7710

CMD ["python", "-v", "server.py"]
