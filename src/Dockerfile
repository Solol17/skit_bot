FROM python:3.11-slim
LABEL authors="SoldatovOA"
ENV PYTHONUNBUFFERED 1

WORKDIR /action/workspace/src

COPY .* .

COPY ../* .
RUN pip install --no-cache-dir -r requirements.txt

# Установка дополнительных зависимостей, если необходимо
RUN apt-get update && \
    apt-get install -y firefox-esr && \
    rm -rf /var/lib/apt/lists/*


