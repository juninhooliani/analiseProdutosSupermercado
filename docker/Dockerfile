# Ambiente de coleta de notas fiscais - projeto Supermercado
# Ver wiki/08_AMBIENTE_DOCKER.md
FROM python:3.12-slim

# tzdata: datas em ISO-8601 no fuso local (wiki/06, secao 2)
ENV TZ=America/Sao_Paulo \
    PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1

RUN apt-get update \
 && apt-get install -y --no-install-recommends tzdata ca-certificates curl \
 && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY docker/requirements.txt /tmp/requirements.txt
RUN pip install --no-cache-dir -r /tmp/requirements.txt

# Usuario nao-root. O UID 1000 casa com o padrao do Docker Desktop no Windows.
RUN useradd --create-home --uid 1000 coletor
USER coletor

CMD ["python", "sonda_nfp.py", "--help"]
