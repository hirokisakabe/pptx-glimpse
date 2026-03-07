FROM ubuntu:24.04

ENV DEBIAN_FRONTEND=noninteractive
ENV LANG=C.UTF-8

RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice-impress \
    libreoffice-common \
    fonts-noto \
    fonts-noto-cjk \
    fonts-noto-color-emoji \
    fonts-liberation \
    fonts-crosextra-carlito \
    fonts-crosextra-caladea \
    fonts-croscore \
    ca-certificates curl \
    python3 \
    python3-pip \
    && mkdir -p /usr/share/fonts/truetype/noto-sans-jp \
    && curl -fsSL -o /usr/share/fonts/truetype/noto-sans-jp/NotoSansJP.ttf \
       'https://github.com/google/fonts/raw/295d98a7a0c17c68f1341eaeea354e7960ea70d3/ofl/notosansjp/NotoSansJP%5Bwght%5D.ttf' \
    && echo 'c2f3b4d463500a2ddcd3849cded1fceeb9fd6d1c32e6cbecd568453ba50fc68f  /usr/share/fonts/truetype/noto-sans-jp/NotoSansJP.ttf' | sha256sum -c - \
    && apt-get purge -y curl \
    && apt-get autoremove -y \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt /tmp/
RUN pip3 install --break-system-packages --no-cache-dir -r /tmp/requirements.txt

WORKDIR /workspace
