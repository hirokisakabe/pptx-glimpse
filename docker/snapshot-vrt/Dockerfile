FROM node:22-slim

ENV DEBIAN_FRONTEND=noninteractive

RUN apt-get update && apt-get install -y --no-install-recommends \
    fonts-noto \
    fonts-noto-cjk \
    fonts-noto-color-emoji \
    fonts-liberation \
    fonts-crosextra-carlito \
    fonts-crosextra-caladea \
    fonts-croscore \
    ca-certificates curl \
    && mkdir -p /usr/share/fonts/truetype/noto-sans-jp \
    && curl -fsSL -o /usr/share/fonts/truetype/noto-sans-jp/NotoSansJP.ttf \
       'https://github.com/google/fonts/raw/295d98a7a0c17c68f1341eaeea354e7960ea70d3/ofl/notosansjp/NotoSansJP%5Bwght%5D.ttf' \
    && echo 'c2f3b4d463500a2ddcd3849cded1fceeb9fd6d1c32e6cbecd568453ba50fc68f  /usr/share/fonts/truetype/noto-sans-jp/NotoSansJP.ttf' | sha256sum -c - \
    && apt-get purge -y curl \
    && apt-get autoremove -y \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /workspace
