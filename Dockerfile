FROM --platform=linux/amd64 python:3.9-slim

# 必要なパッケージをインストール
RUN apt-get update && apt-get install -y \
    libgdiplus \
    wget \
    unzip \
    default-jdk \
    fonts-ipafont-gothic \
    locales \
    && locale-gen ja_JP.UTF-8 \
    && dpkg-reconfigure locales \
    && wget http://archive.ubuntu.com/ubuntu/pool/main/i/icu/libicu70_70.1-2_amd64.deb \
    && dpkg -i libicu70_70.1-2_amd64.deb \
    && rm -rf /var/lib/apt/lists/* \
    && rm libicu70_70.1-2_amd64.deb

# Set JAVA_HOME environment variable
ENV LANG=ja_JP.UTF-8
ENV LANGUAGE=ja_JP.UTF-8
ENV LC_ALL=ja_JP.UTF-8
ENV JAVA_HOME=/usr/lib/jvm/default-java

# 作業ディレクトリを設定
WORKDIR /app

# 必要なPythonパッケージをインストール
COPY requirements.txt .
RUN pip install --upgrade pip && \
    pip install --no-cache-dir -r requirements.txt \
    && pip list

# Asposeのライセンスファイルをダウンロードして配置（必要な場合）
# RUN wget <ライセンスファイルのURL> -O /app/license.lic

# スクリプトをコピー
COPY aspose_excel_to_pdf.py .
COPY inspector.py .
COPY aspose-cells-25.2.jar .

# 出力ディレクトリを作成
RUN mkdir output

# コンテナ実行時のコマンドを設定
#ENTRYPOINT ["python", "aspose_excel_to_pdf.py"]
