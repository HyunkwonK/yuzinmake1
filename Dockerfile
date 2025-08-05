# FROM python:3.11-slim

# # 리눅스 패키지 설치 (Camelot 및 PDF 처리용)
# RUN apt-get update && apt-get install -y \
#     tesseract-ocr \
#     tesseract-ocr-kor \
#     ghostscript \
#     qpdf \
#     python3-tk \
#     libglib2.0-dev \
#     libsm6 \
#     libxext6 \
#     libxrender-dev \
#     && rm -rf /var/lib/apt/lists/*

# # 파이썬 라이브러리 설치
# RUN pip install --no-cache-dir \
#     ocrmypdf \
#     camelot-py[cv] \
#     pandas \
#     openpyxl

# # 작업 디렉토리
# WORKDIR /app

FROM python:3.11-slim

# 리눅스 패키지 설치 (Camelot 및 PDF 처리용)
RUN apt-get update && apt-get install -y \
    tesseract-ocr \
    tesseract-ocr-kor \
    ghostscript \
    qpdf \
    python3-tk \
    libglib2.0-dev \
    libsm6 \
    libxext6 \
    libxrender-dev \
    libgl1-mesa-glx \
    libegl1-mesa \
    libxkbcommon-x11-0 \
    libopencv-dev \
    python3-opencv \
    && rm -rf /var/lib/apt/lists/*

# pip 최신 버전으로 업데이트
RUN pip install --upgrade pip

# 파이썬 라이브러리 설치 (PyQt5 제외)
RUN pip install --no-cache-dir \
    ocrmypdf \
    camelot-py[cv] \
    pandas \
    openpyxl \
    xlsxwriter \
    pdfplumber \
    requests

# 작업 디렉토리
WORKDIR /app