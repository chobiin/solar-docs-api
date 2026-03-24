#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
app.py — 태양광 계약서류 자동입력 Flask API 서버 v2.1
Word 템플릿 방식으로 서류 생성 → ZIP 반환
Render.com / Railway 에 그대로 배포 가능

엔드포인트:
  GET  /api/health
  POST /api/generate
  POST /api/generate_batch
  GET  /api/download_template
  GET  /api/check_templates
  POST /api/debug
"""

import io
import os
import re
import json
import zipfile
import tempfile
import logging
import traceback
from datetime import datetime
from pathlib import Path

from flask import Flask, request, jsonify, send_file, make_response
from flask_cors import CORS

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s"
)
log = logging.getLogger(__name__)

BASE_DIR = Path(__file__).parent
TEMPLATE_DIR = BASE_DIR / "templates"

import sys
sys.path.insert(0, str(BASE_DIR))

ENGINE_LOAD_ERROR = None
generate_all_docs = None
check_templates = None
_normalize_data = None

try:
    from docx_engine import generate_all_docs, check_templates, _normalize_data
    log.info("✅ docx_engine 로드 성공")
except ImportError as e:
    ENGINE_LOAD_ERROR = f"ImportError: {e}\n{traceback.format_exc()}"
    log.error(f"❌ docx_engine 로드 실패: {e}")
except Exception as e:
    ENGINE_LOAD_ERROR = f"{type(e).__name__}: {e}\n{traceback.format_exc()}"
    log.error(f"❌ docx_engine 로드 실패: {e}")

DOCX_AVAILABLE = False
DOCX_ERROR = None
try:
    from docx import Document
    DOCX_AVAILABLE = True
    log.info("✅ python-docx 사용 가능")
except ImportError as e:
    DOCX_ERROR = str(e)
    log.error(f"❌ python-docx 없음: {e}")

try:
    import openpyxl
    EXCEL_AVAILABLE = True
except ImportError:
    EXCEL_AVAILABLE = False
    log.warning("openpyxl 없음 — 엑셀 기능 비활성")

app = Flask(__name__)

CORS(app, resources={r"/*": {"origins": "*"}},
     supports_credentials=False,
     allow_headers=["Content-Type", "Accept", "Authorization", "X-Requested-With"],
     methods=["GET", "POST", "PUT", "DELETE", "OPTIONS"])

app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024


@app.after_request
def add_cors_headers(response):
    response.headers["Access-Control-Allow-Origin"] = "*"
    response.headers
