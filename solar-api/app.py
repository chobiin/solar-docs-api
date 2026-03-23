#!/usr/bin/env python3
     2	# -*- coding: utf-8 -*-
     3	"""
     4	app.py — 태양광 계약서류 자동입력 Flask API 서버 v2.1
     5	━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
     6	Word 템플릿 방식으로 서류 생성 → ZIP 반환
     7	Render.com / Railway 에 그대로 배포 가능
     8	
     9	엔드포인트:
    10	  GET  /api/health              서버 상태 확인 (상세)
    11	  POST /api/generate            단건 생성 → ZIP
    12	  POST /api/generate_batch      엑셀 업로드 일괄 → ZIP
    13	  GET  /api/download_template   엑셀 양식 다운로드
    14	  GET  /api/check_templates     템플릿 파일 상태 확인
    15	  POST /api/debug               디버그 테스트 (템플릿 열기 테스트)
    16	"""
    17	
    18	import io
    19	import os
    20	import re
    21	import json
    22	import zipfile
    23	import tempfile
    24	import logging
    25	import traceback
    26	from datetime import datetime
    27	from pathlib import Path
    28	
    29	from flask import Flask, request, jsonify, send_file, make_response
    30	from flask_cors import CORS
    31	
    32	# ── 로깅 설정 ─────────────────────────────────────────
    33	logging.basicConfig(
    34	    level=logging.INFO,
    35	    format="%(asctime)s [%(levelname)s] %(message)s"
    36	)
    37	log = logging.getLogger(__name__)
    38	
    39	# ── 경로 설정 ─────────────────────────────────────────
    40	BASE_DIR = Path(__file__).parent
    41	TEMPLATE_DIR = BASE_DIR / "templates"
    42	
    43	# ── docx_engine 임포트 ────────────────────────────────
    44	import sys
    45	sys.path.insert(0, str(BASE_DIR))
    46	
    47	ENGINE_LOAD_ERROR = None
    48	generate_all_docs = None
    49	check_templates = None
    50	_normalize_data = None
    51	
    52	try:
    53	    from docx_engine import generate_all_docs, check_templates, _normalize_data
    54	    log.info("✅ docx_engine 로드 성공")
    55	except ImportError as e:
    56	    ENGINE_LOAD_ERROR = f"ImportError: {e}\n{traceback.format_exc()}"
    57	    log.error(f"❌ docx_engine 로드 실패: {e}")
    58	except Exception as e:
    59	    ENGINE_LOAD_ERROR = f"{type(e).__name__}: {e}\n{traceback.format_exc()}"
    60	    log.error(f"❌ docx_engine 로드 실패: {e}")
    61	
    62	# ── python-docx 직접 확인 ────────────────────────────
    63	DOCX_AVAILABLE = False
    64	DOCX_ERROR = None
    65	try:
    66	    from docx import Document
    67	    DOCX_AVAILABLE = True
    68	    log.info("✅ python-docx 사용 가능")
    69	except ImportError as e:
    70	    DOCX_ERROR = str(e)
    71	    log.error(f"❌ python-docx 없음: {e}")
    72	
    73	# ── 엑셀 파싱 (openpyxl) ─────────────────────────────
    74	try:
    75	    import openpyxl
    76	    EXCEL_AVAILABLE = True
    77	except ImportError:
    78	    EXCEL_AVAILABLE = False
    79	    log.warning("openpyxl 없음 — 엑셀 기능 비활성")
    80	
    81	# ── Flask 앱 ──────────────────────────────────────────
    82	app = Flask(__name__)
    83	
    84	# CORS: 모든 도메인 허용 (공개 API)
    85	CORS(app, resources={r"/*": {"origins": "*"}},
    86	     supports_credentials=False,
    87	     allow_headers=["Content-Type", "Accept", "Authorization", "X-Requested-With"],
    88	     methods=["GET", "POST", "PUT", "DELETE", "OPTIONS"])
    89	
    90	app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50MB
    91	
    92	
    93	# ── CORS preflight 완전 처리 ──────────────────────────
    94	@app.after_request
    95	def add_cors_headers(response):
    96	    response.headers["Access-Control-Allow-Origin"] = "*"
    97	    response.headers["Access-Control-Allow-Methods"] = "GET, POST, PUT, DELETE, OPTIONS"
    98	    response.headers["Access-Control-Allow-Headers"] = "Content-Type, Accept, Authorization, X-Requested-With"
    99	    response.headers["Access-Control-Max-Age"] = "86400"
   100	    return response
   101	
   102	@app.route("/api/<path:path>", methods=["OPTIONS"])
   103	@app.route("/", methods=["OPTIONS"])
   104	def handle_options(path=""):
   105	    resp = make_response()
   106	    resp.status_code = 200
   107	    return resp
   108	
   109	
   110	# ════════════════════════════════════════════════════════
   111	#  헬스체크 (상세 버전)
   112	# ════════════════════════════════════════════════════════
   113	
   114	@app.route("/api/health", methods=["GET"])
   115	def health():
   116	    tpl_status = {}
   117	    if generate_all_docs:
   118	        try:
   119	            tpl_status = check_templates()
   120	        except Exception as e:
   121	            tpl_status = {"error": str(e)}
   122	
   123	    # 템플릿 디렉터리 파일 목록
   124	    tpl_files = []
   125	    if TEMPLATE_DIR.exists():
   126	        tpl_files = [f.name for f in TEMPLATE_DIR.iterdir()
   127	                     if f.suffix in ('.docx', '.pdf')]
   128	
   129	    return jsonify({
   130	        "status": "ok",
   131	        "version": "2.1",
   132	        "engine": "docx_word_template",
   133	        "engine_loaded": generate_all_docs is not None,
   134	        "engine_error": ENGINE_LOAD_ERROR,
   135	        "docx_available": DOCX_AVAILABLE,
   136	        "docx_error": DOCX_ERROR,
   137	        "templates": tpl_status,
   138	        "template_files": tpl_files,
   139	        "excel_support": EXCEL_AVAILABLE,
   140	        "python_version": sys.version,
   141	        "timestamp": datetime.now().isoformat()
   142	    })
   143	
   144	
   145	# ════════════════════════════════════════════════════════
   146	#  템플릿 상태 확인
   147	# ════════════════════════════════════════════════════════
   148	
   149	@app.route("/api/check_templates", methods=["GET"])
   150	def api_check_templates():
   151	    result = {
   152	        "template_dir": str(TEMPLATE_DIR),
   153	        "template_dir_exists": TEMPLATE_DIR.exists(),
   154	        "files": [],
   155	        "engine_loaded": generate_all_docs is not None
   156	    }
   157	
   158	    if TEMPLATE_DIR.exists():
   159	        result["files"] = [
   160	            {"name": f.name, "size": f.stat().st_size}
   161	            for f in sorted(TEMPLATE_DIR.iterdir())
   162	        ]
   163	
   164	    if generate_all_docs:
   165	        result["template_check"] = check_templates()
   166	
   167	    return jsonify(result)
   168	
   169	
   170	# ════════════════════════════════════════════════════════
   171	#  디버그 엔드포인트 (문제 진단용)
   172	# ════════════════════════════════════════════════════════
   173	
   174	@app.route("/api/debug", methods=["POST", "GET"])
   175	def api_debug():
   176	    """각 템플릿 파일을 열어보고 오류를 반환합니다."""
   177	    results = {}
   178	
   179	    if not DOCX_AVAILABLE:
   180	        return jsonify({
   181	            "error": "python-docx 미설치",
   182	            "details": DOCX_ERROR
   183	        }), 500
   184	
   185	    from docx import Document
   186	
   187	    for fname in [
   188	        "1_공급계약신고서.docx",
   189	        "2_이용신청서.docx",
   190	        "3_전력공급계약신고서.docx",
   191	        "5_수금용결제계좌신고서.docx",
   192	        "6_전기설비이용계약서.docx",
   193	        "7_준법서약서.docx",
   194	    ]:
   195	        fpath = TEMPLATE_DIR / fname
   196	        if not fpath.exists():
   197	            results[fname] = "❌ 파일 없음"
   198	            continue
   199	        try:
   200	            doc = Document(str(fpath))
   201	            table_count = len(doc.tables)
   202	            para_count = len(doc.paragraphs)
   203	            results[fname] = f"✅ 열기 성공 (표: {table_count}개, 단락: {para_count}개)"
   204	        except Exception as e:
   205	            results[fname] = f"❌ 열기 실패: {type(e).__name__}: {e}"
   206	
   207	    return jsonify({
   208	        "status": "debug_complete",
   209	        "template_dir": str(TEMPLATE_DIR),
   210	        "results": results
   211	    })
   212	
   213	
   214	# ════════════════════════════════════════════════════════
   215	#  단건 생성
   216	# ════════════════════════════════════════════════════════
   217	
   218	@app.route("/api/generate", methods=["POST"])
   219	def api_generate():
   220	    # 엔진 체크
   221	    if not generate_all_docs:
   222	        log.error(f"엔진 미로드: {ENGINE_LOAD_ERROR}")
   223	        return jsonify({
   224	            "error": "서류 생성 엔진 초기화 실패",
   225	            "details": ENGINE_LOAD_ERROR or "알 수 없는 오류",
   226	            "docx_available": DOCX_AVAILABLE
   227	        }), 500
   228	
   229	    # JSON 파싱
   230	    data = request.get_json(force=True, silent=True)
   231	    if not data:
   232	        # form data fallback
   233	        data = {k: v for k, v in request.form.items()}
   234	
   235	    if not data:
   236	        return jsonify({"error": "JSON 데이터가 없습니다."}), 400
   237	
   238	    # 필수 항목 검증
   239	    required = ["상호명", "대표자명"]
   240	    missing = [f for f in required if not str(data.get(f, "")).strip()]
   241	    if missing:
   242	        return jsonify({"error": f"필수 항목 누락: {', '.join(missing)}"}), 400
   243	
   244	    log.info(f"📄 단건 생성 요청: {data.get('상호명')}")
   245	
   246	    try:
   247	        with tempfile.TemporaryDirectory() as tmp:
   248	            success, errors = generate_all_docs(data, tmp)
   249	
   250	            log.info(f"생성 결과: 성공={len(success)}, 오류={len(errors)}")
   251	            if errors:
   252	                log.warning(f"생성 오류들: {errors}")
   253	
   254	            if not success:
   255	                return jsonify({
   256	                    "error": "서류 생성 실패",
   257	                    "details": errors,
   258	                    "hint": "템플릿 파일이 서버에 있는지 확인하세요."
   259	                }), 500
   260	
   261	            # ZIP 묶기
   262	            buf = io.BytesIO()
   263	            with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
   264	                for filepath in success:
   265	                    zf.write(filepath, Path(filepath).name)
   266	            buf.seek(0)
   267	
   268	            # ZIP 내용 확인
   269	            zip_size = buf.getbuffer().nbytes
   270	            log.info(f"ZIP 크기: {zip_size} bytes")
   271	
   272	        safe = re.sub(r'[\\/:*?"<>|]', "_", data.get("상호명", "발전소"))
   273	        fname = f"태양광계약서류_{safe}_{datetime.now().strftime('%Y%m%d')}.zip"
   274	
   275	        log.info(f"✅ 생성 완료: {len(success)}개 파일 → {fname}")
   276	
   277	        return send_file(
   278	            buf,
   279	            mimetype="application/zip",
   280	            as_attachment=True,
   281	            download_name=fname,
   282	        )
   283	
   284	    except Exception as e:
   285	        tb = traceback.format_exc()
   286	        log.error(f"generate 오류: {e}\n{tb}")
   287	        return jsonify({
   288	            "error": "서버 내부 오류",
   289	            "details": str(e),
   290	            "traceback": tb
   291	        }), 500
   292	
   293	
   294	# ════════════════════════════════════════════════════════
   295	#  엑셀 업로드 → 파싱 (미리보기용)
   296	# ════════════════════════════════════════════════════════
   297	
   298	@app.route("/api/upload_excel", methods=["POST"])
   299	def api_upload_excel():
   300	    if not EXCEL_AVAILABLE:
   301	        return jsonify({"error": "openpyxl 미설치"}), 500
   302	    if "file" not in request.files:
   303	        return jsonify({"error": "파일이 없습니다."}), 400
   304	
   305	    f = request.files["file"]
   306	    if not f.filename:
   307	        return jsonify({"error": "파일명이 비어있습니다."}), 400
   308	
   309	    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
   310	        f.save(tmp.name)
   311	        try:
   312	            rows = _load_excel(tmp.name)
   313	        except Exception as e:
   314	            return jsonify({"error": str(e)}), 400
   315	        finally:
   316	            os.unlink(tmp.name)
   317	
   318	    return jsonify({"rows": rows, "count": len(rows)})
   319	
   320	
   321	# ════════════════════════════════════════════════════════
   322	#  엑셀 일괄 생성
   323	# ════════════════════════════════════════════════════════
   324	
   325	@app.route("/api/generate_batch", methods=["POST"])
   326	def api_generate_batch():
   327	    if not generate_all_docs:
   328	        return jsonify({
   329	            "error": "서류 생성 엔진 초기화 실패",
   330	            "details": ENGINE_LOAD_ERROR
   331	        }), 500
   332	    if not EXCEL_AVAILABLE:
   333	        return jsonify({"error": "openpyxl 미설치"}), 500
   334	    if "file" not in request.files:
   335	        return jsonify({"error": "엑셀 파일이 없습니다."}), 400
   336	
   337	    excel_file = request.files["file"]
   338	    selected_str = request.form.get("selected", "")
   339	
   340	    try:
   341	        selected_indices = [int(i) for i in selected_str.split(",")
   342	                            if i.strip().isdigit()]
   343	    except Exception:
   344	        selected_indices = []
   345	
   346	    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
   347	        excel_file.save(tmp.name)
   348	        try:
   349	            rows = _load_excel(tmp.name)
   350	        except Exception as e:
   351	            return jsonify({"error": str(e)}), 400
   352	        finally:
   353	            os.unlink(tmp.name)
   354	
   355	    if not rows:
   356	        return jsonify({"error": "엑셀에 데이터가 없습니다."}), 400
   357	
   358	    if selected_indices:
   359	        rows = [rows[i] for i in selected_indices if 0 <= i < len(rows)]
   360	
   361	    log.info(f"📦 일괄 생성 요청: {len(rows)}건")
   362	
   363	    try:
   364	        buf = io.BytesIO()
   365	        all_errors = []
   366	
   367	        with tempfile.TemporaryDirectory() as batch_tmp:
   368	            with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
   369	                for row in rows:
   370	                    safe = re.sub(r'[\\/:*?"<>|]', "_", row.get("상호명", "발전소"))
   371	                    sub_dir = os.path.join(batch_tmp, safe)
   372	                    success, errors = generate_all_docs(row, sub_dir)
   373	                    all_errors.extend(errors)
   374	                    for fp in success:
   375	                        zf.write(fp, f"{safe}/{Path(fp).name}")
   376	
   377	        buf.seek(0)
   378	        fname = f"태양광계약서류_일괄_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip"
   379	
   380	        log.info(f"✅ 일괄 완료: {len(rows)}건, 오류: {len(all_errors)}건")
   381	
   382	        return send_file(
   383	            buf,
   384	            mimetype="application/zip",
   385	            as_attachment=True,
   386	            download_name=fname,
   387	        )
   388	
   389	    except Exception as e:
   390	        tb = traceback.format_exc()
   391	        log.error(f"batch generate 오류: {e}\n{tb}")
   392	        return jsonify({
   393	            "error": "일괄 생성 중 오류",
   394	            "details": str(e),
   395	            "traceback": tb
   396	        }), 500
   397	
   398	
   399	# ════════════════════════════════════════════════════════
   400	#  엑셀 양식 다운로드
   401	# ════════════════════════════════════════════════════════
   402	
   403	@app.route("/api/download_template", methods=["GET"])
   404	def api_download_template():
   405	    tpl = BASE_DIR / "발전소목록_입력양식.xlsx"
   406	    if tpl.exists():
   407	        return send_file(str(tpl), as_attachment=True,
   408	                         download_name="발전소목록_입력양식.xlsx")
   409	
   410	    if not EXCEL_AVAILABLE:
   411	        return jsonify({"error": "openpyxl 미설치"}), 500
   412	
   413	    try:
   414	        buf = _create_excel_template()
   415	        return send_file(
   416	            buf,
   417	            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
   418	            as_attachment=True,
   419	            download_name="발전소목록_입력양식.xlsx",
   420	        )
   421	    except Exception as e:
   422	        return jsonify({"error": str(e)}), 500
   423	
   424	
   425	# ════════════════════════════════════════════════════════
   426	#  내부 헬퍼
   427	# ════════════════════════════════════════════════════════
   428	
   429	HEADER_ALIASES = {
   430	    "상호명":           ["상호명", "회사명", "사업장명", "법인명"],
   431	    "대표자명":         ["대표자명", "대표자", "대표"],
   432	    "사업자등록번호":   ["사업자등록번호", "사업자번호", "사업자No"],
   433	    "법인등록번호":     ["법인등록번호", "법인번호"],
   434	    "사업장주소":       ["사업장주소", "주소", "사업장 주소"],
   435	    "발전기주소":       ["발전기주소", "발전기 주소", "발전기 설치주소",
   436	                        "발전기 설치주소 ★", "설치주소", "태양광주소"],
   437	    "설비용량":         ["설비용량", "용량(kW)", "용량", "kW",
   438	                        "설비용량(kW)", "설비 용량"],
   439	    "전기사업자등록번호": ["전기사업자등록번호", "전기사업자번호"],
   440	    "연락처":           ["연락처", "전화번호", "전화", "연락전화", "담당자전화"],
   441	    "이메일":           ["이메일", "email", "Email", "E-mail", "메일"],
   442	    "예금주명":         ["예금주명", "예금주"],
   443	    "은행명":           ["은행명", "은행"],
   444	    "계좌번호":         ["계좌번호", "계좌"],
   445	}
   446	
   447	
   448	def _load_excel(path: str) -> list:
   449	    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
   450	    ws = wb.active
   451	
   452	    header_row_idx = None
   453	    col_map = {}
   454	    for ri, row in enumerate(ws.iter_rows(min_row=1, max_row=5, values_only=True)):
   455	        for ci, cell in enumerate(row):
   456	            if cell is None:
   457	                continue
   458	            cell_str = str(cell).strip()
   459	            for key, aliases in HEADER_ALIASES.items():
   460	                if cell_str in aliases and key not in col_map:
   461	                    col_map[key] = ci
   462	                    header_row_idx = ri
   463	    wb.close()
   464	
   465	    if header_row_idx is None:
   466	        raise ValueError("헤더 행을 찾을 수 없습니다. 양식을 확인해주세요.")
   467	
   468	    wb2 = openpyxl.load_workbook(path, read_only=True, data_only=True)
   469	    ws2 = wb2.active
   470	    rows = []
   471	
   472	    for ri, row in enumerate(ws2.iter_rows(min_row=header_row_idx + 2,
   473	                                            values_only=True)):
   474	        if all(v is None or str(v).strip() == "" for v in row):
   475	            continue
   476	        item = {}
   477	        for key, ci in col_map.items():
   478	            if ci < len(row) and row[ci] is not None:
   479	                item[key] = str(row[ci]).strip()
   480	        if item.get("상호명") or item.get("대표자명"):
   481	            rows.append(item)
   482	
   483	    wb2.close()
   484	    return rows
   485	
   486	
   487	def _create_excel_template() -> io.BytesIO:
   488	    wb = openpyxl.Workbook()
   489	    ws = wb.active
   490	    ws.title = "발전소목록"
   491	
   492	    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
   493	    from openpyxl.utils import get_column_letter
   494	
   495	    headers = list(HEADER_ALIASES.keys())
   496	    required = {"상호명", "대표자명", "사업자등록번호", "사업장주소", "설비용량", "연락처"}
   497	
   498	    thin = Side(style="thin")
   499	    border = Border(left=thin, right=thin, top=thin, bottom=thin)
   500	
   501	    for ci, h in enumerate(headers, 1):
   502	        cell = ws.cell(row=1, column=ci, value=h + (" ★" if h in required else ""))
   503	        cell.font = Font(bold=True, size=10,
   504	                         color="FFFFFF" if h in required else "333333")
   505	        cell.fill = PatternFill("solid",
   506	                                fgColor="1E40AF" if h in required else "64748B")
   507	        cell.alignment = Alignment(horizontal="center", vertical="center",
   508	                                   wrap_text=True)
   509	        cell.border = border
   510	        ws.column_dimensions[get_column_letter(ci)].width = max(len(h) * 2.2, 14)
   511	
   512	    example = {
   513	        "상호명": "홍길동태양광발전(주)",
   514	        "대표자명": "홍길동",
   515	        "사업자등록번호": "123-45-67890",
   516	        "법인등록번호": "110111-1234567",
   517	        "사업장주소": "서울특별시 강남구 테헤란로 123",
   518	        "발전기주소": "전라남도 영암군 삼호읍 나불리 100",
   519	        "설비용량": "99.9",
   520	        "전기사업자등록번호": "2023-서울-0001",
   521	        "연락처": "010-1234-5678",
   522	        "이메일": "test@example.com",
   523	        "예금주명": "홍길동태양광발전",
   524	        "은행명": "국민은행",
   525	        "계좌번호": "123456-78-901234",
   526	    }
   527	    for ci, h in enumerate(headers, 1):
   528	        cell = ws.cell(row=2, column=ci, value=example.get(h, ""))
   529	        cell.alignment = Alignment(vertical="center")
   530	        cell.border = border
   531	
   532	    ws.row_dimensions[1].height = 30
   533	    ws.row_dimensions[2].height = 22
   534	    ws.freeze_panes = "A2"
   535	
   536	    buf = io.BytesIO()
   537	    wb.save(buf)
   538	    buf.seek(0)
   539	    return buf
   540	
   541	
   542	# ════════════════════════════════════════════════════════
   543	#  메인
   544	# ════════════════════════════════════════════════════════
   545	
   546	if __name__ == "__main__":
   547	    port = int(os.environ.get("PORT", 5050))
   548	    debug = os.environ.get("DEBUG", "false").lower() == "true"
   549	
   550	    log.info("=" * 55)
   551	    log.info("  ☀  태양광 계약서류 자동입력 API 서버 v2.1")
   552	    log.info(f"  포트: {port}")
   553	    log.info(f"  엔진: Word 템플릿 방식")
   554	    log.info(f"  템플릿: {TEMPLATE_DIR}")
   555	    log.info("=" * 55)
   556	
   557	    if generate_all_docs:
   558	        status = check_templates()
   559	        for name, exists in status.items():
   560	            icon = "✅" if exists else "❌"
   561	            log.info(f"  {icon} {name}")
   562	    else:
   563	        log.warning(f"⚠️  엔진 로드 실패: {ENGINE_LOAD_ERROR}")
   564	
   565	    app.run(host="0.0.0.0", port=port, debug=debug)
