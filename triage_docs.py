#!/usr/bin/env python3
"""
triage_docs.py — 로컬 문서(PDF/Word/PPT/Excel) 대량 트리아지.
Data 360 업로드 전, 파일별로
  (1) 권장 파싱방식  (2) 이유  (3) 중복여부  (4) 원본파일  (5) 정상 파일 여부
를 한 개의 CSV로 정리한다.

지원 형식:
  .pdf                  — pypdf로 벡터/이미지/텍스트 신호 분석
  .docx .pptx .xlsx     — OOXML(=ZIP+XML). 표준 라이브러리(zipfile/xml)만으로
                          내부 chart XML·이미지·텍스트를 직접 카운트 (추가 설치 불필요)
  .doc .ppt .xls        — 구형 바이너리. 구조 분석 불가 → "검토"로 분류

── 판정 근거 ──────────────────────────────────────────────
PDF (휴리스틱):
  - vec_ratio    : 벡터 드로잉 연산자 과밀 페이지 비율  ← 벡터 차트의 결정적 신호
  - dense_ratio  : 텍스트 과밀 페이지 비율
  - img_per_page : 페이지당 래스터 이미지 수 (로고/장식에 오염되므로 단독 신뢰 X)
  → 벡터/텍스트 과밀 = Docling이 차트를 파편화 → "LLM Parser".
     이미지 거의 없는 순수 텍스트 → "Default"(내장 텍스트 추출로 충분).
     이미지 섞인 텍스트+과밀 낮음 → "Docling". 텍스트·이미지 둘다 희박 → "검토"(스캔본/빈).

Data 360 Search Index Builder의 파서 3종에 맞춘 매핑 (Parsing 단계 UI):
  - Default     "Extract text with built-in settings"          = 텍스트만. 표·이미지 없는 순수 텍스트 문서
  - Docling     "Extracts text and tables (open-source models)" = 텍스트 + 표. 표·수치가 있는 문서
  - LLM-based   "Extract text, images, and visual elements"     = 이미지/벡터차트/시각요소 위주

Office (구조가 명확 → 사실 기반):
  OOXML은 차트가 charts/chart*.xml, 표가 <w:tbl>/<a:tbl> 로 "구조적으로" 박혀 있다.
  - charts : ppt|word|xl/charts/chart*.xml 개수 = 네이티브 차트 수
  - media  : ppt|word|xl/media/* 개수 = 삽입 이미지 수
  - tables : <w:tbl>(docx)/<a:tbl>(pptx) 개수 = 본문 표 수
  - text   : 슬라이드/문서 XML의 텍스트런 길이
  판정:
    - 텍스트 있음 + 표·차트·이미지 존재            → "Docling" (표 구조 파싱 필요)
    - 텍스트 있음 + 표·차트·이미지 전무            → "Default" (내장 텍스트 추출로 충분)
    - 텍스트 거의 없고 이미지 위주(슬라이드형/캡처)  → "LLM Parser" (이미지→의미 재구성)
    - 텍스트·차트·이미지 모두 희박(빈 파일)          → "검토"
  XLSX는 본질이 표/수치 → 항상 "Docling"(표는 구조 파싱이 정답), 단 빈 시트는 "검토".

중복여부: 파일 바이트 SHA-256이 같은 파일끼리 그룹. 파일명순 첫 파일 "원본", 나머지 "중복",
          유일 "고유". 중복 행에는 "원본파일" 컬럼에 동일 내용 원본 파일명 표시.

정상여부: 열리고 내용이 있으면 "정상", 아니면 "손상"(+원인).

사용법:
  python3 triage_docs.py <파일_또는_폴더> [...] [--csv 결과.csv] [--jobs 8] [--recursive]
  python3 triage_docs.py ~/docs --csv ~/Downloads/triage.csv
  python3 triage_docs.py ~/docs --xlsx            # 엑셀(요약+상세 2시트)로 저장
  python3 triage_docs.py ~/docs --csv 결과.xlsx    # 확장자가 .xlsx면 자동으로 엑셀

출력:
  기본은 상세 표 한 장(CSV). --xlsx(또는 .xlsx 경로)면 '요약'·'상세' 두 시트의 엑셀 —
  요약 시트에는 콘솔에 찍히는 유형별 분포·파서별 집계·대용량/손상 목록이 그대로 담긴다.
  엑셀 생성은 외부 라이브러리 없이 표준 라이브러리(zipfile)만으로 처리한다.

상세 필드: 번호, 파일명, 경로, 형식, 권장_파싱방식, 이유, 용량, 용량_상태, 용량_추천,
          변환대상, 이름검사, 중복여부, 원본파일, 정상여부
"""
import sys, os, glob, csv, argparse, re, hashlib, zipfile, shutil, unicodedata
from concurrent.futures import ProcessPoolExecutor, as_completed

try:
    from pypdf import PdfReader
except ImportError:
    PdfReader = None  # PDF가 없으면 없어도 됨. PDF 만나면 그때 안내.

# ---- 튜닝 파라미터 (문서군에 맞게 보정) --------------------------------------
HEAVY = 900           # PDF: 이 문자수 초과 = "텍스트 과밀" 페이지
VEC_HEAVY = 300       # PDF: 페이지 path 연산자 이 수 초과 = "벡터 과밀"
SAMPLE_PAGES = 40     # PDF: 페이지 많으면 앞뒤 표본만 검사 (속도)
OFFICE_TEXT_MIN = 200 # Office: 이 문자수 미만이면 "텍스트 실질 없음"으로 간주
# Default 파서(텍스트만) vs Docling(텍스트+표) 분기 임계
PDF_TEXT_ONLY_IMG = 0.15  # PDF: 페이지당 이미지 이 값 미만이면 "이미지 없는 텍스트 문서"로 봄
LARGE_BYTES = 75 * 1024 * 1024        # 경고선: 이 크기 초과면 "대용량"(파싱/임베딩 부담) — 축소 권장
OVERSIZE_BYTES = 2 * 1024 * 1024 * 1024  # 하드 한계: Market Insight Uploader 2GB/file → 이 초과면 "업로드 불가"
HASH_CHUNK = 1 << 20  # 해시 읽기 블록 (1MB)
VEC_OP = re.compile(rb'(?:^|\s)(?:m|l|c|v|y|re|f|F|f\*|B|B\*|b|b\*|S|s|W|W\*)(?=\s)')

# 파일명/경로 유효성 (데이터 마이그레이션 중 저장 실패 예방)
NAME_MAX_LEN = 120   # 파일명(확장자 포함) 이 길이 이상이면 경고
PATH_MAX_LEN = 255   # 전체 경로 이 길이 이상이면 경고
BAD_NAME_CHARS = re.compile(r'[\\/:*?"<>|]|[\x00-\x1f]')  # 윈도우/시스템 금지문자 + 제어문자

EXT_PDF   = {".pdf"}
EXT_OOXML = {".docx", ".pptx", ".xlsx"}
EXT_LEGACY = {".doc", ".ppt", ".xls"}
ALL_EXTS = EXT_PDF | EXT_OOXML | EXT_LEGACY
# ----------------------------------------------------------------------------


def file_sha256(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for blk in iter(lambda: f.read(HASH_CHUNK), b""):
            h.update(blk)
    return h.hexdigest()


def check_filename(name, full_path):
    """파일명/경로 유효성 검사. 위반 사유를 '; '로 이어 반환 (없으면 "").
    데이터 마이그레이션 중 저장 실패를 예방하기 위한 사전 경고 — 실제 파일은 건드리지 않음."""
    issues = []
    if len(name) >= NAME_MAX_LEN:
        issues.append(f"이름 {NAME_MAX_LEN}자 이상({len(name)})")
    bad = BAD_NAME_CHARS.findall(name)
    if bad:
        shown = "".join(sorted(set(c for c in bad if c.isprintable())))
        issues.append("특수문자 포함" + (f"({shown})" if shown else "(제어문자)"))
    if len(full_path) >= PATH_MAX_LEN:
        issues.append(f"경로 {PATH_MAX_LEN}자 이상({len(full_path)})")
    return "; ".join(issues)


# ============================== PDF ==========================================
def count_page_images(page):
    n = 0
    try:
        res = page.get("/Resources")
        if not res:
            return 0
        xo = res.get_object().get("/XObject")
        if not xo:
            return 0
        xo = xo.get_object()
        for k in xo:
            try:
                if xo[k].get_object().get("/Subtype") == "/Image":
                    n += 1
            except Exception:
                pass
    except Exception:
        pass
    return n


def count_vector_ops(page):
    try:
        data = page.get_contents()
        if data is None:
            return 0
        return len(VEC_OP.findall(data.get_data()))
    except Exception:
        return 0


def sample_indices(n):
    if n <= SAMPLE_PAGES:
        return list(range(n))
    head = SAMPLE_PAGES // 2
    tail = SAMPLE_PAGES - head
    return sorted(set(list(range(head)) + list(range(n - tail, n))))


def analyze_pdf(path, out):
    if PdfReader is None:
        out.update({"ok": False, "reason": "pypdf 미설치 (pip install pypdf)",
                    "parser": "-", "why": "-"})
        return out
    try:
        r = PdfReader(path)
        # 암호화(비밀번호) PDF는 손상과 구분해 명시 (No.4). 빈 비번이면 열어보고 넘어감.
        if getattr(r, "is_encrypted", False):
            try:
                opened = r.decrypt("")  # pypdf: 0=실패, 1/2=성공
            except Exception:
                opened = 0
            if not opened:
                out.update({"ok": False, "reason": "암호화(비밀번호) — 해제 후 업로드",
                            "parser": "검토", "why": "-"})
                return out
        pages = len(r.pages)
    except Exception as e:
        out.update({"ok": False, "reason": f"PDF 파싱 실패: {e}", "parser": "-", "why": "-"})
        return out
    if pages == 0:
        out.update({"ok": False, "reason": "페이지 0장", "parser": "-", "why": "-"})
        return out

    idx = sample_indices(pages)
    sampled = len(idx)
    txt_dense = vec_dense = total_imgs = total_txt = read_err = 0
    for i in idx:
        try:
            p = r.pages[i]
            total_imgs += count_page_images(p)
            tl = len(p.extract_text() or "")
            total_txt += tl
            if tl > HEAVY:
                txt_dense += 1
            if count_vector_ops(p) > VEC_HEAVY:
                vec_dense += 1
        except Exception:
            read_err += 1

    if read_err == sampled:
        out.update({"ok": False, "reason": "모든 표본 페이지 읽기 실패", "parser": "-", "why": "-"})
        return out

    img_per_page = total_imgs / sampled
    dense_ratio = txt_dense / sampled
    vec_ratio = vec_dense / sampled
    txt_per_page = total_txt / sampled

    if vec_ratio >= 0.30:
        parser, why = "LLM Parser", f"벡터 드로잉 과밀 {vec_ratio:.0%} (차트가 벡터→Docling 파편화 위험)"
    elif vec_ratio >= 0.15 or dense_ratio >= 0.40:
        parser, why = "LLM Parser", f"벡터/텍스트 과밀 (vec {vec_ratio:.0%}, text {dense_ratio:.0%})"
    elif txt_per_page < 50 and img_per_page < 0.3:
        parser, why = "검토", f"텍스트·이미지 거의 없음 (txt/pg {txt_per_page:.0f}) — 스캔본/빈파일 의심, OCR 확인"
    elif img_per_page < PDF_TEXT_ONLY_IMG:
        # 이미지·벡터차트 거의 없는 순수 텍스트 PDF → Default(내장 설정으로 텍스트만 추출)로 충분
        parser, why = "Default", f"텍스트 위주(이미지 거의 없음) — 내장 텍스트 추출로 충분 (img/pg {img_per_page:.2f}, txt/pg {txt_per_page:.0f}, vec {vec_ratio:.0%})"
    else:
        parser, why = "Docling", f"이미지·텍스트 기반, 과밀 낮음 (img/pg {img_per_page:.2f}, vec {vec_ratio:.0%}, txt/pg {txt_per_page:.0f})"

    out.update({"ok": True, "reason": "정상" + ("  ※일부페이지 읽기실패" if read_err else ""),
                "parser": parser, "why": why, "img_signal": img_per_page >= 0.5})
    return out


# ============================== Office (OOXML) ===============================
def analyze_office(path, ext, out):
    """docx/pptx/xlsx = ZIP+XML. 표준 라이브러리만으로 chart/media/text 카운트."""
    try:
        zf = zipfile.ZipFile(path)
        names = zf.namelist()
    except Exception as e:
        # 암호화된 OOXML은 ZIP이 아니라 OLE 컨테이너(D0 CF 11 E0)로 감싸짐 → 손상과 구분 (No.4).
        try:
            with open(path, "rb") as fh:
                magic = fh.read(8)
        except Exception:
            magic = b""
        if magic.startswith(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"):
            out.update({"ok": False, "reason": "암호화 추정 — 해제 후 업로드", "parser": "검토", "why": "-"})
        else:
            out.update({"ok": False, "reason": f"OOXML(zip) 열기 실패: {e}", "parser": "-", "why": "-"})
        return out

    prefix = {"docx": "word", "pptx": "ppt", "xlsx": "xl"}[ext[1:]]

    # 차트 XML: <prefix>/charts/chart*.xml  (네이티브 삽입 차트 수)
    charts = sum(1 for n in names
                 if re.search(rf"{prefix}/charts/chart\d+\.xml$", n))
    # 이미지: <prefix>/media/*  (삽입 이미지·캡처 그림)
    media = sum(1 for n in names if n.startswith(f"{prefix}/media/"))

    # 텍스트 길이 추출 (형식별 본문 XML의 텍스트런만 대략)
    text_len = 0
    try:
        if ext == ".docx":
            targets = ["word/document.xml"]
            tag = re.compile(rb"<w:t[ >].*?</w:t>", re.S)
        elif ext == ".pptx":
            targets = [n for n in names if re.match(r"ppt/slides/slide\d+\.xml$", n)]
            tag = re.compile(rb"<a:t>.*?</a:t>", re.S)
        else:  # .xlsx
            # 공유문자열(<t>) + 시트 셀 값(<v>) 둘 다 집계.
            # 숫자 위주 시트는 sharedStrings가 비어도 <v>에 값이 있음.
            targets = ([n for n in names if n == "xl/sharedStrings.xml"]
                       + [n for n in names if re.match(r"xl/worksheets/sheet\d+\.xml$", n)])
            tag = re.compile(rb"<t[ >].*?</t>|<v>.*?</v>", re.S)
        for t in targets:
            try:
                raw = zf.read(t)
            except KeyError:
                continue
            for m in tag.findall(raw):
                inner = re.sub(rb"<[^>]+>", b"", m)  # 태그 제거
                text_len += len(inner.decode("utf-8", "ignore").strip())
    except Exception:
        pass

    # 표(table) 감지: Docling(text+tables) vs Default(text only) 분기 신호
    #   docx=<w:tbl>, pptx=<a:tbl>. body XML만 다시 훑어 표 개수 카운트.
    tables = 0
    try:
        if ext == ".docx":
            tbl_re, tbl_targets = re.compile(rb"<w:tbl[ >]"), ["word/document.xml"]
        elif ext == ".pptx":
            tbl_re, tbl_targets = re.compile(rb"<a:tbl[ >]"), \
                [n for n in names if re.match(r"ppt/slides/slide\d+\.xml$", n)]
        else:
            tbl_re, tbl_targets = None, []
        for t in tbl_targets:
            try:
                tables += len(tbl_re.findall(zf.read(t)))
            except KeyError:
                continue
    except Exception:
        pass

    # 슬라이드/시트 수 (규모 참고)
    slides = sum(1 for n in names if re.match(r"ppt/slides/slide\d+\.xml$", n))
    sheets = sum(1 for n in names if re.match(r"xl/worksheets/sheet\d+\.xml$", n))

    # ---- 판정 ----
    if ext == ".xlsx":
        if text_len < 10 and media == 0 and charts == 0:
            parser, why, ok = "검토", "빈 통합문서로 보임 (텍스트·차트·이미지 없음)", True
        else:
            parser, why, ok = "Docling", f"스프레드시트(표/수치) — 구조 파싱 적합 (sheets {sheets}, charts {charts}, media {media})", True
    else:
        # docx / pptx
        has_text = text_len >= OFFICE_TEXT_MIN
        extra = f", slides {slides}" if ext == ".pptx" else ""
        if not has_text and (media > 0 or charts > 0):
            # 텍스트가 거의 없고 그림/차트 위주 = 캡처·이미지 슬라이드형 → 이미지 의미 재구성 유리
            parser, why, ok = "LLM Parser", f"텍스트 희박+이미지/차트 위주 (text {text_len}, media {media}, charts {charts}) — 이미지 처리 유리", True
        elif not has_text and media == 0 and charts == 0:
            parser, why, ok = "검토", f"내용 거의 없음 (text {text_len}, media 0, charts 0) — 빈 파일 의심", True
        elif tables > 0 or charts > 0 or media > 0:
            # 텍스트 있음 + 표/차트/이미지 존재 → 구조(표) 파싱이 필요 → Docling
            parser, why, ok = "Docling", f"텍스트+표/그래픽 — 구조 파싱 적합 (text {text_len}, tables {tables}, charts {charts}, media {media}{extra})", True
        else:
            # 텍스트만 있고 표·차트·이미지 없음 → 내장 텍스트 추출로 충분 → Default
            parser, why, ok = "Default", f"텍스트 위주(표·이미지 없음) — 내장 텍스트 추출로 충분 (text {text_len}, tables 0, charts 0, media 0{extra})", True

    out.update({"ok": ok, "reason": "정상", "parser": parser, "why": why,
                "img_signal": media > 0})
    return out


# ============================== 공통 진입 =====================================
def analyze(path):
    """확장자로 분기. 해시는 항상 계산(중복판정용)."""
    ext = os.path.splitext(path)[1].lower()
    name = os.path.basename(path)
    # 파일명 유효성은 형식과 무관하게 항상 검사 (No.5). convert는 기본 "" (No.6).
    out = {"file": path, "name": name, "ext": ext.lstrip("."),
           "name_issue": check_filename(name, path), "convert": ""}
    try:
        out["size"] = os.path.getsize(path)
    except OSError:
        out["size"] = 0
    try:
        out["sha256"] = file_sha256(path)
    except Exception as e:
        out.update({"sha256": None, "ok": False, "reason": f"읽기 실패: {e}",
                    "parser": "-", "why": "-"})
        return out

    if ext in EXT_PDF:
        return analyze_pdf(path, out)
    if ext in EXT_OOXML:
        return analyze_office(path, ext, out)
    if ext in EXT_LEGACY:
        # 구형 바이너리는 직접 텍스트 추출이 까다로워 PDF/최신 포맷 변환 대상 (No.6).
        out.update({"ok": True, "reason": "정상(구형 바이너리)",
                    "parser": "검토",
                    "convert": "변환 권장(구형 포맷 → PDF 또는 .docx/.pptx/.xlsx)",
                    "why": "구형 포맷(.doc/.ppt/.xls) — 구조 분석 불가. .docx/.pptx/.xlsx로 변환 후 재판정 권장"})
        return out
    out.update({"ok": False, "reason": "미지원 형식", "parser": "-", "why": "-"})
    return out


def _is_target(f, exclude_root=None):
    """지원 확장자이고, Office 임시 잠금파일(~$...)이 아니며,
    이동 목적지(_triage) 안에 있지 않은 실제 파일."""
    if not os.path.isfile(f):
        return False
    if os.path.basename(f).startswith("~$"):   # Word/PPT/Excel이 여는 순간 만드는 lock 파일
        return False
    if exclude_root:
        # 이동 목적지 폴더 안의 파일은 재분류 대상에서 제외 (무한 재이동 방지)
        try:
            if os.path.commonpath([os.path.abspath(f), exclude_root]) == exclude_root:
                return False
        except ValueError:
            pass
    return os.path.splitext(f)[1].lower() in ALL_EXTS


def collect_docs(args, recursive=False, exclude_root=None):
    docs = []
    for a in args:
        a = os.path.expanduser(a)
        if os.path.isdir(a):
            pat = os.path.join(a, "**", "*") if recursive else os.path.join(a, "*")
            for f in glob.glob(pat, recursive=recursive):
                if _is_target(f, exclude_root):
                    docs.append(f)
        elif _is_target(a, exclude_root):
            docs.append(a)
    return sorted(set(docs))


# 권장 파싱방식 → 폴더 이름
DEST_FOLDER = {"LLM Parser": "LLM_Parser", "Docling": "Docling", "Default": "Default", "검토": "검토"}


def size_bucket(r):
    """용량 상태 문자열: 2GB 초과='업로드불가', 경고선 초과='대용량', 그 외=''."""
    sz = r.get("size", 0)
    if sz > OVERSIZE_BYTES:
        return "업로드불가"
    if sz > LARGE_BYTES:
        return "대용량"
    return ""


# ============================== 파일 이동 ====================================
def _unique_dest(dst_dir, name):
    """목적지에 같은 이름이 있으면 ' (2)', ' (3)' … 붙여 충돌 방지."""
    base, ext = os.path.splitext(name)
    cand = os.path.join(dst_dir, name)
    i = 2
    while os.path.exists(cand):
        cand = os.path.join(dst_dir, f"{base} ({i}){ext}")
        i += 1
    return cand


def move_files(rows, out_root):
    """판정 결과대로 파일을 하위 폴더로 이동.
      - 손상/암호화/미지원 → _errors
      - 중복              → _duplicates
      - 업로드불가(2GB↑)   → _oversize (분할/변환 필수; 최우선 분리)
      - 대용량(>경고선)    → _large    (파싱 비용 절감용 별도 분리; 파서 폴더보다 우선)
      - 그 외 원본/고유    → 권장_파싱방식 폴더 (LLM_Parser/Docling/Default/검토)"""
    moved, errs = 0, 0
    summary = {}
    for r in rows:
        src = r["file"]
        if not os.path.isfile(src):
            continue
        if not r.get("ok"):
            sub = "_errors"
        elif r.get("dup") == "중복":
            sub = "_duplicates"
        elif r.get("size", 0) > OVERSIZE_BYTES:
            sub = "_oversize"
        elif r.get("size", 0) > LARGE_BYTES:
            sub = "_large"
        else:  # 원본 또는 고유
            sub = DEST_FOLDER.get(r.get("parser"), "_errors")
        dst_dir = os.path.join(out_root, sub)
        os.makedirs(dst_dir, exist_ok=True)
        dst = _unique_dest(dst_dir, r["name"])
        try:
            shutil.move(src, dst)
            r["file"] = dst  # CSV에 이동 후 경로 반영
            moved += 1
            summary[sub] = summary.get(sub, 0) + 1
        except Exception as e:
            print(f"  이동 실패: {r['name']} → {sub}  ({e})")
            errs += 1
    parts = "  ".join(f"{k}={v}" for k, v in sorted(summary.items()))
    print(f"\n이동 완료: {moved}개  ({parts})" + (f"  실패={errs}" if errs else ""))
    print(f"목적지: {out_root}")


def human_size(n):
    """바이트 → 사람이 읽는 용량 문자열 (1GB 이상은 GB, 그 외 MB)."""
    b = float(n or 0)
    if b >= 1024 ** 3:
        return f"{b / 1024 ** 3:.2f}GB"
    return f"{b / 1024 ** 2:.1f}MB"


def size_recommendation(r):
    """용량 상태별 안내 문구. 2GB 초과=업로드 불가, 경고선 초과=축소 권장, 그 외 빈 문자열."""
    sz = r.get("size", 0)
    if sz > OVERSIZE_BYTES:
        return "2GB 초과 — Market Insight Uploader 업로드 불가. 분할/변환 필수"
    if sz <= LARGE_BYTES:
        return ""
    if r.get("img_signal"):
        return "대용량(이미지 다수 추정) — 이미지 해상도↓/불필요 이미지 제거로 용량 축소 후 파싱 권장"
    return "대용량 — 불필요 요소 제거로 용량↓ 후 파싱 권장"


CSV_NAME = "Triage 결과.csv"
XLSX_NAME = "Triage 결과.xlsx"


def resolve_out_path(csv_arg, paths, out_root=None, want_xlsx=False):
    """결과 파일 저장 경로 결정. (기본 파일명은 CSV_NAME / XLSX_NAME)
      - csv_arg가 .csv/.xlsx 파일명 → 그 경로 그대로 사용 (최우선)
      - csv_arg가 폴더             → 그 폴더에 생성
      - out_root(--move 목적지)가 있으면 → 그 폴더 바로 밑에 생성
      - 그 외                       → 첫 스캔 폴더(또는 첫 파일의 폴더)에 생성
    """
    default_name = XLSX_NAME if want_xlsx else CSV_NAME
    if csv_arg:
        p = os.path.abspath(os.path.expanduser(csv_arg))
        if os.path.isdir(p) or csv_arg.endswith(os.sep):
            return os.path.join(p, default_name)
        return p
    if out_root:
        return os.path.join(out_root, default_name)
    # 미지정: 첫 스캔 대상의 폴더
    first = os.path.abspath(os.path.expanduser(paths[0]))
    folder = first if os.path.isdir(first) else os.path.dirname(first)
    return os.path.join(folder, default_name)


# ============================== XLSX (stdlib) ================================
# openpyxl 등 외부 의존성 없이 zipfile+XML로 최소 스펙 .xlsx를 직접 생성.
# 이 도구는 이미 Office=ZIP+XML을 읽고 있으므로, 쓰는 쪽도 표준 라이브러리로 해결한다.
# 지원: 여러 시트, 헤더 굵게(1행), 문자열/숫자 자동 구분. (서식은 최소한만)
def _xl_col(idx):
    """0-based 열 인덱스 → 엑셀 열 문자(A, B, ..., Z, AA...)."""
    s = ""
    idx += 1
    while idx:
        idx, rem = divmod(idx - 1, 26)
        s = chr(65 + rem) + s
    return s


def _xl_esc(s):
    return (str(s).replace("&", "&amp;").replace("<", "&lt;")
            .replace(">", "&gt;").replace('"', "&quot;"))


def _sheet_xml(rows):
    """2차원 리스트(rows) → worksheet XML. 1행은 헤더(style s=1: 굵게).
    각 셀: 숫자(int/float)면 <c t=n>, 그 외 문자열은 inlineStr로 기록."""
    out = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
           '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>']
    for r_i, row in enumerate(rows, 1):
        out.append(f'<row r="{r_i}">')
        for c_i, val in enumerate(row):
            ref = f"{_xl_col(c_i)}{r_i}"
            style = ' s="1"' if r_i == 1 else ""
            if isinstance(val, bool):
                val = str(val)
            if isinstance(val, (int, float)):
                out.append(f'<c r="{ref}"{style}><v>{val}</v></c>')
            else:
                txt = _xl_esc("" if val is None else val)
                out.append(f'<c r="{ref}"{style} t="inlineStr"><is><t xml:space="preserve">{txt}</t></is></c>')
        out.append('</row>')
    out.append('</sheetData></worksheet>')
    return "".join(out)


def write_xlsx(path, sheets):
    """sheets = [(시트이름, [[행]...]), ...] 를 하나의 .xlsx로 저장.
    NFC 정규화는 호출측에서 끝내고 넘긴다고 가정."""
    n = len(sheets)
    content_types = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
        + "".join(f'<Override PartName="/xl/worksheets/sheet{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
                  for i in range(1, n + 1))
        + '</Types>')
    root_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
        '</Relationships>')
    wb_sheets = "".join(
        f'<sheet name="{_xl_esc(name)[:31]}" sheetId="{i}" r:id="rId{i}"/>'
        for i, (name, _) in enumerate(sheets, 1))
    workbook = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        f'<sheets>{wb_sheets}</sheets></workbook>')
    wb_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        + "".join(f'<Relationship Id="rId{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{i}.xml"/>'
                  for i in range(1, n + 1))
        + f'<Relationship Id="rId{n + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
        + '</Relationships>')
    # styles.xml: 폰트 2개(일반 / 굵게), cellXfs 2개(s=0 일반, s=1 굵게)
    styles = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font>'
        '<font><b/><sz val="11"/><name val="Calibri"/></font></fonts>'
        '<fills count="1"><fill><patternFill patternType="none"/></fill></fills>'
        '<borders count="1"><border/></borders>'
        '<cellStyleXfs count="1"><xf/></cellStyleXfs>'
        '<cellXfs count="2"><xf fontId="0"/><xf fontId="1" applyFont="1"/></cellXfs>'
        '</styleSheet>')
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types)
        z.writestr("_rels/.rels", root_rels)
        z.writestr("xl/workbook.xml", workbook)
        z.writestr("xl/_rels/workbook.xml.rels", wb_rels)
        z.writestr("xl/styles.xml", styles)
        for i, (_, rows) in enumerate(sheets, 1):
            z.writestr(f"xl/worksheets/sheet{i}.xml", _sheet_xml(rows))


def main():
    global LARGE_BYTES  # --large-mb로 경고선 조절 (2GB 업로드불가 상한은 고정)
    ap = argparse.ArgumentParser(description="로컬 문서(PDF/Word/PPT/Excel) 대량 트리아지 → 단일 CSV")
    ap.add_argument("paths", nargs="+", help="파일 또는 폴더")
    ap.add_argument("--csv", help="결과 CSV 경로/폴더. 폴더면 그 안에 'Triage 결과.csv' 생성. "
                                   "생략 시 --move 목적지(_triage) 바로 밑에, 이동이 없으면 스캔 폴더에 생성 (엑셀용 UTF-8 BOM)")
    ap.add_argument("--jobs", type=int, default=max(1, (os.cpu_count() or 2) - 1),
                    help="병렬 프로세스 수 (기본: 코어-1). 제한 환경이면 자동 순차 폴백")
    ap.add_argument("--recursive", action="store_true", help="하위 폴더까지 재귀 탐색")
    ap.add_argument("--move", action="store_true",
                    help="판정 결과대로 파일을 폴더로 이동 (원본/고유→파서폴더, 중복→_duplicates, 손상→_errors)")
    ap.add_argument("--dest", default="~/Downloads/_triage",
                    help="--move 시 이동 목적지 루트 (기본: ~/Downloads/_triage)")
    ap.add_argument("--large-mb", type=float, default=LARGE_BYTES / (1024 * 1024),
                    help=f"대용량 경고 임계값(MB). 기본 {LARGE_BYTES // (1024*1024)}MB 초과 시 축소 권장/_large 분리. "
                         "2GB 초과 업로드불가 기준은 고정")
    ap.add_argument("--xlsx", action="store_true",
                    help="결과를 CSV 대신 엑셀(.xlsx)로 저장 — '요약' 시트(유형별 분포·집계·대용량/손상 목록) + "
                         "'상세' 시트 2장. --csv 값이 .xlsx로 끝나면 자동으로 켜짐 (외부 라이브러리 불필요)")
    a = ap.parse_args()
    LARGE_BYTES = int(a.large_mb * 1024 * 1024)

    out_root = os.path.abspath(os.path.expanduser(a.dest)) if a.move else None
    docs = collect_docs(a.paths, recursive=a.recursive, exclude_root=out_root)
    if not docs:
        sys.exit("지원 문서(.pdf/.docx/.pptx/.xlsx/.doc/.ppt/.xls)를 찾지 못했습니다.")
    print(f"총 {len(docs)}개 스캔 (jobs={a.jobs})...")

    rows = []
    use_pool = a.jobs > 1 and len(docs) > 1
    if use_pool:
        try:
            with ProcessPoolExecutor(max_workers=a.jobs) as ex:
                futs = {ex.submit(analyze, p): p for p in docs}
                for i, fut in enumerate(as_completed(futs), 1):
                    rows.append(fut.result())
                    if i % 200 == 0:
                        print(f"  ...{i}/{len(docs)}")
        except (PermissionError, OSError, NotImplementedError) as e:
            print(f"  (병렬 불가 → 순차 실행으로 전환: {e})")
            rows, use_pool = [], False
    if not use_pool:
        for i, p in enumerate(docs, 1):
            rows.append(analyze(p))
            if i % 200 == 0:
                print(f"  ...{i}/{len(docs)}")

    # 중복 판정 (SHA-256)
    by_hash = {}
    for r in rows:
        if r.get("sha256"):
            by_hash.setdefault(r["sha256"], []).append(r)
    group_no, gid = {}, 0
    for h, items in by_hash.items():
        if len(items) > 1:
            gid += 1
            group_no[h] = gid
    seen_first = {}
    for r in sorted(rows, key=lambda x: x["name"].lower()):
        h = r.get("sha256")
        if h in group_no and group_no[h] not in seen_first:
            seen_first[group_no[h]] = r["file"]
    for r in rows:
        h = r.get("sha256")
        if h in group_no:
            g = group_no[h]
            if seen_first[g] == r["file"]:
                r["dup"], r["origin"] = "원본", ""
            else:
                r["dup"], r["origin"] = "중복", os.path.basename(seen_first[g])
        else:
            r["dup"], r["origin"] = "고유", ""

    # 정렬: 손상 먼저 → LLM Parser > 검토 > Docling > Default
    porder = {"LLM Parser": 0, "검토": 1, "Docling": 2, "Default": 3, "-": 4}
    rows.sort(key=lambda r: (r.get("ok", False), porder.get(r.get("parser", "-"), 9), r["name"].lower()))

    nLLM = sum(1 for r in rows if r.get("parser") == "LLM Parser")
    nDoc = sum(1 for r in rows if r.get("parser") == "Docling")
    nDef = sum(1 for r in rows if r.get("parser") == "Default")
    nRev = sum(1 for r in rows if r.get("parser") == "검토")
    nEnc = sum(1 for r in rows if "암호화" in (r.get("reason") or ""))
    nBad = sum(1 for r in rows if not r.get("ok"))
    nDup = sum(1 for r in rows if r.get("dup") == "중복")
    print(f"\n결과: LLM Parser={nLLM}  Docling={nDoc}  Default={nDef}  검토={nRev}  손상={nBad}(암호화 {nEnc})  |  중복={nDup}")

    # 유형별 개수·용량 분포 (No.1) — 확장자별 누적
    dist = {}
    for r in rows:
        e = r.get("ext") or "(없음)"
        c, s = dist.get(e, (0, 0))
        dist[e] = (c + 1, s + r.get("size", 0))
    print("\n유형별 분포:")
    for e in sorted(dist, key=lambda k: dist[k][1], reverse=True):
        c, s = dist[e]
        print(f"  {e:<6}: {c:>5}개  {human_size(s):>9}")

    # 파일명 유효성 위반 (No.5) / PDF·포맷 변환 대상 (No.6)
    nName = sum(1 for r in rows if r.get("name_issue"))
    nConv = sum(1 for r in rows if r.get("convert"))
    if nName or nConv:
        print(f"\n파일명 검사 위반={nName}  변환 대상={nConv}")

    # 대용량 2단계: 업로드불가(2GB↑) / 대용량(경고선↑)
    over = sorted((r for r in rows if r.get("size", 0) > OVERSIZE_BYTES
                   and r.get("dup") != "중복"),
                  key=lambda r: r.get("size", 0), reverse=True)
    if over:
        tail = "  → --move 시 _oversize 폴더로 분리" if a.move else ""
        print(f"\n⛔ 업로드 불가 {len(over)}개 (>2GB, Uploader 제한) — 분할/변환 필수{tail}:")
        for r in over:
            print(f"  {human_size(r['size']):>9}  {r['name']}")

    large = sorted((r for r in rows if LARGE_BYTES < r.get("size", 0) <= OVERSIZE_BYTES
                    and r.get("ok") and r.get("dup") != "중복"),
                   key=lambda r: r.get("size", 0), reverse=True)
    if large:
        tail = "  → --move 시 _large 폴더로 분리" if a.move else ""
        print(f"\n⚠ 대용량 {len(large)}개 (>{human_size(LARGE_BYTES)}) — 파싱 비용↓ 위해 축소 권장{tail}:")
        for r in large:
            hint = " (이미지 다수 추정 → 이미지 축소)" if r.get("img_signal") else ""
            print(f"  {human_size(r['size']):>9}  {r['name']}{hint}")

    # 이동 (파일 저장보다 먼저 → 결과에 이동 후 경로가 기록되도록)
    if a.move:
        move_files(rows, out_root)

    # macOS 파일 시스템은 파일명을 NFD(분해형)로 저장 → Excel/Windows에서 한글이 자모 분리로 깨져 보임.
    # 결과에 쓰는 텍스트 필드만 NFC(조합형)로 정규화 (실제 파일은 건드리지 않음).
    def nfc(s):
        return unicodedata.normalize("NFC", s) if isinstance(s, str) else s

    header = ["번호", "파일명", "경로", "형식", "권장_파싱방식", "이유",
              "용량", "용량_상태", "용량_추천", "변환대상", "이름검사",
              "중복여부", "원본파일", "정상여부"]
    detail_rows = [[n, nfc(r["name"]), nfc(r["file"]), r.get("ext", ""), r.get("parser", "-"),
                    r.get("why", "-"), human_size(r.get("size", 0)), size_bucket(r),
                    size_recommendation(r), r.get("convert", ""), nfc(r.get("name_issue", "")),
                    r.get("dup", "-"), nfc(r.get("origin", "")), r.get("reason", "-")]
                   for n, r in enumerate(rows, 1)]

    want_xlsx = a.xlsx or (bool(a.csv) and a.csv.lower().endswith(".xlsx"))
    out_path = resolve_out_path(a.csv, a.paths, out_root if a.move else None, want_xlsx)
    os.makedirs(os.path.dirname(out_path) or ".", exist_ok=True)

    if want_xlsx:
        # ---- '요약' 시트: 콘솔에 찍히던 요약을 그대로 표로 (섹션별 소제목 + 데이터) ----
        s = [["Data 360 업로드 사전점검 요약"], [],
             ["항목", "값"],
             ["총 파일 수", len(rows)],
             ["LLM Parser", nLLM], ["Docling", nDoc], ["Default", nDef], ["검토", nRev],
             ["손상", nBad], ["  └ 암호화", nEnc],
             ["중복", nDup],
             ["파일명 검사 위반", nName], ["변환 대상", nConv],
             ["업로드 불가(>2GB)", len(over)], [f"대용량(>{human_size(LARGE_BYTES)})", len(large)],
             [],
             ["유형별 분포"], ["형식", "개수", "총용량"]]
        for e in sorted(dist, key=lambda k: dist[k][1], reverse=True):
            c, sz = dist[e]
            s.append([e, c, human_size(sz)])
        if over:
            s += [[], ["⛔ 업로드 불가 (>2GB, 분할/변환 필수)"], ["용량", "파일명"]]
            s += [[human_size(r["size"]), nfc(r["name"])] for r in over]
        if large:
            s += [[], [f"⚠ 대용량 (>{human_size(LARGE_BYTES)}, 축소 권장)"], ["용량", "파일명"]]
            s += [[human_size(r["size"]), nfc(r["name"])
                   + (" (이미지 다수 추정)" if r.get("img_signal") else "")] for r in large]
        bad_rows = [r for r in rows if not r.get("ok")]
        if bad_rows:
            s += [[], ["손상 · 암호화 · 미지원"], ["파일명", "사유"]]
            s += [[nfc(r["name"]), r.get("reason", "-")] for r in bad_rows]
        name_bad = [r for r in rows if r.get("name_issue")]
        if name_bad:
            s += [[], ["파일명 검사 위반"], ["파일명", "사유"]]
            s += [[nfc(r["name"]), nfc(r["name_issue"])] for r in name_bad]

        write_xlsx(out_path, [("요약", s), ("상세", [header] + detail_rows)])
        print(f"XLSX 저장: {out_path}  (시트: 요약 / 상세)")
    else:
        with open(out_path, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f)
            w.writerow(header)
            for row in detail_rows:
                w.writerow(row)
        print(f"CSV 저장: {out_path}")


if __name__ == "__main__":
    main()
