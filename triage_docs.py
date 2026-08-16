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
     이미지·텍스트 기반+과밀 낮음 → "Docling". 텍스트·이미지 둘다 희박 → "검토"(스캔본/빈).

Office (구조가 명확 → 사실 기반):
  OOXML은 차트가 charts/chart*.xml 로 "구조적으로" 박혀 있어 PDF식 파편화가 없다.
  - charts : ppt|word|xl/charts/chart*.xml 개수 = 네이티브 차트 수
  - media  : ppt|word|xl/media/* 개수 = 삽입 이미지 수
  - text   : 슬라이드/문서 XML의 텍스트런 길이
  → 차트가 많다: 차트는 XML 데이터라 기본/Docling 파서가 구조적으로 잘 읽음.
     단, 차트를 "그림처럼" 캡처해 넣었거나(media 다수) 텍스트가 거의 없으면 이미지 처리가 유리.
  판정:
    - 텍스트가 실질적으로 있음(+차트/이미지 보통)  → "Docling" (구조 파싱으로 충분)
    - 텍스트 거의 없고 이미지 위주(슬라이드형/캡처)  → "LLM Parser" (이미지→의미 재구성)
    - 텍스트·차트·이미지 모두 희박(빈 파일)          → "검토"
  XLSX는 본질이 표/수치 → 항상 "Docling"(표는 구조 파싱이 정답), 단 빈 시트는 "검토".

중복여부: 파일 바이트 SHA-256이 같은 파일끼리 그룹. 파일명순 첫 파일 "원본", 나머지 "중복",
          유일 "고유". 중복 행에는 "원본파일" 컬럼에 동일 내용 원본 파일명 표시.

정상여부: 열리고 내용이 있으면 "정상", 아니면 "손상"(+원인).

사용법:
  python3 triage_docs.py <파일_또는_폴더> [...] [--csv 결과.csv] [--jobs 8] [--recursive]
  python3 triage_docs.py ~/docs --csv ~/Downloads/triage.csv
  (--csv 생략 시 콘솔 요약만)

CSV 필드: 번호, 파일명, 경로, 형식, 권장_파싱방식, 이유, 중복여부, 원본파일, 정상여부
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
LARGE_BYTES = 20 * 1024 * 1024  # 이 크기 초과면 "대용량"으로 표시 (Data 360 업로드/파싱 부담)
HASH_CHUNK = 1 << 20  # 해시 읽기 블록 (1MB)
VEC_OP = re.compile(rb'(?:^|\s)(?:m|l|c|v|y|re|f|F|f\*|B|B\*|b|b\*|S|s|W|W\*)(?=\s)')

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
        if not has_text and (media > 0 or charts > 0):
            # 텍스트가 거의 없고 그림/차트 위주 = 캡처·이미지 슬라이드형 → 이미지 의미 재구성 유리
            parser, why, ok = "LLM Parser", f"텍스트 희박+이미지/차트 위주 (text {text_len}, media {media}, charts {charts}) — 이미지 처리 유리", True
        elif not has_text and media == 0 and charts == 0:
            parser, why, ok = "검토", f"내용 거의 없음 (text {text_len}, media 0, charts 0) — 빈 파일 의심", True
        else:
            extra = f", slides {slides}" if ext == ".pptx" else ""
            parser, why, ok = "Docling", f"텍스트 기반 문서 — 구조 파싱 적합 (text {text_len}, charts {charts}, media {media}{extra})", True

    out.update({"ok": ok, "reason": "정상", "parser": parser, "why": why,
                "img_signal": media > 0})
    return out


# ============================== 공통 진입 =====================================
def analyze(path):
    """확장자로 분기. 해시는 항상 계산(중복판정용)."""
    ext = os.path.splitext(path)[1].lower()
    out = {"file": path, "name": os.path.basename(path), "ext": ext.lstrip(".")}
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
        out.update({"ok": True, "reason": "정상(구형 바이너리)",
                    "parser": "검토",
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
DEST_FOLDER = {"LLM Parser": "LLM_Parser", "Docling": "Docling", "검토": "검토"}


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
      - 손상/미지원 → _errors
      - 중복         → _duplicates
      - 대용량(>임계) → _large   (파싱 비용 절감용 별도 분리; 파서 폴더보다 우선)
      - 그 외 원본/고유 → 권장_파싱방식 폴더 (LLM_Parser/Docling/검토)"""
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
    """바이트 → MB 단위 문자열 (예: 104.8MB). 용량 표시는 MB로 통일."""
    mb = float(n or 0) / (1024 * 1024)
    return f"{mb:.1f}MB"


def size_recommendation(r):
    """대용량이면 용량 축소 안내 문구, 아니면 빈 문자열."""
    if r.get("size", 0) <= LARGE_BYTES:
        return ""
    if r.get("img_signal"):
        return "대용량(이미지 다수 추정) — 이미지 해상도↓/불필요 이미지 제거로 용량 축소 후 파싱 권장"
    return "대용량 — 불필요 요소 제거로 용량↓ 후 파싱 권장"


CSV_NAME = "Triage 결과.csv"


def resolve_csv_path(csv_arg, paths, out_root=None):
    """CSV 저장 경로 결정. (파일명은 CSV_NAME)
      - csv_arg가 .csv 파일명 → 그 경로 그대로 사용 (최우선)
      - csv_arg가 폴더        → 그 폴더에 생성
      - out_root(--move 목적지)가 있으면 → 그 폴더 바로 밑에 생성
      - 그 외                  → 첫 스캔 폴더(또는 첫 파일의 폴더)에 생성
    """
    if csv_arg:
        p = os.path.abspath(os.path.expanduser(csv_arg))
        if os.path.isdir(p) or csv_arg.endswith(os.sep):
            return os.path.join(p, CSV_NAME)
        return p
    if out_root:
        return os.path.join(out_root, CSV_NAME)
    # 미지정: 첫 스캔 대상의 폴더
    first = os.path.abspath(os.path.expanduser(paths[0]))
    folder = first if os.path.isdir(first) else os.path.dirname(first)
    return os.path.join(folder, CSV_NAME)


def main():
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
    a = ap.parse_args()

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

    # 정렬: 손상 먼저 → LLM Parser > 검토 > Docling
    porder = {"LLM Parser": 0, "검토": 1, "Docling": 2, "-": 3}
    rows.sort(key=lambda r: (r.get("ok", False), porder.get(r.get("parser", "-"), 9), r["name"].lower()))

    nLLM = sum(1 for r in rows if r.get("parser") == "LLM Parser")
    nDoc = sum(1 for r in rows if r.get("parser") == "Docling")
    nRev = sum(1 for r in rows if r.get("parser") == "검토")
    nBad = sum(1 for r in rows if not r.get("ok"))
    nDup = sum(1 for r in rows if r.get("dup") == "중복")
    print(f"\n결과: LLM Parser={nLLM}  Docling={nDoc}  검토={nRev}  손상={nBad}  |  중복={nDup}")

    large = sorted((r for r in rows if r.get("size", 0) > LARGE_BYTES
                    and r.get("ok") and r.get("dup") != "중복"),
                   key=lambda r: r.get("size", 0), reverse=True)
    if large:
        tail = "  → --move 시 _large 폴더로 분리" if a.move else ""
        print(f"\n⚠ 대용량 {len(large)}개 (>{human_size(LARGE_BYTES)}) — 파싱 비용↓ 위해 축소 권장{tail}:")
        for r in large:
            hint = " (이미지 다수 추정 → 이미지 압축)" if r.get("img_signal") else ""
            print(f"  {human_size(r['size']):>9}  {r['name']}{hint}")

    # 이동 (CSV보다 먼저 → CSV에 이동 후 경로가 기록되도록)
    if a.move:
        move_files(rows, out_root)

    csv_path = resolve_csv_path(a.csv, a.paths, out_root if a.move else None)
    os.makedirs(os.path.dirname(csv_path) or ".", exist_ok=True)
    # macOS 파일 시스템은 파일명을 NFD(분해형)로 저장 → Excel/Windows에서 한글이 자모 분리로 깨져 보임.
    # CSV에 쓰는 텍스트 필드만 NFC(조합형)로 정규화 (실제 파일은 건드리지 않음).
    def nfc(s):
        return unicodedata.normalize("NFC", s) if isinstance(s, str) else s
    with open(csv_path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["번호", "파일명", "경로", "형식", "권장_파싱방식", "이유",
                    "용량", "용량_추천", "중복여부", "원본파일", "정상여부"])
        for n, r in enumerate(rows, 1):
            w.writerow([n, nfc(r["name"]), nfc(r["file"]), r.get("ext", ""), r.get("parser", "-"),
                        r.get("why", "-"), human_size(r.get("size", 0)), size_recommendation(r),
                        r.get("dup", "-"), nfc(r.get("origin", "")), r.get("reason", "-")])
    print(f"CSV 저장: {csv_path}")


if __name__ == "__main__":
    main()
