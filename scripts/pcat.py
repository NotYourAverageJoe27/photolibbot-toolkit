#!/usr/bin/env python3
"""pcat - Picture Catalog toolkit (one file, stdlib + Pillow/numpy/openpyxl, ffmpeg).

Subcommands (run `pcat.py <cmd> -h` for options):
  doctor     check dependencies and print install commands
  scan       inventory a folder tree -> catalog JSONL (SHA-1, EXIF, GPS, video meta, best date)
             re-running refreshes it: unchanged and moved files are reused, new files read
  catalog    catalog JSONL -> Photo_Catalog.xlsx (Catalog + summary tabs)
  reconcile  check the catalog matches the disk exactly (0 missing, 0 uncataloged)
  dupes      tiered duplicate finder (A exact, B same video stream, C re-saved photo, D similar)
  compare    classify a NEW source against the library (identical/duplicate/upgrade/new/...)
  plan       dry-run date-based reorganization plan (events, topics, trips, months, labels)
  apply      execute a plan CSV (move/copy/quarantine/delete) with verify, log, undo, resume
  verify     compare two trees (backup check) by path+size, optional SHA-1
  sheet      numbered contact sheets for human review
  date-seq   date undated files from dated neighbours in the shot/clip-number sequence
  set-dates  store dates the user supplied (CSV path,date[,evidence]; year/month/day precision kept)
  geonames / from-inventory   helpers (city names; PowerShell inventory -> catalog)

Videos: `scan` stores a 5-frame visual fingerprint per video; a video is never called a duplicate on
name + length alone - identical streams or a matching fingerprint are required.

Safety: nothing is ever overwritten; deletes need --allow-delete and a verified keeper;
every action is logged; moves write an undo CSV that `apply` can run.
"""
import argparse, csv, datetime as dt, hashlib, json, math, os, re, shutil, statistics, subprocess, sys, time
from collections import Counter, defaultdict
from concurrent.futures import ThreadPoolExecutor

VERSION = "1.1"
PHOTO_EXT = {".jpg", ".jpeg", ".png", ".gif", ".bmp", ".tif", ".tiff", ".heic", ".heif", ".webp", ".psd",
             ".pct", ".pict", ".cr2", ".cr3", ".nef", ".arw", ".dng", ".orf", ".rw2", ".raf"}
VIDEO_EXT = {".mov", ".mp4", ".m4v", ".avi", ".mpg", ".mpeg", ".3gp", ".mts", ".m2ts", ".mkv", ".wmv", ".flv", ".dv"}
AUDIO_EXT = {".mp3", ".wav", ".wma", ".m4a", ".aac", ".flac", ".aif", ".aiff", ".ogg"}
DOC_EXT = {".pdf", ".zip", ".rar", ".7z", ".doc", ".docx", ".xls", ".xlsx", ".ppt", ".pptx", ".txt", ".rtf", ".csv"}
JUNK_NAMES = {"desktop.ini", ".ds_store", "thumbs.db", "ehthumbs.db", "ehthumbs_vista.db", ".picasa.ini",
              "picasa.ini", "feed.rss", "view online.url", "zbthumbnail.info"}
JUNK_EXT = {".ithmb", ".aae", ".thm", ".lrv", ".ctg", ".url", ".rss", ".tmp", ".lnk", ".xmp_tmp"}
SKIP_DIRS = {"$recycle.bin", "system volume information", "_catalog", "__pycache__", ".git"}
TRUE_SOURCES = {"exif", "video"}
DEFAULT_DUMP_RE = (r"(?i)(iphone|ipad|\bphone\b|phone\s|backup|camera ?roll|dcim|\d{3}apple|photo ?stream|"
                   r"\bimport|upload|mobile|android|galaxy|pixel \d)")
US_STATES = {"Alabama": "AL", "Alaska": "AK", "Arizona": "AZ", "Arkansas": "AR", "California": "CA", "Colorado": "CO",
             "Connecticut": "CT", "Delaware": "DE", "Florida": "FL", "Georgia": "GA", "Hawaii": "HI", "Idaho": "ID",
             "Illinois": "IL", "Indiana": "IN", "Iowa": "IA", "Kansas": "KS", "Kentucky": "KY", "Louisiana": "LA",
             "Maine": "ME", "Maryland": "MD", "Massachusetts": "MA", "Michigan": "MI", "Minnesota": "MN",
             "Mississippi": "MS", "Missouri": "MO", "Montana": "MT", "Nebraska": "NE", "Nevada": "NV",
             "New Hampshire": "NH", "New Jersey": "NJ", "New Mexico": "NM", "New York": "NY", "North Carolina": "NC",
             "North Dakota": "ND", "Ohio": "OH", "Oklahoma": "OK", "Oregon": "OR", "Pennsylvania": "PA",
             "Rhode Island": "RI", "South Carolina": "SC", "South Dakota": "SD", "Tennessee": "TN", "Texas": "TX",
             "Utah": "UT", "Vermont": "VT", "Virginia": "VA", "Washington": "WA", "West Virginia": "WV",
             "Wisconsin": "WI", "Wyoming": "WY", "District of Columbia": "DC", "Washington, D.C.": "DC",
             "Puerto Rico": "PR"}
COUNTRIES = {"US": "USA", "CA": "Canada", "MX": "Mexico", "GB": "UK", "IE": "Ireland", "FR": "France", "DE": "Germany",
             "IT": "Italy", "ES": "Spain", "PT": "Portugal", "NL": "Netherlands", "BE": "Belgium", "CH": "Switzerland",
             "AT": "Austria", "HU": "Hungary", "CZ": "Czechia", "PL": "Poland", "GR": "Greece", "HR": "Croatia",
             "MT": "Malta", "MC": "Monaco", "VA": "Vatican City", "SE": "Sweden", "NO": "Norway", "DK": "Denmark",
             "FI": "Finland", "IS": "Iceland", "JP": "Japan", "CN": "China", "KR": "South Korea", "TH": "Thailand",
             "VN": "Vietnam", "IN": "India", "AU": "Australia", "NZ": "New Zealand", "BS": "Bahamas",
             "JM": "Jamaica", "DO": "Dominican Republic", "CR": "Costa Rica", "BR": "Brazil", "AR": "Argentina",
             "CL": "Chile", "PE": "Peru", "CO": "Colombia", "EG": "Egypt", "MA": "Morocco", "ZA": "South Africa",
             "IL": "Israel", "TR": "Turkey", "AE": "UAE", "IR": "Iran", "GU": "Guam", "PH": "Philippines",
             "SG": "Singapore", "HK": "Hong Kong", "TW": "Taiwan", "VI": "US Virgin Islands", "KY": "Cayman Islands",
             "TC": "Turks and Caicos", "BM": "Bermuda", "BB": "Barbados", "LC": "St. Lucia", "AW": "Aruba",
             "CW": "Curacao", "SX": "Sint Maarten", "BZ": "Belize", "PA": "Panama"}


# ----------------------------------------------------------------------------- small helpers
def log(*a):
    print(*a, flush=True)


def P(path):
    """Long-path-safe path for file system calls (Windows \\\\?\\ prefix)."""
    if os.name != "nt":
        return path
    path = os.path.abspath(path)
    if path.startswith("\\\\?\\"):
        return path
    if path.startswith("\\\\"):
        return "\\\\?\\UNC\\" + path[2:]
    return "\\\\?\\" + path


def unP(path):
    if path.startswith("\\\\?\\UNC\\"):
        return "\\\\" + path[8:]
    if path.startswith("\\\\?\\"):
        return path[4:]
    return path


def norm_rel(rel):
    return rel.replace("\\", "/")


def walk_files(root, skip_dirs=SKIP_DIRS, follow_skip=True):
    """Yield (abs_path, rel_path_posix, stat) for every file under root (long-path safe, no symlinks)."""
    root = os.path.abspath(root)
    stack = [P(root)]
    rootP = P(root)
    while stack:
        d = stack.pop()
        try:
            it = list(os.scandir(d))
        except OSError as e:
            log(f"  ! cannot list {unP(d)}: {e}")
            continue
        for e in it:
            try:
                if e.is_symlink():
                    continue
                if e.is_dir(follow_symlinks=False):
                    if follow_skip and e.name.lower() in skip_dirs:
                        continue
                    stack.append(e.path)
                elif e.is_file(follow_symlinks=False):
                    rel = os.path.relpath(e.path, rootP)
                    yield unP(e.path), norm_rel(rel), e.stat(follow_symlinks=False)
            except OSError as ex:
                log(f"  ! {unP(e.path)}: {ex}")


def sha1_file(path, partial=False, chunk=1 << 20):
    h = hashlib.sha1()
    with open(P(path), "rb") as f:
        if partial:
            size = os.fstat(f.fileno()).st_size
            if size <= 3 * chunk:
                h.update(f.read())
            else:
                for off in (0, size // 2, size - chunk):
                    f.seek(off)
                    h.update(f.read(chunk))
                h.update(str(size).encode())
            return h.hexdigest()
        while True:
            b = f.read(chunk)
            if not b:
                break
            h.update(b)
    return h.hexdigest()


def haversine_km(a, b):
    lat1, lon1, lat2, lon2 = map(math.radians, (a[0], a[1], b[0], b[1]))
    x = math.sin((lat2 - lat1) / 2) ** 2 + math.cos(lat1) * math.cos(lat2) * math.sin((lon2 - lon1) / 2) ** 2
    return 6371.0 * 2 * math.asin(min(1.0, math.sqrt(x)))


def load_jsonl(path):
    out = []
    with open(path, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if line:
                out.append(json.loads(line))
    return out


def write_jsonl(path, recs):
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        for r in recs:
            f.write(json.dumps(r, ensure_ascii=False) + "\n")
    os.replace(tmp, path)


def read_csv(path):
    with open(path, newline="", encoding="utf-8-sig") as f:
        return list(csv.DictReader(f))


def write_csv(path, rows, fields=None):
    if fields is None:
        fields = []
        for r in rows:
            for k in r:
                if k not in fields:
                    fields.append(k)
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=fields, extrasaction="ignore")
        w.writeheader()
        for r in rows:
            w.writerow(r)


def abspath_of(rec):
    return os.path.join(rec["root"], *rec["path"].split("/"))


def parse_iso(s):
    if not s:
        return None
    try:
        return dt.datetime.fromisoformat(s)
    except ValueError:
        return None


def hamming(a, b):
    if a is None or b is None:
        return 999
    return bin(int(a, 16) ^ int(b, 16)).count("1")


def res(r):
    return f"{r['width']}x{r['height']}" if r.get("width") and r.get("height") else ""


def pixels(r):
    return (r.get("width") or 0) * (r.get("height") or 0)


# ----------------------------------------------------------------------------- doctor
def cmd_doctor(a):
    ok = True
    for mod, pipname, need in [("PIL", "pillow", True), ("numpy", "numpy", True), ("openpyxl", "openpyxl", True),
                               ("reverse_geocoder", "reverse_geocoder", False), ("pillow_heif", "pillow-heif", False)]:
        try:
            __import__(mod)
            log(f"  ok   {mod}")
        except ImportError:
            log(f"  {'MISSING' if need else 'optional'} {mod}  ->  pip install {pipname}"
                + ("  (add --break-system-packages on managed Linux)" if need else ""))
            ok = ok and not need
    for exe in ("ffprobe", "ffmpeg"):
        p = shutil.which(exe)
        log(f"  {'ok  ' if p else 'MISSING'} {exe} {p or ' -> install ffmpeg (winget install Gyan.FFmpeg / apt install ffmpeg)'}")
        ok = ok and bool(p)
    log("doctor: " + ("all required dependencies present" if ok else "install the MISSING items above"))
    return 0 if ok else 1


# ----------------------------------------------------------------------------- scan
EXIF_DT_RE = re.compile(r"(\d{4})[:\-](\d{2})[:\-](\d{2})[ T](\d{2}):(\d{2}):(\d{2})")
YEAR_RE = re.compile(r"(?<!\d)(19[5-9]\d|20[0-4]\d)(?!\d)")


def parse_exif_dt(v):
    if v is None:
        return None
    if isinstance(v, bytes):
        v = v.decode("ascii", "ignore")
    m = EXIF_DT_RE.search(str(v))
    if not m:
        return None
    try:
        d = dt.datetime(*map(int, m.groups()))
    except ValueError:
        return None
    if d.year < 1980 or d > dt.datetime.now() + dt.timedelta(days=2):
        return None
    return d.isoformat()


def sniff(path):
    """Content type from magic bytes; 'zeros' for all-zero files."""
    try:
        with open(P(path), "rb") as f:
            head = f.read(4096)
            size = os.fstat(f.fileno()).st_size
            tail = b""
            if size > 8192:
                f.seek(size - 4096)
                tail = f.read(4096)
    except OSError:
        return "unreadable"
    if not head:
        return "empty"
    if head.count(0) == len(head) and (not tail or tail.count(0) == len(tail)):
        return "zeros"
    h = head
    if h[:3] == b"\xff\xd8\xff":
        return "jpeg"
    if h[:8] == b"\x89PNG\r\n\x1a\n":
        return "png"
    if h[:6] in (b"GIF87a", b"GIF89a"):
        return "gif"
    if h[:2] == b"BM":
        return "bmp"
    if h[:4] in (b"II*\x00", b"MM\x00*"):
        return "tiff"
    if h[:4] == b"8BPS":
        return "psd"
    if h[4:8] == b"ftyp":
        brand = h[8:12]
        if brand in (b"heic", b"heix", b"mif1", b"msf1", b"heim", b"heis", b"hevc"):
            return "heic"
        if brand == b"crx ":
            return "raw"
        if brand.startswith(b"qt"):
            return "mov"
        if brand.startswith(b"3g"):
            return "3gp"
        return "mp4"
    if h[4:8] in (b"moov", b"mdat", b"wide", b"free", b"skip", b"pnot"):
        return "mov"
    if h[:4] == b"RIFF" and h[8:12] == b"AVI ":
        return "avi"
    if h[:4] == b"RIFF" and h[8:12] == b"WEBP":
        return "webp"
    if h[:4] in (b"\x00\x00\x01\xba", b"\x00\x00\x01\xb3"):
        return "mpeg"
    if h[:4] == b"\x1aE\xdf\xa3":
        return "mkv"
    if h[:4] == b"%PDF":
        return "pdf"
    if h[:2] == b"PK":
        return "zip"
    stripped = h.lstrip()[:15].lower()
    if stripped.startswith((b"<!doctype", b"<html", b"<?xml", b"<head")):
        return "html"
    if len(h) > 512 + 12 and size > 512 and h[522:524] in (b"\x00\x11", b"\x11\x01"):
        return "pict"
    return "unknown"


SNIFF_KIND = {"jpeg": "photo", "png": "photo", "gif": "photo", "bmp": "photo", "tiff": "photo", "psd": "photo",
              "heic": "photo", "webp": "photo", "pict": "photo", "raw": "photo", "mov": "video", "mp4": "video",
              "3gp": "video", "avi": "video", "mpeg": "video", "mkv": "video", "pdf": "doc", "zip": "doc",
              "html": "doc"}


def classify(name, ext, size, snf):
    n = name.lower()
    if n in JUNK_NAMES or ext in JUNK_EXT or n.startswith("._") or size == 0:
        return "junk", None
    if snf == "zeros":
        return "problem", "all-zero file"
    if snf == "unreadable":
        return "problem", "unreadable"
    by_ext = ("photo" if ext in PHOTO_EXT else "video" if ext in VIDEO_EXT else "audio" if ext in AUDIO_EXT
              else "doc" if ext in DOC_EXT else "other")
    by_sniff = SNIFF_KIND.get(snf)
    if by_ext in ("photo", "video") and by_sniff and by_sniff != by_ext:
        if by_sniff in ("photo", "video"):
            return by_sniff, f"wrong extension: content is {snf}"
        return "problem", f"not media: content is {snf}"
    if by_ext == "photo" and snf == "unknown" and ext in {".jpg", ".jpeg", ".png", ".gif", ".bmp"}:
        return "problem", "unrecognized content"
    if by_ext == "other" and by_sniff in ("photo", "video"):
        return by_sniff, f"no media extension: content is {snf}"
    return by_ext, None


_PIL = None


def pil():
    global _PIL
    if _PIL is None:
        from PIL import Image, ImageOps
        Image.MAX_IMAGE_PIXELS = None
        try:
            import pillow_heif
            pillow_heif.register_heif_opener()
        except Exception:
            pass
        _PIL = (Image, ImageOps)
    return _PIL


def _rat(x):
    try:
        return float(x)
    except Exception:
        try:
            return x[0] / x[1]
        except Exception:
            return None


def gps_from_ifd(g):
    try:
        lat, lon = g.get(2), g.get(4)
        if not lat or not lon:
            return None, None
        f = lambda t: _rat(t[0]) + _rat(t[1]) / 60 + _rat(t[2]) / 3600
        la, lo = f(lat), f(lon)
        if str(g.get(1, "N")).upper().startswith("S"):
            la = -la
        if str(g.get(3, "E")).upper().startswith("W"):
            lo = -lo
        if abs(la) < 1e-6 and abs(lo) < 1e-6 or abs(la) > 90 or abs(lo) > 180:
            return None, None
        return round(la, 6), round(lo, 6)
    except Exception:
        return None, None


def dhash_image(im):
    Image, _ = pil()
    g = im.convert("L").resize((9, 8), Image.LANCZOS)
    px = list(g.tobytes())
    bits = 0
    for row in range(8):
        for col in range(8):
            bits = (bits << 1) | (1 if px[row * 9 + col] > px[row * 9 + col + 1] else 0)
    return f"{bits:016x}"


def read_photo(path, r):
    Image, ImageOps = pil()
    try:
        with Image.open(P(path)) as im:
            r["width"], r["height"] = im.size
            r["format"] = im.format
            ex = im.getexif()
            if ex:
                sub = ex.get_ifd(0x8769) or {}
                r["exif_dt"] = parse_exif_dt(sub.get(36867))
                r["exif_dt_digitized"] = parse_exif_dt(sub.get(36868))
                r["exif_dt_modified"] = parse_exif_dt(ex.get(306))
                mk, md = ex.get(271), ex.get(272)
                if mk or md:
                    mk = str(mk or "").strip("\x00 ").strip()
                    md = str(md or "").strip("\x00 ").strip()
                    r["camera"] = (md if md.lower().startswith(mk.lower().split(" ")[0]) else f"{mk} {md}").strip() if mk else md
                r["orientation"] = ex.get(274)
                r["gps_lat"], r["gps_lon"] = gps_from_ifd(ex.get_ifd(0x8825) or {})
                sw = ex.get(305)
                if sw:
                    r["software"] = str(sw).strip("\x00 ")[:60]
            try:
                im.draft("RGB", (256, 256))
                tim = ImageOps.exif_transpose(im)
                r["dhash"] = dhash_image(tim)
            except Exception as e:
                r["note"] = (r.get("note") or "") + f" thumb error: {e}"
    except Exception as e:
        r["error"] = f"image read: {str(e)[:120]}"


ISO6709 = re.compile(r"([+-]\d+(?:\.\d+)?)([+-]\d+(?:\.\d+)?)")


def ffprobe_json(path):
    try:
        out = subprocess.run(["ffprobe", "-v", "error", "-print_format", "json", "-show_format", "-show_streams", path],
                             capture_output=True, timeout=120)
        return json.loads(out.stdout.decode("utf-8", "replace") or "{}")
    except Exception:
        return {}


def read_video(path, r, tz):
    j = ffprobe_json(path)
    if not j:
        r["error"] = "ffprobe failed"
        return
    fmt = j.get("format", {})
    tags = {k.lower(): v for k, v in (fmt.get("tags") or {}).items()}
    try:
        r["duration"] = round(float(fmt.get("duration")), 3)
    except (TypeError, ValueError):
        pass
    for s in j.get("streams", []):
        st = {k.lower(): v for k, v in (s.get("tags") or {}).items()}
        if s.get("codec_type") == "video" and "width" not in r and s.get("disposition", {}).get("attached_pic") != 1:
            r["width"], r["height"] = s.get("width"), s.get("height")
            r["vcodec"] = s.get("codec_name")
            if s.get("bit_rate"):
                r["vbitrate"] = int(s["bit_rate"])
            rot = st.get("rotate")
            for sd in s.get("side_data_list") or []:
                if "rotation" in sd:
                    rot = sd["rotation"]
            if rot and abs(int(float(rot))) in (90, 270):
                r["width"], r["height"] = r["height"], r["width"]
        for k in ("creation_time",):
            if k in st and "stream_creation" not in r:
                r["stream_creation"] = st[k]
    mk, md = tags.get("com.apple.quicktime.make"), tags.get("com.apple.quicktime.model")
    if md:
        r["camera"] = md if mk and md.lower().startswith(mk.lower()) else f"{mk or ''} {md}".strip()
    loc = tags.get("com.apple.quicktime.location.iso6709") or tags.get("location") or tags.get("location-eng")
    if loc:
        m = ISO6709.match(loc)
        if m:
            la, lo = float(m.group(1)), float(m.group(2))
            if not (abs(la) < 1e-6 and abs(lo) < 1e-6):
                r["gps_lat"], r["gps_lon"] = round(la, 6), round(lo, 6)
    d = None
    cd = tags.get("com.apple.quicktime.creationdate")
    if cd:
        try:
            d = dt.datetime.fromisoformat(cd.replace("Z", "+00:00")[:25] if len(cd) > 19 else cd)
            d = d.replace(tzinfo=None)  # local wall-clock time of the phone
        except ValueError:
            d = None
    if d is None:
        ct = tags.get("creation_time") or r.get("stream_creation")
        if ct:
            try:
                u = dt.datetime.fromisoformat(ct.replace("Z", "+00:00"))
                if u.tzinfo is None:
                    u = u.replace(tzinfo=dt.timezone.utc)
                d = u.astimezone(tz).replace(tzinfo=None) if tz else u.replace(tzinfo=None)
            except ValueError:
                d = None
    r.pop("stream_creation", None)
    if d and 1980 <= d.year and d <= dt.datetime.now() + dt.timedelta(days=2):
        r["video_dt"] = d.replace(microsecond=0).isoformat()


def best_date(r):
    if r.get("exif_dt"):
        return r["exif_dt"], "exif"
    if r.get("video_dt"):
        return r["video_dt"], "video"
    if r.get("exif_dt_digitized"):
        return r["exif_dt_digitized"], "exif_digitized"
    if r.get("exif_dt_modified"):
        return r["exif_dt_modified"], "exif_modified"
    years = YEAR_RE.findall(os.path.dirname(r["path"]))
    if years:
        return f"{years[-1]}-01-01T00:00:00", "folder_year"
    if r.get("mtime"):
        return dt.datetime.fromtimestamp(r["mtime"]).replace(microsecond=0).isoformat(), "file_mtime"
    return None, "none"


def scan_one(path, rel, st, root, do_hash, tz):
    name = os.path.basename(path)
    ext = os.path.splitext(name)[1].lower()
    r = {"root": root, "path": rel, "name": name, "ext": ext, "size": st.st_size,
         "mtime": round(st.st_mtime, 3)}
    snf = sniff(path) if st.st_size else "empty"
    r["sniff"] = snf
    kind, issue = classify(name, ext, st.st_size, snf)
    r["kind"] = kind
    if issue:
        r["issue"] = issue
    try:
        if do_hash and st.st_size:
            r["sha1"] = sha1_file(path)
        if kind == "photo":
            read_photo(path, r)
        elif kind == "video":
            read_video(path, r, tz)
    except Exception as e:
        r["error"] = str(e)[:200]
    r["date_best"], r["date_source"] = best_date(r)
    return r


def get_tz(name):
    if not name:
        return None
    from zoneinfo import ZoneInfo
    return ZoneInfo(name)


def cmd_scan(a):
    root = os.path.abspath(a.root)
    out = a.out or os.path.join(root, "_Catalog", "catalog.jsonl")
    os.makedirs(os.path.dirname(out), exist_ok=True)
    tz = get_tz(a.tz)
    old = {}
    if os.path.exists(out) and not a.fresh:
        for r in load_jsonl(out):
            if r.get("root") == root:
                old[r["path"]] = r
    log(f"scan {root}  (existing records: {len(old)})")
    files = list(walk_files(root))
    if a.exclude:
        rx = re.compile(a.exclude)
        files = [f for f in files if not rx.search(f[1])]
    log(f"  {len(files):,} files on disk")
    keep, todo = [], []
    seen = set()
    for path, rel, st in files:
        seen.add(rel)
        o = old.get(rel)
        if o and o["size"] == st.st_size and abs(o.get("mtime", 0) - st.st_mtime) < 2:
            keep.append(o)
        else:
            todo.append((path, rel, st))
    # reuse records of moved/renamed files (same size + mtime, old path gone)
    gone = defaultdict(list)
    for rel, o in old.items():
        if rel not in seen:
            gone[(o["size"], round(o.get("mtime", 0)))].append(o)
    still = []
    reused = 0
    for path, rel, st in todo:
        cands = gone.get((st.st_size, round(st.st_mtime)))
        if cands:
            cand = next((c for c in cands if c["name"].lower() == os.path.basename(rel).lower()), cands[0])
            cands.remove(cand)
            n = dict(cand)
            n["path"], n["name"] = rel, os.path.basename(rel)
            if "orig_path" not in n:
                n["orig_path"] = cand["path"]
            n["date_best"], n["date_source"] = best_date(n) if n["date_source"] in ("folder_year", "file_mtime", "none") else (n["date_best"], n["date_source"])
            keep.append(n)
            reused += 1
        else:
            still.append((path, rel, st))
    dropped = sum(len(v) for v in gone.values())
    log(f"  unchanged {len(keep) - reused:,} | moved/renamed (reused) {reused:,} | to read {len(still):,} | "
        f"records dropped (file gone) {dropped:,}")
    t0 = time.time()
    done = []
    total_bytes = sum(s.st_size for _, _, s in still)
    nb = 0
    last = t0
    with ThreadPoolExecutor(max_workers=a.workers) as ex:
        futs = [ex.submit(scan_one, p, rel, st, root, not a.no_hash, tz) for p, rel, st in still]
        for i, f in enumerate(futs, 1):
            r = f.result()
            done.append(r)
            nb += r["size"]
            if time.time() - last > a.progress or i == len(futs):
                el = time.time() - t0
                rate = nb / el / 1e6 if el else 0
                eta = (total_bytes - nb) / (nb / el) if nb and el else 0
                log(f"  {i:,}/{len(futs):,} files  {nb / 1e9:.1f}/{total_bytes / 1e9:.1f} GB  {rate:.0f} MB/s  "
                    f"ETA {eta / 60:.0f} min")
                last = time.time()
                if a.checkpoint:
                    write_jsonl(out, keep + done)
    recs = sorted(keep + done, key=lambda r: r["path"].lower())
    if not a.no_vfp:  # video fingerprints (also back-fills catalogs made before fingerprints existed)
        need = [r for r in recs if r["kind"] == "video" and "vfp" not in r and r.get("duration")]
        if need:
            log(f"  video fingerprints: {len(need):,} videos ({VFP_FRAMES} frames each)")
            t1 = last = time.time()
            with ThreadPoolExecutor(max_workers=a.workers) as ex:
                for i, (r, fp) in enumerate(zip(need, ex.map(lambda r: video_fp(abspath_of(r), r["duration"]), need)), 1):
                    r["vfp"] = fp
                    if time.time() - last > a.progress:
                        el = time.time() - t1
                        log(f"    {i:,}/{len(need):,}  ETA {(len(need) - i) * el / i / 60:.0f} min")
                        last = time.time()
    write_jsonl(out, recs)
    k = Counter(r["kind"] for r in recs)
    log(f"wrote {out}: {len(recs):,} records  " + ", ".join(f"{n} {c:,}" for n, c in k.most_common()))
    probs = [r for r in recs if r.get("issue") or r.get("error")]
    if probs:
        log(f"  {len(probs)} files with issues (see 'Problems' tab after `catalog`)")
    return 0


# ----------------------------------------------------------------------------- from-inventory
def cmd_from_inventory(a):
    """Turn pcat-win.ps1 inventory.csv into minimal catalog JSONL (path/size/sha1/kind) so `compare` can find
    byte-identical files on drives the shell cannot mount. Stage the rest locally and `scan` them fully."""
    rows = read_csv(a.inventory)
    out = []
    for r in rows:
        root = r["Root"]
        if a.root_map:
            for pair in a.root_map:
                win, loc = pair.split("=", 1)
                if root.lower() == win.lower():
                    root = loc
        rel = norm_rel(r["Rel"])
        name = rel.split("/")[-1]
        ext = os.path.splitext(name)[1].lower()
        size = int(r["Size"])
        kind, issue = classify(name, ext, size, "unknown" if ext not in PHOTO_EXT else "jpeg")
        rec = {"root": root, "path": rel, "name": name, "ext": ext, "size": size, "kind": kind,
               "mtime_text": r.get("MTime"), "date_best": None, "date_source": "none", "from_inventory": True}
        if r.get("Sha1") and not r["Sha1"].startswith("ERROR"):
            rec["sha1"] = r["Sha1"]
        out.append(rec)
    write_jsonl(a.out, out)
    log(f"wrote {a.out}: {len(out):,} records ({sum(1 for x in out if x.get('sha1')):,} with SHA-1). "
        "Files without a SHA-1 had a unique size, so they cannot be byte-identical to anything inventoried with them.")
    return 0


# ----------------------------------------------------------------------------- dating the undated
SEQ_RE = re.compile(r"^(.*?)(\d{3,6})$")


def _backup_catalog(path, tag):
    bk = os.path.splitext(path)[0] + f"_before_{tag}_{dt.datetime.now():%Y%m%d_%H%M%S}.jsonl"
    shutil.copy2(path, bk)
    log(f"  catalog backup: {bk}")


def cmd_date_seq(a):
    """Date undated files from the shot/clip-number sequence of dated neighbours from the same camera series.
    Writes a review CSV; with --write, stores the dates in the catalog (date_source = 'sequence')."""
    recs = load_jsonl(a.catalog)
    trusted = set(a.trusted.split(","))
    groups = defaultdict(lambda: {"dated": [], "undated": []})
    for r in recs:
        if r["kind"] not in ("photo", "video") or r["path"].startswith("_Duplicates_Quarantine/"):
            continue
        stem = os.path.splitext(r["name"])[0]
        m = SEQ_RE.match(stem)
        if not m:
            continue
        prefix = m.group(1).strip(" _-").lower()
        scope = "" if a.any_folder else (r.get("orig_path") or r["path"]).split("/")[0]
        key = (prefix, r["kind"], scope)
        item = (int(m.group(2)), r)
        if r.get("date_source") in trusted:
            groups[key]["dated"].append(item)
        elif r.get("date_source") != "sequence" or a.redo:
            groups[key]["undated"].append(item)
    rows = []
    for key, g in groups.items():
        if not g["dated"] or not g["undated"]:
            continue
        dated = sorted(g["dated"], key=lambda x: x[0])
        nums = [n for n, _ in dated]
        import bisect
        for n, r in g["undated"]:
            k = bisect.bisect_left(nums, n)
            if k == 0 or k >= len(nums) or nums[k] == n:
                continue  # needs a dated neighbour on both sides (and not a duplicate number)
            (ln, lr), (un, ur) = dated[k - 1], dated[k]
            if un - ln > a.max_gap:
                continue
            cams = {c for c in (lr.get("camera"), ur.get("camera"), r.get("camera")) if c}
            if len(cams) > 1:
                continue  # neighbours from different cameras: the numbers are not one sequence
            ld, ud = parse_iso(lr["date_best"]), parse_iso(ur["date_best"])
            span = ud - ld
            if span.total_seconds() < 0:
                continue  # counter reset or another camera: not monotonic
            if span <= dt.timedelta(days=1):
                conf = "high"
            elif span <= dt.timedelta(days=a.max_days):
                conf = "moderate"
            else:
                continue
            frac = (n - ln) / (un - ln)
            d = (ld + span * frac).replace(microsecond=0)
            rows.append({"path": r["path"], "proposed_date": d.isoformat(), "confidence": conf,
                         "evidence": f"{key[0] or '(no prefix)'} #{n} between #{ln} ({ld:%Y-%m-%d %H:%M}) and #{un} "
                                     f"({ud:%Y-%m-%d %H:%M}){' camera ' + next(iter(cams)) if cams else ''}",
                         "current_date": r.get("date_best"), "current_source": r.get("date_source")})
    out = a.out or os.path.join(os.path.dirname(os.path.abspath(a.catalog)), "date_seq_plan.csv")
    write_csv(out, rows, ["path", "proposed_date", "confidence", "evidence", "current_date", "current_source"])
    c = Counter(r["confidence"] for r in rows)
    log(f"date-seq: {len(rows):,} undated files can be dated (high {c['high']:,}, moderate {c['moderate']:,}) -> {out}")
    if a.write:
        apply_conf = set(a.apply.split(","))
        chosen = {r["path"]: r for r in rows if r["confidence"] in apply_conf}
        _backup_catalog(a.catalog, "dateseq")
        for r in recs:
            p = chosen.get(r["path"])
            if p:
                r["date_best"], r["date_source"] = p["proposed_date"], "sequence"
                r["date_confidence"], r["date_evidence"] = p["confidence"], p["evidence"]
        write_jsonl(a.catalog, recs)
        log(f"  wrote {len(chosen):,} dates into the catalog (confidence: {', '.join(sorted(apply_conf))}). "
            "Re-run `plan` to move them out of _Needs Date.")
    return 0


def cmd_set_dates(a):
    """Store dates the user supplied (contact-sheet review): CSV with path,date[,evidence].
    date may be YYYY, YYYY-MM, YYYY-MM-DD or a full timestamp; precision is kept."""
    recs = load_jsonl(a.catalog)
    by = {r["path"]: r for r in recs}
    rows = read_csv(a.dates)
    ok, bad = 0, []
    for row in rows:
        p, v = norm_rel(row.get("path", "")).strip(), (row.get("date") or "").strip()
        r = by.get(p)
        m = re.fullmatch(r"(\d{4})(?:-(\d{2}))?(?:-(\d{2}))?(?:[T ](\d{2}):(\d{2})(?::(\d{2}))?)?", v)
        if not r or not m:
            bad.append(f"{p} ({'not in catalog' if not r else 'bad date ' + repr(v)})")
            continue
        y, mo, d, hh, mi, ss = m.groups()
        prec = "second" if hh else "day" if d else "month" if mo else "year"
        try:
            when = dt.datetime(int(y), int(mo or 1), int(d or 1), int(hh or 12), int(mi or 0), int(ss or 0))
        except ValueError:
            bad.append(f"{p} (impossible date {v!r})")
            continue
        r["date_best"], r["date_source"], r["date_precision"] = when.isoformat(), "manual", prec
        r["date_evidence"] = row.get("evidence") or "set by user"
        ok += 1
    if ok:
        _backup_catalog(a.catalog, "setdates")
        write_jsonl(a.catalog, recs)
    log(f"set-dates: {ok:,} dates stored, {len(bad)} rejected")
    for b in bad[:20]:
        log(f"  rejected: {b}")
    return 0 if not bad else 1


# ----------------------------------------------------------------------------- catalog (xlsx)
CAT_COLS = ["path", "name", "kind", "size", "date_best", "date_source", "camera", "width", "height", "duration",
            "vcodec", "gps_lat", "gps_lon", "city", "region", "country", "sha1", "issue", "error", "orig_path",
            "people", "date_evidence", "date_confidence", "date_precision", "reorg_rule"]


def cmd_catalog(a):
    from openpyxl import Workbook
    recs = load_jsonl(a.catalog)
    out = a.out or os.path.join(os.path.dirname(os.path.abspath(a.catalog)), "Photo_Catalog.xlsx")
    if os.path.exists(out) and not a.no_backup:
        bk = out.replace(".xlsx", f"_before_{dt.datetime.now():%Y%m%d_%H%M%S}.xlsx")
        shutil.copy2(out, bk)
        log(f"  backup: {bk}")
    wb = Workbook(write_only=True)

    def sheet(title, header, rows):
        ws = wb.create_sheet(title)
        ws.append(header)
        for row in rows:
            ws.append(row)

    media = [r for r in recs if r["kind"] in ("photo", "video")]
    gb = lambda x: round(x / 2 ** 30, 2)
    summ = [["Files", len(recs)], ["Size (GiB)", gb(sum(r["size"] for r in recs))]]
    for kind, c in Counter(r["kind"] for r in recs).most_common():
        summ.append([f"  {kind} files", c])
        summ.append([f"  {kind} GiB", gb(sum(r["size"] for r in recs if r["kind"] == kind))])
    tdate = sum(1 for r in media if r.get("date_source") in TRUE_SOURCES)
    summ += [["Photos+videos with true capture date", f"{tdate} of {len(media)} ({100 * tdate / max(1, len(media)):.0f}%)"],
             ["Photos+videos with GPS", sum(1 for r in media if r.get("gps_lat") is not None)],
             ["Generated", dt.datetime.now().isoformat(timespec="seconds")]]
    sheet("Summary", ["Metric", "Value"], summ)
    sheet("Catalog", CAT_COLS, ([r.get(c) for c in CAT_COLS] for r in recs))
    by = defaultdict(lambda: [0, 0, 0])
    for r in media:
        y = (r.get("date_best") or "")[:4] or "No date"
        by[y][0 if r["kind"] == "photo" else 1] += 1
        by[y][2] += r.get("date_source") in TRUE_SOURCES
    sheet("By Year", ["Year", "Photos", "Videos", "True date"], [[y] + v for y, v in sorted(by.items())])
    cams = defaultdict(list)
    for r in media:
        if r.get("camera") and r.get("date_source") in TRUE_SOURCES:
            cams[r["camera"]].append(r["date_best"])
    sheet("Cameras", ["Camera", "Files", "First", "Last"],
          sorted(([c, len(v), min(v)[:10], max(v)[:10]] for c, v in cams.items()), key=lambda x: -x[1]))
    sheet("Date sources", ["Source", "Files"], Counter(r.get("date_source") for r in media).most_common())
    sheet("Problems", ["path", "kind", "issue", "error", "size"],
          ([r["path"], r["kind"], r.get("issue"), r.get("error"), r["size"]] for r in recs if r.get("issue") or r.get("error")))
    sheet("Junk", ["path", "size"], ([r["path"], r["size"]] for r in recs if r["kind"] == "junk"))
    fold = defaultdict(lambda: [0, 0])
    for r in recs:
        fd = os.path.dirname(r["path"]) or "(root)"
        fold[fd][0] += 1
        fold[fd][1] += r["size"]
    sheet("Folders", ["Folder", "Files", "GiB"], sorted([k, v[0], gb(v[1])] for k, v in fold.items()))
    wb.save(out)
    log(f"wrote {out}  ({len(recs):,} rows)")
    return 0


# ----------------------------------------------------------------------------- reconcile
def cmd_reconcile(a):
    recs = load_jsonl(a.catalog)
    roots = sorted({r["root"] for r in recs})
    bad = 0
    for root in roots:
        cat = {r["path"]: r for r in recs if r["root"] == root}
        disk = {rel: st.st_size for _, rel, st in walk_files(root)}
        missing = [p for p in cat if p not in disk]
        extra = [p for p in disk if p not in cat]
        size = [p for p in cat if p in disk and disk[p] != cat[p]["size"]]
        log(f"{root}: catalog {len(cat):,} | disk {len(disk):,} | missing {len(missing)} | uncataloged {len(extra)} "
            f"| size changed {len(size)}")
        for lab, lst in (("MISSING", missing), ("UNCATALOGED", extra), ("SIZE", size)):
            for p in lst[:a.show]:
                log(f"   {lab}: {p}")
        if a.out:
            write_csv(a.out, [{"root": root, "path": p, "status": s} for s, lst in
                              (("missing", missing), ("uncataloged", extra), ("size_changed", size)) for p in lst],
                      ["root", "path", "status"])
        bad += len(missing) + len(extra) + len(size)
    log("reconcile: " + ("MATCH - catalog equals disk" if not bad else f"{bad} differences (re-run `scan` to refresh)"))
    return 0 if not bad else 2


# ----------------------------------------------------------------------------- image/video comparison
def _norm_img(path, maxside=512):
    Image, ImageOps = pil()
    im = Image.open(P(path))
    im.draft("RGB", (maxside * 2, maxside * 2))
    im = ImageOps.exif_transpose(im).convert("RGB")
    im.thumbnail((maxside, maxside), Image.LANCZOS)
    return im


_IMG_CACHE = {}


def norm_img(path):
    if path not in _IMG_CACHE:
        if len(_IMG_CACHE) > 400:
            _IMG_CACHE.clear()
        try:
            _IMG_CACHE[path] = _norm_img(path)
        except Exception:
            _IMG_CACHE[path] = None
    return _IMG_CACHE[path]


def pixel_diff(pa, pb, rotations=True):
    """Mean absolute RGB difference (0-255) between two images after orientation, trying 0/90/180/270 on b.
    Returns None when aspect ratios never agree (within 3%)."""
    import numpy as np
    Image, _ = pil()
    A, B = norm_img(pa), norm_img(pb)
    if A is None or B is None:
        return None
    best = None
    for rot in ((0, 90, 180, 270) if rotations else (0,)):
        Br = B.rotate(rot, expand=True) if rot else B
        ra, rb = A.width / A.height, Br.width / Br.height
        if abs(ra - rb) / ra > 0.03:
            continue
        w, h = min(A.width, Br.width), min(A.height, Br.height)
        x = np.asarray(A.resize((w, h), Image.LANCZOS), dtype=np.int16)
        y = np.asarray(Br.resize((w, h), Image.LANCZOS), dtype=np.int16)
        d = float(np.abs(x - y).mean())
        best = d if best is None else min(best, d)
    return best


_SH_CACHE = {}


def stream_hash(path):
    """Per-stream MD5 of copied (not re-encoded) packets: ignores container metadata."""
    if path in _SH_CACHE:
        return _SH_CACHE[path]
    try:
        out = subprocess.run(["ffmpeg", "-v", "error", "-i", path, "-map", "0:v?", "-map", "0:a?", "-c", "copy",
                              "-f", "streamhash", "-hash", "md5", "-"], capture_output=True, timeout=1800)
        lines = tuple(l for l in out.stdout.decode().splitlines() if "MD5=" in l)
        res = lines if lines else None
    except Exception:
        res = None
    _SH_CACHE[path] = res
    return res


VFP_FRAMES = 5          # frames sampled per video for its fingerprint
VFP_MAX_DIST = 10       # mean dHash distance (of 64 bits) for "same picture content"
VFP_MIN_FRAMES = 3      # informative frames needed on both sides before a fingerprint counts


def video_fp(path, duration, n=VFP_FRAMES):
    """Visual fingerprint of a video: dHash of n evenly spaced frames (after the player's auto-rotation).
    Near-blank frames (black, finger over lens, flat colour) are stored as None so they never match."""
    if not duration or duration <= 0:
        return None
    out = []
    for k in range(n):
        t = duration * (k + 1) / (n + 1)
        try:
            p = subprocess.run(["ffmpeg", "-v", "error", "-ss", f"{t:.3f}", "-i", path, "-frames:v", "1",
                                "-vf", "scale=9:8:flags=area,format=gray", "-f", "rawvideo", "-"],
                               capture_output=True, timeout=180)
            px = list(p.stdout)
        except Exception:
            px = []
        if len(px) != 72:
            out.append(None)
            continue
        mean = sum(px) / 72
        if sum((v - mean) ** 2 for v in px) / 72 < 25:  # flat frame: no information
            out.append(None)
            continue
        bits = 0
        for row in range(8):
            for col in range(8):
                bits = (bits << 1) | (1 if px[row * 9 + col] > px[row * 9 + col + 1] else 0)
        out.append(f"{bits:016x}")
    return out


def vfp_dist(a, b):
    """Mean dHash distance over frames informative in both fingerprints; None if fewer than VFP_MIN_FRAMES."""
    if not a or not b or len(a) != len(b):
        return None
    ds = [hamming(x, y) for x, y in zip(a, b) if x and y]
    if len(ds) < VFP_MIN_FRAMES:
        return None
    return sum(ds) / len(ds)


def video_same_content(ri, rj):
    """('identical-streams'|'look-alike'|'different'|'unknown', detail). Never guesses from name/duration alone."""
    d = vfp_dist(ri.get("vfp"), rj.get("vfp"))
    if d is not None and d > VFP_MAX_DIST:
        return "different", f"fingerprint distance {d:.1f}"
    hi, hj = stream_hash(abspath_of(ri)), stream_hash(abspath_of(rj))
    if hi and hi == hj:
        return "identical-streams", "identical video/audio streams"
    if d is not None:
        return "look-alike", f"fingerprint distance {d:.1f} (re-encode/resize of the same clip)"
    return "unknown", "no usable fingerprint (blank frames or unreadable) - check by eye"


def transposed(ri, rj):
    """True when one image looks like a 90-degree-rotated copy of the other (dHash can't see that)."""
    try:
        a1 = ri["width"] / ri["height"]
        a2 = rj["height"] / rj["width"]
        return abs(a1 - a2) / a1 < 0.03 and abs(a1 - 1) > 0.05
    except (KeyError, TypeError, ZeroDivisionError):
        return False


def is_dump_folder(rel, dump_re):
    return bool(dump_re.search(os.path.dirname(rel)))


def keeper_key(r, dump_re):
    return (pixels(r), 0 if is_dump_folder(r["path"], dump_re) else 1, 1 if r.get("date_source") in TRUE_SOURCES else 0,
            1 if r.get("gps_lat") is not None else 0, r["size"], -len(r["path"]))


class UF:
    def __init__(self):
        self.p = {}

    def find(self, x):
        self.p.setdefault(x, x)
        while self.p[x] != x:
            self.p[x] = self.p[self.p[x]]
            x = self.p[x]
        return x

    def union(self, a, b):
        ra, rb = self.find(a), self.find(b)
        if ra != rb:
            self.p[rb] = ra


def cmd_dupes(a):
    recs = load_jsonl(a.catalog)
    dump_re = re.compile(a.dump_re)
    qname = a.quarantine
    media = [r for r in recs if r["kind"] in ("photo", "video") and not r["path"].startswith(qname + "/")]
    by_id = {i: r for i, r in enumerate(media)}
    uf = UF()
    tier = {}  # frozenset pair -> tier

    def link(i, j, t):
        uf.union(i, j)
        tier[frozenset((i, j))] = t

    # Tier A exact
    bysha = defaultdict(list)
    for i, r in by_id.items():
        if r.get("sha1"):
            bysha[r["sha1"]].append(i)
    na = 0
    for ids in bysha.values():
        for j in ids[1:]:
            link(ids[0], j, "A")
            na += 1
    log(f"Tier A (byte-identical): {na} extra copies")
    similar = []
    # candidate buckets: same capture second, same base name
    def buckets(kind, keyfn):
        b = defaultdict(list)
        for i, r in by_id.items():
            if r["kind"] == kind:
                for k in keyfn(r):
                    if k:
                        b[k].append(i)
        return b

    def pkeys(r):
        ks = []
        if r.get("date_source") in TRUE_SOURCES:
            ks.append(("t", r["date_best"][:19], r["kind"]))
        stem = os.path.splitext(r["name"])[0].lower()
        stem = re.sub(r"(_\d|\s\(\d+\)|\s\([a-z]{2}\)|-edited|_edit(ed)?)$", "", stem)
        ks.append(("n", stem, r["kind"]))
        return ks

    checked = set()
    # photos
    pb = buckets("photo", pkeys)
    nc = 0
    t0 = time.time()
    for ids in pb.values():
        if len(ids) < 2 or len(ids) > a.max_bucket:
            continue
        for x in range(len(ids)):
            for y in range(x + 1, len(ids)):
                i, j = ids[x], ids[y]
                key = frozenset((i, j))
                if key in checked or uf.find(i) == uf.find(j):
                    continue
                checked.add(key)
                ri, rj = by_id[i], by_id[j]
                hd = hamming(ri.get("dhash"), rj.get("dhash"))
                if hd > a.dhash and not transposed(ri, rj):
                    continue
                d = pixel_diff(abspath_of(ri), abspath_of(rj))
                if d is not None and d <= a.pixel:
                    link(i, j, "C")
                    nc += 1
                else:
                    similar.append({"tier": "D", "a": ri["path"], "b": rj["path"], "dhash_dist": hd,
                                    "pixel_diff": None if d is None else round(d, 2),
                                    "a_res": res(ri),
                                    "b_res": res(rj)})
    log(f"Tier C (re-saved photos, pixel <= {a.pixel}/255): {nc} links  [{time.time() - t0:.0f}s]")
    # videos: candidate pairs = same name/second (bucketed) + any two clips of near-equal length whose
    # fingerprints look alike (catches renamed/undated copies). A content check always decides.
    vids = sorted((i for i, r in by_id.items() if r["kind"] == "video" and r.get("duration")),
                  key=lambda i: by_id[i]["duration"])
    pairs = []
    for ids in buckets("video", pkeys).values():
        if 2 <= len(ids) <= a.max_bucket:
            pairs += [(ids[x], ids[y], "name/second") for x in range(len(ids)) for y in range(x + 1, len(ids))]
    for x, i in enumerate(vids):  # sliding window on duration
        for j in vids[x + 1:]:
            if by_id[j]["duration"] - by_id[i]["duration"] > a.duration_tol:
                break
            dd = vfp_dist(by_id[i].get("vfp"), by_id[j].get("vfp"))
            if dd is not None and dd <= VFP_MAX_DIST:
                pairs.append((i, j, "look-alike"))
    nb = 0
    for i, j, why in pairs:
        key = frozenset((i, j))
        if key in checked or uf.find(i) == uf.find(j):
            continue
        checked.add(key)
        ri, rj = by_id[i], by_id[j]
        if ri.get("duration") is None or rj.get("duration") is None:
            continue
        if abs(ri["duration"] - rj["duration"]) > a.duration_tol:
            continue
        verdict, detail = video_same_content(ri, rj)
        if verdict == "identical-streams":
            link(i, j, "B")
            nb += 1
        elif verdict in ("look-alike", "unknown"):
            similar.append({"tier": "D-video", "a": ri["path"], "b": rj["path"], "a_res": res(ri), "b_res": res(rj),
                            "note": f"{why}: {detail}"})
        # 'different' = same name/length but another clip: not a duplicate, not listed
    log(f"Tier B (same video streams): {nb} links")
    groups = defaultdict(list)
    for i in by_id:
        if i in uf.p:
            groups[uf.find(i)].append(i)
    rows, plan, held = [], [], []
    for gi, (g, ids) in enumerate(sorted(groups.items(), key=lambda kv: by_id[kv[0]]["path"]), 1):
        if len(ids) < 2:
            continue
        keep = max(ids, key=lambda i: keeper_key(by_id[i], dump_re))
        k = by_id[keep]
        for i in ids:
            r = by_id[i]
            t = "A" if r.get("sha1") and r.get("sha1") == k.get("sha1") else (
                "B" if r["kind"] == "video" else "C")
            row = {"group": gi, "role": "keeper" if i == keep else "extra", "tier": "" if i == keep else t,
                   "path": r["path"], "keeper": k["path"], "size": r["size"], "res": res(r),
                   "date": r.get("date_best"), "date_source": r.get("date_source"),
                   "gps": "yes" if r.get("gps_lat") is not None else ""}
            if i != keep:
                warn = []
                if r.get("date_source") in TRUE_SOURCES and k.get("date_source") not in TRUE_SOURCES:
                    warn.append("extra has a true date the keeper lacks")
                if r.get("gps_lat") is not None and k.get("gps_lat") is None:
                    warn.append("extra has GPS the keeper lacks")
                if pixels(r) > pixels(k):
                    warn.append("extra is larger resolution")
                row["warning"] = "; ".join(warn)
                if warn:
                    held.append(row)
                else:
                    plan.append({"action": "quarantine", "src": abspath_of(r),
                                 "dst": os.path.join(r["root"], qname, *r["path"].split("/")),
                                 "keeper": abspath_of(k), "keeper_sha1": k.get("sha1", ""), "tier": t,
                                 "group": gi, "src_sha1": r.get("sha1", "")})
            rows.append(row)
    # one 'similar' row per pair of duplicate groups (members of one group are already handled together)
    pid = {r["path"]: i for i, r in by_id.items()}
    seen_pairs, slim = set(), []
    for s_ in similar:
        ka = uf.find(pid[s_["a"]]) if pid[s_["a"]] in uf.p else pid[s_["a"]]
        kb = uf.find(pid[s_["b"]]) if pid[s_["b"]] in uf.p else pid[s_["b"]]
        k2 = frozenset((ka, kb))
        if ka != kb and k2 not in seen_pairs:
            seen_pairs.add(k2)
            slim.append(s_)
    similar = slim
    outdir = a.outdir or os.path.dirname(os.path.abspath(a.catalog))
    write_csv(os.path.join(outdir, "dupes_groups.csv"), rows)
    write_csv(os.path.join(outdir, "dupes_plan.csv"), plan)
    write_csv(os.path.join(outdir, "dupes_similar.csv"), similar)
    ext_bytes = sum(os.path.getsize(P(p["src"])) for p in plan if os.path.exists(P(p["src"])))
    log(f"groups {sum(1 for r in rows if r['role'] == 'keeper')} | quarantine plan {len(plan)} files "
        f"({ext_bytes / 2 ** 30:.2f} GiB) | held for review {len(held)} | similar pairs (Tier D) {len(similar)}")
    log(f"wrote dupes_groups.csv, dupes_plan.csv, dupes_similar.csv in {outdir}")
    log("next: review dupes_groups.csv, then `apply dupes_plan.csv` (moves to quarantine, nothing deleted)")
    return 0


# ----------------------------------------------------------------------------- compare (new source vs library)
def cmd_compare(a):
    lib = [r for r in load_jsonl(a.library) if r["kind"] in ("photo", "video")
           and not r["path"].startswith("_Duplicates_Quarantine/")]
    src = load_jsonl(a.source)
    libsha = {r["sha1"]: r for r in lib if r.get("sha1")}
    allsha = set(libsha)
    for extra in (a.also or []):
        allsha |= {r["sha1"] for r in load_jsonl(extra) if r.get("sha1")}
    bytime, byname = defaultdict(list), defaultdict(list)
    bands = defaultdict(list)
    vbydur = defaultdict(list)
    for r in lib:
        if r.get("date_source") in TRUE_SOURCES:
            bytime[(r["kind"], r["date_best"][:19])].append(r)
        byname[(r["kind"], r["name"].lower())].append(r)
        if r["kind"] == "photo" and r.get("dhash"):
            h = r["dhash"]
            for b in range(8):
                bands[(b, h[b * 2:b * 2 + 2])].append(r)
        if r["kind"] == "video" and r.get("duration") and r.get("vfp"):
            vbydur[int(r["duration"])].append(r)
    seen = set()
    out = []
    t0, last = time.time(), time.time()
    for n, s in enumerate(src, 1):
        row = {"src": abspath_of(s), "rel": s["path"], "kind": s["kind"], "class": "", "match": "", "detail": "",
               "src_res": res(s), "size": s["size"], "date": s.get("date_best"),
               "date_source": s.get("date_source")}
        if time.time() - last > 30:
            log(f"  {n:,}/{len(src):,}")
            last = time.time()
        if s["kind"] == "junk":
            row["class"] = "junk"
        elif s["kind"] not in ("photo", "video"):
            row["class"] = "non-media" if s["kind"] != "problem" else "problem"
            row["detail"] = s.get("issue", "")
        elif s.get("from_inventory") and not s.get("sha1"):
            row["class"] = "new"
            row["detail"] = ("unique size in the inventory - new ONLY if the library was among the "
                             "inventoried roots; stage and `scan` it for dates/upgrade checks")
        elif s.get("sha1") in allsha:
            row["class"] = "identical"
            m = libsha.get(s["sha1"])
            row["match"] = abspath_of(m) if m else "(other catalog)"
        elif s.get("sha1") in seen:
            row["class"] = "internal-repeat"
        else:
            cands = []
            if s.get("date_source") in TRUE_SOURCES:
                cands += bytime.get((s["kind"], s["date_best"][:19]), [])
            cands += byname.get((s["kind"], s["name"].lower()), [])
            uniq = {id(c): c for c in cands}.values()
            found = None
            unsure = None
            for c in uniq:
                if s["kind"] == "photo":
                    if hamming(s.get("dhash"), c.get("dhash")) > 12 and not transposed(s, c):
                        continue
                    d = pixel_diff(abspath_of(s), abspath_of(c))
                    if d is not None and d <= a.pixel:
                        found = (c, f"pixel diff {d:.2f}")
                        break
                else:
                    if not (s.get("duration") and c.get("duration") and abs(s["duration"] - c["duration"]) <= 0.5):
                        continue
                    verdict, why = video_same_content(s, c)  # never trust name + length alone
                    if verdict in ("identical-streams", "look-alike"):
                        found = (c, why)
                        break
                    if verdict == "unknown" and unsure is None:
                        unsure = (c, why)
            if not found and s["kind"] == "video" and s.get("vfp"):
                best = None  # look-alike search: renamed/undated copies of the same clip
                for c in vbydur.get(int(s.get("duration") or -9), []) + vbydur.get(int(s.get("duration") or -9) - 1, []) \
                        + vbydur.get(int(s.get("duration") or -9) + 1, []):
                    if abs((c.get("duration") or 0) - s["duration"]) > 1.0:
                        continue
                    dd = vfp_dist(s["vfp"], c.get("vfp"))
                    if dd is not None and dd <= VFP_MAX_DIST and (best is None or dd < best[1]):
                        best = (c, dd)
                if best:
                    unsure = (best[0], f"look-alike by fingerprint (distance {best[1]:.1f}), different name/time")
            if not found and unsure:
                c, why = unsure
                row["class"], row["match"], row["detail"] = "likely-duplicate", abspath_of(c), why
                row["match_res"] = res(c)
                if pixels(s) > pixels(c) * 1.1:
                    row["detail"] += " ; source is larger"
                out.append(row)
                if s.get("sha1"):
                    seen.add(s["sha1"])
                continue
            if found:
                c, why = found
                row["match"], row["detail"] = abspath_of(c), why
                row["match_res"] = res(c)
                sp, cp = pixels(s), pixels(c)
                better = sp > cp * 1.1 or (sp >= cp * 0.99 and s["size"] > c["size"] * 1.25 and abs(s["size"] - c["size"]) > 20000)
                if s["kind"] == "video" and sp == cp and (s.get("vbitrate") or 0) > (c.get("vbitrate") or 0) * 1.25:
                    better = True
                row["class"] = "upgrade" if better else "duplicate"
            elif s["kind"] == "photo" and s.get("dhash"):
                best = None
                pool = {}
                for b in range(8):
                    for c in bands.get((b, s["dhash"][b * 2:b * 2 + 2]), []):
                        pool[id(c)] = c
                for c in pool.values():
                    hd = hamming(s["dhash"], c.get("dhash"))
                    if hd <= a.dhash_loose and (best is None or hd < best[1]):
                        best = (c, hd)
                if best:
                    c, hd = best
                    row["class"] = "likely-duplicate"
                    row["match"] = abspath_of(c)
                    row["match_res"] = res(c)
                    row["detail"] = f"dhash distance {hd}" + (" ; source is larger" if pixels(s) > pixels(c) * 1.1 else "")
                else:
                    row["class"] = "new"
            else:
                row["class"] = "new"
        if s.get("sha1"):
            seen.add(s["sha1"])
        out.append(row)
    write_csv(a.out, out)
    cnt = Counter(r["class"] for r in out)
    size = defaultdict(int)
    for r in out:
        size[r["class"]] += int(r["size"])
    log(f"compare: {len(out):,} source files in {time.time() - t0:.0f}s")
    for k, v in cnt.most_common():
        log(f"  {k:17s} {v:7,}  {size[k] / 2 ** 30:8.2f} GiB")
    log(f"wrote {a.out}")
    return 0


# ----------------------------------------------------------------------------- geocoding
class Geo:
    def __init__(self, geonames=None):
        self.cities = None
        self.rg = None
        if geonames:
            self.grid = defaultdict(list)
            n = 0
            with open(geonames, encoding="utf-8") as f:
                for line in f:
                    p = line.rstrip("\n").split("\t")
                    if len(p) < 15 or not p[14].isdigit():
                        continue
                    if p[7].startswith("PPLX"):  # neighborhoods
                        continue
                    la, lo = float(p[4]), float(p[5])
                    c = (p[1], p[10], p[8], la, lo, int(p[14]), p[11])
                    self.grid[(int(la), int(lo))].append(c)
                    n += 1
            self.cities = n
            log(f"  geonames: {n:,} places loaded")
        try:
            import reverse_geocoder
            self.rg = reverse_geocoder
        except ImportError:
            if not geonames:
                log("  ! no geocoder: install reverse_geocoder or run `geonames` (location labels disabled)")
        self.cache = {}

    def place(self, la, lo):
        key = (round(la, 3), round(lo, 3))
        if key in self.cache:
            return self.cache[key]
        res = None
        if self.cities:
            best = None
            for dla in (-1, 0, 1):
                for dlo in (-1, 0, 1):
                    for c in self.grid.get((int(la) + dla, int(lo) + dlo), []):
                        d = haversine_km((la, lo), (c[3], c[4]))
                        if d > 40:
                            continue
                        score = c[5] / (d + 5) ** 2
                        if best is None or score > best[0]:
                            best = (score, c)
            if best:
                c = best[1]
                res = (c[0], c[1], c[2], "")
        if res is None and self.rg:  # no big city within 40 km: nearest small town
            g = self.rg.search([(la, lo)], mode=1, verbose=False)[0]
            res = (g["name"], g["admin1"], g["cc"], g.get("admin2", ""))
        self.cache[key] = res
        return res

    @staticmethod
    def label(p):
        if not p:
            return None
        city, adm, cc = p[:3]
        if cc == "US":
            st = adm if len(adm) == 2 else US_STATES.get(adm, adm)
            return f"{city}, {st}"
        return f"{city}, {COUNTRIES.get(cc, cc)}"

    @staticmethod
    def region(p):
        if not p:
            return None
        city, adm, cc = p[:3]
        if cc == "US":
            return adm if len(adm) == 2 else US_STATES.get(adm, adm)
        return COUNTRIES.get(cc, cc)


def cmd_geonames(a):
    """Download GeoNames cities15000 (CC-BY 4.0) for population-weighted place names."""
    import io, urllib.request, zipfile
    os.makedirs(a.dir, exist_ok=True)
    url = "https://download.geonames.org/export/dump/cities15000.zip"
    log(f"downloading {url}")
    data = urllib.request.urlopen(url, timeout=120).read()
    zipfile.ZipFile(io.BytesIO(data)).extract("cities15000.txt", a.dir)
    out = os.path.join(a.dir, "cities15000.txt")
    log(f"wrote {out} ({os.path.getsize(out) / 1e6:.1f} MB). Data: GeoNames.org, CC-BY 4.0")
    return 0


# ----------------------------------------------------------------------------- plan (reorganization)
def easter(y):
    a = y % 19; b = y // 100; c = y % 100; d = b // 4; e = b % 4; f = (b + 8) // 25; g = (b - f + 1) // 3
    h = (19 * a + b - d - g + 15) % 30; i = c // 4; k = c % 4; l = (32 + 2 * e + 2 * i - h - k) % 7
    m = (a + 11 * h + 22 * l) // 451; mo = (h + l - 7 * m + 114) // 31; da = ((h + l - 7 * m + 114) % 31) + 1
    return dt.date(y, mo, da)


def nth_weekday(y, mo, wd, n):
    d = dt.date(y, mo, 1)
    d += dt.timedelta(days=(wd - d.weekday()) % 7)
    return d + dt.timedelta(weeks=n - 1)


def holiday(d):
    fixed = {(1, 1): "New Year's Day", (2, 14): "Valentine's Day", (7, 4): "Fourth of July", (10, 31): "Halloween",
             (12, 24): "Christmas Eve", (12, 25): "Christmas", (12, 31): "New Year's Eve"}
    if (d.month, d.day) in fixed:
        return fixed[(d.month, d.day)]
    if d == easter(d.year):
        return "Easter"
    if d == nth_weekday(d.year, 11, 3, 4):
        return "Thanksgiving"
    if d == nth_weekday(d.year, 5, 6, 2):
        return "Mother's Day"
    if d == nth_weekday(d.year, 6, 6, 3):
        return "Father's Day"
    return None


MONTHS = r"(jan(uary)?|feb(ruary)?|mar(ch)?|apr(il)?|may|jun(e)?|jul(y)?|aug(ust)?|sep(t(ember)?)?|oct(ober)?|nov(ember)?|dec(ember)?)"
DATE_TEXT = [re.compile(p, re.I) for p in [
    r"\b(19|20)\d\d[-_. ]?\d{1,2}[-_. ]?\d{1,2}\b", r"\b\d{1,2}[-_./]\d{1,2}[-_./](\d{4}|\d{2})\b",
    r"\b\d{6,8}\b", rf"\b{MONTHS}\.?\s*\d{{1,2}}(st|nd|rd|th)?,?\s*(19|20)?\d\d\b", rf"\b\d{{1,2}}\s*{MONTHS}\s*(19|20)\d\d\b",
    rf"\b{MONTHS}\.?\s*(19|20)\d\d\b", rf"\b{MONTHS}\.?\s*\d{{1,2}}(st|nd|rd|th)?\b", r"\b(19|20)\d\d\b"]]
BAD_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')


def clean_name(s):
    t = s
    for rx in DATE_TEXT:
        t = rx.sub(" ", t)
    t = re.sub(r"[\s_\-.,]+", " ", t).strip(" -_.,()")
    t = re.sub(r"\(\s*\)", "", t).strip()
    return t


def safe_component(s):
    s = BAD_CHARS.sub("-", s).strip().rstrip(".")
    return s or "_"


def densest_window(dates, days):
    ds = sorted(dates)
    best, bi, j = 0, 0, 0
    for i in range(len(ds)):
        while ds[i] - ds[j] > dt.timedelta(days=days):
            j += 1
        if i - j + 1 > best:
            best, bi = i - j + 1, j
    return best, ds[bi], ds[min(len(ds) - 1, bi + best - 1)]


def median_point(pts):
    return (statistics.median(p[0] for p in pts), statistics.median(p[1] for p in pts))


def cmd_plan(a):
    cfg = json.load(open(a.config, encoding="utf-8"))
    recs = load_jsonl(a.catalog)
    if a.include_from:
        classes = set(a.classes.split(","))
        keep = {r["rel"] for r in read_csv(a.include_from) if r.get("class") in classes}
        recs = [r for r in recs if r["path"] in keep]
        log(f"  limited to {len(recs):,} files with class {sorted(classes)} in {os.path.basename(a.include_from)}")
        if not recs:
            log("plan: nothing to do")
            return 0
    trusted = set(cfg.get("trusted_date_sources", ["exif", "video", "sequence", "manual"]))
    dest_root = os.path.abspath(cfg.get("dest_root") or recs[0]["root"])
    dump_re = re.compile(cfg.get("dump_pattern", DEFAULT_DUMP_RE))
    share = cfg.get("event_min_share", 0.8)
    window = cfg.get("event_window_days", 14)
    busy_min = cfg.get("busy_day_min", 20)
    min_trip = cfg.get("min_trip_files", 5)
    unit_depth = cfg.get("unit_depth", 1)
    homes = cfg.get("homes", [])
    for h in homes:
        h["_from"] = dt.date.fromisoformat(h.get("from", "1900-01-01"))
        h["_to"] = dt.date.fromisoformat(h.get("to", "2999-12-31"))
    side = {k: v for k, v in cfg.get("side_folders", {}).items()}  # regex on rel path -> dest prefix
    side_rx = [(re.compile(k), v) for k, v in side.items()]
    skip_rx = re.compile(cfg["skip_pattern"]) if cfg.get("skip_pattern") else None
    gpath = cfg.get("geonames")
    if not gpath:
        auto = os.path.join(os.path.dirname(os.path.abspath(a.catalog)), "cities15000.txt")
        gpath = auto if os.path.exists(auto) else None
    if homes and not gpath:
        log("  ! no cities15000.txt: using reverse_geocoder, which can pick neighborhood names "
            "(run `pcat.py geonames --dir <_Catalog>` for population-weighted city names)")
    geo = Geo(gpath) if homes else None
    if not homes:
        log("  no 'homes' in config: location labels and trips disabled")

    def is_away(point, day):
        if not homes or point is None:
            return False
        act = [h for h in homes if h["_from"] <= day <= h["_to"]]
        if not act:
            return False
        return all(haversine_km(point, (h["lat"], h["lon"])) > h.get("radius_km", 80) for h in act)

    def loc_label(point):
        if geo is None or point is None:
            return None
        return Geo.label(geo.place(*point))

    plan = {}
    rules = {}

    def put(r, dest_rel, rule, loc=None):
        plan[r["path"]] = (dest_rel, rule, loc)

    items = []
    for r in recs:
        if r["path"].startswith("_") or (skip_rx and skip_rx.search(r["path"])) or \
                (cfg.get("skip_organized", True) and re.match(r"(19|20)\d\d/", r["path"])):
            continue  # already-organized side folders / explicitly skipped
        if r["kind"] == "junk":
            put(r, "_Junk/" + r["path"], "junk")
            continue
        hit = next(((rx, v) for rx, v in side_rx if rx.search(r["path"])), None)
        if hit:
            put(r, f"{hit[1]}/{r['path']}", f"side folder ({hit[0].pattern})")
            continue
        if r["kind"] not in ("photo", "video"):
            put(r, "_Review/" + r["path"], f"not photo/video ({r['kind']})")
            continue
        if r.get("date_source") in trusted and r.get("date_precision") in ("year", "month"):
            d = parse_iso(r["date_best"])
            top = r["path"].split("/")[0] if "/" in r["path"] else ""
            label = "" if (not top or dump_re.search(top)) else " " + safe_component(clean_name(top) or top)
            if r["date_precision"] == "year":
                folder = f"{d:%Y}/{d:%Y} (month unknown){label}"
            else:  # month only: the folder's own month group (topic-style), or the plain month folder for dumps
                folder = f"{d:%Y}/{d:%Y-%m}{label}"
            put(r, f"{folder}/{r['name']}", f"user date ({r['date_precision']} only)")
            continue
        items.append(r)

    def unit_of(rel):
        parts = rel.split("/")
        return "/".join(parts[:unit_depth]) if len(parts) > unit_depth else ""

    units = defaultdict(list)
    for r in items:
        units[unit_of(r["path"])].append(r)
    events = []  # dicts: folder, start, end
    topic_units, dump_units = [], []
    for u, rs in units.items():
        name = u.split("/")[-1] if u else ""
        dated = [parse_iso(r["date_best"]) for r in rs if r.get("date_source") in trusted]
        if not u or dump_re.search(u):
            dump_units.append(u)
            continue
        if not dated:
            topic_units.append(u)
            continue
        cnt, ws, we = densest_window(dated, window)
        if cnt / len(dated) >= share:
            pts = [(r["gps_lat"], r["gps_lon"]) for r in rs if r.get("gps_lat") is not None]
            pt = median_point(pts) if pts else None
            label = clean_name(name) or "Event"
            folder = f"{ws:%Y-%m-%d} {label}"
            loc = None
            if pt and is_away(pt, ws.date()):
                pl = geo.place(*pt) if geo else None
                loc = Geo.label(pl)
                words = [loc.split(",")[0]] if loc else []
                if pl and len(pl) > 3 and pl[3]:
                    words.append(pl[3].replace(" County", "").replace(" Parish", ""))
                if loc and not any(w and w.lower() in label.lower() for w in words):
                    folder += f" - {loc}"
            folder = f"{ws:%Y}/{safe_component(folder)}"
            events.append({"folder": folder, "start": ws.date(), "end": we.date(), "unit": u})
            for r in rs:
                sub = r["path"][len(u) + 1:] if u else r["path"]
                put(r, f"{folder}/{sub}", "event folder", loc)
        else:
            topic_units.append(u)
    # topics: split by month
    for u in topic_units:
        label = safe_component(clean_name(u.split("/")[-1]) or u.split("/")[-1])
        for r in units[u]:
            sub = r["path"][len(u) + 1:]
            if r.get("date_source") in trusted:
                d = parse_iso(r["date_best"])
                put(r, f"{d:%Y}/{d:%Y-%m} {label}/{sub}", "topic folder split by month")
            else:
                put(r, f"_Needs Date/{r['path']}", f"no true date (best guess {r.get('date_source')} {(r.get('date_best') or '')[:10]})")
    # existing library events (for imports): folders named YYYY/YYYY-MM-DD <name> in the library catalog
    if cfg.get("library_catalog"):
        spans = defaultdict(list)
        for r in load_jsonl(cfg["library_catalog"]):
            m = re.match(r"((?:19|20)\d\d/(?:19|20)\d\d-\d\d-\d\d [^/]+)/", r["path"])
            if m and r.get("date_source") in trusted:
                spans[m.group(1)].append(parse_iso(r["date_best"]).date())
        added = 0
        for folder, ds in spans.items():
            ds.sort()
            lo, hi = ds[int(len(ds) * 0.05)], ds[min(len(ds) - 1, int(len(ds) * 0.95))]
            if (hi - lo).days <= window:
                events.append({"folder": folder, "start": lo, "end": hi, "unit": None})
                added += 1
        log(f"  {added} existing library event folders available for same-day joining")
        # a source event overlapping an existing library event merges into it (library wins)
        lib_events = [e for e in events if e["unit"] is None]
        merged = 0
        for e in [e for e in events if e["unit"] is not None]:
            hit = next((l for l in lib_events if e["start"] <= l["end"] and e["end"] >= l["start"]), None)
            if hit:
                for rel, (dest, rule, loc) in list(plan.items()):
                    if dest.startswith(e["folder"] + "/"):
                        plan[rel] = (hit["folder"] + dest[len(e["folder"]):], "event merged into existing library event", loc)
                events.remove(e)
                merged += 1
        if merged:
            log(f"  {merged} source event folders merged into existing library events")
    # dumps: join events by date, else day analysis
    ev_by_day = {}
    # longest spans first so shorter (more specific) events win a shared day; library events win over new ones
    for e in sorted(events, key=lambda e: (e["unit"] is None, -(e["end"] - e["start"]).days)):
        d = e["start"]
        while d <= e["end"]:
            ev_by_day[d] = e
            d += dt.timedelta(days=1)
    loose = defaultdict(list)
    for u in dump_units:
        for r in units[u]:
            if r.get("date_source") not in trusted:
                put(r, f"_Needs Date/{r['path']}", f"phone/backup dump, no true date ({r.get('date_source')})")
                continue
            d = parse_iso(r["date_best"]).date()
            if d in ev_by_day:
                put(r, f"{ev_by_day[d]['folder']}/{r['name']}", "phone dump joined same-day event")
            else:
                loose[d].append(r)
    # day locations
    dayinfo = {}
    for d, rs in loose.items():
        pts = [(r["gps_lat"], r["gps_lon"]) for r in rs if r.get("gps_lat") is not None]
        if pts:
            mp = median_point(pts)
            away = is_away(mp, d)
            dayinfo[d] = (mp, away)
        else:
            dayinfo[d] = (None, False)
    # trips: consecutive away days (gap <= 1 day)
    away_days = sorted(d for d, (mp, aw) in dayinfo.items() if aw)
    trips, cur = [], []
    for d in away_days:
        if cur and (d - cur[-1]).days > 2:
            trips.append(cur)
            cur = []
        cur.append(d)
    if cur:
        trips.append(cur)
    trip_of = {}
    for t in trips:
        n = sum(len(loose[d]) for d in t)
        if n < min_trip:
            continue
        labs = Counter()
        regions = set()
        for d in t:
            mp = dayinfo[d][0]
            p = geo.place(*mp) if geo else None
            lab = Geo.label(p)
            if lab:
                labs[lab] += len(loose[d])
                regions.add(Geo.region(p))
        if not labs:
            lab = "Trip"
        elif len(regions) > 1 and labs.most_common(1)[0][1] < 0.6 * sum(labs.values()):
            lab = " & ".join(sorted(regions))
        else:
            lab = labs.most_common(1)[0][0]
        folder = f"{t[0]:%Y}/{safe_component(f'{t[0]:%Y-%m-%d} {lab}')}"
        for d in t:
            trip_of[d] = (folder, lab)
    for d, rs in loose.items():
        if d in trip_of:
            folder, lab = trip_of[d]
            for r in rs:
                put(r, f"{folder}/{r['name']}", "phone dump: trip (away from home)", lab)
        elif len(rs) >= busy_min:
            h = holiday(d)
            folder = f"{d:%Y}/{d:%Y-%m-%d} {h}" if h else f"{d:%Y}/{d:%Y-%m-%d}"
            for r in rs:
                put(r, f"{folder}/{r['name']}", "phone dump: busy day" + (f" ({h})" if h else " (needs a name)"))
        else:
            for r in rs:
                put(r, f"{d:%Y}/{d:%Y-%m}/{r['name']}", "phone dump: month folder")
    # overrides
    ov = cfg.get("overrides", {})
    fold_map = ov.get("folders", {})
    file_map = ov.get("files", {})
    for rel, (dest, rule, loc) in list(plan.items()):
        if rel in file_map:
            plan[rel] = (file_map[rel], "override (file)", loc)
            continue
        for old, new in fold_map.items():
            if dest == old or dest.startswith(old + "/"):
                plan[rel] = (new + dest[len(old):], rule + " + override (folder)", loc)
                break
    # sanitize + collisions
    used = {}
    rows = []
    long_paths = 0
    bymap = {r["path"]: r for r in recs}
    for rel in sorted(plan, key=str.lower):
        dest, rule, loc = plan[rel]
        parts = [safe_component(p) for p in dest.split("/")]
        dest = "/".join(parts)
        base, ext = os.path.splitext(dest)
        cand, n = dest, 2
        r0 = bymap[rel]
        while cand.lower() in used or (
                os.path.exists(P(os.path.join(dest_root, *cand.split("/")))) and
                os.path.normcase(os.path.join(dest_root, *cand.split("/"))) != os.path.normcase(abspath_of(r0))):
            cand = f"{base}_{n}{ext}"
            n += 1
        used[cand.lower()] = rel
        r = bymap[rel]
        src = abspath_of(r)
        dst = os.path.join(dest_root, *cand.split("/"))
        if len(dst) > a.max_path:
            long_paths += 1
        rows.append({"action": a.action, "src": src, "dst": dst, "rule": rule, "renamed": "yes" if cand != dest else "",
                     "date": r.get("date_best"), "date_source": r.get("date_source"), "location": loc or "",
                     "rel": rel, "new_rel": cand, "size": r["size"], "sha1": r.get("sha1", "")})
    unchanged = [x for x in rows if os.path.normcase(os.path.abspath(x["src"])) == os.path.normcase(os.path.abspath(x["dst"]))]
    if unchanged:
        un = {id(x) for x in unchanged}
        rows = [x for x in rows if id(x) not in un]
    out = a.out
    write_csv(out, rows)
    folders = defaultdict(lambda: {"files": 0, "bytes": 0, "rules": Counter()})
    for x in rows:
        fd = os.path.dirname(x["new_rel"])
        folders[fd]["files"] += 1
        folders[fd]["bytes"] += int(x["size"])
        folders[fd]["rules"][x["rule"]] += 1
    fo = os.path.splitext(out)[0] + "_folders.csv"
    write_csv(fo, [{"folder": k, "files": v["files"], "GiB": round(v["bytes"] / 2 ** 30, 3),
                    "main_rule": v["rules"].most_common(1)[0][0]} for k, v in sorted(folders.items())])
    # checks
    total_in = sum(int(x["size"]) for x in rows)
    log(f"plan: {len(rows):,} {a.action} actions ({len(unchanged)} already in place) -> {len(folders):,} folders | "
        f"{total_in / 2 ** 30:.2f} GiB")
    rc = Counter(x["rule"].split(" (")[0] for x in rows)
    for k, v in rc.most_common():
        log(f"  {v:7,}  {k}")
    log(f"  events {len(events)} | trips {len(set(f for f, _ in trip_of.values()))} | renamed for clash "
        f"{sum(1 for x in rows if x['renamed'])} | paths > {a.max_path} chars: {long_paths}")
    if len({x['dst'].lower() for x in rows}) != len(rows):
        log("  !! destination collision detected - do not apply")
        return 3
    log(f"wrote {out} and {fo}.  Review, adjust config/overrides, re-run; then `apply`.")
    return 0


# ----------------------------------------------------------------------------- apply
def same_device(a_path, b_dir):
    try:
        return os.stat(P(a_path)).st_dev == os.stat(P(b_dir)).st_dev
    except OSError:
        return False


def ensure_dir(d):
    os.makedirs(P(d), exist_ok=True)


def cmd_apply(a):
    rows = read_csv(a.plan)
    logp = a.log or os.path.splitext(a.plan)[0] + "_log.csv"
    undop = os.path.splitext(logp)[0] + "_undo.csv"
    done = set()
    if os.path.exists(logp):
        for r in read_csv(logp):
            if r.get("status") == "ok":
                done.add((r["action"], r["src"], r["dst"]))
    fields = ["time", "action", "src", "dst", "size", "sha1", "status", "detail"]
    newlog = not os.path.exists(logp)
    lf = open(logp, "a", newline="", encoding="utf-8")
    lw = csv.DictWriter(lf, fieldnames=fields)
    if newlog:
        lw.writeheader()
    newundo = not os.path.exists(undop)
    uf_ = open(undop, "a", newline="", encoding="utf-8")
    uw = csv.DictWriter(uf_, fieldnames=["action", "src", "dst"])
    if newundo:
        uw.writeheader()
    stats = Counter()
    t0 = last = time.time()
    first_fail_streak = 0
    srcdirs = set()
    for n, r in enumerate(rows, 1):
        act = (r.get("action") or "move").strip().lower()
        src, dst = r.get("src", ""), r.get("dst", "")
        key = (act, src, dst)
        if key in done:
            # trust the log only if the disk agrees (an undo may have reversed it since)
            if act == "delete" and not os.path.exists(P(src)):
                stats["skipped (already done)"] += 1
                continue
            if act in ("move", "quarantine") and os.path.exists(P(dst)) and not os.path.exists(P(src)):
                stats["skipped (already done)"] += 1
                continue
            if act == "copy" and os.path.exists(P(dst)):
                stats["skipped (already done)"] += 1
                continue
        rec = {"time": dt.datetime.now().isoformat(timespec="seconds"), "action": act, "src": src, "dst": dst}
        try:
            if not os.path.exists(P(src)):
                raise RuntimeError("source missing")
            size = os.path.getsize(P(src))
            rec["size"] = size
            if a.dry_run:
                if act in ("move", "copy", "quarantine") and os.path.exists(P(dst)):
                    raise RuntimeError("destination exists")
                rec["status"], rec["detail"] = "dry-run", ""
                stats["dry-run ok"] += 1
                continue
            if act == "delete":
                if not a.allow_delete:
                    raise RuntimeError("delete needs --allow-delete")
                kp = r.get("keeper")
                if not kp or not os.path.exists(P(kp)):
                    raise RuntimeError("keeper missing - refusing to delete")
                ks = r.get("keeper_sha1")
                if ks and sha1_file(kp) != ks:
                    raise RuntimeError("keeper SHA-1 changed - refusing to delete")
                os.remove(P(src))
                rec["status"] = "ok"
            elif act in ("move", "quarantine", "copy"):
                if os.path.exists(P(dst)):
                    raise RuntimeError("destination exists (never overwrite)")
                if act == "quarantine" and r.get("keeper"):
                    if not os.path.exists(P(r["keeper"])):
                        raise RuntimeError("keeper missing - refusing to quarantine")
                    if r.get("keeper_sha1") and a.verify == "full" and sha1_file(r["keeper"]) != r["keeper_sha1"]:
                        raise RuntimeError("keeper SHA-1 changed")
                ensure_dir(os.path.dirname(dst))
                if act != "copy" and same_device(src, os.path.dirname(dst)):
                    expected = r.get("sha1") or r.get("src_sha1")
                    if a.verify == "full" and expected and sha1_file(src) != expected:
                        raise RuntimeError("source SHA-1 differs from plan")
                    os.rename(P(src), P(dst))
                    if not os.path.exists(P(dst)) or os.path.getsize(P(dst)) != size:
                        raise RuntimeError("post-move check failed")
                    rec["detail"] = "renamed"
                    uw.writerow({"action": "move", "src": dst, "dst": src})
                    srcdirs.add(os.path.dirname(src))
                else:
                    h1 = sha1_file(src)
                    tmp = dst + ".pcat_tmp"
                    shutil.copy2(P(src), P(tmp))
                    h2 = sha1_file(tmp)
                    if h1 != h2:
                        os.remove(P(tmp))
                        raise RuntimeError("SHA-1 mismatch after copy")
                    os.replace(P(tmp), P(dst))
                    rec["sha1"] = h1
                    if act in ("move", "quarantine"):
                        os.remove(P(src))
                        uw.writerow({"action": "move", "src": dst, "dst": src})
                        srcdirs.add(os.path.dirname(src))
                        rec["detail"] = "copied+verified, source removed"
                    else:
                        rec["detail"] = "copied+verified"
                rec["status"] = "ok"
            else:
                raise RuntimeError(f"unknown action {act}")
            stats["ok"] += 1
            first_fail_streak = -10 ** 9
        except Exception as e:
            rec["status"], rec["detail"] = "FAILED", str(e)[:200]
            stats["FAILED"] += 1
            first_fail_streak += 1
        lw.writerow(rec)
        if time.time() - last > a.progress:
            lf.flush()
            uf_.flush()
            log(f"  {n:,}/{len(rows):,}  " + ", ".join(f"{k} {v:,}" for k, v in stats.items()))
            last = time.time()
        if first_fail_streak >= 50:
            log("  !! first 50 actions all failed - stopping (fail-fast). Check the log.")
            break
    lf.close()
    uf_.close()
    if a.prune_empty and not a.dry_run:
        removed = 0
        for d in sorted(srcdirs, key=len, reverse=True):
            cur = d
            while cur and os.path.isdir(P(cur)):
                try:
                    os.rmdir(P(cur))  # fails safely when not empty
                    removed += 1
                    cur = os.path.dirname(cur)
                except OSError:
                    break
        log(f"  removed {removed} emptied folders (rmdir only)")
    log(f"apply {'(DRY RUN) ' if a.dry_run else ''}done in {time.time() - t0:.0f}s: " +
        ", ".join(f"{k} {v:,}" for k, v in stats.items()))
    log(f"log: {logp}" + ("" if a.dry_run else f"\nundo: {undop}  (run `apply {os.path.basename(undop)}` to reverse moves)"))
    return 0 if not stats["FAILED"] else 1


# ----------------------------------------------------------------------------- verify (two trees)
def cmd_verify(a):
    ex = {e.lower() for e in (a.exclude or [])}

    def listing(root):
        out = {}
        for _, rel, st in walk_files(root, skip_dirs=SKIP_DIRS - {"_catalog"} if a.include_catalog else SKIP_DIRS):
            parts = rel.lower().split("/")
            if any(p in ex for p in parts[:-1]) or any(rel.lower().startswith(e.rstrip("/") + "/") for e in ex):
                continue
            out[rel] = st.st_size
        return out

    A, B = listing(a.a), listing(a.b)
    rows = []
    for p in A:
        if p not in B:
            rows.append({"path": p, "status": "missing_in_B", "size_a": A[p]})
        elif A[p] != B[p]:
            rows.append({"path": p, "status": "size_diff", "size_a": A[p], "size_b": B[p]})
    for p in B:
        if p not in A:
            rows.append({"path": p, "status": "extra_in_B", "size_b": B[p]})
    common = [p for p in A if p in B and A[p] == B[p]]
    hashed = 0
    if a.hash != "none":
        t0 = last = time.time()
        tot = sum(A[p] for p in common)
        nb = 0

        def job(p):
            fa = os.path.join(a.a, *p.split("/"))
            fb = os.path.join(a.b, *p.split("/"))
            return p, sha1_file(fa, a.hash == "partial"), sha1_file(fb, a.hash == "partial")

        with ThreadPoolExecutor(max_workers=a.workers) as exr:
            for p, ha, hb in exr.map(job, common):
                hashed += 1
                nb += A[p]
                if ha != hb:
                    rows.append({"path": p, "status": "hash_diff", "size_a": A[p]})
                if time.time() - last > a.progress:
                    el = time.time() - t0
                    log(f"  hashed {hashed:,}/{len(common):,}  {nb / 1e9:.1f}/{tot / 1e9:.1f} GB  "
                        f"ETA {(tot - nb) / (nb / el) / 60 if nb else 0:.0f} min")
                    last = time.time()
    out = a.out or "remaining_differences.csv"
    write_csv(out, rows, ["path", "status", "size_a", "size_b"])
    c = Counter(r["status"] for r in rows)
    log(f"verify: A {len(A):,} files | B {len(B):,} files | identical-by-size {len(common):,}"
        + (f" | hashed {hashed:,} ({a.hash})" if hashed else ""))
    log("  " + (", ".join(f"{k} {v:,}" for k, v in c.items()) if rows else "NO DIFFERENCES"))
    log(f"wrote {out}")
    return 0 if not rows else 2


# ----------------------------------------------------------------------------- contact sheets
def video_frame(path, at=1.0):
    Image, _ = pil()
    try:
        out = subprocess.run(["ffmpeg", "-v", "error", "-ss", str(at), "-i", path, "-frames:v", "1", "-f", "image2pipe",
                              "-vcodec", "png", "-"], capture_output=True, timeout=120)
        if not out.stdout:
            out = subprocess.run(["ffmpeg", "-v", "error", "-i", path, "-frames:v", "1", "-f", "image2pipe",
                                  "-vcodec", "png", "-"], capture_output=True, timeout=120)
        import io
        return Image.open(io.BytesIO(out.stdout)).convert("RGB")
    except Exception:
        return None


def cmd_sheet(a):
    Image, ImageOps = pil()
    from PIL import ImageDraw, ImageFont
    items = []
    if a.input.endswith(".jsonl"):
        for r in load_jsonl(a.input):
            items.append({"path": abspath_of(r), "label": r.get("date_best", "")[:10] if r.get("date_best") else "",
                          "kind": r["kind"]})
    else:
        for r in read_csv(a.input):
            p = r.get(a.column) or ""
            if p:
                items.append({"path": p, "label": r.get(a.label_column, "") if a.label_column else "",
                              "kind": "video" if os.path.splitext(p)[1].lower() in VIDEO_EXT else "photo"})
    if a.filter:
        rx = re.compile(a.filter)
        items = [i for i in items if rx.search(i["path"])]
    if a.kinds:
        ks = set(a.kinds.split(","))
        items = [i for i in items if i["kind"] in ks]
    os.makedirs(a.out, exist_ok=True)
    per = a.cols * a.rows
    tw = a.thumb
    th = tw + 44
    try:
        font = ImageFont.load_default(size=14)
        big = ImageFont.load_default(size=22)
    except TypeError:
        font = big = ImageFont.load_default()
    index = []
    nsheets = 0
    for s in range(0, len(items), per):
        page = items[s:s + per]
        sheet = Image.new("RGB", (a.cols * tw, a.rows * th), "white")
        dr = ImageDraw.Draw(sheet)
        for k, it in enumerate(page):
            num = a.start + s + k
            x, y = (k % a.cols) * tw, (k // a.cols) * th
            im = None
            try:
                if it["kind"] == "video":
                    im = video_frame(it["path"])
                    full = f"{im.width}x{im.height}" if im is not None else ""
                else:
                    im = Image.open(P(it["path"]))
                    full = f"{im.width}x{im.height}"
                    im.draft("RGB", (tw * 2, tw * 2))
                    im = ImageOps.exif_transpose(im).convert("RGB")
            except Exception:
                im = None
            if im is not None:
                res = full
                im.thumbnail((tw - 8, tw - 8))
                sheet.paste(im, (x + (tw - im.width) // 2, y + 4 + (tw - 8 - im.height) // 2))
            else:
                res = "unreadable"
                dr.rectangle([x + 4, y + 4, x + tw - 4, y + tw - 4], outline="red")
            dr.rectangle([x + 2, y + 2, x + 46, y + 28], fill="black")
            dr.text((x + 6, y + 4), str(num), fill="yellow", font=big)
            if it["kind"] == "video":
                dr.text((x + tw - 50, y + 6), "VIDEO", fill="red", font=font)
            name = os.path.basename(it["path"])
            dr.text((x + 4, y + tw), name[:34], fill="black", font=font)
            dr.text((x + 4, y + tw + 18), f"{it['label']}  {res}"[:36], fill="gray", font=font)
            index.append({"number": num, "sheet": nsheets + 1, "path": it["path"], "label": it["label"]})
        nsheets += 1
        sheet.save(os.path.join(a.out, f"sheet_{nsheets:03d}.jpg"), quality=85)
    write_csv(os.path.join(a.out, "sheet_index.csv"), index)
    log(f"wrote {nsheets} sheets ({len(items)} items) + sheet_index.csv to {a.out}")
    return 0


# ----------------------------------------------------------------------------- main
def main(argv=None):
    ap = argparse.ArgumentParser(prog="pcat", description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--version", action="version", version=VERSION)
    sp = ap.add_subparsers(dest="cmd", required=True)

    s = sp.add_parser("doctor"); s.set_defaults(fn=cmd_doctor)

    s = sp.add_parser("scan", help="inventory a tree into catalog JSONL (refreshes incrementally)")
    s.add_argument("root"); s.add_argument("--out", help="default <root>/_Catalog/catalog.jsonl")
    s.add_argument("--tz", help="IANA zone for UTC video times, e.g. America/Chicago (default: keep UTC)")
    s.add_argument("--workers", type=int, default=4); s.add_argument("--no-hash", action="store_true")
    s.add_argument("--fresh", action="store_true", help="ignore existing catalog")
    s.add_argument("--exclude", help="regex on relative path to skip")
    s.add_argument("--progress", type=float, default=60); s.add_argument("--checkpoint", action="store_true")
    s.add_argument("--no-vfp", action="store_true", help="skip video fingerprints (faster; weaker video matching)")
    s.set_defaults(fn=cmd_scan)

    s = sp.add_parser("from-inventory", help="pcat-win.ps1 inventory.csv -> minimal catalog JSONL")
    s.add_argument("inventory"); s.add_argument("--out", required=True)
    s.add_argument("--root-map", nargs="*", help='rewrite roots, e.g. "E:\\PHOTOS=/mnt/e_photos"')
    s.set_defaults(fn=cmd_from_inventory)

    s = sp.add_parser("date-seq", help="date undated files from dated neighbours in the shot/clip-number sequence")
    s.add_argument("catalog"); s.add_argument("--out")
    s.add_argument("--max-gap", type=int, default=300, help="max number distance between the two dated neighbours")
    s.add_argument("--max-days", type=int, default=14, help="max days between neighbours (<=1 day = high)")
    s.add_argument("--any-folder", action="store_true", help="allow neighbours from other top-level folders")
    s.add_argument("--trusted", default="exif,video,manual", help="date sources that count as dated neighbours")
    s.add_argument("--write", action="store_true", help="store the dates in the catalog (after review)")
    s.add_argument("--apply", default="high,moderate", help="confidences to store with --write")
    s.add_argument("--redo", action="store_true", help="also recompute files already dated by sequence")
    s.set_defaults(fn=cmd_date_seq)

    s = sp.add_parser("set-dates", help="store user-supplied dates from a CSV (path,date[,evidence])")
    s.add_argument("catalog"); s.add_argument("dates")
    s.set_defaults(fn=cmd_set_dates)

    s = sp.add_parser("catalog", help="catalog JSONL -> Photo_Catalog.xlsx")
    s.add_argument("catalog"); s.add_argument("--out"); s.add_argument("--no-backup", action="store_true")
    s.set_defaults(fn=cmd_catalog)

    s = sp.add_parser("reconcile", help="catalog vs disk")
    s.add_argument("catalog"); s.add_argument("--out"); s.add_argument("--show", type=int, default=10)
    s.set_defaults(fn=cmd_reconcile)

    s = sp.add_parser("dupes", help="tiered duplicate finder -> dupes_*.csv")
    s.add_argument("catalog"); s.add_argument("--outdir")
    s.add_argument("--pixel", type=float, default=2.0, help="max mean pixel diff (0-255) for Tier C")
    s.add_argument("--dhash", type=int, default=12, help="max dHash distance to try a pixel check")
    s.add_argument("--duration-tol", type=float, default=0.5)
    s.add_argument("--max-bucket", type=int, default=60)
    s.add_argument("--quarantine", default="_Duplicates_Quarantine")
    s.add_argument("--dump-re", default=DEFAULT_DUMP_RE)
    s.set_defaults(fn=cmd_dupes)

    s = sp.add_parser("compare", help="classify a new source catalog against the library catalog")
    s.add_argument("--library", required=True); s.add_argument("--source", required=True)
    s.add_argument("--also", nargs="*", help="other catalogs (e.g. Archive) whose SHA-1s count as identical")
    s.add_argument("--out", default="compare.csv"); s.add_argument("--pixel", type=float, default=3.0)
    s.add_argument("--dhash-loose", type=int, default=4)
    s.set_defaults(fn=cmd_compare)

    s = sp.add_parser("geonames", help="download GeoNames cities15000.txt (better place names)")
    s.add_argument("--dir", required=True, help="usually the _Catalog folder")
    s.set_defaults(fn=cmd_geonames)

    s = sp.add_parser("plan", help="dry-run reorganization plan")
    s.add_argument("catalog"); s.add_argument("--config", required=True)
    s.add_argument("--out", default="reorg_plan.csv"); s.add_argument("--max-path", type=int, default=240)
    s.add_argument("--action", choices=["move", "copy"], default="move",
                   help="copy = import from another drive, leaving the source untouched")
    s.add_argument("--include-from", help="compare.csv: plan only files whose class is in --classes")
    s.add_argument("--classes", default="new,upgrade")
    s.set_defaults(fn=cmd_plan)

    s = sp.add_parser("apply", help="execute a plan CSV (action,src,dst[,keeper,keeper_sha1,sha1])")
    s.add_argument("plan"); s.add_argument("--log"); s.add_argument("--dry-run", action="store_true")
    s.add_argument("--verify", choices=["full", "size"], default="full")
    s.add_argument("--allow-delete", action="store_true"); s.add_argument("--prune-empty", action="store_true")
    s.add_argument("--progress", type=float, default=30)
    s.set_defaults(fn=cmd_apply)

    s = sp.add_parser("verify", help="compare two trees (backup verification)")
    s.add_argument("a"); s.add_argument("b"); s.add_argument("--hash", choices=["none", "partial", "full"], default="full")
    s.add_argument("--exclude", nargs="*", help="relative folder paths or folder names to skip on both sides")
    s.add_argument("--include-catalog", action="store_true", help="also compare _Catalog folders")
    s.add_argument("--workers", type=int, default=2); s.add_argument("--out"); s.add_argument("--progress", type=float, default=60)
    s.set_defaults(fn=cmd_verify)

    s = sp.add_parser("sheet", help="numbered contact sheets")
    s.add_argument("input", help="CSV (with --column) or catalog JSONL"); s.add_argument("--out", required=True)
    s.add_argument("--column", default="path"); s.add_argument("--label-column")
    s.add_argument("--filter"); s.add_argument("--kinds", help="photo,video")
    s.add_argument("--cols", type=int, default=6); s.add_argument("--rows", type=int, default=5)
    s.add_argument("--thumb", type=int, default=260); s.add_argument("--start", type=int, default=1)
    s.set_defaults(fn=cmd_sheet)

    a = ap.parse_args(argv)
    return a.fn(a)


if __name__ == "__main__":
    sys.exit(main())
