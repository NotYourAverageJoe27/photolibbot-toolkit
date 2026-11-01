#!/usr/bin/env python3
"""Self-test: builds a synthetic library in a temp folder and runs the whole pcat pipeline on it.
Run this once in any new environment before touching real photos:  python3 selftest.py
Exit 0 = every check passed. Nothing outside the temp folder is touched."""
import csv, json, os, shutil, subprocess, sys, tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
PCAT = [sys.executable, os.path.join(HERE, "pcat.py")]
fails = []


def run(*args, ok=(0,)):
    p = subprocess.run(PCAT + list(args), capture_output=True, text=True)
    if p.returncode not in ok:
        print(p.stdout[-2000:], p.stderr[-2000:])
        fails.append(f"exit {p.returncode}: {' '.join(args[:2])}")
    return p.stdout


def check(cond, msg):
    print(("  PASS " if cond else "  FAIL ") + msg)
    if not cond:
        fails.append(msg)


tmp = tempfile.mkdtemp(prefix="pcat_selftest_")
lib = os.path.join(tmp, "lib")
try:
    subprocess.run([sys.executable, os.path.join(HERE, "_testlib.py"), lib], check=True, capture_output=True)
    cat = os.path.join(lib, "_Catalog", "catalog.jsonl")
    run("doctor")
    run("scan", lib, "--tz", "America/Detroit")
    recs = [json.loads(l) for l in open(cat)]
    by = {r["path"]: r for r in recs}
    check(len(recs) == 84, f"scan found 84 files ({len(recs)})")
    check(by["iPhone Pics/IMG_7000.JPG"]["kind"] == "video", "video saved as .JPG detected")
    check(by["iPhone Pics/IMG_7001.MOV"]["kind"] == "problem", "all-zero video flagged")
    check(by["Cats/webpage.jpg"]["kind"] == "problem", "HTML saved as .jpg flagged")
    check(by["Kauai Sept 2012/IMG_100.JPG"]["date_source"] == "exif" and by["Kauai Sept 2012/IMG_100.JPG"]["gps_lat"], "EXIF date + GPS read")
    check(by["iPhone Pics/IMG_6000.MOV"]["date_best"] == "2014-08-17T15:03:22", "video UTC time converted to local")
    run("catalog", cat)
    check(os.path.exists(os.path.join(lib, "_Catalog", "Photo_Catalog.xlsx")), "xlsx written")
    run("dupes", cat)
    plan = list(csv.DictReader(open(os.path.join(lib, "_Catalog", "dupes_plan.csv"))))
    tiers = sorted(r["tier"] for r in plan)
    check(tiers == ["A", "B", "B", "C", "C", "C"], f"duplicate tiers A,B,B,C,C,C incl. rotated copy ({tiers})")
    check(all(v is None for v in by["GoPro/GOPR0001.MP4"]["vfp"]), "blank video frames give no fingerprint")
    grp = open(os.path.join(lib, "_Catalog", "dupes_groups.csv")).read()
    sim = open(os.path.join(lib, "_Catalog", "dupes_similar.csv")).read()
    check("GOPR000" not in grp and "GOPR000" not in sim, "two different black clips are not called duplicates")
    check("Partner Phone" not in grp and "Partner Phone" not in sim,
          "other phone's IMG_6000.MOV (same name+length, different clip) is not a duplicate")
    check("Exports/clip_export.mp4" in sim, "renamed re-encoded video found by look-alike (listed for review)")
    run("apply", os.path.join(lib, "_Catalog", "dupes_plan.csv"))
    # new-source comparison of videos
    vs = os.path.join(tmp, "vsrc")
    os.makedirs(os.path.join(vs, "A")); os.makedirs(os.path.join(vs, "B"))
    keepv = os.path.join(lib, "iPhone 5S Backup", "IMG_6000.MOV")
    ff = lambda *x: subprocess.run(["ffmpeg", "-v", "error", "-y", *x], check=True, capture_output=True)
    ff("-f", "lavfi", "-i", "smptebars=size=640x360:rate=30", "-t", "3", "-c:v", "libx264", "-pix_fmt", "yuv420p",
       "-metadata", "creation_time=2016-05-05T12:00:00Z", os.path.join(vs, "A", "IMG_6000.MOV"))
    ff("-i", keepv, "-vf", "scale=1920:1080", "-c:v", "libx264", "-crf", "18", "-c:a", "copy",
       "-metadata", "creation_time=2014-08-17T19:03:22Z", os.path.join(vs, "B", "IMG_6000.MOV"))
    ff("-i", keepv, "-vf", "scale=320:180", "-c:v", "libx264", "-an", "-map_metadata", "-1", os.path.join(vs, "B", "movie.mp4"))
    vcat = os.path.join(tmp, "vsrc.jsonl")
    run("scan", vs, "--out", vcat, "--tz", "America/Detroit")
    run("compare", "--library", cat, "--source", vcat, "--out", os.path.join(tmp, "vcmp.csv"))
    cls = {r["rel"]: r["class"] for r in csv.DictReader(open(os.path.join(tmp, "vcmp.csv")))}
    check(cls.get("A/IMG_6000.MOV") == "new", f"compare: same name+length but different clip -> new ({cls.get('A/IMG_6000.MOV')})")
    check(cls.get("B/IMG_6000.MOV") == "upgrade", f"compare: 1080p version of a library clip -> upgrade ({cls.get('B/IMG_6000.MOV')})")
    check(cls.get("B/movie.mp4") == "likely-duplicate", f"compare: renamed small copy -> likely-duplicate ({cls.get('B/movie.mp4')})")
    out = run("scan", lib)
    check("moved/renamed (reused) 6" in out, "rescan reuses moved files without re-reading")
    run("reconcile", cat)
    # dating the undated
    run("date-seq", cat, "--write")
    seq = {r["path"]: r for r in (json.loads(l) for l in open(cat))}
    check(seq["Camcorder/MOV00101.MPG"]["date_source"] == "sequence" and seq["Camcorder/MOV00101.MPG"]["date_confidence"] == "high",
          "clip-number dating: #101 between dated #100 and #103 (high)")
    check(seq["Camcorder/MOV00205.MPG"]["date_source"] != "sequence", "clip-number dating refuses neighbours 47 days apart")
    dcsv = os.path.join(tmp, "dates.csv")
    open(dcsv, "w").write("path,date,evidence\nCamcorder/MOV00205.MPG,2007-07,user\nCats/cat_undated.jpg,2015,\n")
    run("set-dates", cat, dcsv)
    gz = os.path.join(lib, "_Catalog", "cities15000.txt")
    try:
        run("geonames", "--dir", os.path.join(lib, "_Catalog"))
    except Exception:
        pass
    cfg = os.path.join(tmp, "cfg.json")
    json.dump({"homes": [{"from": "1990-01-01", "to": "2010-12-31", "lat": 37.6, "lon": -122.2, "radius_km": 80},
                         {"from": "2011-01-01", "to": "2999-12-31", "lat": 42.5, "lon": -83.2, "radius_km": 80}],
               "side_folders": {"^Docs/": "_Documents"}}, open(cfg, "w"))
    rp = os.path.join(lib, "_Catalog", "reorg_plan.csv")
    run("plan", cat, "--config", cfg, "--out", rp)
    folders = {r["folder"] for r in csv.DictReader(open(rp.replace(".csv", "_folders.csv")))}
    check("2012/2012-09-20 Kauai" in folders, "event folder named, no duplicate place suffix")
    check("2016/2016-12-25 Christmas" in folders, "busy day named by holiday")
    check("2013/2013-03 Cats" in folders, "topic folder split by month")
    newrel = {r["rel"]: r["new_rel"] for r in csv.DictReader(open(rp))}
    check(newrel.get("Camcorder/MOV00101.MPG", "").startswith("2006/2006-06 Camcorder/"), "sequence-dated clip placed by its date")
    check(newrel.get("Camcorder/MOV00205.MPG", "").startswith("2007/2007-07 Camcorder/"), "month-only user date -> month group")
    check(newrel.get("Cats/cat_undated.jpg", "").startswith("2015/2015 (month unknown) Cats/"), "year-only user date -> 'month unknown'")
    trip = [f for f in folders if f.startswith("2019/2019-03-31")]
    check(bool(trip), f"trip folder created ({trip})")
    if os.path.exists(gz):
        check(trip == ["2019/2019-03-31 Orlando, FL"], "trip labeled with the big city (GeoNames)")
    run("apply", rp, "--prune-empty")
    run("scan", lib)
    out = run("reconcile", cat)
    check("MATCH" in out, "catalog matches disk after reorganization")
    undo = rp.replace(".csv", "_log_undo.csv")
    run("apply", undo, "--prune-empty")
    check(os.path.isdir(os.path.join(lib, "Kauai Sept 2012")), "undo restores the original layout")
    run("apply", rp, "--prune-empty")
    check(os.path.isdir(os.path.join(lib, "2012", "2012-09-20 Kauai")), "re-apply after undo works")
    bk = os.path.join(tmp, "backup")
    shutil.copytree(lib, bk)
    out = run("verify", lib, bk, "--out", os.path.join(tmp, "d.csv"))
    check("NO DIFFERENCES" in out, "verify: identical backup")
    victim = next(os.path.join(dp, f) for dp, _, fs in os.walk(os.path.join(bk, "2012")) for f in fs)
    with open(victim, "r+b") as f:
        f.seek(300); f.write(b"\x00\x01")
    out = run("verify", lib, bk, "--out", os.path.join(tmp, "d.csv"), ok=(2,))
    check("hash_diff 1" in out, "verify: silent corruption caught by SHA-1")
    run("sheet", cat, "--out", os.path.join(tmp, "sheets"), "--kinds", "photo")
    check(os.path.exists(os.path.join(tmp, "sheets", "sheet_001.jpg")), "contact sheet written")
finally:
    shutil.rmtree(tmp, ignore_errors=True)
print("\nSELFTEST " + ("PASSED" if not fails else f"FAILED: {len(fails)} problem(s)"))
for f in fails:
    print("  - " + f)
sys.exit(1 if fails else 0)
