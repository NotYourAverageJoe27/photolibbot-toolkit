"""Build a synthetic photo library that exercises every pcat rule."""
import os, random, shutil, subprocess, sys
from fractions import Fraction
from PIL import Image, ImageDraw

root = sys.argv[1]
shutil.rmtree(root, ignore_errors=True)
os.makedirs(root)
random.seed(7)


def rat(x):
    f = Fraction(x).limit_denominator(10000)
    from PIL.TiffImagePlugin import IFDRational
    return IFDRational(f.numerator, f.denominator)


def dms(v):
    v = abs(v); d = int(v); m = int((v - d) * 60); s = (v - d - m / 60) * 3600
    return (rat(d), rat(m), rat(round(s, 2)))


def photo(rel, when=None, gps=None, cam="iPhone 5s", seed=None, size=(800, 600), q=92, rotate=0, orient=None):
    path = os.path.join(root, rel)
    os.makedirs(os.path.dirname(path), exist_ok=True)
    rnd = random.Random(seed if seed is not None else rel)
    im = Image.new("RGB", size, tuple(rnd.randrange(256) for _ in range(3)))
    d = ImageDraw.Draw(im)
    for _ in range(25):
        x, y = rnd.randrange(size[0]), rnd.randrange(size[1])
        d.ellipse([x, y, x + rnd.randrange(40, 200), y + rnd.randrange(40, 200)],
                  fill=tuple(rnd.randrange(256) for _ in range(3)))
    if rotate:
        im = im.rotate(rotate, expand=True)
    ex = Image.Exif()
    if cam:
        ex[271] = "Apple"; ex[272] = cam
    if orient:
        ex[274] = orient
    if when:
        ex.get_ifd(0x8769)[36867] = when
    if gps:
        g = ex.get_ifd(0x8825)
        g[1] = "N" if gps[0] >= 0 else "S"; g[2] = dms(gps[0])
        g[3] = "E" if gps[1] >= 0 else "W"; g[4] = dms(gps[1])
    im.save(path, quality=q, exif=ex)
    return path


DET = (42.49, -83.13); KAUAI = (22.08, -159.32); ORL = (28.47, -81.47); SF = (37.77, -122.42)
# Event folder: Kauai 2012 (away), 8 photos
for i in range(8):
    photo(f"Kauai Sept 2012/IMG_{100 + i}.JPG", f"2012:09:{20 + i // 3:02d} 1{i}:00:00", KAUAI, cam="iPhone 4S")
# Event at home (Bay Area home through 2010)
for i in range(6):
    photo(f"Wedding Dec 10 2010/DSC0{i}.JPG", f"2010:12:10 15:{i:02d}:00", SF, cam="NIKON COOLPIX S8100")
# Topic folder: Cats across years
for i, d in enumerate(["2013:03:01", "2013:03:15", "2015:06:02", "2015:11:20", "2016:01:05"]):
    photo(f"Cats/cat{i}.jpg", f"{d} 10:00:00", None, cam="iPhone 5s")
photo("Cats/cat_undated.jpg", None, None, cam=None)
# Phone dump
for i in range(3):  # same day as Kauai -> joins event
    photo(f"iPhone Pics/IMG_2{i:03d}.JPG", f"2012:09:21 09:0{i}:00", KAUAI, cam="iPhone 4S")
for i in range(6):  # Orlando trip, 3 days
    photo(f"iPhone Pics/IMG_3{i:03d}.JPG", f"2019:03:{31 if i < 2 else '04:0' + str(i // 2)}"
          if False else (f"2019:03:31 1{i}:00:00" if i < 2 else f"2019:04:0{i // 2} 1{i}:00:00"), ORL, cam="iPhone 7 Plus")
for i in range(22):  # busy day at home: Christmas 2016
    photo(f"iPhone Pics/IMG_4{i:03d}.JPG", f"2016:12:25 0{i // 3}:{i:02d}:00", DET if i % 2 else None, cam="iPhone 7 Plus")
for i in range(4):  # ordinary home days -> month folder
    photo(f"iPhone Pics/IMG_5{i:03d}.JPG", f"2014:08:{10 + i} 12:00:00", DET, cam="iPhone 5s")
for i in range(2):  # undated in dump
    photo(f"iPhone Pics/download{i}.jpg", None, None, cam=None)
# Duplicates
k = os.path.join(root, "Kauai Sept 2012/IMG_100.JPG")
shutil.copy2(k, os.path.join(root, "iPhone Pics/IMG_100.JPG"))                     # Tier A exact
photo("iPhone 5S Backup/IMG_101.JPG", "2012:09:20 11:00:00", None, cam="iPhone 4S",
      seed="Kauai Sept 2012/IMG_101.JPG", q=70)                                       # Tier C re-save, lost GPS -> held? keeper has GPS
photo("iPhone 5S Backup/IMG_102_small.JPG", "2012:09:20 12:00:00", KAUAI, cam="iPhone 4S",
      seed="Kauai Sept 2012/IMG_102.JPG", size=(800, 600), q=80)                      # same second, re-save
photo("iPhone 5S Backup/IMG_103.JPG", "2012:09:21 13:00:00", KAUAI, cam="iPhone 4S",
      seed="Kauai Sept 2012/IMG_103.JPG", rotate=90)                                  # sideways copy
photo("iPhone 5S Backup/IMG_104.JPG", "2012:09:21 14:00:00", KAUAI, cam="iPhone 4S",
      seed="different picture")                                                       # same second, different photo -> D/none
# Videos
vd = os.path.join(root, "iPhone Pics"); os.makedirs(vd, exist_ok=True)
def vid(path, ct, extra=(), size="640x360", loc=None):
    meta = ["-metadata", f"creation_time={ct}"]
    if loc: meta += ["-metadata", f"location={loc}"]
    subprocess.run(["ffmpeg", "-v", "error", "-y", "-f", "lavfi", "-i", f"testsrc=size={size}:rate=30:duration=3",
                    "-f", "lavfi", "-i", "sine=frequency=440:duration=3", "-c:v", "libx264", "-pix_fmt", "yuv420p",
                    "-c:a", "aac", "-shortest", *meta, *extra, path], check=True)
vid(os.path.join(vd, "IMG_6000.MOV"), "2014-08-17T19:03:22Z", loc="+42.4900-083.1300/")
os.makedirs(os.path.join(root, "iPhone 5S Backup"), exist_ok=True)
subprocess.run(["ffmpeg", "-v", "error", "-y", "-i", os.path.join(vd, "IMG_6000.MOV"), "-map", "0", "-c", "copy",
                "-metadata", "title=Windows tag", "-metadata", "creation_time=2014-08-17T19:03:22Z", os.path.join(root, "iPhone 5S Backup/IMG_6000.MOV")], check=True)  # Tier B
vid(os.path.join(root, "iPhone 5S Backup/IMG_6000_small.MOV"), "2014-08-17T19:03:22Z", size="320x180")  # re-encode -> D-video
# Video look-alikes / traps (v1.1)
def vsrc(path, src, ct=None, size="640x360", vcodec=("-c:v", "libx264")):
    meta = ["-metadata", f"creation_time={ct}"] if ct else []
    os.makedirs(os.path.dirname(path), exist_ok=True)
    spec = f"{src}{':' if '=' in src else '='}size={size}:rate=30"
    subprocess.run(["ffmpeg", "-v", "error", "-y", "-f", "lavfi", "-i", spec, "-t", "3",
                    *vcodec, "-pix_fmt", "yuv420p", *meta, path], check=True)
# a different person's phone: same file name as IMG_6000.MOV, same length, different picture
vsrc(os.path.join(root, "Partner Phone/IMG_6000.MOV"), "mandelbrot", "2015-02-01T18:00:00Z")
# renamed, re-encoded (smaller) export of IMG_6000.MOV, no date
subprocess.run(["ffmpeg", "-v", "error", "-y", "-i", os.path.join(vd, "IMG_6000.MOV"), "-vf", "scale=480:270",
                "-c:v", "libx264", "-crf", "30", "-an", "-map_metadata", "-1", os.path.join(root, "Exports/clip_export.mp4")]
               if os.makedirs(os.path.join(root, "Exports"), exist_ok=True) is None else [], check=True)
# two different all-black clips (finger over lens): must never match each other
vsrc(os.path.join(root, "GoPro/GOPR0001.MP4"), "color=c=black", "2014-07-19T15:00:00Z")
vsrc(os.path.join(root, "GoPro/GOPR0002.MP4"), "color=c=black", "2014-07-19T16:00:00Z")
# camcorder clips: numbered, only some carry a date (clip-number dating)
for n, ct in [(100, "2006-06-25T14:00:00Z"), (101, None), (102, None), (103, "2006-06-25T16:00:00Z"),
              (200, "2007-07-04T12:00:00Z"), (205, None), (290, "2007-08-20T12:00:00Z")]:
    vsrc(os.path.join(root, f"Camcorder/MOV00{n}.MPG"), f"life=seed={n}:ratio=0.4", ct, size="320x240",
         vcodec=("-c:v", "mpeg2video", "-f", "mp4"))
# Junk + problems
open(os.path.join(root, "iPhone Pics/desktop.ini"), "w").write("[x]")
open(os.path.join(root, "iPhone Pics/IMG_2000.AAE"), "w").write("<plist/>")
open(os.path.join(root, "Cats/empty.jpg"), "w").close()
shutil.copy2(os.path.join(vd, "IMG_6000.MOV"), os.path.join(root, "iPhone Pics/IMG_7000.JPG"))  # video named .JPG
open(os.path.join(root, "iPhone Pics/IMG_7001.MOV"), "wb").write(b"\0" * 20000)                  # all zeros
open(os.path.join(root, "Cats/webpage.jpg"), "w").write("<!DOCTYPE html><html></html>")
open(os.path.join(root, "Docs/notes.pdf") if os.makedirs(os.path.join(root, "Docs"), exist_ok=True) is None else "", "wb").write(b"%PDF-1.4 test")
# Long path
lp = os.path.join(root, "Deep " + "x" * 60, "y" * 70, "z" * 70, "photo_with_a_long_name.jpg")
photo(os.path.relpath(lp, root), "2015:05:05 05:05:05", DET)
print("built", sum(len(f) for _, _, f in os.walk(root)), "files")
