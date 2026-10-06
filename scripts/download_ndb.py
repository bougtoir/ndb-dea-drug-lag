#!/usr/bin/env python3
"""
Download the NDB Open Data 処方薬 (prescription-drug) workbooks for editions 1-11.

For editions 1-10 we fetch the MHLW "bunya" page and download every file whose label
contains both '性年齢' and '薬効分類別数量' (i.e. the 内服 外来院内/外来院外/入院,
外用, 注射 national sex-age tables). Dental (歯科) subset files are skipped by the
label filter used downstream. Files are saved to data/ndb_raw/dai{N}/f{i}.xlsx and a
manifest.json records edition, fiscal year, source URL and label.

Edition 11 (FY2024) is distributed differently: all 処方薬 tables ship inside two
zip archives (公費レセプトを含まない / 含む) linked only as "こちら", and the
薬効分類別数量 tables now list ALL products (the top-N listing cap no longer
applies). fetch_edition11() downloads the no-公費 zip and extracts the five sex-age
workbooks used by the pipeline as dai11/f00..f04.xlsx.

Source: MHLW NDB Open Data, https://www.mhlw.go.jp/ndb/opendatasite/
"""
import json
import os
import re
import time
import urllib.request

EDITIONS = {
    1: "0000139390", 2: "0000177221", 3: "0000177221_00002", 4: "0000177221_00003",
    5: "0000177221_00008", 6: "0000177221_00010", 7: "0000177221_00011",
    8: "0000177221_00012", 9: "0000177221_00014", 10: "0000177221_00016",
}
EDITIONS[11] = "0000177221_00017"  # FY2024; zip-packaged, see fetch_edition11
EDITION_FY = {1: 2014, 2: 2015, 3: 2016, 4: 2017, 5: 2018,
              6: 2019, 7: 2020, 8: 2021, 9: 2022, 10: 2023, 11: 2024}

# Edition 11: no-公費 zip URL and the five sex-age member files (order = f00..f04).
ED11_ZIP = "https://www.mhlw.go.jp/content/12400000/001742577.zip"
ED11_FILES = ["【内服】外来（院外）_性年齢別薬効分類別数量.xlsx",
              "【内服】外来（院内）_性年齢別薬効分類別数量.xlsx",
              "【内服】入院_性年齢別薬効分類別数量.xlsx",
              "【外用】性年齢別薬効分類別数量.xlsx",
              "【注射】性年齢別薬効分類別数量.xlsx"]

BUNYA = "https://www.mhlw.go.jp/stf/seisakunitsuite/bunya/"
OUTDIR = os.path.join(os.path.dirname(__file__), "..", "data", "ndb_raw")
HDR = {"User-Agent": "Mozilla/5.0 (research; NDB open data)"}


def get(url):
    req = urllib.request.Request(url, headers=HDR)
    return urllib.request.urlopen(req, timeout=90).read()


def fetch_edition11(manifest):
    import zipfile
    d = os.path.join(OUTDIR, "dai11")
    os.makedirs(d, exist_ok=True)
    zpath = os.path.join(d, "no_kohi.zip")
    if not os.path.exists(zpath):
        with open(zpath, "wb") as fh:
            fh.write(get(ED11_ZIP))
    z = zipfile.ZipFile(zpath)
    found = {}
    for n in z.namelist():
        dec = n.encode("cp437").decode("cp932")
        base = os.path.basename(dec)
        if base in ED11_FILES:
            found[base] = n
    for i, w in enumerate(ED11_FILES):
        data = z.read(found[w])
        fn = os.path.join(d, f"f{i:02d}.xlsx")
        with open(fn, "wb") as fh:
            fh.write(data)
        manifest.append({"edition": 11, "fy": 2024, "label": w,
                         "url": ED11_ZIP, "file": os.path.relpath(fn, OUTDIR)})
        print(f"dai11 FY2024: {w} -> {fn} ({len(data)} bytes)")


def main():
    os.makedirs(OUTDIR, exist_ok=True)
    manifest = []
    for ed, pid in EDITIONS.items():
        if ed == 11:
            fetch_edition11(manifest)
            continue
        html = get(BUNYA + pid + ".html").decode("utf-8", "replace")
        links = re.findall(r'href="([^"]+\.xlsx)"[^>]*>([^<]*)', html)
        picks = [(u, lbl.strip()) for u, lbl in links
                 if "薬効分類別数量" in lbl and "性年齢" in lbl and "歯科" not in lbl]
        d = os.path.join(OUTDIR, f"dai{ed}")
        os.makedirs(d, exist_ok=True)
        print(f"dai{ed} FY{EDITION_FY[ed]}: {len(picks)} files")
        for i, (u, lbl) in enumerate(picks):
            full = u if u.startswith("http") else "https://www.mhlw.go.jp" + u
            fn = os.path.join(d, f"f{i:02d}.xlsx")
            data = get(full)
            with open(fn, "wb") as fh:
                fh.write(data)
            manifest.append({"edition": ed, "fy": EDITION_FY[ed], "label": lbl,
                             "url": full, "file": os.path.relpath(fn, OUTDIR)})
            print(f"  {lbl} -> {fn} ({len(data)} bytes)")
            time.sleep(0.2)
    with open(os.path.join(OUTDIR, "manifest.json"), "w", encoding="utf-8") as fh:
        json.dump(manifest, fh, ensure_ascii=False, indent=1)
    print("TOTAL", len(manifest))


if __name__ == "__main__":
    main()
