# FY2024_UPDATE_AUDIT — NDB Open Data edition 11

## Source
- MHLW 11th NDB Open Data page: https://www.mhlw.go.jp/stf/seisakunitsuite/bunya/0000177221_00017.html (published 2026-06-16; FY2024 receipts).
- 処方薬 tables ship inside two zip archives ("こちら" links). Used: **公費レセプトを含まない** zip https://www.mhlw.go.jp/content/12400000/001742577.zip, folder `01_処方薬（内服／外用／注射）全/`.
- Files extracted to `data/ndb_raw/dai11/f00..f04.xlsx` (all raw editions 1–11 now persisted under `data/ndb_raw/`; `scripts/download_ndb.py` fetches all).

## Tables/files used (sex-age 薬効分類別数量)
f00 【内服】外来（院外）, f01 【内服】外来（院内）, f02 【内服】入院, f03 【外用】, f04 【注射】 — identical sheet-name scheme (`内服薬 外来 (院外)` etc.) and header layout (`薬効分類名称`/`医薬品名`/`単位`/`総計(処方数量)`) as editions 8–10. `header_cols` parses unmodified.

## Mapping to existing variables
Same brand-name substrings → same groups/products. All prior definitions valid.

## Discontinuities (documented, semantic equivalence checked)
1. **Listing expansion**: edition 11 publishes ALL products per class ("全"); editions ≤10 capped at top-30/100/500. Consequence: products previously absent for falling below the cap may re-appear in FY2024 — a **publication-scope change, not a dispensing change**. Concrete: ribavirin re-appears in FY2024 (28,508 units) after being below threshold since FY2018/FY2019; the manuscript must not describe this as a rebound. IFN_peg/IFN_conv/DAA anchor products were already above the cap in FY2023, so their FY2024 values are continuous.
2. Upper-bound machinery still works: absent products in FY2024 bound = 999 (cap_hit=False since no cap); listed "-" cells bound 999. FY2024 max_hidden DAA ≈20k vs ≈439k in FY2023 (narrower bounds because caps no longer hide products).
3. ボセビー (sofosbuvir/velpatasvir/voxilaprevir) absent from FY2024 tables (below publication threshold/withdrawn); also absent in earlier editions — no change.

## Resulting FY2024 principal estimates (rebuilt pipeline)
- DAA dispensed quantity: 1,164,534 units; estimated DAA courses 7,861.5
- Peginterferon: 4,240 syringes (~88.3 est. courses); IFN_conv 16,816; ribavirin 28,508 (see discontinuity 1)
- Combined estimated volume FY2024: 7,949.9 courses; peak still FY2015 (92,589); FY2024 = −91.4% vs peak; IFN-free share FY2024 = 98.9%
- PegIFN decay FY2014→2024: −99.3% (−33.8%/yr, HAC 95% −41.6..−19.0)
- Segmented combined volume (knot FY2015, descriptive): post-peak −23.4%/yr, n=10 post-knot obs
- Cumulative estimated courses FY2014–2024: ≈305,499

FY2014–2023 values unchanged (byte-identical rebuilt CSVs vs committed for FY2014–2023 columns).
