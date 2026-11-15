# HW corporate PowerPoint template — design-token & structure report

**Input file:** `D:\Projects\ai\BrandDeck\PPT模板-HW.pptx` (2 387 726 bytes on disk; 66 package parts)
**Scope:** analysis only — the `.pptx` was never written to. All evidence below comes from reading the
package with `zipfile` + `xml.etree` (and cross-checked with `python-pptx` 1.0.2).
**Machine-readable companion:** `research/expected-tokens.json`
**Verifier:** `research/dump-template.mjs` (its terminal output is `research/dump-template.output.txt`;
verbatim XML fragments are in `research/raw-theme-fragments.txt`).

Document metadata for context: `docProps/core.xml` creator `AutoBVT`, last modified `Hanke (A)`, modified
2021-01-28; `docProps/app.xml` `<Company>Huawei Technologies Co.,Ltd.</Company>`,
`<Template>MWC 19-PPT模板-白底 16比9</Template>` ("white background, 16:9"), Application
"Microsoft Office PowerPoint", `<PresentationFormat>自定义</PresentationFormat>`.

---

## 1. Package inventory

| Kind | Count | Parts |
|---|---|---|
| slide masters | **4** | `ppt/slideMasters/slideMaster1..4.xml` |
| slide layouts | **7** | `ppt/slideLayouts/slideLayout1..7.xml` |
| slides | 5 | `ppt/slides/slide1..5.xml` |
| themes | **6** | `ppt/theme/theme1..6.xml` (**no** `themeOverride*.xml` anywhere) |
| media | 6 | `ppt/media/image1.png`, `image2.png`, `image3.jpeg`, `image4.jpg`, `image5.jpeg`, `image6.png` |
| charts | 1 | `ppt/charts/chart1.xml` (+ `colors1.xml`, `style1.xml`) |
| embedded pkgs | 1 | `ppt/embeddings/Microsoft_Excel____1.xlsx` (chart data) |
| notes / handout masters | 1 + 1 | `notesMaster1.xml`, `handoutMaster1.xml` (each carries its own theme) |
| notes slides | 1 | `notesSlides/notesSlide1.xml` |

Master order as declared in `ppt/presentation.xml` → `ppt/_rels/presentation.xml.rels`:

| # | r:id | master | theme |
|---|---|---|---|
| 1 | rId1 | `slideMasters/slideMaster1.xml` | `theme/theme1.xml` |
| 2 | rId2 | `slideMasters/slideMaster2.xml` | `theme/theme2.xml` |
| 3 | rId3 | `slideMasters/slideMaster3.xml` | `theme/theme3.xml` |
| 4 | rId4 | `slideMasters/slideMaster4.xml` | `theme/theme4.xml` |

`theme5.xml` → notes master, `theme6.xml` → handout master.
`theme1.xml` is additionally the presentation-level theme (`presentation.xml.rels` rId14), i.e. the
default theme a parser should treat as "the deck's theme".

---

## 2. Slide size

```xml
<p:sldSz cx="12196763" cy="6858000"/>
<p:notesSz cx="6858000" cy="9144000"/>
```

| | EMU | inches | points | px @96 dpi |
|---|---|---|---|---|
| width | `12196763` | **13.3385** | 960.38 | 1280.50 |
| height | `6858000` | **7.5000** | 540.00 | 720.00 |

* Aspect = 12196763 / 6858000 = **1.778472**; 16:9 is 1.777778. The deck is *nominally* 16:9 (the
  document template literally says `16比9`), but **width is not the canonical 12192000 EMU**:
  it is **+4 763 EMU (+0.00521 in)** wider.
* Consequence: **never hardcode `12192000`** or `13.3333 in`. Use `cx=12196763`, `widthIn=13.3385`.
  Note `layout1..4` background pictures are drawn with `cx="12196763"` — exactly the slide width — which
  confirms the odd value is real and not a rounding artifact of the viewer.
* Width in px is fractional (1280.50), so rasterisation at 96 dpi lands on a half pixel.

---

## 3. Theme colours — `<a:clrScheme>` of `ppt/theme/theme1.xml`

Theme header: `<a:theme ... name="1_Title Slide">`, `<a:clrScheme name="2210">`,
`<a:fmtScheme name="Office">`, `<a:extraClrSchemeLst/>` (empty), `<a:objectDefaults>` present,
`<a:extLst>` present. Note the three names differ (`1_Title Slide` ≠ `2210` ≠ `Arial Black-Arial`);
"theme name" is ambiguous unless you say which attribute you mean.

All twelve slots, raw form and effective value (verbatim from `research/raw-theme-fragments.txt`):

| slot | raw XML form | effective |
|---|---|---|
| `dk1` | `<a:srgbClr val="1D1D1A"/>` | **#1D1D1A** |
| `lt1` | `<a:srgbClr val="1D1D1A"/>` | **#1D1D1A** |
| `dk2` | `<a:srgbClr val="FFFFFF"/>` | **#FFFFFF** |
| `lt2` | `<a:srgbClr val="FFFFFF"/>` | **#FFFFFF** |
| `accent1` | `<a:srgbClr val="C7000A"/>` | **#C7000A** |
| `accent2` | `<a:srgbClr val="E9002F"/>` | **#E9002F** |
| `accent3` | `<a:srgbClr val="F4A100"/>` | **#F4A100** |
| `accent4` | `<a:srgbClr val="FFFF00"/>` | **#FFFF00** |
| `accent5` | `<a:srgbClr val="232323"/>` | **#232323** |
| `accent6` | `<a:srgbClr val="666666"/>` | **#666666** |
| `hlink` | `<a:srgbClr val="919191"/>` | **#919191** |
| `folHlink` | `<a:srgbClr val="C4C4C4"/>` | **#C4C4C4** |

**The defining oddity: the pairs are duplicated.** `dk1 == lt1 == #1D1D1A` and `dk2 == lt2 == #FFFFFF`.
With `slideMaster1`'s `<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" …/>` this means

* `bg1` (usually white) = **#1D1D1A** (near-black)
* `tx1` (usually black) = **#1D1D1A**
* `bg2` = `tx2` = **#FFFFFF**

So "background 1" and "text 1" are the *same* colour, and "background 2"/"text 2" are both white.
A parser that assumes `lt1=#FFFFFF` / `dk1=#000000` will invert this deck. This template is a
"white background" deck: the white comes from an explicit `<p:bg><p:bgPr><a:solidFill><a:srgbClr val="FFFFFF"/>`,
not from `lt1`.

### 3.1 The other five themes (they are not the same scheme)

| part | theme@name | clrScheme@name | fontScheme@name | deviation from theme1 |
|---|---|---|---|---|
| theme1 | `1_Title Slide` | `2210` | `Arial Black-Arial` | — (reference) |
| theme2 | `Chart page` | `updated` | `Office` | dk/lt identical, but **accent2..accent6 = #C8102E, #EA594F, #F8B53C, #231815, #595757**, hlink `#898989`, folHlink `#B5B5B5` |
| theme3 | `4_Chart page` | `Custom 29` | `Office` | accents identical to theme1, but **lt1=#666666, lt2=#EBEBEB** |
| theme4 | `End page` | `2210` | `Arial` | clrScheme identical to theme1; fontScheme major **and** minor = `Arial` |
| theme5 | `Office Theme` | `Office` | `Office` | stock Office: **sysClr** — see below |
| theme6 | `Office Theme` | `Office` | `Office` | identical to theme5 |

theme5/theme6 (notes & handout masters) are the only parts containing `<a:sysClr>`:

```xml
<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
<a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>
<a:accent1><a:srgbClr val="4472C4"/></a:accent1> … <a:accent6><a:srgbClr val="70AD47"/></a:accent6>
```

`lastClr` is **optional** in the schema, so a robust parser must (a) read `lastClr` when present and
(b) fall back (`windowText`→`#000000`, `window`→`#FFFFFF`) when it is missing. In *this* file every
`sysClr` happens to carry `lastClr`, but a validator must not crash on it being absent.

---

## 4. Effective colour tokens (what `expected-tokens.json` asserts)

| token | value | evidence |
|---|---|---|
| `background` | **#FFFFFF** | `<p:bg>` solidFill on all four masters; layout7 also white |
| `surface` | **#EBEBEB** | `slideLayout5` `<p:bg>` solidFill — the **only** non-white background fill in the whole package (matches theme3 `lt2=#EBEBEB`) |
| `textPrimary` | **#1D1D1A** | `schemeClr tx1` → `dk1`; also `slideMaster1` title `schemeClr tx1` |
| `textSecondary` | **#666666** | `accent6`; by far the most-used non-`tx1` scheme colour in the package (164 `schemeClr accent6` hits) and the dominant grey in chart data points |
| `primary` | **#C7000A** | `accent1` (brand red) |
| `accent` | **#E9002F** | `accent2` |
| `chartPalette` | **#C7000A, #E9002F, #F4A100, #FFFF00, #232323, #666666** | `accent1..6`; `ppt/charts/colors1.xml` `<cs:colorStyle id="10" meth="cycle">` maps exactly `accent1..accent6` (plus lumMod/lumOff variations) |

Two caveats worth knowing before you use `primary`:

* `slideLayout1..4` paint their corner mark with the literal **`C7000B`**, one hex digit off the theme's
  `#C7000A`. `slideMaster3`'s swatch board additionally uses `C8102E`, `78000F`, `A72126`, `9F0001`.
* `slideMaster1`'s body style uses the literal **`1D1D1B`**, one hex digit off `dk1 = 1D1D1A`
  (`1D1D1B` appears 40× across the masters — once in master1's `bodyStyle`, the rest on master3's
  swatch board). Do not "normalise" either value; both are intentional in the artwork.

Chart detail (`chart1.xml`): one `barChart` series with `schemeClr accent1`, overridden per point by
`dPt` fills `C5C5C5`, `929292`, `666666`, `accent2`, `C00000`, `FBA000`.

---

## 5. Fonts

### 5.1 `<a:fontScheme>` of theme1 (`name="Arial Black-Arial"`)

| | `<a:latin>` | `<a:ea>` | `<a:cs>` |
|---|---|---|---|
| `majorFont` | `Arial Black` (panose `020B0A04020102020204`) | `""` (**empty**) | `""` (**empty**) |
| `minorFont` | `Arial` (panose `020B0604020202020204`) | `""` (**empty**) | `""` (**empty**) |

Per-script fallbacks that actually carry the CJK faces:

| script | majorFont | minorFont |
|---|---|---|
| `Hans` (zh-CN, the one that matters here) | **微软雅黑** | **黑体** |
| `Hant` | 微軟正黑體 | 微軟正黑體 |
| `Jpan` | ＭＳ ゴシック | ＭＳ Ｐゴシック |
| `Hang` | 굴림 | 굴림 |
| `Thai` | FreesiaUPC | Cordia New |
| …30 more | | |

**Trap:** `<a:ea typeface=""/>` and `<a:cs typeface=""/>` are *present but empty strings*, not absent
attributes. Empty means "use the per-script `<a:font script=…>` face". Reporting `ea: ""` is
technically faithful and practically useless; a good parser reports `ea → (empty → Hans: 微软雅黑)`.

### 5.2 Master `txStyles` — these beat the theme

`<p:txStyles>` on each master overrides the theme faces. Measured at `lvl1pPr/<a:defRPr>`:

| master | titleStyle sz / colour / latin / ea | bodyStyle sz / colour / latin / ea / cs |
|---|---|---|
| master1 | 32.00 pt, `schemeClr tx1`, **Microsoft YaHei** / **Microsoft YaHei** | 10.00 pt, **`srgbClr 1D1D1B`**, Arial / Microsoft YaHei / Arial |
| master2 | 57.16 pt, `schemeClr tx1`, `+mj-lt` (=Arial Black) / `+mj-ea` (empty) | 36.36 pt, `schemeClr tx1`, `+mn-lt` (=Arial) / `+mn-ea` (empty) |
| master3 | 57.16 pt, `schemeClr tx1`, `+mj-lt` / `+mj-ea` | 36.36 pt, `schemeClr tx1`, `+mn-lt` / `+mn-ea` |
| master4 | 50.00 pt, `schemeClr tx1`, Arial / Microsoft YaHei / Arial | 18.19 pt, `srgbClr FFFFFF`, **Microsoft YaHei / Microsoft YaHei** / `+mn-cs` |

So the theme's `Arial Black` title face is overridden to **Microsoft YaHei** on the primary master
(master1) and on master4's ea. master1's `titleStyle` also sets `defTabSz=914400`, `lnSpc spcPts=3440`
(34.4 pt line spacing), `buNone`.

`ppt/presentation.xml` `<p:defaultTextStyle>` `lvl1pPr`: `sz="1800"` (**18 pt**),
`<a:solidFill><a:schemeClr val="tx1"/></a:solidFill>`,
`+mn-lt` / `+mn-ea` / `+mn-cs`, `marL="0"`, `defTabSz="914478"`.

### 5.3 Font tokens

| token | value | meaning |
|---|---|---|
| `fonts.title` | **Microsoft YaHei** | effective title face (master1 `titleStyle` latin *and* ea) |
| `fonts.body` | **Arial** | effective body Latin face (master1 `bodyStyle` latin; = theme `minorFont/latin`) |
| `fonts.latin` | **Arial** | theme1 `minorFont/latin` |
| `fonts.cs` | **Arial** | master1 `bodyStyle` `cs` (the theme-level `cs` is the empty string) |

Decision record, because "the" body font is genuinely ambiguous in this file:
the CJK body face is **Microsoft YaHei** (master1 `bodyStyle/ea`, and the Hans major fallback 微软雅黑);
the Latin body face is **Arial**; content pages (layout6 → master3 → theme3) inherit `Calibri`/`黑体`
instead because master3 uses `+mn-lt`. `expected-tokens.json` records master1 + theme1 (the primary
master and the presentation-level theme), which is the strongest single answer.

Fonts actually used, per `docProps/app.xml` `<HeadingPairs>`: 黑体, Microsoft YaHei, Arial, Calibri,
Calibri Light (5 used faces).

---

## 6. Master / layout inventory

### 6.1 Masters

| # | part | `<p:cSld>` name | theme | layouts | background | shapes | `<p:hf>` |
|---|---|---|---|---|---|---|---|
| 1 | slideMaster1 | *(none)* | theme1 | 4 (L1-L4) | solidFill **#FFFFFF** | 2 | `hdr=0 ftr=0 dt=0` |
| 2 | slideMaster2 | *(none)* | theme2 | 1 (L5) | solidFill **#FFFFFF** | **0 (empty spTree)** | `hdr=0 ftr=0 dt=0` |
| 3 | slideMaster3 | *(none)* | theme3 | 1 (L6) | solidFill **#FFFFFF** | 4 (+ nested groups) | `hdr=0 ftr=0 dt=0` |
| 4 | slideMaster4 | *(none)* | theme4 | 1 (L7) | solidFill **#FFFFFF** | 5 | `hdr=0 ftr=0 dt=0` |

No master declares a `name` attribute on `<p:cSld>`, and no master has a `title` placeholder except
master4 (so "master title" cannot be derived for 3 of 4 masters). All use the identical `clrMap`
`bg1=lt1 tx1=dk1 bg2=lt2 tx2=dk2 …`; no `overrideClrMapping` anywhere.

### 6.2 Layouts

None of the seven `<p:sldLayout>` elements carries a `@type` attribute (verified: `type=None` on all 7).
Identify them by `<p:cSld name="…">`. All declare `userDrawn="1"`.

| # | part | `cSld/@name` | master | theme | background | media | `clrMapOvr` |
|---|---|---|---|---|---|---|---|
| 1 | slideLayout1 | **探索** ("explore") | M1 | theme1 | *(none → master #FFFFFF)* | `image2.png` | `masterClrMapping` |
| 2 | slideLayout2 | **智能** ("intelligent") | M1 | theme1 | *(none → master #FFFFFF)* | `image3.jpeg` | `masterClrMapping` |
| 3 | slideLayout3 | **攀登** ("climb") | M1 | theme1 | *(none → master #FFFFFF)* | `image4.jpg` | `masterClrMapping` |
| 4 | slideLayout4 | **灯塔** ("lighthouse") | M1 | theme1 | *(none → master #FFFFFF)* | `image5.jpeg` | `masterClrMapping` |
| 5 | slideLayout5 | **1_Contents page** | M2 | theme2 | solidFill **#EBEBEB** | — | `masterClrMapping` |
| 6 | slideLayout6 | **1_Chinese text page** | M3 | theme3 | *(none → master #FFFFFF)* | — | `masterClrMapping` |
| 7 | slideLayout7 | **End page** | M4 | theme4 | solidFill **#FFFFFF** | — | `masterClrMapping` |

`layoutNames` for the JSON, in package order:
`["探索", "智能", "攀登", "灯塔", "1_Contents page", "1_Chinese text page", "End page"]`.

### 6.3 Layout → master → theme, and the 5 slides

| slide | layout | master | theme | section (`p14:sectionLst`) |
|---|---|---|---|---|
| slide1 | layout1 (探索) | M1 | theme1 | 封面页_图片版 (cover, image version) |
| slide2 | layout5 (1_Contents page) | M2 | theme2 | 目录页 (TOC) |
| slide3 | layout6 (1_Chinese text page) | M3 | theme3 | 章节页 (section divider) |
| slide4 | layout6 (1_Chinese text page) | M3 | theme3 | 章节页 (table + chart demo) |
| slide5 | layout7 (End page) | M4 | theme4 | 结束页 (end page) |

"One theme for the whole deck" is therefore false in practice: the actual content pages (slide3/4) run
on theme3, whose `lt1=#666666` and `lt2=#EBEBEB`.

---

## 7. Media inventory

Pixel dimensions were read straight out of the file headers (PNG IHDR / JPEG SOFn), not guessed.

| part | bytes | px | format | referenced by (rel id) | classification |
|---|---|---|---|---|---|
| `ppt/media/image1.png` | **22 308** | **529 × 116** | png | `slideMasters/_rels/slideMaster1.xml.rels` (rId6), `slideMasters/_rels/slideMaster4.xml.rels` (rId3) | **logo** (canonical) |
| `ppt/media/image2.png` | 1 284 697 | 1949 × 917 | png | `slideLayouts/_rels/slideLayout1.xml.rels` (rId2) | background/photo |
| `ppt/media/image3.jpeg` | 76 774 | 1949 × 917 | jpeg | `slideLayouts/_rels/slideLayout2.xml.rels` (rId2) | background/photo |
| `ppt/media/image4.jpg` | 639 346 | 1949 × 1122 | jpeg | `slideLayouts/_rels/slideLayout3.xml.rels` (rId2) | background/photo |
| `ppt/media/image5.jpeg` | 383 121 | 2982 × 1657 | jpeg | `slideLayouts/_rels/slideLayout4.xml.rels` (rId2) | background/photo |
| `ppt/media/image6.png` | 6 492 | 242 × 53 | png | `slideMasters/_rels/slideMaster3.xml.rels` (rId3) | **logo-variant** (same artwork, 242 px wide) |

No media part is referenced from any `slide*.xml.rels` — all five slides have zero image relationships
(slide4's only extra rel is the chart; slide5's is a notesSlide). So *every* image in the deck arrives
via a master or a layout.

### 7.1 Which one is the corporate logo

**`ppt/media/image1.png` — the HUAWEI logo, 529 × 116 px, 22 308 bytes, PNG.**

Evidence:

* It is a brand mark: the eight-petal red Huawei flower plus the black "HUAWEI" wordmark.
  (Rendered to confirm; `image6.png` is the *same* artwork at 242 × 53 px.)
* Small and shallow: 529 × 116 px, aspect 4.56 : 1 — a horizontal lockup, characteristic of a logo strip,
  versus 1949 × 917 … 2982 × 1657 px full-bleed photography for the other four.
* Referenced **only** from slide masters, never from a layout or slide.
* Placed on the master as a `userDrawn` picture with no `<p:ph>`:
  `slideMaster1` frame `off(9206684, 5976168) ext(2258389, 482533)` →
  **x 10.0686 in, y 6.5356 in, w 2.4698 in, h 0.5277 in** — i.e. **bottom-right** (not top-left/top-right
  as is often assumed; the slide is 13.3385 in wide, so the logo sits from 75.5 % to 94.0 % of the width
  and from 87.1 % to 94.2 % of the height).
* `slideMaster4` places the same part at `off(7975497, 5251150) ext(1869596, 399462)` →
  **x 8.7221 in, y 5.7427 in, w 2.0446 in, h 0.4369 in** (also bottom-right, slightly higher/smaller).

Logo judgement caveat for downstream code: **"referenced by exactly one master" is not a usable
heuristic here**, because `image1.png` is referenced by *two* masters (M1 rId6, M4 rId3) and the second
logo (`image6.png`) has a *different* resolution. Ranking by (master-only) × (small area) × (largest
byte size) picks `image1.png` correctly; `expected-tokens.json` pins it explicitly.

The `logo.frameInches` recorded in `expected-tokens.json` is **master1's** frame
(`x 10.0686, y 6.5356, w 2.4698, h 0.5277`), since master1 is the first/primary master and the one whose
layout family (L1-L4) is used for the cover. Master4's frame is listed above for reference.
`emittedOnMaster: true` means: the part is actually placed by a master's shape tree (not merely related).

---

## 8. Geometry (EMU → inches, 914400 EMU = 1 in)

### 8.1 `slideMaster1.xml`

| shape | id / name | placeholder | x | y | w | h |
|---|---|---|---|---|---|---|
| footer text `Security Level:` (10 pt) | 15 / `Text Placeholder 14` | `type="body" idx="1"` | 0.9944 | 6.8577 | 2.1679 | 0.1680 |
| logo picture (→ image1.png) | 38 / `图片 37` | none (`userDrawn`) | **10.0686** | **6.5356** | **2.4698** | **0.5277** |

Notable: this master has **no title placeholder at all**, and its only placeholder is a `body`-typed one
at the bottom-left holding the literal text "Security Level:" — i.e. the "body" placeholder is used as a
footer. A parser that maps `type="body"` to "main content area" will mis-place body text.

### 8.2 `slideMaster2.xml`

Empty `spTree` (`children: ['nvGrpSpPr', 'grpSpPr']`). No placeholders, no decorations. Legal but
unusual — a master with zero shapes.

### 8.3 `slideMaster3.xml`

| shape | placeholder | x | y | w | h | notes |
|---|---|---|---|---|---|---|
| `TextBox 2` — "Huawei Confidential" | none | 1.1980 | 6.9520 | 1.6005 | 0.2656 | footer text, runs at `sz="974"` = **9.74 pt**, colour `#1D1D1B` |
| `TextBox 3` — `<a:fld type="slidenum">` "‹#›" | none | 0.8029 | 7.0022 | 0.5465 | 0.1644 | **page-number field** |
| `图片 40` → `image6.png` | none (`userDrawn`) | 11.1505 | 6.9115 | 1.3879 | 0.2965 | logo variant, bottom-right |
| `组合 24` (`<p:grpSp>`, nested) | none | **13.3599** | 3.2064 | 2.1685 | 4.3030 | **off-canvas** (slide is 13.3385 wide) |

The group shapes (`组合 24/25/26`, nested 3 deep) are a **brand colour-swatch board**: their text reads
`PANTONE 186C / RGB 200/16/46`, `PANTONE 185C / RGB 199/0/11`, `品牌色` ("brand colours"),
`辅助色` ("secondary colours"), plus `RGB 234/90/79`, `RGB 120/0/15`, `RGB 248/181/60`, `RGB 235/92/1`,
`RGB 137/137/137`, `RGB 35/24/21`, `RGB 221/221/221`, `RGB 233/140/128`, `RGB 159/0/1`,
`RGB 245/220/87`, `RGB 240/133/0`, `RGB 181/181/181`, `RGB 89/87/87`, `RGB 255/255/255`.
That board is the origin of most of the literal hex values in the master census (C8102E, EA5A4F, 78000F,
F8B53C, EB5C01, 898989, 231815, DDDDDD, E98C80, A72126, 9F0001, F5DC57, F08500, B5B5B5, 595757) —
they are **reference swatches, not applied styling**. Do not promote them to design tokens.
It also confirms the intended brand reds: **PANTONE 186C = RGB 200/16/46 = `#C8102E`** and
**PANTONE 185C = RGB 199/0/11 = `#C7000B`** — the theme's `accent1 = #C7000A` is the 185C-ish value.

### 8.4 `slideMaster4.xml` (end page)

| shape | placeholder | x | y | w | h |
|---|---|---|---|---|---|
| `Title Placeholder 1` | `type="title"` (no `idx`) | 0.6743 | 1.6123 | 4.3572 | 3.0797 |
| `Text Placeholder 1` — copyright / disclaimer | none (userDrawn) | 8.7263 | 3.0566 | 3.5271 | 2.2200 |
| `Subtitle 6` — 把数字世界带入每个人… | none | 8.7350 | 1.7846 | 3.8033 | 0.5371 |
| `Subtitle 6` — Bring digital to every person… | none | 8.7245 | 2.3033 | 3.8078 | 0.6374 |
| `图片 98` → `image1.png` | none (`userDrawn`) | 8.7221 | 5.7427 | 2.0446 | 0.4369 |

### 8.5 Layouts — persistent decorations and placeholders

Every layout 1-4 shares the same geometry skeleton:

| element | x | y | w | h |
|---|---|---|---|---|
| full-bleed photo (`pic`, stretches to slide width) | 0.0000 | 0.0000 | **13.3385** | **6.1267** |
| corner mark `L 形` (`prstGeom prst="corner"`, `rot=5400000` = 90°, fill `#C7000B`) | L1 8.5892 / L2 6.3692 / L3 8.6723 / L4 6.5021 | L1 2.3300 / L2 2.8091 / L3 1.8127 / L4 2.5410 | 0.7667 | 0.7851 |
| `Title 1` placeholder `type="ctrTitle"` | 0.9832 | 0.9920 | 7.1739 | 0.7549 |
| `Text Placeholder 5` `type="body" idx="10"` | 1.0163 | 2.1319 | 7.1477 | 0.7042 |

The photo covers only the top **6.1267 in of 7.5000 in** (81.7 % of the height), i.e. it is deliberately
not full-height; the bottom band is white and carries master1's logo/footer. `slideLayout3`'s photo is
nudged to `y = -0.0396` and is 6.1663 in tall.

| layout | elements (x, y, w, h in inches) |
|---|---|
| L1 探索 | pic 0, 0, 13.3385, 6.1267; corner 8.5892, 2.3300, 0.7667, 0.7851 (rot 90°); ctrTitle 0.9832, 0.9920, 7.1739, 0.7549; body idx10 1.0163, 2.1319, 7.1477, 0.7042 |
| L2 智能 | pic 0, 0, 13.3385, 6.1267; corner 6.3692, 2.8091, 0.7667, 0.7851; ctrTitle & body identical to L1 |
| L3 攀登 | pic 0, **-0.0396**, 13.3385, **6.1663**; corner 8.6723, 1.8127, 0.7667, 0.7851; ctrTitle & body identical to L1 |
| L4 灯塔 | pic 0, 0, 13.3385, 6.1267; corner 6.5021, 2.5410, 0.7667, 0.7851; ctrTitle & body identical to L1 |
| L5 1_Contents page | body idx10 1.1304, 2.0156, 11.0696, 3.2958; `cxnSp` rule 1.1263, 1.4756, 0.9689, **0.0000** (flipH); text box 1.0049, 0.6894, 1.2549, 0.7147 |
| L6 1_Chinese text page | `Subtitle 2` `type="subTitle" idx="1"` 0.7974, 0.4988, 11.7461, 1.0864; `Content Placeholder 2` (`<p:ph idx="10"/>` with **no `type`**) 0.8059, 1.6426, 11.7384, 5.1295 |
| L7 End page | single text box, no placeholders at all: 0.6644, 1.5333, 4.2881, 0.9347 |

**Persistent decorations summary:** the *only* persistent logo is the master picture (§7.1). There are
**no footer/date/slide-number placeholders** anywhere — every `<p:hf>` is `hdr="0" ftr="0" dt="0"`, and
the page number is a hand-placed `<a:fld type="slidenum">` text box on master3 only (0.8029, 7.0022).
master1's "Security Level:" text is likewise a hand-placed text box that happens to be a `body`
placeholder.

### 8.6 Slides — geometry inheritance

Most slide placeholders have **no `<a:xfrm>` of their own** and inherit from their layout:

| slide | shape | placeholder | resolved geometry | source |
|---|---|---|---|---|
| slide1 | `标题 2` | `ctrTitle` | 0.9832, 0.9920, 7.1739, 0.7549 | **layout** (L1) |
| slide1 | `Text Placeholder 7` | `body idx=10` | 1.0163, 2.1319, 7.1477, 0.7042 | **layout** |
| slide1 | `Text Placeholder 3` | none | 1.0376, 6.8505, 2.8509, 0.1495 | own |
| slide2 | `Text Placeholder 1` | `body idx=10` | 1.1304, 2.0156, 10.9861, 3.2958 | own |
| slide3 | `Subtitle 1` | `subTitle idx=1` | 0.7974, 0.4988, 11.7461, 1.0864 | **layout** (L6) |
| slide3 | `Content Placeholder 2` | `idx=10` (no type) | 0.8059, 1.6426, 11.7384, 5.1295 | **layout** |
| slide3 | `TextBox 7` | none | 6.1903, **-1.3000**, 0.2020, 0.5778 | own — above the top edge |
| slide4 | `Subtitle 11` | `subTitle idx=1` | 0.7974, 0.5159, 11.7461, 1.0864 | own (y shifted 0.017 in) |
| slide4 | `Table 4` | — | 5.4171, 2.4981, 7.0081, 4.0527 | own |
| slide4 | `Chart 5` | — | 0.8498, 2.8515, 4.2127, 3.1224 | own |
| slide4 | `Text Box 22` (rot **-20°**) | none | 7.4120, 5.5531, 5.8526, 0.4941 | own |
| slide5 | *(empty spTree — no shapes at all)* | | | |

Note slide4's chart sits at 0.8498 in — left of the table — and overlaps nothing, but the rotated text box
at -20° overlaps the table's lower-right corner region; rotation is not reflected in `a:ext`.

---

## 9. Anything that would break a naive parser (consolidated trap list)

**Colour / theme**
1. `dk1 == lt1 == #1D1D1A`, `dk2 == lt2 == #FFFFFF`. `bg1 == tx1` and `bg2 == tx2`. Any code assuming
   `lt1 = background = #FFFFFF` and `dk1 = text = #000000` produces an inverted palette.
2. Six themes exist; `theme1` is "the" deck theme only via `presentation.xml.rels` rId14. Merging all
   themes (or "walk every `*.xml`") over-counts colours — theme2 genuinely differs in `accent2..6`.
3. `theme3` (which the real content pages use) has `lt1=#666666`, `lt2=#EBEBEB`.
4. `theme5`/`theme6` use `<a:sysClr val="windowText" lastClr="000000"/>` / `<a:sysClr val="window"
   lastClr="FFFFFF"/>`. `lastClr` is **optional** in the schema — read it, don't assume it, and fall back
   for `window`/`windowText`/`highlight`/`btnFace`.
5. Near-miss literals that must not be "corrected": theme `accent1=C7000A` vs layer `C7000B`;
   theme `dk1=1D1D1A` vs master1 body text `1D1D1B`; swatch board `C8102E`.
6. `<a:extraClrSchemeLst/>` is empty; `a:fmtScheme` (`name="Office"`) contains **3 `<a:gradFill>`
   entries** (and 8 `<a:solidFill>`) that are never used by the masters — a "collect all fills" pass will
   report gradients the deck does not paint.
7. Master1 `bodyStyle` uses a literal `srgbClr 1D1D1B` instead of `schemeClr tx1`, so scheme resolution
   alone misses the most common text colour in the masters.
8. `hlink=#919191` / `folHlink=#C4C4C4` are both greys (unusual) — a link-colour check that expects a
   blue will flag this deck.

**Fonts**
9. `<a:ea typeface=""/>` and `<a:cs typeface=""/>` are empty strings in both major and minor fonts;
   the CJK faces live in `<a:font script="Hans">` (微软雅黑 / 黑体).
10. Master `txStyles` override the theme (master1 title = Microsoft YaHei, not Arial Black). Reading
    `fontScheme` alone yields the wrong title face.
11. `presentation.xml defaultTextStyle` is 18 pt `tx1` `+mn-*`, while master1 body is 10 pt — three
    different "default body sizes" coexist (10 pt master1, 18.19 pt master4, 18 pt defaultTextStyle),
    plus layout1/2/3/4 body `defRPr sz="1400"` (14 pt).

**Structure / geometry**
12. `sldSz cx=12196763` ≠ 12192000 (13.3385 in, not 13.3333 in). 12196763 px @96 dpi = 1280.5.
13. **4 masters, 7 layouts, 6 themes** — not the usual 1 master / 11-12 layouts. `masters` must be 4.
14. No `<p:sldLayout>` has a `@type`; identify by `cSld/@name`. Masters have no `cSld/@name` at all.
15. Most slide placeholders have no `<a:xfrm>` → resolve slide → layout → master by `(type, idx)`;
    `@idx` is absent in places (defaults to `0`), `@type` is absent in others (defaults to `obj`).
    `ctrTitle` (layout) vs `title` (master4) are distinct types for the same visual role.
16. master1's only placeholder is `type="body"` used as a footer ("Security Level:"); master4 has the
    only `title` placeholder. So "first body placeholder = content" is wrong here.
17. `slideMaster2` has an **empty spTree** — zero shapes, zero placeholders.
18. `slideMaster3` has 3 nested `<p:grpSp>` positioned **off-canvas** (`x=13.3599 in` > 13.3385 in slide
    width), and groups have `chOff/chExt ≠ off/ext`, so child coordinates need the affine transform.
    A bbox/overflow check will (correctly) report out-of-page shapes.
19. Decorations are `userDrawn` pictures/text boxes **without** `<p:ph>` — placeholder-only extraction
    misses the logo, the footer, and the page-number field.
20. Page number is `<a:fld type="slidenum">` (rendered as "‹#›"), not a `<p:ph type="sldNum">`;
    all `<p:hf>` flags are 0, so there is no footer placeholder to find.
21. Off-page content exists by design: slide3 `TextBox 7` at `y = -1.3000`; layout3 photo at
    `y = -0.0396`; master3 swatch groups at `x >` slide width.
22. Zero-height line shapes exist (`cxnSp` on layout5: `h=0.0000`); `prstGeom prst="corner"` with
    `<a:gd adj1/adj2>` and `rot="5400000"` (90°) on the layouts' corner marks.
23. slide4 mixes `<a:tbl>` (graphicFrame), a `<c:chart>` graphicFrame bound to an embedded xlsx, and a
    rotated text box (`rot=-20°`); slide5's spTree is empty (an "end page" carrying only layout content).
24. `a:blip` elements carry `<a14:useLocalDpi val="0"/>` extensions (logo) — DPI extensions can fool
    naive "natural size from DPI" calculations; prefer the frame's `a:ext`.
25. Every `slide*.xml.rels` has **no** image relationships: images arrive only via masters/layouts, so a
    slide-level media walk finds nothing.
26. `docProps/custom.xml` contains opaque base64-looking strings and an empty `<vt:lpwstr/>`; don't
    assume every custom property has text.
27. Encoding: the package's XML is UTF-8 without BOM; layout names and master3 text are Chinese
    (`探索`, `智能`, `攀登`, `灯塔`, `品牌色`, `辅助色`, `HUAWEI`). Read as UTF-8; on Windows do not let
    the console default to GBK.

---

## 10. Verification performed

Everything in this report is reproducible with the delivered script. Exact commands and observed results:

```
$ node research/dump-template.mjs                # exit 0, 309 lines
$ node research/dump-template.mjs --json         # exit 0, JSON identical to research/expected-tokens.json
```

* `research/dump-template.output.txt` — the actual terminal output of the non-JSON run (25 589 bytes,
  309 lines, UTF-8, no BOM).
* Independent cross-check inside the script: `python-pptx 1.0.2` reports
  `slide_width=12196763, slide_height=6858000, masters=4, layouts=7, slides=5` and the script asserts
  `agrees with raw XML: True`.
* `research/raw-theme-fragments.txt` — verbatim (byte-for-byte substring) extracts of
  `theme1` `clrScheme` / `fontScheme` / `fmtScheme` / `extraClrSchemeLst`, `presentation.xml`
  `sldSz` / `notesSz` / `sldMasterIdLst` / `defaultTextStyle`, `slideMaster1` `bg` / `clrMap` /
  `titleStyle` / `bodyStyle` / logo `p:pic` / footer `p:sp`, and `theme5`/`theme6` `sysClr` forms.
  These were produced by `zipfile` + regex on the raw part text, so they are literal, not re-serialised.
* Image pixel sizes were read from the PNG `IHDR` chunk and JPEG `SOFn` markers (not from python-pptx).
* The logo identification was additionally confirmed by rendering both small images and reading the
  eight-petal Huawei flower + "HUAWEI" wordmark.

Numbers to trust (they are asserted in `research/expected-tokens.json`):
`cx=12196763`, `cy=6858000`, `13.3385 × 7.5 in`, 16:9; 4 masters, 7 layouts,
logo `ppt/media/image1.png` 529×116 @ 22 308 bytes,
frame `x 10.0686, y 6.5356, w 2.4698, h 0.5277 in` (master1).
