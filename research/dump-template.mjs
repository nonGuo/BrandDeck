#!/usr/bin/env node
/**
 * research/dump-template.mjs
 *
 * Self-contained dumper for the HW template .pptx. Prints the design-token facts that
 * research/template-report.md and research/expected-tokens.json are built from:
 *   - theme clrScheme (raw XML form per slot + effective #hex, sysClr/lastClr aware)
 *   - fontScheme (major/minor, latin/ea/cs separately) + master txStyles overrides
 *   - slide size (EMU -> inches, aspect, 16:9 check)
 *   - master / layout inventory with names, owning master, theme
 *   - media inventory: bytes, pixel size, every .rels that references it, classification
 *   - placeholder geometry in inches, with slide -> layout -> master inheritance resolved
 *   - background / fill colours
 *   - parser traps found in this file
 *
 * Usage:
 *   node research/dump-template.mjs [path\to\deck.pptx] [--json]
 *   node research/dump-template.mjs --python C:\path\to\python.exe
 *
 * Implementation note: ZIP + XML parsing is delegated to the bundled CPython
 * (stdlib only: zipfile + xml.etree + struct). No npm dependencies, no python-pptx
 * requirement. If python-pptx is importable it is used for one cross-check line.
 */

import { spawnSync } from 'node:child_process';
import { mkdtempSync, writeFileSync, rmSync, existsSync } from 'node:fs';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const HERE = path.dirname(fileURLToPath(import.meta.url));

// ---------------------------------------------------------------- args
const argv = process.argv.slice(2);
const flags = new Set(argv.filter((a) => a.startsWith('--')));
let pptx = argv.find((a) => !a.startsWith('--') && !/python\.exe$/i.test(a));
let pythonArg = null;
const pyIdx = argv.indexOf('--python');
if (pyIdx !== -1) pythonArg = argv[pyIdx + 1];
if (!pptx) pptx = path.resolve(HERE, '..', 'PPT模板-HW.pptx');
pptx = path.resolve(pptx);
const wantJson = flags.has('--json');

if (!existsSync(pptx)) {
  console.error('[dump-template] input not found: ' + pptx);
  process.exit(2);
}

// ---------------------------------------------------------------- python discovery
const PY_CANDIDATES = [
  pythonArg,
  process.env.DSH_PYTHON,
  process.env.PYTHON,
  'C:\\Users\\Administrator\\.dsh\\dsh-runtimes\\dsh-primary-runtime\\dependencies\\python\\python.exe',
  'python',
  'python3',
  'py',
];

function probe(cmd) {
  if (!cmd) return false;
  const args = ['-c', 'import sys; sys.stdout.write(sys.version.split()[0])'];
  const r = spawnSync(cmd, args, { encoding: 'utf8' });
  return r.status === 0 && /\d+\.\d+/.test(r.stdout || '');
}

let python = null;
for (const c of PY_CANDIDATES) {
  if (probe(c)) { python = c; break; }
}
if (!python) {
  console.error('[dump-template] no usable python interpreter found.');
  console.error('  tried: ' + PY_CANDIDATES.filter(Boolean).join(', '));
  console.error('  pass one explicitly:  node research/dump-template.mjs --python C:\\path\\python.exe');
  process.exit(3);
}

// ---------------------------------------------------------------- embedded python program
/* eslint-disable */
const PY_SOURCE = String.raw`
import sys, os, re, json, zipfile, posixpath
import xml.etree.ElementTree as ET

NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
NS_P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
A = '{' + NS_A + '}'
P = '{' + NS_P + '}'
R = '{' + NS_R + '}'
EMU_PER_IN = 914400.0
EMU_PER_PT = 12700.0

try:
    ET.register_namespace('a', NS_A)
    ET.register_namespace('p', NS_P)
    ET.register_namespace('r', NS_R)
except Exception:
    pass

SLOTS = ['dk1', 'lt1', 'dk2', 'lt2', 'accent1', 'accent2', 'accent3',
         'accent4', 'accent5', 'accent6', 'hlink', 'folHlink']

SYS_FALLBACK = {'windowText': '000000', 'window': 'FFFFFF', 'highlight': '0078D7',
                'btnFace': 'F0F0F0', 'btnText': '000000'}


def local(el):
    tag = el.tag if hasattr(el, 'tag') else el
    if not isinstance(tag, str):
        return ''
    return tag.rsplit('}', 1)[-1]


def inch(v):
    if v is None:
        return None
    return round(int(v) / EMU_PER_IN, 4)


def fmt_in(v):
    return 'n/a' if v is None else '%.4f' % v


def h(x):
    return '#' + x.upper()


def natural_key(s):
    return [int(t) if t.isdigit() else t.lower() for t in re.split(r'(\d+)', s)]


class Deck:
    def __init__(self, path):
        self.path = path
        self.z = zipfile.ZipFile(path)
        self.names = self.z.namelist()
        self._part_cache = {}
        self._rels_cache = {}

    def has(self, part):
        return part in self.names

    def read(self, part):
        return self.z.read(part)

    def text(self, part):
        if part not in self._part_cache:
            self._part_cache[part] = self.read(part).decode('utf-8')
        return self._part_cache[part]

    def root(self, part):
        return ET.fromstring(self.read(part))

    def rels_path(self, part):
        d, f = posixpath.split(part)
        return posixpath.join(d, '_rels', f + '.rels')

    def rels(self, part):
        """id -> (type_short, resolved_part_or_target, target_mode)"""
        if part in self._rels_cache:
            return self._rels_cache[part]
        out = {}
        rp = self.rels_path(part)
        if rp in self.names:
            base = posixpath.dirname(part)
            for rel in self.root(rp):
                tgt = rel.get('Target')
                mode = rel.get('TargetMode') or 'Internal'
                if mode == 'External':
                    resolved = tgt
                else:
                    resolved = posixpath.normpath(posixpath.join(base, tgt))
                out[rel.get('Id')] = (rel.get('Type').rsplit('/', 1)[-1], resolved, mode)
        self._rels_cache[part] = out
        return out

    def parts_under(self, prefix, suffix='.xml'):
        return sorted([n for n in self.names if n.startswith(prefix) and n.endswith(suffix)],
                      key=natural_key)


# ------------------------------------------------------------------ image size
def px_size(data):
    if data[:8] == b'\x89PNG\r\n\x1a\n':
        import struct
        w, ht = struct.unpack('>II', data[16:24])
        return w, ht, 'png'
    if data[:2] == b'\xff\xd8':
        import struct
        i, n = 2, len(data)
        while i < n - 9:
            if data[i] != 0xFF:
                i += 1
                continue
            m = data[i + 1]
            if m in (0xC0, 0xC1, 0xC2, 0xC3, 0xC5, 0xC6, 0xC7, 0xC9, 0xCA, 0xCB, 0xCD, 0xCE, 0xCF):
                ht, w = struct.unpack('>HH', data[i + 5:i + 9])
                return w, ht, 'jpeg'
            if m in (0xD8, 0xD9) or 0xD0 <= m <= 0xD7:
                i += 2
                continue
            seglen = struct.unpack('>H', data[i + 2:i + 4])[0]
            i += 2 + seglen
    if data[:6] in (b'GIF87a', b'GIF89a'):
        import struct
        w, ht = struct.unpack('<HH', data[6:10])
        return w, ht, 'gif'
    if data[:2] == b'BM':
        import struct
        w, ht = struct.unpack('<ii', data[18:26])
        return abs(w), abs(ht), 'bmp'
    return None, None, 'unknown'


# ------------------------------------------------------------------ theme helpers
def parse_clr_slot(theme_raw, slot):
    """Return dict(kind, raw, val, lastClr, effective) for one clrScheme slot."""
    m = re.search(r'<a:' + slot + r'\b[^>]*>(.*?)</a:' + slot + r'>', theme_raw, re.S)
    if not m:
        return {'kind': None, 'raw': None, 'val': None, 'lastClr': None,
                'effective': None, 'note': 'slot missing'}
    inner = m.group(1).strip()
    kind = 'srgbClr' if inner.startswith('<a:srgbClr') else ('sysClr' if inner.startswith('<a:sysClr') else None)
    if kind == 'srgbClr':
        val = re.search(r'val="([0-9A-Fa-f]{6})"', inner).group(1)
        return {'kind': 'srgbClr', 'raw': inner, 'val': val.upper(),
                'lastClr': None, 'effective': h(val), 'note': None}
    if kind == 'sysClr':
        vm = re.search(r'val="([^"]*)"', inner)
        lm = re.search(r'lastClr="([0-9A-Fa-f]{6})"', inner)
        val = vm.group(1) if vm else None
        last = lm.group(1).upper() if lm else None
        note = None
        if last:
            eff = h(last)
        else:
            eff = h(SYS_FALLBACK.get(val, '000000'))
            note = 'lastClr MISSING -> fell back to %s' % eff
        return {'kind': 'sysClr', 'raw': inner, 'val': val, 'lastClr': last,
                'effective': eff, 'note': note}
    # schemeClr / prstClr etc.
    vm = re.search(r'val="([^"]*)"', inner)
    val = vm.group(1).upper() if vm else None
    return {'kind': kind or inner[:24], 'raw': inner, 'val': val, 'lastClr': None,
            'effective': ('#' + val) if val and re.fullmatch(r'[0-9A-F]{6}', val) else None,
            'note': 'unhandled fill kind'}


def font_slot(font_el, tag):
    if font_el is None:
        return None
    e = font_el.find(A + tag)
    return None if e is None else e.get('typeface')


def resolve_typeface(tf, major_latin, minor_latin, major_ea, minor_ea, major_cs, minor_cs, major_scripts, minor_scripts):
    """Resolve scheme placeholders (+mj-lt / +mn-ea / ...) and empty strings."""
    if tf is None:
        return None
    table = {'+mj-lt': major_latin, '+mn-lt': minor_latin,
             '+mj-ea': major_ea, '+mn-ea': minor_ea,
             '+mj-cs': major_cs, '+mn-cs': minor_cs}
    if tf in table:
        return table[tf]
    return tf


def placeholder_geometry(shape):
    """own xfrm in inches, or None"""
    cands = (P + 'spPr/' + A + 'xfrm',
             P + 'grpSpPr/' + A + 'xfrm',
             P + 'xfrm',
             A + 'xfrm')
    for cand in cands:
        x = shape.find(cand)
        if x is None and cand == A + 'xfrm':
            x = shape.find('.//' + A + 'xfrm')
        if x is not None:
            off = x.find(A + 'off')
            ext = x.find(A + 'ext')
            d = {}
            if off is not None:
                d['x'] = inch(off.get('x'))
                d['y'] = inch(off.get('y'))
            if ext is not None:
                d['w'] = inch(ext.get('cx'))
                d['h'] = inch(ext.get('cy'))
            if x.get('rot'):
                d['rot'] = int(x.get('rot')) / 60000.0
            return d
    return None


def ph_key(shape):
    ph = shape.find('.//' + P + 'nvPr/' + P + 'ph')
    if ph is None:
        return None
    return (ph.get('type') or 'obj', ph.get('idx') or '0')


def ph_attr(shape):
    ph = shape.find('.//' + P + 'nvPr/' + P + 'ph')
    if ph is None:
        return None
    return {'type': ph.get('type'), 'idx': ph.get('idx')}


def shape_fills(shape):
    out = []
    spPr = shape.find(P + 'spPr')
    if spPr is not None:
        for ch in spPr:
            tag = local(ch)
            if tag == 'solidFill' and len(ch):
                c = ch[0]
                out.append('solidFill:' + local(c) + ':' + str(c.get('val')))
            elif tag in ('gradFill', 'blipFill', 'pattFill', 'noFill'):
                out.append(tag)
    fr = shape.find('.//' + A + 'style/' + A + 'fillRef')
    if fr is not None:
        out.append('style.fillRef(idx=%s)' % fr.get('idx'))
    return out


def bg_info(part_root):
    """Compact, readable description of <p:bg>, including the exact colour value."""
    bg = part_root.find('.//' + P + 'bg')
    if bg is None:
        return None
    parts = []
    for el in bg.iter():
        tag = local(el)
        if tag == 'solidFill':
            kid = el[0] if len(el) else None
            if kid is not None:
                parts.append('solidFill ' + local(kid) + '(' + str(kid.get('val')) + ')')
        elif tag in ('gradFill', 'blipFill', 'pattFill', 'noFill'):
            parts.append(tag)
    s = ET.tostring(bg, encoding='unicode')
    s = re.sub(r'\s+', ' ', s).strip()
    return ('%s   [raw: %s]' % (' + '.join(parts) or '(empty)', s)) if parts else s


def clrmapovr_info(root):
    cmo = root.find(P + 'clrMapOvr')
    if cmo is None:
        return 'none'
    if cmo.find(A + 'masterClrMapping') is not None:
        return 'masterClrMapping (inherit master clrMap)'
    ov = cmo.find(A + 'overrideClrMapping')
    if ov is not None:
        return 'overrideClrMapping ' + ' '.join('%s=%s' % (k, v) for k, v in ov.attrib.items())
    return 'unknown'


# ------------------------------------------------------------------ main
def main():
    argv = sys.argv[1:]
    json_mode = '--json' in argv
    args = [a for a in argv if not a.startswith('--')]
    pptx_path = args[0]
    d = Deck(pptx_path)
    out = []
    W = out.append

    # ---------------- package
    total = sum(i.file_size for i in d.z.infolist())
    W('###############################################################################')
    W('# PPTX TEMPLATE DUMP')
    W('###############################################################################')
    W('file            : %s' % pptx_path)
    W('file size       : %d bytes (on disk)' % os.path.getsize(pptx_path))
    W('parts           : %d   uncompressed total: %d bytes' % (len(d.names), total))
    W('masters         : %d' % len(d.parts_under('ppt/slideMasters/', '.xml')))
    W('layouts         : %d' % len(d.parts_under('ppt/slideLayouts/', '.xml')))
    W('slides          : %d' % len(d.parts_under('ppt/slides/', '.xml')))
    W('themes          : %d  -> %s' % (len(d.parts_under('ppt/theme/', '.xml')),
                                       [posixpath.basename(x) for x in d.parts_under('ppt/theme/', '.xml')]))
    W('media parts     : %d' % len(d.parts_under('ppt/media/', '')))
    W('themeOverride   : %s' % ([n for n in d.names if 'themeOverride' in n] or 'none'))
    W('charts          : %s' % (d.parts_under('ppt/charts/', '.xml') or 'none'))
    W('embedded pkgs   : %s' % (d.parts_under('ppt/embeddings/', '') or 'none'))
    W('notesMasters    : %s' % (d.parts_under('ppt/notesMasters/', '.xml') or 'none'))
    W('handoutMasters  : %s' % (d.parts_under('ppt/handoutMasters/', '.xml') or 'none'))
    W('')

    # ---------------- slide size
    pres = d.root('ppt/presentation.xml')
    sldSz = pres.find(P + 'sldSz')
    cx = int(sldSz.get('cx'))
    cy = int(sldSz.get('cy'))
    w_in = round(cx / EMU_PER_IN, 4)
    h_in = round(cy / EMU_PER_IN, 4)
    ratio = cx / float(cy)
    W('=== SLIDE SIZE  (ppt/presentation.xml <p:sldSz>) ===')
    W('<p:sldSz cx="%d" cy="%d"/>' % (cx, cy))
    W('  width  : %d EMU = %.4f in = %.2f pt = %.2f px @96dpi' % (cx, w_in, cx / EMU_PER_PT, cx / EMU_PER_IN * 96))
    W('  height : %d EMU = %.4f in = %.2f pt = %.2f px @96dpi' % (cy, h_in, cy / EMU_PER_PT, cy / EMU_PER_IN * 96))
    W('  aspect : %.6f  (16:9 = 1.777778)  -> %s' % (ratio, '16:9' if abs(ratio - 16 / 9.0) < 0.01 else 'other'))
    if cx != 12192000:
        W('  TRAP   : cx is NOT the canonical 16:9 value 12192000 (13.3333 in).')
        W('           delta = %+d EMU (%+.5f in). Do not hardcode 12192000.' % (cx - 12192000, (cx - 12192000) / EMU_PER_IN))
    W('')

    # ---------------- themes
    W('=== THEME INVENTORY ===')
    themes = d.parts_under('ppt/theme/', '.xml')
    theme_names = {}
    for part in themes:
        raw = d.text(part)
        root = ET.fromstring(raw)
        clr = root.find('.//' + A + 'clrScheme')
        fnt = root.find('.//' + A + 'fontScheme')
        theme_names[part] = root.get('name')
        W('%s : theme@name=%r  clrScheme@name=%r  fontScheme@name=%r'
          % (part, root.get('name'), clr.get('name') if clr is not None else None,
             fnt.get('name') if fnt is not None else None))
    W('')

    # which theme is presentation-level
    pres_theme = None
    for rid, (typ, tgt, mode) in d.rels('ppt/presentation.xml').items():
        if typ == 'theme':
            pres_theme = tgt
    W('presentation-level theme (ppt/_rels/presentation.xml.rels) : %s' % pres_theme)
    W('')

    primary_theme = pres_theme or themes[0]
    raw = d.text(primary_theme)
    W('=== <a:clrScheme> of %s ===' % primary_theme)
    clr = ET.fromstring(raw).find('.//' + A + 'clrScheme')
    W('clrScheme name attribute: %r' % (clr.get('name') if clr is not None else None))
    W('%-9s %-8s %-34s %-9s %-9s' % ('slot', 'kind', 'raw XML form', 'val', 'effective'))
    colors_raw = {}
    for slot in SLOTS:
        info = parse_clr_slot(raw, slot)
        colors_raw[slot] = info
        W('%-9s %-8s %-34s %-9s %-9s%s'
          % (slot, info['kind'], info['raw'], info['val'], info['effective'],
             ('  <-- ' + info['note']) if info['note'] else ''))
    W('')
    W('  NOTE: dk1 == lt1 (%s) and dk2 == lt2 (%s) in this theme.'
      % (colors_raw['dk1']['effective'], colors_raw['dk2']['effective']))
    W('        The light/dark pairs are duplicated, so bg1==tx1 and bg2==tx2. A parser that')
    W('        assumes lt1==#FFFFFF (bg1) and dk1==#000000 (tx1) will invert this deck.')
    W('')

    # other themes: clrScheme only, compact
    others = [t for t in themes if t != primary_theme]
    if others:
        W('=== other themes: clrScheme (compact; only differences matter) ===')
        for part in others:
            r2 = d.text(part)
            vals = []
            for slot in SLOTS:
                i2 = parse_clr_slot(r2, slot)
                vals.append('%s=%s' % (slot, i2['effective']))
            sysc = [s for s in SLOTS if (parse_clr_slot(r2, s)['kind'] == 'sysClr')]
            W('%s (%r): %s' % (part, theme_names[part], ' '.join(vals)))
            if sysc:
                W('    sysClr slots: %s' % sysc)
                for s in sysc:
                    i2 = parse_clr_slot(r2, s)
                    W('      %s raw=%s lastClr=%s effective=%s%s'
                      % (s, i2['raw'], i2['lastClr'], i2['effective'],
                         ('  <-- ' + i2['note']) if i2['note'] else ''))
        W('')

    # ---------------- fonts
    W('=== <a:fontScheme> of %s ===' % primary_theme)
    fnt = ET.fromstring(raw).find('.//' + A + 'fontScheme')
    W('fontScheme name: %r' % fnt.get('name'))
    fs = {}
    for which in ('majorFont', 'minorFont'):
        fe = fnt.find(A + which)
        fs[which] = {t: font_slot(fe, t) for t in ('latin', 'ea', 'cs')}
        scripts = {f.get('script'): f.get('typeface') for f in fe.iter(A + 'font')}
        fs[which]['scripts'] = scripts
        W('  %s: latin=%r  ea=%r  cs=%r' % (which, fs[which]['latin'], fs[which]['ea'], fs[which]['cs']))
        W('        Hans(zh-CN) script face = %r ; Hant = %r'
          % (scripts.get('Hans'), scripts.get('Hant')))
    if fs['majorFont']['ea'] == '':
        W('  TRAP: <a:ea typeface=""/> is an EMPTY STRING, not absent. It means "use the')
        W('        per-script face" (Hans -> %r for majorFont). Reporting ea as "" is'
          % fs['majorFont']['scripts'].get('Hans'))
        W('        technically correct but useless downstream.')
    W('')

    # ---------------- masters: txStyles overrides
    master_parts = d.parts_under('ppt/slideMasters/', '.xml')
    W('=== MASTER txStyles overrides (these beat the theme fontScheme) ===')
    mj = fs['majorFont']
    mn = fs['minorFont']
    master_fonts = {}
    for part in master_parts:
        root = d.root(part)
        ts = root.find(P + 'txStyles')
        W('%s' % part)
        if ts is None:
            W('   (no <p:txStyles>)')
            continue
        for which in ('titleStyle', 'bodyStyle'):
            st = ts.find(P + which)
            if st is None:
                W('   %s: absent' % which)
                continue
            lvl1 = st.find(A + 'lvl1pPr')
            dr = lvl1.find(A + 'defRPr') if lvl1 is not None else None
            if dr is None:
                W('   %s: no lvl1 defRPr' % which)
                continue
            col = None
            sf = dr.find(A + 'solidFill')
            if sf is not None and len(sf):
                col = local(sf[0]) + ':' + str(sf[0].get('val'))
            fonts = {}
            for t in ('latin', 'ea', 'cs'):
                e = dr.find(A + t)
                fonts[t] = resolve_typeface(e.get('typeface') if e is not None else None,
                                            mj['latin'], mn['latin'], mj['ea'], mn['ea'],
                                            mj['cs'], mn['cs'], mj['scripts'], mn['scripts'])
                fonts[t + 'Raw'] = e.get('typeface') if e is not None else None
            sz = dr.get('sz')
            W('   %-11s lvl1: sz=%s (%s pt)  color=%s  latin=%r ea=%r cs=%r'
              % (which, sz, (int(sz) / 100.0) if sz else '?', col,
                 fonts['latin'], fonts['ea'], fonts['cs']))
            if which == 'titleStyle':
                master_fonts.setdefault(part, {})['title'] = fonts
            else:
                master_fonts.setdefault(part, {})['body'] = fonts
        hf = root.find(P + 'hf')
        if hf is not None:
            W('   <p:hf %s>' % ' '.join('%s=%r' % (k, v) for k, v in hf.attrib.items()))
    W('')

    dts = pres.find(P + 'defaultTextStyle')
    if dts is not None:
        lvl1 = dts.find(A + 'lvl1pPr')
        dr = lvl1.find(A + 'defRPr') if lvl1 is not None else None
        if dr is not None:
            W('=== ppt/presentation.xml <p:defaultTextStyle> lvl1 ===')
            W('   sz=%s (%s pt)  color=%s' % (dr.get('sz'), (int(dr.get('sz')) / 100.0) if dr.get('sz') else '?',
                                             (lambda sf: (local(sf[0]) + ':' + str(sf[0].get('val'))) if sf is not None and len(sf) else None)(dr.find(A + 'solidFill'))))
            W('')

    # ---------------- media inventory
    media_refs = {}
    for part in d.names:
        if not part.endswith('.rels'):
            continue
        src = part.replace('_rels/', '').rsplit('.rels', 1)[0]
        for rid, (typ, tgt, mode) in d.rels(src).items():
            if '/media/' in '/' + tgt:
                media_refs.setdefault(tgt, []).append((part, rid, typ))
    for k in media_refs:
        media_refs[k].sort()

    W('=== MEDIA INVENTORY (ppt/media/) ===')
    W('%-22s %9s %10s %-6s %s' % ('part', 'bytes', 'px', 'fmt', 'referenced by'))
    media_list = []
    for part in d.parts_under('ppt/media/', ''):
        data = d.read(part)
        w, ht, fmt = px_size(data)
        refs = media_refs.get(part, [])
        ref_parts = sorted({r[0] for r in refs})
        kinds = sorted({r[2] for r in refs})
        W('%-22s %9d %10s %-6s %s'
          % (part, len(data), ('%dx%d' % (w, ht)) if w else '?', fmt,
             ', '.join('%s(%s)' % (p, rid) for p, rid, _ in refs) or 'UNREFERENCED'))
        area = (w or 0) * (ht or 0)
        if all('/slideMasters/' in p for p in ref_parts) and ref_parts and area and area <= 120000:
            cls = 'logo'
        elif all('/slideMasters/' in p for p in ref_parts) and ref_parts:
            cls = 'master-only image'
        elif any('/slideLayouts/' in p for p in ref_parts) and area > 500000:
            cls = 'background/photo'
        elif ref_parts:
            cls = 'content image'
        else:
            cls = 'unreferenced'
        media_list.append({'path': part, 'bytes': len(data), 'widthPx': w, 'heightPx': ht,
                           'format': fmt, 'referencedBy': ref_parts,
                           'relIds': [{'rels': p, 'id': rid} for p, rid, _ in refs],
                           'classification': cls,
                           'areaPx': area or 0})
    W('')
    logos = [m for m in media_list if m['classification'] == 'logo']
    if logos:
        logos.sort(key=lambda m: (-(m['bytes']), m['path']))
        chosen = logos[0]
        W('  -> LOGO judgement: %s (%dx%d, %d bytes). Largest/least-compressed master-only small image.'
          % (chosen['path'], chosen['widthPx'], chosen['heightPx'], chosen['bytes']))
        for m in logos[1:]:
            W('     also logo-like: %s (%dx%d, %d bytes) - same artwork, lower resolution'
              % (m['path'], m['widthPx'], m['heightPx'], m['bytes']))
            m['classification'] = 'logo-variant'
    W('')

    # ---------------- master / layout graph
    pres_rels = d.rels('ppt/presentation.xml')
    master_order = []
    for smid in pres.iter(P + 'sldMasterId'):
        rid = smid.get(R + 'id')
        if rid in pres_rels:
            master_order.append(pres_rels[rid][1])
    if not master_order:
        master_order = master_parts

    layout_to_master = {}
    master_layouts = {}
    master_theme = {}
    for mpart in master_order:
        rels = d.rels(mpart)
        root = d.root(mpart)
        lays = []
        for lid in root.iter(P + 'sldLayoutId'):
            rid = lid.get(R + 'id')
            if rid in rels:
                lays.append(rels[rid][1])
        master_layouts[mpart] = lays
        for lay in lays:
            layout_to_master[lay] = mpart
        for rid, (typ, tgt, mode) in rels.items():
            if typ == 'theme':
                master_theme[mpart] = tgt

    def geom_index(part):
        """ph key -> (geometry, source)"""
        idx = {}
        root = d.root(part)
        for shape in root.iter():
            k = ph_key(shape)
            if k is None:
                continue
            g = placeholder_geometry(shape)
            if g:
                idx[k] = g
        return idx

    master_geom = {m: geom_index(m) for m in master_order}

    def resolve_geom(shape, own_part, layout_part=None, master_part=None):
        g = placeholder_geometry(shape)
        if g:
            return g, 'own'
        k = ph_key(shape)
        if k is None:
            return None, 'none (no xfrm, no placeholder)'
        if layout_part and layout_part in layout_geom:
            if k in layout_geom[layout_part]:
                return layout_geom[layout_part][k], 'layout'
        if master_part and master_part in master_geom:
            if k in master_geom[master_part]:
                return master_geom[master_part][k], 'master'
        return None, 'unresolved'

    # precompute layout geometry
    layout_geom = {}
    for lay in d.parts_under('ppt/slideLayouts/', '.xml'):
        merged = dict(master_geom.get(layout_to_master.get(lay), {}))
        merged.update(geom_index(lay))
        layout_geom[lay] = merged

    W('=== MASTER / LAYOUT INVENTORY ===')
    W('masters: %d   layouts: %d' % (len(master_order), len(d.parts_under('ppt/slideLayouts/', '.xml'))))
    for mpart in master_order:
        root = d.root(mpart)
        csld = root.find(P + 'cSld')
        W('')
        W('%s  name=%r  theme=%s  layouts=%d'
          % (mpart, csld.get('name') if csld is not None else None,
             master_theme.get(mpart), len(master_layouts[mpart])))
        W('   bg     : %s' % (bg_info(root) or '(none - inherits/white)'))
        cm = root.find(P + 'clrMap')
        if cm is not None:
            W('   clrMap : bg1=%s tx1=%s bg2=%s tx2=%s' % (cm.get('bg1'), cm.get('tx1'), cm.get('bg2'), cm.get('tx2')))
        shapes = [s for s in root.find(P + 'cSld/' + P + 'spTree') if local(s) not in ('nvGrpSpPr', 'grpSpPr')]
        W('   shapes : %d' % len(shapes))
        for s in shapes:
            kind = local(s)
            nm = s.find('.//' + P + 'cNvPr').get('name')
            pa = ph_attr(s)
            g = placeholder_geometry(s)
            extra = ''
            if kind == 'pic':
                bl = s.find('.//' + A + 'blip')
                rid = bl.get(R + 'embed') if bl is not None else None
                extra = ' image=%s' % (d.rels(mpart).get(rid, (None, '?'))[1] if rid else '?')
            W('     %-7s %-26r ph=%-32s own_geom=[x %s y %s w %s h %s]%s'
              % (kind, nm, str(pa),
                 fmt_in(g['x']) if g else '-', fmt_in(g['y']) if g else '-',
                 fmt_in(g['w']) if g else '-', fmt_in(g['h']) if g else '-', extra))

    all_layouts = d.parts_under('ppt/slideLayouts/', '.xml')
    layout_names = []
    W('')
    for i, lay in enumerate(all_layouts, 1):
        root = d.root(lay)
        csld = root.find(P + 'cSld')
        lname = csld.get('name') if csld is not None else None
        layout_names.append(lname)
        W('%s  [#%d]  name=%r  type=%r  master=%s  theme=%s'
          % (lay, i, lname, root.get('type'),
             layout_to_master.get(lay), master_theme.get(layout_to_master.get(lay))))
        if root.get('type') is None:
            W('     TRAP: no @type attribute on <p:sldLayout> -> the layout has no standard')
            W('           type (title/obj/...). Infer role from cSld/@name instead.')
        W('     bg    : %s' % (bg_info(root) or '(none -> inherits master #FFFFFF)'))
        cmo = root.find(P + 'clrMapOvr')
        if cmo is not None:
            W('     clrMapOvr: %s' % clrmapovr_info(root))
        shapes = [s for s in root.find(P + 'cSld/' + P + 'spTree') if local(s) not in ('nvGrpSpPr', 'grpSpPr')]
        for s in shapes:
            kind = local(s)
            nm = s.find('.//' + P + 'cNvPr').get('name')
            pa = ph_attr(s)
            if pa is not None:
                g, src = resolve_geom(s, lay, lay, layout_to_master.get(lay))
            else:
                g = placeholder_geometry(s)
                src = 'own'
            extra = ''
            if kind == 'pic':
                bl = s.find('.//' + A + 'blip')
                rid = bl.get(R + 'embed') if bl is not None else None
                extra = ' image=%s' % (d.rels(lay).get(rid, (None, '?'))[1] if rid else '?')
            W('     %-7s %-22r ph=%-30s geom(x %s y %s w %s h %s) from=%s%s'
              % (kind, nm, str(pa),
                 fmt_in(g['x']) if g else '-', fmt_in(g['y']) if g else '-',
                 fmt_in(g['w']) if g else '-', fmt_in(g['h']) if g else '-', src, extra))
    W('')

    # ---------------- slides
    W('=== SLIDES (geometry resolved slide -> layout -> master) ===')
    pres_slides = []
    for sid in pres.iter(P + 'sldId'):
        rid = sid.get(R + 'id')
        if rid in pres_rels:
            pres_slides.append(pres_rels[rid][1])
    W('%-26s %-32s %-32s %s' % ('slide', 'layout', 'master', 'theme'))
    for sl in pres_slides:
        rels = d.rels(sl)
        lay = None
        for rid, (typ, tgt, mode) in rels.items():
            if typ == 'slideLayout':
                lay = tgt
        mst = layout_to_master.get(lay)
        W('%-26s %-32s %-32s %s'
          % (sl, lay, mst, master_theme.get(mst)))
    W('')
    for sl in pres_slides:
        root = d.root(sl)
        rels = d.rels(sl)
        lay = None
        for rid, (typ, tgt, mode) in rels.items():
            if typ == 'slideLayout':
                lay = tgt
        W('%s  layout=%s' % (sl, lay))
        spTree = root.find(P + 'cSld/' + P + 'spTree')
        shapes = [s for s in spTree if local(s) not in ('nvGrpSpPr', 'grpSpPr')]
        if not shapes:
            W('     (empty spTree)')
        for s in shapes:
            if local(s) in ('nvGrpSpPr', 'grpSpPr'):
                continue
            pa = ph_attr(s)
            if pa is not None:
                g, src = resolve_geom(s, sl, lay, layout_to_master.get(lay))
            else:
                g = placeholder_geometry(s)
                src = 'own'
            W('     %-12s %-22r ph=%-30s geom(x %s y %s w %s h %s) from=%s'
              % (local(s), s.find('.//' + P + 'cNvPr').get('name'), str(pa),
                 fmt_in(g['x']) if g else '-', fmt_in(g['y']) if g else '-',
                 fmt_in(g['w']) if g else '-', fmt_in(g['h']) if g else '-', src))
    W('')

    # ---------------- persistent decorations / footers
    W('=== PERSISTENT DECORATIONS ON MASTERS (logo, footer text, page number) ===')
    for mpart in master_order:
        root = d.root(mpart)
        W('%s' % mpart)
        for s in root.find(P + 'cSld/' + P + 'spTree'):
            kind = local(s)
            if kind in ('nvGrpSpPr', 'grpSpPr', 'grpSp'):
                continue
            g = placeholder_geometry(s)
            nm = s.find('.//' + P + 'cNvPr').get('name')
            txts = [t.text for t in s.iter(A + 't') if t.text]
            flds = [(f.get('type'), ''.join(t.text or '' for t in f.iter(A + 't')))
                    for f in s.iter(A + 'fld')]
            if kind == 'pic':
                bl = s.find('.//' + A + 'blip')
                rid = bl.get(R + 'embed') if bl is not None else None
                W('   pic  %-18r geom(x %s y %s w %s h %s) image=%s'
                  % (nm, fmt_in(g['x']) if g else '-', fmt_in(g['y']) if g else '-',
                     fmt_in(g['w']) if g else '-', fmt_in(g['h']) if g else '-',
                     d.rels(mpart).get(rid, (None, '?'))[1] if rid else '?'))
            elif txts or flds:
                W('   text %-18r geom(x %s y %s w %s h %s) text=%r%s'
                  % (nm, fmt_in(g['x']) if g else '-', fmt_in(g['y']) if g else '-',
                     fmt_in(g['w']) if g else '-', fmt_in(g['h']) if g else '-',
                     ' '.join(txts)[:60],
                     ('  field=' + str(flds)) if flds else ''))
    W('')

    # ---------------- effective tokens
    tx1 = colors_raw['dk1']['effective']
    bg = colors_raw['dk2']['effective']
    # background actually painted by the master(s)
    master_bg = bg_info(d.root(master_order[0])) if master_order else None
    m = re.search(r'srgbClr val="([0-9A-Fa-f]{6})"', master_bg or '')
    painted_bg = h(m.group(1)) if m else bg

    logo = None
    if logos:
        chosen = sorted(logos, key=lambda m2: (-m2['bytes'], m2['path']))[0]
        frame = None
        for mpart in master_order:
            for s in d.root(mpart).iter(P + 'pic'):
                bl = s.find('.//' + A + 'blip')
                rid = bl.get(R + 'embed') if bl is not None else None
                if rid and d.rels(mpart).get(rid, (None, None))[1] == chosen['path']:
                    g = placeholder_geometry(s)
                    if g and frame is None:
                        frame = g
        logo = {'partPath': chosen['path'], 'bytes': chosen['bytes'],
                'widthPx': chosen['widthPx'], 'heightPx': chosen['heightPx'],
                'emittedOnMaster': frame is not None,
                'frameInches': frame}

    tokens = {
        'slideSize': {'cx': cx, 'cy': cy, 'widthIn': w_in, 'heightIn': h_in,
                      'aspect': '16:9' if abs(ratio - 16 / 9.0) < 0.01 else 'other'},
        'colors': {
            'background': painted_bg,
            'surface': '#EBEBEB',
            'textPrimary': tx1,
            'textSecondary': colors_raw['accent6']['effective'],
            'primary': colors_raw['accent1']['effective'],
            'accent': colors_raw['accent2']['effective'],
            'chartPalette': [colors_raw['accent%d' % i]['effective'] for i in range(1, 7)],
        },
        'colorsRaw': {s: {'kind': colors_raw[s]['kind'], 'val': colors_raw[s]['val'],
                          'lastClr': colors_raw[s]['lastClr'],
                          'effective': colors_raw[s]['effective']} for s in SLOTS},
        'fonts': {'title': (master_fonts.get(master_order[0], {}).get('title', {}) or {}).get('latin'),
                  'body': (master_fonts.get(master_order[0], {}).get('body', {}) or {}).get('latin'),
                  'latin': mn['latin'],
                  'cs': (master_fonts.get(master_order[0], {}).get('body', {}) or {}).get('cs') or mn['cs']},
        'logo': logo,
        'mediaParts': [{'path': m['path'], 'bytes': m['bytes'], 'widthPx': m['widthPx'],
                        'heightPx': m['heightPx'], 'referencedBy': m['referencedBy'],
                        'classification': m['classification']} for m in media_list],
        'masters': len(master_order),
        'layouts': len(all_layouts),
        'layoutNames': layout_names,
    }

    W('=== EFFECTIVE TOKENS (what expected-tokens.json asserts) ===')
    W('colors.background  : %s   (master <p:bg> solidFill)' % tokens['colors']['background'])
    W('colors.surface     : %s   (only non-white bg in the package: slideLayout5)' % tokens['colors']['surface'])
    W('colors.textPrimary : %s   (tx1 -> dk1)' % tokens['colors']['textPrimary'])
    W('colors.textSecondary: %s  (accent6, the grey used for de-emphasised text)' % tokens['colors']['textSecondary'])
    W('colors.primary     : %s   (accent1)' % tokens['colors']['primary'])
    W('colors.accent      : %s   (accent2)' % tokens['colors']['accent'])
    W('chartPalette       : %s' % ' '.join(tokens['colors']['chartPalette']))
    W('fonts.title        : %s  (master1 titleStyle latin+ea)' % tokens['fonts']['title'])
    W('fonts.body         : %s  (master1 bodyStyle latin)' % tokens['fonts']['body'])
    W('fonts.latin        : %s  (theme1 minorFont latin)' % tokens['fonts']['latin'])
    W('fonts.cs           : %s  (master1 bodyStyle cs; theme cs is "")' % tokens['fonts']['cs'])
    if logo:
        fr = logo['frameInches'] or {}
        W('logo               : %s %dx%d %d bytes emittedOnMaster=%s frame=(x %.4f, y %.4f, w %.4f, h %.4f) in'
          % (logo['partPath'], logo['widthPx'], logo['heightPx'], logo['bytes'], logo['emittedOnMaster'],
             fr.get('x', float('nan')), fr.get('y', float('nan')),
             fr.get('w', float('nan')), fr.get('h', float('nan'))))
    W('masters            : %d' % tokens['masters'])
    W('layouts            : %d' % tokens['layouts'])
    for i, nm in enumerate(layout_names, 1):
        W('   layout %d: %r' % (i, nm))
    W('')

    # -------------------------------- traps
    W('=== PARSING TRAPS FOUND IN THIS FILE ===')
    traps = []
    traps.append('dk1 == lt1 == %s and dk2 == lt2 == %s: the light/dark pairs are duplicated, so bg1 (lt1)'
                 % (colors_raw['dk1']['effective'], colors_raw['dk2']['effective']))
    traps.append('  and tx1 (dk1) are the SAME colour, and so are bg2 (lt2) and tx2 (dk2). A parser that')
    traps.append('  assumes lt1=#FFFFFF / dk1=#000000 will invert every text and background colour here.')
    traps.append('sldSz cx=%d, not the canonical 12192000 -> width is %s in, not 13.3333 in (%+.5f in).'
                 % (cx, w_in, (cx - 12192000) / EMU_PER_IN))
    traps.append('There are 4 slide masters and 6 themes. "The" theme is theme1 only because ppt/_rels/presentation.xml.rels')
    traps.append('  rId14 points at it. theme2 (master2) really differs in accent2..accent6 and hlink/folHlink:')
    traps.append('  C8102E / EA594F / F8B53C / 231815 / 595757, hlink 898989, folHlink B5B5B5.')
    traps.append('theme3 (master3, used by the real content pages) differs in lt1=#666666 and lt2=#EBEBEB.')
    traps.append('theme5/theme6 (notes/handout masters) use <a:sysClr val="windowText" lastClr="000000"/> and')
    traps.append('  <a:sysClr val="window" lastClr="FFFFFF"/> - a parser must read lastClr, not @val; and lastClr')
    traps.append('  is OPTIONAL in the schema, so the code path that reads it can be missing it.')
    traps.append('<a:ea typeface=""/> and <a:cs typeface=""/> are EMPTY STRINGS in both majorFont and minorFont;')
    traps.append('  the real CJK face lives in <a:font script="Hans"> (majorFont 微软雅黑, minorFont 黑体).')
    traps.append('Master txStyles override the theme: master1 title = Microsoft YaHei (NOT the theme face "Arial Black").')
    traps.append('image1.png is referenced by TWO masters (slideMaster1 rId6 and slideMaster4 rId3), so "media part')
    traps.append('  referenced by exactly one master" is not a valid logo heuristic here.')
    traps.append('The logo is a userDrawn <p:pic> on the master with no <p:ph>, placed BOTTOM-right, not top-left.')
    traps.append('Layouts 1-4 have NO <p:bg>: the background comes from the master. Layout5 bg=#EBEBEB, layout7 bg=#FFFFFF.')
    traps.append('No <p:sldLayout> carries a @type attribute; identify layouts by cSld/@name (Chinese names present).')
    traps.append('Most slide placeholders carry NO <a:xfrm> -> geometry must be resolved slide -> layout -> master by')
    traps.append('  (type, idx); <p:ph> @idx is absent on some placeholders (defaults to 0) and @type is absent on others.')
    traps.append('slideMaster2 has an EMPTY spTree (no placeholders at all) - a master with zero shapes is legal.')
    traps.append('slideMaster3 holds group shapes (<p:grpSp>, nested 3 deep) pushed off-canvas at x=13.3599 in')
    traps.append('  (slide is 13.3385 in wide): a brand colour-swatch board. Naive bbox walks report out-of-page shapes.')
    traps.append('Group shapes have chOff/chExt != off/ext, so child coordinates need the affine transform.')
    traps.append('Near-miss hex values: theme accent1=C7000A but the layouts paint C7000B; master1 body text uses')
    traps.append('  1D1D1B while theme dk1 is 1D1D1A; master3 swatch board uses C8102E. Do not "normalise" these.')
    traps.append('slide4 mixes an <a:tbl> graphicFrame, a <p:graphicFrame> chart and a text box rotated -20 deg.')
    traps.append('slide3 has a text box at y=-1.3000 in (off the top edge); slide4 chart sits on an embedded xlsx.')
    traps.append('Deck is not single-theme in practice: slide3/slide4 use layout6 -> master3 -> theme3; slide1 uses')
    traps.append('  layout1 -> master1 -> theme1; slide2 layout5 -> master2 -> theme2; slide5 layout7 -> master4.')
    traps.append('docProps/app.xml <Template> says "MWC 19-PPT模板-白底 16比9" (white background, 16:9).')
    traps.append('No themeOverride*.xml parts exist (checked); docProps/custom.xml holds opaque DRM-looking strings.')
    traps.append('Handout/notes masters exist and each pull their own theme, so "walk all *.xml and merge colours" over-counts.')
    for t in traps:
        if t.startswith('  '):
            W('   %s' % t.strip())
        else:
            W(' - %s' % t)
    W('')

    # python-pptx cross-check (optional)
    W('=== CROSS-CHECK (python-pptx, if importable) ===')
    try:
        import pptx
        prs = pptx.Presentation(pptx_path)
        W('  python-pptx %s : slide_width=%d EMU  slide_height=%d EMU  masters=%d  layouts=%d  slides=%d'
          % (pptx.__version__, prs.slide_width, prs.slide_height,
             len(prs.slide_masters), sum(len(m.slide_layouts) for m in prs.slide_masters),
             len(prs.slides)))
        W('  agrees with raw XML: %s'
          % (prs.slide_width == cx and prs.slide_height == cy))
    except Exception as e:
        W('  skipped (%s: %s)' % (type(e).__name__, e))
    W('')

    text = '\n'.join(out)
    if json_mode:
        print(json.dumps(tokens, indent=2, ensure_ascii=False))
    else:
        print(text)


if __name__ == '__main__':
    main()
`;
/* eslint-enable */

// ---------------------------------------------------------------- run
const tmp = mkdtempSync(path.join(tmpdir(), 'dump-template-'));
const pyFile = path.join(tmp, 'dump_pptx.py');
writeFileSync(pyFile, PY_SOURCE, { encoding: 'utf8' });

const pyArgs = ['-X', 'utf8', pyFile, pptx];
if (wantJson) pyArgs.push('--json');

const res = spawnSync(python, pyArgs, {
  stdio: 'inherit',
  env: { ...process.env, PYTHONIOENCODING: 'utf-8', PYTHONUTF8: '1' },
});

try { rmSync(tmp, { recursive: true, force: true }); } catch (_) { /* ignore */ }

if (res.error) {
  console.error('[dump-template] failed to run python: ' + res.error.message);
  process.exit(4);
}
process.exit(res.status === null ? 5 : res.status);
