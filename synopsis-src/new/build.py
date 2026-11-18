#!/usr/bin/env python3
"""Assemble the condensed synopsis HTML (24-page build)."""
import pathlib

BASE = pathlib.Path('/home/user/SRS-Ambiguity-Detector/synopsis-src')

HEAD = '''<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<title>SRS Ambiguity Detector — Synopsis</title>
<style>
@font-face { font-family:'Tinos'; src:url('fonts/tinos-latin-400-normal.woff2') format('woff2'); font-weight:400; font-style:normal; }
@font-face { font-family:'Tinos'; src:url('fonts/tinos-latin-400-italic.woff2') format('woff2'); font-weight:400; font-style:italic; }
@font-face { font-family:'Tinos'; src:url('fonts/tinos-latin-700-normal.woff2') format('woff2'); font-weight:700; font-style:normal; }
@font-face { font-family:'Tinos'; src:url('fonts/tinos-latin-700-italic.woff2') format('woff2'); font-weight:700; font-style:italic; }
@page { size: A4; margin: 46px 54px 66px 54px; }
* { box-sizing: border-box; }
html, body { margin:0; padding:0; }
body { font-family:'Tinos','Times New Roman',serif; font-size:10.6pt; line-height:1.38; color:#111; }
p { text-align:justify; margin:0 0 6px 0; }
h1.ch { font-size:15pt; font-weight:700; color:#10355f; text-transform:uppercase; letter-spacing:.4px; border-bottom:2.2px solid #10355f; padding-bottom:6px; margin:0 0 12px 0; page-break-before:always; page-break-after:avoid; }
h1.flow { page-break-before:avoid; margin-top:18px; }
h2 { font-size:11.5pt; font-weight:700; color:#000; margin:11px 0 4px 0; page-break-after:avoid; }
h3 { font-size:11pt; font-weight:700; color:#10355f; margin:9px 0 3px 0; page-break-after:avoid; }
ul.b { list-style:none; margin:2px 0 8px 0; padding-left:20px; }
ul.b li { margin-bottom:3px; text-align:justify; }
ul.b li::before { content:"\\25CF"; font-size:6.5pt; color:#222; display:inline-block; width:17px; margin-left:-17px; vertical-align:2px; }
ol.n { margin:2px 0 8px 0; padding-left:24px; }
ol.n li { margin-bottom:3px; text-align:justify; }
table { border-collapse:collapse; width:100%; margin:7px 0 10px 0; font-size:9.2pt; line-height:1.3; }
th, td { border:1px solid #555; padding:3px 5px; vertical-align:top; text-align:left; }
th { background:#d9e2ef; font-weight:700; }
thead { display:table-header-group; }
tr { page-break-inside:avoid; }
.mono { font-family:'DejaVu Sans Mono',monospace; font-size:8.6pt; }
.frontpage { page-break-after:always; }
.fronttitle { text-align:center; font-size:15pt; font-weight:700; color:#10355f; letter-spacing:1.2px; margin:6px 0 22px 0; }
.center { text-align:center; }
.figure { text-align:center; margin:7px 0 10px 0; page-break-inside:avoid; }
.figure img { max-width:100%; border:1px solid #888; }
.caption { font-size:9.8pt; font-style:italic; margin-top:5px; text-align:center; }
.sigrow { display:flex; justify-content:space-between; margin-top:52px; }
.sigbox { width:44%; }
.sigline { border-top:1.4px solid #111; margin-top:32px; padding-top:5px; }
.tocline { margin:0 0 7px 0; font-size:12pt; }
.eqbox { background:#f2f5fa; border:1px solid #b9c6da; padding:7px 11px; margin:7px 0 10px 0; font-size:10pt; text-align:justify; }
</style>
</head>
<body>
'''

front = (BASE / 'new' / 'front.html').read_text(encoding='utf-8')
parts = [(BASE / 'new' / f'ch-{x}.html').read_text(encoding='utf-8') for x in 'abcd']
html = HEAD + front + ''.join(parts)

for name in ['fig71', 'fig72', 'fig73', 'fig74', 'fig75', 'fig76', 'fig77']:
    svg = pathlib.Path(f'/tmp/svgs/{name}.svg').read_text(encoding='utf-8')
    token = '%%' + name.upper() + '%%'
    assert token in html or name == "fig74", token
    html = html.replace(token, svg)

# Compact figure sizing for the condensed edition
for old, new in [('max-width:600px', 'max-width:415px'),
                 ('max-width:590px', 'max-width:385px'),
                 ('max-width:580px', 'max-width:414px'),
                 ('max-width:470px', 'max-width:308px'),
                 ('style="width:84%;"', 'style="width:60%;"'),
                 ('style="width:36%;"', 'style="width:27%;"')]:
    html = html.replace(old, new)

out = BASE / 'synopsis.html'
out.write_text(html, encoding='utf-8')
print('assembled:', out, len(html), 'bytes')
