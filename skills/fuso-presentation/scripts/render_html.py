#!/usr/bin/env python3
"""Render a FuSo deck (deck.json, see fuso_deck.py) as an online presentation.

  python render_html.py deck.json                 # writes <deck folder>/html/index.html and html/assets/
  python render_html.py deck.json --out DIR
  python render_html.py deck.json --shots DIR     # also screenshots for the check: one PNG per slide,
                                                  # a contact sheet and the phone view

In a window of 1000px or wider it shows one 16:9 slide, centred. In a narrower window (phone) it still shows one
slide at a time, filling the window with the content in one column; a slide taller than the window scrolls
within itself. Printing gives one slide per page (PDF).
Keys: arrows or space to move, O overview, F full screen, D dark and light, N speaker notes.
"""
import argparse, html, re, shutil, sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import fuso_deck as fd

esc = html.escape

CSS = r"""
*{box-sizing:border-box}
html,body{margin:0;height:100%}
/* One unit (U) is 1% of a 16:9 slide. Type scale, fixed for every deck (pt = PowerPoint at 13.33in):
   headline 4U (38pt, bold), card heading 1.9U (18pt), list text 1.3U (12.5pt), body 1.15U (11pt), small .95U (9pt),
   source .8U (7.5pt). Headlines are bold, not black: black is kept for the title slide. */
:root{--ctl:60px;--bar:var(--ctl);--u:min(1vw,calc((100vh - var(--bar))/56.25));--u:min(1vw,calc((100dvh - var(--bar))/56.25));
  --s-h1:6U;--s-hs:3.6U;--s-h2:3.1U;--s-lead:1.35U;--s-h3:1.6U;--s-text:1.3U;--s-body:1.15U;--s-small:.95U;--s-tiny:.8U;
  --s-agenda:1.3U;--s-agenda2:.9U}
body{font-family:InterLocal,Inter,system-ui,sans-serif;-webkit-font-smoothing:antialiased;overflow:hidden;
  background:var(--k-stage);color:var(--c-fg)}
h1,h2,h3,p,ul,ol,dl,dd,figure{margin:0;padding:0}
a{color:inherit}
a:focus-visible,button:focus-visible{outline:2px solid currentColor;outline-offset:2px}

.slide{position:absolute;left:50%;top:calc((100% - var(--bar))/2);width:100U;height:56.25U;transform:translate(-50%,-50%);display:grid;grid-template-rows:minmax(0,1fr) auto;justify-items:center;align-items:stretch;
  overflow:hidden;visibility:hidden;background:var(--c-bg);color:var(--c-fg)}
.slide.active{visibility:visible}
.frame{width:100%;height:100%;padding:3.6U 5U 1.2U;display:flex;flex-direction:column;min-width:0;position:relative;z-index:1}
.wm{position:absolute;right:0;bottom:4U;width:50U;height:50U;fill:none;stroke:var(--c-wm);stroke-width:.3;pointer-events:none}
.wm.tr{top:0;bottom:auto;width:34U;height:34U;stroke-width:.45}

h1,h2,.hs{font-weight:700;letter-spacing:-.05em;line-height:1.04;text-wrap:balance;color:var(--c-fg)}
.bg-green h2,.bg-green .hs{line-height:1.22}
h1{font-size:var(--s-h1);font-weight:900;letter-spacing:-.035em;line-height:1.04}
/* a picture behind the end of the first title line, as the paper bag on the Brown Bag Lunch Talk graphics */
h1{position:relative;z-index:0}
.kicker{align-self:flex-start;width:fit-content;font-size:var(--s-text);font-weight:700;line-height:1.2;color:var(--c-inv-fg);
  background:var(--c-inv);padding:.5U 1.5U;border-radius:2U;margin-bottom:2U}
.tmark{position:relative;display:inline-block;width:0;height:0}
.tmark img{position:absolute;left:-.3em;bottom:-.3em;height:1.05em;width:auto;z-index:-1}
.hs{font-size:var(--s-hs)}
h2{font-size:var(--s-h2);max-width:76U}
h3{font-size:var(--s-h3);font-weight:700;line-height:1.14;letter-spacing:-.025em;color:var(--c-fg)}
p,li,dd,dt{font-size:var(--s-body);line-height:1.42;color:var(--c-body)}
.hl{color:var(--c-hl);background:linear-gradient(var(--c-hlbg),var(--c-hlbg)) no-repeat 0 100%/100% 90%;
  padding:0 var(--hl-pad);box-decoration-break:clone;-webkit-box-decoration-break:clone}
.questions,.nums{list-style:none;counter-reset:q}
.questions{display:flex;flex-wrap:wrap;gap:.5U 2.6U;margin-bottom:1.5U}
.questions li{font-size:var(--s-text);font-weight:700;color:var(--c-fg)}
.questions li,.nums li{counter-increment:q;position:relative}
.questions li::before,.nums li::before{content:counter(q);color:var(--c-hl);font-weight:700;margin-right:.6em}
.nums{display:grid;gap:.55U}
.nums li{padding-left:1.4em}
.nums li::before{position:absolute;left:0;margin:0}
.aside-in .tag{font-size:var(--s-small);font-weight:700}
.headrow{display:flex;justify-content:space-between;align-items:flex-end;gap:4U}
.callout{flex:none;max-width:36U;display:flex;align-items:center;gap:1.3U;padding:1.2U 1.8U 1.2U 1.3U}
.callout .disc{width:3.8U;height:3.8U}
.callout strong{display:block;font-size:1.4U;line-height:1.2;letter-spacing:-.02em;color:var(--c-panel-fg)}
.callout span{display:block;font-size:var(--s-body);line-height:1.35;color:var(--c-panel-body);margin-top:.2U}
.lead{font-size:var(--s-lead);line-height:1.3;margin-top:1.1U;max-width:76U}
/* the content is centred in the space between headline and footer, a little above the middle */
.main{flex:1;min-height:0;display:flex;flex-direction:column;justify-content:safe center;padding:2.4U 0 3.4U}
.source{font-size:var(--s-tiny);color:var(--c-quiet);margin-top:.8U}
.center{flex:1;min-width:0;display:flex;flex-direction:column;justify-content:safe center}
/* the closing slide mirrors the title slide: the green panel with the logo, now on the left */
.l-closing{grid-template-rows:minmax(0,1fr)}
.l-closing .foot,.l-closing .wm{display:none}
.l-closing .frame{width:100%;padding:0;display:grid;grid-template-columns:36% minmax(0,1fr)}
.l-closing .tmain{flex-direction:row;align-items:center;gap:4U}
html[data-mode=dark] .only-l,html[data-mode=light] .only-d{display:none!important}

/* circuit line, as on futuresociety.ch and the LinkedIn figures: a thick line ending in a filled dot */
.circuit{position:relative;width:9U;height:1.65U;margin:2.6U 0 2.4U;flex:none}
.circuit::before{content:"";position:absolute;left:0;top:.55U;width:7.6U;height:.55U;background:var(--c-line)}
.circuit::after{content:"";position:absolute;right:0;top:0;width:1.65U;height:1.65U;border-radius:50%;background:var(--c-line)}

/* title slide: logos, title and links on the left, the FuSo mark on a green field on the right */
.l-title{grid-template-rows:minmax(0,1fr);--s-lead:1.9U;--s-text:1.7U;--s-small:1.3U}
.l-title .foot{display:none}
.l-title .frame{width:100%;padding:0;display:grid;grid-template-columns:minmax(0,1fr) 36%}
.tmain{padding:3.6U 5U;display:flex;flex-direction:column;min-width:0}
.toprow{display:flex;flex-wrap:wrap;align-items:center;gap:1.2U 2.2U}
.toprow>*+*{padding-left:2.2U;border-left:1px solid var(--c-divider)}
.toprow img{height:3.3U;width:auto;display:block}
.wordmark{line-height:1}
.wordmark b{display:block;font-size:1.75U;font-weight:900;letter-spacing:-.03em;color:var(--c-fg)}
.wordmark b i{font-style:normal;color:var(--c-hl)}
.wordmark span{display:block;font-size:1.05U;color:var(--c-quiet);margin-top:.45U}
.tart{position:relative;overflow:hidden;background:#00802F;display:grid;place-items:center}
.tart svg{position:absolute;left:50%;top:50%;width:210%;aspect-ratio:1;transform:translate(-50%,-50%);fill:none;stroke:#fff;stroke-width:.22;opacity:.3}
.tart img{position:relative;width:58%;height:auto;display:block}
.sub{font-size:var(--s-lead);line-height:1.35;max-width:62U}
.presenters{font-size:var(--s-text);font-weight:700;color:var(--c-fg)}
.sub+.presenters{margin-top:1.6U}
.meta{font-size:var(--s-small);color:var(--c-quiet);margin-top:.5U}
.attr{font-size:var(--s-body);color:var(--c-quiet);margin-top:2U}
.tbottom{display:flex;align-items:center;gap:3U}
.qr{flex:none;display:flex;flex-direction:column;gap:1U;text-decoration:none;color:var(--c-fg);font-size:var(--s-small);font-weight:700}
.qr svg{width:15U;height:15U;background:#fff;border-radius:.6U;padding:.8U;display:block}
.tbottom .qr{flex-direction:row;align-items:center;gap:1.3U}
.tbottom .qr svg{width:8.6U;height:8.6U;padding:.45U}
.banner{display:block;flex:none}
.banner img{height:8.6U;width:auto;display:block;border-radius:.6U}
.contact{list-style:none;display:grid;grid-template-columns:auto auto;justify-content:start;gap:1U 3.2U;margin-top:2.4U}
.contact a,.contact li>span{display:inline-flex;align-items:center;gap:.8U}
.contact svg{width:1.7U;height:1.7U;fill:var(--c-hl);flex:none}
.contact li{font-weight:700;color:var(--c-fg)}

.panel{background:var(--c-panel);color:var(--c-panel-fg);border:1px solid var(--c-panel-line);border-radius:.6U;padding:1.8U}
.panel h3{color:var(--c-panel-fg)}
.panel p,.panel li,.panel dd{color:var(--c-panel-body)}
.panel.inv{background:var(--c-inv);border-color:var(--c-inv-line)}
.panel.inv h3,.panel.inv p,.panel.inv li,.panel.inv .num{color:var(--c-inv-fg)}
.panel.inv .disc{background:#fff;color:#00802F}
ul.dots{list-style:none;display:grid;gap:.75U}
ul.dots li{position:relative;padding-left:1.4em}
ul.dots li::before{content:"";position:absolute;left:0;top:.45em;width:.5em;height:.5em;border-radius:50%;background:var(--c-hl)}
.panel.inv ul.dots li::before{background:var(--c-inv-fg)}
.disc{width:3.6U;height:3.6U;border-radius:50%;background:var(--c-disc);color:var(--c-icon);display:grid;place-items:center;flex:none}
.disc svg{width:50%;height:50%;fill:currentColor;display:block}

/* text: copy on the left, a picture or panel on the right that reaches the bottom of the zone */
.split{display:grid;grid-template-columns:1.12fr 1fr;gap:4U;align-items:center}
.split.flip{grid-template-columns:1fr 1.12fr}
.split.flip .aside{order:-1}
.copy{display:grid;gap:1.6U;align-content:start;max-width:66U}
.copy .contact{grid-template-columns:auto;margin-top:0}
.split .copy:has(.contact){align-self:end}   /* links end on the bottom edge of the panel beside them */
.copy p,.copy li{font-size:var(--s-text)}
/* points start with the same green dot as every list; a few short ones are set as large statements */
.rows{list-style:none;display:grid;gap:1U;margin-top:.6U}
.rows li{position:relative;padding-left:1.8U;font-size:var(--s-body)}
.rows li::before{content:"";position:absolute;left:0;top:.45em;width:.6U;height:.6U;border-radius:50%;background:var(--c-hl)}
.rows.big{gap:2U}
.rows.big li{padding-left:2.7U;font-size:1.9U;font-weight:700;letter-spacing:-.03em;line-height:1.25;color:var(--c-fg)}
.rows.big li::before{top:.34em;width:1.1U;height:1.1U}
.split .copy p{font-size:var(--s-lead)}
.copy strong{font-weight:700;color:var(--c-fg)}
.tl{color:var(--c-hl);font-weight:700;text-underline-offset:.2em}
/* points with an icon: one block per point, the icon in a disc on the left */
.prows{display:grid;gap:1U;margin-top:.8U}
.prow{display:flex;align-items:center;gap:1.5U;padding:1.2U 1.8U 1.2U 1.3U}
.prow .disc{width:3.8U;height:3.8U}
.prow p{font-size:var(--s-body);line-height:1.4}
.split .copy{gap:1.4U}
.aside{min-height:0}
.aside.pic{display:flex;flex-direction:column}
.aside.pic img{width:100%;max-height:32U;object-fit:cover;border-radius:.6U;display:block}
.aside.pic figcaption,.bigpic figcaption{font-size:var(--s-tiny);color:var(--c-quiet);margin-top:.8U}
.aside.ico{display:grid;place-items:center}
.disc.big{width:20U;height:20U}
.aside.panel{padding:0;overflow:hidden;display:flex;flex-direction:column}
.aimg{position:relative;flex:1 1 9U;min-height:9U}
.aimg img{position:absolute;inset:0;width:100%;height:100%;object-fit:cover;display:block}
.badge{position:absolute;left:1.4U;top:1.4U;display:flex;align-items:center;gap:.6U;background:#fff;color:#111;font-size:1.05U;font-weight:700;padding:.45U 1.2U .45U .55U;border-radius:2U}
.aimg .badge img{position:static;width:1.9U;height:1.9U}
.aside-in{padding:2U 2.2U;display:flex;flex-direction:column;gap:.9U}
.aside.panel:not(.has-img) .aside-in{flex:1}
.aside-in h3{font-size:1.9U;letter-spacing:-.035em}
.aside-in .atitle{font-size:var(--s-h2);letter-spacing:-.05em;line-height:1.08}
.ahead{display:flex;justify-content:space-between;align-items:flex-start;gap:2U}
.ahead-t{display:grid;gap:.9U;min-width:0}
.ahead .face{width:8U;height:8U}
.aside-in>p{font-size:var(--s-body)}
.aside-in .alabel{font-size:var(--s-text);font-weight:700;color:var(--c-panel-fg);margin-top:-.4U}
.aside-in .anote{font-size:var(--s-small);color:var(--c-panel-quiet)}
.facts{display:grid;grid-template-columns:auto 1fr;gap:.7U 2U;margin-top:.5U}
.facts dt{font-size:var(--s-text);font-weight:700;color:var(--c-panel-fg);line-height:1.3}
.facts dd{font-size:var(--s-body);color:var(--c-panel-body);line-height:1.35;align-self:center}
.aside-logos{display:flex;flex-wrap:wrap;align-items:center;gap:1.6U;margin-top:.4U}
.aside-logos img{height:2.6U;width:auto;display:block}
/* the call to action is a band of its own at the foot of the panel */
.qrrow{display:flex;gap:1.6U;align-items:center;text-decoration:none;padding:1.5U 2.2U;background:rgba(0,0,0,.18);color:var(--c-panel-fg)}
.qrrow svg{width:6.4U;height:6.4U;background:#fff;border-radius:.4U;padding:.4U;flex:none;display:block}
.qrtext{display:grid;gap:.25U}
.qrtext strong{font-size:1.5U;font-weight:700;letter-spacing:-.02em;line-height:1.2}
.qrtext span{font-size:var(--s-body);line-height:1.35}
.qrtext b{font-size:var(--s-text);font-weight:700;text-decoration:underline;text-underline-offset:.2em;margin-top:.3U}
.panel.inv .alabel,.panel.inv .facts dt,.panel.inv .qrrow{color:var(--c-inv-fg)}
.panel.inv .facts dd,.panel.inv .anote,.panel.inv .qrtext span{color:rgba(255,255,255,.86)}
.panel.inv .hl{color:#fff;font-weight:700}
/* a panel with a picture takes the full height of the zone; the copy is centred beside it */
.split:has(.has-img){flex:1;min-height:0;align-items:stretch}
.split:has(.has-img) .copy{align-self:center}
.main:has(.split .has-img){padding-bottom:1U}

/* cards, as in the AI+X deck: a row of equal cards right under the headline. Each card carries a statement
   and a caption at its foot; with an icon, the icon sits at the top and the text at the foot. */
.cols,.people,.steps{display:grid;gap:1.6U;grid-template-columns:repeat(var(--n),minmax(0,1fr))}
.col,.step{display:flex;flex-direction:column;gap:1.2U;padding:2.2U}
.chead{display:flex;align-items:flex-start;gap:1.4U}
/* a card that is only an icon and a word: both larger, centred on each other */
.chead:not(:has(p)){align-items:center}
.step .chead:not(:has(p)) h3{font-size:2.3U}
.chead:not(:has(p)) .ctext{padding-top:0}
.ctext{display:grid;gap:.8U;min-width:0;padding-top:.2U}
.col p,.step p{font-size:var(--s-body)}
.band{margin-top:1.6U;background:var(--c-band);color:var(--c-band-fg);border-radius:.6U;padding:1.2U 1.8U;display:flex;gap:2U;
  align-items:baseline;font-size:var(--s-text);line-height:1.35}
.band strong{font-weight:700;flex:none}

/* steps: the same cards, wired to each other: a line across the gap that ends in a dot on the next card */
.steps{list-style:none;gap:3U}
.steps li{position:relative;display:flex}
.steps .step{flex:1}
.steps li+li::before{content:"";position:absolute;left:-3U;width:3U;top:calc(50% - .2U);height:.4U;background:var(--c-line)}
.steps li+li::after{content:"";position:absolute;left:-.65U;top:calc(50% - .65U);width:1.3U;height:1.3U;border-radius:50%;background:var(--c-line)}
.num{font-size:var(--s-small);font-weight:700;color:var(--c-panel-quiet)}

.cmp{display:grid;grid-template-columns:1fr 1fr;gap:1.6U}
.side{display:grid;gap:1.2U;align-content:start;padding:2.2U;min-height:15U}
.side p,.side li{font-size:var(--s-text)}
.side-head{display:flex;align-items:center;gap:1.4U}
.side-head .disc{width:4U;height:4U}
.side ul.dots{gap:1.1U}
.conclusion{font-size:1.5U;font-weight:700;color:var(--c-fg);letter-spacing:-.015em;line-height:1.3;margin-top:2.2U}
.copy .after{font-size:var(--s-body);color:var(--c-quiet);margin-top:1.2U}
/* fuse: two fields as overlapping circles; the overlap is green and carries the FuSo logo */
.fuse{display:grid;grid-template-columns:minmax(0,1fr) auto minmax(0,1fr);gap:3.4U;align-items:center}
.fz{display:grid;gap:.9U}.fz h3{font-size:1.9U}.fz p{font-size:var(--s-text)}
.fz.l{text-align:right}
.venn{position:relative;width:28.5U;height:18U}
.vc,.vx{position:absolute;top:0;left:0;width:18U;height:18U;border-radius:50%}
.vc{background:var(--c-panel);border:.25U solid var(--c-inv);color:var(--c-inv)}
.vc.r{left:10.5U}
.vc svg{position:absolute;top:6.75U;left:3U;width:4.5U;height:4.5U;fill:currentColor}
.vc.r svg{left:auto;right:3U}
.vx{overflow:hidden}
.vx::before{content:"";position:absolute;top:0;left:10.5U;width:18U;height:18U;border-radius:50%;background:var(--c-inv)}
.venn img{position:absolute;top:6.4U;left:11.65U;width:5.2U;height:5.2U}
.l-fuse .conclusion{text-align:center;width:fit-content;max-width:58U;margin:2.6U auto 0;padding:1.1U 2U;border-radius:.6U;font-size:var(--s-text);background:var(--c-panel);color:var(--c-panel-fg);text-wrap:balance}

.agenda{list-style:none;display:grid;gap:1.2U}
.agenda>li{display:grid;grid-template-columns:var(--when,6.5U) 1fr;gap:1.6U;align-items:baseline;
  font-size:var(--s-agenda)}
.when{font-weight:700;color:var(--c-hl);font-variant-numeric:tabular-nums}
.agenda strong{display:block;font-weight:700;color:var(--c-fg);line-height:1.25}
.who{display:block;font-size:var(--s-agenda2);color:var(--c-quiet);line-height:1.3;margin-top:.15U}
.subs{list-style:none;display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:.7U;margin-top:.7U}
.subs li{font-size:var(--s-agenda2);font-weight:700;line-height:1.25;color:var(--c-panel-fg);background:var(--c-panel);
  border:1px solid var(--c-panel-line);border-radius:.5U;padding:.7U 1U}
.subs span{display:block;font-size:var(--s-tiny);font-weight:400;color:var(--c-panel-quiet);margin-top:.15U}

.people{gap:1.6U}
.person{display:grid;gap:1.3U;align-content:start}
.phead{display:flex;align-items:center;gap:1.2U}
.who{display:grid;gap:.35U;min-width:0}
.l-people:not(.speaker) .lead{font-size:var(--s-text)}
.face{width:5U;height:5U;border-radius:50%;object-fit:cover;display:block;flex:none;background:var(--c-disc)}
.name{font-size:1.3U;font-weight:700;color:var(--c-fg);line-height:1.15;letter-spacing:-.015em}
.role{font-size:var(--s-tiny);color:var(--c-quiet);line-height:1.35}
.person li{font-size:.85U;line-height:1.38}
.ptag{font-size:var(--s-tiny);font-weight:700;color:var(--c-hl)}
.people.many{gap:1.1U}
.people.many .person{padding:1.1U 1.3U}
.people.many .face{width:4.6U;height:4.6U}
.people.n2 .face,.people.n3 .face{width:7U;height:7U}
.people.n2 .name,.people.n3 .name{font-size:1.7U}
.people.n2 .role,.people.n3 .role{font-size:var(--s-body)}
/* one person: label, title, circuit line, name and role form one group in the middle of the slide */
.speaker .frame{justify-content:safe center}
.speaker h2{font-size:var(--s-hs);max-width:70U}
.speaker .lead{order:-1;font-size:1.6U;font-weight:700;margin:0 0 1.4U}
.speaker .main{flex:none;padding:0}
.speaker .source{position:absolute;left:5U;bottom:1.2U}
.people.n1 .circuit{margin:2.8U 0 2.4U}
.people.n1 .phead{gap:2.4U}
.people.n1 .face{width:10U;height:10U}
.people.n1 .name{font-size:2.5U;font-weight:700;letter-spacing:-.04em}
.people.n1 .role{font-size:var(--s-text);max-width:60U;margin-top:.3U}
.people.n1 li{font-size:var(--s-body)}

.flow{display:grid;grid-template-columns:auto minmax(0,1fr) auto;gap:3.2U;align-items:stretch}
.fend{position:relative;background:var(--c-inv);color:var(--c-inv-fg);border-radius:.6U;display:grid;place-items:center;padding:0 3.2U;
  font-size:2.4U;font-weight:700;letter-spacing:-.035em}
.chips{position:relative;list-style:none;display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:1.1U}
.chips li{background:var(--c-panel);color:var(--c-panel-fg);border-radius:.6U;padding:1.5U 1.6U;min-height:8U;display:flex;align-items:center;
  font-size:1.45U;font-weight:700;letter-spacing:-.02em;line-height:1.2}
.chips::before,.fend:last-child::before{content:"";position:absolute;left:-3.2U;top:calc(50% - .2U);width:3.2U;height:.4U;background:var(--c-line)}
.chips::after,.fend:last-child::after{content:"";position:absolute;left:-.65U;top:calc(50% - .65U);width:1.3U;height:1.3U;border-radius:50%;background:var(--c-line)}
.matrix{display:grid;grid-template-columns:minmax(0,1fr) repeat(var(--n),11.5U);row-gap:.55U}
.mx-h{align-self:end;text-align:center;padding:0 .5U .7U;font-size:var(--s-small);font-weight:700;line-height:1.2;color:var(--c-fg)}
.mx-l,.mx-c{background:var(--c-panel);display:flex;align-items:center;padding:.85U 1.6U}
.mx-l{border-radius:.6U 0 0 .6U;font-size:1.3U;font-weight:700;letter-spacing:-.015em;color:var(--c-panel-fg)}
.mx-c{justify-content:center}
.mx-c.last{border-radius:0 .6U .6U 0}
.mx-c i{width:1.3U;height:1.3U;border-radius:50%;background:var(--c-hl);display:block}
.lwall{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:1.4U}
.ltile{background:var(--c-panel);border-radius:.6U;height:10.5U;display:grid;place-items:center;padding:2U 2.4U}
.ltile img{width:100%;height:6.5U;object-fit:contain;display:block}
html[data-mode=dark] .ltile.lt-w{background:#fff}
.link{display:flex;flex-direction:column;gap:.7U;text-decoration:none;min-height:17U}
.link .src{font-size:var(--s-small);font-weight:700;color:var(--c-panel-fg);line-height:1.3}
.link .date{font-size:var(--s-small);color:var(--c-panel-quiet);line-height:1.3;margin-top:-.5U}
.link h3{margin-top:1.4U}
.link p{font-size:var(--s-body)}
.link .go{margin-top:auto;padding-top:1.4U;display:flex;align-items:center;justify-content:space-between;gap:1U}
.link .host{font-size:var(--s-small);color:var(--c-panel-fg);text-decoration:underline;text-underline-offset:.2em}
.link .disc{width:3U;height:3U}

.bigpic{flex:1;min-height:0;display:flex;flex-direction:column}
.bigpic img{flex:1;min-height:0;width:100%;object-fit:contain;object-position:left center;display:block}

.foot{justify-self:stretch;height:4U;background:var(--c-foot);color:var(--c-foot-fg);display:flex;justify-content:center}
.foot-in{width:100%;padding:0 5U;display:flex;align-items:center;justify-content:space-between;gap:2U}
.hub{display:flex;align-items:center;gap:.9U;text-decoration:none;font-weight:700;font-size:1.25U}
.hub img{width:2.5U;height:2.5U;display:block}
.pn{font-size:1.1U;font-variant-numeric:tabular-nums;height:2.6U;aspect-ratio:.9;box-sizing:border-box;padding:.11em 0 0 .25em;display:grid;place-items:center;color:#fff;background:url(assets/HSG_Impact_Space_Flache_RGB.png) center/contain no-repeat}

/* the controls: a pill that floats under the slide. Previous and next are green discs like the icon discs on the slides,
   every other button carries its label. In full screen the pill hides and comes back when the pointer moves. */
.controls{position:fixed;left:0;right:0;bottom:0;height:var(--ctl);display:grid;place-items:center;z-index:7;pointer-events:none;
  color:var(--k-fg);font-size:13px;font-weight:600}
html[data-mode=light]{--k-stage:#e4e4e4;--k-bg:#ffffff;--k-fg:#111111;--k-line:#dddddd;--k-hover:#f0f0f0;--k-quiet:#666666}
html[data-mode=dark]{--k-stage:#000000;--k-bg:#1a1a1a;--k-fg:#ededed;--k-line:#2e2e2e;--k-hover:#2e2e2e;--k-quiet:#bdbdbd}
.pill{pointer-events:auto;display:flex;align-items:center;height:44px;padding:0 6px;border-radius:22px;background:var(--k-bg);
  border:1px solid var(--k-line);box-shadow:0 4px 18px rgba(0,0,0,.1);transition:opacity .25s,transform .25s}
.pill>div{display:flex;align-items:center;gap:2px}
.pill .sep{width:1px;height:20px;background:var(--k-line);margin:0 8px}
.controls button{border:0;background:transparent;color:inherit;font:inherit;height:34px;padding:0 12px;border-radius:17px;cursor:pointer;
  display:flex;align-items:center;justify-content:center;gap:7px;white-space:nowrap}
.controls [hidden]{display:none}
.controls svg{width:18px;height:18px;fill:none;stroke:currentColor;stroke-width:1.75;stroke-linecap:round;stroke-linejoin:round;flex:none}
.controls button:hover,body.show-notes #notesBtn,body:has(.overview.open) #ovBtn{background:var(--k-hover)}
.controls button:focus-visible{outline:2px solid #00802F;outline-offset:2px}
.controls .nav button{width:32px;height:32px;padding:0;background:#00802F;color:#fff}
.controls .nav button:not(:disabled):hover{background:#006a27}
.controls .nav button:disabled{opacity:.3;cursor:default}
.nav svg{stroke-width:2.25}
.count{min-width:60px;text-align:center;font-variant-numeric:tabular-nums}
html[data-mode=light] .to-light,html[data-mode=dark] .to-dark,:root:not(:fullscreen) .fs-on,:root:fullscreen .fs-off{display:none}
.progress{position:fixed;left:0;right:0;bottom:0;height:3px;background:var(--k-line)}
.progress span{display:block;height:100%;width:0;background:#00802F;transition:width .3s}
:root:fullscreen{--bar:0px}
:root:fullscreen .progress{bottom:0;background:transparent}
:root:fullscreen .pill{opacity:0;transform:translateY(10px);pointer-events:none}
:root:fullscreen body.awake .pill,:root:fullscreen .pill:has(:focus-visible){opacity:1;transform:none;pointer-events:auto}
:root:fullscreen body:not(.awake){cursor:none}

.overview{position:fixed;inset:0;z-index:10;display:none;overflow:auto;padding:5U 5U 8U;background:var(--c-bg);color:var(--c-fg)}
.overview.open{display:block}
.ov-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(20U,1fr));gap:1.4U;margin-top:3U}
.ov-grid button{all:unset;cursor:pointer;box-sizing:border-box;min-height:10U;padding:1.6U;border-radius:.6U;background:var(--c-panel);
  border:1px solid var(--c-panel-line);color:var(--c-panel-fg);font-size:var(--s-small);font-weight:700;line-height:1.25;
  display:flex;flex-direction:column;gap:.6U}
.ov-grid button:focus-visible{outline:2px solid currentColor}
.ov-grid small{font-size:var(--s-tiny);font-weight:400;color:var(--c-panel-quiet)}
.notes-view{position:fixed;left:0;right:0;bottom:var(--bar);z-index:6;display:none;max-height:42%;overflow:auto;padding:2U 5U;
  background:var(--c-panel);color:var(--c-panel-fg);border-top:1px solid var(--c-divider);font-size:var(--s-body);line-height:1.45;white-space:pre-wrap}
body.show-notes .notes-view{display:block}
.sr{position:absolute;width:1px;height:1px;overflow:hidden;clip-path:inset(50%)}

/* windows narrower than 1000px: still one slide at a time, but the slide fills the window and its content
   reflows into one column; a slide taller than the window scrolls within itself */
@media (max-width:999px),(max-height:480px){
  :root{--u:8px;--s-h1:38px;--s-hs:30px;--s-h2:25px;--s-lead:17px;--s-h3:18px;--s-text:16px;--s-body:16px;--s-small:14px;--s-tiny:12.5px;
    --s-agenda:16px;--s-agenda2:14px}
  body{background:var(--c-bg)}
  .slide{left:0;top:0;width:100%;height:calc(100% - var(--bar));transform:none;display:flex;flex-direction:column;
    justify-items:normal;overflow-y:auto}
  .frame{width:auto;height:auto;flex:1 0 auto;display:block;padding:36px 22px 40px}
  .l-closing .frame{display:block}
  .contact{grid-template-columns:1fr;gap:12px}.contact svg{width:22px;height:22px}
  .l-title .frame{display:block;width:auto}
  .tmain{padding:28px 22px 36px;display:block}
  .toprow{margin-bottom:40px}.toprow .wordmark{flex-basis:100%;padding-left:0;border-left:0}.toprow img{height:26px}.wordmark b{font-size:15px}.wordmark span{font-size:10px}
  .tbottom{display:block;margin-top:34px}
  .tbottom .qr svg,.qr svg{width:120px;height:120px;padding:8px}
  .banner{margin-top:22px}.banner img{height:auto;width:100%;max-width:320px}
  .tart{height:190px}.tart img{width:auto;height:120px}.tart svg{width:560px}
  h2{max-width:none}
  .flow{grid-template-columns:1fr;gap:26px}.fend{padding:16px}.chips{grid-template-columns:1fr 1fr}
  .chips::before,.fend:last-child::before{left:26px;top:-26px;width:4px;height:26px}
  .chips::after,.fend:last-child::after{left:21px;top:-7px;width:14px;height:14px}
  .chips li{min-height:0;font-size:15px}
  .matrix{grid-template-columns:minmax(0,1fr) repeat(var(--n),34px)}.mx-h{writing-mode:vertical-rl;font-size:11px;padding:0 0 8px;justify-self:center}
  .mx-l{font-size:14px;padding:10px 12px}.mx-c{padding:0}.mx-c i{width:12px;height:12px}
  .lwall{grid-template-columns:1fr 1fr}.ltile{height:90px}.ltile img{height:50px}
  .people.many{grid-template-columns:1fr}
  .headrow{display:block}.callout{max-width:none;margin-top:18px}
  .main{display:block;padding-top:22px}
  .split,.split.flip,.cols,.cmp,.people,.steps,.subs{grid-template-columns:1fr}
  .fuse{grid-template-columns:1fr;gap:22px;justify-items:center}.fz,.fz.l{text-align:center}
  .split{gap:24px}.cols,.people{gap:14px}.people{gap:30px}
  .qr{margin-top:28px}
  .disc{width:52px;height:52px}.disc.big{width:96px;height:96px}
  .col,.step,.link,.side{min-height:0}
  .col h3:first-child,.step h3{font-size:20px}
  .aside.ico{place-items:start}
  .face{width:64px;height:64px}.people.n1 .face{width:96px;height:96px}
  .people.n1 .name{font-size:26px}.person li{font-size:15px}
  .steps{gap:26px}
  .steps li+li::before{left:26px;top:-26px;width:2px;height:26px}
  .steps li+li::after{left:21px;top:-6px;width:12px;height:12px}
  .agenda{gap:18px}.agenda>li{grid-template-columns:1fr;gap:2px}
  .band{flex-direction:column;gap:4px;margin-top:14px;padding:14px 18px}
  .bigpic img{flex:none;height:auto}
  .wm{display:none}
  .foot{flex:none;height:52px}
  .rows.big li{font-size:20px}
  .wm.tr{display:block;width:150px;height:150px}
  .aimg{flex:none;aspect-ratio:16/8}
  .speaker .lead{font-size:18px}
  .speaker .source{position:static}
  .foot-in{padding:0 22px}
  .hub{white-space:nowrap;font-size:14px}.pn{font-size:14px;height:34px}
  .hub img{width:26px;height:26px}
  :root{--ctl:60px}
  .pill{width:100%;height:100%;padding:3px 10px 0;border-radius:0;border:0;box-shadow:none;justify-content:space-between}
  .pill .sep{display:none}
  .tools button{flex-direction:column;gap:3px;height:48px;min-width:56px;padding:0 6px;border-radius:10px;font-size:10.5px}
  .tools svg{width:20px;height:20px}
  .controls .nav button{width:38px;height:38px}
  .progress{bottom:calc(var(--ctl) - 3px)}
}
@media print{
  @page{size:13.333in 7.5in;margin:0}
  :root{--u:.13333in}
  html,body{height:auto}
  body{overflow:visible}
  .slide{position:relative;left:auto;top:auto;transform:none;width:100%;visibility:visible;height:7.5in;break-after:page}
  :root{--bar:0px}
  .controls,.overview,.notes-view{display:none!important}
  *{-webkit-print-color-adjust:exact;print-color-adjust:exact}
}
@media (prefers-reduced-motion:no-preference){html{scroll-behavior:smooth}}
"""

JS = r"""
(()=>{const slides=[...document.querySelectorAll('.slide')],root=document.documentElement,body=document.body,
 ov=document.getElementById('overview'),nv=document.getElementById('notesView');
const by=id=>document.getElementById(id);
let cur=0;const m=location.hash.match(/(\d+)$/);if(m&&+m[1]>=1&&+m[1]<=slides.length)cur=+m[1]-1;
function go(n,scroll){cur=Math.max(0,Math.min(slides.length-1,n));
 slides.forEach((s,i)=>{s.classList.toggle('active',i===cur);
  s.setAttribute('aria-hidden',i!==cur)});
 slides[cur].scrollTop=0;
 by('count').textContent=(cur+1)+' / '+slides.length;by('prev').disabled=cur===0;by('next').disabled=cur===slides.length-1;by('progress').style.width=((cur+1)/slides.length*100)+'%';
 const t=slides[cur].querySelector('.notes');nv.textContent=t?t.textContent:'';
 history.replaceState(null,'','#/'+(cur+1));ov.classList.remove('open');
}
by('prev').onclick=()=>go(cur-1);by('next').onclick=()=>go(cur+1);
by('ovBtn').onclick=()=>ov.classList.toggle('open');
by('modeBtn').onclick=()=>{root.dataset.mode=root.dataset.mode==='dark'?'light':'dark'};
const notes=()=>body.classList.toggle('show-notes');if(by('notesBtn'))by('notesBtn').onclick=notes;
if(!document.fullscreenEnabled)by('fsBtn').hidden=true;
let idle;const pill=document.querySelector('.pill'),wake=()=>{body.classList.add('awake');clearTimeout(idle);
 idle=setTimeout(()=>{if(!pill.matches(':hover'))body.classList.remove('awake')},2500)};
['mousemove','touchstart'].forEach(t=>document.addEventListener(t,wake,{passive:true}));
by('fsBtn').onclick=()=>document.fullscreenElement?document.exitFullscreen():root.requestFullscreen();
const grid=ov.querySelector('.ov-grid');
slides.forEach((s,i)=>{const b=document.createElement('button'),sm=document.createElement('small');
 sm.textContent=(i+1)+(s.dataset.time?'  ·  '+s.dataset.time:'');b.append(sm,s.dataset.title);b.onclick=()=>go(i,true);grid.append(b)});
document.addEventListener('keydown',e=>{if(e.ctrlKey||e.metaKey||e.altKey)return;
 const inCtl=e.target instanceof Element&&e.target.closest('a,button');const k=e.key;
 if(['ArrowRight','PageDown'].includes(k)||(k===' '&&!inCtl)){e.preventDefault();go(cur+1)}
 else if(['ArrowLeft','PageUp'].includes(k)){e.preventDefault();go(cur-1)}
 else if(k==='Home')go(0);else if(k==='End')go(slides.length-1);
 else if(k==='o'||k==='O')ov.classList.toggle('open');else if(k==='Escape')ov.classList.remove('open');
 else if(k==='f'||k==='F')by('fsBtn').click();else if(k==='d'||k==='D')by('modeBtn').click();
 else if(k==='n'||k==='N')notes()});
let x0=null;document.addEventListener('touchstart',e=>{x0=e.touches[0].clientX},{passive:true});
document.addEventListener('touchend',e=>{if(x0===null)return;const d=e.changedTouches[0].clientX-x0;x0=null;
 if(Math.abs(d)>60)go(cur+(d<0?1:-1))},{passive:true});
window.fusoDeck={go,count:slides.length};go(cur,cur>0)})();
"""

CORNER = ('<svg class="wm tr" viewBox="0 0 100 100" aria-hidden="true">'
          '<circle cx="100" cy="0" r="45"/><circle cx="100" cy="0" r="70"/><circle cx="100" cy="0" r="95"/></svg>')
WATERMARK = ('<svg class="wm" viewBox="0 0 100 100" aria-hidden="true">'
             '<circle cx="100" cy="100" r="45"/><circle cx="100" cy="100" r="70"/><circle cx="100" cy="100" r="95"/></svg>')



def css():
    out = re.sub(r"(-?\d*\.?\d+)U\b", r"calc(\1*var(--u))", CSS)
    fonts = "".join(f"@font-face{{font-family:InterLocal;font-weight:{w};font-display:swap;"
                    f"src:url(assets/inter-latin-{w}-normal.woff2) format('woff2')}}\n" for w in (400, 700, 900))
    rules = []
    for mode, bgs in fd.MODES.items():
        for bg, t in bgs.items():
            v = {"bg": t["bg"], "fg": t["fg"], "body": t["body"], "quiet": t["quiet"], "divider": t["divider"], "line": t["line"],
                 "hl": t["hl_fg"], "hlbg": t["hl_bg"] or "transparent", "panel": t["panel"], "panel-fg": t["panel_fg"],
                 "panel-body": t["panel_body"], "panel-quiet": t["panel_quiet"], "panel-line": t["panel_line"] or "transparent",
                 "disc": t["disc"], "icon": t["icon"], "inv": t["inv"], "inv-fg": t["inv_fg"], "inv-line": t["inv_line"], "band": t["band"],
                 "band-fg": t["band_fg"], "foot": t["foot"], "foot-fg": t["foot_fg"], "wm": t["wm"]}
            decl = ";".join(f"--c-{k}:{val}" for k, val in v.items()) + f";--hl-pad:{'.12em' if t['hl_bg'] else '0'}"
            m = f'html[data-mode="{mode}"]'
            sel = f"{m} .bg-{bg}" + (f",{m} body" if bg == "plain" else "")
            rules.append(f"{sel}{{{decl}}}")
    return fonts + "\n".join(rules) + out


class Html:
    def __init__(self, deck, base, out):
        self.deck, self.base, self.out = deck, base, out
        self.assets = out / "assets"
        self.files = {}

    def asset(self, rel):
        """Copy a deck image next to the page and return its relative URL."""
        src = (self.base / rel).resolve()
        if src not in self.files:
            name = re.sub(r"[^A-Za-z0-9._-]", "-", src.name)
            if name in {v for v in self.files.values()}:
                name = f"{len(self.files)}-{name}"
            shutil.copyfile(src, self.assets / name)
            self.files[src] = name
        return "assets/" + self.files[src]

    def logo(self, src, alt, src_dark=None):
        if src_dark and src_dark != src:
            return (f'<img class="only-l" src="{src}" alt="{esc(alt)}"><img class="only-d" src="{src_dark}" alt="{esc(alt)}">')
        return f'<img src="{src}" alt="{esc(alt)}">'

    def hsg(self):
        return self.logo("assets/hsg_logo.png", "University of St.Gallen", "assets/hsg_logo_white.png")

    def partners(self, names=False):
        out = []
        for pt in self.deck.get("partners", []):
            img = self.logo(self.asset(pt["src"]), pt["alt"], self.asset(pt["src_dark"]) if pt.get("src_dark") else None)
            out.append(f'<span class="plogo">{img}{esc(pt["name"]) if names and pt.get("name") else ""}</span>')
        return out

    def title(self, s, tag="h2", cls=""):
        a, hl, b = fd.split_highlight(s["title"], s.get("highlight"))
        inner = (esc(a) + (f'<span class="hl">{esc(hl)}</span>' if hl else "") + esc(b)).replace("\n", "<br>")
        c = f' class="{cls}"' if cls else ""
        return f"<{tag}{c}>{inner}</{tag}>"

    @staticmethod
    def svg(name):
        svg = re.sub(r"<\?xml.*?\?>", "", fd.fetch_symbol(name)).strip()
        return svg.replace("<svg ", '<svg aria-hidden="true" ', 1)

    @staticmethod
    def icon(name, big=False):
        return f'<div class="disc{" big" if big else ""}">{Html.svg(name)}</div>'

    @staticmethod
    def qr_svg(url):
        return fd.qr_matrix(url).svg_inline(border=1, dark="#111111", light="#ffffff", omitsize=True)

    def qr(self, q):
        label = esc(q.get("label") or fd.host(q["url"]))
        return (f'<a class="qr" href="{esc(q["url"])}" target="_blank" rel="noopener" aria-label="QR code and link: {label}">'
                f'{self.qr_svg(q["url"])}<span>{label}</span></a>')

    @staticmethod
    def mark(text):
        """Escape a text; a phrase between *asterisks* is shown in green."""
        t = re.sub(r"\[([^\]]+)\]\((https?://[^)\s]+)\)", r'<a class="tl" href="\2" target="_blank" rel="noopener">\1</a>', esc(text))
        t = re.sub(r"\*\*([^*]+)\*\*", r"<strong>\1</strong>", t)
        return re.sub(r"\*([^*]+)\*", r'<span class="hl">\1</span>', t)

    @staticmethod
    def points(items):
        return '<ul class="dots">' + "".join(f"<li>{esc(x)}</li>" for x in items) + "</ul>" if items else ""

    # ---- layouts: each returns the inside of .frame

    def l_title(self, s):
        h = self.title(s, "h1")
        if s.get("kicker"):
            h = f'<p class="kicker">{esc(s["kicker"])}</p>' + h
        if s.get("mark"):
            mark = f'<span class="tmark"><img src="assets/{fd.MARKS[s["mark"]]}" alt=""></span>'
            h = h.replace("<br>", mark + "<br>", 1) if "<br>" in h else h.replace("</h1>", mark + "</h1>")
        h += '<div class="circuit" aria-hidden="true"></div>'
        if s.get("subtitle"):
            h += f'<p class="sub">{esc(s["subtitle"])}</p>'
        if s.get("presenters"):
            h += f'<p class="presenters">{esc(", ".join(s["presenters"]))}</p>'
        if s.get("meta"):
            h += f'<p class="meta">{esc(s["meta"])}</p>'
        wordmark = '<div class="wordmark"><b>Future <i>Society</i> Hub</b><span>where tech and society fuse</span></div>'
        top = "".join(f"<div>{x}</div>" for x in [self.hsg()] + self.partners()) + wordmark
        bottom = self.qr(s["qr"]) if s.get("qr") else ""
        if s.get("banner"):
            b = s["banner"]
            img = f'<img src="{self.asset(b["image"])}" alt="{esc(b["alt"])}">'
            bottom += (f'<a class="banner" href="{esc(b["url"])}" target="_blank" rel="noopener">{img}</a>' if b.get("url")
                       else f'<div class="banner">{img}</div>')
        art = self.ART
        return (f'<div class="tmain"><div class="toprow">{top}</div><div class="center">{h}</div>'
                + (f'<div class="tbottom">{bottom}</div>' if bottom else "") + f"</div>{art}")

    def l_section(self, s):
        h = self.title(s, "h2", "hs") + '<div class="circuit" aria-hidden="true"></div>'
        if s.get("text"):
            h += f'<p class="sub">{esc(s["text"])}</p>'
        return f'<div class="center">{h}</div>'

    def l_statement(self, s):
        h = self.title(s, "h2", "hs")
        if s.get("text"):
            h += f'<p class="sub" style="margin-top:calc(2.4*var(--u))">{esc(s["text"])}</p>'
        if s.get("attribution"):
            h += f'<p class="attr">{esc(s["attribution"])}</p>'
        return f'<div class="center">{h}</div>'

    ART = ('<div class="tart"><svg viewBox="0 0 100 100" aria-hidden="true"><circle cx="50" cy="50" r="20"/>'
           '<circle cx="50" cy="50" r="30"/><circle cx="50" cy="50" r="40"/><circle cx="50" cy="50" r="49.8"/></svg>'
           '<img src="assets/fuso-logo-white.svg" alt="Future Society Hub logo"></div>')

    def l_closing(self, s):
        h = self.title(s, "h2", "hs") + '<div class="circuit" aria-hidden="true"></div>'
        if s.get("text"):
            h += f'<p class="sub">{esc(s["text"])}</p>'
        h += self.contacts(s.get("contacts"))
        return (f'{self.ART}<div class="tmain"><div class="center">{h}</div>'
                + (self.qr(s["qr"]) if s.get("qr") else "") + "</div>")

    def contacts(self, items):
        """Links with an icon each, as on the closing slide."""
        def item(c):
            inner = (self.svg(c["icon"]) if c.get("icon") else "") + esc(c["label"])
            return (f'<a href="{esc(c["url"])}" target="_blank" rel="noopener">{inner}</a>' if c.get("url")
                    else f"<span>{inner}</span>")
        return '<ul class="contact">' + "".join(f"<li>{item(c)}</li>" for c in items) + "</ul>" if items else ""

    def head(self, s):
        h = self.title(s) + (f'<p class="lead">{esc(s["lead"])}</p>' if s.get("lead") else "")
        c = s.get("callout")
        if not c:
            return h
        box = ((self.icon(c["icon"]) if c.get("icon") else "") + f'<div><strong>{esc(c["heading"])}</strong>'
               + (f'<span>{esc(c["text"])}</span>' if c.get("text") else "") + "</div>")
        return f'<div class="headrow"><div>{h}</div><aside class="callout panel">{box}</aside></div>'

    def main(self, s, inner):
        src = f'<p class="source">{self.mark(s["source"])}</p>' if s.get("source") else ""
        return f'{self.head(s)}<div class="main">{inner}</div>{src}'

    def aside(self, a):
        panel = any(a.get(k) for k in ("title", "heading", "text", "facts", "qr"))
        if not panel and a.get("image"):
            cap = f'<figcaption>{esc(a["caption"])}</figcaption>' if a.get("caption") else ""
            return f'<figure class="aside pic"><img src="{self.asset(a["image"])}" alt="{esc(a["alt"])}">{cap}</figure>'
        if not panel:
            return f'<div class="aside ico">{self.icon(a["icon"], big=True)}</div>'
        # a panel reads in levels: heading, label, facts (bold term, quiet detail), a quiet note, then the call to action
        h = ""
        if a.get("title"):
            h += f'<h3 class="atitle">{esc(a["title"])}</h3>'
        if a.get("heading"):
            h += f'<h3>{esc(a["heading"])}</h3>'
        if a.get("label"):
            h += f'<p class="alabel">{esc(a["label"])}</p>'
        if a.get("photo"):   # a round portrait in the top right corner, beside title and label
            ph = a["photo"]
            h = (f'<div class="ahead"><div class="ahead-t">{h}</div>'
                 f'<img class="face" src="{self.asset(ph["image"])}" alt="{esc(ph["alt"])}"></div>')
        if a.get("text"):
            h += f'<p>{self.mark(a["text"])}</p>'
        if a.get("facts"):
            h += '<dl class="facts">' + "".join(f"<dt>{esc(k)}</dt><dd>{self.mark(v)}</dd>" for k, v in a["facts"]) + "</dl>"
        if a.get("note"):
            h += f'<p class="anote">{esc(a["note"])}</p>'
        if a.get("logos"):
            h += '<div class="aside-logos">' + "".join(
                self.logo(self.asset(l["src"]), l["alt"], self.asset(l["src_dark"]) if l.get("src_dark") else None)
                for l in a["logos"]) + "</div>"
        act = ""
        if a.get("qr"):
            q = a["qr"]
            txt = ((f'<strong>{esc(q["heading"])}</strong>' if q.get("heading") else "")
                   + (f'<span>{esc(q["text"])}</span>' if q.get("text") else "")
                   + f'<b>{esc(q.get("label") or fd.host(q["url"]))}</b>')
            act = (f'<a class="qrrow" href="{esc(q["url"])}" target="_blank" rel="noopener">{self.qr_svg(q["url"])}'
                   f'<span class="qrtext">{txt}</span></a>')
        img = ""
        if a.get("image"):
            b = a.get("badge")
            b = {"text": b} if isinstance(b, str) else b
            badge = ""
            if b:
                pic = f'<img src="{self.asset(b["image"])}" alt="">' if b.get("image") else ""
                col = f' style="color:{esc(b["color"])}"' if b.get("color") else ""
                badge = f'<span class="badge"{col}>{pic}{esc(b["text"])}</span>'
            img = f'<div class="aimg"><img src="{self.asset(a["image"])}" alt="{esc(a["alt"])}">{badge}</div>'
        cls = (" has-img" if img else "") + (" inv" if a.get("qr") else "")   # a panel with a QR code is a call to action
        return f'<aside class="aside panel{cls}">{img}<div class="aside-in">{h}</div>{act}</aside>'

    def l_text(self, s):
        pts = s.get("points") or []
        big = " big" if pts and len(pts) <= 4 and all(isinstance(x, str) and len(x) <= 64 for x in pts) else ""
        if pts and all(isinstance(x, dict) for x in pts):
            rows = '<div class="prows">' + "".join(
                f'<div class="prow panel">{self.icon(x["icon"])}<p>{self.mark(x["text"])}</p></div>' for x in pts) + "</div>"
        else:
            rows = f'<ul class="rows{big}">' + "".join(f"<li>{self.mark(x)}</li>" for x in pts) + "</ul>" if pts else ""
        copy = "".join(f"<p>{self.mark(p)}</p>" for p in s.get("paragraphs", [])) + rows
        if s.get("after"):
            copy += f'<p class="after">{self.mark(s["after"])}</p>'
        copy = f'<div class="copy">{copy}{self.contacts(s.get("contacts"))}</div>'
        if s.get("aside"):
            flip = " flip" if s["aside"].get("side") == "left" else ""
            return self.main(s, f'<div class="split{flip}">{copy}{self.aside(s["aside"])}</div>')
        return self.main(s, copy)

    def chead(self, it):
        """Heading and text of a card; with an icon, both stand beside it."""
        h = f'<h3>{esc(it["heading"])}</h3>' + (f'<p>{esc(it["text"])}</p>' if it.get("text") else "")
        return f'<div class="chead">{self.icon(it["icon"])}<div class="ctext">{h}</div></div>' if it.get("icon") else h

    def l_columns(self, s):
        cols = ""
        for it in s["items"]:
            cols += '<div class="col panel">' + self.chead(it) + "</div>"
        band = ""
        if s.get("base"):
            b = s["base"]
            band = ('<div class="band">' + (f'<strong>{esc(b["label"])}</strong>' if b.get("label") else "")
                    + f'<span>{esc(b["text"])}</span></div>')
        n = len(s["items"])
        per_row = 3 if n > 4 else 2 if n == 4 and max(len(it["heading"]) for it in s["items"]) > 28 else n   # long headings need width
        return self.main(s, f'<div class="cols" style="--n:{per_row}">{cols}</div>{band}')

    def l_steps(self, s):
        num = bool(s.get("numbered"))
        li = ""
        for i, it in enumerate(s["items"], 1):
            li += ('<li><div class="step panel">' + (f'<span class="num">{i}</span>' if num else "")
                   + self.chead(it) + "</div></li>")
        return self.main(s, f'<ol class="steps{" numbered" if num else ""}" style="--n:{len(s["items"])}">{li}</ol>')

    def l_compare(self, s):
        def side(d, inv):
            body = (f'<p>{esc(d["text"])}</p>' if d.get("text") else "") + self.points(d.get("points"))
            head = f'<div class="side-head">{self.icon(d["icon"]) if d.get("icon") else ""}<h3>{esc(d["heading"])}</h3></div>'
            return f'<div class="panel side">{head}{body}</div>'
        end = f'<p class="conclusion">{esc(s["conclusion"])}</p>' if s.get("conclusion") else ""
        return self.main(s, f'<div class="cmp">{side(s["left"], False)}{side(s["right"], True)}</div>{end}')

    def l_fuse(self, s):
        def field(d, cls):
            return f'<div class="fz {cls}"><h3>{esc(d["heading"])}</h3>' + (f'<p>{esc(d["text"])}</p>' if d.get("text") else "") + "</div>"
        def circle(d, cls):
            return f'<div class="vc {cls}">{self.svg(d["icon"]) if d.get("icon") else ""}</div>'
        venn = (f'<div class="venn" aria-hidden="true">{circle(s["left"], "l")}{circle(s["right"], "r")}<div class="vx"></div>'
                '<img src="assets/fuso-logo-white.svg" alt=""></div>')
        end = f'<p class="conclusion">{esc(s["conclusion"])}</p>' if s.get("conclusion") else ""
        return self.main(s, f'<div class="fuse">{field(s["left"], "l")}{venn}{field(s["right"], "r")}</div>{end}')

    def l_agenda(self, s):
        li = ""
        for it in s["items"]:
            sub = ""
            if it.get("sub"):
                sub = '<ul class="subs">' + "".join(
                    f'<li>{esc(x["title"])}' + (f'<span>{esc(x["who"])}</span>' if x.get("who") else "") + "</li>"
                    for x in it["sub"]) + "</ul>"
            li += (f'<li><span class="when">{esc(it.get("when", ""))}</span><div><strong>{esc(it["title"])}</strong>'
                   + (f'<span class="who">{esc(it["who"])}</span>' if it.get("who") else "") + sub + "</div></li>")
        wide = max(6.5, max(len(it.get("when", "")) for it in s["items"]) * .78 + .6)
        return self.main(s, f'<ol class="agenda" style="--when:calc({wide:.1f}*var(--u))">{li}</ol>')

    def l_people(self, s):
        n, out = len(s["items"]), ""
        for it in s["items"]:
            face = ""
            if it.get("photo"):
                alt = it.get("alt") or f"Portrait of {it['name']}"
                face = f'<img class="face" src="{self.asset(it["photo"])}" alt="{esc(alt)}">'
            who = (f'<div class="ptag">{esc(it["tag"])}</div>' if it.get("tag") else "") + f'<div class="name">{esc(it["name"])}</div>'
            if it.get("role"):
                who += '<div class="role">' + esc(it["role"]).replace("\n", "<br>") + "</div>"
            if s.get("questions") and it.get("points"):
                pts = '<ol class="nums">' + "".join(f"<li>{esc(x)}</li>" for x in it["points"]) + "</ol>"
            else:
                pts = self.points(it.get("points"))
            line = '<div class="circuit" aria-hidden="true"></div>' if n == 1 else ""
            out += (f'<div class="person{" panel" if n > 1 else ""}">{line}<div class="phead">{face}<div class="who">{who}</div></div>'
                    f"{pts}</div>")
        qs = ('<ol class="questions">' + "".join(f"<li>{esc(x)}</li>" for x in s["questions"]) + "</ol>") if s.get("questions") else ""
        many = " many" if n > 4 else ""
        return self.main(s, f'{qs}<div class="people n{n}{many}" style="--n:{min(n, 4)}">{out}</div>')

    def l_flow(self, s):
        chips = "".join(f"<li>{esc(x)}</li>" for x in s["items"])
        return self.main(s, f'<div class="flow"><div class="fend">{esc(s["start"])}</div><ul class="chips">{chips}</ul>'
                            f'<div class="fend">{esc(s["end"])}</div></div>')

    def l_matrix(self, s):
        cols = s["columns"]
        head = '<div class="mx-h"></div>' + "".join(f'<div class="mx-h">{esc(c)}</div>' for c in cols)
        rows = ""
        for r in s["rows"]:
            rows += f'<div class="mx-l">{esc(r["label"])}</div>'
            for i, c in enumerate(cols):
                mark = f'<i></i><span class="sr">{esc(c)}</span>' if c in r["for"] else ""
                rows += f'<div class="mx-c{" last" if i == len(cols) - 1 else ""}">{mark}</div>'
        return self.main(s, f'<div class="matrix" style="--n:{len(cols)}">{head}{rows}</div>')

    def l_logos(self, s):
        t = ""
        for l in s["items"]:
            img = self.logo(self.asset(l["src"]), l["alt"], self.asset(l["src_dark"]) if l.get("src_dark") else None)
            t += f'<div class="ltile{"" if l.get("src_dark") else " lt-w"}">{img}</div>'
        return self.main(s, f'<div class="lwall">{t}</div>')

    def l_links(self, s):
        out = ""
        for it in s["items"]:
            out += (f'<a class="panel link" href="{esc(it["url"])}" target="_blank" rel="noopener">'
                    + (f'<span class="src">{esc(it["source"])}</span>' if it.get("source") else "")
                    + (f'<span class="date">{esc(it["date"])}</span>' if it.get("date") else "")
                    + f'<h3>{esc(it["heading"])}</h3>' + (f'<p>{esc(it["text"])}</p>' if it.get("text") else "")
                    + f'<span class="go"><span class="host">{esc(it.get("label") or fd.host(it["url"]))}</span>{self.icon("arrow-up-right")}</span></a>')
        return self.main(s, f'<div class="cols cards" style="--n:{len(s["items"])}">{out}</div>')

    def l_image(self, s):
        cap = f'<figcaption>{esc(s["caption"])}</figcaption>' if s.get("caption") else ""
        return self.main(s, f'<figure class="bigpic"><img src="{self.asset(s["image"])}" alt="{esc(s["alt"])}">{cap}</figure>')

    def page(self):
        d = self.deck
        n_all = len(d["slides"])
        foot = ('<footer class="foot"><div class="foot-in"><a class="hub" href="https://futuresociety.ch" target="_blank" rel="noopener">'
                '<img src="assets/fuso-logo-footer.svg" alt="Future Society Hub logo"><span>Future Society Hub</span></a>'
                '<span class="pn" aria-hidden="true">{n}</span></div></footer>')
        sections = ""
        for n, s in enumerate(d["slides"], 1):
            inner = getattr(self, "l_" + s["layout"])(s)
            notes = f'<template class="notes">{esc(s["notes"])}</template>' if s.get("notes") else ""
            t = f' data-time="{esc(str(s["time"]))}"' if s.get("time") else ""
            speaker = s["layout"] == "people" and len(s["items"]) == 1
            wm = "" if s["layout"] == "title" else WATERMARK if s["layout"] in fd.RINGS or speaker else CORNER
            sections += (f'<section class="slide bg-{s["bg"]} l-{s["layout"]}{" speaker" if speaker else ""}{" active" if n == 1 else ""}" id="s{n}" '
                         f'data-bg="{s["bg"]}" data-title="{esc(s["title"])}"{t} aria-label="Slide {n} of {n_all}">'
                         f'{wm}<div class="frame">{inner}</div>{foot.format(n=n, all=n_all)}{notes}</section>\n')

        def ico(paths, cls=""):
            c = f' class="{cls}"' if cls else ""
            return f'<svg{c} viewBox="0 0 24 24" aria-hidden="true">{paths}</svg>'
        sq = "".join(f'<rect x="{x}" y="{y}" width="7" height="7" rx="1.5"/>' for x in (4, 13) for y in (4, 13))
        back, fwd = '<path d="M15 5l-7 7 7 7"/>', '<path d="M9 5l7 7-7 7"/>'
        note = '<rect x="4" y="4" width="16" height="16" rx="2.5"/><path d="M8 9h8M8 12.5h8M8 16h4"/>'
        moon = '<path d="M20 14.5A8 8 0 0 1 9.5 4a8 8 0 1 0 10.5 10.5z"/>'
        sun = ('<circle cx="12" cy="12" r="4"/><path d="M12 2v2M12 20v2M2 12h2M20 12h2M4.9 4.9l1.4 1.4M17.7 17.7l1.4 1.4'
               'M4.9 19.1l1.4-1.4M17.7 6.3l1.4-1.4"/>')
        grow, shrink = '<path d="M4 9V4h5M20 9V4h-5M4 15v5h5M20 15v5h-5"/>', '<path d="M9 4v5H4M15 4v5h5M9 20v-5H4M15 20v-5h5"/>'
        notes_btn = f'<button id="notesBtn" title="Speaker notes (N)">{ico(note)}Notes</button>' if any(s.get("notes") for s in d["slides"]) else ""
        ctl = ('<div class="controls" role="navigation" aria-label="Presentation controls"><div class="pill"><div class="nav">'
               f'<button id="prev" title="Previous slide (←)" aria-label="Previous slide">{ico(back)}</button>'
               f'<span class="count" id="count">1 / {n_all}</span>'
               f'<button id="next" title="Next slide (→)" aria-label="Next slide">{ico(fwd)}</button>'
               '</div><i class="sep"></i><div class="tools">'
               f'<button id="ovBtn" title="All slides (O)">{ico(sq)}Overview</button>{notes_btn}'
               f'<button id="modeBtn" title="Dark or light (D)">{ico(moon, "to-dark")}<span class="to-dark">Dark</span>'
               f'{ico(sun, "to-light")}<span class="to-light">Light</span></button>'
               f'<button id="fsBtn" title="Full screen (F)">{ico(grow, "fs-off")}<span class="fs-off">Full screen</span>'
               f'{ico(shrink, "fs-on")}<span class="fs-on">Exit</span></button></div></div>'
               '<div class="progress" aria-label="Slide progress"><span id="progress"></span></div></div>')
        return f"""<!doctype html>
<html lang="{esc(d.get('lang', 'en'))}" data-mode="{d['mode']}">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{esc(d['title'])} | Future Society Hub</title>
<style>{css()}</style>
</head>
<body data-bg="{d['slides'][0]['bg']}">
<main aria-label="{esc(d['title'])}">
{sections}</main>
{ctl}
<div class="overview bg-plain" id="overview" role="dialog" aria-label="Slide overview"><h2>{esc(d['title'])}</h2><div class="ov-grid"></div></div>
<div class="notes-view bg-plain" id="notesView" aria-label="Speaker notes"></div>
<script>{JS}</script>
</body>
</html>
"""

    def write(self):
        self.assets.mkdir(parents=True, exist_ok=True)
        for w in (400, 700, 900):
            shutil.copyfile(fd.ASSETS / "fonts" / f"inter-latin-{w}-normal.woff2", self.assets / f"inter-latin-{w}-normal.woff2")
        for f in ("fuso-logo-footer.svg", "fuso-logo-white.svg", "hsg_logo.png", "hsg_logo_white.png", "HSG_Impact_Space_Flache_RGB.png"):
            shutil.copyfile(fd.ASSETS / f, self.assets / f)
        for s in self.deck["slides"]:
            if s.get("mark"):
                shutil.copyfile(fd.ASSETS / fd.MARKS[s["mark"]], self.assets / fd.MARKS[s["mark"]])
        page = self.page()
        (self.out / "index.html").write_text(page, encoding="utf-8")
        return self.out / "index.html"


OVERFLOW_JS = """() => {
  const s=document.querySelector('.slide.active'), f=s.querySelector('.frame'), bad=[];
  const fr=f.getBoundingClientRect();
  if(f.scrollHeight>f.clientHeight+1||f.scrollWidth>f.clientWidth+1) bad.push('content is larger than the slide');
  for(const el of f.querySelectorAll('.main,.center,.panel,.col,.person,.steps li,.side')){
    if(el.scrollHeight>el.clientHeight+1&&(el.matches('.main')||getComputedStyle(el).overflow!=='visible')) bad.push('the content runs into the source line or footer');
    const r=el.getBoundingClientRect();
    if(r.bottom>fr.bottom+1||r.right>fr.right+1) bad.push('a block reaches past the slide');
  }
  const m=f.querySelector('.main');
  if(m&&m.lastElementChild&&!s.classList.contains('speaker')){const top=m.getBoundingClientRect().top+parseFloat(getComputedStyle(m).paddingTop);
    let lo=Infinity,hi=0;for(const c of m.children){const r=c.getBoundingClientRect();lo=Math.min(lo,r.top);hi=Math.max(hi,r.bottom)}
    const zone=m.getBoundingClientRect().bottom-top, fill=(hi-lo)/zone;
    if(fill<0.55) bad.push('THIN: the content fills only '+Math.round(fill*100)+'% of the slide');}
  return [...new Set(bad)];
}"""


def shots(index, folder, n):
    """Screenshots for the visual check. Returns a list of overflow messages."""
    from playwright.sync_api import sync_playwright
    from PIL import Image
    folder.mkdir(parents=True, exist_ok=True)
    for old in list(folder.glob("slide-*.png")) + list(folder.glob("phone-*.png")) + list(folder.glob("contact*.png")):  # stale shots
        old.unlink()
    problems, thin = [], []
    with sync_playwright() as p:
        try:
            b = p.chromium.launch(channel="msedge")
        except Exception:
            b = p.chromium.launch()
        pg = b.new_page(viewport={"width": 1440, "height": 810})
        pg.goto(index.as_uri())
        pg.evaluate("document.fonts.ready")
        pg.wait_for_timeout(300)
        paths = []
        for i in range(n):
            pg.evaluate("i => window.fusoDeck.go(i)", i)
            pg.wait_for_timeout(60)
            for msg in pg.evaluate(OVERFLOW_JS):
                (thin if msg.startswith("THIN") else problems).append(f"slide {i + 1}: {msg.replace('THIN: ', '')}")
            path = folder / f"slide-{i + 1:02d}.png"
            pg.screenshot(path=str(path))
            paths.append(path)
        pg.close()
        ph = b.new_page(viewport={"width": 390, "height": 844})
        ph.goto(index.as_uri())
        ph.evaluate("document.fonts.ready")
        ph.wait_for_timeout(300)
        phones = []
        for i in range(n):
            ph.set_viewport_size({"width": 390, "height": 844})
            ph.evaluate("i => window.fusoDeck.go(i)", i)
            ph.wait_for_timeout(60)
            wide, tall = ph.evaluate("() => { const s = document.querySelector('.slide.active');"
                                     " return [s.scrollWidth > s.clientWidth + 1, s.scrollHeight - s.clientHeight]; }")
            if wide:
                problems.append(f"slide {i + 1}, phone: something is wider than the screen and does not wrap")
            if tall > 0:                      # show the whole slide, not only the part above the fold
                ph.set_viewport_size({"width": 390, "height": 844 + tall})
                ph.wait_for_timeout(60)
            path = folder / f"phone-{i + 1:02d}.png"
            ph.screenshot(path=str(path))
            phones.append(path)
        b.close()
    pw = 300
    for c, start in enumerate(range(0, len(phones), 6), 1):
        ims = [Image.open(x).convert("RGB") for x in phones[start:start + 6]]
        ims = [im.resize((pw, im.height * pw // im.width)) for im in ims]
        sheet = Image.new("RGB", (len(ims) * (pw + 12) + 12, max(im.height for im in ims) + 24), "#888888")
        for k, im in enumerate(ims):
            sheet.paste(im, (12 + k * (pw + 12), 12))
        sheet.save(folder / f"contact-phone-{c:02d}.png")
    cols, tw = 3, 640
    th = tw * 810 // 1440
    for c, start in enumerate(range(0, len(paths), 9), 1):
        chunk = paths[start:start + 9]
        rows = -(-len(chunk) // cols)
        sheet = Image.new("RGB", (cols * tw + (cols + 1) * 12, rows * th + (rows + 1) * 12), "#888888")
        for j, pth in enumerate(chunk):
            t = Image.open(pth).convert("RGB").resize((tw, th))
            sheet.paste(t, (12 + (j % cols) * (tw + 12), 12 + (j // cols) * (th + 12)))
        sheet.save(folder / f"contact-{c:02d}.png")
    if thin:
        print("Thin slides (add the evidence, example or picture that proves the headline, or merge with a neighbour):")
        for m in thin:
            print("  - " + m)
    return problems


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("deck")
    ap.add_argument("--out", help="output folder (default: html/ next to the deck)")
    ap.add_argument("--shots", help="folder for check screenshots (keep it outside the repo)")
    a = ap.parse_args()
    deck, base, warnings = fd.load(a.deck)
    out = Path(a.out).resolve() if a.out else base / "html"
    index = Html(deck, base, out).write()
    print(f"Saved {index} ({len(deck['slides'])} slides, {deck['mode']} mode)")
    fd.report(warnings)
    if a.shots:
        problems = shots(index, Path(a.shots).resolve(), len(deck["slides"]))
        print(f"Screenshots in {Path(a.shots).resolve()}: slide-NN.png, contact-NN.png, phone-NN.png, contact-phone-NN.png")
        if problems:
            print("OVERFLOW, shorten the text or split the slide (never change sizes):", file=sys.stderr)
            for m in problems:
                print("  - " + m, file=sys.stderr)
            sys.exit(2)


if __name__ == "__main__":
    main()
