# -*- coding: utf-8 -*-
from pathlib import Path
import html
import json
import os
import re
import zipfile
import xml.etree.ElementTree as ET

OUT = Path(__file__).resolve().parent
MSSK = OUT / "mssk"
MSSK.mkdir(parents=True, exist_ok=True)

docx = Path(os.environ["BIM_DOCX"])
xlsx = Path(os.environ["BIM_XLSX"])

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"


def esc(value):
    return html.escape(value or "", quote=True)


def slug(value, idx):
    text = re.sub(r"[^\w\dа-яА-ЯёЁ]+", "-", value.lower(), flags=re.U).strip("-")
    text = text[:60].strip("-")
    return f"{text or 'section'}-{idx}"


def struck(node):
    for el in node.iter():
        if el.tag in (W + "strike", W + "dstrike"):
            val = el.attrib.get(W + "val")
            if val not in ("false", "0"):
                return True
    return False


def text_from(node):
    return "".join(t.text or "" for t in node.iter(W + "t")).strip()


def style_of(paragraph):
    ppr = paragraph.find(W + "pPr")
    style = ppr.find(W + "pStyle") if ppr is not None else None
    return style.attrib.get(W + "val", "") if style is not None else ""


def cell_colspan(cell):
    props = cell.find(W + "tcPr")
    grid_span = props.find(W + "gridSpan") if props is not None else None
    value = grid_span.attrib.get(W + "val", "1") if grid_span is not None else "1"
    return int(value) if value.isdigit() and int(value) > 1 else 1


CSS = ':root{--bg:#f2f4ef;--paper:#fffefa;--ink:#18211d;--muted:#65706a;--line:#d7dcd4;--green:#176b52;--green-soft:#e3f1eb;--blue:#285e88;--blue-soft:#e5eff7;--amber:#956414;--amber-soft:#fff1ce;--red:#9a3f3f;--shadow:0 24px 70px rgba(35,43,38,.12)}*{box-sizing:border-box}html{scroll-behavior:auto}body{margin:0;color:var(--ink);background:radial-gradient(circle at 8% 0%,rgba(23,107,82,.11),transparent 30rem),radial-gradient(circle at 92% 12%,rgba(40,94,136,.10),transparent 26rem),var(--bg);font-family:Inter,"Segoe UI",Arial,sans-serif;line-height:1.55}a{color:inherit}.layout{width:min(1280px,calc(100% - 32px));margin:0 auto;padding:28px 0 70px}.hero{position:relative;overflow:hidden;padding:54px;border:1px solid var(--line);border-radius:30px;background:var(--paper);box-shadow:var(--shadow)}.hero:after{content:"IFC";position:absolute;right:-18px;bottom:-78px;color:rgba(23,107,82,.06);font-size:220px;font-weight:900;letter-spacing:-12px;pointer-events:none}.kicker,.tag{display:inline-flex;align-items:center;width:max-content;border-radius:999px;font-size:12px;font-weight:800;letter-spacing:.09em;text-transform:uppercase}.kicker{padding:8px 12px;color:var(--green);background:var(--green-soft)}h1{max-width:950px;margin:22px 0 16px;font-size:clamp(42px,7vw,76px);line-height:1;letter-spacing:-.045em}.intro{max-width:900px;margin:0;color:var(--muted);font-size:19px}.meta{display:flex;flex-wrap:wrap;gap:10px;margin-top:28px}.meta span,.meta a{padding:9px 13px;border:1px solid var(--line);border-radius:12px;background:#faf9f4;font-size:13px;text-decoration:none}.notice{position:relative;z-index:1;max-width:920px;margin-top:24px;padding:14px 16px;border-left:4px solid var(--amber);background:var(--amber-soft);color:#66450f;font-size:14px}.body-grid{display:grid;grid-template-columns:270px minmax(0,1fr);gap:28px;margin-top:28px;align-items:start}nav{position:sticky;top:18px;max-height:calc(100vh - 36px);overflow-y:auto;padding:18px;border:1px solid var(--line);border-radius:20px;background:rgba(255,254,250,.9);backdrop-filter:blur(12px)}nav strong{display:block;margin:0 8px 10px;font-size:13px;color:var(--muted);text-transform:uppercase;letter-spacing:.08em}nav a{display:block;padding:8px 10px;border-radius:10px;color:#37413d;font-size:13px;text-decoration:none}nav a:hover{color:var(--green);background:var(--green-soft)}nav a.active-nav{background:#e8eee9;color:#1f352d;font-weight:850;box-shadow:inset 3px 0 0 var(--green)}.nav-group{margin-bottom:10px}.nav-main{font-weight:800}.nav-sub{display:grid;gap:1px;margin:2px 0 0 11px;padding-left:9px;border-left:1px solid var(--line)}nav .nav-sub a{padding:4px 8px;color:var(--muted);font-size:12px;line-height:1.3}nav .nav-sub a.nav-sub2{margin-left:12px;padding-left:12px;font-size:11px;border-left:1px solid #e1e3dd;color:#78827d}nav .nav-sub a.active-nav{color:#1f352d}nav .nav-sub a.nav-sub2.active-nav{color:#1f352d;background:#eef3ef}.content{min-width:0}[id]{scroll-margin-top:22px}section{scroll-margin-top:22px;margin-bottom:22px;padding:34px;border:1px solid var(--line);border-radius:24px;background:var(--paper)}.section-head{display:flex;justify-content:space-between;gap:18px;align-items:flex-start;margin-bottom:24px}.section-head h2{margin:4px 0 0;font-size:32px;line-height:1.12;letter-spacing:-.03em}.tag{padding:7px 10px}.tag.all{color:var(--green);background:var(--green-soft)}.tag.data{color:var(--blue);background:var(--blue-soft)}.stats{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:12px;margin-top:22px}.stat{padding:16px;border:1px solid var(--line);border-radius:14px;background:#faf9f4}.stat b{display:block;font-size:24px;line-height:1}.stat span{display:block;margin-top:4px;color:var(--muted);font-size:13px}.search{position:relative;z-index:1;margin:0 0 24px;padding:0 0 18px;border-bottom:1px solid var(--line);background:transparent}.search label{display:block;margin-bottom:8px;color:var(--muted);font-size:13px;font-weight:800;text-transform:uppercase;letter-spacing:.08em}.search input{width:100%;height:46px;padding:0 14px;border:1px solid var(--line);border-radius:12px;background:white;color:var(--ink);font:inherit}.search small{display:block;margin-top:8px;color:var(--muted)}.search-status{display:flex;align-items:center;justify-content:space-between;gap:12px;margin-top:8px}.search-status small{margin-top:0}.search-actions{display:inline-flex;gap:6px;flex:0 0 auto}.search-actions button{display:grid;place-items:center;width:34px;height:34px;border:1px solid var(--line);border-radius:10px;padding:0;background:#faf9f4;color:var(--ink);font:inherit;font-size:18px;font-weight:850;line-height:1;cursor:pointer}.search-actions button:hover{border-color:var(--green);color:var(--green)}.search-actions button:disabled{cursor:not-allowed;opacity:.45}.active-search-result{outline:3px solid rgba(23,107,82,.28);outline-offset:4px;border-radius:8px;background:#f0f7f3}mark.search-hit{padding:1px 3px;border-radius:4px;background:#36d399;color:#072116;box-decoration-break:clone;-webkit-box-decoration-break:clone}.doc-heading{margin:28px 0 12px}.doc-heading:first-child{margin-top:0}.doc-heading.doc-block{padding-top:8px}.doc-p{margin:0 0 10px;color:#2f3935}.doc-p.bullet{padding-left:18px}.table-wrap{overflow:auto;border:1px solid var(--line);border-radius:14px;margin:14px 0 22px;max-height:76vh}table{width:100%;border-collapse:collapse;font-size:14px;background:#fffefa}th,td{padding:10px 12px;text-align:left;border-bottom:1px solid var(--line);border-right:1px solid var(--line);vertical-align:top;min-width:90px}th{position:sticky;top:0;z-index:1;color:var(--muted);background:#f5f3ed;font-size:12px;text-transform:uppercase;letter-spacing:.04em}tr:last-child td{border-bottom:0}code{padding:2px 6px;border-radius:5px;background:#eceae3;font-family:Consolas,monospace;font-size:.92em}.doc-table table{font-size:13px}.sheet-grid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:12px}.sheet-card{display:block;padding:18px;border:1px solid var(--line);border-radius:14px;background:#faf9f4;text-decoration:none}.sheet-card b{display:block;margin-bottom:4px}.sheet-card span{color:var(--muted);font-size:13px}.mssk-table{font-size:12px}.mssk-table th,.mssk-table td{padding:7px 9px;min-width:80px}.empty{display:none;padding:16px 18px;border-radius:14px;background:#f7e8e5;color:#723333}.flow{counter-reset:step;display:grid;gap:12px}.flow article{position:relative;padding:18px 18px 18px 56px;border:1px solid var(--line);border-radius:14px;background:#faf9f4}.flow article:before{counter-increment:step;content:counter(step);position:absolute;left:16px;top:16px;display:grid;place-items:center;width:28px;height:28px;border-radius:50%;color:white;background:var(--green);font-weight:800}.flow h3{margin:0 0 4px;font-size:16px}.flow p{margin:0;color:var(--muted);font-size:14px}footer{padding:20px 6px 0;color:var(--muted);font-size:13px;text-align:center}@media(max-width:850px){.hero{padding:34px 26px}.body-grid{grid-template-columns:1fr}nav{position:static;display:flex;gap:4px;max-height:none;overflow-x:auto;overflow-y:hidden}nav strong{display:none}.nav-group{flex:0 0 auto;margin:0;padding-right:5px;border-right:1px solid var(--line)}nav a{white-space:nowrap}.nav-sub{margin-left:10px}.stats,.sheet-grid{grid-template-columns:1fr}}@media(max-width:520px){.layout{width:min(100% - 18px,1280px);padding-top:9px}.hero,section{border-radius:18px}section{padding:24px 18px}.section-head{display:block}.section-head .tag{margin-top:12px}.stats{grid-template-columns:1fr}.search-status{align-items:flex-start;flex-direction:column}}'
JS = """const searchInput=document.querySelector('#searchInput');const resultCount=document.querySelector('#resultCount');const empty=document.querySelector('#emptySearch');const blocks=Array.from(document.querySelectorAll('.doc-block,.mssk-row'));blocks.forEach(el=>{el.dataset.originalHtml=el.innerHTML;el.dataset.searchText=norm(el.textContent)});let timer;let matches=[];let current=-1;let activeEl=null;function norm(s){return(s||'').toLowerCase().replace(/ё/g,'е')}function escapeRegExp(s){return s.replace(/[.*+?^${}()|[\\]\\\\]/g,'\\\\$&')}function clearActive(){document.querySelectorAll('.active-search-result').forEach(x=>x.classList.remove('active-search-result'));if(activeEl){activeEl.innerHTML=activeEl.dataset.originalHtml;activeEl=null}}function highlightElement(el,q){if(activeEl&&activeEl!==el)activeEl.innerHTML=activeEl.dataset.originalHtml;activeEl=el;const source=el.dataset.originalHtml;const re=new RegExp(escapeRegExp(q),'gi');el.innerHTML=source.replace(re,match=>`<mark class=\"search-hit\">${match}</mark>`)}function setButtons(){document.querySelectorAll('[data-search-nav]').forEach(btn=>btn.disabled=!matches.length)}function jumpTo(i){if(!matches.length)return;clearActive();current=(i+matches.length)%matches.length;const el=matches[current];el.classList.add('active-search-result');highlightElement(el,searchInput.value.trim());const firstHit=el.querySelector('mark.search-hit');(firstHit||el).scrollIntoView({behavior:'smooth',block:'center'});resultCount.textContent=`Показано ${current+1} из ${matches.length}`;}function applySearch(){const q=norm(searchInput.value.trim());matches=[];current=-1;clearActive();if(!q){resultCount.textContent='Поиск по текущей странице';empty.style.display='none';setButtons();return}matches=blocks.filter(el=>el.dataset.searchText.includes(q));resultCount.textContent=matches.length?`Найдено совпадений: ${matches.length}`:'Ничего не найдено';empty.style.display=matches.length?'none':'block';setButtons()}function activateNav(id,scrollMenu=false){if(!id)return;const link=document.querySelector(`nav a[href=\"#${CSS.escape(id)}\"]`);if(!link)return;document.querySelectorAll('nav a.active-nav').forEach(a=>a.classList.remove('active-nav'));link.classList.add('active-nav');if(scrollMenu)link.scrollIntoView({block:'nearest'})}function setupActiveNav(){const navLinks=Array.from(document.querySelectorAll('nav a[href^=\"#\"]'));const observed=navLinks.map(a=>document.getElementById(decodeURIComponent(a.getAttribute('href').slice(1)))).filter(Boolean);if(!observed.length)return;let activeId='';const pick=()=>{let candidate=observed[0].id;for(const el of observed){if(el.getBoundingClientRect().top<=180)candidate=el.id;else break}if(candidate!==activeId){activeId=candidate;activateNav(activeId,false)}};let ticking=false;window.addEventListener('scroll',()=>{if(ticking)return;ticking=true;requestAnimationFrame(()=>{pick();ticking=false})},{passive:true});window.addEventListener('hashchange',()=>{const hashId=decodeURIComponent(location.hash.slice(1));const exact=document.querySelector(`nav a[href=\"#${CSS.escape(hashId)}\"]`);if(exact){activeId=hashId;activateNav(hashId,true)}else pick()});if(location.hash){const hashId=decodeURIComponent(location.hash.slice(1));const exact=document.querySelector(`nav a[href=\"#${CSS.escape(hashId)}\"]`);if(exact){activeId=hashId;activateNav(hashId,true)}else pick()}else{pick()}}setupActiveNav();if(searchInput){const status=document.createElement('div');status.className='search-status';resultCount.parentNode.insertBefore(status,resultCount);status.appendChild(resultCount);const actions=document.createElement('div');actions.className='search-actions';actions.innerHTML='<button type=\"button\" data-search-nav=\"prev\" title=\"Предыдущее совпадение\" aria-label=\"Предыдущее совпадение\">↑</button><button type=\"button\" data-search-nav=\"next\" title=\"Следующее совпадение\" aria-label=\"Следующее совпадение\">↓</button>';status.appendChild(actions);actions.addEventListener('click',e=>{const btn=e.target.closest('button');if(!btn)return;if(btn.dataset.searchNav==='prev')jumpTo(current<0?matches.length-1:current-1);if(btn.dataset.searchNav==='next')jumpTo(current<0?0:current+1)});searchInput.addEventListener('input',()=>{clearTimeout(timer);timer=setTimeout(applySearch,450)});searchInput.addEventListener('keydown',e=>{if(e.key==='Enter'||e.key==='ArrowDown'){e.preventDefault();if(!matches.length)applySearch();jumpTo(current<0?0:current+1)}if(e.key==='ArrowUp'){e.preventDefault();if(!matches.length)applySearch();jumpTo(current<0?matches.length-1:current-1)}});setButtons();}"""


def page(title, body):
    return f"""<!doctype html>
<html lang="ru">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>{esc(title)}</title>
  <style>{CSS}</style>
</head>
<body>
  <div class="layout">{body}</div>
  <script>{JS}</script>
</body>
</html>"""


with zipfile.ZipFile(docx) as archive:
    root = ET.fromstring(archive.read("word/document.xml"))

body = root.find(W + "body")
doc_parts = []
doc_nav_groups = []
current_nav_group = None
doc_blocks = 0
doc_tables = 0
skipped = 0
headings = 0

for child in list(body):
    if child.tag == W + "p":
        text = text_from(child)
        if not text:
            continue
        if struck(child):
            skipped += 1
            continue
        style = style_of(child)
        doc_blocks += 1
        if style in ("1", "2"):
            headings += 1
            tag = "h2" if style == "1" else "h3"
            sid = slug(text, headings)
            doc_parts.append(f'<{tag} id="{sid}" class="doc-heading doc-block">{esc(text)}</{tag}>')
            if style == "1":
                current_nav_group = {"title": text, "id": sid, "children": []}
                doc_nav_groups.append(current_nav_group)
            elif current_nav_group is not None:
                current_nav_group["children"].append({"title": text, "id": sid})
        else:
            cls = "doc-p bullet doc-block" if re.match(r"^[\-−–]", text) else "doc-p doc-block"
            doc_parts.append(f'<p class="{cls}">{esc(text)}</p>')
    elif child.tag == W + "tbl":
        rows_html = []
        for tr in child.findall(W + "tr"):
            if struck(tr):
                skipped += 1
                continue
            cells = []
            for tc in tr.findall(W + "tc"):
                values = []
                for p in tc.findall(W + "p"):
                    if struck(p):
                        continue
                    text = text_from(p)
                    if text:
                        values.append(esc(text))
                span = cell_colspan(tc)
                attr = f' colspan="{span}"' if span > 1 else ""
                cells.append({"attr": attr, "html": "<br>".join(values)})
            if any(cell["html"].strip() for cell in cells):
                cell_tag = "th" if not rows_html else "td"
                rows_html.append(
                    "<tr>"
                    + "".join(f"<{cell_tag}{cell['attr']}>{cell['html']}</{cell_tag}>" for cell in cells)
                    + "</tr>"
                )
        if rows_html:
            doc_tables += 1
            doc_blocks += 1
            doc_parts.append('<div class="table-wrap doc-table doc-block"><table>' + "".join(rows_html) + "</table></div>")

with zipfile.ZipFile(xlsx) as archive:
    wb = ET.fromstring(archive.read("xl/workbook.xml"))
    rels = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    relmap = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rels}
    shared = []
    if "xl/sharedStrings.xml" in archive.namelist():
        sst = ET.fromstring(archive.read("xl/sharedStrings.xml"))
        for si in sst.findall(f"{{{NS_MAIN}}}si"):
            shared.append("".join(t.text or "" for t in si.iter(f"{{{NS_MAIN}}}t")))

    sheet_cards = []
    sheet_links = []
    total_rows = 0
    total_cells = 0
    sheets = wb.find(f"{{{NS_MAIN}}}sheets")

    for index, sheet in enumerate(sheets.findall(f"{{{NS_MAIN}}}sheet"), 1):
        name = sheet.attrib["name"]
        rid = sheet.attrib[R + "id"]
        target = "xl/" + relmap[rid]
        sheet_root = ET.fromstring(archive.read(target))
        rows_out = []
        max_col = 0

        for row in sheet_root.findall(f".//{{{NS_MAIN}}}row"):
            values = {}
            for cell in row.findall(f"{{{NS_MAIN}}}c"):
                ref = cell.attrib.get("r", "")
                match = re.match(r"[A-Z]+", ref)
                if not match:
                    continue
                col = 0
                for char in match.group(0):
                    col = col * 26 + ord(char) - 64
                cell_type = cell.attrib.get("t")
                v = cell.find(f"{{{NS_MAIN}}}v")
                value = v.text if v is not None and v.text is not None else ""
                if cell_type == "s" and value.isdigit():
                    idx = int(value)
                    value = shared[idx] if idx < len(shared) else ""
                elif cell_type == "inlineStr":
                    value = "".join(t.text or "" for t in cell.iter(f"{{{NS_MAIN}}}t"))

                if value != "":
                    max_col = max(max_col, col)
                    values[col] = value
                    total_cells += 1
            if values:
                rows_out.append(values)

        total_rows += len(rows_out)
        file_name = f"sheet-{index:02d}.html"
        sheet_links.append(f'<a href="mssk/{file_name}">{esc(name)}</a>')
        sheet_cards.append(
            f'<a class="sheet-card" href="mssk/{file_name}"><b>{esc(name)}</b></a>'
        )

        table = ['<div class="table-wrap"><table class="mssk-table">']
        for row_index, values in enumerate(rows_out):
            tag = "th" if row_index == 0 else "td"
            cells = "".join(f"<{tag}>{esc(values.get(col, ''))}</{tag}>" for col in range(1, max_col + 1))
            table.append(f'<tr class="mssk-row">{cells}</tr>')
        table.append("</table></div>")

        sheet_body = f"""<header class="hero">
  <span class="kicker">МССК вер. 5.0</span>
  <h1>{esc(name)}</h1>
  <div class="meta"><a href="../index.html#mssk">Назад к списку</a></div>
</header>
<section>
  <div class="search"><label for="searchInput">Поиск по листу</label><input id="searchInput" type="search" placeholder="Введите код, параметр или название"><small id="resultCount">Поиск по текущей странице</small></div>
  <div id="emptySearch" class="empty">По этому запросу ничего не найдено.</div>
  {''.join(table)}
</section>"""
        (MSSK / file_name).write_text(page(f"МССК - {name}", sheet_body), encoding="utf-8")

doc_nav_html = []
for group in doc_nav_groups:
    children = "".join(f'<a class="nav-sub2" href="#{child["id"]}">{esc(child["title"])}</a>' for child in group["children"])
    doc_nav_html.append(f'<a href="#{group["id"]}">{esc(group["title"])}</a>{children}')

main_body = f"""<header class="hero">
  <span class="kicker">Полная веб-версия</span>
  <h1>Регламент BIM для ЦИМ АГР</h1>
</header>
<div class="body-grid">
  <nav aria-label="Навигация">
    <strong>Содержание</strong>
    <div class="nav-group"><a class="nav-main" href="#overview">Сводка</a></div>
    <div class="nav-group"><a class="nav-main" href="#docx">Требования DOCX</a><div class="nav-sub">{''.join(doc_nav_html)}</div></div>
    <div class="nav-group"><a class="nav-main" href="#mssk">МССК 5.0</a><div class="nav-sub">{''.join(sheet_links)}</div></div>
    <div class="nav-group"><a class="nav-main" href="#finish">Как пользоваться</a></div>
  </nav>
  <main class="content">
    <section id="overview">
      <div class="section-head"><div><span class="tag all">Состав сайта</span><h2>Полная версия материалов</h2></div></div>
      <div class="search"><label for="searchInput">Поиск по DOCX</label><input id="searchInput" type="search" placeholder="Например: IfcSpace, RUS_FNO, ведомость, уровень, площадь"><small id="resultCount">Поиск по текущей странице</small></div>
      <div id="emptySearch" class="empty">По этому запросу ничего не найдено.</div>
    </section>
    <section id="docx"><div class="section-head"><div><span class="tag all">DOCX</span><h2>Требования к материалам в формате IFC</h2></div></div>{''.join(doc_parts)}</section>
    <section id="mssk"><div class="section-head"><div><span class="tag data">XLSX</span><h2>Московская строительная система классификаторов, версия 5.0</h2></div></div><p class="doc-p">Все листы Excel вынесены в отдельные HTML-страницы, чтобы основной регламент открывался быстро. На каждой странице есть собственный поиск по строкам.</p><div class="sheet-grid">{''.join(sheet_cards)}</div></section>
    <section id="finish"><div class="section-head"><div><span class="tag all">Рабочий порядок</span><h2>Как пользоваться полной версией</h2></div></div><div class="flow"><article><h3>Ищите по DOCX на главной</h3><p>Поиск сверху фильтрует полный текст требований и таблицы Word.</p></article><article><h3>Открывайте листы МССК отдельно</h3><p>Каждый лист Excel доступен целиком в подпапке mssk и имеет свой поиск.</p></article><article><h3>Сверяйте с исходниками при споре</h3><p>HTML создан из файлов в W:\\Test_AI_addon\\Сайт, но юридический приоритет остается за исходным утвержденным документом.</p></article></div></section>
  </main>
</div>
<footer>Источник: {esc(str(docx))} и {esc(str(xlsx))}. Сгенерировано локально для быстрого просмотра и поиска.</footer>"""

(OUT / "index.html").write_text(page("Полный регламент BIM: ЦИМ АГР IFC + МССК", main_body), encoding="utf-8")

print(json.dumps({
    "index": str(OUT / "index.html"),
    "index_bytes": (OUT / "index.html").stat().st_size,
    "mssk_pages": len(sheet_cards),
    "doc_blocks": doc_blocks,
    "doc_tables": doc_tables,
    "skipped_struck": skipped,
    "xlsx_rows": total_rows,
    "xlsx_cells": total_cells,
}, ensure_ascii=False, indent=2))
