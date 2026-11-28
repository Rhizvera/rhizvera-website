"""Build the Rhizvera Insights blog (static pages for GitHub Pages).

Only published posts are built. Article text comes from each article's Word file;
nothing unpublished is written to this repository.

  Publish a post:  python scripts/build_blog.py add <article.docx> <cover image> <YYYY-MM-DD> [rank]
                   (rank orders posts that share a date; lower comes first)
  Rebuild index:   python scripts/build_blog.py index
  Preview (no repo changes): python scripts/build_blog.py preview <outdir> <docx>... (covers optional: docx=image)
"""
import html, json, os, re, sys, shutil, datetime
import docx
from PIL import Image

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
BLOG = os.path.join(ROOT, "blog")
MANIFEST = os.path.join(BLOG, "posts.json")
SITE = "https://rhizvera.com"
GTM = '''<script async src="https://www.googletagmanager.com/gtm.js?id=GTM-W5VX366B"></script>
<script>window.dataLayer = window.dataLayer || []; function gtag(){dataLayer.push(arguments);} gtag('js', new Date());</script>'''

CSS = """
  :root { --ink:#172B5C; --root:#23736E; --teal:#45A9A1; --gold:#EBBC58; --root-dark:#1B5C58;
          --paper:#FBFCFC; --sage:#E8F2F1; --text:#2B3035; --muted:#5B6670; --line:#C8D8D6; }
  * { box-sizing: border-box; }
  html { -webkit-text-size-adjust: 100%; }
  body { margin:0; background:var(--paper); color:var(--text); font:400 1.0625rem/1.65 "Public Sans", system-ui, sans-serif; }
  a { color: var(--root); }
  a:focus-visible { outline:3px solid var(--root); outline-offset:3px; }
  .wrap { max-width:72rem; margin:0 auto; padding:0 1.5rem; }
  header.site { padding:1.5rem 0; }
  header.site .wrap { display:flex; align-items:center; justify-content:space-between; gap:1rem; flex-wrap:wrap; }
  .wordmark img { height:4rem; width:auto; display:block; }
  @media (max-width:520px) { .wordmark img { height:2.5rem; } }
  .header-links { display:flex; align-items:center; gap:1.5rem; }
  .header-link { font-weight:600; font-size:.95rem; text-decoration:none; color:var(--root); border-bottom:2px solid transparent; white-space:nowrap; }
  .header-link:hover, .header-link[aria-current] { border-bottom-color:var(--root); }
  @media (max-width:560px) { .wordmark img { height:2.25rem; } .header-links { gap:1rem; } .header-link { font-size:.85rem; } }
  h1, h2 { font-family:"Newsreader", Georgia, serif; font-weight:400; color:var(--ink); }
  .eyebrow { font-weight:600; font-size:.8rem; letter-spacing:.08em; text-transform:uppercase; color:var(--root); margin:0 0 .75rem; }
  /* index */
  .blog-head { padding:2.5rem 0 2rem; border-bottom:1px solid var(--line); }
  .blog-head h1 { font-size:clamp(2.25rem,4.5vw,3.5rem); line-height:1.05; margin:0 0 1rem; }
  .blog-head p { max-width:40rem; color:var(--muted); margin:0; }
  .posts { list-style:none; margin:0; padding:2.5rem 0 4rem; display:grid; grid-template-columns:repeat(auto-fill,minmax(19rem,1fr)); gap:2.5rem 2rem; }
  .posts a { text-decoration:none; color:inherit; display:block; }
  .posts img { width:100%; height:auto; aspect-ratio:16/9; object-fit:cover; display:block; border-radius:3px; background:var(--sage); }
  .posts .date { font-size:.85rem; color:var(--muted); margin:.9rem 0 .35rem; }
  .posts h2 { font-size:1.5rem; line-height:1.2; margin:0 0 .5rem; }
  .posts a:hover h2 { color:var(--root); }
  .posts .dek { color:var(--muted); font-size:.975rem; margin:0; }
  .empty { padding:3rem 0 5rem; color:var(--muted); }
  /* article */
  article { padding:1.5rem 0 4rem; }
  .article-head, .article-body { max-width:42rem; margin:0 auto; }
  .article-head h1 { font-size:clamp(2.1rem,4.5vw,3.25rem); line-height:1.08; letter-spacing:-.01em; margin:0 0 1rem; }
  .dek { font-size:1.2rem; line-height:1.5; color:var(--muted); margin:0 0 1rem; }
  .meta { font-size:.9rem; color:var(--muted); margin:0 0 1.75rem; }
  .cover { max-width:56rem; margin:0 auto 2.25rem; }
  .cover img { width:100%; height:auto; display:block; border-radius:3px; }
  .who { font-weight:600; color:var(--ink); border-left:4px solid var(--root); padding:.15rem 0 .15rem 1rem; margin:0 0 1.75rem; }
  .article-body h2 { font-size:1.75rem; line-height:1.2; margin:2.5rem 0 .75rem; }
  .article-body p { margin:0 0 1.1rem; }
  .article-body ul { margin:0 0 1.25rem; padding-left:1.25rem; }
  .article-body li { margin:0 0 .5rem; }
  .article-body .callout { font-weight:600; color:var(--ink); }
  .about { background:var(--sage); padding:1.5rem; border-radius:3px; margin-top:2.5rem; }
  .about h2 { margin-top:0 !important; font-size:1.4rem !important; }
  .disclaimer { font-size:.85rem; font-style:italic; color:var(--muted); margin-bottom:0 !important; }
  .copyright { margin-top:2rem; font-size:.9rem; color:var(--muted); }
  .copyright p { margin:0 !important; }
  .copyright .co { font-weight:600; color:var(--ink); }
  .sources { margin-top:2.5rem; border-top:1px solid var(--line); padding-top:.5rem; font-size:.9rem; }
  .sources h2 { font-size:1.35rem !important; }
  .sources li { overflow-wrap:anywhere; }
  .back { display:inline-block; margin-top:2.5rem; font-weight:600; text-decoration:none; }
  footer.site { padding:2.5rem 0 3rem; border-top:1px solid var(--line); color:var(--muted); font-size:.9rem; }
  footer.site .wrap { display:flex; justify-content:space-between; flex-wrap:wrap; gap:1rem; }
"""

def page(title, desc, body, canonical, og_image=None, og_type="website", extra_head=""):
    og = [f'<meta property="og:title" content="{html.escape(title)}">',
          f'<meta property="og:description" content="{html.escape(desc)}">',
          f'<meta property="og:type" content="{og_type}">',
          f'<meta property="og:url" content="{canonical}">',
          '<meta property="og:site_name" content="Rhizvera">']
    if og_image:
        og += [f'<meta property="og:image" content="{og_image}">', '<meta property="og:image:width" content="1200">',
               '<meta property="og:image:height" content="627">', f'<meta property="og:image:alt" content="{html.escape(title)}">',
               '<meta name="twitter:card" content="summary_large_image">']
    return f"""<!DOCTYPE html>
<html lang="en">
<head>
{GTM}
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{html.escape(title)}</title>
<meta name="description" content="{html.escape(desc)}">
<link rel="canonical" href="{canonical}">
{chr(10).join(og)}
{extra_head}<link rel="icon" type="image/png" href="/rhizvera-logo.png">
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Newsreader:opsz,wght@6..72,400;6..72,500&family=Public+Sans:wght@400;500;600&display=swap" rel="stylesheet">
<style>{CSS}</style>
</head>
<body>
<header class="site">
  <div class="wrap">
    <a class="wordmark" href="/"><img src="/rhizvera-logo.png" alt="Rhizvera: Where Data Becomes Truth" width="720" height="171"></a>
    <nav class="header-links" aria-label="Site">
      <a class="header-link" href="/blog/" aria-current="page">Blog</a>
      <a class="header-link" href="https://learn.rhizvera.com/login">Sign in</a>
      <a class="header-link" href="/#request">Request access</a>
    </nav>
  </div>
</header>
<main>
{body}
</main>
<footer class="site">
  <div class="wrap">
    <span>© {datetime.date.today().year} Rhizvera, Inc.</span>
    <span>Staffing Risk Intelligence · <a href="/blog/">Blog</a> · <a href="/privacy.html">Privacy</a></span>
  </div>
</footer>
</body>
</html>
"""

def esc(t): return html.escape(t, quote=False)

def slugify(t):
    t = t.lower().replace("’", "").replace("'", "")
    return re.sub(r"[^a-z0-9]+", "-", t).strip("-")

def read_article(path):
    d = docx.Document(path)
    ps = [p for p in d.paragraphs if p.text.strip()]
    title = next(p.text for p in ps if p.style.name == "Title")
    i = [p.text for p in ps].index(title)
    sub, who = ps[i + 1].text, ps[i + 3].text   # i+2 is the byline
    rest = ps[i + 4:]
    out, lst, section = [], [], "body"
    def flush():
        nonlocal lst
        if lst:
            out.append("<ul>\n" + "\n".join(f"<li>{x}</li>" for x in lst) + "\n</ul>")
            lst = []
    about, copyright_, sources = [], [], []
    for p in rest:
        t, st = p.text.strip(), p.style.name
        if st == "Heading 1" and t == "About Rhizvera":
            flush(); section = "about"; continue
        if st == "Heading 1" and t == "Sources and Notes":
            flush(); section = "sources"; continue
        if section == "about":
            if t == "Rhizvera, Inc.": section = "copyright"; copyright_.append(t); continue
            italic = all(r.italic for r in p.runs if r.text.strip())
            about.append(f'<p class="disclaimer">{esc(t)}</p>' if italic else f"<p>{esc(t)}</p>")
            continue
        if section == "copyright":
            copyright_.append(t); continue
        if section == "sources":
            m = re.match(r"^(.*?):\s+(https?://\S+)$", t)
            sources.append(f'<li>{esc(m.group(1))}: <a href="{html.escape(m.group(2))}">{esc(m.group(2))}</a></li>' if m else f"<p>{esc(t)}</p>")
            continue
        if st == "List Paragraph":
            lst.append(esc(t)); continue
        flush()
        if st == "Heading 1": out.append(f"<h2>{esc(t)}</h2>")
        elif p.runs and all(r.bold for r in p.runs if r.text.strip()): out.append(f'<p class="callout">{esc(t)}</p>')
        else: out.append(f"<p>{esc(t)}</p>")
    flush()
    src_items = [s for s in sources if s.startswith("<li>")]
    src_notes = [s for s in sources if not s.startswith("<li>")]
    src_html = "".join(src_notes) + (("<ul>" + "".join(src_items) + "</ul>") if src_items else "")
    cp = copyright_ + [""] * 3
    return dict(title=title, sub=sub, who=who, body="\n".join(out), about="\n".join(about),
                copyright=f'<p class="co">{esc(cp[0])}</p><p><em>{esc(cp[1])}</em></p><p>{esc(cp[2])}</p>',
                sources=src_html)

def fmt_date(iso):
    d = datetime.date.fromisoformat(iso)
    return f"{d.strftime('%B')} {d.day}, {d.year}"

def article_page(a, slug, date, cover_name):
    url = f"{SITE}/blog/{slug}/"
    cover = (f'<figure class="cover"><img src="{cover_name}" alt="" width="1600" height="900"></figure>' if cover_name
             else '<figure class="cover"><div style="aspect-ratio:16/9;background:var(--sage);border-radius:3px;display:grid;place-items:center;color:var(--muted)">Cover image goes here</div></figure>')
    body = f"""<article>
  <div class="wrap">
    <div class="article-head">
      <p class="eyebrow">Rhizvera Insights</p>
      <h1>{esc(a['title'])}</h1>
      <p class="dek">{esc(a['sub'])}</p>
      <p class="meta">Rhizvera · <time datetime="{date}">{fmt_date(date)}</time></p>
    </div>
    {cover}
    <div class="article-body">
      <p class="who">{esc(a['who'])}</p>
{a['body']}
      <div class="about"><h2>About Rhizvera</h2>
{a['about']}
      </div>
      <div class="copyright">{a['copyright']}</div>
      <div class="sources"><h2>Sources and Notes</h2>{a['sources']}</div>
      <a class="back" href="/blog/">← All articles</a>
    </div>
  </div>
</article>"""
    ld = json.dumps({"@context": "https://schema.org", "@type": "Article", "headline": a["title"], "description": a["sub"],
                     "datePublished": date, "author": {"@type": "Organization", "name": "Rhizvera, Inc."},
                     "publisher": {"@type": "Organization", "name": "Rhizvera, Inc."},
                     **({"image": url + cover_name} if cover_name else {})})
    return page(f"{a['title']} | Rhizvera Insights", a["sub"], body, url,
                og_image=(url + "social.jpg") if cover_name else None, og_type="article",
                extra_head=f'<script type="application/ld+json">{ld}</script>\n')

def index_page(posts, prefix=""):
    if posts:
        items = "\n".join(f"""  <li><a href="{prefix}{p['slug']}/">
    <img src="{prefix}{p['slug']}/{p['cover']}" alt="" width="1600" height="900" loading="lazy">
    <p class="date"><time datetime="{p['date']}">{fmt_date(p['date'])}</time></p>
    <h2>{esc(p['title'])}</h2>
    <p class="dek">{esc(p['sub'])}</p></a></li>""" for p in posts)
        listing = f'<ul class="posts">\n{items}\n</ul>'
    else:
        listing = '<p class="empty">The first articles are on their way.</p>'
    body = f"""<div class="wrap">
  <div class="blog-head">
    <p class="eyebrow">Rhizvera Insights</p>
    <h1>Blog</h1>
    <p>Practical articles on workers' compensation for temporary staffing: classification, payroll and audit, claims procedures, and the records that hold it all together.</p>
  </div>
  {listing}
</div>"""
    return page("Blog | Rhizvera Insights", "Practical articles on workers' compensation for temporary staffing from Rhizvera.", body, f"{SITE}/blog/")

def trim_flat_sides(im, tol=4, min_band=24):
    """Cut a flat, empty band off either side (image generators sometimes leave one)."""
    from PIL import ImageStat
    g = im.convert("L"); w, h = g.size
    flat = lambda x: ImageStat.Stat(g.crop((x, 0, x + 1, h))).stddev[0] < tol
    r = w
    while r > w // 2 and flat(r - 1): r -= 1
    l = 0
    while l < w // 2 and flat(l): l += 1
    l = l if l >= min_band else 0
    r = r if w - r >= min_band else w
    return im.crop((l, 0, r, h)) if (l or r != w) else im

def save_cover(src, dest):
    im = trim_flat_sides(Image.open(src).convert("RGB"))
    w, h = im.size; target = 16 / 9
    if w / h > target: nw = int(h * target); im = im.crop(((w - nw) // 2, 0, (w - nw) // 2 + nw, h))
    elif w / h < target: nh = int(w / target); im = im.crop((0, (h - nh) // 2, w, (h - nh) // 2 + nh))
    im = im.resize((1600, 900), Image.LANCZOS)
    im.save(dest, "JPEG", quality=85, optimize=True, progressive=True)

def save_social(src, dest):
    """LinkedIn link previews are 1.91:1 (1200 x 627); other shapes get padded."""
    im = trim_flat_sides(Image.open(src).convert("RGB"))
    w, h = im.size; target = 1200 / 627
    if w / h > target: nw = int(h * target); im = im.crop(((w - nw) // 2, 0, (w - nw) // 2 + nw, h))
    else: nh = int(w / target); im = im.crop((0, (h - nh) // 2, w, (h - nh) // 2 + nh))
    im.resize((1200, 627), Image.LANCZOS).save(dest, "JPEG", quality=85, optimize=True, progressive=True)

def load():
    return json.load(open(MANIFEST)) if os.path.exists(MANIFEST) else []

def write_index(posts):
    posts = sorted(posts, key=lambda p: p.get("rank", 99))
    posts = sorted(posts, key=lambda p: p["date"], reverse=True)
    open(os.path.join(BLOG, "index.html"), "w", encoding="utf-8").write(index_page(posts))
    json.dump(posts, open(MANIFEST, "w"), indent=2)

def add(docx_path, cover, date, rank=None):
    datetime.date.fromisoformat(date)
    a = read_article(docx_path); slug = slugify(a["title"])
    d = os.path.join(BLOG, slug); os.makedirs(d, exist_ok=True)
    save_cover(cover, os.path.join(d, "cover.jpg"))
    save_social(cover, os.path.join(d, "social.jpg"))
    open(os.path.join(d, "index.html"), "w", encoding="utf-8").write(article_page(a, slug, date, "cover.jpg"))
    posts = [p for p in load() if p["slug"] != slug]
    entry = dict(slug=slug, title=a["title"], sub=a["sub"], date=date, cover="cover.jpg")
    if rank is not None: entry["rank"] = int(rank)
    posts.append(entry)
    write_index(posts); print("published", slug)

def preview(outdir, items):
    os.makedirs(outdir, exist_ok=True); posts = []
    today = datetime.date.today().isoformat()
    for it in items:
        path, _, cover = it.partition("=")
        a = read_article(path); slug = slugify(a["title"]); d = os.path.join(outdir, slug); os.makedirs(d, exist_ok=True)
        cn = None
        if cover: save_cover(cover, os.path.join(d, "cover.jpg")); cn = "cover.jpg"
        open(os.path.join(d, "index.html"), "w", encoding="utf-8").write(article_page(a, slug, today, cn))
        posts.append(dict(slug=slug, title=a["title"], sub=a["sub"], date=today, cover=cn or "cover.jpg"))
    open(os.path.join(outdir, "index.html"), "w", encoding="utf-8").write(index_page(posts, prefix=""))
    print("preview in", outdir)

if __name__ == "__main__":
    cmd = sys.argv[1] if len(sys.argv) > 1 else ""
    if cmd == "add" and len(sys.argv) in (5, 6): add(*sys.argv[2:])
    elif cmd == "index": os.makedirs(BLOG, exist_ok=True); write_index(load())
    elif cmd == "preview": preview(sys.argv[2], sys.argv[3:])
    else: print(__doc__)
