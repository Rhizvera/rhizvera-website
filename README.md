# rhizvera.com

Public website for Rhizvera, Inc. It holds public text only: no SRI platform code, data or confidential material.

## Blog (Rhizvera Insights)

The blog lives in `blog/` and is built by `scripts/build_blog.py` from each article's Word file. Only published
posts are ever written to this repository (it is public), so unpublished drafts never appear here.

    python scripts/build_blog.py add <article.docx> <cover image> <YYYY-MM-DD>

builds `blog/<slug>/index.html` and `blog/<slug>/cover.jpg` (cropped to 16:9, 1600 x 900), and updates
`blog/index.html` and `blog/posts.json`. Needs Python with `python-docx` and `Pillow`.
