# -*- coding: utf-8 -*-
r"""
generate_pdfs.py
- Injects print CSS into a temp copy of each HTML (forces .rv content visible)
- Uses Chrome headless to render each to PDF
- Combines each batch into one PDF with pypdf

Output: C:/Users/HP/Downloads/EPP_Pdf/
Batches: 26_50_EPP.pdf, 51_75_EPP.pdf  (add more to BATCHES as links arrive)
"""

import os, subprocess, sys, io, tempfile
from pypdf import PdfWriter, PdfReader
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

CHROME      = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
# Search ready_to_upload first (has GDrive URLs), fall back to html_pages
BASE        = r"D:\Claude_Code\Ekta_Diwas_website\Ektadiwas_site\Ekta_Prakash_Parv_QR\EPP2026_GIPL_Pages"
READY_DIR   = os.path.join(BASE, "ready_to_upload")
HTML_DIR    = os.path.join(BASE, "html_pages")
OUT_DIR     = r"C:\Users\HP\Downloads\EPP_Pdf"
TEMP_DIR    = os.path.join(OUT_DIR, "_temp")
os.makedirs(TEMP_DIR, exist_ok=True)

# CSS injected before </head> — forces all content visible for printing
PRINT_CSS = """
<style id="epp-print">
/* ── Force all reveal-on-scroll content visible ── */
.rv { opacity: 1 !important; transform: none !important; transition: none !important; }

@media print {
  /* Remove interactive/animated elements */
  * { animation: none !important; transition: none !important; }
  #fx { display: none !important; }          /* particle canvas */
  .top { display: none !important; }         /* fixed nav bar */
  .lights { display: none !important; }      /* string lights */
  .scroll { display: none !important; }      /* scroll indicator */
  .cta-row { display: none !important; }     /* CTA buttons */
  dialog { display: none !important; }       /* zoom dialog */

  /* Hero: allow natural height so image shows */
  .hero { min-height: 60vh !important; page-break-after: avoid; }
  .hero-img img { transform: none !important; }
  .hero::after { display: none !important; }

  /* Spinning border on image frame — remove */
  .frame::before { display: none !important; }
  .frame { box-shadow: none !important; }

  /* Content sections — no page break inside */
  .wrap, .story, .hl, .text, .frame, .next, footer { break-inside: avoid; }

  /* Colours — keep the warm background */
  body { background: #fff8ee !important; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
}
</style>
"""

BATCHES = [
    ("76_100_EPP",  list(range(76,  101))),
    ("101_125_EPP", list(range(101, 126))),   # Sr 114 inactive — skipped automatically
    ("126_152_EPP", list(range(126, 153))),
]

def html_file_for_sr(sr):
    """Find HTML file: prefer ready_to_upload (GDrive URLs), fall back to html_pages."""
    prefix = f"EPP2026_{sr:03d}_"
    for folder in (READY_DIR, HTML_DIR):
        for fname in os.listdir(folder):
            if fname.startswith(prefix) and fname.endswith(".html"):
                return os.path.join(folder, fname)
    return None

def make_printable_copy(html_path):
    """Inject print CSS and save to a temp file; return temp file path."""
    with open(html_path, encoding="utf-8") as f:
        html = f.read()
    # Inject before </head>
    html = html.replace("</head>", PRINT_CSS + "\n</head>", 1)
    tmp = tempfile.NamedTemporaryFile(
        suffix=".html", dir=TEMP_DIR, delete=False, mode="w", encoding="utf-8"
    )
    tmp.write(html)
    tmp.close()
    return tmp.name

def print_to_pdf(html_path, out_pdf):
    file_url = "file:///" + html_path.replace("\\", "/")
    cmd = [
        CHROME,
        "--headless=new",
        "--disable-gpu",
        "--no-sandbox",
        "--disable-extensions",
        f"--print-to-pdf={out_pdf}",
        "--no-pdf-header-footer",
        "--run-all-compositor-stages-before-draw",
        "--virtual-time-budget=15000",   # 15 s virtual time for images
        file_url,
    ]
    try:
        subprocess.run(cmd, capture_output=True, timeout=120)
    except subprocess.TimeoutExpired:
        print("(timeout — retrying without virtual-time-budget) ", end="", flush=True)
        # Retry without virtual-time-budget for pages that hang
        cmd2 = [c for c in cmd if "virtual-time-budget" not in c]
        try:
            subprocess.run(cmd2, capture_output=True, timeout=90)
        except subprocess.TimeoutExpired:
            return False
    return os.path.exists(out_pdf) and os.path.getsize(out_pdf) > 1024

# ── Main loop ────────────────────────────────────────────────────────────────
for batch_name, sr_list in BATCHES:
    print(f"\n{'='*60}")
    print(f"Batch: {batch_name}  ({len(sr_list)} exhibits)")
    print(f"{'='*60}")

    writer   = PdfWriter()
    tmp_html_files = []
    tmp_pdf_files  = []
    ok = fail = skip = 0

    for sr in sr_list:
        html = html_file_for_sr(sr)
        if not html:
            print(f"  SKIP  Sr {sr:03d}  (not in ready_to_upload)")
            skip += 1
            continue

        title    = os.path.basename(html).replace(".html", "")
        temp_htm = make_printable_copy(html)
        temp_pdf = os.path.join(TEMP_DIR, f"{sr:03d}.pdf")
        tmp_html_files.append(temp_htm)

        print(f"  Sr {sr:03d}  {title} ...", end=" ", flush=True)
        success = print_to_pdf(temp_htm, temp_pdf)

        if success:
            reader = PdfReader(temp_pdf)
            for page in reader.pages:
                writer.add_page(page)
            tmp_pdf_files.append(temp_pdf)
            print(f"OK ({len(reader.pages)}p)")
            ok += 1
        else:
            print("FAIL")
            fail += 1

    if ok > 0:
        out_path = os.path.join(OUT_DIR, f"{batch_name}.pdf")
        with open(out_path, "wb") as f:
            writer.write(f)
        size_mb = os.path.getsize(out_path) / (1024 * 1024)
        print(f"\nSaved: {out_path}  ({size_mb:.1f} MB)")
    else:
        print("\nNo pages rendered — PDF not saved.")

    # Clean up temp files
    for t in tmp_html_files + tmp_pdf_files:
        try:
            os.remove(t)
        except Exception:
            pass

    print(f"Done: {ok} OK,  {fail} failed,  {skip} skipped")

try:
    os.rmdir(TEMP_DIR)
except Exception:
    pass

print("\nAll batches complete.")
