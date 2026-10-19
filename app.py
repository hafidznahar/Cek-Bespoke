"""
Cek Bespoke - BAST Checker & CCTV Checker (versi web, UI v2)
Jalankan:  streamlit run app.py
"""
import hmac
import importlib
import shutil
import tempfile
import zipfile
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st

st.set_page_config(page_title="Cek Bespoke", page_icon="✅", layout="wide")

# ============================================================== STYLE
CSS = """
@import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap');
.stApp, .stApp p, .stApp h1, .stApp h2, .stApp h3, .stApp label, .stApp li, .stApp button, .stApp textarea, .stApp input
{ font-family: 'Plus Jakarta Sans', system-ui, sans-serif; }
.stApp { background:#F3F5F9; color:#14202E; }
[data-testid="stAppDeployButton"], .stDeployButton { display:none; }
[data-testid="stHeader"] { background:transparent; }
.block-container { padding:2.2rem 2.8rem 3rem; max-width:1240px; }

/* sidebar */
section[data-testid="stSidebar"] { background:#14202E; }
section[data-testid="stSidebar"] * { color:#E8EDF3; }
.brand { display:flex; align-items:center; gap:10px; font-weight:800; font-size:19px; margin:6px 4px 22px; color:#fff; }
.logo { width:34px; height:34px; border-radius:10px; background:#3FA7D6; display:grid; place-items:center; }
[data-testid="stSidebar"] [role="radiogroup"] { gap:6px; }
[data-testid="stSidebar"] [role="radiogroup"] label { padding:12px 14px; border-radius:12px; width:100%; cursor:pointer; }
[data-testid="stSidebar"] [role="radiogroup"] label > div:first-child { display:none; }
[data-testid="stSidebar"] [role="radiogroup"] label:has(input:checked) { background:#243A52; }
[data-testid="stSidebar"] [role="radiogroup"] label p { font-weight:600; font-size:15px; }
.side-note { margin-top:28px; font-size:12px; line-height:1.5; color:#A9B6C4 !important; }

/* header & stepper */
.page-title { font-size:30px; font-weight:800; margin:0; color:#14202E; }
.page-sub { margin:6px 0 18px; color:#4A5B6D; font-size:15px; }
.steps { display:flex; gap:10px; flex-wrap:wrap; margin-bottom:20px; }
.step { display:flex; align-items:center; gap:10px; padding:8px 16px 8px 8px; border-radius:99px; background:#fff;
        border:1px solid #D9E0E8; font-size:13px; font-weight:600; color:#4A5B6D; }
.step b { width:26px; height:26px; border-radius:50%; background:#E3E9F0; display:grid; place-items:center; font-size:12px; color:#14202E; }
.step.on { border-color:#1F4E78; color:#1F4E78; }
.step.on b, .step.done b { background:#1F4E78; color:#fff; }
.step.done { color:#1F4E78; }

/* cards */
.card { background:#fff; border:1px solid #D9E0E8; border-radius:18px; padding:22px 24px; margin-bottom:14px; }
.card h4 { margin:0 0 12px; font-size:17px; font-weight:800; }
.how { display:flex; gap:12px; font-size:14px; line-height:1.5; margin-bottom:10px; }
.how b { color:#1F4E78; }
.file-count { font-weight:800; font-size:17px; }
.file-row { margin-top:10px; padding:12px 14px; background:#F3F5F9; border-radius:12px; font-size:14px; word-break:break-all; }
.muted { color:#4A5B6D; font-size:14px; }
.banner { border-radius:14px; padding:16px 20px; font-weight:700; font-size:15px; margin-bottom:16px; }
.banner.bad { background:#FDECEA; color:#8C1D18; }
.banner.ok { background:#E6F4EA; color:#14532D; }
.banner.warn { background:#FFF4DB; color:#7A4B00; }
.m-card { background:#fff; border:1px solid #D9E0E8; border-radius:18px; padding:18px 22px; }
.m-card.bad { background:#FFF8F7; border-color:#E8A9A4; }
.m-card.bad .m-l, .m-card.bad .m-v { color:#8C1D18; }
.m-l { color:#4A5B6D; font-size:13px; font-weight:700; }
.m-v { font-size:36px; font-weight:800; line-height:1.2; }
.h-sec { font-weight:800; font-size:18px; margin:22px 0 10px; }
.doc-name { font-size:13px; font-weight:700; word-break:break-all; margin-top:6px; }

/* dropzone */
[data-testid="stFileUploader"] section { border:2px dashed #7FA6C9; background:#F7FAFD; border-radius:18px; padding:38px 24px;
  flex-direction:column; align-items:center; justify-content:center; text-align:center; gap:6px; }
[data-testid="stFileUploaderDropzoneInstructions"] > div > * { display:none; }
[data-testid="stFileUploaderDropzoneInstructions"] > div::before { content:"Tarik dan lepas file di sini"; display:block; font-size:20px; font-weight:800; color:#14202E; }
[data-testid="stFileUploaderDropzoneInstructions"] > div::after { content:"atau klik tombol di bawah untuk memilih file"; display:block; font-size:14px; color:#4A5B6D; margin-top:4px; }
[data-testid="stFileUploaderDropzone"] button { font-size:0; background:#1F4E78; border:0; border-radius:12px; padding:12px 28px; min-height:46px; margin-top:10px; }
[data-testid="stFileUploaderDropzone"] button::after { content:"Pilih file"; font-size:15px; font-weight:700; color:#fff; }
[data-testid="stFileUploaderFile"], [data-testid="stFileUploaderPagination"], [data-testid="stFileUploaderFileList"],
[data-testid="stFileUploader"] section + div { display:none; }

/* buttons */
[data-testid="stButton"] button, [data-testid="stDownloadButton"] button { border:1.5px solid #1F4E78; color:#1F4E78; background:#fff;
  border-radius:12px; font-weight:700; min-height:48px; padding:.6rem 1.4rem; }
[data-testid="stButton"] button p, [data-testid="stDownloadButton"] button p { color:inherit; font-weight:700; font-size:15px; }
[data-testid="stButton"] button[kind="primary"], [data-testid="stDownloadButton"] button[kind="primary"], [data-testid="stBaseButton-primary"]
{ background:#1F4E78; color:#fff; border-color:#1F4E78; }
[data-testid="stButton"] button:disabled { opacity:.5; }
"""
st.markdown(f"<style>{CSS}</style>", unsafe_allow_html=True)

# ============================================================== KONFIGURASI ALAT
TOOLS = {
    "bast": dict(
        core="bast_core", runner="run_bast_check", title="BAST Checker",
        desc="Deteksi foto yang sama atau mirip di antara banyak dokumen BAST. Logo dan kop surat otomatis diabaikan.",
        types=["pdf", "docx", "zip"], exts=(".pdf", ".docx"),
        run_label="Mulai Cek BAST",
        steps=["Upload file BAST, satu per satu atau sekaligus lewat ZIP.",
               "Klik <b>Mulai Cek BAST</b> dan tunggu progress selesai.",
               "Unduh laporan Excel dan foto duplikatnya."],
    ),
    "cctv": dict(
        core="cctv_core", runner="run_cctv_check", title="CCTV Checker",
        desc="Cek screenshot CCTV: kamera burem, pecah, pucat, silau, gelap, dan kamera fisheye. Hasil OK atau N/A per screenshot.",
        types=["pptx", "pdf", "zip"], exts=(".pptx", ".pdf"),
        run_label="Mulai Cek CCTV",
        steps=["Upload file PPTX / PDF CCTV, atau 1 file ZIP berisi folder-folder dokumen.",
               "Klik <b>Mulai Cek CCTV</b> dan tunggu progress selesai.",
               "Unduh laporan Excel. Screenshot N/A juga tampil di halaman hasil."],
    ),
}

# ============================================================== HELPER
def stamp():
    return datetime.now().strftime("%Y%m%d_%H%M")


def fmt_int(n):
    return f"{n:,}".replace(",", ".")


def fmt_size(nbytes):
    mb = nbytes / (1024 * 1024)
    return f"{mb:,.1f} MB".replace(",", ".") if mb < 1024 else f"{mb / 1024:.1f} GB"


def stepper_html(active):
    names = ["Upload file", "Proses", "Unduh hasil"]
    out = []
    for i, n in enumerate(names, 1):
        cls = "step on" if i == active else ("step done" if i < active else "step")
        out.append(f'<span class="{cls}"><b>{i}</b>{n}</span>')
    return '<div class="steps">' + "".join(out) + "</div>"


def metric_html(label, value, tone=""):
    return f'<div class="m-card {tone}"><div class="m-l">{label}</div><div class="m-v">{value}</div></div>'


def save_uploads(uploaded_files, dest: Path, allowed_ext):
    """Simpan upload ke folder sementara. ZIP diekstrak (struktur subfolder dipertahankan)."""
    dest.mkdir(parents=True, exist_ok=True)
    dest_root = dest.resolve()
    n_saved = 0

    def unique_path(p: Path) -> Path:
        if not p.exists():
            return p
        i = 2
        while True:
            cand = p.with_name(f"{p.stem}_{i}{p.suffix}")
            if not cand.exists():
                return cand
            i += 1

    for uf in uploaded_files:
        name = Path(uf.name).name
        if name.lower().endswith(".zip"):
            with zipfile.ZipFile(uf) as z:
                for info in z.infolist():
                    if info.is_dir():
                        continue
                    fixed = info.filename
                    if not (info.flag_bits & 0x800):
                        try:
                            fixed = info.filename.encode("cp437").decode("utf-8")
                        except Exception:
                            pass
                    rel = Path(fixed.replace("\\", "/"))
                    if rel.name.startswith("~$") or "__MACOSX" in rel.parts:
                        continue
                    if rel.suffix.lower() not in allowed_ext:
                        continue
                    target = (dest / rel).resolve()
                    if dest_root not in target.parents:  # cegah zip-slip
                        continue
                    target = unique_path(target)
                    target.parent.mkdir(parents=True, exist_ok=True)
                    with z.open(info) as src, open(target, "wb") as out:
                        shutil.copyfileobj(src, out)
                    n_saved += 1
        elif Path(name).suffix.lower() in allowed_ext:
            target = unique_path(dest / name)
            target.write_bytes(uf.getbuffer())
            n_saved += 1
    return n_saved


def reset_tool(key, clear_result):
    st.session_state[f"{key}_n"] = st.session_state.get(f"{key}_n", 0) + 1  # ganti key uploader = kosongkan
    if clear_result:
        st.session_state.pop(f"{key}_result", None)


def run_tool(key, files):
    """Proses file. Return True kalau sukses (hasil disimpan di session_state)."""
    cfg = TOOLS[key]
    try:
        core = importlib.import_module(cfg["core"])
    except ImportError as e:
        st.error(f"Library belum terpasang: {e}. Jalankan: pip install -r requirements.txt")
        return False
    if key == "bast" and not core.pdf_support_available():
        st.warning("PyMuPDF belum terpasang, file PDF tidak akan terbaca (DOCX tetap jalan).")

    work = Path(tempfile.mkdtemp(prefix=f"{key}_"))
    bar = st.progress(0.0)
    txt = st.empty()

    def cb(frac, msg):
        bar.progress(min(max(frac, 0.0), 1.0))
        txt.caption(msg)

    try:
        cb(0.01, "Menyimpan file upload...")
        n = save_uploads(files, work / "input", cfg["exts"])
        if n == 0:
            st.error("Tidak ada file yang valid di upload kamu. Pastikan formatnya sesuai.")
            return False
        res = getattr(core, cfg["runner"])(work / "input", work, progress=cb)
        res["stamp"] = stamp()
        st.session_state[f"{key}_result"] = res
        return True
    except ValueError as e:
        st.error(str(e))
        return False
    except Exception as e:
        st.error(f"Terjadi error saat memproses: {e}")
        return False
    finally:
        shutil.rmtree(work, ignore_errors=True)
        bar.empty()
        txt.empty()


# ============================================================== TAMPILAN HASIL
def render_bast(res):
    if res["n_groups"] == 0:
        st.markdown('<div class="banner ok">Tidak ditemukan foto duplikat.</div>', unsafe_allow_html=True)
    else:
        st.markdown(f'<div class="banner bad">Ditemukan {res["n_groups"]} grup foto duplikat. '
                    'Periksa detailnya di bawah atau di laporan Excel.</div>', unsafe_allow_html=True)

    cols = st.columns(4)
    cols[0].markdown(metric_html("Dokumen dibaca", fmt_int(res["n_docs"])), unsafe_allow_html=True)
    cols[1].markdown(metric_html("Foto valid", fmt_int(res["n_photos"])), unsafe_allow_html=True)
    cols[2].markdown(metric_html("Grup duplikat", fmt_int(res["n_groups"]), "bad" if res["n_groups"] else ""), unsafe_allow_html=True)
    cols[3].markdown(metric_html("Foto terlibat duplikat", fmt_int(res["n_dup_photos"])), unsafe_allow_html=True)

    for w in res["warnings"]:
        st.warning(w)

    st.write("")
    b1, b2, _, b4 = st.columns([1.1, 1.3, 1.2, 0.9])
    b1.download_button("Unduh laporan Excel", res["excel_bytes"], type="primary",
                       file_name=f"hasil_laporan_bast_{res['stamp']}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="bast_dl_x")
    if res["n_groups"]:
        b2.download_button("Unduh foto duplikat (ZIP)", res["dup_zip_bytes"],
                           file_name=f"foto_duplikat_{res['stamp']}.zip", mime="application/zip", key="bast_dl_z")
    b4.button("Cek file lain", on_click=reset_tool, args=("bast", True), key="bast_again")

    if res["preview"]:
        extra = f" · menampilkan {len(res['preview'])} dari {res['n_groups']}" if res["n_groups"] > len(res["preview"]) else ""
        st.markdown(f'<div class="h-sec">Preview grup duplikat<span class="muted">{extra}</span></div>', unsafe_allow_html=True)
        for i, g in enumerate(res["preview"]):
            holder = st.container(border=True) if i < 2 else st.expander(f"Grup #{g['group_id']} · {g['size']} foto")
            with holder:
                if i < 2:
                    st.markdown(f"**Grup #{g['group_id']} · {g['size']} foto**")
                icols = st.columns(len(g["items"]))
                for col, (doc, fname, img) in zip(icols, g["items"]):
                    col.image(img)
                    col.markdown(f'<div class="doc-name">{doc}</div>', unsafe_allow_html=True)
                    col.caption(fname)


def render_cctv(res):
    if res["n_na"] == 0:
        st.markdown('<div class="banner ok">Semua screenshot berstatus OK.</div>', unsafe_allow_html=True)
    else:
        st.markdown(f'<div class="banner warn">{res["n_na"]} dari {res["n_total"]} screenshot berstatus N/A. '
                    'Periksa detailnya di bawah atau di laporan Excel.</div>', unsafe_allow_html=True)

    cols = st.columns(4)
    cols[0].markdown(metric_html("Dokumen dibaca", fmt_int(res["n_docs"])), unsafe_allow_html=True)
    cols[1].markdown(metric_html("Screenshot dicek", fmt_int(res["n_total"])), unsafe_allow_html=True)
    cols[2].markdown(metric_html("OK", fmt_int(res["n_ok"])), unsafe_allow_html=True)
    cols[3].markdown(metric_html("N/A", fmt_int(res["n_na"]), "bad" if res["n_na"] else ""), unsafe_allow_html=True)

    for w in res["warnings"]:
        st.warning(w)

    st.write("")
    b1, _, b3 = st.columns([1.1, 2.4, 0.9])
    b1.download_button("Unduh laporan Excel", res["excel_bytes"], type="primary",
                       file_name=f"hasil_qc_cctv_{res['stamp']}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="cctv_dl_x")
    b3.button("Cek file lain", on_click=reset_tool, args=("cctv", True), key="cctv_again")

    st.markdown('<div class="h-sec">Detail per screenshot</div>', unsafe_allow_html=True)
    only_na = st.checkbox("Tampilkan hanya yang N/A", value=True, key="cctv_only_na")
    df = pd.DataFrame(res["summary"])
    view = df[df["Status Akhir"] == "N/A"] if only_na else df
    st.dataframe(view, hide_index=True)

    if res["preview_na"]:
        extra = f" · menampilkan {len(res['preview_na'])} dari {res['n_na']}" if res["n_na"] > len(res["preview_na"]) else ""
        st.markdown(f'<div class="h-sec">Preview screenshot N/A<span class="muted">{extra}</span></div>', unsafe_allow_html=True)
        cols = st.columns(4)
        for i, item in enumerate(res["preview_na"]):
            with cols[i % 4]:
                st.image(item["bytes"])
                st.markdown(f'<div class="doc-name">{item["caption"]}</div>', unsafe_allow_html=True)
                st.caption(item["note"])


RENDERERS = {"bast": render_bast, "cctv": render_cctv}

# ============================================================== HALAMAN
def tool_page(key):
    cfg = TOOLS[key]
    st.markdown(f'<h1 class="page-title">{cfg["title"]}</h1><p class="page-sub">{cfg["desc"]}</p>', unsafe_allow_html=True)

    res = st.session_state.get(f"{key}_result")
    if res:
        st.markdown(stepper_html(3), unsafe_allow_html=True)
        RENDERERS[key](res)
        return

    slot = st.empty()
    slot.markdown(stepper_html(1), unsafe_allow_html=True)

    left, right = st.columns([3, 1.1], gap="large")
    with left:
        n = st.session_state.get(f"{key}_n", 0)
        files = st.file_uploader("Upload file", type=cfg["types"], accept_multiple_files=True,
                                 key=f"{key}_up_{n}", label_visibility="collapsed")
        if files:
            total = sum(f.size for f in files)
            c1, c2 = st.columns([4, 1])
            c1.markdown(
                f'<div class="card"><div class="file-count">{len(files)} file siap dicek '
                f'<span class="muted">· {fmt_size(total)}</span></div>'
                f'<div class="file-row">{files[0].name}</div>'
                + (f'<div class="muted" style="margin-top:8px">+ {len(files) - 1} file lainnya</div>' if len(files) > 1 else "")
                + "</div>", unsafe_allow_html=True)
            c2.button("Hapus semua", on_click=reset_tool, args=(key, False), key=f"{key}_clear")

        go = st.button(cfg["run_label"], type="primary", disabled=not files, key=f"{key}_go")
        if go:
            slot.markdown(stepper_html(2), unsafe_allow_html=True)
            if run_tool(key, files):
                st.rerun()
            slot.markdown(stepper_html(1), unsafe_allow_html=True)

    with right:
        how = "".join(f'<div class="how"><b>{i}</b><span>{s}</span></div>' for i, s in enumerate(cfg["steps"], 1))
        st.markdown(f'<div class="card"><h4>Cara pakai</h4>{how}</div>', unsafe_allow_html=True)


def guide_page():
    st.markdown('<h1 class="page-title">Panduan</h1><p class="page-sub">Tips singkat supaya pengecekan lancar.</p>', unsafe_allow_html=True)
    st.markdown("""
<div class="card"><h4>Upload banyak file sekaligus</h4>
<div class="muted">Kalau file tersebar di banyak folder (misal Kalimantan, Sulawesi, Sumbagsel), jadikan <b>satu file ZIP</b> lalu upload.
Nama subfolder akan ikut tampil di laporan.</div></div>
<div class="card"><h4>Cara membaca hasil</h4>
<div class="muted"><b>BAST:</b> setiap grup duplikat berisi foto yang sama atau mirip di beberapa dokumen. Cek manual untuk memastikan.<br>
<b>CCTV:</b> status N/A artinya ada masalah kualitas (burem, pecah, pucat, silau, gelap) atau kamera fisheye. Alasannya tertulis di kolom Catatan.</div></div>
<div class="card"><h4>Kalau terasa lambat atau error</h4>
<div class="muted">Upload per batch yang lebih kecil, tutup aplikasi berat lain, dan pastikan koneksi tidak terputus selama proses.
Hasil perlu tetap diverifikasi manual untuk kasus yang meragukan.</div></div>
""", unsafe_allow_html=True)


# ============================================================== KATA SANDI (opsional)
def password_gate():
    """Aktif hanya kalau secret APP_PASSWORD diisi (di hosting). Kalau kosong (jalan lokal), tidak ada kata sandi."""
    try:
        expected = st.secrets.get("APP_PASSWORD", None)
    except Exception:
        expected = None
    if not expected or st.session_state.get("auth_ok"):
        return
    st.markdown('<h1 class="page-title">Cek Bespoke</h1>'
                '<p class="page-sub">Masukkan kata sandi untuk melanjutkan.</p>', unsafe_allow_html=True)
    pw = st.text_input("Kata sandi", type="password", key="pw_input")
    if st.button("Masuk", type="primary", key="pw_go"):
        if hmac.compare_digest(str(pw), str(expected)):
            st.session_state["auth_ok"] = True
            st.rerun()
        else:
            st.error("Kata sandi salah.")
    st.stop()


# ============================================================== SIDEBAR & ROUTING
password_gate()

with st.sidebar:
    st.markdown(
        '<div class="brand"><span class="logo"><svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#fff" '
        'stroke-width="3" stroke-linecap="round" stroke-linejoin="round"><path d="M5 12l5 5 9-10"/></svg></span>Cek Bespoke</div>',
        unsafe_allow_html=True)
    page = st.radio("Menu", ["BAST Checker", "CCTV Checker", "Panduan"], label_visibility="collapsed")
    st.markdown('<div class="side-note">File diproses sementara dan dihapus setelah selesai. '
                'Kasus meragukan tetap perlu dicek manual.</div>', unsafe_allow_html=True)

if page == "BAST Checker":
    tool_page("bast")
elif page == "CCTV Checker":
    tool_page("cctv")
else:
    guide_page()
