import streamlit as st
import os
import re
import time
from io import BytesIO

# --- IMPORT ENGINE ---
from engine import PDFConverterEngine
from engine2 import DocxOptimizerEngine
from engine3 import DocxTranslatorEngine
from engine4 import CoverPageEngine
from engine5 import DaftarIsiEngine
from engine6 import PrakataPendahuluanEngine
from engine7 import InfoPendukungEngine
from engine9 import DocxFinalTranslatorEngine, CustomDictionary

# --- KONFIGURASI HALAMAN ---
st.set_page_config(
    page_title="ISO Doc Master",
    page_icon="📑",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- CSS CUSTOM ---
st.markdown("""
    <style>
    .stButton>button {width: 100%; border-radius: 5px; height: 3em;}
    .success-box {padding: 1rem; background-color: #d4edda; border-radius: 5px; color: #155724; margin-bottom: 1rem;}
    </style>
""", unsafe_allow_html=True)

# --- KONSTANTA STANDAR ISO (HARDCODED) ---
ISO_FONT_NAME = "Arial"
ISO_FONT_SIZE = 11


# ─────────────────────────────────────────────────────────────
# AUTO EKSTRAK JUDUL DARI DOKUMEN
# Strategi:
#   1. Cari paragraf dengan font size ≥ 14pt atau heading style
#   2. Judul Bahasa Indonesia  → baris pertama yang memenuhi kriteria
#   3. Judul Bahasa Inggris   → baris kedua yang memenuhi kriteria
#      (atau baris pertama yang terdeteksi italic / font size lebih kecil)
# ─────────────────────────────────────────────────────────────
def extract_titles_from_docx(docx_path: str):
    """
    Mengekstrak judul Bahasa Indonesia dan Bahasa Inggris dari dokumen .docx.

    Returns:
        (title_id, title_en) – keduanya string, bisa kosong jika tidak ditemukan.
    """
    try:
        from docx import Document
        from docx.shared import Pt

        doc = Document(docx_path)
        candidates = []

        for para in doc.paragraphs:
            text = para.text.strip()
            if not text or len(text) < 5:
                continue

            # Cek apakah paragraf menggunakan heading style
            is_heading = para.style.name.lower().startswith('heading')

            # Cek font size dari run pertama yang punya teks
            max_size = 0
            is_italic = False
            for run in para.runs:
                if run.text.strip():
                    sz = run.font.size
                    if sz:
                        max_size = max(max_size, sz.pt if hasattr(sz, 'pt') else sz / 12700)
                    if run.font.italic:
                        is_italic = True

            # Fallback: cek dari style paragraf jika run tidak punya size
            if max_size == 0 and para.style.font.size:
                max_size = para.style.font.size.pt

            # Kandidat judul: heading style ATAU font ≥ 13pt
            if is_heading or max_size >= 12:
                candidates.append({
                    'text': text,
                    'size': max_size,
                    'italic': is_italic,
                    'heading': is_heading,
                })

            # Ambil maksimal 10 kandidat pertama (area awal dokumen)
            if len(candidates) >= 10:
                break

        if not candidates:
            return "", ""

        # Judul ID → kandidat pertama (biasanya lebih besar, tidak italic)
        title_id = candidates[0]['text']

        # Judul EN → sama dengan judul ID (diambil dari paragraf judul pertama)
        title_en = title_id

        return title_id, title_en

    except Exception as e:
        return "", ""

# --- INISIALISASI ENGINE ---
@st.cache_resource
def load_engines():
    # Sesuaikan path Tesseract Anda di sini
    e1 = PDFConverterEngine(tesseract_path=r'C:\Program Files\Tesseract-OCR\tesseract.exe')
    e2 = DocxOptimizerEngine()
    e3 = DocxTranslatorEngine()
    e4 = CoverPageEngine()  # Engine Cover/Sampul
    e5 = DaftarIsiEngine()  # Engine Daftar Isi
    e6 = PrakataPendahuluanEngine()  # Engine Prakata & Pendahuluan
    e7 = InfoPendukungEngine()       # Engine Info Pendukung Perumus
    e8 = DocxFinalTranslatorEngine() # Engine 9: Terjemahan Final + Custom Dictionary
    return e1, e2, e3, e4, e5, e6, e7, e8

engine1, engine2, engine3, engine4, engine5, engine6, engine7, engine8 = load_engines()

# --- HELPER FUNCTION: AUTO CONVERT PDF ---
def handle_pdf_conversion(uploaded_file):
    """Fungsi pembantu untuk mengubah PDF ke Word di background"""
    temp_pdf = f"temp_{uploaded_file.name}"
    temp_docx = temp_pdf.replace(".pdf", ".docx")
    
    with open(temp_pdf, "wb") as f:
        f.write(uploaded_file.getbuffer())
    
    # Deteksi Scan
    is_scan = engine1.is_scanned_pdf(temp_pdf)
    msg = "🔍 PDF Scan (OCR)" if is_scan else "⚡ PDF Digital"
    
    # Konversi
    success, error = engine1.convert(temp_pdf, temp_docx)
    
    # Bersihkan PDF mentah
    if os.path.exists(temp_pdf): os.remove(temp_pdf)
    
    return success, temp_docx, msg, error

# --- SIDEBAR ---
st.sidebar.title("🎛️ Control Panel")
menu = st.sidebar.radio(
    "Pilih Mode:", 
    ["1. Konversi (PDF -> Word)", "2. Rapikan (Word -> ISO Std)", "3. Terjemahkan (EN -> ID + Rapikan)", "4. Terjemahkan Dokumen Final (→ Bahasa Indonesia)"]
)

# ── PANEL KAMUS (Sidebar) ─────────────────────────────────────────────────────
st.sidebar.divider()
st.sidebar.subheader("📚 Kamus Istilah (Engine 9)")
st.sidebar.caption("Istilah di kamus diprioritaskan sebelum Google Translate.")

# Inisialisasi kamus di session state
if 'custom_dict' not in st.session_state:
    st.session_state['custom_dict'] = None

# Load kamus bawaan
if st.sidebar.checkbox("Aktifkan kamus bawaan (±60 istilah teknik)", value=True, key="chk_defaults"):
    if st.session_state['custom_dict'] is None:
        d = CustomDictionary()
        d.load_defaults()
        st.session_state['custom_dict'] = d
else:
    st.session_state['custom_dict'] = None

current_dict = st.session_state['custom_dict']

# ── Google Sheet ────────────────────────────────────────────────────────────
with st.sidebar.expander("🔗 Kamus dari Google Sheet", expanded=False):
    st.caption(
        "Tempel link Google Sheet di sini. "
        "Sheet harus diset **Anyone with the link → Viewer**."
    )
    gs_url = st.text_input(
        "URL Google Sheet",
        key="gs_url_input",
        placeholder="https://docs.google.com/spreadsheets/d/...",
    )
    if gs_url:
        st.session_state['gs_url_saved'] = gs_url
    saved_url = st.session_state.get('gs_url_saved', '')

    col_gs1, col_gs2 = st.columns(2)
    with col_gs1:
        btn_load_gs = st.button("📥 Muat", key="btn_load_gs",
                                 use_container_width=True, disabled=not saved_url)
    with col_gs2:
        btn_refresh_gs = st.button("🔄 Refresh", key="btn_refresh_gs",
                                    use_container_width=True, disabled=not saved_url,
                                    help="Ambil ulang data terbaru dari Google Sheet (tanpa hapus kamus lain)")

    if (btn_load_gs or btn_refresh_gs) and saved_url:
        with st.spinner("Mengambil data dari Google Sheets..."):
            try:
                _d = st.session_state.get('custom_dict') or CustomDictionary()
                if btn_refresh_gs:
                    # Refresh: reload defaults + ambil ulang dari sheet
                    _d.clear()
                    if st.session_state.get('chk_defaults', True):
                        _d.load_defaults()
                n = _d.load_from_google_sheet(saved_url)
                st.session_state['custom_dict'] = _d
                action = "Di-refresh" if btn_refresh_gs else "Dimuat"
                st.success(f"✅ {action}: **{n} istilah** dari Google Sheet.")
            except ConnectionError as ce:
                st.error(f"❌ Gagal terhubung:\n\n{ce}")
            except ValueError as ve:
                st.error(f"❌ Format kolom tidak dikenali:\n\n{ve}")
            except Exception as ex:
                st.error(f"❌ Error: {ex}")

    if saved_url:
        st.caption(f"🔗 `{saved_url[:55]}{'...' if len(saved_url)>55 else ''}`")

    st.markdown("---")
    st.markdown("""
**Cara setup Google Sheet:**
1. Buat sheet dengan header `source` | `target`
2. Isi istilah asing & terjemahannya
3. **Share → Anyone with the link → Viewer**
4. Paste link di atas, klik **Muat**
5. Jika sheet diupdate → klik **Refresh**
""")

# Upload CSV
kamus_csv = st.sidebar.file_uploader("Upload Kamus CSV", type=["csv"], key="upload_csv")
if kamus_csv:
    import tempfile, os as _os
    with tempfile.NamedTemporaryFile(delete=False, suffix=".csv") as tmp:
        tmp.write(kamus_csv.getbuffer())
        tmp_path = tmp.name
    try:
        if current_dict is None:
            current_dict = CustomDictionary()
        n = current_dict.load_from_csv(tmp_path)
        st.session_state['custom_dict'] = current_dict
        st.sidebar.success(f"✅ {n} istilah dari CSV dimuat.")
    except Exception as e:
        st.sidebar.error(f"Gagal: {e}")
    finally:
        _os.unlink(tmp_path)

# Upload Excel
kamus_xlsx = st.sidebar.file_uploader("Upload Kamus Excel (.xlsx)", type=["xlsx","xls"], key="upload_xlsx")
if kamus_xlsx:
    import tempfile, os as _os
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        tmp.write(kamus_xlsx.getbuffer())
        tmp_path = tmp.name
    try:
        if current_dict is None:
            current_dict = CustomDictionary()
        n = current_dict.load_from_excel(tmp_path)
        st.session_state['custom_dict'] = current_dict
        st.sidebar.success(f"✅ {n} istilah dari Excel dimuat.")
    except Exception as e:
        st.sidebar.error(f"Gagal: {e}")
    finally:
        _os.unlink(tmp_path)

# Tambah istilah manual
with st.sidebar.expander("✏️ Tambah Istilah Manual"):
    col_s, col_t = st.columns(2)
    with col_s:
        manual_src = st.text_input("Asing", key="man_src", placeholder="shear wall")
    with col_t:
        manual_tgt = st.text_input("Indonesia", key="man_tgt", placeholder="dinding geser")
    if st.button("➕ Tambah", key="btn_add_term"):
        if manual_src and manual_tgt:
            if current_dict is None:
                current_dict = CustomDictionary()
            current_dict.add_term(manual_src, manual_tgt)
            st.session_state['custom_dict'] = current_dict
            st.success(f'"{manual_src}" → "{manual_tgt}" ditambahkan!')
        else:
            st.warning("Isi kedua kolom.")

# Info jumlah & lihat isi kamus
if st.session_state['custom_dict'] and len(st.session_state['custom_dict']) > 0:
    d_now = st.session_state['custom_dict']
    st.sidebar.info(f"📖 Kamus aktif: **{len(d_now)} istilah**")
    with st.sidebar.expander("👁️ Lihat Isi Kamus"):
        terms = d_now.list_terms()
        for src, tgt in terms[:50]:
            st.markdown(f"- `{src}` → **{tgt}**")
        if len(terms) > 50:
            st.caption(f"... dan {len(terms)-50} istilah lainnya.")
    if st.sidebar.button("🗑️ Reset Kamus", key="btn_reset_dict"):
        st.session_state['custom_dict'] = None
        st.sidebar.success("Kamus dikosongkan.")
else:
    st.sidebar.caption("⚪ Tidak ada kamus aktif.")

# --- HALAMAN 1: KONVERSI PDF ---
if menu == "1. Konversi (PDF -> Word)":
    st.header("📄 Konversi PDF ke Word (Raw)")
    st.info("Digunakan untuk mendapatkan file mentah dari PDF. Mendukung OCR otomatis.")
    
    uploaded_file = st.file_uploader("Upload PDF", type=["pdf"])
    
    if uploaded_file:
        if st.button("🚀 Konversi Sekarang"):
            with st.spinner("Menganalisis dan mengonversi dokumen..."):
                success, output_path, mode_msg, error = handle_pdf_conversion(uploaded_file)
                
                if success:
                    st.success(f"Berhasil! {mode_msg}")
                    with open(output_path, "rb") as f:
                        st.download_button(
                            "📥 Download Word (Mentah)", 
                            f, 
                            file_name=uploaded_file.name.replace(".pdf", ".docx")
                        )
                    # Simpan path ke session state agar bisa lanjut ke menu lain
                    st.session_state['last_docx'] = output_path
                else:
                    st.error(f"Gagal: {error}")

# --- HALAMAN 2: OPTIMIZER ---
elif menu == "2. Rapikan (Word -> ISO Std)":
    st.header("🛠️ Word Optimizer (ISO Standard)")
    st.info("Merapikan spasi, judul center, justify, dan font standar ISO (Arial 11).")
    
    # Opsi Input: Upload Baru atau Pakai Hasil Konversi Sebelumnya
    input_source = st.radio("Sumber File:", ["Upload File Word", "Gunakan Hasil Konversi Terakhir"])
    
    target_file = None
    
    if input_source == "Upload File Word":
        uploaded_file = st.file_uploader("Upload .docx", type=["docx"])
        if uploaded_file:
            target_file = f"temp_{uploaded_file.name}"
            with open(target_file, "wb") as f:
                f.write(uploaded_file.getbuffer())
                
    elif input_source == "Gunakan Hasil Konversi Terakhir":
        if 'last_docx' in st.session_state and os.path.exists(st.session_state['last_docx']):
            target_file = st.session_state['last_docx']
            st.markdown(f"📄 *Menggunakan file: {target_file}*")
        else:
            st.warning("Belum ada file yang dikonversi di sesi ini.")

    # --- OPSI HEADER/FOOTER ---
    if target_file:
        st.divider()
        st.subheader("⚙️ Pengaturan Header/Footer")
        
        col1, col2 = st.columns(2)
        with col1:
            enable_headers = st.checkbox("Aktifkan Header/Footer", value=False, 
                                        help="Header halaman genap + Footer halaman pertama")
        
        doc_title = ""
        copyright_text = "©BSN 2025"
        
        if enable_headers:
            with col2:
                doc_title = st.text_input("Judul Dokumen (untuk header)", 
                                         value="SNI ISO 15118-1:2019",
                                         help="Muncul di header halaman genap")
            copyright_text = st.text_input("Copyright Text", 
                                          value="©BSN 2025",
                                          help="Muncul di footer halaman pertama (kiri)")
            
            st.info("""
            **Preview Header/Footer:**
            - **First Page Footer**: `{} | 1 dari X` (kiri-kanan)
            - **Even Page Header**: `{}` (center, bold)
            - Different First Page & Odd/Even: ✓ Enabled
            """.format(copyright_text, doc_title))

        # --- OPSI COVER/SAMPUL ---
        st.divider()
        st.subheader("📄 Pengaturan Cover/Sampul (Engine 4)")
        enable_cover = st.checkbox("Tambahkan Halaman Cover (Sampul)", value=False,
                                   help="Membuat halaman sampul sesuai standar BSN sebelum isi dokumen")
        
        cover_settings = {}
        if enable_cover:
            # ── Auto-sync Nomor SNI dari Judul Dokumen (header) ──
            _sni_key = 'cover_sni_input'
            _prev_title = st.session_state.get('_prev_doc_title', None)

            # Jika doc_title berubah, sinkronkan ke field SNI
            if doc_title and doc_title != _prev_title:
                st.session_state[_sni_key] = doc_title
            st.session_state['_prev_doc_title'] = doc_title

            # Default awal jika belum ada di session state
            if _sni_key not in st.session_state:
                st.session_state[_sni_key] = doc_title if doc_title else "SNI ISO XXXXX:20XX"

            col_c1, col_c2 = st.columns(2)
            with col_c1:
                cover_sni_number = st.text_input(
                    "Nomor SNI",
                    key=_sni_key,
                    help="Otomatis diambil dari 'Judul Dokumen' di Pengaturan Header/Footer. Bisa diedit manual."
                )
                if doc_title and cover_sni_number == doc_title:
                    st.caption("🔗 *Auto-fill dari Judul Dokumen. Ketik untuk mengganti.*")
                else:
                    st.caption("✏️ *Diisi manual.*")
                cover_bsn_year = st.text_input(
                    "Tahun Penetapan BSN", value="20XX",
                    help="Tahun yang tertera di cover, misal: 2020"
                )
                cover_ics = st.text_input(
                    "Nomor ICS", value="XX.XXX.XX",
                    help="Contoh: 45.060.01"
                )
            with col_c2:
                cover_ref = st.text_input(
                    "Referensi Standar (kosongkan untuk otomatis)",
                    help="Contoh: ISO 19659-2:2020, IDT — kosongkan untuk otomatis dari Nomor SNI"
                )
                st.info(
                    "ℹ️ **Judul cover diambil otomatis dari dokumen.**\n\n"
                    "- **Judul Bahasa Indonesia** → baris judul pertama (font besar / heading)\n"
                    "- **Judul Bahasa Inggris** → baris judul italic di bawahnya\n\n"
                    "Judul akan tampil setelah dokumen diproses."
                )

            cover_settings = {
                "sni_number": cover_sni_number,
                "bsn_year": cover_bsn_year,
                "ics_number": cover_ics,
                "ref_standard": cover_ref,
            }

        # --- OPSI DAFTAR ISI ---
        st.divider()
        st.subheader("📋 Pengaturan Daftar Isi (Engine 5)")
        enable_daftar_isi = st.checkbox(
            "Tambahkan Halaman Daftar Isi",
            value=False,
            disabled=not enable_cover,
            help="Daftar Isi disisipkan antara Cover dan Isi dokumen. Membutuhkan Cover aktif."
        )
        if not enable_cover and enable_daftar_isi:
            st.info("ℹ️ Aktifkan Cover terlebih dahulu untuk menambahkan Daftar Isi.")

        # --- OPSI PRAKATA & PENDAHULUAN ---
        st.divider()
        st.subheader("📝 Pengaturan Prakata & Pendahuluan (Engine 6)")
        enable_prakata = st.checkbox(
            "Tambahkan Prakata & Pendahuluan",
            value=False,
            disabled=not enable_daftar_isi,
            help="Prakata & Pendahuluan disisipkan setelah Daftar Isi, satu section (Romawi). Membutuhkan Daftar Isi aktif."
        )
        if not enable_daftar_isi and enable_prakata:
            st.info("ℹ️ Aktifkan Daftar Isi terlebih dahulu untuk menambahkan Prakata & Pendahuluan.")
        if enable_prakata and enable_daftar_isi:
            st.info("""
            **Preview Prakata & Pendahuluan yang akan dibuat:**
            - **Satu section** dengan Daftar Isi (penomoran Romawi berlanjut)
            - **Prakata**: template BSN standar dengan substitusi nomor SNI & judul
            - **Pendahuluan**: judul + 3 baris placeholder (merah, bold)
            - Font: Arial 11pt (body), Arial 12pt Bold (judul bagian)
            - Judul SNI & referensi diambil otomatis dari pengaturan Cover
            """)

        # --- OPSI INFO PENDUKUNG PERUMUS STANDAR ---
        st.divider()
        st.subheader("📋 Informasi Pendukung Perumus Standar (Engine 7)")
        enable_info_pendukung = st.checkbox(
            "Tambahkan Halaman Info Pendukung Perumus",
            value=False,
            disabled=not enable_cover,
            help="Halaman terakhir setelah Bibliografi. Section baru, header/footer bersih."
        )
        if not enable_cover and enable_info_pendukung:
            st.info("ℹ️ Aktifkan Cover terlebih dahulu.")
        if enable_info_pendukung and enable_cover:
            st.info("""
            **Preview Halaman Info Pendukung:**
            - **Judul**: *Informasi pendukung terkait perumus standar* (Arial 12 Bold Center)
            - **[1]** Komtek perumus SNI → Komite Teknis xx-yy zzzzzz
            - **[2]** Susunan keanggotaan → Tabel 3 kolom (Ketua, Wakil Ketua, Sekretaris, Anggota)
            - **[3]** Konseptor terjemahan rancangan SNI
            - **[4]** Editor rancangan SNI
            - **[5]** Sekretariat pengelola Komtek perumus SNI
            - **Header/Footer**: Bersih (tanpa logo, tanpa nomor halaman)
            - **Margin**: Top 3cm | Left 3cm | Bottom 2cm | Right 2cm
            - **Section**: Baru (setelah Bibliografi)
            """)

    if target_file and st.button("✨ Jalankan Optimasi"):
        output_file = f"opt_{os.path.basename(target_file)}"
        
        with st.spinner("Mendeteksi judul, bab, dan merapikan spasi..."):
            # Pass parameters ke engine
            success, msg = engine2.process(
                target_file, 
                output_file, 
                ISO_FONT_NAME, 
                ISO_FONT_SIZE,
                enable_headers=enable_headers,
                doc_title=doc_title if enable_headers else "",
                copyright_text=copyright_text if enable_headers else "©BSN 2025"
            )
            
            if success:
                final_output_file = output_file

                # ── Tambahkan Cover jika diaktifkan
                if enable_cover and cover_settings:
                    with st.spinner("Mengekstrak judul & membuat halaman cover/sampul..."):
                        cover_output = f"cover_{os.path.basename(output_file)}"

                        # ── Auto-ekstrak judul dari dokumen yang sudah dioptimasi
                        auto_title_id, auto_title_en = extract_titles_from_docx(output_file)

                        # Tampilkan preview judul yang diekstrak
                        st.info(
                            "**Judul yang diekstrak dari dokumen:**\n\n"
                            f"**🇮🇩 Indonesia:** {auto_title_id or '(tidak ditemukan)'}\n\n"
                            f"**🇬🇧 Inggris:** {auto_title_en or '(tidak ditemukan)'}"
                        )

                        ok_cover, msg_cover = engine4.prepend_cover(
                            input_docx=output_file,
                            output_docx=cover_output,
                            sni_number=cover_settings["sni_number"],
                            bsn_year=cover_settings["bsn_year"],
                            title_id=auto_title_id,
                            title_en=auto_title_en,
                            ref_standard=cover_settings["ref_standard"],
                            ics_number=cover_settings["ics_number"],
                        )
                        if ok_cover:
                            final_output_file = cover_output
                            st.success("✅ Halaman cover berhasil ditambahkan!")

                            # ── Tambahkan Daftar Isi jika diaktifkan
                            if enable_daftar_isi:
                                with st.spinner("Membuat halaman Daftar Isi..."):
                                    di_output = f"di_{os.path.basename(cover_output)}"
                                    ok_di, msg_di = engine5.insert(
                                        input_docx=final_output_file,
                                        output_docx=di_output,
                                        doc_title=cover_settings["sni_number"],
                                        copyright_text=f"©BSN {cover_settings['bsn_year']}",
                                    )
                                    if ok_di:
                                        final_output_file = di_output
                                        st.success("✅ Daftar Isi berhasil ditambahkan!")

                                        # ── Tambahkan Prakata & Pendahuluan jika diaktifkan
                                        if enable_prakata:
                                            with st.spinner("Menyisipkan Prakata & Pendahuluan..."):
                                                pp_output = f"pp_{os.path.basename(di_output)}"

                                                # Hitung ref_standard otomatis jika kosong
                                                ref_std = cover_settings.get("ref_standard", "").strip()
                                                if not ref_std:
                                                    # Buat dari nomor SNI: "SNI ISO 12707:2019" → "ISO 12707:2019"
                                                    ref_std = re.sub(r'^SNI\s+', '', cover_settings["sni_number"]).strip()

                                                ok_pp, msg_pp = engine6.insert(
                                                    input_docx=final_output_file,
                                                    output_docx=pp_output,
                                                    sni_number=cover_settings["sni_number"],
                                                    title_id=auto_title_id or 'Judul Bahasa Indonesia',
                                                    title_en=auto_title_en or 'Title in English',
                                                    ref_standard=ref_std,
                                                    bsn_year=cover_settings["bsn_year"],
                                                )
                                                if ok_pp:
                                                    final_output_file = pp_output
                                                    st.success("✅ Prakata & Pendahuluan berhasil ditambahkan!")
                                                else:
                                                    st.warning(f"⚠️ Prakata & Pendahuluan gagal: {msg_pp}. File DI tetap disimpan.")
                                    else:
                                        st.warning(f"⚠️ Daftar Isi gagal: {msg_di}. File cover tetap disimpan.")
                        else:
                            st.warning(f"⚠️ Cover gagal dibuat: {msg_cover}. File isi tetap disimpan.")

                # ── Tambahkan Info Pendukung jika diaktifkan (selalu di akhir pipeline)
                if enable_info_pendukung and enable_cover:
                    with st.spinner("Menambahkan halaman Informasi Pendukung Perumus..."):
                        ip_output = f"ip_{os.path.basename(final_output_file)}"
                        ok_ip, msg_ip = engine7.append(
                            input_docx=final_output_file,
                            output_docx=ip_output,
                        )
                        if ok_ip:
                            final_output_file = ip_output
                            st.success("✅ Halaman Info Pendukung berhasil ditambahkan!")
                        else:
                            st.warning(f"⚠️ Info Pendukung gagal: {msg_ip}. File sebelumnya tetap disimpan.")

                st.balloons()
                st.success("Dokumen berhasil dirapikan!")
                
                # Show summary
                summary = [
                    "✓ Format heading & spacing",
                    "✓ Bibliography hanging indent",
                    "✓ NOTE formatting",
                    "✓ List items (a,b,c) tidak bold",
                    "✓ Margin: Top 3cm, Inside 3cm, Bottom 2cm, Outside 2cm"
                ]
                if enable_headers:
                    summary.append(f"✓ Header: {doc_title}")
                    summary.append(f"✓ Footer: {copyright_text} | X dari Y")
                if enable_cover and cover_settings:
                    summary.append(f"✓ Cover: {cover_settings['sni_number']} | ICS {cover_settings['ics_number']}")
                if enable_daftar_isi and enable_cover:
                    summary.append("✓ Daftar Isi: penomoran Romawi (i, ii, iii, ...)")
                if enable_prakata and enable_daftar_isi:
                    summary.append("✓ Prakata & Pendahuluan: satu section dengan Daftar Isi")
                if enable_info_pendukung and enable_cover:
                    summary.append("✓ Info Pendukung: halaman terakhir, section baru, header/footer bersih")
                
                st.markdown("**Fitur yang diterapkan:**\n" + "\n".join(f"- {s}" for s in summary))
                
                with open(final_output_file, "rb") as f:
                    st.download_button("📥 Download Hasil Akhir", f, file_name="ISO_Fixed_Document.docx")

                # Simpan path final ke session_state agar bisa diakses
                # oleh panel Engine 8 di bawah (di luar blok tombol ini)
                st.session_state['e8_ready_file'] = final_output_file
                st.success("💡 Dokumen siap diterjemahkan — lihat panel **Engine 8** di bawah.")

            else:
                st.error(msg)

    # ── PANEL ENGINE 9: Terjemahkan ke Bahasa Indonesia ──────────────────────
    # Panel ini ada DI LUAR blok "if st.button Jalankan Optimasi" agar
    # tombol terjemah bisa diklik pada rerun berikutnya tanpa harus
    # mengulang optimasi.
    if st.session_state.get('e8_ready_file') and os.path.exists(st.session_state['e8_ready_file']):
        st.divider()
        st.subheader("🌐 Terjemahkan ke Bahasa Indonesia (Engine 9)")
        st.info(
            "Dokumen siap diterjemahkan. Engine 9 akan menerjemahkan seluruh teks "
            "ke Bahasa Indonesia sambil menyisipkan **teks asli bahasa asing** "
            "sebelum bagian Bibliografi."
            + (f"\n\n📚 Kamus aktif: **{len(st.session_state['custom_dict'])} istilah**"
               if st.session_state.get('custom_dict') else
               "\n\n⚪ Tanpa kamus — aktifkan di sidebar untuk konsistensi istilah teknis.")
        )
        col_tr1, col_tr2 = st.columns([2, 1])
        with col_tr1:
            translate_headers_opt = st.checkbox(
                "Terjemahkan juga teks di Header/Footer",
                value=False,
                key="e8_translate_headers",
            )
        with col_tr2:
            do_translate = st.button("🌍 Terjemahkan Sekarang", key="e8_btn_inline", type="primary")

        if do_translate:
            _ready = st.session_state['e8_ready_file']
            tr_output = f"ID_{os.path.basename(_ready)}"
            _prog_bar = st.progress(0, text="Memulai terjemahan...")

            def _e8_cb(pct, msg):
                _prog_bar.progress(min(pct, 100), text=f"{pct}% — {msg[:70]}")

            # Gunakan kamus dari session state (Engine 9)
            _engine9_inline = DocxFinalTranslatorEngine(
                custom_dict=st.session_state.get('custom_dict')
            )

            with st.spinner("Menerjemahkan + menyisipkan teks asli..."):
                ok_tr, msg_tr = _engine9_inline.translate(
                    input_docx=_ready,
                    output_docx=tr_output,
                    progress_callback=_e8_cb,
                    translate_headers=translate_headers_opt,
                )
            _prog_bar.empty()
            if ok_tr:
                st.success("✅ Dokumen bilingual berhasil dibuat!")
                with open(tr_output, "rb") as f:
                    st.download_button(
                        "📥 Download Dokumen Bilingual (ID + Original)",
                        f,
                        file_name="ID_ISO_Fixed_Document.docx",
                        key="e8_download_inline"
                    )
            else:
                st.error(f"Gagal menerjemahkan: {msg_tr}")

# --- HALAMAN 3: TRANSLATOR ---
elif menu == "3. Terjemahkan (EN -> ID + Rapikan)":
    st.header("🌐 Smart Translator (Inggris -> Indo)")
    st.info("Menerjemahkan dokumen, menjaga nomor bab, dan langsung merapikan format.")
    
    uploaded_file = st.file_uploader("Upload PDF atau Word", type=["pdf", "docx"])
    
    if uploaded_file:
        file_ext = uploaded_file.name.split(".")[-1].lower()
        
        if st.button("🌍 Terjemahkan & Rapikan"):
            temp_input = f"raw_trans_{uploaded_file.name}"
            final_output = f"Translated_{uploaded_file.name.replace('.pdf', '.docx')}"
            
            # PROSES 1: SIAPKAN FILE WORD
            ready_to_translate = False
            word_path = ""
            
            with st.status("Memproses Dokumen...", expanded=True) as status:
                
                # Langkah A: Cek Tipe File
                if file_ext == "pdf":
                    status.write("🔄 Mengonversi PDF ke Word terlebih dahulu...")
                    success, docx_path, _, err = handle_pdf_conversion(uploaded_file)
                    if success:
                        word_path = docx_path
                        ready_to_translate = True
                    else:
                        status.update(label="Gagal Konversi PDF", state="error")
                        st.error(err)
                else:
                    # Jika sudah Word, simpan saja
                    with open(temp_input, "wb") as f:
                        f.write(uploaded_file.getbuffer())
                    word_path = temp_input
                    ready_to_translate = True
                
                # Langkah B: Translate & Optimize
                if ready_to_translate:
                    status.write("🤖 Menerjemahkan & Merapikan (Engine 3)...")
                    # Menggunakan konstanta hardcoded
                    success, msg = engine3.process(word_path, final_output, ISO_FONT_NAME, ISO_FONT_SIZE)
                    
                    if success:
                        status.update(label="Selesai!", state="complete")
                        st.balloons()
                        with open(final_output, "rb") as f:
                            st.download_button(
                                label="📥 Download Dokumen Terjemahan",
                                data=f,
                                file_name=f"ID_{uploaded_file.name.replace('.pdf', '.docx')}"
                            )
                    else:
                        status.update(label="Gagal Penerjemahan", state="error")
                        st.error(msg)
            
            # Cleanup File Sementara
            if os.path.exists(temp_input): os.remove(temp_input)
            if word_path and os.path.exists(word_path) and word_path != temp_input: os.remove(word_path)
# --- HALAMAN 4: TERJEMAHKAN DOKUMEN FINAL (ENGINE 9) ---
elif menu == "4. Terjemahkan Dokumen Final (→ Bahasa Indonesia)":
    st.header("🌐 Terjemahkan Dokumen Final ke Bahasa Indonesia (Engine 9)")
    st.info(
        "Upload dokumen DOCX hasil pembangunan akhir. Engine 9 akan menerjemahkan "
        "seluruh teks ke Bahasa Indonesia dengan mempertahankan **struktur dokumen identik**: "
        "heading, tabel, gambar, format, margin, header/footer, dan penomoran section.\n\n"
        "🆕 **Engine 9**: Istilah teknis dari **Kamus Istilah** (sidebar) diprioritaskan "
        "sebelum Google Translate untuk hasil terjemahan yang lebih konsisten."
    )

    col_up, col_opt = st.columns([2, 1])
    with col_up:
        uploaded_final = st.file_uploader(
            "Upload Dokumen DOCX Final",
            type=["docx"],
            help="File .docx yang sudah selesai dibangun (sudah ada cover, daftar isi, dll)"
        )
    with col_opt:
        st.markdown("**⚙️ Opsi Terjemahan**")
        src_lang = st.selectbox(
            "Bahasa Sumber",
            options=["auto", "en", "fr", "de", "es", "it", "nl", "pt", "ru", "ja", "zh-CN", "ko", "ar"],
            index=0,
            format_func=lambda x: {
                "auto": "🔍 Deteksi Otomatis",
                "en": "🇬🇧 Inggris",
                "fr": "🇫🇷 Prancis",
                "de": "🇩🇪 Jerman",
                "es": "🇪🇸 Spanyol",
                "it": "🇮🇹 Italia",
                "nl": "🇳🇱 Belanda",
                "pt": "🇵🇹 Portugis",
                "ru": "🇷🇺 Rusia",
                "ja": "🇯🇵 Jepang",
                "zh-CN": "🇨🇳 Mandarin",
                "ko": "🇰🇷 Korea",
                "ar": "🇸🇦 Arab",
            }.get(x, x)
        )
        translate_hf = st.checkbox(
            "Terjemahkan teks Header/Footer",
            value=False,
            help="Header/footer biasanya berisi nomor SNI & copyright — aktifkan hanya jika diperlukan."
        )

    if uploaded_final:
        # Preview nama file output
        out_name = f"ID_{uploaded_final.name}"
        st.markdown(f"📄 File output: **`{out_name}`**")

        st.divider()
        if st.button("🚀 Mulai Terjemahkan", type="primary"):
            # Simpan file sumber ke temp
            temp_src = f"e8_src_{uploaded_final.name}"
            temp_out = f"e8_out_{uploaded_final.name}"

            with open(temp_src, "wb") as f:
                f.write(uploaded_final.getbuffer())

            progress_bar = st.progress(0, text="Memulai terjemahan...")
            status_txt   = st.empty()

            def _progress_cb(pct: int, msg: str):
                progress_bar.progress(min(pct, 100), text=f"{pct}% — {msg[:70]}")
                status_txt.caption(f"⏳ {msg}")

            # Init engine 9 dengan bahasa dan kamus yang dipilih
            engine9_custom = DocxFinalTranslatorEngine(
                source_lang=src_lang,
                target_lang='id',
                custom_dict=st.session_state.get('custom_dict'),
            )

            with st.spinner("Sedang menerjemahkan... (proses mungkin memakan beberapa menit tergantang panjang dokumen)"):
                ok, result = engine9_custom.translate(
                    input_docx=temp_src,
                    output_docx=temp_out,
                    progress_callback=_progress_cb,
                    translate_headers=translate_hf,
                )

            progress_bar.empty()
            status_txt.empty()

            if ok:
                st.balloons()
                st.success("✅ Dokumen berhasil diterjemahkan ke Bahasa Indonesia!")

                dict_info = st.session_state.get('custom_dict')
                if dict_info and len(dict_info) > 0:
                    st.info(f"📚 Kamus digunakan: **{len(dict_info)} istilah** diterapkan sebelum Google Translate.")

                st.markdown("""
                **Apa yang dipertahankan:**
                - ✓ Struktur section, margin, orientasi halaman
                - ✓ Style heading (Heading 1, 2, 3, dst)
                - ✓ Format run: bold, italic, underline, font, ukuran, warna
                - ✓ Tabel: struktur, border, lebar kolom, merge sel
                - ✓ Gambar, drawing, objek tertanam
                - ✓ Header & footer (nomor halaman, copyright)
                - ✓ Nomor/label (angka, kode, SNI, ICS) tidak diubah
                - ✓ 🆕 Istilah teknis dari kamus diterjemahkan konsisten
                """)

                with open(temp_out, "rb") as f:
                    st.download_button(
                        label="📥 Download Dokumen Terjemahan Bahasa Indonesia",
                        data=f,
                        file_name=out_name,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
            else:
                st.error(f"❌ Gagal menerjemahkan:\n\n{result}")
                st.info(
                    "**Kemungkinan penyebab:**\n"
                    "- `deep-translator` belum terinstall → jalankan `pip install deep-translator`\n"
                    "- `pandas` atau `openpyxl` belum terinstall (untuk kamus Excel) → `pip install pandas openpyxl`\n"
                    "- Koneksi internet tidak tersedia\n"
                    "- Google Translate membatasi permintaan (coba lagi beberapa saat)"
                )

            # Cleanup
            for f in [temp_src, temp_out]:
                if os.path.exists(f):
                    os.remove(f)

    else:
        st.markdown("""
        ### 📋 Cara Penggunaan:
        1. Selesaikan pembangunan dokumen di **Menu 2** terlebih dahulu
        2. Download dokumen hasil akhir (`.docx`)
        3. Upload kembali dokumen tersebut di sini
        4. Pilih bahasa sumber dan klik **Mulai Terjemahkan**
        5. Download hasil terjemahan Bahasa Indonesia

        ---
        > 💡 **Tips:** Anda juga bisa menggunakan tombol **"Terjemahkan Sekarang"** langsung
        > di Menu 2 setelah proses pembangunan selesai, tanpa perlu download dan upload ulang.
        """)