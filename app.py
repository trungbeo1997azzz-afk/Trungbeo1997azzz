# -*- coding: utf-8 -*-
"""
===============================================================================
  HỆ THỐNG TẠO THÔNG BÁO TẠM NGỪNG CUNG CẤP ĐIỆN
  Dành cho ngành Điện lực Việt Nam - Tổng công ty Điện lực miền Bắc (EVNNPC)
-------------------------------------------------------------------------------
  Chức năng:
    - Tự động đọc dữ liệu khách hàng từ file Excel
    - Tự động thay thế các placeholder vào file Word mẫu
    - Lọc theo Lộ đường dây hoặc Trạm biến áp
    - Gom danh sách thôn/xã không trùng lặp
    - Xuất file Word hoàn chỉnh đúng văn phong EVN
-------------------------------------------------------------------------------
  Cách chạy:  streamlit run app.py
===============================================================================
"""

import io
import re
from copy import deepcopy
from datetime import datetime, time

import pandas as pd
import streamlit as st
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


# ====================================================================
# 1. CẤU HÌNH TRANG VÀ GIAO DIỆN THEO PHONG CÁCH EVN
# ====================================================================
st.set_page_config(
    page_title="Tạo Thông báo Cắt điện - EVNNPC",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded",
)

EVN_CSS = """
<style>
    :root {
        --evn-blue: #003D7A;
        --evn-light-blue: #0066CC;
        --evn-orange: #FF6B00;
        --evn-red:  #E30613;
    }
    html, body, [class*="css"] {
        font-family: 'Times New Roman', 'Segoe UI', sans-serif !important;
    }
    .main-header {
        background: linear-gradient(135deg, #003D7A 0%, #0066CC 100%);
        padding: 25px 20px;
        border-radius: 12px;
        color: white;
        text-align: center;
        margin-bottom: 25px;
        box-shadow: 0 4px 12px rgba(0,61,122,0.25);
    }
    .main-header h1 {
        margin: 0; font-size: 26px; font-weight: 800; letter-spacing: 0.5px;
    }
    .main-header p {
        margin: 6px 0 0 0; font-size: 14px; opacity: 0.95;
    }
    .info-box {
        background-color: #E8F4FD; border-left: 4px solid #0066CC;
        padding: 12px 15px; margin: 10px 0; border-radius: 6px;
    }
    .warning-box {
        background-color: #FFF4E5; border-left: 4px solid #FF6B00;
        padding: 12px 15px; margin: 10px 0; border-radius: 6px;
    }
    .success-box {
        background-color: #E8F5E9; border-left: 4px solid #2E7D32;
        padding: 12px 15px; margin: 10px 0; border-radius: 6px;
    }
    .preview-box {
        background-color: #FFFFFF; border: 1px solid #DDD;
        padding: 30px 40px; border-radius: 8px;
        font-family: 'Times New Roman', serif; color: #000;
        max-height: 600px; overflow-y: auto;
        box-shadow: 0 2px 8px rgba(0,0,0,0.06);
    }
    .preview-box p { line-height: 1.6; margin: 8px 0; }
    .stButton>button {
        background: linear-gradient(135deg, #003D7A 0%, #0066CC 100%);
        color: white; font-weight: bold; border: none;
        padding: 12px 30px; border-radius: 6px; font-size: 16px;
        transition: all 0.2s;
    }
    .stButton>button:hover {
        background: linear-gradient(135deg, #0066CC 0%, #003D7A 100%);
        transform: translateY(-1px);
        box-shadow: 0 4px 8px rgba(0,61,122,0.3);
    }
    .stDownloadButton>button {
        background: linear-gradient(135deg, #FF6B00 0%, #E30613 100%) !important;
        color: white !important; font-weight: bold !important;
        padding: 12px 30px !important; font-size: 16px !important;
    }
    h2, h3, h4 { color: #003D7A; }
    [data-testid="stMetricValue"] { color: #003D7A; font-weight: bold; }
    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #F0F6FC 0%, #E8F0F8 100%);
    }
</style>
"""
st.markdown(EVN_CSS, unsafe_allow_html=True)

# ---- Header trang -------------------------------------------------
st.markdown(
    """
<div class="main-header">
    <h1>⚡ HỆ THỐNG TẠO THÔNG BÁO TẠM NGỪNG CUNG CẤP ĐIỆN ⚡</h1>
    <p>TỔNG CÔNG TY ĐIỆN LỰC MIỀN BẮC – EVNNPC</p>
    <p style="font-size:12px;">Tự động tạo thông báo từ file Excel khách hàng và file Word mẫu</p>
</div>
""",
    unsafe_allow_html=True,
)


# ====================================================================
# 2. CÁC HÀM XỬ LÝ DỮ LIỆU
# ====================================================================
def normalize_text(text) -> str:
    """Chuẩn hóa chuỗi để so sánh: bỏ khoảng trắng thừa và chuyển hoa."""
    if pd.isna(text):
        return ""
    return str(text).strip().upper()


def find_column(df: pd.DataFrame, possible_names):
    """Tìm tên cột linh hoạt, không phân biệt hoa/thường, có/không dấu."""
    cols_lower = {str(c).strip().lower(): c for c in df.columns}
    for name in possible_names:
        key = name.strip().lower()
        if key in cols_lower:
            return cols_lower[key]
    # Tìm chứa từ khóa
    for name in possible_names:
        key = name.strip().lower()
        for c_low, c_orig in cols_lower.items():
            if key in c_low:
                return c_orig
    return None


def filter_data(df, search_term, search_type, col_map):
    """
    Lọc dữ liệu theo Lộ đường dây hoặc Trạm biến áp.
    Trả về: (df_lọc, list_lo, list_tba, list_thon_xa_đã_gom)
    """
    search_norm = normalize_text(search_term)
    col_lo = col_map.get("lo")
    col_tba = col_map.get("tba")
    col_thon = col_map.get("thon")
    col_xa = col_map.get("xa")
    col_huyen = col_map.get("huyen")

    if search_type == "lo":
        mask = df[col_lo].apply(lambda x: search_norm in normalize_text(x))
    else:
        mask = df[col_tba].apply(lambda x: search_norm in normalize_text(x))

    filtered = df[mask].copy()
    if filtered.empty:
        return filtered, [], [], []

    list_lo = sorted({str(x).strip() for x in filtered[col_lo].dropna() if str(x).strip()}) if col_lo else []
    list_tba = sorted({str(x).strip() for x in filtered[col_tba].dropna() if str(x).strip()}) if col_tba else []

    # Gom thôn/xã: gộp các thôn cùng xã thành một dòng
    list_thon_xa = []
    if col_xa:
        grouped = {}
        sub = filtered[[c for c in [col_thon, col_xa, col_huyen] if c]].dropna(subset=[col_xa])
        for _, row in sub.iterrows():
            xa = str(row[col_xa]).strip()
            huyen = str(row[col_huyen]).strip() if col_huyen and col_huyen in row else ""
            xa_key = f"{xa}|{huyen}" if huyen else xa
            if xa_key not in grouped:
                grouped[xa_key] = {"xa": xa, "huyen": huyen, "thons": set()}
            if col_thon and col_thon in row:
                thon = str(row[col_thon]).strip()
                if thon and thon.lower() not in ("nan", "none", ""):
                    grouped[xa_key]["thons"].add(thon)

        for key in sorted(grouped.keys()):
            info = grouped[key]
            thons = sorted(info["thons"])
            huyen_text = f", huyện {info['huyen']}" if info["huyen"] else ""
            if thons:
                if len(thons) == 1:
                    list_thon_xa.append(f"thôn {thons[0]} – xã {info['xa']}{huyen_text}")
                else:
                    list_thon_xa.append(
                        f"các thôn {', '.join(thons)} – xã {info['xa']}{huyen_text}"
                    )
            else:
                list_thon_xa.append(f"xã {info['xa']}{huyen_text}")

    return filtered, list_lo, list_tba, list_thon_xa


# ====================================================================
# 3. XỬ LÝ FILE WORD - THAY THẾ PLACEHOLDER
# ====================================================================
def _replace_in_paragraph(paragraph, replacements: dict):
    """
    Thay thế placeholder trong paragraph, hỗ trợ trường hợp placeholder bị
    tách qua nhiều run. Giữ định dạng của run đầu tiên có chứa placeholder.
    """
    full_text = paragraph.text
    if not any(k in full_text for k in replacements):
        return

    new_text = full_text
    for k, v in replacements.items():
        new_text = new_text.replace(k, str(v))
    if new_text == full_text:
        return

    runs = paragraph.runs
    if not runs:
        return

    # Gán toàn bộ text mới vào run đầu, xóa các run còn lại
    runs[0].text = new_text
    for run in runs[1:]:
        run.text = ""


def _replace_in_table(table, replacements: dict):
    """Đệ quy thay thế trong bảng (bao gồm bảng lồng nhau)."""
    for row in table.rows:
        for cell in row.cells:
            for para in cell.paragraphs:
                _replace_in_paragraph(para, replacements)
            for inner_tbl in cell.tables:
                _replace_in_table(inner_tbl, replacements)


def replace_placeholders(doc: Document, replacements: dict):
    """Thay thế placeholder ở body + tất cả tables + header + footer."""
    # Body
    for para in doc.paragraphs:
        _replace_in_paragraph(para, replacements)
    for tbl in doc.tables:
        _replace_in_table(tbl, replacements)

    # Header/Footer của mọi section
    for section in doc.sections:
        for hdr in (section.header, section.first_page_header, section.even_page_header):
            if hdr is None:
                continue
            for para in hdr.paragraphs:
                _replace_in_paragraph(para, replacements)
            for tbl in hdr.tables:
                _replace_in_table(tbl, replacements)
        for ftr in (section.footer, section.first_page_footer, section.even_page_footer):
            if ftr is None:
                continue
            for para in ftr.paragraphs:
                _replace_in_paragraph(para, replacements)
            for tbl in ftr.tables:
                _replace_in_table(tbl, replacements)


def ensure_vietnamese_font(doc: Document, font_name="Times New Roman"):
    """Đảm bảo font Times New Roman cho toàn bộ document (hỗ trợ Unicode tiếng Việt)."""

    def _set_run_font(run):
        run.font.name = font_name
        rPr = run._element.get_or_add_rPr()
        rFonts = rPr.find(qn("w:rFonts"))
        if rFonts is None:
            rFonts = OxmlElement("w:rFonts")
            rPr.append(rFonts)
        rFonts.set(qn("w:ascii"), font_name)
        rFonts.set(qn("w:hAnsi"), font_name)
        rFonts.set(qn("w:cs"), font_name)
        rFonts.set(qn("w:eastAsia"), font_name)

    def _process_paragraphs(paras):
        for p in paras:
            for r in p.runs:
                _set_run_font(r)

    _process_paragraphs(doc.paragraphs)
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                _process_paragraphs(cell.paragraphs)


# ====================================================================
# 4. CÁC HÀM ĐỊNH DẠNG NGÀY/GIỜ THEO VĂN PHONG EVN
# ====================================================================
def format_date_vn(d) -> str:
    """Định dạng: ngày DD tháng MM năm YYYY"""
    return f"ngày {d.day:02d} tháng {d.month:02d} năm {d.year}"


def format_time_vn(t) -> str:
    """Định dạng: HHhMM (ví dụ 07h30, 17h00)"""
    return f"{t.hour:02d}h{t.minute:02d}"


# ====================================================================
# 5. SIDEBAR - UPLOAD FILES VÀ HƯỚNG DẪN
# ====================================================================
with st.sidebar:
    st.markdown("### 📁 TẢI FILE LÊN")

    excel_file = st.file_uploader(
        "📊 File Excel dữ liệu khách hàng",
        type=["xlsx", "xls"],
        help="File Excel chứa danh sách khách hàng có các cột: Lộ, TBA, Thôn, Xã",
    )

    word_file = st.file_uploader(
        "📄 File Word mẫu thông báo (.docx)",
        type=["docx"],
        help="File Word chứa các placeholder dạng {{NGAY_CAT}}, {{LO_DUONG_DAY}}, ...",
    )

    st.markdown("---")
    st.markdown("### 🔑 CÁC PLACEHOLDER")
    st.code(
        """{{NGAY_CAT}}
{{GIO_BAT_DAU}}
{{GIO_KET_THUC}}
{{LO_DUONG_DAY}}
{{DS_TRAM}}
{{DS_THON_XA}}
{{LY_DO}}""",
        language="text",
    )

    st.markdown("---")
    with st.expander("📖 Cấu trúc file Excel"):
        st.markdown(
            """
**Các cột Excel cần có (tên linh hoạt):**

| Cột | Tên gợi ý |
|-----|-----------|
| **Lộ đường dây** | `Lộ`, `Lộ đường dây`, `Feeder` |
| **Trạm biến áp** | `TBA`, `Trạm biến áp`, `Tên TBA` |
| **Thôn/Xóm** | `Thôn`, `Xóm`, `Thôn/Xóm` |
| **Xã/Phường** | `Xã`, `Phường`, `Xã/Phường` |
| **Huyện** | `Huyện`, `Quận` *(tùy chọn)* |
| **Số khách hàng** | `Số KH` *(tùy chọn)* |

Hệ thống tự nhận diện cột, không phân biệt hoa/thường.
        """
        )

    with st.expander("📄 Cấu trúc Word mẫu"):
        st.markdown(
            """
File Word mẫu cần có các placeholder ở vị trí mong muốn.
Ví dụ một đoạn trong template:

> *"Thời gian: từ **{{GIO_BAT_DAU}}** đến **{{GIO_KET_THUC}}** **{{NGAY_CAT}}**.*
> *Lộ đường dây: **{{LO_DUONG_DAY}}***
> *Các TBA: **{{DS_TRAM}}***
> *Khu vực ảnh hưởng: **{{DS_THON_XA}}***
> *Lý do: **{{LY_DO}}*** "

Hệ thống sẽ giữ nguyên định dạng (font, size, in đậm, căn lề...) của file mẫu.
        """
        )


# ====================================================================
# 6. FORM NHẬP LIỆU CHÍNH
# ====================================================================
col_main, col_side = st.columns([2, 1])

with col_main:
    st.markdown("### 📝 THÔNG TIN THÔNG BÁO")

    with st.form("form_thong_bao", clear_on_submit=False):
        c1, c2, c3 = st.columns(3)
        with c1:
            ngay_cat = st.date_input("📅 Ngày cắt điện", value=datetime.now())
        with c2:
            gio_bat_dau = st.time_input("🕐 Giờ bắt đầu", value=time(7, 0))
        with c3:
            gio_ket_thuc = st.time_input("🕔 Giờ kết thúc", value=time(17, 0))

        st.markdown("#### 🔍 Tìm kiếm khu vực bị ảnh hưởng")
        search_type = st.radio(
            "Loại tìm kiếm:",
            options=["lo", "tba"],
            format_func=lambda x: "🔌 Theo Lộ đường dây" if x == "lo" else "🏭 Theo Trạm biến áp",
            horizontal=True,
        )

        if search_type == "lo":
            search_term = st.text_input(
                "Nhập tên Lộ đường dây",
                placeholder="VD: 371E9.12, 373E1.5 ...",
                help="Có thể nhập một phần tên lộ, hệ thống sẽ tìm tất cả lộ chứa từ khóa.",
            )
        else:
            search_term = st.text_input(
                "Nhập tên Trạm biến áp",
                placeholder="VD: TBA Đồng Tâm, TBA Hòa Bình 1 ...",
                help="Có thể nhập một phần tên trạm, hệ thống sẽ tìm tất cả TBA chứa từ khóa.",
            )

        ly_do = st.text_area(
            "📋 Lý do công tác",
            value="Sửa chữa, bảo dưỡng định kỳ lưới điện nhằm đảm bảo an toàn và "
                  "nâng cao chất lượng cung cấp điện",
            height=80,
        )

        submitted = st.form_submit_button(
            "⚡ TẠO THÔNG BÁO ⚡", use_container_width=True
        )

with col_side:
    st.markdown("### 📊 TRẠNG THÁI FILE")
    if excel_file:
        try:
            df_check = pd.read_excel(excel_file)
            st.metric("📋 Bản ghi Excel", f"{len(df_check):,}")
            st.metric("📑 Số cột", len(df_check.columns))
            excel_file.seek(0)
        except Exception as e:
            st.error(f"Lỗi: {e}")
    else:
        st.info("⬅️ Chưa có file Excel")

    if word_file:
        st.success("✅ Đã có file Word mẫu")
    else:
        st.info("⬅️ Chưa có file Word")


# ====================================================================
# 7. XỬ LÝ KHI BẤM NÚT [TẠO THÔNG BÁO]
# ====================================================================
if submitted:
    if not excel_file:
        st.error("❌ Vui lòng tải lên file **Excel dữ liệu khách hàng**!")
        st.stop()
    if not word_file:
        st.error("❌ Vui lòng tải lên file **Word mẫu thông báo**!")
        st.stop()
    if not search_term or not search_term.strip():
        st.error(
            f"❌ Vui lòng nhập tên "
            f"{'lộ đường dây' if search_type == 'lo' else 'trạm biến áp'}!"
        )
        st.stop()
    if gio_ket_thuc <= gio_bat_dau:
        st.warning("⚠️ Giờ kết thúc nên muộn hơn giờ bắt đầu. Vui lòng kiểm tra lại.")

    try:
        # ---- 1. Đọc Excel --------------------------------------------------
        with st.spinner("🔄 Đang đọc dữ liệu Excel..."):
            excel_file.seek(0)
            df = pd.read_excel(excel_file)
            df.columns = [str(c).strip() for c in df.columns]

            col_map = {
                "lo": find_column(df, ["Lộ", "Lộ đường dây", "Lo", "Lo duong day",
                                       "Feeder", "Đường dây", "Duong day"]),
                "tba": find_column(df, ["TBA", "Trạm", "Trạm biến áp", "Tên TBA",
                                        "Ten TBA", "Tram bien ap", "Tram"]),
                "thon": find_column(df, ["Thôn", "Xóm", "Thôn/Xóm", "Thon", "Xom"]),
                "xa": find_column(df, ["Xã", "Phường", "Xã/Phường", "Xa", "Phuong"]),
                "huyen": find_column(df, ["Huyện", "Quận", "Huyen", "Quan"]),
            }

            missing = []
            if not col_map["lo"]:
                missing.append("Lộ đường dây")
            if not col_map["tba"]:
                missing.append("Trạm biến áp")
            if not col_map["xa"]:
                missing.append("Xã/Phường")
            if missing:
                st.error(f"❌ Thiếu cột bắt buộc trong Excel: **{', '.join(missing)}**")
                st.info(f"Các cột hiện có: `{', '.join(df.columns)}`")
                st.stop()

        # ---- 2. Lọc dữ liệu ------------------------------------------------
        with st.spinner("🔍 Đang lọc và gom dữ liệu..."):
            filtered, list_lo, list_tba, list_thon_xa = filter_data(
                df, search_term, search_type, col_map
            )

        if filtered.empty:
            st.warning(
                f"⚠️ Không tìm thấy dữ liệu với từ khóa **'{search_term}'**. "
                "Vui lòng kiểm tra lại tên lộ/trạm trong file Excel."
            )
            st.stop()

        # ---- 3. Hiển thị kết quả lọc ---------------------------------------
        st.markdown("---")
        st.markdown("### ✅ KẾT QUẢ LỌC DỮ LIỆU")

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("🔌 Số lộ ĐD", len(list_lo))
        m2.metric("🏭 Số TBA", len(list_tba))
        m3.metric("🏘️ Khu vực", len(list_thon_xa))
        m4.metric("📋 Bản ghi", len(filtered))

        with st.expander("📋 Xem chi tiết danh sách đã gom (không trùng)"):
            cA, cB = st.columns(2)
            with cA:
                st.markdown("**🔌 Lộ đường dây:**")
                for x in list_lo:
                    st.markdown(f"- {x}")
                st.markdown("**🏭 Trạm biến áp:**")
                for x in list_tba:
                    st.markdown(f"- {x}")
            with cB:
                st.markdown("**🏘️ Thôn/Xã bị ảnh hưởng:**")
                for x in list_thon_xa:
                    st.markdown(f"- {x}")

        with st.expander("📊 Xem bảng dữ liệu Excel đã lọc"):
            st.dataframe(filtered, use_container_width=True, height=300)

        # ---- 4. Chuẩn bị placeholder ---------------------------------------
        ds_tram_str = "; ".join(list_tba)
        ds_thon_xa_str = "; ".join(list_thon_xa)
        ds_lo_str = ", ".join(list_lo)

        replacements = {
            "{{NGAY_CAT}}":     format_date_vn(ngay_cat),
            "{{GIO_BAT_DAU}}":  format_time_vn(gio_bat_dau),
            "{{GIO_KET_THUC}}": format_time_vn(gio_ket_thuc),
            "{{LO_DUONG_DAY}}": ds_lo_str,
            "{{DS_TRAM}}":      ds_tram_str,
            "{{DS_THON_XA}}":   ds_thon_xa_str,
            "{{LY_DO}}":        ly_do.strip(),
        }

        # ---- 5. Preview thông báo theo văn phong EVN -----------------------
        st.markdown("### 👁️ XEM TRƯỚC NỘI DUNG THÔNG BÁO")

        preview_html = f"""
        <div class="preview-box">
            <table style="width:100%; border:none;">
                <tr>
                    <td style="width:50%; text-align:center; vertical-align:top; border:none;">
                        <b>TỔNG CÔNG TY ĐIỆN LỰC MIỀN BẮC</b><br>
                        <b style="text-decoration:underline;">CÔNG TY ĐIỆN LỰC ...</b>
                    </td>
                    <td style="width:50%; text-align:center; vertical-align:top; border:none;">
                        <b>CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM</b><br>
                        <b style="text-decoration:underline;">Độc lập – Tự do – Hạnh phúc</b>
                    </td>
                </tr>
            </table>
            <p style="text-align:right; font-style:italic;">............, {format_date_vn(ngay_cat)}</p>

            <h3 style="text-align:center; margin-top:20px;">THÔNG BÁO</h3>
            <p style="text-align:center;"><b><i>V/v tạm ngừng cung cấp điện</i></b></p>

            <p><b>Kính gửi:</b> Quý khách hàng sử dụng điện.</p>

            <p style="text-indent: 30px;">Thực hiện kế hoạch công tác trên lưới điện nhằm đảm bảo
            vận hành an toàn, ổn định và nâng cao chất lượng cung cấp điện phục vụ Quý khách hàng,
            Công ty Điện lực trân trọng thông báo lịch tạm ngừng cung cấp điện như sau:</p>

            <p><b>1. Thời gian tạm ngừng cấp điện:</b><br>
            &nbsp;&nbsp;&nbsp;&nbsp;Từ <b>{format_time_vn(gio_bat_dau)}</b>
            đến <b>{format_time_vn(gio_ket_thuc)}</b> {format_date_vn(ngay_cat)}.</p>

            <p><b>2. Lộ đường dây liên quan:</b> {ds_lo_str}</p>

            <p><b>3. Trạm biến áp bị ảnh hưởng ({len(list_tba)} TBA):</b><br>
            &nbsp;&nbsp;&nbsp;&nbsp;{ds_tram_str}</p>

            <p><b>4. Khu vực mất điện:</b><br>
            &nbsp;&nbsp;&nbsp;&nbsp;{ds_thon_xa_str}.</p>

            <p><b>5. Lý do:</b> {ly_do}.</p>

            <p style="text-indent: 30px;">Công ty Điện lực rất mong Quý khách hàng thông cảm
            và chủ động bố trí công việc, sinh hoạt, sản xuất phù hợp trong thời gian tạm
            ngừng cấp điện nêu trên.</p>

            <p style="text-indent: 30px;">Mọi thông tin chi tiết, Quý khách hàng vui lòng
            liên hệ Tổng đài Chăm sóc khách hàng <b>19006769</b> hoặc website
            <b>cskh.npc.com.vn</b> để được hỗ trợ.</p>

            <p style="text-indent: 30px;"><i>Trân trọng thông báo!</i></p>

            <table style="width:100%; border:none; margin-top:20px;">
                <tr>
                    <td style="width:50%; border:none;"><b>Nơi nhận:</b><br>
                        <i>- Như kính gửi;</i><br>
                        <i>- Lưu: VT, KD.</i>
                    </td>
                    <td style="width:50%; text-align:center; border:none;">
                        <b>GIÁM ĐỐC</b><br><br><br><br>
                    </td>
                </tr>
            </table>
        </div>
        """
        st.markdown(preview_html, unsafe_allow_html=True)

        st.markdown(
            '<div class="info-box">📌 <b>Lưu ý:</b> Phần preview trên chỉ minh họa nội dung. '
            'File Word xuất ra sẽ <b>giữ nguyên định dạng gốc</b> của file mẫu bạn upload '
            '(logo, font, lề, bảng biểu, chữ ký...).</div>',
            unsafe_allow_html=True,
        )

        # ---- 6. Tạo file Word từ template ----------------------------------
        with st.spinner("📄 Đang tạo file Word..."):
            word_file.seek(0)
            doc = Document(word_file)
            replace_placeholders(doc, replacements)
            ensure_vietnamese_font(doc, "Times New Roman")

            output = io.BytesIO()
            doc.save(output)
            output.seek(0)

        safe_term = re.sub(r"[^\w\-]", "_", search_term.strip())
        file_name = f"ThongBao_CatDien_{ngay_cat.strftime('%d-%m-%Y')}_{safe_term}.docx"

        # ---- 7. Nút tải file -----------------------------------------------
        st.markdown("### 📥 TẢI FILE WORD")
        st.markdown(
            '<div class="success-box">✅ <b>Tạo thông báo thành công!</b> '
            'Bấm nút bên dưới để tải file Word về máy.</div>',
            unsafe_allow_html=True,
        )

        st.download_button(
            label="⬇️ TẢI FILE WORD THÔNG BÁO",
            data=output,
            file_name=file_name,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )

    except Exception as e:
        st.error(f"❌ Có lỗi xảy ra: **{str(e)}**")
        import traceback
        with st.expander("🔧 Chi tiết kỹ thuật (để báo lỗi)"):
            st.code(traceback.format_exc())


# ====================================================================
# 8. FOOTER
# ====================================================================
st.markdown("---")
st.markdown(
    """
<div style="text-align:center; color:#555; font-size:13px; padding:15px;">
    <b>⚡ HỆ THỐNG TỰ ĐỘNG HÓA THÔNG BÁO CẮT ĐIỆN ⚡</b><br>
    Phát triển dành cho ngành Điện lực Việt Nam – EVNNPC<br>
    <i>Powered by Python · Streamlit · python-docx · pandas</i>
</div>
""",
    unsafe_allow_html=True,
)
