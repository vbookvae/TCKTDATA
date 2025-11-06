# app.py
import io
import re
from datetime import datetime, date
from typing import List, Tuple, Optional, Dict

import pandas as pd
import openpyxl
from openpyxl.styles import numbers
import streamlit as st

st.set_page_config(page_title="Xử lý PX/PN", layout="wide")
DATE_FMT_OUT = "%d-%m-%Y"  # dùng cho hiển thị; khi ghi Excel sẽ set number_format

# ===================== Helpers =====================
# ===== PN helpers & expanders (đÃ chuẩn hóa theo yêu cầu) =====

def _replace_tail_full(base_full: int, end_token: str) -> int:
    """Thay toàn bộ len(end_token) chữ số cuối của base_full bằng end_token."""
    base_str = str(base_full)
    k = min(len(end_token), len(base_str))
    return int(base_str[:-k] + end_token[-k:].zfill(k))

def pn_expand_range_pattern(base_full: int, end_token: str, count: int) -> List[int]:
    """
    BASE…END(k): sinh 'count' số liên tiếp bắt đầu từ số nhỏ hơn giữa BASE và BASE(thay END).
    """
    candidate = _replace_tail_full(base_full, end_token)
    start = min(base_full, candidate)
    return list(range(start, start + count))

def as_date(d):
    """Chuẩn hoá giá trị ngày từ Excel về datetime.date hoặc None."""
    if isinstance(d, datetime): return d.date()
    if isinstance(d, date):     return d
    # Chuỗi dd/mm/yyyy hoặc dd-mm-yyyy
    if isinstance(d, str):
        s = d.strip().replace("-", "/")
        m = re.fullmatch(r"(\d{2})/(\d{2})/(\d{4})", s)
        if m:
            return datetime.strptime(s, "%d/%m/%Y").date()
    return None

def write_excel_with_formats(df: pd.DataFrame, file_name: str, sheet_name: str,
                             ticket_col: str, date_cols: List[str]):
    """Ghi DataFrame ra Excel, ép định dạng:
       - ticket_col: Text '@'
       - date_cols: Date format dd-mm-yyyy.
    """
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        # ghi tạm dữ liệu
        df.to_excel(writer, index=False, sheet_name=sheet_name)
        ws = writer.book[sheet_name]
        # xác định cột
        headers = {cell.value: cell.col_idx for cell in ws[1]}
        # ticket -> text
        if ticket_col in headers:
            col = headers[ticket_col]
            for r in range(2, ws.max_row + 1):
                ws.cell(r, col).number_format = numbers.FORMAT_TEXT  # '@'
        # date -> dd-mm-yyyy
        for dc in date_cols:
            if dc in headers:
                col = headers[dc]
                for r in range(2, ws.max_row + 1):
                    ws.cell(r, col).number_format = "DD-MM-YYYY"
    output.seek(0)
    st.download_button(
        f"Tải {file_name}",
        data=output.getvalue(),
        file_name=file_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# ===================== PX: parser =====================
def parse_header_ngay(text: str) -> Tuple[Optional[date], Optional[date]]:
    """'Ngày dd/mm/yyyy (bàn giao ngày dd/mm[/yyyy])' → (ngày xuất, ngày bàn giao)."""
    if not text: return None, None
    m1 = re.search(r"Ngày\s+(\d{2}/\d{2}/\d{4}).*?bàn giao ngày\s+(\d{2}/\d{2})(?:/(\d{4}))?", text, flags=re.I)
    if not m1:
        m1 = re.search(r"Ngày\s+(\d{2}/\d{2}/\d{4}).*?(\d{2}/\d{2})(?:/(\d{4}))?", text, flags=re.I)
    if m1:
        nx = datetime.strptime(m1.group(1), "%d/%m/%Y").date()
        gy = int(m1.group(3)) if m1.group(3) else nx.year
        bg = datetime.strptime(f"{m1.group(2)}/{gy}", "%d/%m/%Y").date()
        return nx, bg
    m2 = re.search(r"(\d{2}/\d{2}/\d{4})", text)
    if m2:
        nx = datetime.strptime(m2.group(1), "%d/%m/%Y").date()
        return nx, nx
    return None, None

def extract_px_tickets_from_row(values: List[str]) -> List[str]:
    out = []
    for v in values:
        if v is None: continue
        s = str(v).strip()
        if not s: continue
        # nhiều cụm trong 1 ô
        tokens = re.findall(r"\b\d{4,}\s*-\s*[\dA-Za-z]+", s)
        if tokens:
            for token in tokens:
                first = token.split("-", 1)[0].strip()
                if first.isdigit(): out.append(first)
        elif "-" in s:
            first = s.split("-", 1)[0].strip()
            if first.isdigit(): out.append(first)
    return out

def parse_px_sheet(ws) -> pd.DataFrame:
    date_rows = []
    for r in range(1, ws.max_row + 1):
        v = ws.cell(r, 1).value
        if isinstance(v, str) and v.strip().startswith("Ngày"):
            date_rows.append(r)
    if not date_rows:
        return pd.DataFrame(columns=["Mã CT","Phiếu xuất","Ngày xuất","Ngày bàn giao"])
    date_rows.append(ws.max_row + 1)

    out_rows = []
    for i in range(len(date_rows) - 1):
        start, end = date_rows[i], date_rows[i + 1] - 1
        header = (ws.cell(start, 1).value or "").strip()
        nx, bg = parse_header_ngay(header)
        if not nx: continue
        for r in range(start + 1, end + 1):
            row_vals = [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
            for tk in extract_px_tickets_from_row(row_vals):
                out_rows.append({
                    "Mã CT": "PX",
                    "Phiếu xuất": str(tk),      # sẽ định dạng Text khi ghi file
                    "Ngày xuất": as_date(nx),
                    "Ngày bàn giao": as_date(bg),
                })
    return pd.DataFrame(out_rows)

# ===================== PN: parser (đã mở rộng luật) =====================
# ===== PN: mở rộng luật (đồng bộ tham số left_full) =====
def _replace_tail(prev_num: int, sfx: str) -> int:
    """Thay đúng số chữ số cuối của prev_num bằng sfx (dùng cho danh sách hậu tố)."""
    k = len(sfx)
    mod = 10 ** k
    return prev_num - (prev_num % mod) + int(sfx)

def pn_expand_rhs(rhs: str, left_full: Optional[str]) -> List[int]:
    """
    Phân tích phần bên phải dấu '-':
    - Dạng range: BASE...END(k) hoặc BASE…END(k)
    - Dạng liệt kê: FULL, sfx1, sfx2, ...
    - Dạng rút gọn mạnh: chỉ sfx,sfx,... → dùng left_full làm nền (không thêm vào kết quả)
    """
    rhs = rhs.strip()
    out: List[int] = []

    # ---- Case: range (chấp nhận ... hoặc …)
    m = re.fullmatch(r"(\d+)[.…]{3}(\d+)\((\d+)\)", rhs)
    if m:
        base_full = int(m.group(1))
        end_token = m.group(2)
        count = int(m.group(3))
        return pn_expand_range_pattern(base_full, end_token, count)

    # ---- Case: danh sách hậu tố
    parts = [p.strip() for p in rhs.split(",") if p.strip()]
    if not parts:
        return out

    left_num = int(left_full) if (left_full and str(left_full).isdigit()) else None

    first = parts[0]
    if first.isdigit() and len(first) >= 5:
        prev = int(first)
        out.append(prev)
        iterable = parts[1:]
    else:
        if left_num is None:
            return out
        prev = left_num
        iterable = parts

    for token in iterable:
        if not token.isdigit():
            continue
        k = len(token)
        mod = 10 ** k
        new_num = prev - (prev % mod) + int(token)
        # đảm bảo ≥ 5 chữ số
        if len(str(new_num)) < 5 and left_num is not None:
            mod = 10 ** len(token)
            new_num = left_num - (left_num % mod) + int(token)
        out.append(new_num)
        prev = new_num

    return out

def parse_pn_cell(cell_value: str, want_return_suffix: bool) -> List[str]:
    """
    Tách mọi cụm '<left>-<right>' trong 1 ô, mở rộng theo pn_expand_rhs.
    Thêm '-R' nếu want_return_suffix = True.
    """
    s = str(cell_value).strip()
    if not s:
        return []
    results: List[str] = []
    for left, right in re.findall(r"(\d+)\s*-\s*([0-9,().\.]+)", s):
        nums = pn_expand_rhs(right, left_full=left)
        for n in nums:
            results.append(str(n))

    seen = set(); uniq = []
    for x in results:
        tag = f"{x}-R" if want_return_suffix else x
        if tag not in seen:
            seen.add(tag); uniq.append(tag)
    return uniq 

def guess_pn_header(ws) -> Dict[str, int]:
    """Tìm dòng tiêu đề và vị trí các cột chính."""
    for r in range(1, ws.max_row + 1):
        vals = [ws.cell(r, c).value for c in range(1, min(ws.max_column, 12) + 1)]
        texts = [str(v).strip().upper() if v is not None else "" for v in vals]
        try: c_phieu = next(i+1 for i,t in enumerate(texts) if "SỐ PHIẾU" in t or "SO PHIEU" in t)
        except StopIteration: continue
        c_nguon = next((i+1 for i,t in enumerate(texts) if "NGUỒN" in t), None)
        c_ngay  = next((i+1 for i,t in enumerate(texts) if t.startswith("NGÀY") or "NGAY" in t), None)
        c_giao  = next((i+1 for i,t in enumerate(texts) if "NGÀY GIAO" in t or "NGAY GIAO" in t), None)
        return {"row": r, "so_phieu": c_phieu, "nguon": c_nguon, "ngay": c_ngay, "ngay_giao": c_giao}
    return {}

def sheet_has_return_flag(ws) -> bool:
    """Nếu vùng tiêu đề (vài dòng đầu) có chữ RETURN → True."""
    for r in range(1, min(10, ws.max_row) + 1):
        for c in range(1, min(6, ws.max_column) + 1):
            v = ws.cell(r, c).value
            if isinstance(v, str) and "RETURN" in v.upper():
                return True
    return False

def parse_pn_sheet(ws) -> pd.DataFrame:
    pos = guess_pn_header(ws)
    if not pos:
        return pd.DataFrame(columns=["File","Sheet","Số phiếu gốc","Mã CT","Nguồn","Phiếu nhập","Ngày nhập","Ngày bàn giao"])

    want_R = sheet_has_return_flag(ws)
    r0 = pos["row"] + 1
    c_phieu, c_nguon, c_ngay, c_giao = pos["so_phieu"], pos["nguon"], pos["ngay"], pos["ngay_giao"]

    out_rows = []
    r = r0
    while r <= ws.max_row:
        raw = ws.cell(r, c_phieu).value
        if raw is None or str(raw).strip() == "":
            if all(ws.cell(r, c).value in (None, "") for c in range(1, min(ws.max_column, 6)+1)):
                break
            r += 1
            continue

        raw_str = str(raw).strip()          # <-- SỐ PHIẾU GỐC
        nguon = ws.cell(r, c_nguon).value if c_nguon else ""
        ngay  = as_date(ws.cell(r, c_ngay).value)  if c_ngay  else None
        giao  = as_date(ws.cell(r, c_giao).value)  if c_giao  else None

        nums = parse_pn_cell(raw_str, want_return_suffix=want_R)
        for phieu in nums:
            out_rows.append({
                "Số phiếu gốc": raw_str,                # <-- luôn có cột này
                "Mã CT": "PN",
                "Nguồn": "" if nguon is None else str(nguon),
                "Phiếu nhập": phieu,                    # để openpyxl định dạng Text khi ghi
                "Ngày nhập": ngay,
                "Ngày bàn giao": giao,
            })
        r += 1

    return pd.DataFrame(out_rows)

# ===================== UI ================================
st.title("XỬ LÝ DỮ LIỆU PX / PN")

st.markdown("""
**Bước 1.** Upload **một hoặc nhiều** file Excel.  
**Bước 2.** Tick **sheet** cần xử lý trong mỗi file.  
**Bước 3.** Chọn chế độ **PX** hoặc **PN** → bấm **Xử lý** để tải Excel kết quả.
""")

uploaded_files = st.file_uploader("Chọn file Excel", type=["xlsx", "xlsm"], accept_multiple_files=True)

if not uploaded_files:
    st.info("Hãy tải lên ít nhất một file Excel.")
    st.stop()

# load workbooks
workbooks = {}
file_sheets = {}
for f in uploaded_files:
    bio = io.BytesIO(f.read())
    wb = openpyxl.load_workbook(bio, data_only=True)
    workbooks[f.name] = wb
    file_sheets[f.name] = wb.sheetnames

st.write("### Chọn sheet để xử lý")
selected_sheets = {}
cols = st.columns(min(3, len(file_sheets)))
for i, (fname, sheets) in enumerate(file_sheets.items()):
    with cols[i % len(cols)]:
        st.caption(f"**{fname}**")
        selected = st.multiselect(f"Sheet trong {fname}", sheets, default=sheets, key=f"ms_{fname}")
        selected_sheets[fname] = selected

mode = st.radio("Chọn loại xử lý", options=["PX","PN"], horizontal=True)

if mode == "PX":
    if st.button("Xử lý dữ liệu phiếu xuất", type="primary"):
        all_rows = []
        for fname, sheets in selected_sheets.items():
            wb = workbooks[fname]
            for sheet in sheets:
                ws = wb[sheet]
                df = parse_px_sheet(ws)
                if not df.empty:
                    df.insert(0, "File", fname)
                    df.insert(1, "Sheet", sheet)
                    all_rows.append(df)
        if not all_rows:
            st.warning("Không trích xuất được dữ liệu PX từ các sheet đã chọn.")
        else:
            df_all = pd.concat(all_rows, ignore_index=True)
            st.success(f"Đã trích xuất {len(df_all)} dòng PX.")
            st.dataframe(df_all.head(200).assign(
                **{"Ngày xuất": df_all["Ngày xuất"].map(lambda d: d.strftime(DATE_FMT_OUT) if d else ""),
                   "Ngày bàn giao": df_all["Ngày bàn giao"].map(lambda d: d.strftime(DATE_FMT_OUT) if d else "")}
            ))
            write_excel_with_formats(
                df_all, file_name="PX_raw_output.xlsx", sheet_name="PX_raw",
                ticket_col="Phiếu xuất", date_cols=["Ngày xuất", "Ngày bàn giao"]
            )

else:  # PN
    if st.button("Xử lý dữ liệu phiếu nhập", type="primary"):
        all_rows = []
        for fname, sheets in selected_sheets.items():
            wb = workbooks[fname]
            for sheet in sheets:
                ws = wb[sheet]
                df = parse_pn_sheet(ws)
                if not df.empty:
                    df.insert(0, "File", fname)
                    df.insert(1, "Sheet", sheet)
                    all_rows.append(df)
        if not all_rows:
            st.warning("Không trích xuất được dữ liệu PN từ các sheet đã chọn.")
        else:
            df_all = pd.concat(all_rows, ignore_index=True)
            # ... trong nhánh if mode == "PN": sau khi df_all = pd.concat(...)
            order = ["File","Sheet","Số phiếu gốc","Mã CT","Nguồn","Phiếu nhập","Ngày nhập","Ngày bàn giao"]
            for col in order:
                if col not in df_all.columns:
                    df_all[col] = ""    # phòng khi sheet nào đó thiếu
            df_all = df_all.reindex(columns=order)

            st.success(f"Đã trích xuất {len(df_all)} dòng PN.")
            st.dataframe(df_all.head(200).assign(
                **{"Ngày nhập": df_all["Ngày nhập"].map(lambda d: d.strftime(DATE_FMT_OUT) if d else ""),
                "Ngày bàn giao": df_all["Ngày bàn giao"].map(lambda d: d.strftime(DATE_FMT_OUT) if d else "")}
            ))

            # Ghi file: cột phiếu dạng Text, ngày dạng Date dd-mm-yyyy
            write_excel_with_formats(
                df_all,
                file_name="PN_raw_output.xlsx",
                sheet_name="PN_raw",
                ticket_col="Phiếu nhập",
                date_cols=["Ngày nhập", "Ngày bàn giao"]
            )

