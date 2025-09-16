# streamlit_app_ce_batch_secure.py
import io
from dataclasses import dataclass
from typing import List, Dict
import pandas as pd
import streamlit as st
from openpyxl import load_workbook

# --- 비밀번호 보호 ---
PASSWORD = "gkdudwlals"  # 필요 시 여기서 변경 가능
st.title("🔒 비밀번호 보호됨")
pw = st.text_input("비밀번호를 입력하세요", type="password")
if pw != PASSWORD:
    st.warning("올바른 비밀번호를 입력해야 계속할 수 있습니다.")
    st.stop()

st.success("인증 성공! 학회비 조회를 시작할 수 있습니다.")

st.set_page_config(page_title="학회비 일괄 조회", layout="wide")
st.title("학회비 납부 일괄 조회 (학년/반 규칙 + 이름 표시)")

@dataclass
class Hit:
    grade: str
    class_name: str
    row: int
    student_id: str
    name: str
    status_cell: str
    verdict: str

def _norm(s):
    if s is None:
        return ""
    return str(s).strip()

def scan_workbook_bytes(file_bytes: bytes, filename: str, target_id: str,
                        start_row: int = 5, id_col: str = "C", name_col: str = "D", status_col: str = "E") -> List[Hit]:
    wb = load_workbook(io.BytesIO(file_bytes), data_only=True)
    tid = _norm(target_id).replace(" ", "")
    results: List[Hit] = []
    grade = filename.split(".")[0]
    for ws in wb.worksheets:
        class_name = ws.title
        r = start_row
        while True:
            id_cell = ws[f"{id_col}{r}"].value
            sid = _norm(id_cell)
            if sid == "":
                break
            if sid.replace(" ", "") == tid:
                name_val = ws[f"{name_col}{r}"].value
                name_text = _norm(name_val)
                status_val = ws[f"{status_col}{r}"].value
                status_text = _norm(status_val)
                verdict = "미납" if status_text == "미납" else "납부"
                results.append(Hit(grade, class_name, r, sid, name_text, status_text, verdict))
            r += 1
    return results

def batch_check(files, student_ids, start_row=5):
    out_rows = []
    file_blobs: Dict[str, bytes] = {uf.name: uf.read() for uf in files}
    for sid in student_ids:
        sid = _norm(sid)
        if not sid:
            continue
        found = False
        for fname, fbytes in file_blobs.items():
            hits = scan_workbook_bytes(fbytes, fname, sid, start_row=start_row)
            if hits:
                found = True
                for h in hits:
                    out_rows.append({
                        "학번": h.student_id,
                        "이름": h.name,
                        "상태": h.verdict,
                        "원본(E열)": h.status_cell,
                        "학년(파일명)": h.grade,
                        "반(시트명)": h.class_name,
                        "행": h.row
                    })
        if not found:
            out_rows.append({
                "학번": sid,
                "이름": "",
                "상태": "명단에 없음",
                "원본(E열)": "",
                "학년(파일명)": "",
                "반(시트명)": "",
                "행": ""
            })
    return pd.DataFrame(out_rows)

start_row = st.sidebar.number_input("시작 행", min_value=1, value=5, step=1)

st.subheader("1) 학년별 엑셀 파일 업로드")
grade_files = st.file_uploader("1학년.xlsx, 2학년.xlsx ... 업로드", type=["xlsx"], accept_multiple_files=True)

st.subheader("2) 학번 목록 업로드 또는 입력")
uploaded_ids = st.file_uploader("학번 CSV/XLSX 파일 (첫 열이 학번)", type=["csv", "xlsx"])
manual_ids = st.text_area("또는 직접 입력 (쉼표/줄바꿈 구분)", placeholder="20230001, 20230002\n20230003")

student_ids = []
if uploaded_ids:
    if uploaded_ids.name.endswith(".csv"):
        try:
            df = pd.read_csv(uploaded_ids, dtype=str)
        except Exception:
            df = pd.read_csv(uploaded_ids, dtype=str, encoding="cp949")
    else:
        df = pd.read_excel(uploaded_ids, dtype=str)
    student_ids.extend(df.iloc[:, 0].dropna().astype(str).tolist())

if manual_ids.strip():
    for part in manual_ids.splitlines():
        for x in part.split(","):
            if x.strip():
                student_ids.append(x.strip())

student_ids = list(dict.fromkeys(student_ids))

if st.button("일괄 조회"):
    if not grade_files:
        st.warning("엑셀 파일을 업로드하세요.")
    elif not student_ids:
        st.warning("학번을 입력하거나 업로드하세요.")
    else:
        with st.spinner("조회 중..."):
            df_out = batch_check(grade_files, student_ids, start_row=start_row)
        st.success(f"{len(student_ids)}건 조회 완료!")
        st.dataframe(df_out, use_container_width=True)
        csv = df_out.to_csv(index=False).encode("utf-8-sig")
        st.download_button("CSV 다운로드", data=csv, file_name="학회비_일괄조회.csv", mime="text/csv")
