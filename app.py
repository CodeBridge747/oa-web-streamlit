import io
import urllib.request
from pathlib import Path
from typing import List, Optional

import pandas as pd
import streamlit as st
from fpdf import FPDF 
import plotly.express as px


st.set_page_config(page_title="IT 자산 통합 관리 시스템", page_icon="🖥️")
st.title("🖥️ IT 자산 통합 관리 시스템")

st.write(
    "여러 지점의 엑셀을 한 번에 올려 자동으로 합치고, 상태별 요약/추출/다운로드를 제공합니다."
)


# -------------------------------------------------------------------
# 폰트 준비 (UTF-8 PDF 지원용)
# -------------------------------------------------------------------
FONT_URL = (
    "https://github.com/dejavu-fonts/dejavu-fonts/raw/master/ttf/DejaVuSans.ttf"
)


@st.cache_resource(show_spinner=False)
def get_font_path() -> Optional[str]:
    """DejaVuSans.ttf를 다운로드/캐시하여 경로를 반환."""
    font_path = Path("fonts/DejaVuSans.ttf")
    if font_path.exists():
        return str(font_path)
    try:
        font_path.parent.mkdir(parents=True, exist_ok=True)
        urllib.request.urlretrieve(FONT_URL, font_path)
        return str(font_path)
    except Exception as exc:
        st.warning(f"PDF 생성용 폰트 준비에 실패했습니다: {exc}")
        return None


# -------------------------------------------------------------------
# 1) 데이터 불러오기/합치기 함수
# -------------------------------------------------------------------
def load_excel_or_csv(file_bytes: bytes) -> pd.DataFrame:
    """엑셀 우선 시도, 실패하면 CSV로 읽어 반환."""
    try:
        return pd.read_excel(io.BytesIO(file_bytes), engine="openpyxl")
    except Exception:
        return pd.read_csv(io.BytesIO(file_bytes))


def concat_uploads(files: List[st.runtime.uploaded_file_manager.UploadedFile]) -> Optional[pd.DataFrame]:
    frames = []
    for f in files:
        content = f.read()
        try:
            frames.append(load_excel_or_csv(content))
        except Exception as exc:
            st.error(f"{f.name} 읽기 오류: {exc}")
            return None
    if not frames:
        return None
    return pd.concat(frames, ignore_index=True)


# -------------------------------------------------------------------
# 2) 시리얼 번호 규칙 적용 함수
# -------------------------------------------------------------------
def apply_serial_rule(df: pd.DataFrame, rule: str) -> pd.DataFrame:
    """
    단순 규칙 예시:
    - 접두사 추가: 'prefix=ABC-'
    - 접미사 추가: 'suffix=-2025'
    """
    serial_cols = [c for c in df.columns if "serial" in c.lower() or "시리얼" in c]
    if not serial_cols:
        return df
    col = serial_cols[0]
    if rule.startswith("prefix="):
        prefix = rule.split("prefix=", 1)[1]
        df[col] = prefix + df[col].astype(str)
    elif rule.startswith("suffix="):
        suffix = rule.split("suffix=", 1)[1]
        df[col] = df[col].astype(str) + suffix
    return df


# -------------------------------------------------------------------
# 3) PDF 생성 (수리 대상 추출)
# -------------------------------------------------------------------
def build_repair_pdf(df: pd.DataFrame, status_col: str) -> Optional[bytes]:
    font_path = get_font_path()
    if not font_path:
        return None

    pdf = FPDF()
    pdf.add_page()
    pdf.add_font("DejaVu", "", font_path, uni=True)
    pdf.set_font("DejaVu", size=12)
    pdf.cell(0, 10, "수리 대상자 목록", ln=1)
    subset = df[df[status_col] == "수리 필요"]
    for _, row in subset.iterrows():
        line = ", ".join([f"{col}: {row[col]}" for col in subset.columns])
        pdf.multi_cell(0, 8, line)
    # dest="S"는 문자열을 반환하며, fpdf는 latin-1로 인코딩이 필요합니다.
    return pdf.output(dest="S").encode("latin-1")


uploaded_files = st.file_uploader(
    "지점별 자산 엑셀을 올려 주세요",
    type=["xlsx", "xls", "csv"],
    accept_multiple_files=True,
)

serial_rule = st.text_input(
    "변경할 시리얼 번호 규칙 (예: prefix=HQ-, suffix=-2025)",
    placeholder="원하면 입력",
)

if uploaded_files:
    df = concat_uploads(uploaded_files)

    if df is None or df.empty:
        st.warning("데이터가 비어 있거나 읽기에 실패했습니다.")
        st.stop()

    if serial_rule:
        df = apply_serial_rule(df, serial_rule)

    # 상태 열 선택
    status_col = "Status" if "Status" in df.columns else None
    if not status_col:
        candidates = [c for c in df.columns if df[c].dtype == object]
        if not candidates:
            st.warning("상태를 나타낼 텍스트 열을 찾을 수 없습니다.")
            st.stop()
        status_col = st.selectbox(
            "자산 상태가 적힌 열을 선택하세요",
            options=candidates,
            help="예: 정상, 수리 필요, 폐기 예정",
        )

    total_devices = len(df)
    need_repair = (df[status_col] == "수리 필요").sum()
    to_dispose = (df[status_col] == "폐기 예정").sum()

    col1, col2, col3 = st.columns(3)
    col1.metric("전체 기기 수", f"{total_devices:,}")
    col2.metric("수리 필요 기기 수", f"{need_repair:,}")
    col3.metric("폐기 예정 기기 수", f"{to_dispose:,}")

    with st.expander("데이터 미리보기", expanded=False):
        st.dataframe(df.head(100))

    # 시각화 섹션
    st.subheader("시각화")
    # 부서/기기종류 자동 추정 및 선택
    dept_candidates = [c for c in df.columns if any(k in c.lower() for k in ["dept", "department", "부서"])]
    type_candidates = [c for c in df.columns if any(k in c.lower() for k in ["model", "type", "종류", "모델"])]

    col_sel1, col_sel2 = st.columns(2)
    dept_col = col_sel1.selectbox(
        "부서 열을 선택하세요",
        options=dept_candidates or list(df.columns),
        index=0 if dept_candidates else 0,
    )
    type_col = col_sel2.selectbox(
        "기기 종류/모델 열을 선택하세요",
        options=type_candidates or list(df.columns),
        index=0 if type_candidates else 0,
    )

    vis_col1, vis_col2 = st.columns(2)
    with vis_col1:
        dept_count = df.groupby(dept_col, dropna=False).size().reset_index(name="count")
        fig_bar = px.bar(
            dept_count,
            x=dept_col,
            y="count",
            title="부서별 기기 보유량",
            labels={"count": "수량"},
        )
        st.plotly_chart(fig_bar, use_container_width=True)

    with vis_col2:
        type_count = df.groupby(type_col, dropna=False).size().reset_index(name="count")
        fig_pie = px.pie(
            type_count,
            names=type_col,
            values="count",
            title="기기 종류별 비율",
        )
        st.plotly_chart(fig_pie, use_container_width=True)

    # 엑셀 다운로드
    excel_buffer = io.BytesIO()
    df.to_excel(excel_buffer, index=False, engine="openpyxl")
    st.download_button(
        "정리된 데이터 엑셀 다운로드",
        data=excel_buffer.getvalue(),
        file_name="merged_assets.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    # 수리 대상자 PDF 추출
    if st.button("수리 대상자 추출 (PDF 다운로드)"):
        pdf_bytes = build_repair_pdf(df, status_col)
        if pdf_bytes:
            st.download_button(
                "PDF 다운로드",
                data=pdf_bytes,
                file_name="repair_list.pdf",
                mime="application/pdf",
            )
        else:
            st.error("PDF 생성에 실패했습니다. 인터넷 연결을 확인 후 다시 시도하세요.")

else:
    st.info("좌측의 업로드 영역을 통해 여러 엑셀/CSV 파일을 올려 주세요.")

