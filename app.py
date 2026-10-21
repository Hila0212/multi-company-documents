
import streamlit as st

st.set_page_config(
    page_title="Multi Company Documents",
    page_icon="📄",
    layout="wide"
)

st.title("HỆ THỐNG TẠO CHỨNG TỪ")
st.caption("Quản lý chứng từ tập trung cho nhiều pháp nhân")

companies = [
    "TLO",
    "MD GLOBAL",
    "ELIM",
    "Công ty 4",
    "Công ty 5",
    "Công ty 6",
    "Công ty 7"
]

company = st.selectbox(
    "Chọn pháp nhân",
    companies
)

st.success(f"Đang làm việc với: {company}")

st.subheader("Chức năng")

tab1, tab2, tab3 = st.tabs([
    "Tạo chứng từ",
    "Quản lý mẫu",
    "Cấu hình công ty"
])

with tab1:
    st.info("Chức năng tạo chứng từ sẽ được tích hợp tại đây.")

with tab2:
    st.info("Quản lý mẫu chứng từ riêng cho từng pháp nhân.")

with tab3:
    st.info("Cấu hình thông tin pháp nhân.")
