# main.py
import streamlit as st
import report_revenue
import report_settlement
import search_shorts
import search_meme
import search_creator

st.title("East Blue 업무지원 서비스")

page = st.sidebar.selectbox("서비스 선택", ["홈", "음원정산시트 생성기", "정산보고서 생성기", "쇼츠 수집 및 다운로드", "밈 검색기", "크리에이터 검색기"])

if page == "홈":
    st.write("환영합니다!")
elif page == "음원정산시트 생성기":
    report_revenue.main()  # 파일 내 정의된 main() 함수 호출
elif page == "정산보고서 생성기":
    report_settlement.main()
elif page == "쇼츠 수집 및 다운로드":
    search_shorts.main()
elif page == "밈 검색기":
    search_meme.main()
elif page == "크리에이터 검색기":
    search_creator.main()
