import streamlit as st
import pandas as pd
import nltk
import re
import os
import time
import requests
from datetime import datetime, timedelta
from deep_translator import GoogleTranslator
from collections import defaultdict
from nltk.corpus import stopwords
from nltk.tokenize import word_tokenize
from pytube import YouTube
import zipfile
from io import BytesIO
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side


# NLTK 데이터 다운로드 (punkt, stopwords) - 필요한 데이터 추가 가능
try:
    nltk.data.find('tokenizers/punkt')
except LookupError:
    nltk.download('punkt')

try:
    nltk.data.find('corpora/stopwords')
except LookupError:
    nltk.download('stopwords')

###############################################################################
# 1) APIKeyManager 클래스
###############################################################################
class APIKeyManager:
    """
    YouTube API 키들을 리스트로 관리.
    할당량 초과(QuotaExceeded) 시 다음 키로 전환.
    """
    def __init__(self, api_keys):
        self.api_keys = api_keys
        self.current_index = 0
        self.quota_exceeded_keys = set()

    def get_current_key(self):
        return self.api_keys[self.current_index]

    def switch_to_next_key(self):
        self.quota_exceeded_keys.add(self.current_index)
        available_keys = [i for i in range(len(self.api_keys)) if i not in self.quota_exceeded_keys]

        if not available_keys:
            raise Exception("모든 API 키의 할당량이 초과되었습니다.")

        next_keys = [i for i in available_keys if i > self.current_index]
        self.current_index = next_keys[0] if next_keys else available_keys[0]
        st.warning(f"할당량 초과로 인해 다음 API 키로 전환합니다. (키 {self.current_index + 1}/{len(self.api_keys)})")

        return self.get_current_key()

    def has_available_keys(self):
        return len(self.quota_exceeded_keys) < len(self.api_keys)

###############################################################################
# 2) 유틸 함수
###############################################################################
def get_date_range(period=None):
    """검색 기간(ISO8601) 반환"""
    from datetime import datetime, timezone, timedelta
    
    now = datetime.now(timezone.utc)

    
    if period == "hour":
        start_date = now - timedelta(hours=1)
    elif period == "today":
        start_date = now.replace(hour=0, minute=0, second=0, microsecond=0)
    elif period == "week":
        start_date = now - timedelta(days=7)
    elif period == "month":
        start_date = now.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
    elif period == "year":
        start_date = now.replace(month=1, day=1, hour=0, minute=0, second=0, microsecond=0)
    elif period == "3month":
        start_date = now - timedelta(days=90)
    elif period == "6month":
        start_date = now - timedelta(days=180)
    elif period == "2year":
        start_date = now - timedelta(days=730)
    elif period == "3year":
        start_date = now - timedelta(days=1095)
    else:
        return None, None
    
    end_date = now

    start_date_rfc = start_date.strftime('%Y-%m-%dT%H:%M:%SZ')
    end_date_rfc = end_date.strftime('%Y-%m-%dT%H:%M:%SZ')
    
    return start_date_rfc, end_date_rfc

def recommend_keywords(base_keyword, geo="US", n=10):
    """
    Google Gemini API를 통해 base_keyword 연관 키워드를 검색하여
    상위 n개를 리스트로 반환한다.
    """
    import google.generativeai as genai

    # (1) Gemini API 접근에 필요한 key(또는 token) 가져오기
    genai.configure(api_key=st.secrets["gemini"]["api_key"])

    # (2) 모델 선택 (Gemini Pro Vision 또는 Gemini Pro)
    model = genai.GenerativeModel('gemini-pro')  # 텍스트-텍스트 모델 사용

    # (3) 쿼리 작성 (예시: "animal" 관련 유튜브 쇼츠 키워드 추천)
    prompt_parts = [
        f"Please recommend {n} popular and interesting YouTube Shorts keywords related to '{base_keyword}' in {geo}. "
        "Focus on keywords that are likely to be used for searching funny and engaging animal-related Shorts."
        "Provide the keywords as a numbered list."
    ]

    # (4) API 호출 및 응답 처리
    try:
        response = model.generate_content(prompt_parts)
        response.resolve() # response가 None이 아닌지 확인

        if response and hasattr(response, "text") and response.text:
            # (5) 텍스트 응답에서 키워드 추출 및 리스트로 변환
            #     - 예시: "1. 웃긴 동물\n2. 귀여운 강아지\n..."  -> ["웃긴 동물", "귀여운 강아지", ...]
            keyword_lines = response.text.strip().split('\n')
            recommended_keywords = []
            for line in keyword_lines:
                # 번호 제거 (예: "1. 키워드" -> "키워드")
                keyword = re.sub(r'^\d+\.\s*', '', line).strip()
                if keyword: # 빈 문자열이 아닌 경우만 추가
                    recommended_keywords.append(keyword)

            return recommended_keywords[:n] # n개 이하로 잘라서 반환
        else:
            st.warning("Gemini API 응답에서 키워드를 추출하지 못했습니다.")
            return []

    except Exception as e:
        st.error(f"Gemini API 요청 중 오류 발생: {e}")
        return []



###############################################################################
# 3) 메인 로직: get_youtube_shorts
#    (댓글/코멘트, 키워드 분석 관련 부분 삭제)
###############################################################################
import re
import math

def parse_iso8601_duration(duration_str):
    pattern = r'PT(?:(\d+)H)?(?:(\d+)M)?(?:(\d+)S)?'
    match = re.match(pattern, duration_str)
    hours = int(match.group(1)) if match and match.group(1) else 0
    minutes = int(match.group(2)) if match and match.group(2) else 0
    seconds = int(match.group(3)) if match and match.group(3) else 0
    return hours * 3600 + minutes * 60 + seconds

def custom_score(item):
    view_count = item.get('view_count', 0)
    like_count = item.get('like_count', 0)
    like_ratio = (like_count + 1) / (view_count + 100)  # 간단한 비율
    popularity_boost = math.sqrt(view_count)           # 조회수 가중
    return like_ratio * 1000 + popularity_boost

def get_youtube_shorts(api_key_manager, search_query=None, tag_query=None,
                       max_results=50, region_code=None, period=None):
    from googleapiclient.discovery import build

    while api_key_manager.has_available_keys():
        try:
            api_key = api_key_manager.get_current_key()
            youtube = build('youtube', 'v3', developerKey=api_key)

            base_url = "https://www.googleapis.com/youtube/v3"

            # (A) 검색어 결정
            if search_query:
                final_query = f"{search_query} #shorts"  
                # → 필요하다면 여기서 #shorts를 빼고, 오직 videoDuration='short'만으로 판별해도 됨
            elif tag_query:
                cleaned_tag = tag_query.replace('#', '')
                final_query = f"#{cleaned_tag}"
            else:
                st.error("search_query 또는 tag_query 중 하나는 반드시 입력해야 합니다.")
                return []

            # (B) 기간 필터
            published_after, published_before = get_date_range(period)

            shorts_data = []
            next_page_token = None
            processed_videos = set()

            # (C) pagination 처리
            while len(shorts_data) < max_results:
                search_params = {
                    'q': final_query,
                    'part': 'id,snippet',
                    'maxResults': min(50, max_results - len(shorts_data)),
                    'type': 'video',
                    'videoDuration': 'short',     # 일단 short 조건
                    'order': 'viewCount',         # 1차: 조회수 높은 순
                    'safeSearch': 'none',
                    'key': api_key
                }

                if published_after:
                    search_params['publishedAfter'] = published_after
                if published_before:
                    search_params['publishedBefore'] = published_before
                if region_code:
                    search_params['regionCode'] = region_code
                if next_page_token:
                    search_params['pageToken'] = next_page_token

                response = requests.get(f"{base_url}/search", params=search_params, timeout=30)

                # API 할당량 초과 확인
                if response.status_code == 403:
                    response_json = response.json()
                    if "quotaExceeded" in str(response_json.get('error', {})):
                        api_key_manager.switch_to_next_key()
                        continue

                if response.status_code != 200:
                    st.warning(f"검색 API 오류 응답: {response.text}")
                    break

                search_response = response.json()
                if not search_response.get('items'):
                    st.info("검색 결과가 없습니다.")
                    break

                video_ids = [item['id']['videoId'] for item in search_response['items']
                             if item['id']['videoId'] not in processed_videos]
                if not video_ids:
                    break

                # (D) 비디오 상세 조회 (snippet, statistics, contentDetails)
                video_params = {
                    'part': 'snippet,statistics,contentDetails',
                    'id': ','.join(video_ids),
                    'key': api_key
                }

                video_response = requests.get(f"{base_url}/videos", params=video_params, timeout=30)
                if video_response.status_code == 403:
                    response_json = video_response.json()
                    if "quotaExceeded" in str(response_json.get('error', {})):
                        api_key_manager.switch_to_next_key()
                        continue

                if video_response.status_code != 200:
                    st.warning(f"비디오 API 오류 응답: {video_response.text}")
                    break

                video_data = video_response.json()


                # (E) 60초 이하 필터링
                for video_item in video_data.get('items', []):
                    video_id = video_item['id']
                    if video_id in processed_videos:
                        continue

                    snippet = video_item['snippet']
                    statistics = video_item['statistics']
                    content_details = video_item['contentDetails']

                    duration_str = content_details.get('duration', '')
                    total_seconds = parse_iso8601_duration(duration_str)
                    if total_seconds > 60:
                        # Shorts가 아닌 것으로 판단
                        continue

                    # ★ 여기서 좋아요가 0인 경우, 해당 영상을 건너뛰도록 추가합니다.
                    if int(statistics.get('likeCount', 0)) == 0:
                        continue

                    processed_videos.add(video_id)

                    video_info = {
                        'title': snippet['title'],
                        'description': snippet.get('description', ''),
                        'video_id': video_id,
                        'published_at': snippet['publishedAt'],
                        'channel_title': snippet['channelTitle'],
                        'view_count': int(statistics.get('viewCount', 0)),
                        'like_count': int(statistics.get('likeCount', 0)),
                        'url': f'https://www.youtube.com/watch?v={video_id}',
                        'thumbnail_img': snippet['thumbnails']['high']['url']
                    }
                    shorts_data.append(video_info)

                if len(shorts_data) >= max_results:
                    break

                # (F) 페이지 토큰 처리
                next_page_token = search_response.get('nextPageToken')
                if not next_page_token:
                    break

            # (G) 후처리 정렬 (커스텀 스코어)
            # 원한다면, 아래 한 줄만 추가:
            shorts_data = sorted(shorts_data, key=custom_score, reverse=True)

            return shorts_data

        except Exception as e:
            if "quotaExceeded" in str(e):
                if api_key_manager.has_available_keys():
                    api_key_manager.switch_to_next_key()
                    continue
                else:
                    st.error("모든 API 키의 할당량이 초과되었습니다.")
                    return []
            else:
                st.error(f"API 처리 중 오류 발생: {str(e)}")
                return []

    return []


###############################################################################
# 4) Streamlit App
###############################################################################
def main():
    st.title("유튜브 쇼츠 수집 및 다운로드 서비스")

    # **[수정] 세션 상태 명시적 초기화 (main 함수 시작 시점)**
    if 'recommended_keywords_df' not in st.session_state:
        st.session_state.recommended_keywords_df = None

    #######################################
    # 0) 추천 키워드 찾기 섹션
    #######################################
    st.subheader("1. 추천 키워드 검색(10개)")
    # 가이드를 접을 수 있게 만듭니다.
    with st.expander("사용자 가이드"):
        st.info(
            """
            1) **추천받고 싶은 키워드**를 입력하세요 (예: **`animal`**).  
            2) **검색 국가**를 선택할 수 있습니다. **"None"**은 전 세계를 의미합니다.  
            3) **"추천 키워드 찾기"** 버튼을 누르면, 10개의 관련 키워드가 추천됩니다.  
            4) 결과는 **엑셀**로 **다운로드**할 수 있습니다.
            """
        )
    
    # **[추가] 국가 선택 selectbox**
    region_code_0 = st.selectbox(
        "국가 선택 (추천 키워드용)",
        ["US", "KR", "JP", "None"], # 국가 목록 (필요에 따라 추가)
        index=0,
        key="region_for_trend"
    )
    if region_code_0 == "None":
        geo_code = ""  # None 선택 시 geo_code를 빈 문자열로 설정 (전 세계)
    else:
        geo_code = region_code_0 # 선택된 국가 코드 사용


    # 키워드 입력
    base_kw = st.text_input("추천받고 싶은 키워드를 입력하세요 (예: animal)")

    # 추천 버튼 (기존 코드와 거의 동일, 세션 상태 초기화 부분 수정)
    if st.button("추천 키워드 찾기"):
        if not base_kw.strip():
            st.warning("키워드를 입력해주세요.")
        else:
            recommended_list = recommend_keywords(
                base_keyword=base_kw.strip(),
                geo=geo_code,
                n=10
            )

            if recommended_list:
                st.success(f"'{base_kw}' 관련 추천 키워드 (상위 {len(recommended_list)}개, 국가: {geo_code or '전 세계'})")

                # [수정] 추천 키워드 DataFrame 생성 및 세션 상태 저장 (기존 코드와 동일)
                df_recommend = pd.DataFrame({"keyword": recommended_list})
                st.session_state.recommended_keywords_df = df_recommend
            else:
                st.info("연관 키워드를 찾을 수 없습니다.")
                # **[수정] 추천 결과 없을 때도 세션 상태 None으로 명시적 초기화**
                st.session_state.recommended_keywords_df = None
    # **[추가] 추천 버튼을 누르지 않았을 때도 세션 상태 None으로 명시적 초기화**
    elif 'recommended_keywords_df' not in st.session_state or st.session_state.recommended_keywords_df is None:
        st.session_state.recommended_keywords_df = None

    # **[수정] 추천 키워드 표 (세션 상태에 따라 조건부 표시, 조건 로직 명확화)**
    if st.session_state.recommended_keywords_df is not None: # 세션 상태가 None이 아닐 때만 표시
        st.dataframe(st.session_state.recommended_keywords_df, width=400) # 표로 표시

        # 엑셀 다운로드 기능 (기존 코드와 동일, 세션 상태 DataFrame 사용)
        excel_buffer = BytesIO()
        st.session_state.recommended_keywords_df.to_excel(excel_buffer, index=False, sheet_name="RecommendedKeywords")

        st.download_button(
            label="추천 키워드 엑셀 다운로드",
            data=excel_buffer.getvalue(),
            file_name="recommended_keywords.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )


    # secrets.toml에서 API 키 목록 불러오기
    api_keys = st.secrets["youtube_api"]["keys"]
    api_key_manager = APIKeyManager(api_keys)

    ############################################################################
    # 섹션 A) 키워드 엑셀 업로드 → Shorts 데이터 수집 → 엑셀 다운로드
    ############################################################################
    st.divider()  # 가로선 추가
    st.subheader("2. 키워드 기반 YouTube Shorts 콘텐츠 수집")
    with st.expander("사용자 가이드"):
        st.info(
            """
            1) **키워드 목록이 담긴 엑셀 파일**을 업로드합니다.  
                - 해당 엑셀 내 `keyword` 칼럼의 값을 사용합니다.
            2) **국가선택**(Shorts 수집용), **검색기간**, **키워드별 최대 수집 개수**를 설정합니다.  
            3) **"Shorts 데이터 수집 시작"** 버튼을 누르면,  
               - 해당 키워드로 YouTube Shorts를 검색하여  
               - (최대 설정 개수만큼) **조회수가 높고 좋아요가 1개 이상**인 영상들을 수집합니다.  
            4) 수집 완료 후, **결과 엑셀**(썸네일이 포함된)이 **다운로드** 가능합니다.
            """
        )
    
    # 세션 스테이트에 결과 저장 여부 확인
    if "shorts_data_df" not in st.session_state:
        st.session_state["shorts_data_df"] = None

    uploaded_file = st.file_uploader("키워드가 담긴 엑셀 파일을 업로드하세요", type=["xlsx"])

    if uploaded_file is not None:
        df_keywords = pd.read_excel(uploaded_file)

        if "keyword" not in df_keywords.columns:
            st.error("엑셀 파일에 'keyword' 칼럼이 필요합니다.")
        else:
            region_code_1 = st.selectbox(
                "국가 선택 (Shorts 수집용)",
                ["US", "JP", "KR", "None"],
                index=0,
                key="region_for_shorts"  # key를 다르게
            )
            if region_code_1 == "None":
                region_code_1 = None

            period_options = ["None", "hour", "today", "week", "month", "3month", "6month", "2year", "3year"]
            period_choice_1 = st.selectbox(
                "검색기간 선택", 
                period_options, 
                index=0,
                key="period_for_shorts"
            )
            if period_choice_1 == "None":
                period_choice_1 = None


            max_results = st.number_input("키워드당 최대 수집 Shorts 수", min_value=1, max_value=500, value=50)

            if st.button("Shorts 데이터 수집 시작"):
                st.info("데이터 수집을 시작합니다. 잠시만 기다려주세요...")
                
                # **[추가] 프로그레스 바 초기화 (섹션 1)**
                progress_bar_section1 = st.progress(0) # 섹션 1 프로그레스 바
                
                results = []
                num_keywords = len(df_keywords) # 전체 키워드 개수

                for idx, row in df_keywords.iterrows():
                    keyword_val = row['keyword']

                    st.write(f"**[{idx+1}/{len(df_keywords)}]** keyword: {keyword_val}")

                    # Shorts 수집
                    data = get_youtube_shorts(
                        api_key_manager,
                        search_query=keyword_val if pd.notna(keyword_val) else None,
                        tag_query=None,
                        max_results=max_results,
                        region_code=region_code_1,
                        period=period_choice_1
                    )
                    if data:
                        # 수집된 data(=비디오 dict들)에 "keyword" 추가
                        for d in data:
                            d["keyword"] = keyword_val

                        results.extend(data)

                    # API 할당량 보호를 위해 잠시 대기
                    time.sleep(1)

                    # **[추가] 프로그레스 바 업데이트 (섹션 1)**
                    progress_percent = int(((idx + 1) / num_keywords) * 100)
                    progress_bar_section1.progress(progress_percent) # 프로그레스 바 업데이트


                if results:
                    st.success(f"총 {len(results)} 개의 Shorts 데이터를 수집했습니다.")
                    df_result = pd.DataFrame(results)

                    # 여기서 published_at 날짜 포맷 변환 추가
                    df_result["published_at"] = pd.to_datetime(
                        df_result["published_at"], errors="coerce"
                    ).dt.strftime("%Y-%m-%d")

                    # ★★ 3개 칼럼 추가 ★★
                    df_result.insert(0, "keyword", df_result.pop("keyword"))
                    df_result.insert(1, "Region Code", region_code_1) # 예: 선택한 region_code
                    df_result.insert(2, "period", period_choice_1)    # 예: 선택한 period

                    st.session_state["shorts_data_df"] = df_result
                else:
                    st.warning("수집된 데이터가 없습니다.")
                    st.session_state["shorts_data_df"] = None
                
                # **[추가] 프로그레스 바 완료 (섹션 1, 완료 시 100% 채움)**
                progress_bar_section1.progress(100)


    # 수집 결과가 세션에 있다면 다운로드 버튼 표시
    if st.session_state["shorts_data_df"] is not None:
        st.write("아래 버튼을 클릭하여 결과 엑셀을 다운로드하세요.")

        # 썸네일 이미지를 엑셀에 실제로 삽입하기
        df_to_download = st.session_state["shorts_data_df"].copy()

        # Excel 생성
        excel_buffer = BytesIO()
        with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
            # 우선 DF를 시트에 쓴다 (단, 'thumbnail_url'을 제외할 수도 있지만
            # 여기서는 레이아웃 잡기 위해 일단 넣고, 밑에서 이미지를 덮어쓰는 방식)
            df_to_download.to_excel(writer, index=False, sheet_name="ShortsData")

            # 워크북, 워크시트 객체
            workbook = writer.book
            worksheet = writer.sheets["ShortsData"]

            # (2) 헤더 행(제목행) 서식 지정
            header_fill = PatternFill(start_color="CCFFCC", end_color="CCFFCC", fill_type="solid")  # 파스텔 연두
            header_font = Font(bold=True)  # 볼드

            # 얇은(또는 'thin') 검정 테두리
            thin_black_border = Border(
                left=Side(border_style="thin", color="000000"),
                right=Side(border_style="thin", color="000000"),
                top=Side(border_style="thin", color="000000"),
                bottom=Side(border_style="thin", color="000000")
            )

            # worksheet의 1행 = 헤더 셀들
            for cell in worksheet[1]:
                cell.fill = header_fill
                cell.font = header_font
                cell.border = thin_black_border

            # (1) 전체 셀에 대해 중앙정렬 적용
            for row in worksheet.iter_rows():
                for cell in row:
                    cell.alignment = Alignment(horizontal="center", vertical="center")

            # 열 너비 & 행 높이 설정 (예: A~I 컬럼, 필요한 만큼 조정)
            for col_letter in ["A","B","C"]:
                worksheet.column_dimensions[col_letter].width = 15  # 조금 넓게

            worksheet.column_dimensions["D"].width = 25
            worksheet.column_dimensions["E"].width = 30
            for col_letter in ["F","G","H","I","J"]:
                worksheet.column_dimensions[col_letter].width = 14  # 조금 넓게
            worksheet.column_dimensions["K"].width = 7
            worksheet.column_dimensions["L"].width = 30
            # 행 높이도 조금 높임
            worksheet.row_dimensions[1].height = 30
            for row_idx in range(2, len(df_to_download) + 2):  # 헤더 포함
                worksheet.row_dimensions[row_idx].height = 135

            # 썸네일 이미지를 셀에 삽입 (헤더가 1행, 실제 데이터는 2행부터)
            from openpyxl.drawing.image import Image as OpxImage
            for i, row in df_to_download.iterrows():
                cell_row = i + 2  # 2행부터 시작
                thumb_url = row["thumbnail_img"]
                if thumb_url:
                    try:
                        resp = requests.get(thumb_url, timeout=10)
                        if resp.status_code == 200:
                            img_data = BytesIO(resp.content)
                            opx_img = OpxImage(img_data)
                            opx_img.width, opx_img.height = (240, 180)
                            # 이미지 크기를 조정하고 싶다면 아래 예시:
                            # 썸네일 이미지를 J열(10번째 컬럼)에 삽입
                            insert_cell = f"L{cell_row}"
                            worksheet.add_image(opx_img, insert_cell)
                    except:
                        pass

            
            #     A, B, C 열 적용
            target_columns = ["A", "B", "C"]

            # 전체 데이터가 몇 행까지 있는지 계산
            max_row = worksheet.max_row  # 현재 시트 내 가장 마지막 행 번호

            for col_letter in target_columns:
                for row_idx in range(1, max_row + 1):  
                    cell = worksheet[f"{col_letter}{row_idx}"]
                    current_alignment = cell.alignment

                    # 이미 중앙 정렬 등의 설정이 있다면, 기존 alignment 정보를 유지하며 wrapText만 활성화
                    # (만약 기존 alignment가 None이면 기본값으로 Alignment()를 만들어줌)
                    new_alignment = Alignment(
                        horizontal=current_alignment.horizontal if current_alignment else "center",
                        vertical=current_alignment.vertical if current_alignment else "center",
                        wrapText=True
                    )
                    cell.alignment = new_alignment


            #     A, B, I 열에만 줄바꿈 적용하고 싶을 경우
            target_columns = ["D", "E", "L"]

            # 전체 데이터가 몇 행까지 있는지 계산
            max_row = worksheet.max_row  # 현재 시트 내 가장 마지막 행 번호

            for col_letter in target_columns:
                for row_idx in range(1, max_row + 1):  
                    cell = worksheet[f"{col_letter}{row_idx}"]
                    current_alignment = cell.alignment

                    # 이미 중앙 정렬 등의 설정이 있다면, 기존 alignment 정보를 유지하며 wrapText만 활성화
                    # (만약 기존 alignment가 None이면 기본값으로 Alignment()를 만들어줌)
                    new_alignment = Alignment(
                        horizontal=current_alignment.horizontal if current_alignment else "center",
                        vertical=current_alignment.vertical if current_alignment else "center",
                        wrapText=True
                    )
                    cell.alignment = new_alignment

            #H 열에만 적용
            target_columns_2 = ["K"]

            max_row = worksheet.max_row

            for col_letter in target_columns_2:
                for row_idx in range(2, max_row + 1):  
                    cell = worksheet[f"{col_letter}{row_idx}"]
                    current_alignment = cell.alignment

                    link_value = cell.value  # URL이 들어있는 값

                    if link_value and isinstance(link_value, str) and link_value.startswith("http"):
                        # (A) 셀에 하이퍼링크 설정
                        cell.hyperlink = link_value
                        cell.style = "Hyperlink"  # 기본 하이퍼링크 스타일 (파란색 밑줄 등)

                    # 이미 중앙 정렬 등의 설정이 있다면, 기존 alignment 정보를 유지하며 wrapText만 활성화
                    # (만약 기존 alignment가 None이면 기본값으로 Alignment()를 만들어줌)
                    new_alignment = Alignment(
                        horizontal="left",
                        vertical=current_alignment.vertical if current_alignment else "center",
                        wrapText=True
                    )
                    cell.alignment = new_alignment

        # 현재 시간을 파일명에 추가
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        file_name = f"shorts_result_{timestamp}.xlsx"

        excel_buffer.seek(0)  # 버퍼 포인터를 시작점으로 되돌림

        st.download_button(
            label="결과 엑셀 다운로드",
            data=excel_buffer.getvalue(),
            file_name=file_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        
        # 버퍼 정리
        excel_buffer.close()

    ############################################################################
    # 섹션 B) 유튜브 링크 엑셀 업로드 → mp4 다운로드 → ZIP 파일로 제공
    # (항상 표시되도록, 섹션 A와 독립)
    ############################################################################
    st.divider()  # 가로선 추가
    st.subheader("3. YouTube 영상 다운로드 (mp4)")
    with st.expander("사용자 가이드"):
        st.info(
            """
            1) **유튜브 링크(영상 URL)**가 담긴 엑셀 파일을 업로드합니다.  
               - 해당 엑셀 내 `url` 칼럼의 값을 사용합니다.
            2) **"영상 다운로드 및 ZIP 만들기"** 버튼을 누르면,  
               - 각 URL에 해당하는 영상을 mp4로 다운받고,  
               - 파일들을 하나의 **ZIP**으로 묶어 줍니다.  
            3) 다운로드가 완료되면, **ZIP 파일**을 바로 **다운로드**할 수 있습니다.
            """
        )
    
    uploaded_links_file = st.file_uploader("유튜브 링크가 담긴 엑셀 파일을 업로드하세요", type=["xlsx"], key="video_links")
    if uploaded_links_file is not None:
        df_links = pd.read_excel(uploaded_links_file)
        # 주의: 아래 예시는 'url' 컬럼을 사용한다고 가정
        # 필요 시 컬럼명을 변경해서 사용
        if "url" not in df_links.columns:
            st.error("엑셀 파일에 'url' 컬럼이 필요합니다.")
        else:
            if st.button("영상 다운로드 및 ZIP 만들기"):
                st.info("영상을 다운로드 중입니다. 파일 크기에 따라 시간이 걸릴 수 있습니다...")
            
                import yt_dlp
            
                timestamp = int(time.time())
                download_dir = f"downloaded_videos_{timestamp}"
                os.makedirs(download_dir, exist_ok=True)
            
                # (1) 파일에서 URL 추출 & Shorts → watch?v=... 변환 (필요 시)
                video_links = []
                for i, row in df_links.iterrows():
                    url = row["url"]
                    if pd.isna(url) or not isinstance(url, str):
                        continue
                    if "youtube.com/shorts/" in url:
                        video_id = url.split("/")[-1]
                        url = f"https://www.youtube.com/watch?v={video_id}"
                    video_links.append(url)
            
                num_videos = len(video_links)
                st.write(f"총 {num_videos}개의 링크를 다운로드합니다...")
            
                # (2) yt-dlp 옵션
                ydl_opts = {
                    'format': 'bestvideo[ext=mp4][vcodec^=avc]+bestaudio[ext=m4a]/best[ext=mp4]',
                    'outtmpl': os.path.join(download_dir, '%(title)s.%(ext)s'),
                    'merge_output_format': 'mp4',
                    'http_headers': {
                         'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) '
                                       'AppleWebKit/537.36 (KHTML, like Gecko) '
                                       'Chrome/112.0.0.0 Safari/537.36'
                    },
                    'force-ipv4': True,
                    'postprocessors': [{
                         'key': 'FFmpegVideoConvertor',
                         'preferedformat': 'mp4'
                    }],
                }

                
                
                # (3) 개별 링크 다운로드
                failed_list = []
                downloaded_count = 0
                progress_bar_section2 = st.progress(0)
            
                for idx, link in enumerate(video_links, start=1):
                    try:
                        with yt_dlp.YoutubeDL(ydl_opts) as ydl:
                            ydl.download([link])
                        downloaded_count += 1
                    except Exception as e:
                        failed_list.append({"url": link, "error_msg": str(e)})
            
                    progress_percent = int(idx / num_videos * 100)
                    progress_bar_section2.progress(progress_percent)
            
                st.success(f"다운로드 완료: 총 {num_videos}개 중 {downloaded_count}개 성공")
            
                # (3-1) 실패 목록 표시
                if failed_list:
                    st.warning(f"{len(failed_list)}개 영상에서 에러가 발생했습니다.")
                    df_failed = pd.DataFrame(failed_list)
                    with st.expander("다운로드 실패 목록 보기"):
                        st.dataframe(df_failed)
            
                # (4) ZIP 압축
                zip_file_name = f"youtube_videos_{timestamp}.zip"
                with zipfile.ZipFile(zip_file_name, 'w', zipfile.ZIP_DEFLATED) as zipf:
                    for root, dirs, files in os.walk(download_dir):
                        for file in files:
                            file_path = os.path.join(root, file)
                            zipf.write(file_path, arcname=file)
            
                # (5) ZIP 다운로드 버튼
                with open(zip_file_name, "rb") as f:
                    st.download_button(
                        label="ZIP 파일 다운로드",
                        data=f,
                        file_name=zip_file_name,
                        mime="application/zip"
                    )
            
                # (6) 임시 파일/폴더 정리
                try:
                    os.remove(zip_file_name)
                    for f_name in os.listdir(download_dir):
                        os.remove(os.path.join(download_dir, f_name))
                    os.rmdir(download_dir)
                except:
                    pass

if __name__ == "__main__":
    main()
