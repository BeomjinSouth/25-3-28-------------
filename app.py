import streamlit as st
import logging
import io
from datetime import datetime
from openai import OpenAI
from tenacity import retry, wait_random_exponential, stop_after_attempt
from sheet_controller import SheetController
from docx_controller import DocxController

# 로깅 설정
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# API 클라이언트 설정
client = OpenAI(api_key=st.secrets["openai"]["api_key"])
GPT_MODEL = "gpt-4o"

# GPT API 호출 함수
@retry(wait=wait_random_exponential(multiplier=1, max=40), stop=stop_after_attempt(3))
def chat_completion_request(messages, model=GPT_MODEL):
    response = client.chat.completions.create(model=model, messages=messages)
    return response

# SheetController 초기화
sheet = SheetController("2025 03 28 GPT sheets docs")

# 사이드바 로그인 처리
with st.sidebar:
    st.header("🔑 학생 로그인")
    if 'logged_in' not in st.session_state:
        st.session_state.logged_in = False

    if not st.session_state.logged_in:
        student_id = st.text_input("학번 입력")
        password = st.text_input("비밀번호 입력", type='password')
        if st.button("로그인"):
            user = sheet.verify_user(student_id, password)
            if user:
                st.session_state.logged_in = True
                st.session_state.student_id = student_id
                st.session_state.usage_limit = user['답변제한횟수']
                st.session_state.usage_count = user.get('사용횟수', 0)
                st.session_state.chat_history = []
                st.success("로그인 성공")
            else:
                st.error("학번 또는 비밀번호가 틀렸습니다.")
    else:
        st.success(f"{st.session_state.student_id} 로그인 중")
        if st.button("로그아웃"):
            st.session_state.logged_in = False
            st.experimental_rerun()

# 로그인된 사용자만 접근 가능
if st.session_state.logged_in:
    st.title("GPT 챗봇")

    if st.session_state.usage_count >= st.session_state.usage_limit:
        st.warning("사용 제한 횟수를 초과했습니다.")
    else:
        # 프롬프트 선택
        prompt_type = st.selectbox("프롬프트 유형 선택", ['전반', '교과별'])
        subject = st.text_input("교과명 입력 (교과별 선택 시)") if prompt_type == '교과별' else None
        prompts = sheet.get_prompts(prompt_type, subject)

        # 채팅 UI
        chat_placeholder = st.empty()
        user_input = st.text_input("메시지 입력 후 엔터", key="user_input")

        if user_input:
            messages = [{"role": "system", "content": "\n".join(prompts)}]
            messages.extend([{"role": "user", "content": user_input}])

            response = chat_completion_request(messages)
            answer = response.choices[0].message.content or ""

            # 채팅 내역 저장
            st.session_state.chat_history.append({"질문": user_input, "답변": answer})

            # 구글 시트 기록
            sheet.increment_usage(st.session_state.student_id)
            sheet.log_chat(st.session_state.student_id, user_input, answer, datetime.now().strftime("%Y-%m-%d %H:%M"))
            st.session_state.usage_count += 1
            st.session_state.user_input = ""

        # 카카오톡 스타일 채팅 렌더링
        with chat_placeholder.container():
            for chat in reversed(st.session_state.chat_history):
                st.markdown(f"<div style='background-color:#DCF8C6;padding:10px;border-radius:10px;margin:5px;text-align:right;color:black;'>{chat['질문']}</div>", unsafe_allow_html=True)
                st.markdown(f"<div style='background-color:#FFFFFF;padding:10px;border-radius:10px;margin:5px;text-align:left;color:black;'>{chat['답변']}</div>", unsafe_allow_html=True)

        # Word 파일 다운로드 버튼
        def generate_docx():
            docx = DocxController()
            docx.add_heading("GPT 답변 내역", level=1)
            for chat in st.session_state.chat_history:
                docx.add_heading("질문", level=2)
                docx.add_paragraph(chat["질문"])
                docx.add_heading("답변", level=2)
                docx.add_paragraph(chat["답변"])
            docx_io = io.BytesIO()
            docx.document.save(docx_io)
            docx_io.seek(0)
            return docx_io.getvalue()

        st.download_button(
            "Word 파일로 저장", generate_docx(), "chat_history.docx",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
else:
    st.info("로그인 후 GPT 챗봇을 사용할 수 있습니다.")
