import streamlit as st
import platform
import logging
import io

from sheet_controller import SheetController
from datetime import datetime
from openai import OpenAI
from tenacity import retry, wait_random_exponential, stop_after_attempt
from docx_controller import DocxController

# 로깅 설정
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# API 키 설정
client = OpenAI(api_key=st.secrets["openai"]["api_key"])
GPT_MODEL = "gpt-4o"

# GPT API 호출 함수
@retry(wait=wait_random_exponential(multiplier=1, max=40), stop=stop_after_attempt(3))
def chat_completion_request(messages, model=GPT_MODEL):
    try:
        response = client.chat.completions.create(model=model, messages=messages)
        return response
    except Exception as e:
        logger.error("ChatCompletion 응답 생성 실패", exc_info=True)
        st.error("ChatCompletion 응답 생성 실패")
        st.error(f"Exception: {e}")
        raise e

# 대화 내역 출력 함수
def render_chat_history(chat_history):
    for msg in chat_history:
        color, align = ('#DCF8C6', 'right') if msg["role"] == "user" else ('#FFFFFF', 'left')
        st.markdown(
            f"<div style='text-align: {align}; background-color: {color}; color: black; padding: 10px; border-radius: 10px; margin: 5px;'>{msg['content']}</div>",
            unsafe_allow_html=True,
        )

# 세션 상태 초기화
if "chat_history" not in st.session_state:
    st.session_state.chat_history = [{"role": "system", "content": "당신은 도움이 되는 챗봇입니다."}]

# 사용자 입력 처리 함수 (엔터 입력 시 호출)
def submit():
    user_input = st.session_state.user_input.strip()
    if user_input:
        st.session_state.chat_history.append({"role": "user", "content": user_input})
        try:
            response = chat_completion_request(st.session_state.chat_history)
            assistant_content = response.choices[0].message.content or ""
            st.session_state.chat_history.append({"role": "assistant", "content": assistant_content})
        except:
            st.session_state.chat_history.append({"role": "assistant", "content": "오류가 발생했습니다."})
    st.session_state.user_input = ""  # 입력 후 초기화

st.title("최신 GPT 챗봇 & Word 저장")

# 사용자 입력 (엔터키로 입력 가능)
st.text_input("메시지를 입력하고 엔터를 누르세요.", key="user_input", on_change=submit)

# 대화 내역 렌더링
render_chat_history(st.session_state.chat_history)

# 대화 초기화 버튼
if st.button("대화 초기화"):
    st.session_state.chat_history = [{"role": "system", "content": "당신은 도움이 되는 챗봇입니다."}]
    st.experimental_rerun()

# Word 파일 생성 및 다운로드 버튼
def generate_docx():
    docx = DocxController()
    docx.add_heading("최종 GPT 답변", level=1)
    docx.add_heading("내용", level=2)
    final_answer = "\n\n".join(msg["content"] for msg in st.session_state.chat_history if msg["role"] == "assistant")
    docx.add_paragraph(final_answer, font_size=12)
    docx_io = io.BytesIO()
    docx.document.save(docx_io)
    docx_io.seek(0)
    return docx_io.getvalue()

docx_data = generate_docx()
st.download_button(
    label="Word 파일로 저장",
    data=docx_data,
    file_name="final_answer.docx",
    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
)



sheet = SheetController('credentials.json', '내구글스프레드시트명')

# 로그인 처리
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
            st.success("로그인 성공")
        else:
            st.error("학번 또는 비밀번호가 틀렸습니다.")
            
else:
    # 사용 제한 횟수 체크
    if st.session_state.usage_count >= st.session_state.usage_limit:
        st.warning("사용 제한 횟수를 초과했습니다.")
    else:
        prompt_type = st.selectbox("프롬프트 유형", ['전반', '교과별'])
        subject = None
        if prompt_type == '교과별':
            subject = st.text_input("교과명 입력 (예: 수학)")

        prompts = sheet.get_prompts(prompt_type, subject)

        question = st.text_input("질문 입력 후 엔터")
        if question:
            messages = [{"role": "system", "content": "\n".join(prompts)}, {"role": "user", "content": question}]
            response = chat_completion_request(messages)
            answer = response.choices[0].message.content or ""
            
            st.markdown(f"**GPT 답변:** {answer}")

            # 사용 횟수 증가 및 로그 저장
            sheet.increment_usage(st.session_state.student_id)
            sheet.log_chat(st.session_state.student_id, question, answer, datetime.now().strftime("%Y-%m-%d %H:%M"))
            st.session_state.usage_count += 1
