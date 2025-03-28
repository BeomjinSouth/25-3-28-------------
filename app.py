import json
import streamlit as st
import platform
import os
import logging

from openai import OpenAI
from tenacity import retry, wait_random_exponential, stop_after_attempt
from hwp_controller import HwpController  # HWP 문서 제어 모듈

# 로깅 설정
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Streamlit Secrets에 저장한 OpenAI API 키를 불러와서 설정 (일관성 있는 딕셔너리 표기법 사용)
import openai
openai.api_key = st.secrets["openai_api_key"]

# 최신 GPT 모델과 OpenAI 클라이언트 설정
GPT_MODEL = "gpt-4o"
client = OpenAI()

# 재시도 로직을 포함한 GPT API 호출 함수
@retry(wait=wait_random_exponential(multiplier=1, max=40), stop=stop_after_attempt(3))
def chat_completion_request(messages, tools=None, tool_choice=None, model=GPT_MODEL):
    try:
        response = client.chat.completions.create(
            model=model,
            messages=messages,
            tools=tools,
            tool_choice=tool_choice,
        )
        return response
    except Exception as e:
        logger.error("ChatCompletion 응답 생성 실패", exc_info=True)
        st.error("ChatCompletion 응답 생성 실패")
        st.error(f"Exception: {e}")
        raise e

# 대화 내역을 예쁘게 출력하는 함수 (Streamlit HTML 활용)
def render_chat_history(chat_history):
    for msg in chat_history:
        if msg["role"] == "user":
            st.markdown(
                f"""<div style='text-align: right; background-color: #DCF8C6; padding: 10px; 
                border-radius: 10px; margin: 5px;'>{msg["content"]}</div>""",
                unsafe_allow_html=True,
            )
        elif msg["role"] == "assistant":
            st.markdown(
                f"""<div style='text-align: left; background-color: #FFF; padding: 10px; 
                border-radius: 10px; margin: 5px;'>{msg["content"]}</div>""",
                unsafe_allow_html=True,
            )
        elif msg["role"] == "system":
            st.markdown(
                f"""<div style='text-align: center; color: red; margin: 5px;'>{msg["content"]}</div>""",
                unsafe_allow_html=True,
            )

# OS 환경 체크: HWP 기능은 Windows에서만 동작
is_windows = platform.system() == "Windows"
if not is_windows:
    st.warning("HWP 관련 기능은 Windows 환경에서만 작동합니다.")

# Streamlit 초기 설정: session_state에 대화 내역 초기화
if "chat_history" not in st.session_state:
    st.session_state.chat_history = [
        {"role": "system", "content": "당신은 도움이 되는 챗봇입니다."}
    ]

st.title("최신 GPT API & HWP 저장 기능 챗봇")

# 사용자 입력 (빈 문자열 및 공백만 있는 경우 필터링)
user_input = st.text_input("메시지를 입력하세요:")
if st.button("전송") and user_input.strip():
    st.session_state.chat_history.append({"role": "user", "content": user_input.strip()})
    try:
        response = chat_completion_request(st.session_state.chat_history)
        resp_message = response.choices[0].message
        if resp_message.content is None and hasattr(resp_message, "tool_calls") and resp_message.tool_calls:
            assistant_content = "[툴 호출 결과 처리 필요]"
        else:
            assistant_content = resp_message.content or ""
        st.session_state.chat_history.append({
            "role": "assistant",
            "content": assistant_content
        })
    except Exception as e:
        st.session_state.chat_history.append({
            "role": "assistant",
            "content": "응답 생성 중 오류가 발생했습니다."
        })

# 대화 내역 출력
st.markdown("### 대화 내역")
render_chat_history(st.session_state.chat_history)

# 대화 초기화 버튼
if st.button("대화 초기화"):
    st.session_state.chat_history = [
        {"role": "system", "content": "당신은 도움이 되는 챗봇입니다."}
    ]
    st.experimental_rerun()

# 한글 파일로 저장 버튼: 지금까지의 GPT 응답(assistant 메시지들)을 모아서 HWP 파일로 저장
if st.button("한글 파일로 저장"):
    if not is_windows:
        st.error("한글 파일 저장 기능은 Windows 환경에서만 사용 가능합니다.")
    else:
        final_answer = "\n\n".join(
            [msg["content"] for msg in st.session_state.chat_history if msg["role"] == "assistant"]
        )
        hwp = HwpController()
        if hwp.connect():
            hwp.create_new_document()
            # 제목 설정: 굵고 큰 글씨
            hwp.set_font_style(font_name="맑은 고딕", font_size=20, bold=True)
            hwp.insert_text("최종 GPT 답변")
            hwp.insert_paragraph()
            # 소제목 설정
            hwp.set_font_style(font_name="맑은 고딕", font_size=16, bold=True)
            hwp.insert_text("내용")
            hwp.insert_paragraph()
            # 본문 설정: 일반 글씨
            hwp.set_font_style(font_name="맑은 고딕", font_size=12)
            hwp.insert_text(final_answer)
            save_path = "final_answer.hwp"
            if hwp.save_document(save_path):
                st.success(f"한글 파일이 저장되었습니다: {save_path}")
            else:
                st.error("한글 파일 저장에 실패하였습니다.")
        else:
            st.error("한글 프로그램 연결에 실패하였습니다.")
